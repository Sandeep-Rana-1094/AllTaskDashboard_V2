import React, { useState, useEffect, useMemo, useCallback, useRef } from 'react';
import { createRoot } from 'react-dom/client';
import { ChecklistSystem } from './ChecklistSystem.tsx';
import { DelegationSystem } from './DelegationSystem.tsx';
import { TaskDashboardSystem } from './TaskDashboardSystem.tsx';
import { 
    Task, Checklist, MasterTask, DelegationTask, DashboardTask, AuthenticatedUser, UserAuth, Person, AttendanceData, DailyAttendance, AppMode, TaskHistory, Holiday
} from './types';
import { useLocalStorage, robustCsvParser } from './utils';

// Fix: Declare Google Apps Script global variables to resolve TypeScript errors.
// These are available in the Google Apps Script environment but not in a standard TS/React project.
declare var DriveApp: any;
declare var LockService: any;
declare var SpreadsheetApp: any;
declare var Utilities: any;
declare var ContentService: any;
declare var ScriptApp: any; // Added for trigger setup

// =================================================================================================
// == CRITICAL ACTION: DEPLOY SCRIPT AND SET UP ASYNCHRONOUS TRIGGER ==
// =================================================================================================
// The application has been upgraded to handle many simultaneous users without lag.
// This is achieved by queuing requests and processing them in the background.
// To enable this, you MUST follow these steps carefully.
//
// --- STEP 1: UPDATE THE SCRIPT CODE ---
// Copy the ENTIRE Google Apps Script block below (from "--- START OF GOOGLE APPS SCRIPT ---" 
// to "--- END OF GOOGLE APPS SCRIPT ---") and PASTE it into your "Code.gs" file in the
// MASTER SPREADSHEET'S script editor, replacing ALL existing code.
//
// --- STEP 2: RE-DEPLOY THE SCRIPT AS A WEB APP ---
// This step ensures the latest code is active.
// 1. In the Apps Script editor, click the "Deploy" button (top right).
// 2. Select "New deployment".
// 3. In the dialog, click the gear icon next to "Select type" and choose "Web app".
// 4. Fill in the deployment settings:
//    - Description: (Optional) e.g., "Task Delegator v3.6 - Batch Processing Fix"
//    - Execute as: ME (your.email@domain.com) <-- CRITICAL
//    - Who has access: Anyone <-- CRITICAL FOR THE APP TO WORK
// 5. Click "Deploy".
// 6. AUTHORIZE the script's permissions if asked.
// 7. COPY the new "Web app URL" and PASTE it into the `SCRIPT_URL` constant below.
//
// --- STEP 3: CREATE THE AUTOMATED TRIGGER (ONE-TIME SETUP) ---
// This is the most important new step. It creates a timer that will process the queue.
// 1. In the Apps Script editor, go to the left-hand menu and click the "Triggers" icon (looks like a clock).
// 2. Click the "+ Add Trigger" button (bottom right).
// 3. Configure the trigger with these exact settings:
//    - Choose which function to run: processQueue
//    - Choose which deployment should run: Head
//    - Select event source: Time-driven
//    - Select type of time-based trigger: Minutes timer
//    - Select minute interval: Every minute
//    - Failure notification settings: Notify me immediately (Recommended)
// 4. Click "Save".
// 5. You may be asked to authorize permissions again. This is normal; please approve them.
//
// Your system is now set up! The first time a user submits a task, a new sheet named
// "RequestQueue" will be created automatically. You do not need to create it manually.
//
// --- START OF GOOGLE APPS SCRIPT (Code.gs) ---
//
var MASTER_SHEET_ID = '1XTc_cmSnyfAOduFTqpjnbAI8-dMgNz2LCBv_8DFTeNs';
var DELEGATION_SHEET_ID = '1Znih9FtcuqTJSJtS7peoBuJ8TijOXQl9eiWrGcAmXAg';
var TASK_ATTACHMENTS_FOLDER_NAME = "Task Attachments"; // Define the folder name for uploads
var QUEUE_SHEET_NAME = 'RequestQueue'; // Name for the new request queue sheet
var BATCH_SIZE = 50; // Process up to 50 requests per run to avoid timeouts

// -----------------------------------------------------------------------------
// ** NEW ** Configuration for each sheet, including which spreadsheet it's in.
// This is a more robust way to handle multiple spreadsheets.
// -----------------------------------------------------------------------------
var SHEET_CONFIG = {
  // --- Sheets in the Master/Checklist Spreadsheet ---
  'Task': {
    spreadsheetId: MASTER_SHEET_ID,
    matchColumn: 'Task'
  },
  'Master Data': {
    spreadsheetId: MASTER_SHEET_ID,
    matchColumn: 'Task ID',
    matchColumnIndex: 1,
    readOnlyColumns: ['pc']
  },
  'Done Task Status': {
    spreadsheetId: MASTER_SHEET_ID,
    matchColumn: 'Task ID'
  },
  // --- Sheets in the Delegation Spreadsheet ---
  'Working Task Form': {
    spreadsheetId: DELEGATION_SHEET_ID,
    matchColumn: 'Task ID',
    matchColumnIndex: 8,
    readOnlyColumns: ['delegate email']
  }
};

// -----------------------------------------------------------------------------
// Gets a Google Drive folder by name, creating it if it doesn't exist in root.
// -----------------------------------------------------------------------------
function getOrCreateFolder(folderName) {
  var folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  } else {
    return DriveApp.createFolder(folderName);
  }
}

// -----------------------------------------------------------------------------
// Finds the row number for a given value in a specific column.
// -----------------------------------------------------------------------------
function findRow(sheet, config, matchValue) {
  var matchColumnName = config.matchColumn;
  var matchColumnIndex; 

  if (config.matchColumnIndex) {
    matchColumnIndex = config.matchColumnIndex - 1;
    if (matchColumnIndex < 0 || matchColumnIndex >= sheet.getLastColumn()) {
      throw new Error("Configuration error: Invalid matchColumnIndex '" + config.matchColumnIndex + "' for sheet '" + sheet.getName() + "'.");
    }
  } else {
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var normalizedHeaders = headers.map(function(h) { return String(h).trim().toLowerCase(); });
    var normalizedMatchColumnName = String(matchColumnName).trim().toLowerCase();
    matchColumnIndex = normalizedHeaders.indexOf(normalizedMatchColumnName);
    if (matchColumnIndex === -1) {
      throw new Error("Could not find match column '" + matchColumnName + "' in sheet '" + sheet.getName() + "'. Headers: [" + headers.join(", ") + "]");
    }
  }

  var data = sheet.getDataRange().getValues();
  var normalizedMatchValue = String(matchValue).trim().toLowerCase();
  for (var i = data.length - 1; i > 0; i--) {
    if (data[i].length > matchColumnIndex) {
        var cellValue = String(data[i][matchColumnIndex]).trim().toLowerCase();
        if (cellValue !== "" && cellValue === normalizedMatchValue) {
          return i + 1; // 1-indexed row number.
        }
    }
  }
  return -1; // Not found
}

// -----------------------------------------------------------------------------
// ** MODIFIED: doPost now correctly queues some requests and processes others directly. **
// This makes the UI extremely fast for "Mark Done", while guaranteeing "Undone" works correctly.
// -----------------------------------------------------------------------------
function doPost(e) {
  try {
    var request = JSON.parse(e.postData.contents);
    
    // --- QUEUE LOGIC ---
    // For 'Done Task Status', only queue 'create' and 'batchCreate' actions for performance.
    // We exit early from the function if an action is queued.
    if (request.sheetName === 'Done Task Status' && (request.action === 'create' || request.action === 'batchCreate')) {
      console.log("Queueing request for 'Done Task Status'. Action: " + request.action);
      
      var masterDoc = SpreadsheetApp.openById(MASTER_SHEET_ID);
      var queueSheet = masterDoc.getSheetByName(QUEUE_SHEET_NAME);
      if (!queueSheet) {
        queueSheet = masterDoc.insertSheet(QUEUE_SHEET_NAME);
        console.log("Created 'RequestQueue' sheet as it did not exist.");
      }
      
      // Append the entire request payload as a string to the queue
      queueSheet.appendRow([e.postData.contents]);

      return ContentService
        .createTextOutput(JSON.stringify({ 
          status: 'success', 
          message: 'Request queued successfully.' 
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // --- SYNCHRONOUS LOGIC ---
    // All other requests fall through and are processed immediately.
    // This now correctly handles 'delete' for 'Done Task Status'.
    console.log("Processing synchronously. Action: " + request.action + " for sheet: " + (request.sheetName || 'N/A'));
    _processSingleRequest(request); // This function will throw an error on failure
    
    return ContentService
      .createTextOutput(JSON.stringify({
        status: 'success',
        message: 'Request processed successfully.'
      }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    // This catch block handles errors from both queuing and synchronous processing.
    console.error("Error in doPost: " + err.toString() + "\nStack: " + err.stack);
    return ContentService
      .createTextOutput(JSON.stringify({ 
        status: 'error', 
        message: 'Script failed: ' + err.message 
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// -----------------------------------------------------------------------------
// ** NEW: BATCH PROCESSING that processes a fixed number of items per run to prevent timeouts. **
// ** It reads a batch, prepares rows, writes them, and then deletes ONLY the processed rows from the queue. **
// -----------------------------------------------------------------------------
function processQueue() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) { // Try to get a lock for 10 seconds.
    console.log("Could not obtain lock. Another queue process is likely running.");
    return;
  }
  
  var masterDoc;
  var queueSheet;
  var numRowsToProcess = 0;

  try {
    masterDoc = SpreadsheetApp.openById(MASTER_SHEET_ID);
    queueSheet = masterDoc.getSheetByName(QUEUE_SHEET_NAME);
    
    if (!queueSheet || queueSheet.getLastRow() === 0) {
      console.log("RequestQueue sheet is empty. Exiting.");
      return; // The 'finally' block will still run to release the lock.
    }

    var totalRowsInQueue = queueSheet.getLastRow();
    numRowsToProcess = Math.min(totalRowsInQueue, BATCH_SIZE);

    console.log("Optimized Batch Processing: Starting for " + numRowsToProcess + " of " + totalRowsInQueue + " item(s).");

    // Get only the rows for the current batch from the top of the sheet
    var dataRange = queueSheet.getRange(1, 1, numRowsToProcess, 1);
    var data = dataRange.getValues();

    // --- BATCH PROCESSING LOGIC ---
    var doneTaskRows = [];
    var historyRows = [];

    // Get target sheets and headers once to avoid repeated calls in the loop
    var doneTaskSheet = masterDoc.getSheetByName('Done Task Status');
    var historySheet = masterDoc.getSheetByName('History');
    
    if (!doneTaskSheet) {
      console.error("FATAL: 'Done Task Status' sheet not found. Aborting queue processing.");
      numRowsToProcess = 0; // Prevent row deletion on critical error
      return;
    }
    if (!historySheet) {
      console.error("FATAL: 'History' sheet not found. Aborting queue processing.");
      numRowsToProcess = 0; // Prevent row deletion on critical error
      return;
    }

    var doneTaskHeaders = doneTaskSheet.getRange(1, 1, 1, doneTaskSheet.getLastColumn()).getValues()[0];
    var normalizedDoneTaskHeaders = doneTaskHeaders.map(function(h) { return String(h).trim().toLowerCase(); });

    // 1. PREPARE DATA: Loop through the batch of queued items.
    data.forEach(function(row, index) {
      if (row[0]) { // Ensure row is not empty
        try {
          var request = JSON.parse(row[0]);

          // Safety check: This queue is only for 'Done Task Status' creation
          if (request.sheetName !== 'Done Task Status' || (request.action !== 'create' && request.action !== 'batchCreate')) {
            console.warn("Skipping unexpected request type in queue: " + row[0]);
            return; // Skips to the next item in forEach
          }
          
          // --- Handle Single 'create' or 'batchCreate' from the queue ---
          var newDatas = (request.action === 'batchCreate') ? (request.newDatas || []) : [request.newData];
          var historyRecords = (request.action === 'batchCreate') ? (request.historyRecords || []) : [request.historyRecord];
          
          for (var i = 0; i < newDatas.length; i++) {
            var currentData = newDatas[i];
            if (!currentData) continue;
            
            // --- Handle Attachment (only for single 'create' actions) ---
            if (request.attachment && request.action === 'create') {
                var decodedContent = Utilities.base64Decode(request.attachment.content);
                var blob = Utilities.newBlob(decodedContent, request.attachment.mimeType, request.attachment.fileName);
                var uploadFolder = getOrCreateFolder(TASK_ATTACHMENTS_FOLDER_NAME);
                var file = uploadFolder.createFile(blob);
                file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
                currentData['Attachment URL'] = file.getUrl();
            }

            // --- Prepare 'Done Task Status' row ---
            var normalizedNewData = {};
            for (var key in currentData) {
                normalizedNewData[String(key).trim().toLowerCase()] = currentData[key];
            }
            var newRow = normalizedDoneTaskHeaders.map(function(h) { return normalizedNewData[h] || ""; });
            doneTaskRows.push(newRow);
          }

          // --- Prepare 'History' rows ---
          for (var j = 0; j < historyRecords.length; j++) {
             var historyRecord = historyRecords[j];
             if (historyRecord) {
                historyRows.push([
                    new Date(),
                    historyRecord.systemType || 'N/A',
                    historyRecord.task || 'N/A',
                    historyRecord.changedBy || 'N/A',
                    historyRecord.change || 'N/A'
                ]);
            }
          }

        } catch (e) {
          console.error("Error preparing queued item #" + (index + 1) + ": " + e.toString() + "\nOriginal request: " + row[0]);
        }
      }
    });

    // 2. WRITE DATA IN BATCHES
    if (doneTaskRows.length > 0) {
      console.log("Writing " + doneTaskRows.length + " rows to 'Done Task Status'.");
      doneTaskSheet.getRange(doneTaskSheet.getLastRow() + 1, 1, doneTaskRows.length, doneTaskHeaders.length).setValues(doneTaskRows);
    }

    if (historyRows.length > 0) {
      console.log("Writing " + historyRows.length + " rows to 'History'.");
      historySheet.getRange(historySheet.getLastRow() + 1, 1, historyRows.length, 5).setValues(historyRows);
    }
    
    console.log("Batch processing complete for this run.");

  } catch (e) {
    console.error("A critical error occurred during processQueue execution: " + e.toString() + "\nStack: " + e.stack);
    numRowsToProcess = 0; // Prevent deletion of rows that may have caused an error
  } finally {
    // This is the key change: only delete the rows that were processed in this batch.
    if (queueSheet && numRowsToProcess > 0) {
      try {
        queueSheet.deleteRows(1, numRowsToProcess);
        console.log("Deleted " + numRowsToProcess + " processed rows from the queue.");
      } catch (deleteError) {
        console.error("Failed to delete processed rows from queue: " + deleteError.toString());
      }
    }
    
    lock.releaseLock();
    console.log("Queue processing finished for this run. Lock released.");
  }
}


// -----------------------------------------------------------------------------
// ** This helper contains the processing logic for SYNCHRONOUS requests. It's called directly by doPost. **
// -----------------------------------------------------------------------------
function _processSingleRequest(request) {
  var sheetName = String(request.sheetName || '').trim();
  if (!sheetName) throw new Error("Request is missing 'sheetName'.");

  var config = SHEET_CONFIG[sheetName];
  if (!config) throw new Error("Configuration for sheet '" + sheetName + "' not found.");
  if (!config.spreadsheetId) throw new Error("Config for sheet '" + sheetName + "' is missing 'spreadsheetId'.");

  console.log("Processing request for '" + sheetName + "'. Opening Spreadsheet ID: " + config.spreadsheetId);

  var doc = SpreadsheetApp.openById(config.spreadsheetId);
  var sheet = doc.getSheetByName(sheetName);
  if (!sheet) throw new Error("Sheet '" + sheetName + "' not found in spreadsheet (ID: " + config.spreadsheetId + ").");

  var masterDoc = SpreadsheetApp.openById(MASTER_SHEET_ID);
  var historySheet = masterDoc.getSheetByName('History');
  if (!historySheet) throw new Error("'History' sheet not found in Master spreadsheet.");

  if (request.attachment) {
    console.log("Processing attachment: " + request.attachment.fileName);
    var decodedContent = Utilities.base64Decode(request.attachment.content);
    var blob = Utilities.newBlob(decodedContent, request.attachment.mimeType, request.attachment.fileName);
    var uploadFolder = getOrCreateFolder(TASK_ATTACHMENTS_FOLDER_NAME);
    var file = uploadFolder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var attachmentUrl = file.getUrl();
    console.log("File uploaded. URL: " + attachmentUrl);
    if (request.newData) request.newData['Attachment URL'] = attachmentUrl;
  }

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  switch (request.action) {
    case 'create':
      var normalizedNewData = {};
      for (var key in request.newData) normalizedNewData[String(key).trim().toLowerCase()] = request.newData[key];
      var newRow = headers.map(function(h) { return normalizedNewData[String(h).trim().toLowerCase()] || ""; });
      sheet.appendRow(newRow);
      break;
      
    case 'batchCreate':
      if (!request.newDatas || !Array.isArray(request.newDatas)) throw new Error("batchCreate requires a 'newDatas' array.");
      var newRows = request.newDatas.map(function(newData) {
        var normalizedData = {};
        for (var key in newData) normalizedData[String(key).trim().toLowerCase()] = newData[key];
        return headers.map(function(h) { return normalizedData[String(h).trim().toLowerCase()] || ""; });
      });
      if(newRows.length > 0) {
        sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
      }
      break;

    case 'update':
      var rowToUpdate = findRow(sheet, config, request.matchValue);
      if (rowToUpdate !== -1) {
        var readOnlyColumns = (config.readOnlyColumns || []).map(function(c) { return String(c).trim().toLowerCase(); });
        var normalizedUpdatedData = {};
        for (var key in request.updatedData) normalizedUpdatedData[String(key).trim().toLowerCase()] = request.updatedData[key];
        headers.forEach(function(header, index) {
          var normalizedHeader = String(header).trim().toLowerCase();
          if (readOnlyColumns.indexOf(normalizedHeader) !== -1) {
            if (normalizedUpdatedData.hasOwnProperty(normalizedHeader)) console.log("Skipping read-only column '" + header + "'.");
            return;
          }
          if (normalizedUpdatedData.hasOwnProperty(normalizedHeader)) {
            sheet.getRange(rowToUpdate, index + 1).setValue(normalizedUpdatedData[normalizedHeader]);
          }
        });
      } else {
        throw new Error("Row with " + config.matchColumn + " = '" + request.matchValue + "' not found to update in '" + sheetName + "'.");
      }
      break;

    case 'delete':
      var rowToDelete = findRow(sheet, config, request.matchValue);
      if (rowToDelete !== -1) {
        sheet.deleteRow(rowToDelete);
      } else {
        console.log("Row with " + config.matchColumn + " = '" + request.matchValue + "' not found to delete in '" + sheetName + "'. No action taken.");
      }
      break;

    default:
      throw new Error("Invalid action: '" + request.action + "'.");
  }

  // LOG TO HISTORY (handles single and batch)
  var historyRecords = [];
  if (request.historyRecord) historyRecords.push(request.historyRecord);
  if (request.historyRecords && Array.isArray(request.historyRecords)) historyRecords = historyRecords.concat(request.historyRecords);
  
  if (historyRecords.length > 0) {
    var historyRows = historyRecords.map(function(rec) {
      return [ new Date(), rec.systemType || 'N/A', rec.task || 'N/A', rec.changedBy || 'N/A', rec.change || 'N/A' ];
    });
    historySheet.getRange(historySheet.getLastRow() + 1, 1, historyRows.length, historyRows[0].length).setValues(historyRows);
  }
}
// --- END OF GOOGLE APPS SCRIPT ---
// =================================================================================================


// --- HELPER FUNCTIONS (LOCAL) ---
const simpleHash = (str: string): string => {
    let hash = 0;
    if (str.length === 0) return '0';
    for (let i = 0; i < str.length; i++) {
        const char = str.charCodeAt(i);
        hash = (hash << 5) - hash + char;
        hash |= 0; // Convert to 32bit integer
    }
    return String(hash);
};

const sleep = (ms: number) => new Promise(resolve => setTimeout(resolve, ms));


// --- UI COMPONENTS ---

const RefreshControl: React.FC<{
    lastUpdated: Date | null;
    onRefresh: () => void;
    isRefreshing: boolean;
    isAdmin: boolean;
}> = ({ lastUpdated, onRefresh, isRefreshing, isAdmin }) => {
    const [timeAgo, setTimeAgo] = useState('');

    useEffect(() => {
        const formatTimeAgo = () => {
            if (!lastUpdated) {
                setTimeAgo('never');
                return;
            }
            const seconds = Math.floor((new Date().getTime() - lastUpdated.getTime()) / 1000);
            if (seconds < 5) {
                setTimeAgo('just now');
                return;
            }
            if (seconds < 60) {
                setTimeAgo(`${seconds} seconds ago`);
                return;
            }
            const minutes = Math.floor(seconds / 60);
            if (minutes < 60) {
                setTimeAgo(`${minutes} minute${minutes > 1 ? 's' : ''} ago`);
                return;
            }
            setTimeAgo(`on ${lastUpdated.toLocaleString()}`);
        };

        formatTimeAgo();
        const interval = setInterval(formatTimeAgo, 5000); // update every 5 seconds
        return () => clearInterval(interval);
    }, [lastUpdated]);

    return (
        <div className="refresh-control">
            <span className="last-updated-text" aria-live="polite">
                Last updated: {timeAgo}
            </span>
            {isAdmin && (
                <button
                    className="btn-refresh"
                    onClick={onRefresh}
                    disabled={isRefreshing}
                    aria-label="Refresh data"
                >
                    <svg className={isRefreshing ? 'spinning' : ''} xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor" viewBox="0 0 16 16">
                        <path fillRule="evenodd" d="M8 3a5 5 0 1 0 4.546 2.914.5.5 0 0 1 .908-.417A6 6 0 1 1 8 2z"/>
                        <path d="M8 4.466V.534a.25.25 0 0 1 .41-.192l2.36 1.966c.12.1.12.284 0 .384L8.41 4.658A.25.25 0 0 1 8 4.466"/>
                    </svg>
                </button>
            )}
        </div>
    );
};

const LoginPanel: React.FC<{ onLoginSuccess: (user: AuthenticatedUser) => void }> = ({ onLoginSuccess }) => {
    const [step, setStep] = useState<'email' | 'password'>('email');
    const [email, setEmail] = useState('');
    const [password, setPassword] = useState('');
    const [error, setError] = useState('');
    const [adminUser, setAdminUser] = useState<UserAuth | null>(null);

    const handleEmailSubmit = async (e: React.FormEvent) => {
        e.preventDefault();
        setError('');

        const sheetId = '1XTc_cmSnyfAOduFTqpjnbAI8-dMgNz2LCBv_8DFTeNs';
        const usersSheetName = 'Users';
        
        const usersUrl = `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?tqx=out:csv&sheet=${usersSheetName}`;

        try {
            const usersResponse = await fetch(usersUrl);
            if (!usersResponse.ok) throw new Error('Failed to fetch user data. The "Users" sheet might be private or may not exist.');

            // Parse Users sheet
            const usersCsvText = await usersResponse.text();
            const userRows = usersCsvText.trim().replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n').slice(1);
            const csvSplitter = /,(?=(?:(?:[^"]*"){2})*[^"]*$)/;
            const users: UserAuth[] = userRows.map(row => {
                 const fields = row.split(csvSplitter);
                 const [mailId, role, password] = fields.map(field => field.trim().replace(/^"|"$/g, ''));
                 return { mailId, role, password };
            }).filter(u => u.mailId); // Ensure mailId is not empty
            
            const lowerCaseEmail = email.toLowerCase();

            const foundUserInUsersSheet = users.find(u => u.mailId.toLowerCase() === lowerCaseEmail);

            if (foundUserInUsersSheet) {
                if (foundUserInUsersSheet.role === 'Admin') {
                    // Admin found, ask for password
                    setAdminUser(foundUserInUsersSheet);
                    setStep('password');
                } else {
                    // Any other role in Users sheet is logged in without password
                    onLoginSuccess({ mailId: foundUserInUsersSheet.mailId, role: foundUserInUsersSheet.role });
                }
            } else {
                // Not found in the Users sheet, treat as a normal user.
                onLoginSuccess({ mailId: email, role: 'User' });
            }
        } catch (err: any) {
            console.error('Login error:', err);
            setError(err.message || 'An error occurred during login.');
        }
    };

    const handlePasswordSubmit = (e: React.FormEvent) => {
        e.preventDefault();
        if (!adminUser) return;
        
        if (adminUser.password === password) {
            onLoginSuccess({ mailId: adminUser.mailId, role: adminUser.role });
        } else {
            setError('Incorrect password.');
        }
    };
    
    const handleGoBack = () => {
        setStep('email');
        setError('');
        setPassword('');
        setAdminUser(null);
    }

    if (step === 'password') {
        return (
            <div className="login-container">
                <div className="login-panel">
                    <h1>Admin Login</h1>
                    <p>Enter password for <strong>{email}</strong></p>
                    <form className="login-form" onSubmit={handlePasswordSubmit}>
                        <input
                            type="password"
                            placeholder="Password"
                            value={password}
                            onChange={e => setPassword(e.target.value)}
                            required
                            autoFocus
                            aria-label="Password"
                        />
                        {error && <div className="login-error" role="alert">{error}</div>}
                        <div className="login-actions">
                            <button type="button" className="btn btn-secondary" onClick={handleGoBack}>Back</button>
                            <button type="submit" className="btn btn-primary">Sign In</button>
                        </div>
                    </form>
                </div>
            </div>
        );
    }

    return (
        <div className="login-container">
            <div className="login-panel">
                <h1>Welcome</h1>
                <p>Please enter your email to sign in</p>
                <form className="login-form" onSubmit={handleEmailSubmit}>
                    <input
                        type="email"
                        placeholder="Email"
                        value={email}
                        onChange={e => setEmail(e.target.value)}
                        required
                        aria-label="Email Address"
                    />
                    {error && <div className="login-error" role="alert">{error}</div>}
                    <button type="submit" className="btn btn-primary">Continue</button>
                </form>
            </div>
        </div>
    );
};

// --- MAIN APP COMPONENT ---
const App = () => {
    const [authenticatedUser, setAuthenticatedUser] = useLocalStorage<AuthenticatedUser | null>('task-delegator-auth', null);
    const isAdmin = authenticatedUser?.role === 'Admin';

    const [mode, setMode] = useState<AppMode>('dashboard');
    
    // Data State
    const [people, setPeople] = useState<Person[]>([]);
    const [tasks, setTasks] = useLocalStorage<Task[]>('task-delegator-tasks', []);
    const [checklists, setChecklists] = useState<Checklist[]>([]);
    const [masterTasks, setMasterTasks] = useState<MasterTask[]>([]);
    const [delegationTasks, setDelegationTasks] = useState<DelegationTask[]>([]);
    const [allDashboardTasks, setAllDashboardTasks] = useState<DashboardTask[]>([]);
    const [attendanceData, setAttendanceData] = useState<AttendanceData[]>([]);
    const [dailyAttendanceData, setDailyAttendanceData] = useState<DailyAttendance[]>([]);
    const [holidays, setHolidays] = useState<Holiday[]>([]);
    const [taskHistory, setTaskHistory] = useState<TaskHistory[]>([]);

    // Loading and Error State
    const [isLoadingPeople, setIsLoadingPeople] = useState(true);
    const [peopleError, setPeopleError] = useState<string | null>(null);
    const [checklistsError, setChecklistsError] = useState<string | null>(null);
    const [masterTasksError, setMasterTasksError] = useState<string | null>(null);
    const [delegationTasksError, setDelegationTasksError] = useState<string | null>(null);
    const [allDashboardTasksError, setAllDashboardTasksError] = useState<string | null>(null);
    const [attendanceError, setAttendanceError] = useState<string | null>(null);
    const [dailyAttendanceError, setDailyAttendanceError] = useState<string | null>(null);
    const [holidaysError, setHolidaysError] = useState<string | null>(null);
    const [taskHistoryError, setTaskHistoryError] = useState<string | null>(null);

    // General state
    const [lastUpdated, setLastUpdated] = useState<Date | null>(null);
    const [isRefreshing, setIsRefreshing] = useState(false);
    
    // --- ACTION REQUIRED (STEP 2 from instructions at top of file) ---
    // PASTE YOUR NEW DEPLOYMENT URL HERE.
    // The URL you get after deploying the script from the MASTER workbook's script editor.
    const SCRIPT_URL = "https://script.google.com/macros/s/AKfycbxBeTwQyh7wbX4DuUNEh-3OeixyCjmGnHatgn3Dff14Dn5NVI5wC-DUzz-aFxk5p2s6/exec";
    
    const DELEGATION_FORM_URL = "https://script.google.com/macros/s/AKfycbxbNrwhuhCxoTQlXwgN2XAClofwGIUe2-H2QpqMX8-KUN6-wgczXjW1NSl-NvhVOf3g/exec";

    // Enforce view for non-admin roles
    useEffect(() => {
        if (authenticatedUser && authenticatedUser.role !== 'Admin') {
            setMode('dashboard');
        }
    }, [authenticatedUser]);

    const fetchData = useCallback(async (isInitialLoad = false) => {
        if (!isInitialLoad && isRefreshing) return;

        setIsRefreshing(true);
        if (isInitialLoad) {
            setIsLoadingPeople(true);
        }
        setPeopleError(null);
        setChecklistsError(null);
        setMasterTasksError(null);
        setDelegationTasksError(null);
        setAllDashboardTasksError(null);
        setAttendanceError(null);
        setDailyAttendanceError(null);
        setHolidaysError(null);
        setTaskHistoryError(null);


        const sheetId = '1XTc_cmSnyfAOduFTqpjnbAI8-dMgNz2LCBv_8DFTeNs';
        const delegationSheetId = '1Znih9FtcuqTJSJtS7peoBuJ8TijOXQl9eiWrGcAmXAg';
        const masterDashboardSheetId = '1tlHs1iKCEnhrNAZRMy8YiTMeLGtyd5QWJ09okevio_M';

        // URLs for fetching data
        const peopleUrl = `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?tqx=out:csv&sheet=Employee%20Data&range=B:Q`;
        const checklistUrl = `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?tqx=out:csv&sheet=Task&range=A:J`;
        const masterUrl = `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?tqx=out:csv&sheet=${encodeURIComponent('Master Data')}&range=A:N`;
        const delegationUrl = `https://docs.google.com/spreadsheets/d/${delegationSheetId}/gviz/tq?tqx=out:csv&sheet=Working%20Task%20Form`;
        const leavesUrl = `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?tqx=out:csv&sheet=Leaves&tq=${encodeURIComponent('SELECT J, U WHERE U IS NOT NULL')}`;
        const dailyAttendanceUrl = `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?tqx=out:csv&sheet=Leaves&tq=${encodeURIComponent('SELECT P, Q, R, U WHERE R IS NOT NULL AND P IS NOT NULL')}`;
        const holidaysUrl = `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?tqx=out:csv&sheet=Leaves&tq=${encodeURIComponent('SELECT S, T WHERE T IS NOT NULL')}`;
        const historyUrl = `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?tqx=out:csv&sheet=History`;

        const fetchWithHandling = async (url: string, processor: (csv: string) => void) => {
            const response = await fetch(url);
            if (!response.ok) throw new Error(`Network response was not ok. Status: ${response.status}`);
            const contentType = response.headers.get('content-type');
            if (!contentType || !contentType.includes('text/csv')) throw new Error('Received non-CSV response. The Google Sheet may be private or incorrectly named.');
            const csvText = await response.text();
            processor(csvText);
        };

        try {
            // --- SEQUENTIAL FETCHING TO PREVENT RATE-LIMITING ---

            // 1. Fetch People
            try {
                await fetchWithHandling(peopleUrl, (csvText) => {
                    const parsedData = robustCsvParser(csvText);
                    const parsedPeople: Person[] = parsedData.map(fields => {
                        let name = (fields[0] || '').trim(); // Column B for Name
                        const email = (fields[4] || '').trim(); // Column F for Email
                        const photoUrl = (fields[15] || '').trim(); // Column Q for Photo URL
                        if (!name && email) {
                            const namePart = email.split('@')[0];
                            name = namePart.replace(/[._-]/g, ' ').split(' ').map(word => word.charAt(0).toUpperCase() + word.slice(1)).join(' ');
                        }
                        return { name, email, photoUrl };
                    }).filter(p => p.name && p.email);
                    
                    if (parsedPeople.length === 0) {
                        setPeopleError('No valid employee data found. Please ensure the "Employee Data" sheet has columns for Name (B) and especially Email (F) populated.');
                        setPeople([]);
                    } else {
                        setPeople(parsedPeople);
                    }
                });
            } catch (err: any) {
                console.error("Data fetch error (People):", err);
                setPeopleError('Failed to load team. Please make sure the "Employee Data" Google Sheet is public (set Share > General access > Anyone with the link) AND published to the web (File > Share > Publish to web).');
            }
            await sleep(250);

            // 2. Fetch Checklists
            try {
                await fetchWithHandling(checklistUrl, (csvText) => {
                     const parsedData = robustCsvParser(csvText);
                     const importedChecklists: Checklist[] = parsedData.filter(fields => fields.length > 0 && fields[0] && fields[0].trim() !== '').map((fields, index) => ({
                         id: `sheet-item-${simpleHash((fields[0] || '') + '-' + (fields[1] || '') + '-' + index)}`,
                         task: fields[0] || '',
                         doer: fields[1] || '',
                         frequency: fields[2] || 'D',
                         date: fields[3] || '',
                         buddy: fields[4] || '',
                         secondBuddy: fields[5] || '',
                     }));
                     setChecklists(importedChecklists);
                });
            } catch (err: any) {
               console.error("Data fetch error (Checklists):", err);
               setChecklistsError('Failed to load Task List. Please ensure the "Task" sheet in the main Google Sheet is public and published to the web.');
            }
            await sleep(250);
            
            // 3. Fetch Master Tasks
            try {
                await fetchWithHandling(masterUrl, (csvText) => {
                    const parsedData = robustCsvParser(csvText);
                    const importedMasterTasks: MasterTask[] = parsedData
                        .filter(fields => fields.length > 2 && fields[2] && fields[2].trim() !== '')
                        .map((fields, index) => ({
                            id: `master-task-${fields[0] || `row-${index}`}`, taskId: fields[0] || '', plannedDate: fields[1] || '',
                            actualDate: fields[7] || '', taskDescription: fields[2] || '', doer: fields[3] || '',
                            originalDoer: fields[11] || '', frequency: fields[4] || '', pc: fields[9] || '', status: fields[13] || '',
                        }));
                    setMasterTasks(importedMasterTasks);
                });
            } catch (err: any) {
                console.error("Data fetch error (Master Tasks):", err);
                setMasterTasksError('Failed to load Master Tasks. Please ensure the "Master Data" sheet in the main Google Sheet is public and published to the web.');
            }
            await sleep(250);

            // 4. Fetch Delegation Tasks
            try {
                await fetchWithHandling(delegationUrl, (csvText) => {
                     const parsedData = robustCsvParser(csvText);
                     const allDelegationTasks: (DelegationTask & { status?: string })[] = parsedData.map((fields, index) => ({
                        id: `delegation-${fields[7] || `row-${index}`}`, timestamp: fields[0] || '', assignee: fields[1] || '',
                        task: fields[2] || '', plannedDate: fields[3] || '', assignerEmail: fields[4] || '', assigner: fields[5] || '',
                        delegateEmail: fields[6] || '', taskId: fields[7] || '', actualDate: fields[8] || '', status: fields[9] || '',
                     }));
                     const importedDelegationTasks = allDelegationTasks.filter(task => {
                        const hasTask = task.task && task.task.trim() !== '';
                        const isCancelled = task.status && task.status.toLowerCase() === 'cancel';
                        return hasTask && !isCancelled;
                     });
                     setDelegationTasks(importedDelegationTasks);
                });
            } catch (err: any) {
                console.error("Data fetch error (Delegation Tasks):", err);
                setDelegationTasksError('Failed to load Delegation Tasks. Please make sure the "Working Task Form" Google Sheet is public (set Share > General access > Anyone with the link) AND published to the web (File > Share > Publish to web).');
            }
            await sleep(250);
            
            // 5. Fetch All Dashboard Data
            try {
                const sources = [
                    { name: 'Checklist', id: sheetId, sheet: 'DB' },
                    { name: 'Delegation', id: delegationSheetId, sheet: 'DB' },
                    { name: 'Master', id: masterDashboardSheetId, sheet: 'Master' }
                ];

                const parseDashboardTaskData = (csvText: string, source: string): DashboardTask[] => {
                    const parsedData = robustCsvParser(csvText);
                    return parsedData
                        .filter(fields => fields.length > 1 && fields[1] && fields[1].trim() !== '')
                        .map((fields, index) => ({
                            id: `${source}-task-${fields[1] || `row-${index}`}`, timestamp: fields[0] || '', taskId: fields[1] || '', task: fields[2] || '', stepCode: fields[3] || '',
                            planned: fields[4] || '', actual: (fields[5] || '').trim(), name: fields[6] || '', link: fields[7] || '', forPc: fields[8] || '',
                            systemType: fields[9] || '', userName: (fields[14] || '').trim(), userEmail: (fields[15] || '').trim(),
                            photoUrl: (fields[16] || '').trim(), attachmentUrl: (fields[17] || '').trim(),
                        }));
                };

                const allTasks: DashboardTask[] = [];
                for (const sourceInfo of sources) {
                    const url = `https://docs.google.com/spreadsheets/d/${sourceInfo.id}/gviz/tq?tqx=out:csv&sheet=${encodeURIComponent(sourceInfo.sheet)}`;
                    try {
                        const res = await fetch(url);
                        if (!res.ok) throw new Error(`Network error fetching sheet: ${sourceInfo.name} (Status: ${res.status}). Ensure the sheet is public.`);
                        const contentType = res.headers.get('content-type');
                        if (!contentType || !contentType.includes('text/csv')) throw new Error(`Received a non-CSV response for sheet "${sourceInfo.name}". Please ensure it is published to the web and the name is spelled correctly.`);
                        const csv = await res.text();
                        allTasks.push(...parseDashboardTaskData(csv, sourceInfo.name.toLowerCase()));
                    } catch (err) {
                        console.error(`Failed to fetch or process sheet "${sourceInfo.name}" from URL: ${url}`, err);
                        throw err; 
                    }
                    await sleep(250); // Delay between each dashboard source
                }
                setAllDashboardTasks(allTasks);
            } catch (err: any) {
                console.error("Data fetch error (Dashboard/MIS Tasks):", err);
                setAllDashboardTasksError('Failed to load Dashboard tasks. Please ensure the "Checklist", "Delegation", and "Master" sheets in their respective Google Sheets are public and published to the web.');
            }
            await sleep(250);

            // 6. Fetch Attendance
            try {
                await fetchWithHandling(leavesUrl, (csvText) => {
                    const parsedData: AttendanceData[] = robustCsvParser(csvText).map(fields => ({
                        email: fields[1] || '',
                        daysPresent: !isNaN(parseFloat(fields[0])) ? parseFloat(fields[0]) : 0,
                    })).filter(item => item.email);
                    setAttendanceData(parsedData);
                });
            } catch (err: any) {
                console.error("Data fetch error (Attendance):", err);
                setAttendanceError('Failed to load Attendance Data. Please ensure the "Leaves" sheet in the main Google Sheet is public and published to the web.');
            }
            await sleep(250);
            
            // 7. Fetch Daily Attendance
            try {
                await fetchWithHandling(dailyAttendanceUrl, (csvText) => {
                    const parsedData: DailyAttendance[] = robustCsvParser(csvText).map(fields => ({
                        date: (fields[0] || '').trim(), status: (fields[1] || '').trim(),
                        name: (fields[2] || '').trim(), email: (fields[3] || '').trim().toLowerCase(),
                    })).filter(item => item.name && item.date && item.status);
                    setDailyAttendanceData(parsedData);
                });
            } catch (err: any) {
                 console.error("Data fetch error (Daily Attendance):", err);
                 setDailyAttendanceError('Failed to load Daily Attendance. Please ensure columns P (Date), Q (Status), R (Name), and U (Email) in the "Leaves" sheet are correctly formatted and the sheet is public.');
            }
            await sleep(250);
            
            // 8. Fetch Holidays
            try {
                await fetchWithHandling(holidaysUrl, (csvText) => {
                    const parsedData: Holiday[] = robustCsvParser(csvText).map(fields => ({
                        name: (fields[0] || 'Holiday').trim(), date: (fields[1] || '').trim(),
                    })).filter(item => item.date);
                    setHolidays(parsedData);
                });
            } catch (err: any) {
                console.error("Data fetch error (Holidays):", err);
                setHolidaysError('Failed to load Holidays. Please ensure columns S (Name) and T (Date) in the "Leaves" sheet are correctly formatted and the sheet is public.');
            }
            await sleep(250);
            
            // 9. Fetch History
            try {
                await fetchWithHandling(historyUrl, (csvText) => {
                    const parsedData: TaskHistory[] = robustCsvParser(csvText).map(fields => ({
                        timestamp: (fields[0] || '').trim(), systemType: (fields[1] || '').trim(),
                        task: (fields[2] || '').trim(), changedBy: (fields[3] || '').trim(),
                        change: (fields[4] || '').trim(),
                    })).filter(item => item.timestamp);
                    setTaskHistory(parsedData);
                });
            } catch (err: any) {
                 console.error("Data fetch error (History):", err);
                 setTaskHistoryError('Failed to load Task History. Please ensure the "History" sheet in the main Google Sheet is public and published to the web.');
            }

        } catch (err) {
            console.error("An unexpected error occurred during data fetch:", err);
        } finally {
            setLastUpdated(new Date());
            setIsRefreshing(false);
            if (isInitialLoad) {
                setIsLoadingPeople(false);
            }
        }
    }, [isRefreshing]);

    // Initial load and auto-refresh timer
    useEffect(() => {
        fetchData(true);
        const refreshInterval = setInterval(() => fetchData(false), 60000); // 60 seconds
        return () => clearInterval(refreshInterval);
    // eslint-disable-next-line react-hooks/exhaustive-deps
    }, []);

    // --- Google Sheet Communication ---
    const postToGoogleSheet = async (data: Record<string, any>) => {
        if (SCRIPT_URL.includes("PASTE_YOUR_NEW_WEB_APP_URL_HERE")) {
            const errorMessage = "The application is not configured. Please follow the instructions at the top of index.tsx to add your Google Apps Script Web App URL.";
            alert(errorMessage);
            throw new Error(errorMessage);
        }

        if (!authenticatedUser) {
            alert("Authentication error. Please log in again.");
            throw new Error("No authenticated user.");
        }
        
        try {
            const response = await fetch(SCRIPT_URL, {
                method: 'POST',
                headers: {
                  'Content-Type': 'text/plain',
                },
                body: JSON.stringify(data),
                // mode: 'no-cors' is removed to allow reading the response
            });

            if (!response.ok) {
                // Try to get more info from the response body if it's not OK
                const errorText = await response.text();
                throw new Error(`Network error: ${response.status} ${response.statusText}. Response: ${errorText}`);
            }
            
            const result = await response.json();

            if (result.status === 'error') {
                // This catches errors reported by our script's JSON response
                throw new Error(result.message || 'An unknown script error occurred.');
            }
            
            return result;

        } catch (error) {
            console.error("Error communicating with Google Sheet:", error);
            // Re-throw the error so the calling function's catch block can handle it
            if (error instanceof Error) {
                 // Just rethrow the specific error
                 throw error;
            }
            throw new Error("An unknown network or parsing error occurred.");
        }
    };


    const handleManualRefresh = () => fetchData(false);
    const handleLogout = () => setAuthenticatedUser(null);
    
    useEffect(() => {
        if (mode !== 'delegation') {
            // This is a simple way to reset state, could be more granular
        }
    }, [mode]);


    if (!authenticatedUser) {
        return <LoginPanel onLoginSuccess={setAuthenticatedUser} />;
    }

    const containerClass = mode === 'dashboard' ? 'container-dashboard' : 'container';

    return (
        <>
            <header>
                <div className="header-left">
                    <h1>Task Delegator</h1>
                    {isAdmin && (
                         <div className="mode-switcher">
                            <button onClick={() => setMode('dashboard')} className={mode === 'dashboard' ? 'active' : ''}>Task Dashboard</button>
                            <button onClick={() => setMode('checklist')} className={mode === 'checklist' ? 'active' : ''}>Checklist</button>
                            <button onClick={() => setMode('delegation')} className={mode === 'delegation' ? 'active' : ''}>Delegation</button>
                        </div>
                    )}
                </div>
                <div className="header-controls">
                     <div className="header-user-info">
                        <span>{authenticatedUser.mailId}</span>
                        <span className="user-role">{authenticatedUser.role}</span>
                        <button className="btn btn-logout" onClick={handleLogout}>Logout</button>
                    </div>
                    <RefreshControl 
                        lastUpdated={lastUpdated}
                        isRefreshing={isRefreshing}
                        onRefresh={handleManualRefresh}
                        isAdmin={isAdmin}
                    />
                </div>
            </header>
            <div className={containerClass}>
               <main>
                    {mode === 'dashboard' ? (
                        <TaskDashboardSystem
                            dashboardTasks={allDashboardTasks}
                            misTasks={allDashboardTasks}
                            isRefreshing={isRefreshing}
                            dashboardTasksError={allDashboardTasksError}
                            misTasksError={allDashboardTasksError}
                            authenticatedUser={authenticatedUser}
                            postToGoogleSheet={postToGoogleSheet}
                            fetchData={fetchData}
                            people={people}
                            attendanceData={attendanceData}
                            dailyAttendanceData={dailyAttendanceData}
                            holidays={holidays}
                            taskHistory={taskHistory}
                        />
                    ) : (isAdmin && mode === 'delegation') ? (
                        <DelegationSystem 
                            people={people}
                            delegationTasks={delegationTasks}
                            setDelegationTasks={setDelegationTasks}
                            authenticatedUser={authenticatedUser}
                            postToGoogleSheet={postToGoogleSheet}
                            fetchData={fetchData}
                            delegationFormUrl={DELEGATION_FORM_URL}
                            delegationTasksError={delegationTasksError}
                            isRefreshing={isRefreshing}
                        />
                    ) : (isAdmin && mode === 'checklist') ? (
                        <ChecklistSystem
                            isAdmin={isAdmin}
                            people={people}
                            checklists={checklists}
                            setChecklists={setChecklists}
                            masterTasks={masterTasks}
                            setMasterTasks={setMasterTasks}
                            tasks={tasks}
                            setTasks={setTasks}
                            authenticatedUser={authenticatedUser}
                            postToGoogleSheet={postToGoogleSheet}
                            fetchData={fetchData}
                            checklistsError={checklistsError}
                            masterTasksError={masterTasksError}
                            isRefreshing={isRefreshing}
                        />
                    ) : (
                        // Failsafe for non-admins if mode is not 'dashboard', or for admins with an invalid mode
                        <TaskDashboardSystem
                            dashboardTasks={allDashboardTasks}
                            misTasks={allDashboardTasks}
                            isRefreshing={isRefreshing}
                            dashboardTasksError={allDashboardTasksError}
                            misTasksError={allDashboardTasksError}
                            authenticatedUser={authenticatedUser}
                            postToGoogleSheet={postToGoogleSheet}
                            fetchData={fetchData}
                            people={people}
                            attendanceData={attendanceData}
                            dailyAttendanceData={dailyAttendanceData}
                            holidays={holidays}
                            taskHistory={taskHistory}
                        />
                    )}
                </main>
            </div>
        </>
    );
};

const container = document.getElementById('root');
if (container) {
    const root = createRoot(container);
    root.render(<App />);
}
