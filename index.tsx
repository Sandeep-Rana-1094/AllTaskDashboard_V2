import React, { useState, useEffect, useMemo, useCallback, useRef } from 'react';
import { createRoot } from 'react-dom/client';
import { ChecklistSystem } from './ChecklistSystem.tsx';
import { DelegationSystem } from './DelegationSystem.tsx';
import { TaskDashboardSystem } from './TaskDashboardSystem.tsx';
import { 
    Task, Checklist, MasterTask, DelegationTask, DashboardTask, AuthenticatedUser, UserAuth, Person, AttendanceData, DailyAttendance, AppMode, TaskHistory, Holiday, UserRole
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
        
        const usersUrl = `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?tqx=out:csv&sheet=${usersSheetName}&range=A:C`;
        const teamMapUrl = `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?tqx=out:csv&sheet=${usersSheetName}&tq=${encodeURIComponent('SELECT E, F WHERE E IS NOT NULL')}`;

        try {
            // Fetch both user roles and manager mappings concurrently
            const [usersResponse, teamMapResponse] = await Promise.all([
                fetch(usersUrl),
                fetch(teamMapUrl)
            ]);

            if (!usersResponse.ok) throw new Error('Failed to fetch user data. The "Users" sheet might be private or may not exist.');
            if (!teamMapResponse.ok) throw new Error('Failed to fetch team mapping data from the "Users" sheet.');

            const usersCsvText = await usersResponse.text();
            const teamMapCsvText = await teamMapResponse.text();
            
            const csvSplitter = /,(?=(?:(?:[^"]*"){2})*[^"]*$)/;

            // Parse Users for authentication (Admin/User)
            const userRows = usersCsvText.trim().replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n').slice(1);
            const users: UserAuth[] = userRows.map(row => {
                 const fields = row.split(csvSplitter);
                 const [mailId, role, password] = fields.map(field => field.trim().replace(/^"|"$/g, ''));
                 return { mailId, role, password };
            }).filter(u => u.mailId);

            // Parse Manager-Team mapping
            const teamMapRows = teamMapCsvText.trim().replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n').slice(1);
            const teamMap = new Map<string, string[]>();
            teamMapRows.forEach(row => {
                const fields = row.split(csvSplitter);
                const managerEmail = (fields[0] || '').trim().replace(/^"|"$/g, '').toLowerCase();
                const teamMemberEmail = (fields[1] || '').trim().replace(/^"|"$/g, '').toLowerCase();
                if (managerEmail && teamMemberEmail) {
                    if (!teamMap.has(managerEmail)) {
                        teamMap.set(managerEmail, []);
                    }
                    teamMap.get(managerEmail)!.push(teamMemberEmail);
                }
            });

            const lowerCaseEmail = email.toLowerCase();
            
            // --- LOGIN LOGIC: Manager > Admin > User ---

            // 1. Check if the user is a Manager
            if (teamMap.has(lowerCaseEmail)) {
                onLoginSuccess({
                    mailId: email,
                    role: 'Manager',
                    teamEmails: teamMap.get(lowerCaseEmail) || []
                });
                return;
            }

            // 2. Check if the user is in the auth list (Admin or User)
            const foundUserInUsersSheet = users.find(u => u.mailId.toLowerCase() === lowerCaseEmail);
            if (foundUserInUsersSheet) {
                if (foundUserInUsersSheet.role === 'Admin') {
                    // Admin found, ask for password
                    setAdminUser(foundUserInUsersSheet);
                    setStep('password');
                } else {
                    // Any other role in Users sheet is logged in without password
                    onLoginSuccess({ mailId: foundUserInUsersSheet.mailId, role: (foundUserInUsersSheet.role as UserRole) || 'User' });
                }
            } else {
                // 3. Not found anywhere, treat as a new standard user.
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
            onLoginSuccess({ mailId: adminUser.mailId, role: 'Admin' });
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
    const isManager = authenticatedUser?.role === 'Manager';

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
    const SCRIPT_URL = "https://script.google.com/macros/s/AKfycbwYN0BY-mpKiAmmj8zXF97dhukWH-m-q2fX6DTdfGVB6nHJIwBwhZ29ySaz1rNBr_Qv/exec";
    
    const DELEGATION_FORM_URL = "https://script.google.com/macros/s/AKfycbzTtcv7en0te98MUU8DeK_rPGrEW-xs2aH3EQCt4FqX2vIf-WPg9uFYtmG1WGY_8SlW/exec";

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
        const delegationSheetId = '18QL7gwHfWQyCCckTbwZr2eFgVnF55O8Vq9II_yQVzdU';
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
                        let name = (fields[0] || '').trim(); // Column B
                        const email = (fields[4] || '').trim(); // Column F
                        const status = (fields[6] || '').trim(); // Column H
                        const photoUrl = (fields[15] || '').trim(); // Column Q
                        if (!name && email) {
                            const namePart = email.split('@')[0];
                            name = namePart.replace(/[._-]/g, ' ').split(' ').map(word => word.charAt(0).toUpperCase() + word.slice(1)).join(' ');
                        }
                        return { name, email, photoUrl, status };
                    }).filter(p => {
                        // A person must have a name to be included in the system.
                        if (!p.name || p.name.trim() === '') {
                            return false;
                        }
                        // If status is explicitly 'Left' (case-insensitive), filter them out.
                        if ((p.status || '').trim().toLowerCase() === 'left') {
                            return false;
                        }
                        // Otherwise, keep them (this includes 'Active', empty statuses, and any other status).
                        return true;
                    });
                    
                    if (parsedPeople.length === 0) {
                        setPeopleError('No active employee data found. Please ensure the "Employee Data" sheet has names in Column B and that not all employees are marked as "Left" in Column H.');
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
                    const tasks: (DashboardTask | null)[] = parsedData
                        .filter(fields => fields.length > 1 && fields[1] && fields[1].trim() !== '')
                        .map((fields, index): DashboardTask | null => {
                            const baseId = fields[1] || `row-${index}`;
                            if (source === 'master') {
                                // Updated mapping based on user request:
                                // Name -> Column O (Index 14)
                                // Actual -> Column F (Index 5)
                                // Link -> Column H (Index 7)
                                const doerName = (fields[14] || '').trim(); 
                                if (!doerName) return null;
                                return {
                                    id: `master-task-${baseId}`,
                                    timestamp: fields[0] || '',
                                    taskId: fields[1] || '',
                                    task: fields[2] || '',
                                    stepCode: fields[3] || '',
                                    planned: fields[4] || '',
                                    actual: (fields[5] || '').trim(), 
                                    name: doerName,
                                    link: fields[7] || '', 
                                    forPc: fields[8] || '',
                                    systemType: fields[9] || '',
                                    userName: doerName,
                                    daysGiven: fields[19] || '',
                                    workDoneDay: fields[20] || '',
                                };
                            } else {
                                const userName = (fields[14] || '').trim();
                                if (!userName) return null;
                                return {
                                    id: `${source}-task-${baseId}`,
                                    timestamp: fields[0] || '',
                                    taskId: fields[1] || '',
                                    task: fields[2] || '',
                                    stepCode: fields[3] || '',
                                    planned: fields[4] || '',
                                    actual: (fields[5] || '').trim(),
                                    name: fields[6] || '',
                                    link: fields[7] || '',
                                    forPc: fields[8] || '',
                                    systemType: fields[9] || '',
                                    userName: userName,
                                    userEmail: (fields[15] || '').trim(),
                                    photoUrl: (fields[16] || '').trim(),
                                    attachmentUrl: (fields[17] || '').trim(),
                                    daysGiven: fields[19] || '',
                                    workDoneDay: fields[20] || '',
                                };
                            }
                        });

                    return tasks.filter((t): t is DashboardTask => t !== null);
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
