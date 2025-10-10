import React, { useState, useEffect, useMemo, useCallback, useRef } from 'react';
import { 
    AuthenticatedUser, Checklist, MasterTask, Task, Person
} from './types';
import { getStartOf, getIsoDate, parseDate } from './utils';

type ChecklistSubMode = 'templates' | 'master';

// --- ChecklistModal Component (moved from App) ---
const ChecklistModal: React.FC<{
    isOpen: boolean;
    onClose: () => void;
    onSave: (checklist: Omit<Checklist, 'id'>) => void;
    people: Person[];
}> = ({ isOpen, onClose, onSave, people }) => {
    const defaultState = {
        task: '',
        doer: people[0]?.name || '',
        frequency: 'D',
        date: new Date().toISOString().split('T')[0],
        buddy: people[0]?.name || '',
        secondBuddy: '',
    };
    const [newItem, setNewItem] = useState(defaultState);
    const prevIsOpen = useRef(isOpen);

    useEffect(() => {
        if (isOpen && !prevIsOpen.current) {
            setNewItem({
                task: '',
                doer: people[0]?.name || '',
                frequency: 'D',
                date: new Date().toISOString().split('T')[0],
                buddy: people[0]?.name || '',
                secondBuddy: '',
            });
        }
        prevIsOpen.current = isOpen;
    }, [isOpen, people]);

    if (!isOpen) return null;

    const handleSave = () => {
        if (newItem.task.trim() === '') return;
        onSave(newItem);
        onClose();
    };

    const handleFieldChange = (field: keyof Omit<Checklist, 'id'>, value: string) => {
        setNewItem(prev => ({ ...prev, [field]: value }));
    };

    return (
        <div className="modal-overlay" onClick={onClose} role="dialog" aria-modal="true" aria-labelledby="modal-title">
            <div className="modal-content" onClick={e => e.stopPropagation()}>
                <h2 id="modal-title">Add New Checklist Item</h2>
                <form className="modal-form" onSubmit={(e) => { e.preventDefault(); handleSave(); }}>
                    <div className="form-group">
                        <label htmlFor="task">Task</label>
                        <textarea id="task" value={newItem.task} onChange={e => handleFieldChange('task', e.target.value)} required />
                    </div>
                    <div className="form-grid">
                        <div className="form-group">
                            <label htmlFor="doer">Doer</label>
                            <select id="doer" value={newItem.doer} onChange={e => handleFieldChange('doer', e.target.value)} disabled={people.length === 0}>
                                {people.map(p => <option key={`doer-${p.name}`} value={p.name}>{p.name}</option>)}
                            </select>
                        </div>
                        <div className="form-group">
                            <label htmlFor="frequency">Frequency</label>
                            <select id="frequency" value={newItem.frequency} onChange={e => handleFieldChange('frequency', e.target.value)}>
                                <option value="D">Daily</option>
                                <option value="W">Weekly</option>
                                <option value="M">Monthly</option>
                                <option value="Q">Quarterly</option>
                                <option value="Y">Yearly</option>
                            </select>
                        </div>
                         <div className="form-group">
                            <label htmlFor="date">Start Date</label>
                            <input id="date" type="date" value={newItem.date} onChange={e => handleFieldChange('date', e.target.value)} />
                        </div>
                        <div className="form-group">
                            <label htmlFor="buddy">Buddy</label>
                            <select id="buddy" value={newItem.buddy} onChange={e => handleFieldChange('buddy', e.target.value)} disabled={people.length === 0}>
                                {people.map(p => <option key={`buddy-${p.name}`} value={p.name}>{p.name}</option>)}
                            </select>
                        </div>
                        <div className="form-group">
                            <label htmlFor="secondBuddy">Second Buddy</label>
                            <select id="secondBuddy" value={newItem.secondBuddy} onChange={e => handleFieldChange('secondBuddy', e.target.value)} disabled={people.length === 0}>
                                <option value="">None</option>
                                {people.map(p => <option key={`2nd-${p.name}`} value={p.name}>{p.name}</option>)}
                            </select>
                        </div>
                    </div>
                    <div className="modal-actions">
                        <button type="button" className="btn btn-cancel" onClick={onClose}>Cancel</button>
                        <button type="submit" className="btn btn-primary" disabled={!newItem.task.trim()}>Save Item</button>
                    </div>
                </form>
            </div>
        </div>
    );
};


// --- ChecklistSystem Component ---
interface ChecklistSystemProps {
    isAdmin: boolean;
    people: Person[];
    checklists: Checklist[];
    setChecklists: React.Dispatch<React.SetStateAction<Checklist[]>>;
    masterTasks: MasterTask[];
    setMasterTasks: React.Dispatch<React.SetStateAction<MasterTask[]>>;
    tasks: Task[];
    setTasks: React.Dispatch<React.SetStateAction<Task[]>>;
    authenticatedUser: AuthenticatedUser | null;
    postToGoogleSheet: (data: Record<string, any>) => Promise<any>;
    fetchData: (isInitialLoad?: boolean) => Promise<void>;
    checklistsError: string | null;
    masterTasksError: string | null;
    isRefreshing: boolean;
}

export const ChecklistSystem: React.FC<ChecklistSystemProps> = ({
    isAdmin,
    people,
    checklists,
    setChecklists,
    masterTasks,
    setMasterTasks,
    tasks,
    setTasks,
    authenticatedUser,
    postToGoogleSheet,
    fetchData,
    checklistsError,
    masterTasksError,
    isRefreshing,
}) => {
    const [checklistSubMode, setChecklistSubMode] = useState<ChecklistSubMode>('templates');
    const [isChecklistModalOpen, setIsChecklistModalOpen] = useState(false);

    // Checklist state
    const [editingChecklistId, setEditingChecklistId] = useState<string | null>(null);
    const [editedChecklist, setEditedChecklist] = useState<Partial<Checklist> | null>(null);
    const [deletingChecklistId, setDeletingChecklistId] = useState<string | null>(null);
    
    // Master Task state
    const [editingMasterTaskId, setEditingMasterTaskId] = useState<string | null>(null);
    const [editedMasterTask, setEditedMasterTask] = useState<Partial<MasterTask> | null>(null);
    const [deletingMasterTaskId, setDeletingMasterTaskId] = useState<string | null>(null);
    const [savingMasterTaskId, setSavingMasterTaskId] = useState<string | null>(null);
    const [undoneMasterTaskId, setUndoneMasterTaskId] = useState<string | null>(null);

    // Filter states
    const [masterFilters, setMasterFilters] = useState({
        description: '',
        doer: 'all',
        originalDoer: 'all',
        startDate: '',
        pc: '',
    });
     const [checklistFilters, setChecklistFilters] = useState({
        task: '',
        doer: 'all',
        frequency: 'all',
        buddy: 'all',
    });

    // Enforce view for 'User' role
    useEffect(() => {
        if (!isAdmin) {
            setChecklistSubMode('master');
        }
    }, [isAdmin]);

    // Effect for auto-generating tasks from checklists
    useEffect(() => {
        if (!isAdmin) return;
        if (checklists.length === 0) return;

        const now = new Date();
        const generatedTasks: Task[] = [];
        
        checklists.forEach(checklist => {
            let period: 'day' | 'week' | 'month' | 'quarter' | 'year' | null = null;
            switch(checklist.frequency.toUpperCase()) {
                case 'D': period = 'day'; break;
                case 'W': period = 'week'; break;
                case 'M': period = 'month'; break;
                case 'Q': period = 'quarter'; break;
                case 'Y': period = 'year'; break;
                default: break;
            }

            if(!period) return;

            const periodStart = getStartOf(now, period);

            const taskExists = tasks.some(task => 
                task.sourceChecklistId === checklist.id && 
                task.createdAt >= periodStart.getTime()
            );

            if (!taskExists) {
                generatedTasks.push({
                    id: `task-${Date.now()}-${Math.random()}`,
                    description: checklist.task,
                    assignee: checklist.doer,
                    buddy: checklist.buddy,
                    secondBuddy: checklist.secondBuddy,
                    completed: false,
                    createdAt: now.getTime(),
                    sourceChecklistId: checklist.id
                });
            }
        });

        if (generatedTasks.length > 0) {
            setTasks(prevTasks => [...prevTasks, ...generatedTasks]);
        }
    }, [checklists, tasks, setTasks, isAdmin]);

    // --- Checklist Handlers ---
    const handleSaveNewChecklist = (newItem: Omit<Checklist, 'id'>) => {
        const newChecklist: Checklist = {
            id: `cl-${Date.now()}`,
            ...newItem,
        };
        setChecklists(c => [newChecklist, ...c]);

        const sheetData = {
            action: 'create',
            sheetName: 'Task',
            newData: {
                "Task": newItem.task,
                "Doer": newItem.doer,
                "Frequency": newItem.frequency,
                "Date": newItem.date,
                "Buddy": newItem.buddy,
                "Second Buddy": newItem.secondBuddy || "",
            },
            historyRecord: {
                systemType: 'Task List',
                task: newItem.task,
                changedBy: authenticatedUser?.mailId,
                change: `Created on ${new Date().toLocaleString()}`
            }
        };
        postToGoogleSheet(sheetData).catch(error => {
            console.error("Failed to save new checklist item:", error);
            if (error.message !== 'No authenticated user.') {
               alert(`Failed to save new item to Google Sheets. Please refresh and try again. Error: ${error.message}`);
            }
            setChecklists(c => c.filter(item => item.id !== newChecklist.id));
        });
    };

    const handleDeleteChecklist = async (id: string) => {
        setDeletingChecklistId(id);
        const itemToDelete = checklists.find(cl => cl.id === id);
        if (!itemToDelete) {
            setDeletingChecklistId(null);
            return;
        }
        
        const sheetData = {
            action: 'delete',
            sheetName: 'Task',
            matchValue: itemToDelete.task,
            historyRecord: {
                systemType: 'Task List',
                task: itemToDelete.task,
                changedBy: authenticatedUser?.mailId,
                change: `Deleted on ${new Date().toLocaleString()}`
            }
        };

        try {
            setChecklists(c => c.filter(cl => cl.id !== id));
            setTasks(t => t.filter(task => task.sourceChecklistId !== id));
            await postToGoogleSheet(sheetData);
            await fetchData(false);
        } catch (error) {
            console.error("Failed to delete checklist item:", error);
            if (error instanceof Error && error.message !== 'No authenticated user.') {
              alert(`Failed to delete "${itemToDelete.task}" from Google Sheets. Your view has been restored.\n\nError: ${error.message}`);
              fetchData(false); 
            }
        } finally {
            setDeletingChecklistId(null);
        }
    };

    const handleEditChecklist = (checklist: Checklist) => {
        setEditingChecklistId(checklist.id);
        setEditedChecklist({...checklist});
    };

    const handleSaveChecklist = () => {
        if (!editedChecklist || !editingChecklistId) return;
        
        const originalChecklist = checklists.find(c => c.id === editingChecklistId);
        if (!originalChecklist) return;

        const updatedChecklist = { ...originalChecklist, ...editedChecklist } as Checklist;
        
        const changes: string[] = [];
        const fieldNames: Record<keyof Omit<Checklist, 'id'>, string> = {
            task: 'Task', doer: 'Doer', frequency: 'Frequency', date: 'Date', buddy: 'Buddy', secondBuddy: 'Second Buddy',
        };

        (Object.keys(fieldNames) as Array<keyof typeof fieldNames>).forEach(key => {
            const originalValue = originalChecklist[key] || '';
            const updatedValue = updatedChecklist[key] || '';
            if (originalValue !== updatedValue) {
                changes.push(`${fieldNames[key]}: "${originalValue}" -> "${updatedValue}"`);
            }
        });

        if (changes.length === 0) {
            setEditingChecklistId(null); setEditedChecklist(null); return;
        }

        setChecklists(checklists.map(c => c.id === editingChecklistId ? updatedChecklist : c));

        const sheetData = {
            action: 'update', sheetName: 'Task', matchValue: originalChecklist.task,
            updatedData: {
                "Task": updatedChecklist.task, "Doer": updatedChecklist.doer, "Frequency": updatedChecklist.frequency,
                "Date": updatedChecklist.date, "Buddy": updatedChecklist.buddy, "Second Buddy": updatedChecklist.secondBuddy || "",
            },
            historyRecord: {
                systemType: 'Task List',
                task: updatedChecklist.task,
                changedBy: authenticatedUser?.mailId,
                change: `Updated on ${new Date().toLocaleString()}: ${changes.join('; ')}`
            }
        };
        postToGoogleSheet(sheetData).catch(error => {
             console.error("Failed to update checklist item:", error);
             if (error.message !== 'No authenticated user.') {
                alert(`Failed to update item in Google Sheets. Your changes have been reverted. Error: ${error.message}`);
             }
             setChecklists(checklists.map(c => c.id === editingChecklistId ? originalChecklist : c));
        });
        
        setEditingChecklistId(null); setEditedChecklist(null);
    };

    const handleCancelEdit = () => {
        if (editedChecklist && editedChecklist.task === '') {
            setChecklists(c => c.filter(cl => cl.id !== editingChecklistId));
        }
        setEditingChecklistId(null); setEditedChecklist(null);
    };

    const handleChecklistFieldChange = (field: keyof Checklist, value: string) => {
        if(editedChecklist) {
            setEditedChecklist({...editedChecklist, [field]: value});
        }
    };
    
    // --- Master Task Handlers ---
    const handleEditMasterTask = (task: MasterTask) => {
        setEditingMasterTaskId(task.id);
        setEditedMasterTask({ ...task });
    };

    const handleCancelEditMasterTask = () => {
        setEditingMasterTaskId(null);
        setEditedMasterTask(null);
    };

    const handleMasterTaskFieldChange = (field: keyof MasterTask, value: string) => {
        if (editedMasterTask) {
            setEditedMasterTask({ ...editedMasterTask, [field]: value });
        }
    };
    
    const handleSaveMasterTask = async () => {
        if (!editedMasterTask || !editingMasterTaskId) return;

        setSavingMasterTaskId(editingMasterTaskId);

        const originalTask = masterTasks.find(t => t.id === editingMasterTaskId);
        if (!originalTask) {
            setSavingMasterTaskId(null);
            return;
        }

        const updatedTask = { ...originalTask, ...editedMasterTask } as MasterTask;
        
        setMasterTasks(masterTasks.map(t => t.id === editingMasterTaskId ? updatedTask : t));

        const sheetData = {
            action: 'update', sheetName: 'Master Data', matchValue: originalTask.taskId,
            updatedData: {
                "Planned": updatedTask.plannedDate,
                "Task": updatedTask.taskDescription,
                "Doer Name": updatedTask.doer,
                "Original Doer Name": updatedTask.originalDoer,
                "Frequency": updatedTask.frequency,
            },
            historyRecord: {
                systemType: 'Master Tasks',
                task: `Master Task ID: ${originalTask.taskId}`,
                changedBy: authenticatedUser?.mailId,
                change: `Updated on ${new Date().toLocaleString()}`
            }
        };

        try {
            await postToGoogleSheet(sheetData);
        } catch (error) {
            console.error("Failed to update master task:", error);
            if (error instanceof Error && error.message !== 'No authenticated user.') {
              alert(`Failed to update master task in Google Sheets. Your changes have been reverted. Error: ${error.message}`);
            }
            setMasterTasks(masterTasks.map(t => t.id === editingMasterTaskId ? originalTask : t));
        } finally {
            setSavingMasterTaskId(null);
            setEditingMasterTaskId(null);
            setEditedMasterTask(null);
        }
    };

    const handleDeleteMasterTask = async (id: string) => {
        setDeletingMasterTaskId(id);
        const taskToDelete = masterTasks.find(t => t.id === id);
        if (!taskToDelete) {
            setDeletingMasterTaskId(null); return;
        }

        const sheetData = {
            action: 'delete', sheetName: 'Master Data', matchValue: taskToDelete.taskId,
            historyRecord: {
                systemType: 'Master Tasks',
                task: `Master Task ID: ${taskToDelete.taskId}`,
                changedBy: authenticatedUser?.mailId,
                change: `Deleted on ${new Date().toLocaleString()}`
            }
        };

        try {
            setMasterTasks(tasks => tasks.filter(t => t.id !== id));
            await postToGoogleSheet(sheetData);
            await fetchData(false);
        } catch (error) {
            console.error("Failed to delete master task:", error);
             if (error instanceof Error && error.message !== 'No authenticated user.') {
                alert(`Failed to delete master task from Google Sheets. Your view have been restored.\n\nError: ${error.message}`);
                fetchData(false);
             }
        } finally {
            setDeletingMasterTaskId(null);
        }
    };

    const handleUndoneMasterTask = async (taskToUpdate: MasterTask) => {
        setUndoneMasterTaskId(taskToUpdate.id);
        const updatedTask = { ...taskToUpdate, status: 'Undone' };
    
        // Action: Delete the corresponding entry from the 'Done Task Status' sheet
        const deleteDoneStatusData = {
            action: 'delete',
            sheetName: 'Done Task Status',
            matchValue: taskToUpdate.taskId,
            historyRecord: {
                systemType: 'Master Tasks',
                task: `Master Task ID: ${taskToUpdate.taskId}`,
                changedBy: authenticatedUser?.mailId,
                change: `'Done' record deleted on ${new Date().toLocaleString()}`
            }
        };
    
        try {
            // Optimistic UI update
            setMasterTasks(tasks => tasks.map(t => t.id === taskToUpdate.id ? updatedTask : t));
    
            // Perform the Google Sheet operation to delete the 'Done' record
            await postToGoogleSheet(deleteDoneStatusData);
    
        } catch (error) {
            console.error("Failed to process 'Undone' action:", error);
            if (error instanceof Error && error.message !== 'No authenticated user.') {
                alert(`Failed to update task status in Google Sheets. Your change has been reverted. Please check if the task has a 'Done' record to remove. Error: ${error.message}`);
            }
            // Rollback on failure
            setMasterTasks(tasks => tasks.map(t => t.id === taskToUpdate.id ? taskToUpdate : t));
        } finally {
            setUndoneMasterTaskId(null);
        }
    };

    // --- Filter Logic ---

    // Master Task Filters
    const uniqueDoers = useMemo(() => {
        const doers = new Set(masterTasks.map(task => task.doer).filter(Boolean));
        return ['all', ...Array.from(doers).sort()];
    }, [masterTasks]);
    const uniqueOriginalDoers = useMemo(() => {
        const originalDoers = new Set(masterTasks.map(task => task.originalDoer).filter(Boolean));
        return ['all', ...Array.from(originalDoers).sort()];
    }, [masterTasks]);
    const handleMasterFilterChange = (filterName: keyof typeof masterFilters, value: string) => {
        setMasterFilters(prev => ({...prev, [filterName]: value}));
    };
    const clearMasterFilters = () => {
        setMasterFilters({ description: '', doer: 'all', originalDoer: 'all', startDate: '', pc: '' });
    };
    const filteredMasterTasks = useMemo(() => {
        const today = new Date();
        today.setHours(0, 0, 0, 0);

        return masterTasks.filter(task => {
            const plannedDate = parseDate(task.plannedDate);
            if (plannedDate) {
                plannedDate.setHours(0, 0, 0, 0);
                if (plannedDate.getTime() > today.getTime()) {
                    return false;
                }
            }

            const searchTerm = masterFilters.description.toLowerCase();
            const searchMatch = masterFilters.description === '' ||
                task.taskDescription.toLowerCase().includes(searchTerm) ||
                task.taskId.toLowerCase().includes(searchTerm);
            const doerMatch = masterFilters.doer === 'all' || task.doer === masterFilters.doer;
            const originalDoerMatch = masterFilters.originalDoer === 'all' || task.originalDoer === masterFilters.originalDoer;
            const pcMatch = masterFilters.pc === '' || (task.pc && task.pc.toLowerCase().includes(masterFilters.pc.toLowerCase()));
            let dateMatch = true;
            if (masterFilters.startDate) {
                try {
                    const startDate = new Date(masterFilters.startDate);
                    startDate.setHours(0,0,0,0);
                    const taskPlannedDate = parseDate(task.plannedDate);
                    if (taskPlannedDate) { 
                        taskPlannedDate.setHours(0, 0, 0, 0);
                        dateMatch = taskPlannedDate >= startDate; 
                    } else { 
                        dateMatch = false; 
                    }
                } catch (e) { dateMatch = false; }
            }
            return searchMatch && doerMatch && originalDoerMatch && dateMatch && pcMatch;
        });
    }, [masterTasks, masterFilters]);

    // Checklist (Task List) Filters
    const allPeopleNamesForFilter = useMemo(() => {
        if (people.length === 0) return ['all'];
        const names = new Set(people.map(p => p.name).filter(Boolean));
        return ['all', ...Array.from(names).sort()];
    }, [people]);
    const handleChecklistFilterChange = (filterName: keyof typeof checklistFilters, value: string) => {
        setChecklistFilters(prev => ({...prev, [filterName]: value}));
    };
    const clearChecklistFilters = () => {
        setChecklistFilters({ task: '', doer: 'all', frequency: 'all', buddy: 'all' });
    };
    const filteredChecklists = useMemo(() => {
        return checklists.filter(c => {
            const taskMatch = checklistFilters.task === '' || c.task.toLowerCase().includes(checklistFilters.task.toLowerCase());
            const doerMatch = checklistFilters.doer === 'all' || c.doer === checklistFilters.doer;
            const freqMatch = checklistFilters.frequency === 'all' || c.frequency === checklistFilters.frequency;
            const buddyMatch = checklistFilters.buddy === 'all' || c.buddy === checklistFilters.buddy;
            return taskMatch && doerMatch && freqMatch && buddyMatch;
        });
    }, [checklists, checklistFilters]);


    return (
        <div className="main-content">
             <div className="checklist-view" role="region" aria-labelledby="checklist-title">
                {isAdmin && (
                    <div className="sub-nav">
                        <button onClick={() => setChecklistSubMode('templates')} className={checklistSubMode === 'templates' ? 'active' : ''} aria-pressed={checklistSubMode === 'templates'}>Task List</button>
                        <button onClick={() => setChecklistSubMode('master')} className={checklistSubMode === 'master' ? 'active' : ''} aria-pressed={checklistSubMode === 'master'}>Master Tasks</button>
                    </div>
                )}

                {checklistSubMode === 'templates' && isAdmin && (
                    <>
                        <div className="page-header">
                            <h2 id="checklist-title">Task List ({filteredChecklists.length})</h2>
                            <button className="btn btn-primary btn-attention" aria-label="Add new checklist item" onClick={() => setIsChecklistModalOpen(true)} disabled={editingChecklistId !== null || deletingChecklistId !== null}>
                                <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor" viewBox="0 0 16 16"><path d="M8 4a.5.5 0 0 1 .5.5v3h3a.5.5 0 0 1 0 1h-3v3a.5.5 0 0 1-1 0v-3h-3a.5.5 0 0 1 0-1h3v-3A.5.5 0 0 1 8 4"/></svg>
                                <span>Add Task</span>
                            </button>
                        </div>
                        {checklistsError && <div className="error-message" style={{marginBottom: '24px'}}>{checklistsError}</div>}
                        <div className="filter-bar" role="search">
                            <div className="filter-group">
                                <label htmlFor="cl-task-search">Task Search</label>
                                <input type="text" id="cl-task-search" placeholder="Filter by task..." value={checklistFilters.task} onChange={e => handleChecklistFilterChange('task', e.target.value)} />
                            </div>
                            <div className="filter-group">
                                <label htmlFor="cl-doer-filter">Doer</label>
                                <select id="cl-doer-filter" value={checklistFilters.doer} onChange={e => handleChecklistFilterChange('doer', e.target.value)}>
                                    {allPeopleNamesForFilter.map(d => <option key={`cldoer-${d}`} value={d}>{d === 'all' ? 'All Doers' : d}</option>)}
                                </select>
                            </div>
                            <div className="filter-group">
                                <label htmlFor="cl-freq-filter">Frequency</label>
                                <select id="cl-freq-filter" value={checklistFilters.frequency} onChange={e => handleChecklistFilterChange('frequency', e.target.value)}>
                                    <option value="all">All Frequencies</option>
                                    <option value="D">Daily</option><option value="W">Weekly</option><option value="M">Monthly</option><option value="Q">Quarterly</option><option value="Y">Yearly</option>
                                </select>
                            </div>
                            <div className="filter-group">
                                <label htmlFor="cl-buddy-filter">Buddy</label>
                                <select id="cl-buddy-filter" value={checklistFilters.buddy} onChange={e => handleChecklistFilterChange('buddy', e.target.value)}>
                                    {allPeopleNamesForFilter.map(b => <option key={`clbuddy-${b}`} value={b}>{b === 'all' ? 'All Buddies' : b}</option>)}
                                </select>
                            </div>
                            <button className="btn" onClick={clearChecklistFilters}>Clear</button>
                        </div>

                        <div className="table-container">
                            <table className={`checklist-table ${!isAdmin ? 'user-view' : ''}`}>
                                <thead>
                                    <tr><th>Task</th><th>Doer</th><th>Frequency</th><th>Date</th><th>Buddy</th><th>Second Buddy</th>{isAdmin && <th>Actions</th>}</tr>
                                </thead>
                                <tbody>
                                    {filteredChecklists.map(cl => (
                                        editingChecklistId === cl.id && editedChecklist ? (
                                            <tr key={cl.id} className="editing-row">
                                                <td><input type="text" value={editedChecklist.task ?? ''} onChange={e => handleChecklistFieldChange('task', e.target.value)} /></td>
                                                <td>
                                                    <select value={editedChecklist.doer} onChange={e => handleChecklistFieldChange('doer', e.target.value)} disabled={people.length === 0}>
                                                        <option value="" disabled>Select a person</option>
                                                        {people.map(p => <option key={`edit-doer-${p.email || p.name}`} value={p.name}>{p.name}</option>)}
                                                    </select>
                                                </td>
                                                <td><select value={editedChecklist.frequency} onChange={e => handleChecklistFieldChange('frequency', e.target.value)}><option value="D">Daily</option><option value="W">Weekly</option><option value="M">Monthly</option><option value="Q">Quarterly</option><option value="Y">Yearly</option></select></td>
                                                <td><input type="date" value={getIsoDate(editedChecklist.date)} onChange={e => handleChecklistFieldChange('date', e.target.value)} /></td>
                                                <td>
                                                    <select value={editedChecklist.buddy} onChange={e => handleChecklistFieldChange('buddy', e.target.value)} disabled={people.length === 0}>
                                                        <option value="" disabled>Select a person</option>
                                                        {people.map(p => <option key={`edit-buddy-${p.email || p.name}`} value={p.name}>{p.name}</option>)}
                                                    </select>
                                                </td>
                                                <td>
                                                    <select value={editedChecklist.secondBuddy} onChange={e => handleChecklistFieldChange('secondBuddy', e.target.value)} disabled={people.length === 0}>
                                                        <option value="">None</option>
                                                        {people.map(p => <option key={`edit-2buddy-${p.email || p.name}`} value={p.name}>{p.name}</option>)}
                                                    </select>
                                                </td>
                                                {isAdmin && <td className="actions-cell"><button className="btn btn-save" onClick={handleSaveChecklist}>Save</button><button className="btn btn-cancel" onClick={handleCancelEdit}>Cancel</button></td>}
                                            </tr>
                                        ) : (
                                            <tr key={cl.id}>
                                                <td>{cl.task}</td><td>{cl.doer}</td><td>{cl.frequency}</td><td>{cl.date}</td><td>{cl.buddy}</td><td>{cl.secondBuddy}</td>
                                                {isAdmin && (
                                                    <td className="actions-cell">
                                                        <button className="btn btn-edit" onClick={() => handleEditChecklist(cl)} disabled={editingChecklistId !== null || deletingChecklistId !== null}>Edit</button>
                                                        <button className="delete-btn" onClick={() => handleDeleteChecklist(cl.id)} disabled={editingChecklistId !== null || deletingChecklistId === cl.id}><svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor" viewBox="0 0 16 16"><path d="M5.5 5.5A.5.5 0 0 1 6 6v6a.5.5 0 0 1-1 0V6a.5.5 0 0 1 .5-.5m2.5 0a.5.5 0 0 1 .5.5v6a.5.5 0 0 1-1 0V6a.5.5 0 0 1 .5-.5m3 .5a.5.5 0 0 0-1 0v6a.5.5 0 0 0 1 0z"/><path d="M14.5 3a1 1 0 0 1-1 1H13v9a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V4h-.5a1 1 0 0 1-1-1V2a1 1 0 0 1 1-1H6a1 1 0 0 1 1-1h2a1 1 0 0 1 1 1h3.5a1 1 0 0 1 1 1zM4.118 4 4 4.059V13a1 1 0 0 0 1 1h6a1 1 0 0 0 1-1V4.059L11.882 4zM2.5 3h11V2h-11z"/></svg></button>
                                                    </td>
                                                )}
                                            </tr>
                                        )
                                    ))}
                                </tbody>
                            </table>
                        </div>
                    </>
                )}
                
                {(checklistSubMode === 'master' || !isAdmin) && (
                    <>
                        <h2 id="checklist-title">Master Tasks ({filteredMasterTasks.length})</h2>
                        {masterTasksError && <div className="error-message" style={{marginBottom: '24px'}}>{masterTasksError}</div>}
                        <div className="filter-bar" role="search" aria-labelledby="filter-heading">
                            <div className="filter-group"><label htmlFor="task-search">Task Search</label><input type="text" id="task-search" placeholder="Filter by Task or ID..." value={masterFilters.description} onChange={e => handleMasterFilterChange('description', e.target.value)} /></div>
                            <div className="filter-group"><label htmlFor="doer-filter">Doer</label><select id="doer-filter" value={masterFilters.doer} onChange={e => handleMasterFilterChange('doer', e.target.value)}>{uniqueDoers.map(d => <option key={`doer-${d}`} value={d}>{d === 'all' ? 'All Doers' : d}</option>)}</select></div>
                            <div className="filter-group"><label htmlFor="orig-doer-filter">Original Doer</label><select id="orig-doer-filter" value={masterFilters.originalDoer} onChange={e => handleMasterFilterChange('originalDoer', e.target.value)}>{uniqueOriginalDoers.map(d => <option key={`orig-doer-${d}`} value={d}>{d === 'all' ? 'All Original' : d}</option>)}</select></div>
                            <div className="filter-group"><label htmlFor="pc-filter">PC</label><input type="text" id="pc-filter" placeholder="Filter by PC..." value={masterFilters.pc} onChange={e => handleMasterFilterChange('pc', e.target.value)} /></div>
                            <div className="filter-group"><label htmlFor="start-date">Planned on or after</label><input type="date" id="start-date" value={masterFilters.startDate} onChange={e => handleMasterFilterChange('startDate', e.target.value)}/></div>
                            <button className="btn" onClick={clearMasterFilters}>Clear</button>
                        </div>

                        <div className="table-container">
                            <table className={`checklist-table master-tasks-table ${!isAdmin ? 'user-view' : ''}`}>
                                <thead>
                                    <tr><th>Task ID</th><th>Planned</th><th>Actual</th><th>Task</th><th>Doer Name</th><th>Original Doer Name</th><th>Frequency</th><th>PC</th>{isAdmin && <th>Actions</th>}</tr>
                                </thead>
                                <tbody>
                                    {filteredMasterTasks.map(task => (
                                        editingMasterTaskId === task.id && editedMasterTask ? (
                                            <tr key={task.id} className={savingMasterTaskId === task.id ? "saving-row" : "editing-row"}>
                                                <td>{task.taskId}</td>
                                                <td><input type="date" value={getIsoDate(editedMasterTask.plannedDate)} onChange={e => handleMasterTaskFieldChange('plannedDate', e.target.value)} disabled={savingMasterTaskId === task.id} /></td>
                                                <td>{editedMasterTask.actualDate}</td>
                                                <td><input type="text" value={editedMasterTask.taskDescription ?? ''} onChange={e => handleMasterTaskFieldChange('taskDescription', e.target.value)} disabled={savingMasterTaskId === task.id} /></td>
                                                <td><select value={editedMasterTask.doer} onChange={e => handleMasterTaskFieldChange('doer', e.target.value)} disabled={people.length === 0 || savingMasterTaskId === task.id}>{people.map(p => <option key={`medit-doer-${p.name}`} value={p.name}>{p.name}</option>)}</select></td>
                                                <td><select value={editedMasterTask.originalDoer} onChange={e => handleMasterTaskFieldChange('originalDoer', e.target.value)} disabled={people.length === 0 || savingMasterTaskId === task.id}>{people.map(p => <option key={`medit-orig-doer-${p.name}`} value={p.name}>{p.name}</option>)}</select></td>
                                                <td><input type="text" value={editedMasterTask.frequency ?? ''} onChange={e => handleMasterTaskFieldChange('frequency', e.target.value)} disabled={savingMasterTaskId === task.id} /></td>
                                                <td>{editedMasterTask.pc}</td>
                                                {isAdmin && <td className="actions-cell">
                                                    {savingMasterTaskId === task.id ? (
                                                        <div className="saving-indicator">
                                                            <div className="spinner"></div>
                                                            <span>Saving...</span>
                                                        </div>
                                                    ) : (
                                                        <>
                                                            <button className="btn btn-save" onClick={handleSaveMasterTask}>Save</button>
                                                            <button className="btn btn-cancel" onClick={handleCancelEditMasterTask}>Cancel</button>
                                                        </>
                                                    )}
                                                </td>}
                                            </tr>
                                        ) : (
                                            <tr key={task.id}>
                                                <td>{task.taskId}</td>
                                                <td>{task.plannedDate}</td>
                                                <td>{task.actualDate}</td>
                                                <td>{task.taskDescription}</td>
                                                <td>{task.doer}</td>
                                                <td>{task.originalDoer}</td>
                                                <td>{task.frequency}</td>
                                                <td>{task.pc}</td>
                                                {isAdmin && (
                                                    <td className="actions-cell">
                                                        <button className="btn btn-edit" onClick={() => handleEditMasterTask(task)} disabled={editingMasterTaskId !== null || deletingMasterTaskId !== null || savingMasterTaskId !== null || undoneMasterTaskId !== null}>Edit</button>
                                                        {(task.actualDate && task.actualDate.trim() !== '') && (!task.status || task.status.trim() === '') && (
                                                            <button 
                                                                className="btn btn-secondary btn-undone"
                                                                onClick={() => handleUndoneMasterTask(task)} 
                                                                disabled={editingMasterTaskId !== null || deletingMasterTaskId !== null || savingMasterTaskId !== null || undoneMasterTaskId !== null}
                                                            >
                                                                {undoneMasterTaskId === task.id ? '...' : 'Undone'}
                                                            </button>
                                                        )}
                                                        <button className="delete-btn" onClick={() => handleDeleteMasterTask(task.id)} disabled={editingMasterTaskId !== null || deletingMasterTaskId !== null || savingMasterTaskId !== null || undoneMasterTaskId !== null}><svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor" viewBox="0 0 16 16"><path d="M5.5 5.5A.5.5 0 0 1 6 6v6a.5.5 0 0 1-1 0V6a.5.5 0 0 1 .5-.5m2.5 0a.5.5 0 0 1 .5.5v6a.5.5 0 0 1-1 0V6a.5.5 0 0 1 .5-.5m3 .5a.5.5 0 0 0-1 0v6a.5.5 0 0 0 1 0z"/><path d="M14.5 3a1 1 0 0 1-1 1H13v9a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V4h-.5a1 1 0 0 1-1-1V2a1 1 0 0 1 1-1H6a1 1 0 0 1 1-1h2a1 1 0 0 1 1 1h3.5a1 1 0 0 1 1 1zM4.118 4 4 4.059V13a1 1 0 0 0 1 1h6a1 1 0 0 0 1-1V4.059L11.882 4zM2.5 3h11V2h-11z"/></svg></button>
                                                    </td>
                                                )}
                                            </tr>
                                        )
                                    ))}
                                </tbody>
                            </table>
                        </div>
                         {masterTasks.length > 0 && filteredMasterTasks.length === 0 && !isRefreshing && ( <div className="no-content" style={{marginTop: '20px'}}><h3>No Tasks Match Filters</h3><p>Try adjusting or clearing your filters to see more results.</p></div> )}
                        {masterTasks.length === 0 && !isRefreshing && ( <div className="no-content" style={{marginTop: '20px'}}><h3>No Master Tasks Found</h3><p>The "Master Data" sheet might be empty or data is still loading.</p></div>)}
                    </>
                )}
            </div>
            <ChecklistModal 
                isOpen={isChecklistModalOpen}
                onClose={() => setIsChecklistModalOpen(false)}
                onSave={handleSaveNewChecklist}
                people={people}
            />
        </div>
    );
};