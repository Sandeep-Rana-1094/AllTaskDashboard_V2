
import React, { useState, useMemo } from 'react';
import { AuthenticatedUser, DelegationTask, Person } from './types';
import { getIsoDate, parseDate } from './utils';

interface DelegationSystemProps {
    people: Person[];
    delegationTasks: DelegationTask[];
    setDelegationTasks: React.Dispatch<React.SetStateAction<DelegationTask[]>>;
    authenticatedUser: AuthenticatedUser | null;
    postToGoogleSheet: (data: Record<string, any>) => Promise<any>;
    fetchData: (isInitialLoad?: boolean) => Promise<void>;
    delegationFormUrl: string;
    delegationTasksError: string | null;
    isRefreshing: boolean;
}

export const DelegationSystem: React.FC<DelegationSystemProps> = ({
    people,
    delegationTasks,
    setDelegationTasks,
    authenticatedUser,
    postToGoogleSheet,
    fetchData,
    delegationFormUrl,
    delegationTasksError,
    isRefreshing,
}) => {
    const [editingDelegationTaskId, setEditingDelegationTaskId] = useState<string | null>(null);
    const [editedDelegationTask, setEditedDelegationTask] = useState<Partial<DelegationTask> | null>(null);
    const [savingDelegationTaskId, setSavingDelegationTaskId] = useState<string | null>(null);
    const [cancellingTaskId, setCancellingTaskId] = useState<string | null>(null);
    const [undoneDelegationTaskId, setUndoneDelegationTaskId] = useState<string | null>(null);
    const [delegationFilters, setDelegationFilters] = useState({
        task: '',
        assignee: 'all',
        assigner: 'all',
        plannedDate: '',
        status: 'all',
    });

    // --- Delegation Task Handlers ---
    const handleEditDelegationTask = (task: DelegationTask) => {
        setEditingDelegationTaskId(task.id);
        setEditedDelegationTask({ ...task });
    };

    const handleCancelEditDelegationTask = () => {
        setEditingDelegationTaskId(null);
        setEditedDelegationTask(null);
    };

    const handleDelegationTaskFieldChange = (field: keyof DelegationTask, value: string) => {
        if (editedDelegationTask) {
            setEditedDelegationTask({ ...editedDelegationTask, [field]: value });
        }
    };

    const handleSaveDelegationTask = async () => {
        if (!editedDelegationTask || !editingDelegationTaskId) return;

        setSavingDelegationTaskId(editingDelegationTaskId);

        const originalTask = delegationTasks.find(t => t.id === editingDelegationTaskId);
        if (!originalTask || !originalTask.taskId) {
            alert("Cannot update this task: original task data or Task ID is missing.");
            setSavingDelegationTaskId(null);
            setEditingDelegationTaskId(null);
            setEditedDelegationTask(null);
            return;
        }

        const updatedDataForSheet: { [key: string]: string } = {};
        const updatedTaskForUI: Partial<DelegationTask> = {};
        const changes: string[] = [];

        const addChange = (field: string, from: string, to: string) => {
            changes.push(`${field}: "${from || ''}" -> "${to || ''}"`);
        };

        // 1. Check for Task change
        if (editedDelegationTask.task !== undefined && editedDelegationTask.task !== originalTask.task) {
            updatedDataForSheet['Task'] = editedDelegationTask.task;
            updatedTaskForUI.task = editedDelegationTask.task;
            addChange('Task', originalTask.task, editedDelegationTask.task);
        }

        // 2. Check for Planned Date change
        const originalIsoDate = getIsoDate(originalTask.plannedDate);
        const newIsoDate = editedDelegationTask.plannedDate ? getIsoDate(editedDelegationTask.plannedDate) : originalIsoDate;
        if (newIsoDate !== originalIsoDate) {
            updatedDataForSheet['Planned Date'] = newIsoDate;
            updatedTaskForUI.plannedDate = newIsoDate;
            addChange('Planned Date', originalTask.plannedDate, newIsoDate);
        }

        // 3. Check for Assignee (Assign To) change
        if (editedDelegationTask.assignee !== undefined && editedDelegationTask.assignee !== originalTask.assignee) {
            const person = people.find(p => p.name === editedDelegationTask.assignee);
            if (person) {
                updatedDataForSheet['Assign To'] = person.name;
                // DO NOT update 'Delegate Email' as it is a formula column in the sheet.
                updatedTaskForUI.assignee = person.name;
                updatedTaskForUI.delegateEmail = person.email || '';
                addChange('Assign To', originalTask.assignee, person.name);
            } else {
                alert(`Error: Could not find details for selected assignee "${editedDelegationTask.assignee}". Update cancelled.`);
                setSavingDelegationTaskId(null);
                return;
            }
        }

        // 4. Check for Assigner (Assign by) change
        if (editedDelegationTask.assigner !== undefined && editedDelegationTask.assigner !== originalTask.assigner) {
            const person = people.find(p => p.name === editedDelegationTask.assigner);
            if (person) {
                updatedDataForSheet['Assign By'] = person.name;
                // Per user request, do not update the email column.
                // It is assumed to be a formula or not required to be updated from the app.
                updatedTaskForUI.assigner = person.name;
                updatedTaskForUI.assignerEmail = person.email || '';
                addChange('Assign By', originalTask.assigner, person.name);
            } else {
                alert(`Error: Could not find details for selected assigner "${editedDelegationTask.assigner}". Update cancelled.`);
                setSavingDelegationTaskId(null);
                return;
            }
        }
        
        if (changes.length === 0) {
            setSavingDelegationTaskId(null);
            setEditingDelegationTaskId(null);
            setEditedDelegationTask(null);
            return;
        }

        const finalTaskForUI = { ...originalTask, ...updatedTaskForUI };

        const sheetData = {
            action: 'update',
            sheetName: 'Working Task Form',
            matchValue: originalTask.taskId,
            updatedData: updatedDataForSheet,
            historyRecord: {
                systemType: 'Delegation',
                task: `Task ID: ${originalTask.taskId} - ${finalTaskForUI.task}`,
                changedBy: authenticatedUser?.mailId,
                change: `Updated on ${new Date().toLocaleString()}: ${changes.join('; ')}`
            }
        };
        
        try {
            // Optimistic UI update
            setDelegationTasks(tasks => tasks.map(t => t.id === editingDelegationTaskId ? finalTaskForUI : t));
            await postToGoogleSheet(sheetData);
        } catch (error) {
            console.error("Failed to update delegation task:", error);
            if (error instanceof Error && error.message !== 'No authenticated user.') {
              alert(`Failed to update delegation task in Google Sheets. Your changes have been reverted. Error: ${error.message}`);
            }
            // Rollback on failure
            setDelegationTasks(tasks => tasks.map(t => t.id === editingDelegationTaskId ? originalTask : t));
        } finally {
            setSavingDelegationTaskId(null);
            setEditingDelegationTaskId(null);
            setEditedDelegationTask(null);
        }
    };

    const handleCancelDelegationTask = async (taskToCancel: DelegationTask) => {
        if (!taskToCancel.taskId) {
            alert("Cannot cancel this task: Task ID is missing.");
            return;
        }
    
        setCancellingTaskId(taskToCancel.id);
    
        // Update 'Working Task Form' to set Status to "Cancel".
        // This will filter it out of the view on the next data refresh.
        const workingFormUpdateData = {
            action: 'update',
            sheetName: 'Working Task Form',
            matchValue: taskToCancel.taskId,
            updatedData: { 'Status': 'Cancel' },
            historyRecord: { // Attach history record to this update
                systemType: 'Delegation',
                task: `Task ID: ${taskToCancel.taskId} - ${taskToCancel.task}`,
                changedBy: authenticatedUser?.mailId,
                change: `Cancelled on ${new Date().toLocaleString()}`
            }
        };
    
        try {
            // Send the update request to the backend script.
            await postToGoogleSheet(workingFormUpdateData);
            // On success, refresh the data to remove the cancelled task from the view.
            await fetchData(false);
            
        } catch (error) {
            console.error("Failed to cancel delegation task:", error);
            if (error instanceof Error && error.message !== 'No authenticated user.') {
              alert(`Failed to cancel task. Please try again. Error: ${error.message}`);
            }
        } finally {
            setCancellingTaskId(null);
        }
    };

    const handleUndoneDelegationTask = async (taskToUpdate: DelegationTask) => {
        if (!taskToUpdate.taskId) {
            alert("Cannot 'Undone' this task: Task ID is missing.");
            return;
        }
        setUndoneDelegationTaskId(taskToUpdate.id);
        const updatedTask = { ...taskToUpdate, actualDate: '', status: '' }; // Create an 'undone' version for UI

        const deleteDoneStatusData = {
            action: 'delete',
            sheetName: 'Done Task Status',
            matchValue: taskToUpdate.taskId,
            historyRecord: {
                systemType: 'Delegation',
                task: `Task ID: ${taskToUpdate.taskId}`,
                changedBy: authenticatedUser?.mailId,
                change: `'Done' record deleted (Undone) on ${new Date().toLocaleString()}`
            }
        };

        try {
            // Optimistic UI update
            setDelegationTasks(tasks => tasks.map(t => t.id === taskToUpdate.id ? updatedTask : t));
            await postToGoogleSheet(deleteDoneStatusData);
        } catch (error) {
            console.error("Failed to process 'Undone' action:", error);
            if (error instanceof Error && error.message !== 'No authenticated user.') {
                alert(`Failed to update task status in Google Sheets. Your change has been reverted. Error: ${error.message}`);
            }
            // Rollback on failure
            setDelegationTasks(tasks => tasks.map(t => t.id === taskToUpdate.id ? taskToUpdate : t));
        } finally {
            setUndoneDelegationTaskId(null);
        }
    };


    // --- Filter Logic ---
    const handleDelegationFilterChange = (filterName: keyof typeof delegationFilters, value: string) => {
        setDelegationFilters(prev => ({...prev, [filterName]: value}));
    };
    const clearDelegationFilters = () => {
        setDelegationFilters({ task: '', assignee: 'all', assigner: 'all', plannedDate: '', status: 'all' });
    };
    const allPeopleNamesForFilter = useMemo(() => {
        if (people.length === 0) return ['all'];
        const names = new Set(people.map(p => p.name).filter(Boolean));
        return ['all', ...Array.from(names).sort()];
    }, [people]);

    const filteredDelegationTasks = useMemo(() => {
        return delegationTasks.filter(task => {
            const searchTerm = delegationFilters.task.toLowerCase();
            const taskMatch = delegationFilters.task === '' ||
                task.task.toLowerCase().includes(searchTerm) ||
                (task.taskId && task.taskId.toLowerCase().includes(searchTerm));
            if (!taskMatch) return false;

            const assigneeMatch = delegationFilters.assignee === 'all' || task.assignee === delegationFilters.assignee;
            if (!assigneeMatch) return false;

            const assignerMatch = delegationFilters.assigner === 'all' || task.assigner === delegationFilters.assigner;
            if (!assignerMatch) return false;

            const statusMatch = (() => {
                if (delegationFilters.status === 'all') {
                    return true;
                }
                const isDone = task.actualDate && task.actualDate.trim() !== '';
                if (delegationFilters.status === 'done') {
                    return isDone;
                }
                if (delegationFilters.status === 'not-done') {
                    return !isDone;
                }
                return true;
            })();
            if (!statusMatch) return false;

            if (delegationFilters.plannedDate) {
                try {
                    const filterDate = parseDate(delegationFilters.plannedDate);
                    if (!filterDate) {
                        return false; 
                    }
                    filterDate.setHours(0, 0, 0, 0);

                    const taskPlannedDate = parseDate(task.plannedDate);
                    if (!taskPlannedDate) return false;
                    
                    taskPlannedDate.setHours(0, 0, 0, 0);
                    return taskPlannedDate.getTime() === filterDate.getTime();
                } catch (e) {
                    console.error("Error during date filtering:", e);
                    return false;
                }
            } else {
                // If no specific plannedDate filter is active, show ALL tasks from the sheet.
                // This removes the previous logic that hid future-dated tasks.
                return true;
            }
        });
    }, [delegationTasks, delegationFilters]);
    

    return (
        <div className="main-content">
            <div role="region" aria-labelledby="delegation-title">
                <div className="page-header">
                    <h2 id="delegation-title">Delegation System ({filteredDelegationTasks.length})</h2>
                     <button className="btn btn-primary btn-attention" aria-label="Add new delegation task" disabled={editingDelegationTaskId !== null || savingDelegationTaskId !== null} onClick={() => window.open(delegationFormUrl, '_blank', 'noopener,noreferrer')}>
                        <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor" viewBox="0 0 16 16"><path d="M8 4a.5.5 0 0 1 .5.5v3h3a.5.5 0 0 1 0 1h-3v3a.5.5 0 0 1-1 0v-3h-3a.5.5 0 0 1 0-1h3v-3A.5.5 0 0 1 8 4"/></svg>
                        <span>Add Delegation</span>
                    </button>
                </div>

                {delegationTasksError && <div className="error-message" style={{marginBottom: '24px'}}>{delegationTasksError}</div>}

                <div className="filter-bar" role="search" aria-labelledby="delegation-filter-heading">
                    <div className="filter-group"><label htmlFor="del-task-search">Task Search</label><input type="text" id="del-task-search" placeholder="Filter by task or unique ID..." value={delegationFilters.task} onChange={e => handleDelegationFilterChange('task', e.target.value)} /></div>
                    <div className="filter-group"><label htmlFor="del-assignee-filter">Assignee</label><select id="del-assignee-filter" value={delegationFilters.assignee} onChange={e => handleDelegationFilterChange('assignee', e.target.value)}>{allPeopleNamesForFilter.map(d => <option key={`del-assignee-${d}`} value={d}>{d === 'all' ? 'All Assignees' : d}</option>)}</select></div>
                    <div className="filter-group"><label htmlFor="del-assigner-filter">Assigner</label><select id="del-assigner-filter" value={delegationFilters.assigner} onChange={e => handleDelegationFilterChange('assigner', e.target.value)}>{allPeopleNamesForFilter.map(d => <option key={`del-assigner-${d}`} value={d}>{d === 'all' ? 'All Assigners' : d}</option>)}</select></div>
                    <div className="filter-group"><label htmlFor="del-planned-date">Planned Date</label><input type="date" id="del-planned-date" value={delegationFilters.plannedDate} onChange={e => handleDelegationFilterChange('plannedDate', e.target.value)}/></div>
                    <div className="filter-group">
                        <label htmlFor="del-status-filter">Task Status</label>
                        <select id="del-status-filter" value={delegationFilters.status} onChange={e => handleDelegationFilterChange('status', e.target.value)}>
                            <option value="all">All Statuses</option>
                            <option value="done">Done</option>
                            <option value="not-done">Not Done</option>
                        </select>
                    </div>
                    <button className="btn" onClick={clearDelegationFilters}>Clear</button>
                </div>
                <div className="table-container">
                    <table className="checklist-table delegation-tasks-table">
                        <thead>
                            <tr>
                                <th>Timestamp</th><th>Unique Id</th><th>Task</th><th>Assign To</th><th>Assign By</th><th>Planned Date</th><th>Actual Date</th><th>Actions</th>
                            </tr>
                        </thead>
                        <tbody>
                            {filteredDelegationTasks.map(task => (
                                editingDelegationTaskId === task.id && editedDelegationTask ? (
                                    <tr key={task.id} className={savingDelegationTaskId === task.id ? "saving-row" : "editing-row"}>
                                        <td>{task.timestamp}</td>
                                        <td>{task.taskId}</td>
                                        <td><input type="text" value={editedDelegationTask.task ?? ''} onChange={e => handleDelegationTaskFieldChange('task', e.target.value)} disabled={savingDelegationTaskId === task.id} /></td>
                                        <td><select value={editedDelegationTask.assignee} onChange={e => handleDelegationTaskFieldChange('assignee', e.target.value)} disabled={people.length === 0 || savingDelegationTaskId === task.id}>{people.map(p => <option key={`del-assignee-${p.name}`} value={p.name}>{p.name}</option>)}</select></td>
                                        <td><select value={editedDelegationTask.assigner} onChange={e => handleDelegationTaskFieldChange('assigner', e.target.value)} disabled={people.length === 0 || savingDelegationTaskId === task.id}>{people.map(p => <option key={`del-assigner-${p.name}`} value={p.name}>{p.name}</option>)}</select></td>
                                        <td><input type="date" value={getIsoDate(editedDelegationTask.plannedDate)} onChange={e => handleDelegationTaskFieldChange('plannedDate', e.target.value)} disabled={savingDelegationTaskId === task.id} /></td>
                                        <td>{task.actualDate}</td>
                                        <td className="actions-cell">
                                            {savingDelegationTaskId === task.id ? (
                                                <div className="saving-indicator">
                                                    <div className="spinner"></div>
                                                    <span>Saving...</span>
                                                </div>
                                            ) : (
                                                <>
                                                    <button className="btn btn-save" onClick={handleSaveDelegationTask}>Save</button>
                                                    <button className="btn btn-cancel" onClick={handleCancelEditDelegationTask}>Cancel</button>
                                                </>
                                            )}
                                        </td>
                                    </tr>
                                ) : (
                                    <tr key={task.id}>
                                        <td>{task.timestamp}</td>
                                        <td>{task.taskId}</td>
                                        <td>{task.task}</td>
                                        <td>{task.assignee}</td>
                                        <td>{task.assigner}</td>
                                        <td>{task.plannedDate}</td>
                                        <td>{task.actualDate}</td>
                                        <td className="actions-cell">
                                            <button 
                                                className="btn btn-edit" 
                                                onClick={() => handleEditDelegationTask(task)} 
                                                disabled={editingDelegationTaskId !== null || savingDelegationTaskId !== null || cancellingTaskId !== null || undoneDelegationTaskId !== null}
                                            >
                                                Edit
                                            </button>

                                            {(task.actualDate && task.actualDate.trim() !== '') && (
                                                <button 
                                                    className="btn btn-secondary btn-undone"
                                                    onClick={() => handleUndoneDelegationTask(task)} 
                                                    disabled={editingDelegationTaskId !== null || savingDelegationTaskId !== null || cancellingTaskId !== null || undoneDelegationTaskId !== null}
                                                >
                                                    {undoneDelegationTaskId === task.id ? (
                                                        <>
                                                            <span className="spinner" style={{ width: '1em', height: '1em', borderWidth: '2px' }}></span>
                                                            <span>Processing...</span>
                                                        </>
                                                    ) : (
                                                        'Undone'
                                                    )}
                                                </button>
                                            )}
                                            
                                            <button 
                                                className="btn btn-danger-confirm" 
                                                onClick={() => handleCancelDelegationTask(task)} 
                                                disabled={editingDelegationTaskId !== null || savingDelegationTaskId !== null || cancellingTaskId !== null || undoneDelegationTaskId !== null}
                                            >
                                                {cancellingTaskId === task.id ? (
                                                    <>
                                                        <span className="spinner" style={{ width: '1em', height: '1em', borderWidth: '2px' }}></span>
                                                        <span>Cancelling...</span>
                                                    </>
                                                ) : (
                                                    'Cancel'
                                                )}
                                            </button>
                                        </td>
                                    </tr>
                                )
                            ))}
                        </tbody>
                    </table>
                </div>
                {delegationTasks.length > 0 && filteredDelegationTasks.length === 0 && !isRefreshing && ( <div className="no-content" style={{marginTop: '20px'}}><h3>No Tasks Match Filters</h3><p>Try adjusting or clearing your filters to see more results.</p></div> )}
                {delegationTasks.length === 0 && !isRefreshing && ( <div className="no-content" style={{marginTop: '20px'}}><h3>No Delegation Tasks Found</h3><p>The "Working Task Form" sheet might be empty or data is still loading.</p></div>)}
            </div>
        </div>
    );
};
