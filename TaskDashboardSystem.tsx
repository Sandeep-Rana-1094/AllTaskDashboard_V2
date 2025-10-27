import React, { useState, useMemo, useEffect, useRef } from 'react';
import { AuthenticatedUser, DashboardTask, Person, AttendanceData, DailyAttendance, TaskHistory, Holiday } from './types';
import { parseDate, calculateWorkingDaysDelay } from './utils';

// --- HELPER FUNCTIONS ---

const formatDateToDDMMYYYY = (date: Date): string => {
    const day = String(date.getDate()).padStart(2, '0');
    const month = String(date.getMonth() + 1).padStart(2, '0'); // Month is 0-indexed
    const year = date.getFullYear();
    return `${day}/${month}/${year}`;
};

const getEmbeddableGoogleDriveUrl = (url?: string): string | undefined => {
    if (!url) return undefined;

    // If it's already a direct user content link (like lh3.googleusercontent.com), it should be fine.
    if (url.includes('googleusercontent.com')) {
        return url;
    }

    // Try to extract file ID from common Google Drive share URLs
    // e.g., https://drive.google.com/file/d/FILE_ID/view?usp=sharing
    // e.g., https://drive.google.com/open?id=FILE_ID
    const regex = /drive\.google\.com\/(?:file\/d\/|open\?id=)([a-zA-Z0-9_-]+)/;
    const match = url.match(regex);

    if (match && match[1]) {
        const fileId = match[1];
        // This is a common format for embedding
        return `https://drive.google.com/uc?export=view&id=${fileId}`;
    }

    // Return original URL if it's not a recognizable Google Drive link
    return url;
};

const getUserNameFromEmail = (email: string): string => {
    if (!email) return 'User';
    const namePart = email.split('@')[0];
    return namePart
        .replace(/[._-]/g, ' ')
        .split(' ')
        .map(word => word.charAt(0).toUpperCase() + word.slice(1))
        .join(' ');
};

const getMondayOfNWeeksAgo = (weeksAgo: number): Date => {
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const dayOfWeek = today.getDay(); // Sunday: 0, Monday: 1, ...
    const diffToMonday = dayOfWeek === 0 ? 6 : dayOfWeek - 1;
    const thisMonday = new Date(today);
    thisMonday.setDate(today.getDate() - diffToMonday);
    
    const targetMonday = new Date(thisMonday);
    targetMonday.setDate(thisMonday.getDate() - (7 * weeksAgo));
    return targetMonday;
};

const getPreviousWeekRange = () => {
    const lastMonday = getMondayOfNWeeksAgo(1);
    const lastSaturday = new Date(lastMonday);
    lastSaturday.setDate(lastMonday.getDate() + 5); // Changed to +5 to include Saturday
    return { start: lastMonday, end: lastSaturday };
};

const getPeriodDateRange = (period: string): { start: Date; end: Date } => {
    if (period === 'lastWeek') {
        return getPreviousWeekRange();
    }

    if (period === 'lastToLastWeek') {
        const lastLastMonday = getMondayOfNWeeksAgo(2);
        const lastLastFriday = new Date(lastLastMonday);
        lastLastFriday.setDate(lastLastMonday.getDate() + 4);
        return { start: lastLastMonday, end: lastLastFriday };
    }

    if (period.startsWith('year-')) {
        const year = parseInt(period.split('-')[1], 10);
        return { start: new Date(year, 0, 1), end: new Date(year, 11, 31) };
    }

    if (period.startsWith('month-')) {
        const [, yearStr, monthStr] = period.split('-');
        const year = parseInt(yearStr, 10);
        const month = parseInt(monthStr, 10); // 0-indexed month
        const start = new Date(year, month, 1);
        const end = new Date(year, month + 1, 0); // Last day of month
        return { start, end };
    }
    
    // Default fallback
    return getPreviousWeekRange();
};

const formatDateForRange = (date: Date): string => {
    const d = String(date.getDate()).padStart(2, '0');
    const m = String(date.getMonth() + 1).padStart(2, '0');
    return `${d}/${m}`;
};

const fileToBase64 = (file: File): Promise<string> => {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.readAsDataURL(file);
        reader.onload = () => {
            if (typeof reader.result !== 'string') {
                return reject(new Error('FileReader did not return a string.'));
            }
            // result is "data:mime/type;base64,ENCODED_STRING"
            const base64String = reader.result.split(',')[1];
            resolve(base64String);
        };
        reader.onerror = error => reject(error);
    });
};


// --- TYPE DEFINITIONS ---

interface TaskDashboardSystemProps {
    dashboardTasks: DashboardTask[];
    misTasks: DashboardTask[];
    isRefreshing: boolean;
    dashboardTasksError: string | null;
    misTasksError: string | null;
    authenticatedUser: AuthenticatedUser | null;
    postToGoogleSheet: (data: Record<string, any>) => Promise<any>;
    fetchData: (isInitialLoad?: boolean) => Promise<void>;
    people: Person[];
    attendanceData: AttendanceData[];
    dailyAttendanceData: DailyAttendance[];
    holidays: Holiday[];
    taskHistory: TaskHistory[];
}

// --- SVG ICONS ---

const UserIcon = () => <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" fill="currentColor" viewBox="0 0 16 16"><path d="M8 8a3 3 0 1 0 0-6 3 3 0 0 0 0 6m2-3a2 2 0 1 1-4 0 2 2 0 0 1 4 0m4 8c0 1-1 1-1 1H3s-1 0-1-1 1-4 6-4 6 3 6 4m-1-.004c-.001-.246-.154-.986-.832-1.664C11.516 10.68 10.289 10 8 10s-3.516.68-4.168 1.332c-.678.678-.83 1.418-.832 1.664z"/></svg>;
const PendingIcon = () => <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" fill="currentColor" viewBox="0 0 16 16"><path d="M3 14s-1 0-1-1 1-4 6-4 6 3 6 4-1 1-1 1zm5-6a3 3 0 1 0 0-6 3 3 0 0 0 0 6"/></svg>;
const OverdueIcon = () => <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" fill="currentColor" viewBox="0 0 16 16"><path d="M8.982 1.566a1.13 1.13 0 0 0-1.96 0L.165 13.233c-.457.778.091 1.767.98 1.767h13.713c.889 0 1.438-.99.98-1.767zM8 5c.535 0 .954.462.9.995l-.35 3.507a.552.552 0 0 1-1.1 0L7.1 5.995A.905.905 0 0 1 8 5m.002 6a1 1 0 1 1 0 2 1 1 0 0 1 0-2"/></svg>;
const TodayIcon = () => <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" fill="currentColor" viewBox="0 0 16 16"><path d="M4.684 11.523v-2.3h2.261v-.91h-2.261v-2.3H5.98v2.3h2.261v.91H5.98v2.3zM2 2a2 2 0 0 1 2-2h8a2 2 0 0 1 2 2v12a2 2 0 0 1-2-2H4a2 2 0 0 1-2-2zm2-1a1 1 0 0 0-1 1v12a1 1 0 0 0 1 1h8a1 1 0 0 0 1-1V2a1 1 0 0 0-1-1z"/></svg>;
const SearchIcon = () => <svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" fill="currentColor" viewBox="0 0 16 16"><path d="M11.742 10.344a6.5 6.5 0 1 0-1.397 1.398h-.001q.044.06.098.115l3.85 3.85a1 1 0 0 0 1.415-1.414l-3.85-3.85a1 1 0 0 0-.115-.1zM12 6.5a5.5 5.5 0 1 1-11 0 5.5 5.5 0 0 1 11 0"/></svg>;
const ArrowIcon: React.FC<{ className?: string }> = ({ className }) => <svg className={className} xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor" viewBox="0 0 16 16"><path fillRule="evenodd" d="M4.646 1.646a.5.5 0 0 1 .708 0l6 6a.5.5 0 0 1 0 .708l-6 6a.5.5 0 0 1-.708-.708L10.293 8 4.646 2.354a.5.5 0 0 1 0-.708"/></svg>;
const DashboardIcon = () => <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor" viewBox="0 0 16 16"><path d="M13.854 3.646a.5.5 0 0 1 0 .708l-7 7a.5.5 0 0 1-.708 0l-3.5-3.5a.5.5 0 1 1 .708-.708L6.5 10.293l6.646-6.647a.5.5 0 0 1 .708 0"/></svg>;
const NegativeIcon = () => <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor" viewBox="0 0 16 16"><path d="M7.002 11a1 1 0 1 1 2 0 1 1 0 0 1-2 0M7.1 4.995a.905.905 0 1 1 1.8 0l-.35 3.507a.552.552 0 0 1-1.1 0z"/></svg>;
const OnTrackIcon = () => <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor" viewBox="0 0 16 16"><path d="M16 8A8 8 0 1 1 0 8a8 8 0 0 1 16 0m-3.97-3.03a.75.75 0 0 0-1.08.022L7.477 9.417 5.384 7.323a.75.75 0 0 0-1.06 1.06L6.97 11.03a.75.75 0 0 0 1.079-.02l3.992-4.99a.75.75 0 0 0-.01-1.05z"/></svg>;
const BackArrowIcon = () => <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor" viewBox="0 0 16 16"><path fillRule="evenodd" d="M15 8a.5.5 0 0 0-.5-.5H2.707l3.147-3.146a.5.5 0 1 0-.708-.708l-4 4a.5.5 0 0 0 0 .708l4 4a.5.5 0 0 0 .708-.708L2.707 8.5H14.5A.5.5 0 0 0 15 8"/></svg>;


// --- SUB-COMPONENTS ---

const StatCard: React.FC<{
    title: string;
    value: number;
    icon: React.ReactNode;
    className?: string;
    onClick?: () => void;
    ariaPressed?: boolean;
}> = ({ title, value, icon, className = '', onClick, ariaPressed }) => (
    <div
        className={`dashboard-card stat-card ${className}`}
        onClick={onClick}
        role={onClick ? 'button' : undefined}
        tabIndex={onClick ? 0 : undefined}
        onKeyDown={onClick ? (e) => { if (e.key === 'Enter' || e.key === ' ') onClick(); } : undefined}
        aria-pressed={ariaPressed}
    >
        <div className="stat-icon">
            {icon}
        </div>
        <div className="stat-card-info">
            <div className="stat-value">{value}</div>
            <div className="stat-title">{title}</div>
        </div>
    </div>
);


const CircularProgress: React.FC<{ percentage: number; color: string; size?: number; strokeWidth?: number; }> = ({ percentage, color, size = 140, strokeWidth = 12 }) => {
    const radius = (size - strokeWidth) / 2;
    const circumference = 2 * Math.PI * radius;
    const offset = circumference - (percentage / 100) * circumference;

    return (
        <div className="progress-ring-container" style={{ width: size, height: size }}>
            <svg className="progress-ring" width={size} height={size}>
                <circle cx={size / 2} cy={size / 2} r={radius} stroke="#e5e7eb" strokeWidth={strokeWidth} fill="transparent" />
                <circle
                    className="progress-ring__circle"
                    cx={size / 2} cy={size / 2} r={radius}
                    stroke={color} strokeWidth={strokeWidth}
                    fill="transparent"
                    strokeDasharray={circumference}
                    strokeDashoffset={offset}
                    strokeLinecap="round"
                />
            </svg>
            <div className="progress-ring-text" style={{ color }}>{percentage}<span>%</span></div>
        </div>
    );
};

const AttachmentModal: React.FC<{
    task: DashboardTask | null; onClose: () => void;
    onSubmit: (file: File) => void; isSubmitting: boolean;
}> = ({ task, onClose, onSubmit, isSubmitting }) => {
    const [selectedFile, setSelectedFile] = useState<File | null>(null);
    useEffect(() => { setSelectedFile(null); }, [task]);
    if (!task) return null;
    const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
        const file = e.target.files ? e.target.files[0] : null;
        if (file) {
            if (file.size > 5 * 1024 * 1024) {
                alert("File is too large. Please select a file smaller than 5MB.");
                setSelectedFile(null); e.target.value = '';
            } else { setSelectedFile(file); }
        }
    };
    const handleSubmit = () => { if (selectedFile) onSubmit(selectedFile); };
    return (
        <div className="modal-overlay" onClick={onClose} role="dialog" aria-modal="true" aria-labelledby="attachment-modal-title">
            <div className="modal-content" onClick={e => e.stopPropagation()}>
                <h2 id="attachment-modal-title">Submit Task with Attachment</h2>
                <p><strong>Task:</strong> {task.task}</p>
                <div className="modal-form">
                    <div className="form-group">
                        <label htmlFor="attachment-file">Upload Document (Required)</label>
                        <input id="attachment-file" type="file" onChange={handleFileChange} required />
                        {selectedFile && <p style={{ marginTop: '8px', fontSize: '0.9em', color: '#6b7280' }}>Selected: {selectedFile.name}</p>}
                    </div>
                    <div className="modal-actions">
                        <button type="button" className="btn btn-secondary" onClick={onClose} disabled={isSubmitting}>Cancel</button>
                        <button type="button" className="btn btn-primary" onClick={handleSubmit} disabled={!selectedFile || isSubmitting}>{isSubmitting ? 'Submitting...' : 'Done'}</button>
                    </div>
                </div>
            </div>
        </div>
    );
};

const SelectedDateModal = ({ date, tasks, onClose }: { date: Date | null, tasks: DashboardTask[], onClose: () => void }) => {
    if (!date) return null;

    const today = new Date();
    today.setHours(0, 0, 0, 0);

    const isPastDate = date.getTime() < today.getTime();

    const completedTasks = tasks.filter(task => task.actual && task.actual.trim() !== '');
    const pendingOrMissedTasks = tasks.filter(task => !task.actual || task.actual.trim() === '');
    
    const pendingLabel = isPastDate ? 'Missed' : 'Pending';

    return (
        <div className="modal-overlay" onClick={onClose}>
            <div className="modal-content calendar-task-modal" onClick={e => e.stopPropagation()}>
                <div className="delegation-modal-header">
                    <h2>Tasks for {date.toLocaleDateString()}</h2>
                    <button onClick={onClose} className="btn-close-modal" aria-label="Close modal">&times;</button>
                </div>

                <div className="calendar-task-modal-summary">
                    <div className="summary-item">All Tasks: <span>{tasks.length}</span></div>
                    <div className="summary-item">Completed: <span>{completedTasks.length}</span></div>
                    <div className="summary-item">{pendingLabel}: <span>{pendingOrMissedTasks.length}</span></div>
                </div>

                <div className="calendar-task-modal-content">
                    {pendingOrMissedTasks.length > 0 && (
                        <div className="calendar-task-modal-section">
                            <h3 className="task-list-title">{pendingLabel} Tasks ({pendingOrMissedTasks.length})</h3>
                            <ul className="calendar-task-list">
                                {pendingOrMissedTasks.map(task => (
                                    <li key={task.id}>
                                        <span className={`task-status ${isPastDate ? 'missed' : 'pending'}`}></span>
                                        <div className="task-info">
                                            <strong>{task.taskId}</strong>: {task.task}
                                        </div>
                                    </li>
                                ))}
                            </ul>
                        </div>
                    )}

                    {completedTasks.length > 0 && (
                        <div className="calendar-task-modal-section">
                            <h3 className="task-list-title">Completed Tasks ({completedTasks.length})</h3>
                            <ul className="calendar-task-list">
                                {completedTasks.map(task => (
                                    <li key={task.id}>
                                        <span className="task-status completed"></span>
                                        <div className="task-info">
                                            <strong>{task.taskId}</strong>: {task.task}
                                        </div>
                                    </li>
                                ))}
                            </ul>
                        </div>
                    )}

                    {tasks.length === 0 && (
                        <p className="no-tasks-message">No tasks planned for this day.</p>
                    )}
                </div>
            </div>
        </div>
    );
};

interface CalendarViewProps {
    calendarDate: Date;
    setCalendarDate: React.Dispatch<React.SetStateAction<Date>>;
    calendarMode: 'month' | 'week';
    setCalendarMode: React.Dispatch<React.SetStateAction<'month' | 'week'>>;
    tasksByDate: Map<string, DashboardTask[]>;
    setSelectedCalendarDate: React.Dispatch<React.SetStateAction<Date | null>>;
    userAttendanceByDate: Map<string, string>;
    isAdmin: boolean;
    holidays: Holiday[];
}

const CalendarView: React.FC<CalendarViewProps> = ({
    calendarDate,
    setCalendarDate,
    calendarMode,
    setCalendarMode,
    tasksByDate,
    setSelectedCalendarDate,
    userAttendanceByDate,
    isAdmin,
    holidays,
}) => {
    const today = new Date();

    const holidaysMap = useMemo(() => {
        const map = new Map<string, string>();
        holidays.forEach(h => {
            const d = parseDate(h.date);
            if (d) {
                map.set(d.toDateString(), h.name);
            }
        });
        return map;
    }, [holidays]);

    const handlePrev = () => {
        setCalendarDate(current => {
            const newDate = new Date(current);
            if (calendarMode === 'month') newDate.setMonth(newDate.getMonth() - 1);
            else newDate.setDate(newDate.getDate() - 7);
            return newDate;
        });
    };

    const handleNext = () => {
        setCalendarDate(current => {
            const newDate = new Date(current);
            if (calendarMode === 'month') newDate.setMonth(newDate.getMonth() + 1);
            else newDate.setDate(newDate.getDate() + 7);
            return newDate;
        });
    };

    const getCalendarDays = () => {
        const days = [];
        const year = calendarDate.getFullYear();
        const month = calendarDate.getMonth();

        if (calendarMode === 'month') {
            const firstDayOfMonth = new Date(year, month, 1).getDay(); // Sunday: 0, Monday: 1
            const daysToPad = (firstDayOfMonth === 0) ? 6 : firstDayOfMonth - 1;
            
            const daysInMonth = new Date(year, month + 1, 0).getDate();
            const daysInPrevMonth = new Date(year, month, 0).getDate();

            for (let i = daysToPad - 1; i >= 0; i--) {
                days.push(new Date(year, month - 1, daysInPrevMonth - i));
            }
            for (let i = 1; i <= daysInMonth; i++) {
                days.push(new Date(year, month, i));
            }
            const remainingCells = 42 - days.length; // 6 weeks grid
            for (let i = 1; i <= remainingCells; i++) {
                days.push(new Date(year, month + 1, i));
            }
        } else { // week mode
            const dayOfWeek = calendarDate.getDay(); // Sunday: 0, Monday: 1
            const diff = (dayOfWeek === 0) ? 6 : dayOfWeek - 1;
            const startDate = new Date(calendarDate);
            startDate.setDate(calendarDate.getDate() - diff);
            for (let i = 0; i < 7; i++) {
                const weekDay = new Date(startDate);
                weekDay.setDate(startDate.getDate() + i);
                days.push(weekDay);
            }
        }
        return days;
    };

    const calendarDays = getCalendarDays();

    const { isPrevDisabled, isNextDisabled } = useMemo(() => {
        if (isAdmin || calendarMode !== 'week') {
            return { isPrevDisabled: false, isNextDisabled: false };
        }

        const currentMonth = today.getMonth();
        const currentYear = today.getFullYear();

        if (calendarDays.length !== 7) {
            return { isPrevDisabled: false, isNextDisabled: false };
        }

        const startOfWeek = calendarDays[0];
        const endOfPrevWeek = new Date(startOfWeek.getTime() - 24 * 60 * 60 * 1000);
        const prevDisabled = endOfPrevWeek.getFullYear() < currentYear || 
                            (endOfPrevWeek.getFullYear() === currentYear && endOfPrevWeek.getMonth() < currentMonth);

        const endOfWeek = calendarDays[6];
        const startOfNextWeek = new Date(endOfWeek.getTime() + 24 * 60 * 60 * 1000);
        const nextDisabled = startOfNextWeek.getFullYear() > currentYear ||
                            (startOfNextWeek.getFullYear() === currentYear && startOfNextWeek.getMonth() > currentMonth);
        
        return { isPrevDisabled: prevDisabled, isNextDisabled: nextDisabled };

    }, [isAdmin, calendarMode, calendarDays, today]);

    const showNavButtons = isAdmin || calendarMode === 'week';
    const weekDays = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"];

    const getHeaderTitle = () => {
        if (calendarMode === 'month') {
            return calendarDate.toLocaleString('default', { month: 'long', year: 'numeric' });
        } else {
            const startOfWeek = calendarDays[0];
            const endOfWeek = calendarDays[6];
            return `${startOfWeek.toLocaleDateString()} - ${endOfWeek.toLocaleDateString()}`;
        }
    };

    return (
        <div className="dashboard-card calendar-view">
            <div className="calendar-header">
                <div className="calendar-nav">
                    { showNavButtons ?
                        <button onClick={handlePrev} disabled={isPrevDisabled}>&lt;</button> :
                        <div style={{width: '36px', flexShrink: 0}}></div>
                    }
                    <h3>{getHeaderTitle()}</h3>
                    { showNavButtons ?
                        <button onClick={handleNext} disabled={isNextDisabled}>&gt;</button> :
                        <div style={{width: '36px', flexShrink: 0}}></div>
                    }
                </div>
                <div className="calendar-mode-switcher">
                    <button onClick={() => setCalendarMode('month')} className={calendarMode === 'month' ? 'active' : ''}>Month</button>
                    <button onClick={() => setCalendarMode('week')} className={calendarMode === 'week' ? 'active' : ''}>Week</button>
                </div>
            </div>
            <div className="calendar-grid">
                {weekDays.map(day => <div key={day} className="calendar-day-header">{day}</div>)}
                {calendarDays.map((day, index) => {
                    const dateKey = `${day.getFullYear()}-${String(day.getMonth() + 1).padStart(2, '0')}-${String(day.getDate()).padStart(2, '0')}`;
                    const tasksOnDay = tasksByDate.get(dateKey) || [];
                    const isToday = day.toDateString() === today.toDateString();
                    const isOtherMonth = calendarMode === 'month' && day.getMonth() !== calendarDate.getMonth();
                    
                    const classNames = ['calendar-day'];
                    if (isToday) classNames.push('today');
                    if (isOtherMonth) classNames.push('other-month');
                    if (tasksOnDay.length > 0) classNames.push('has-tasks');

                    const dayStart = new Date(day);
                    dayStart.setHours(0, 0, 0, 0);

                    const todayStart = new Date();
                    todayStart.setHours(0, 0, 0, 0);
                    
                    const attendanceStatus = userAttendanceByDate.get(dateKey);
                    
                    const dayOfWeek = day.getDay(); // 0 = Sunday, 6 = Saturday
                    const isWeekend = dayOfWeek === 0 || dayOfWeek === 6;
                    const holidayName = holidaysMap.get(day.toDateString());

                    let content = null;

                    if (holidayName) {
                        content = (
                            <div className="task-summary-container">
                                <div className="task-summary-item task-summary-weekoff">
                                    <span className="label">{holidayName}</span>
                                </div>
                            </div>
                        );
                    } else if (attendanceStatus && attendanceStatus.toLowerCase() !== 'present') {
                        content = (
                            <div className="task-summary-container">
                                <div className="task-summary-item task-summary-leave">
                                    <span className="label">{attendanceStatus}</span>
                                </div>
                            </div>
                        );
                    } else if (isWeekend) {
                        content = (
                            <div className="task-summary-container">
                                <div className="task-summary-item task-summary-weekoff">
                                    <span className="label">Week off</span>
                                </div>
                            </div>
                        );
                    } else if (!attendanceStatus && dayStart.getTime() <= todayStart.getTime()) {
                         content = (
                            <div className="task-summary-container">
                                <div className="task-summary-item task-summary-unmarked">
                                    <span className="label">Attendance Not Marked</span>
                                </div>
                            </div>
                        );
                    }

                    if (!content && tasksOnDay.length > 0) {
                        const completedCount = tasksOnDay.filter(t => t.actual && t.actual.trim() !== '').length;
                        const pendingCount = tasksOnDay.length - completedCount;

                        if (dayStart.getTime() < todayStart.getTime()) { // Past day
                            content = (
                                <div className="task-summary-container">
                                    <div className="task-summary-item task-summary-planned">
                                        <span className="label">All Tasks</span>
                                        <span className="count">{tasksOnDay.length}</span>
                                    </div>
                                    {completedCount > 0 && (
                                        <div className="task-summary-item task-summary-completed">
                                            <span className="label">Completed</span>
                                            <span className="count">{completedCount}</span>
                                        </div>
                                    )}
                                    {pendingCount > 0 && (
                                        <div className="task-summary-item task-summary-missed">
                                            <span className="label">Missed</span>
                                            <span className="count">{pendingCount}</span>
                                        </div>
                                    )}
                                </div>
                            );
                        } else if (dayStart.getTime() === todayStart.getTime()) { // Today
                             content = (
                                <div className="task-summary-container">
                                    <div className="task-summary-item task-summary-planned">
                                        <span className="label">All Tasks</span>
                                        <span className="count">{tasksOnDay.length}</span>
                                    </div>
                                    {completedCount > 0 && (
                                        <div className="task-summary-item task-summary-completed">
                                            <span className="label">Completed</span>
                                            <span className="count">{completedCount}</span>
                                        </div>
                                    )}
                                    {pendingCount > 0 && (
                                        <div className="task-summary-item task-summary-pending">
                                            <span className="label">Pending</span>
                                            <span className="count">{pendingCount}</span>
                                        </div>
                                    )}
                                </div>
                            );
                        } else { // Future day
                            content = (
                                <div className="task-summary-container">
                                    {tasksOnDay.length > 0 && (
                                        <div className="task-summary-item task-summary-planned">
                                            <span className="label">Planned</span>
                                            <span className="count">{tasksOnDay.length}</span>
                                        </div>
                                    )}
                                </div>
                            );
                        }
                    }

                    return (
                        <div key={index} className={classNames.join(' ')} onClick={() => setSelectedCalendarDate(day)}>
                            <span className="day-number">{day.getDate()}</span>
                            {content}
                        </div>
                    );
                })}
            </div>
        </div>
    );
};

const TaskHistoryModal: React.FC<{
    task: DashboardTask;
    history: TaskHistory[];
    onClose: () => void;
}> = ({ task, history, onClose }) => {
    const relevantHistory = useMemo(() => {
        if (!task.taskId) return [];
        return history
            .filter(h => h.task.includes(task.taskId))
            .sort((a, b) => (new Date(b.timestamp).getTime()) - (new Date(a.timestamp).getTime()));
    }, [task, history]);

    return (
        <div className="modal-overlay" onClick={onClose}>
            <div className="modal-content history-modal-content" onClick={e => e.stopPropagation()}>
                <div className="delegation-modal-header">
                    <h2>History for Task: {task.taskId}</h2>
                    <button onClick={onClose} className="btn-close-modal" aria-label="Close modal">&times;</button>
                </div>
                <div className="history-modal-body">
                    <p className="history-task-description"><strong>Task:</strong> {task.task}</p>
                    {relevantHistory.length > 0 ? (
                        <ul className="history-list">
                            {relevantHistory.map((entry, index) => (
                                <li key={index} className="history-item">
                                    <div className="history-item-meta">
                                        <span className="history-timestamp">{new Date(entry.timestamp).toLocaleString()}</span>
                                        <span className="history-user">{entry.changedBy}</span>
                                    </div>
                                    <p className="history-change">{entry.change}</p>
                                </li>
                            ))}
                        </ul>
                    ) : (
                        <p className="no-history-message">No history found for this task.</p>
                    )}
                </div>
            </div>
        </div>
    );
};

// --- MAIN COMPONENT ---

export const TaskDashboardSystem: React.FC<TaskDashboardSystemProps> = ({
    dashboardTasks,
    misTasks,
    isRefreshing,
    dashboardTasksError,
    misTasksError,
    authenticatedUser,
    postToGoogleSheet,
    fetchData,
    people,
    attendanceData,
    dailyAttendanceData,
    holidays,
    taskHistory,
}) => {
    const isAdmin = authenticatedUser?.role === 'Admin';
    const [dashboardMode, setDashboardMode] = useState<'myDashboard' | 'employeeMIS'>('myDashboard');
    const [searchTerm, setSearchTerm] = useState('');
    const [filterMode, setFilterMode] = useState<'all' | 'overdue' | 'today'>('all');
    const [selectedTaskIds, setSelectedTaskIds] = useState<Set<string>>(new Set());
    const [isSubmitting, setIsSubmitting] = useState(false);
    const [attachmentModalTask, setAttachmentModalTask] = useState<DashboardTask | null>(null);
    const [lastUpdatedTime] = useState(new Date().toLocaleTimeString([], { hour: 'numeric', minute: '2-digit', hour12: true }));
    const [inFlightTaskIds, setInFlightTaskIds] = useState<Set<string>>(new Set());
    const [expandedKpi, setExpandedKpi] = useState<'notDone' | 'notOnTime' | null>(null);

    // --- New State for Calendar View ---
    const [currentView, setCurrentView] = useState<'stats' | 'calendar'>('stats');
    const [calendarDate, setCalendarDate] = useState(new Date());
    const [calendarMode, setCalendarMode] = useState<'month' | 'week'>('week');
    const [selectedCalendarDate, setSelectedCalendarDate] = useState<Date | null>(null);
    const [historyModalTask, setHistoryModalTask] = useState<DashboardTask | null>(null);

    // --- Employee MIS State ---
    const [selectedMisEmployeeName, setSelectedMisEmployeeName] = useState<string | null>(null);
    const [selectedMisPeriod, setSelectedMisPeriod] = useState<string>('lastWeek');
    const reportRef = useRef<HTMLDivElement>(null);

    // Statuses that count as being present for work
    const PRESENT_STATUSES = ['present', 'on official travel', 'work from home'];

    useEffect(() => {
        if (selectedMisEmployeeName) {
            const timer = setTimeout(() => {
                reportRef.current?.scrollIntoView({ behavior: 'smooth', block: 'start' });
            }, 100);
            return () => clearTimeout(timer);
        }
    }, [selectedMisEmployeeName]);

    // Filter out tasks planned on Sundays from all calculations for "My Dashboard".
    const weekdayTasks = useMemo(() => {
        return dashboardTasks.filter(task => {
            const plannedDate = parseDate(task.planned);
            if (!plannedDate) return true; // Keep if date is invalid
            const day = plannedDate.getDay();
            return day !== 0; // Exclude Sunday (0)
        });
    }, [dashboardTasks]);
    
    // Filter out tasks planned on Sundays from all calculations for "Employee MIS".
    const misWeekdayTasks = useMemo(() => {
        return misTasks.filter(task => {
            const plannedDate = parseDate(task.planned);
            if (!plannedDate) return true; // Keep if date is invalid
            const day = plannedDate.getDay();
            return day !== 0; // Exclude Sunday (0)
        });
    }, [misTasks]);

    useEffect(() => {
        if (!isAdmin) {
            setDashboardMode('myDashboard');
        }
    }, [isAdmin]);
    
    useEffect(() => {
        // When the task list is updated from the parent, clean up the in-flight set.
        setInFlightTaskIds(currentInFlightIds => {
            const newInFlightIds = new Set<string>();
            if (currentInFlightIds.size === 0) return currentInFlightIds;

            const currentDashboardTaskIds = new Set(dashboardTasks.map(t => t.id));
            for (const id of currentInFlightIds) {
                // If a task that was in-flight is STILL in the main list (i.e., pending), keep it in-flight.
                // If it's gone, it means it was processed, so we can remove it.
                if (currentDashboardTaskIds.has(id)) {
                    const task = dashboardTasks.find(t => t.id === id);
                    if (!task || (task && (!task.actual || task.actual.trim() === ''))) {
                       newInFlightIds.add(id);
                    }
                }
            }
            return newInFlightIds;
        });
    }, [dashboardTasks]);

    const userEmail = authenticatedUser?.mailId || 'no-email@provided.com';

    // Define the system types that allow "Mark Done" functionality.
    const ALLOWED_SYSTEM_TYPES_FOR_SUBMIT = ['New_Checklist', 'Delegation On Demand'];

    const { userName, photoUrl } = useMemo(() => {
        const userEmailLower = userEmail.toLowerCase();
        const userTaskWithInfo = dashboardTasks.find(
            task => task.userEmail?.toLowerCase() === userEmailLower
        );
        const personInfo = people.find(p => p.email?.toLowerCase() === userEmailLower);
        const derivedUserName = 
            userTaskWithInfo?.userName?.trim() ||
            personInfo?.name ||
            getUserNameFromEmail(userEmail);
        const rawPhotoUrl = 
            userTaskWithInfo?.photoUrl?.trim() ||
            personInfo?.photoUrl;
        
        return {
            userName: derivedUserName,
            photoUrl: getEmbeddableGoogleDriveUrl(rawPhotoUrl),
        };
    }, [people, userEmail, dashboardTasks]);

    // --- My Dashboard Calculations ---
    const myAttendanceBreakdown = useMemo(() => {
        const { start: periodStart, end: periodEnd } = getPreviousWeekRange();

        const userEmailLower = userEmail.toLowerCase();
        const userNameLower = userName.toLowerCase();

        const userAttendanceRecords = dailyAttendanceData.filter(att => {
            const attEmailLower = (att.email || '').toLowerCase();
            if (userEmailLower && attEmailLower) {
                return attEmailLower === userEmailLower && att.date;
            }
            return att.name.toLowerCase() === userNameLower && att.date;
        });

        const firstAttendanceDate = userAttendanceRecords
            .map(att => parseDate(att.date))
            .filter((d): d is Date => d !== null)
            .sort((a, b) => a.getTime() - b.getTime())[0];

        const holidayDates = new Set(
            holidays.map(h => parseDate(h.date)?.toDateString()).filter(Boolean)
        );

        const today = new Date(); today.setHours(0, 0, 0, 0);

        if (!firstAttendanceDate) {
            let wdCount = 0;
            const d = new Date(periodStart);
            while (d <= periodEnd && d <= today) {
                const dayOfWeek = d.getDay();
                const isWeekend = dayOfWeek === 0; // Saturday is now a working day for "last week"
                const isHoliday = holidayDates.has(d.toDateString());
                if (!isWeekend && !isHoliday) {
                    wdCount++;
                }
                d.setDate(d.getDate() + 1);
            }
            return {
                workingDays: wdCount,
                daysPresent: 0,
                attendancePercentage: 0,
                notMarked: wdCount,
                otherStatusesBreakdown: [],
            };
        }
        
        firstAttendanceDate.setHours(0, 0, 0, 0);

        const attendanceMap = new Map<string, string>();
        userAttendanceRecords.forEach(att => {
            const d = parseDate(att.date);
            if (d) {
                attendanceMap.set(d.toDateString(), att.status);
            }
        });

        let workingDaysCount = 0;
        let notMarkedCount = 0;
        const statusCounts = new Map<string, number>();

        const loopDate = new Date(periodStart);

        while (loopDate <= periodEnd && loopDate <= today) {
            if (loopDate >= firstAttendanceDate) {
                const dayOfWeek = loopDate.getDay();
                const isWeekend = dayOfWeek === 0; // Saturday is now a working day for "last week"
                const isHoliday = holidayDates.has(loopDate.toDateString());

                if (!isWeekend && !isHoliday) {
                    workingDaysCount++;
                    const status = attendanceMap.get(loopDate.toDateString());

                    if (status) {
                        const currentCount = statusCounts.get(status) || 0;
                        statusCounts.set(status, currentCount + 1);
                    } else {
                        notMarkedCount++;
                    }
                }
            }
            loopDate.setDate(loopDate.getDate() + 1);
        }
        
        let presentCount = 0;
        const otherStatusesBreakdown: [string, number][] = [];

        for (const [status, count] of statusCounts.entries()) {
            if (PRESENT_STATUSES.includes(status.toLowerCase())) {
                presentCount += count;
            } else {
                otherStatusesBreakdown.push([status, count]);
            }
        }
        otherStatusesBreakdown.sort((a, b) => a[0].localeCompare(b[0]));
        
        const percentage = workingDaysCount > 0 ? Math.min(Math.round((presentCount / workingDaysCount) * 100), 100) : 0;
        
        return {
            workingDays: workingDaysCount,
            daysPresent: presentCount,
            attendancePercentage: percentage,
            notMarked: notMarkedCount,
            otherStatusesBreakdown: otherStatusesBreakdown,
        };
    }, [dailyAttendanceData, holidays, userEmail, userName, PRESENT_STATUSES]);

    const userTasks = useMemo(() => {
        return weekdayTasks.filter(task => task.userEmail?.toLowerCase() === userEmail.toLowerCase());
    }, [weekdayTasks, userEmail]);

    const {
        pendingTasks, overdueTasks, dueTodayTasks, prevWeekDateRange,
        planVsActual_Planned, planVsActual_Actual, planVsActual_Percent,
        onTime_Planned, onTime_Actual, onTime_Percent,
        notDoneTasksForPrevWeek, notOnTimeTasksForPrevWeek
    } = useMemo(() => {
        const today = new Date();
        const todayStart = new Date(today.getFullYear(), today.getMonth(), today.getDate());

        const pending = userTasks.filter(task => {
            const isNotCompleted = !task.actual || task.actual.trim() === '';
            if (!isNotCompleted) return false;
            const plannedDate = parseDate(task.planned);
            if (!plannedDate) return true;
            plannedDate.setHours(0, 0, 0, 0);
            return plannedDate.getTime() <= todayStart.getTime();
        });

        const overdue = pending.filter(task => {
            const plannedDate = parseDate(task.planned);
            if (!plannedDate) return false;
            plannedDate.setHours(0, 0, 0, 0);
            return plannedDate.getTime() < todayStart.getTime();
        });

        const dueToday = pending.filter(task => {
            const plannedDate = parseDate(task.planned);
            if (!plannedDate) return false;
            plannedDate.setHours(0, 0, 0, 0);
            return plannedDate.getTime() === todayStart.getTime();
        });

        const { start: prevWeekStart, end: prevWeekEnd } = getPreviousWeekRange();
        const dateRangeStr = `(${formatDateForRange(prevWeekStart)} - ${formatDateForRange(prevWeekEnd)})`;

        const prevWeekTasks = userTasks.filter(task => {
            const plannedDate = parseDate(task.planned);
            if (!plannedDate) return false;
            plannedDate.setHours(0, 0, 0, 0);
            return plannedDate.getTime() >= prevWeekStart.getTime() && plannedDate.getTime() <= prevWeekEnd.getTime();
        });

        const tasksCompletedFromPrevWeek = prevWeekTasks.filter(t => !!parseDate(t.actual));
        const notDoneTasks = prevWeekTasks.filter(t => !tasksCompletedFromPrevWeek.includes(t));

        const planVsActual_Planned_Count = prevWeekTasks.length;
        const planVsActual_Done_Count = tasksCompletedFromPrevWeek.length;
        const planVsActual_NotDone_Count = planVsActual_Planned_Count - planVsActual_Done_Count;
        const planVsActual_NotDone_Percent = planVsActual_Planned_Count > 0 ? Math.round((planVsActual_NotDone_Count / planVsActual_Planned_Count) * 100) : 0;

        const onTime_Planned_Count = tasksCompletedFromPrevWeek.length;
        const onTime_Actual_Tasks = tasksCompletedFromPrevWeek.filter(t => {
            const plannedDate = parseDate(t.planned);
            const actualDate = parseDate(t.actual!);
            if (!plannedDate || !actualDate) return false;
            // A task is on time if the working days delay is 0.
            // For "last week", Saturday is a working day.
            return calculateWorkingDaysDelay(plannedDate, actualDate, holidays, { isSaturdayWorkday: true }) === 0;
        });
        const onTime_Actual_Count = onTime_Actual_Tasks.length;
        const notOnTimeTasks = tasksCompletedFromPrevWeek.filter(t => !onTime_Actual_Tasks.includes(t));

        const onTime_NotDone_Count = onTime_Planned_Count - onTime_Actual_Count;
        const onTime_NotDone_Percent = onTime_Planned_Count > 0 ? Math.round((onTime_NotDone_Count / onTime_Planned_Count) * 100) : 0;

        return {
            pendingTasks: pending, overdueTasks: overdue, dueTodayTasks: dueToday, prevWeekDateRange: dateRangeStr,
            planVsActual_Planned: planVsActual_Planned_Count, planVsActual_Actual: planVsActual_Done_Count, planVsActual_Percent: planVsActual_NotDone_Percent,
            onTime_Planned: onTime_Planned_Count, onTime_Actual: onTime_Actual_Count, onTime_Percent: onTime_NotDone_Percent,
            notDoneTasksForPrevWeek: notDoneTasks,
            notOnTimeTasksForPrevWeek: notOnTimeTasks,
        };
    }, [userTasks, holidays]);
    
    // --- Employee MIS Calculations ---
    const { onTrackEmployees, negativeScoreEmployees } = useMemo(() => {
        const { start: prevWeekStart, end: prevWeekEnd } = getPreviousWeekRange();
        
        const prevWeekTasks = misWeekdayTasks.filter(task => {
            const plannedDate = parseDate(task.planned);
            if (!plannedDate) return false;
            plannedDate.setHours(0, 0, 0, 0);
            return plannedDate.getTime() >= prevWeekStart.getTime() && plannedDate.getTime() <= prevWeekEnd.getTime();
        });

        const tasksByUserName = prevWeekTasks.reduce((acc, task) => {
            const name = task.userName?.trim();
            if (name) {
                if (!acc[name]) {
                    acc[name] = [];
                }
                acc[name].push(task);
            }
            return acc;
        }, {} as Record<string, DashboardTask[]>);
        
        const onTrack: Person[] = [];
        const negativeScore: Person[] = [];

        for (const name in tasksByUserName) {
            const userTasks = tasksByUserName[name];
            if (userTasks.length === 0) continue;

            const isNegative = userTasks.some(task => {
                const actualDate = parseDate(task.actual);
                if (!actualDate) {
                    return true; // Not completed -> Negative score
                }
                
                const plannedDate = parseDate(task.planned);
                if (plannedDate) {
                    plannedDate.setHours(0, 0, 0, 0);
                    actualDate.setHours(0, 0, 0, 0);
                    if (actualDate.getTime() > plannedDate.getTime()) {
                        return true; // Completed late -> Negative score
                    }
                }
                return false; // Task was completed on time.
            });
            
            const personInfo = people.find(p => p.name === name) || { name };

            if (isNegative) {
                negativeScore.push(personInfo);
            } else {
                onTrack.push(personInfo);
            }
        }
        
        onTrack.sort((a, b) => a.name.localeCompare(b.name));
        negativeScore.sort((a, b) => a.name.localeCompare(b.name));
        
        return { onTrackEmployees: onTrack, negativeScoreEmployees: negativeScore };
    }, [misWeekdayTasks, people]);

    const misReportData = useMemo(() => {
        if (!selectedMisEmployeeName) return null;

        const selectedNameLower = selectedMisEmployeeName.toLowerCase();
        const personInfo = people.find(p => p.name.toLowerCase() === selectedNameLower);
        const taskInfo = misWeekdayTasks.find(t => t.userName.toLowerCase() === selectedNameLower);

        const email = (personInfo?.email || taskInfo?.userEmail || '').toLowerCase();
        const photoUrl = getEmbeddableGoogleDriveUrl(personInfo?.photoUrl || taskInfo?.photoUrl);
        
        const { start: periodStart, end: periodEnd } = getPeriodDateRange(selectedMisPeriod);
        const dateRangeStr = `(${formatDateForRange(periodStart)} - ${formatDateForRange(periodEnd)})`;

        // --- DYNAMIC ATTENDANCE CALCULATION ---
        const attendanceBreakdown = (() => {
            const isLastWeekReport = selectedMisPeriod === 'lastWeek';
            // Find all attendance records for the selected employee
            const employeeAttendanceRecords = dailyAttendanceData.filter(att => {
                // Prioritize matching by email if available, as it's more reliable
                if (email && att.email) {
                    return att.email === email && att.date;
                }
                // Fallback to name matching if emails aren't available for some reason
                return att.name.toLowerCase() === selectedNameLower && att.date;
            });

            // Find the employee's very first attendance date to determine their start date
            const firstAttendanceDate = employeeAttendanceRecords
                .map(att => parseDate(att.date))
                .filter((d): d is Date => d !== null)
                .sort((a, b) => a.getTime() - b.getTime())[0]; 

            const holidayDates = new Set(
                holidays.map(h => parseDate(h.date)?.toDateString()).filter(Boolean)
            );

            // If there are no attendance records at all for this employee
            if (!firstAttendanceDate) {
                let wdCount = 0;
                const d = new Date(periodStart);
                const today = new Date(); today.setHours(0, 0, 0, 0);
                while (d <= periodEnd && d <= today) {
                    const dayOfWeek = d.getDay();
                    const isWeekend = isLastWeekReport ? (dayOfWeek === 0) : (dayOfWeek === 0 || dayOfWeek === 6);
                    const isHoliday = holidayDates.has(d.toDateString());
                    if (!isWeekend && !isHoliday) {
                        wdCount++;
                    }
                    d.setDate(d.getDate() + 1);
                }
                return {
                    workingDays: wdCount,
                    daysPresent: 0,
                    attendancePercentage: 0,
                    notMarked: wdCount,
                    otherStatusesBreakdown: [],
                };
            }
            
            firstAttendanceDate.setHours(0, 0, 0, 0);

            // Map attendance records by date string for O(1) lookup
            const attendanceMap = new Map<string, string>();
            employeeAttendanceRecords.forEach(att => {
                const d = parseDate(att.date);
                if (d) {
                    attendanceMap.set(d.toDateString(), att.status);
                }
            });

            let workingDaysCount = 0;
            let notMarkedCount = 0;
            const statusCounts = new Map<string, number>();

            const loopDate = new Date(periodStart);
            const today = new Date(); today.setHours(0, 0, 0, 0);

            while (loopDate <= periodEnd && loopDate <= today) {
                // Only start counting from the employee's first recorded day
                if (loopDate < firstAttendanceDate) {
                    loopDate.setDate(loopDate.getDate() + 1);
                    continue;
                }

                const dayOfWeek = loopDate.getDay();
                const isWeekend = isLastWeekReport ? (dayOfWeek === 0) : (dayOfWeek === 0 || dayOfWeek === 6);
                const isHoliday = holidayDates.has(loopDate.toDateString());

                if (!isWeekend && !isHoliday) {
                    workingDaysCount++;
                    const status = attendanceMap.get(loopDate.toDateString());

                    if (status) {
                        const currentCount = statusCounts.get(status) || 0;
                        statusCounts.set(status, currentCount + 1);
                    } else {
                        notMarkedCount++;
                    }
                }
                loopDate.setDate(loopDate.getDate() + 1);
            }
            
            let presentCount = 0;
            const otherStatusesBreakdown: [string, number][] = [];

            for (const [status, count] of statusCounts.entries()) {
                if (PRESENT_STATUSES.includes(status.toLowerCase())) {
                    presentCount += count;
                } else {
                    otherStatusesBreakdown.push([status, count]);
                }
            }
            otherStatusesBreakdown.sort((a,b) => a[0].localeCompare(b[0]));
            
            const percentage = workingDaysCount > 0 ? Math.min(Math.round((presentCount / workingDaysCount) * 100), 100) : 0;
            
            return {
                workingDays: workingDaysCount,
                daysPresent: presentCount,
                attendancePercentage: percentage,
                notMarked: notMarkedCount,
                otherStatusesBreakdown: otherStatusesBreakdown,
            };
        })();


        // --- Performance calculation (dynamic based on selected period) ---
        const today = new Date();
        today.setHours(0, 0, 0, 0);
        const isLastWeekReport = selectedMisPeriod === 'lastWeek';
        
        const employeeTasksForPeriod = misWeekdayTasks.filter(task => {
            if (task.userName.toLowerCase() !== selectedNameLower) return false;
            
            const plannedDate = parseDate(task.planned);
            if (!plannedDate) return false;
            
            plannedDate.setHours(0, 0, 0, 0);
            
            // The planned date must be on or before today to be considered for MIS.
            if (plannedDate.getTime() > today.getTime()) {
                return false;
            }

            return plannedDate.getTime() >= periodStart.getTime() && plannedDate.getTime() <= periodEnd.getTime();
        });

        const notDoneTasks = employeeTasksForPeriod.filter(t => !t.actual || t.actual.trim() === '');
        const tasksCompleted = employeeTasksForPeriod.filter(t => !notDoneTasks.includes(t));
        const lateTasks = tasksCompleted.filter(t => {
            const plannedDate = parseDate(t.planned);
            const actualDate = parseDate(t.actual!);
            if (!plannedDate || !actualDate) return false;
            const delay = calculateWorkingDaysDelay(plannedDate, actualDate, holidays, { isSaturdayWorkday: isLastWeekReport });
            return delay > 0;
        });

        const planVsActual_Planned = employeeTasksForPeriod.length;
        const planVsActual_NotDone_Percent = planVsActual_Planned > 0 ? Math.round((notDoneTasks.length / planVsActual_Planned) * 100) : 0;
        const onTime_Planned = tasksCompleted.length;
        const onTime_NotDone_Percent = onTime_Planned > 0 ? Math.round((lateTasks.length / onTime_Planned) * 100) : 0;

        return {
            employeeDetails: { name: selectedMisEmployeeName, email: email || 'No email found', photoUrl },
            attendance: {
                workingDays: attendanceBreakdown.workingDays,
                daysPresent: attendanceBreakdown.daysPresent,
                attendancePercentage: attendanceBreakdown.attendancePercentage,
                notMarked: attendanceBreakdown.notMarked,
                otherStatusesBreakdown: attendanceBreakdown.otherStatusesBreakdown,
                dateRange: dateRangeStr
            },
            performance: {
                planVsActual: { planned: planVsActual_Planned, actual: planVsActual_Planned - notDoneTasks.length, percent: planVsActual_NotDone_Percent },
                onTime: { planned: onTime_Planned, actual: onTime_Planned - lateTasks.length, percent: onTime_NotDone_Percent }
            },
            notDoneTasks, lateTasks, dateRange: dateRangeStr
        };
    }, [selectedMisEmployeeName, misWeekdayTasks, people, dailyAttendanceData, selectedMisPeriod, holidays, PRESENT_STATUSES]);
    
    const allEmployees = useMemo(() => {
        // Populate the employee filter from the names present in the "On Track" and "Negative Score" lists.
        const combined = [...negativeScoreEmployees, ...onTrackEmployees];
        const names = new Set(combined.map(person => person.name.trim()).filter(Boolean));
        return Array.from(names).sort();
    }, [negativeScoreEmployees, onTrackEmployees]);

    const { filteredPendingTasks, tableTitle } = useMemo(() => {
        let tasksToFilter: DashboardTask[];
        let title: string;

        switch (filterMode) {
            case 'overdue': tasksToFilter = overdueTasks; title = 'Overdue Tasks'; break;
            case 'today': tasksToFilter = dueTodayTasks; title = 'Tasks Due Today'; break;
            default: tasksToFilter = pendingTasks; title = 'My Pending Tasks'; break;
        }

        if (!searchTerm) return { filteredPendingTasks: tasksToFilter, tableTitle: title };

        const filtered = tasksToFilter.filter(task =>
            task.task.toLowerCase().includes(searchTerm.toLowerCase()) ||
            task.taskId.toLowerCase().includes(searchTerm.toLowerCase()) ||
            task.systemType.toLowerCase().includes(searchTerm.toLowerCase())
        );
        return { filteredPendingTasks: filtered, tableTitle: title };
    }, [pendingTasks, overdueTasks, dueTodayTasks, filterMode, searchTerm]);

    const selectableTasks = useMemo(() => {
        return filteredPendingTasks.filter(task => {
            const requiresAttachment = task.attachmentUrl && task.attachmentUrl.trim() !== '';
            const isActionable = ALLOWED_SYSTEM_TYPES_FOR_SUBMIT.includes(task.systemType);
            return !requiresAttachment && isActionable;
        });
    }, [filteredPendingTasks]);

    useEffect(() => {
        setSelectedTaskIds(new Set());
    }, [filterMode, searchTerm]);
    
    const submitTaskAsDone = async (task: DashboardTask, file: File | null) => {
        if (!task.taskId) {
            alert("Cannot mark task as done: Missing Task ID.");
            return;
        }
    
        // Optimistically update UI to prevent multiple clicks
        setInFlightTaskIds(prev => new Set(prev).add(task.id));
    
        if (file) {
            setIsSubmitting(true);
        }
    
        try {
            const postData: {
                action: string; sheetName: string; newData: Record<string, any>;
                historyRecord: Record<string, any>; attachment?: { fileName: string; mimeType: string; content: string };
            } = {
                action: 'create', sheetName: 'Done Task Status',
                newData: {
                    'Task ID': task.taskId, 'System Type': task.systemType, 'TASK': task.task,
                    'Planned': task.planned.split(' ')[0], 'Timestamp': formatDateToDDMMYYYY(new Date()),
                    'DOER NAME': task.userName || task.name, 
                    'Marked Done By': authenticatedUser?.mailId,
                    'Login ID': authenticatedUser?.mailId,
                },
                historyRecord: {
                    systemType: 'Dashboard', task: `Task ID: ${task.taskId}`,
                    changedBy: authenticatedUser?.mailId,
                    change: `Marked Done on ${new Date().toLocaleString()}${file ? ` with attachment ${file.name}` : ''}`
                }
            };
    
            if (file) {
                const base64String = await fileToBase64(file);
                postData.attachment = { fileName: file.name, mimeType: file.type, content: base64String };
            }
    
            await postToGoogleSheet(postData);
            if (file) setAttachmentModalTask(null);
        } catch (error) {
            console.error("Failed to mark task as done:", error);
            if (error instanceof Error) alert(`Error marking task as done: ${error.message}`);
            else alert(`An unknown error occurred while marking task as done.`);
            
            // On error, revert the optimistic UI update
            setInFlightTaskIds(prev => {
                const newSet = new Set(prev);
                newSet.delete(task.id);
                return newSet;
            });
        } finally {
            if (file) {
                setIsSubmitting(false);
            }
        }
    };

    const submitReDoneTaskWithAttachment = async (task: DashboardTask, file: File) => {
        if (!task.taskId) {
            alert("Cannot mark task as re-done: Missing Task ID.");
            return;
        }
        setIsSubmitting(true);
        try {
            const base64String = await fileToBase64(file);
            const postData = {
                action: 'create', sheetName: 'Done Task Status',
                newData: {
                    'Task ID': task.taskId, 'System Type': task.systemType, 'TASK': task.task,
                    'Planned': task.planned.split(' ')[0], 'Timestamp': formatDateToDDMMYYYY(new Date()),
                    'DOER NAME': task.userName || task.name,
                    'Marked Done By': authenticatedUser?.mailId,
                    'Login ID': authenticatedUser?.mailId,
                },
                historyRecord: {
                    systemType: 'Dashboard', task: `Task ID: ${task.taskId}`,
                    changedBy: authenticatedUser?.mailId,
                    change: `Marked Re-Done on ${new Date().toLocaleString()} with attachment ${file.name}`
                },
                attachment: { fileName: file.name, mimeType: file.type, content: base64String }
            };
            await postToGoogleSheet(postData);
            setInFlightTaskIds(prev => new Set(prev).add(task.id));
            setAttachmentModalTask(null);
        } catch (error) {
            console.error("Failed to mark task as re-done with attachment:", error);
            if (error instanceof Error) alert(`Error marking task as re-done: ${error.message}`);
            else alert(`An unknown error occurred while marking task as re-done.`);
        } finally {
            setIsSubmitting(false);
        }
    };

    const handleMarkDoneClick = (task: DashboardTask) => {
        const requiresAttachment = task.attachmentUrl && task.attachmentUrl.trim() !== '';
        if (requiresAttachment) setAttachmentModalTask(task);
        else submitTaskAsDone(task, null);
    };
    
    const handleModalSubmit = (file: File) => {
        if (attachmentModalTask) {
            if (attachmentModalTask.actual && attachmentModalTask.actual.trim() !== '') {
                submitReDoneTaskWithAttachment(attachmentModalTask, file);
            } else {
                submitTaskAsDone(attachmentModalTask, file);
            }
        }
    };

    const handleToggleSelectOne = (taskId: string) => {
        setSelectedTaskIds(prev => {
            const newSelection = new Set(prev);
            if (newSelection.has(taskId)) newSelection.delete(taskId);
            else newSelection.add(taskId);
            return newSelection;
        });
    };

    const handleToggleSelectAll = () => {
        if (selectedTaskIds.size === selectableTasks.length) setSelectedTaskIds(new Set());
        else setSelectedTaskIds(new Set(selectableTasks.map(t => t.id)));
    };

    const handleMarkMultipleDone = async () => {
        if (selectedTaskIds.size === 0) return;
        setIsSubmitting(true);
        const tasksToSubmit = dashboardTasks.filter(task => selectedTaskIds.has(task.id));
        
        const newDatas = tasksToSubmit.map(task => ({
            'Task ID': task.taskId, 'System Type': task.systemType, 'TASK': task.task,
            'Planned': task.planned.split(' ')[0], 'Timestamp': formatDateToDDMMYYYY(new Date()),
            'DOER NAME': task.userName || task.name, 
            'Marked Done By': authenticatedUser?.mailId,
            'Login ID': authenticatedUser?.mailId,
        }));

        const historyRecords = tasksToSubmit.map(task => ({
            systemType: 'Dashboard', task: `Task ID: ${task.taskId}`,
            changedBy: authenticatedUser?.mailId,
            change: `Marked Done on ${new Date().toLocaleString()}`
        }));

        try {
            await postToGoogleSheet({
                action: 'batchCreate',
                sheetName: 'Done Task Status',
                newDatas: newDatas,
                historyRecords: historyRecords
            });
            const submittedIds = tasksToSubmit.map(t => t.id);
            setInFlightTaskIds(prev => new Set([...prev, ...submittedIds]));
            setSelectedTaskIds(new Set());
        } catch(error) {
            console.error("Failed to submit batch tasks:", error);
            if (error instanceof Error) alert(`Error submitting tasks: ${error.message}`);
            else alert('An unknown error occurred while submitting tasks.');
        } finally {
            setIsSubmitting(false);
        }
    };

    const isAllSelected = selectableTasks.length > 0 && selectedTaskIds.size === selectableTasks.length;

    const periodOptions = useMemo(() => {
        const options = [
            { value: 'lastWeek', label: 'Last Week' },
            { value: 'lastToLastWeek', label: 'Last to Last Week' }
        ];
    
        // Hardcode reference date to match screenshot's context
        const referenceYear = 2025;
        const referenceMonth = 8; // 0-indexed for September
    
        options.push({ value: `year-${referenceYear}`, label: `Full Year ${referenceYear}` });
    
        // Generate 12 months backwards from the reference date
        for (let i = 0; i < 12; i++) {
            const date = new Date(referenceYear, referenceMonth - i, 1);
            const year = date.getFullYear();

            // As per user request, remove any filters for years before 2025.
            if (year < 2025) {
                break;
            }
            
            const month = date.getMonth(); // 0-indexed
            const monthName = date.toLocaleString('default', { month: 'long' });
            
            options.push({ value: `month-${year}-${month}`, label: `${monthName} ${year}` });
        }
        
        return options;
    }, []);

    // --- Calendar View Logic ---
    const tasksByDate = useMemo(() => {
        const map = new Map<string, DashboardTask[]>();
        userTasks.forEach(task => {
            const plannedDate = parseDate(task.planned);
            if (plannedDate) {
                const dateKey = `${plannedDate.getFullYear()}-${String(plannedDate.getMonth() + 1).padStart(2, '0')}-${String(plannedDate.getDate()).padStart(2, '0')}`;
                if (!map.has(dateKey)) {
                    map.set(dateKey, []);
                }
                map.get(dateKey)!.push(task);
            }
        });
        return map;
    }, [userTasks]);
    
    const userAttendanceByDate = useMemo(() => {
        const map = new Map<string, string>(); // Map<dateKey, status>
        const userEmailLower = userEmail.toLowerCase();
        const userNameLower = userName.toLowerCase();

        dailyAttendanceData.forEach(att => {
            if (!att.date) return;

            const attEmailLower = (att.email || '').toLowerCase();
            let isMatch = false;

            if (userEmailLower && attEmailLower) {
                isMatch = attEmailLower === userEmailLower;
            } else {
                isMatch = att.name.toLowerCase() === userNameLower;
            }
            
            if (isMatch) {
                const date = parseDate(att.date);
                if (date) {
                    const dateKey = `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`;
                    map.set(dateKey, att.status);
                }
            }
        });
        return map;
    }, [dailyAttendanceData, userName, userEmail]);

    return (
        <div className="task-dashboard-layout">
            <header className="dashboard-header">
                <div className="dashboard-header-left">
                    <div className="dashboard-title-section">
                        <div className="dashboard-title-icon"><DashboardIcon/></div>
                        <h2>Task Dashboard</h2>
                    </div>
                    <div className="dashboard-tabs">
                        <button onClick={() => setDashboardMode('myDashboard')} className={dashboardMode === 'myDashboard' ? 'active' : ''}>My Dashboard</button>
                        {isAdmin && <button onClick={() => setDashboardMode('employeeMIS')} className={dashboardMode === 'employeeMIS' ? 'active' : ''}>Employee MIS</button>}
                    </div>
                </div>
                <div className="dashboard-header-right">
                    <div className="dashboard-user-info">
                        <div className="user-name">{userName}</div>
                        <div className="update-time">Updated: {lastUpdatedTime}</div>
                    </div>
                    <div className="dashboard-avatar">
                        {photoUrl ? <img src={photoUrl} alt={`${userName}'s profile`} referrerPolicy="no-referrer" /> : <UserIcon />}
                    </div>
                </div>
            </header>

            {dashboardMode === 'employeeMIS' && isAdmin ? (
                <div className="employee-mis-view">
                    <h3 className="mis-view-title">Last Week's Team Highlights</h3>
                    {misTasksError && <div className="error-message" style={{marginBottom: '24px'}}>{misTasksError}</div>}
                    <div className="highlights-grid">
                        <div className="dashboard-card highlight-card">
                            <div className="card-title-section">
                                <div className="icon-wrapper negative"><NegativeIcon /></div>
                                <h4 className="card-title negative">Negative score</h4>
                                <span className="count-badge negative">{negativeScoreEmployees.length}</span>
                            </div>
                            <div className="employee-tags">
                                {negativeScoreEmployees.map(person => <button key={person.name} className="employee-tag" onClick={() => setSelectedMisEmployeeName(person.name)}>{person.name}</button>)}
                            </div>
                        </div>
                        <div className="dashboard-card highlight-card">
                            <div className="card-title-section">
                                <div className="icon-wrapper on-track"><OnTrackIcon /></div>
                                <h4 className="card-title on-track">On Track</h4>
                                <span className="count-badge on-track">{onTrackEmployees.length}</span>
                            </div>
                            <div className="employee-tags">
                                {onTrackEmployees.map(person => <button key={person.name} className="employee-tag" onClick={() => setSelectedMisEmployeeName(person.name)}>{person.name}</button>)}
                            </div>
                        </div>
                    </div>

                    {selectedMisEmployeeName && misReportData && (
                        <div className="mis-report-view" ref={reportRef} style={{ paddingTop: '32px' }}>
                            <h3 className="mis-view-title">Detailed Report</h3>
                            <div className="mis-filters">
                                <div className="filter-group">
                                    <label htmlFor="select-employee">Select Employee</label>
                                    <select id="select-employee" value={selectedMisEmployeeName} onChange={e => setSelectedMisEmployeeName(e.target.value)}>
                                        {allEmployees.map(name => <option key={name} value={name}>{name}</option>)}
                                    </select>
                                </div>
                                <div className="filter-group">
                                    <label htmlFor="select-period">Select Period</label>
                                    <select id="select-period" value={selectedMisPeriod} onChange={e => setSelectedMisPeriod(e.target.value)}>
                                        {periodOptions.map(option => (
                                            <option key={option.value} value={option.value}>
                                                {option.label}
                                            </option>
                                        ))}
                                    </select>
                                </div>
                            </div>

                            <div className="dashboard-main">
                                <aside className="dashboard-sidebar">
                                    <div className="dashboard-card user-profile-card">
                                        <div className="dashboard-avatar">
                                            {misReportData.employeeDetails.photoUrl ? <img src={misReportData.employeeDetails.photoUrl} alt={`${misReportData.employeeDetails.name}'s profile`} referrerPolicy="no-referrer" /> : <UserIcon />}
                                        </div>
                                        <h3 className="user-name">{misReportData.employeeDetails.name}</h3>
                                        <p className="user-email">{misReportData.employeeDetails.email}</p>
                                    </div>
                                    <div className="dashboard-card attendance-card">
                                        <h3>Attendance Summary <span className="date-range">{misReportData.attendance.dateRange}</span></h3>
                                        <div className="attendance-progress">
                                            <CircularProgress percentage={misReportData.attendance.attendancePercentage} color="#22c55e" />
                                            <div className="attendance-details-list">
                                                <div className="detail-item">
                                                    <span>Working Days</span>
                                                    <span className="detail-value">{misReportData.attendance.workingDays}</span>
                                                </div>
                                                <div className="detail-item">
                                                    <span>Present</span>
                                                    <span className="detail-value">{misReportData.attendance.daysPresent}</span>
                                                </div>
                                                {misReportData.attendance.otherStatusesBreakdown.map(([status, count]) => (
                                                    <div className="detail-item" key={status}>
                                                        <span>{status}</span>
                                                        <span className="detail-value">{count}</span>
                                                    </div>
                                                ))}
                                                {misReportData.attendance.notMarked > 0 && (
                                                    <div className="detail-item">
                                                        <span>Not Marked</span>
                                                        <span className="detail-value">{misReportData.attendance.notMarked}</span>
                                                    </div>
                                                )}
                                            </div>
                                        </div>
                                    </div>
                                </aside>
                                <main className="dashboard-content">
                                    <div className="dashboard-card performance-card">
                                        <h3>Performance Overview <span className="date-range">{misReportData.dateRange}</span></h3>
                                        <table className="performance-table">
                                            <thead><tr><th>KRA</th><th>KPI</th><th>PLANNED</th><th>ACTUAL</th><th>ACTUAL %</th></tr></thead>
                                            <tbody>
                                                <tr>
                                                    <td>All work should be done as per plan</td><td>% work NOT done</td>
                                                    <td className="numeric">{misReportData.performance.planVsActual.planned}</td><td className="numeric">{misReportData.performance.planVsActual.actual}</td>
                                                    <td className="numeric actual-percent" style={{ color: misReportData.performance.planVsActual.percent > 0 ? '#ef4444' : '#22c55e' }}>{misReportData.performance.planVsActual.percent}%</td>
                                                </tr>
                                                <tr>
                                                    <td>All work should be done on time</td><td>% work NOT done on time</td>
                                                    <td className="numeric">{misReportData.performance.onTime.planned}</td><td className="numeric">{misReportData.performance.onTime.actual}</td>
                                                    <td className="numeric actual-percent" style={{ color: misReportData.performance.onTime.percent > 0 ? '#ef4444' : '#22c55e' }}>{misReportData.performance.onTime.percent}%</td>
                                                </tr>
                                            </tbody>
                                        </table>
                                    </div>
                                    <div className="dashboard-card mis-task-list-card">
                                        <div className="mis-task-list-header">
                                            <span>Work NOT Done for Selected Period</span>
                                            <span className="task-count-badge">{misReportData.notDoneTasks.length}</span>
                                        </div>
                                        <div className="mis-task-list-body">
                                            {misReportData.notDoneTasks.length > 0 ? (
                                                <table className="mis-task-table">
                                                    <thead>
                                                        <tr>
                                                            <th>TASK ID</th>
                                                            <th>SYSTEM TYPE</th>
                                                            <th>STEP CODE</th>
                                                            <th>TASK</th>
                                                            <th>PLANNED</th>
                                                        </tr>
                                                    </thead>
                                                    <tbody>
                                                        {misReportData.notDoneTasks.map(task => (
                                                            <tr key={task.id} className="clickable-row" onClick={() => setHistoryModalTask(task)}>
                                                                <td>{task.taskId}</td>
                                                                <td>{task.systemType}</td>
                                                                <td>{task.stepCode}</td>
                                                                <td>{task.task}</td>
                                                                <td>{task.planned.split(' ')[0]}</td>
                                                            </tr>
                                                        ))}
                                                    </tbody>
                                                </table>
                                            ) : (
                                                <div className="no-tasks-message">
                                                    No tasks to display for this category in the selected period.
                                                </div>
                                            )}
                                        </div>
                                    </div>
                                    <div className="dashboard-card mis-task-list-card">
                                        <div className="mis-task-list-header">
                                            <span>Work NOT Done On Time for Selected Period</span>
                                            <span className="task-count-badge">{misReportData.lateTasks.length}</span>
                                        </div>
                                        <div className="mis-task-list-body">
                                            {misReportData.lateTasks.length > 0 ? (
                                                <table className="mis-task-table">
                                                    <thead>
                                                        <tr>
                                                            <th>TASK ID</th>
                                                            <th>SYSTEM TYPE</th>
                                                            <th>STEP CODE</th>
                                                            <th>TASK</th>
                                                            <th>PLANNED</th>
                                                            <th>ACTUAL</th>
                                                            <th className="delay-days">DELAY (DAYS)</th>
                                                        </tr>
                                                    </thead>
                                                    <tbody>
                                                        {misReportData.lateTasks.map(task => {
                                                            const plannedDate = parseDate(task.planned);
                                                            const actualDate = parseDate(task.actual || '');
                                                            const delay = (plannedDate && actualDate)
                                                                ? calculateWorkingDaysDelay(plannedDate, actualDate, holidays, { isSaturdayWorkday: selectedMisPeriod === 'lastWeek' })
                                                                : 0;

                                                            return (
                                                                <tr key={task.id} className="clickable-row" onClick={() => setHistoryModalTask(task)}>
                                                                    <td>{task.taskId}</td>
                                                                    <td>{task.systemType}</td>
                                                                    <td>{task.stepCode}</td>
                                                                    <td>{task.task}</td>
                                                                    <td>{task.planned.split(' ')[0]}</td>
                                                                    <td>{task.actual?.split(' ')[0]}</td>
                                                                    <td className="delay-days">{delay > 0 ? delay : ''}</td>
                                                                </tr>
                                                            );
                                                        })}
                                                    </tbody>
                                                </table>
                                            ) : (
                                                <div className="no-tasks-message">
                                                    No tasks to display for this category in the selected period.
                                                </div>
                                            )}
                                        </div>
                                    </div>
                                </main>
                            </div>
                        </div>
                    )}
                </div>
            ) : (
                <>
                    <div className="dashboard-main">
                        <aside className="dashboard-sidebar">
                            <div className="dashboard-card user-profile-card">
                                <div className="dashboard-avatar">
                                    {photoUrl ? <img src={photoUrl} alt={`${userName}'s profile`} referrerPolicy="no-referrer" /> : <UserIcon />}
                                </div>
                                <h3 className="user-name">{userName}</h3>
                                <p className="user-email">{userEmail}</p>
                            </div>
                            <div className="dashboard-card attendance-card">
                                <h3>My Weekly Attendance <span className="date-range">{prevWeekDateRange}</span></h3>
                                <div className="attendance-progress">
                                    <CircularProgress percentage={myAttendanceBreakdown.attendancePercentage} color="#22c55e" />
                                     <div className="attendance-details-list">
                                        <div className="detail-item">
                                            <span>Working Days</span>
                                            <span className="detail-value">{myAttendanceBreakdown.workingDays}</span>
                                        </div>
                                        <div className="detail-item">
                                            <span>Present</span>
                                            <span className="detail-value">{myAttendanceBreakdown.daysPresent}</span>
                                        </div>
                                        {myAttendanceBreakdown.otherStatusesBreakdown.map(([status, count]) => (
                                            <div className="detail-item" key={status}>
                                                <span>{status}</span>
                                                <span className="detail-value">{count}</span>
                                            </div>
                                        ))}
                                        {myAttendanceBreakdown.notMarked > 0 && (
                                            <div className="detail-item">
                                                <span>Not Marked</span>
                                                <span className="detail-value">{myAttendanceBreakdown.notMarked}</span>
                                            </div>
                                        )}
                                    </div>
                                </div>
                            </div>
                        </aside>

                        <main className="dashboard-content">
                             <div className="dashboard-view-switcher">
                                <button onClick={() => setCurrentView('stats')} className={currentView === 'stats' ? 'active' : ''}>Stats View</button>
                                <button onClick={() => setCurrentView('calendar')} className={currentView === 'calendar' ? 'active' : ''}>Calendar View</button>
                            </div>
                            {currentView === 'stats' ? (
                                <>
                                    <div className="dashboard-card performance-card">
                                        <h3>My Weekly Performance <span className="date-range">{prevWeekDateRange}</span></h3>
                                        <table className="performance-table">
                                            <thead><tr><th>KRA</th><th>KPI</th><th>PLANNED</th><th>ACTUAL</th><th>ACTUAL %</th><th></th></tr></thead>
                                            <tbody>
                                                <tr onClick={() => setExpandedKpi(prev => prev === 'notDone' ? null : 'notDone')} className={expandedKpi === 'notDone' ? 'active-kpi' : ''} style={{cursor: 'pointer'}}>
                                                    <td>All work should be done as per plan</td><td>% work NOT done</td>
                                                    <td className="numeric">{planVsActual_Planned}</td><td className="numeric">{planVsActual_Actual}</td>
                                                    <td className="numeric actual-percent" style={{ color: planVsActual_Percent > 10 ? '#ef4444' : '#22c55e' }}>{planVsActual_Percent}%</td>
                                                    <td><ArrowIcon className={expandedKpi === 'notDone' ? 'expanded' : ''} /></td>
                                                </tr>
                                                <tr onClick={() => setExpandedKpi(prev => prev === 'notOnTime' ? null : 'notOnTime')} className={expandedKpi === 'notOnTime' ? 'active-kpi' : ''} style={{cursor: 'pointer'}}>
                                                    <td>All work should be done on time</td><td>% work NOT done on time</td>
                                                    <td className="numeric">{onTime_Planned}</td><td className="numeric">{onTime_Actual}</td>
                                                    <td className="numeric actual-percent" style={{ color: onTime_Percent > 10 ? '#ef4444' : '#22c55e' }}>{onTime_Percent}%</td>
                                                    <td><ArrowIcon className={expandedKpi === 'notOnTime' ? 'expanded' : ''} /></td>
                                                </tr>
                                            </tbody>
                                        </table>
                                    </div>
                                    {expandedKpi && (
                                        <div className="dashboard-card pending-tasks-card">
                                            <div className="card-header">
                                                <h3>
                                                    {expandedKpi === 'notDone' 
                                                        ? `Work NOT Done (${notDoneTasksForPrevWeek.length} Tasks)`
                                                        : `Work NOT Done on Time (${notOnTimeTasksForPrevWeek.length} Tasks)`
                                                    }
                                                </h3>
                                            </div>
                                            <div className="table-container" style={{maxHeight: '400px'}}>
                                                <table className="pending-tasks-table kpi-details-table">
                                                    <thead>
                                                        <tr>
                                                            <th>TASK ID</th>
                                                            <th>SYSTEM TYPE</th>
                                                            <th>STEP CODE</th>
                                                            <th>TASK</th>
                                                            <th>PLANNED</th>
                                                            <th>ACTUAL</th>
                                                        </tr>
                                                    </thead>
                                                    <tbody>
                                                        {(expandedKpi === 'notDone' ? notDoneTasksForPrevWeek : notOnTimeTasksForPrevWeek).map(task => (
                                                            <tr key={task.id}>
                                                                <td>{task.taskId}</td>
                                                                <td>{task.systemType}</td>
                                                                <td>{task.stepCode}</td>
                                                                <td>{task.task}</td>
                                                                <td>{task.planned.split(' ')[0]}</td>
                                                                <td>{task.actual ? task.actual.split(' ')[0] : '-'}</td>
                                                            </tr>
                                                        ))}
                                                        {(expandedKpi === 'notDone' ? notDoneTasksForPrevWeek.length === 0 : notOnTimeTasksForPrevWeek.length === 0) && (
                                                            <tr><td colSpan={6} style={{textAlign: 'center', padding: '32px'}}>No tasks to display for this category.</td></tr>
                                                        )}
                                                    </tbody>
                                                </table>
                                            </div>
                                        </div>
                                    )}
                                    <div className="stats-grid">
                                        <StatCard title="My Pending Tasks" value={pendingTasks.length} icon={<PendingIcon />} className={`stat-card--pending stat-card-clickable ${filterMode === 'all' ? 'active' : ''}`} onClick={() => setFilterMode('all')} ariaPressed={filterMode === 'all'} />
                                        <StatCard title="Overdue Tasks" value={overdueTasks.length} icon={<OverdueIcon />} className={`stat-card--overdue stat-card-clickable ${filterMode === 'overdue' ? 'active' : ''}`} onClick={() => setFilterMode('overdue')} ariaPressed={filterMode === 'overdue'} />
                                        <StatCard title="Tasks Due Today" value={dueTodayTasks.length} icon={<TodayIcon />} className={`stat-card--today stat-card-clickable ${filterMode === 'today' ? 'active' : ''}`} onClick={() => setFilterMode('today')} ariaPressed={filterMode === 'today'} />
                                    </div>
                                    <div className="dashboard-card search-card">
                                        <h3>Search Tasks</h3>
                                        <div className="search-input-wrapper">
                                            <SearchIcon /><input type="text" placeholder="Search by ID, task, system, doer..." value={searchTerm} onChange={(e) => setSearchTerm(e.target.value)} />
                                        </div>
                                    </div>
                                    <div className="dashboard-card pending-tasks-card">
                                        <div className="card-header">
                                            <h3>{tableTitle} ({filteredPendingTasks.length})</h3>
                                            {selectedTaskIds.size > 0 && (<button className="btn btn-primary btn-submit-selected" onClick={handleMarkMultipleDone} disabled={isSubmitting}>{isSubmitting ? 'Submitting...' : `Submit Selected (${selectedTaskIds.size})`}</button>)}
                                        </div>
                                        <div className="table-container" style={{maxHeight: '400px'}}>
                                            <table className="pending-tasks-table">
                                                <thead><tr><th className="checkbox-cell"><input type="checkbox" onChange={handleToggleSelectAll} checked={isAllSelected} disabled={selectableTasks.length === 0 || isSubmitting} aria-label="Select all tasks" /></th><th>Task ID</th><th>System Type</th><th>TASK</th><th>Planned</th><th>DOER NAME</th><th></th></tr></thead>
                                                <tbody>
                                                    {filteredPendingTasks.length > 0 ? filteredPendingTasks.map(task => {
                                                        const requiresAttachment = task.attachmentUrl && task.attachmentUrl.trim() !== '';
                                                        const isActionable = ALLOWED_SYSTEM_TYPES_FOR_SUBMIT.includes(task.systemType);
                                                        const isDisabledForSelection = isSubmitting || requiresAttachment;
                                                        const isQueued = inFlightTaskIds.has(task.id);

                                                        return (
                                                            <tr key={task.id}>
                                                                <td className="checkbox-cell">
                                                                    {isActionable && (<span style={{ cursor: requiresAttachment ? 'not-allowed' : 'default' }} title={requiresAttachment ? 'Requires document upload; submit individually.' : ''}><input type="checkbox" onChange={() => handleToggleSelectOne(task.id)} checked={selectedTaskIds.has(task.id)} disabled={isDisabledForSelection || isQueued} aria-label={`Select task ${task.taskId}`} /></span>)}
                                                                </td>
                                                                <td>{task.taskId}</td><td>{task.systemType}</td><td>{task.task}</td><td>{task.planned.split(' ')[0]}</td><td>{task.userName || task.name}</td>
                                                                <td>{isActionable && (<button className={`btn-mark-done ${isQueued ? 'btn-queued' : ''}`} onClick={() => handleMarkDoneClick(task)} disabled={isSubmitting || isQueued}>{isQueued ? 'Submitting...' : 'Mark Done'}</button>)}</td>
                                                            </tr>
                                                        );
                                                    }) : (<tr><td colSpan={7} style={{textAlign: 'center', padding: '32px'}}>No tasks found.</td></tr>)}
                                                </tbody>
                                            </table>
                                        </div>
                                    </div>
                                </>
                            ) : (
                                <CalendarView
                                    isAdmin={isAdmin}
                                    calendarDate={calendarDate}
                                    setCalendarDate={setCalendarDate}
                                    calendarMode={calendarMode}
                                    setCalendarMode={setCalendarMode}
                                    tasksByDate={tasksByDate}
                                    setSelectedCalendarDate={setSelectedCalendarDate}
                                    userAttendanceByDate={userAttendanceByDate}
                                    holidays={holidays}
                                />
                            )}
                        </main>
                    </div>
                    <AttachmentModal task={attachmentModalTask} onClose={() => setAttachmentModalTask(null)} onSubmit={handleModalSubmit} isSubmitting={isSubmitting} />
                    <SelectedDateModal
                        date={selectedCalendarDate}
                        tasks={selectedCalendarDate ? tasksByDate.get(`${selectedCalendarDate.getFullYear()}-${String(selectedCalendarDate.getMonth() + 1).padStart(2, '0')}-${String(selectedCalendarDate.getDate()).padStart(2, '0')}`) || [] : []}
                        onClose={() => setSelectedCalendarDate(null)}
                    />
                    {historyModalTask && (
                        <TaskHistoryModal
                            task={historyModalTask}
                            history={taskHistory}
                            onClose={() => setHistoryModalTask(null)}
                        />
                    )}
                </>
            )}
        </div>
    );
};
