

import React, { useState } from 'react';
import { Task, Holiday } from './types';

// --- HOOK FOR LOCALSTORAGE (Exported) ---
export function useLocalStorage<T>(key: string, initialValue: T): [T, React.Dispatch<React.SetStateAction<T>>] {
    const [storedValue, setStoredValue] = useState<T>(() => {
        try {
            const item = window.localStorage.getItem(key);
            if (!item) return initialValue;

            let initial = JSON.parse(item);

            // Cleanup logic on initial load
            if (key === 'task-delegator-tasks' && Array.isArray(initial)) {
                const fourteenDaysAgo = Date.now() - 14 * 24 * 60 * 60 * 1000;
                const cleanedTasks = initial.filter((task: Task) => {
                    return !task.completed || task.createdAt > fourteenDaysAgo;
                });
                return cleanedTasks as T;
            }

            return initial;
        } catch (error) {
            console.error(`Error reading localStorage key “${key}”:`, error);
            return initialValue;
        }
    });

    const setValue: React.Dispatch<React.SetStateAction<T>> = (value) => {
        try {
            const valueToStore = value instanceof Function ? value(storedValue) : value;
            
            let finalValueToStore = valueToStore;

            // Proactive cleanup logic before every write to prevent quota errors
            if (key === 'task-delegator-tasks' && Array.isArray(valueToStore)) {
                const fourteenDaysAgo = Date.now() - 14 * 24 * 60 * 60 * 1000;
                const cleanedTasks = valueToStore.filter((task: Task) => {
                    // Keep task if it's NOT completed, OR if it was completed within the last 14 days.
                    return !task.completed || task.createdAt > fourteenDaysAgo;
                });
                finalValueToStore = cleanedTasks as T;
            }

            setStoredValue(finalValueToStore);
             if (value === null) {
                window.localStorage.removeItem(key);
            } else {
                window.localStorage.setItem(key, JSON.stringify(finalValueToStore));
            }


        } catch (error) {
            console.error(`Error setting localStorage key “${key}”:`, error);
             if (error instanceof DOMException && (error.name === 'QuotaExceededError' || error.name === 'NS_ERROR_DOM_QUOTA_REACHED')) {
                alert("Could not save tasks because the browser's local storage is full. Old completed tasks have been cleared to make space. Please try your action again.");
            }
        }
    };

    return [storedValue, setValue];
}

export const getStartOf = (date: Date, unit: 'day' | 'week' | 'month' | 'quarter' | 'year'): Date => {
    const d = new Date(date);
    if (unit === 'day') d.setHours(0, 0, 0, 0);
    if (unit === 'week') {
        const day = d.getDay();
        const diff = d.getDate() - day + (day === 0 ? -6 : 1); // adjust when day is sunday
        d.setDate(diff);
        d.setHours(0, 0, 0, 0);
    }
    if (unit === 'month') {
        d.setDate(1);
        d.setHours(0, 0, 0, 0);
    }
    if (unit === 'quarter') {
        const quarter = Math.floor(d.getMonth() / 3);
        d.setMonth(quarter * 3, 1);
        d.setHours(0, 0, 0, 0);
    }
    if (unit === 'year') {
        d.setFullYear(d.getFullYear(), 0, 1);
        d.setHours(0, 0, 0, 0);
    }
    return d;
};

export const getIsoDate = (dateString?: string): string => {
    if (!dateString) return '';
    try {
        const datePart = dateString.trim().split(' ')[0];

        // If the format is already YYYY-MM-DD, just return it after validation.
        // This is what the date input provides, so it's the most common case.
        if (/^\d{4}-\d{2}-\d{2}$/.test(datePart)) {
            const [year, month, day] = datePart.split('-').map(Number);
            const date = new Date(Date.UTC(year, month - 1, day));
            if (date.getUTCFullYear() === year && date.getUTCMonth() === month - 1 && date.getUTCDate() === day) {
                return datePart;
            }
        }

        let year: number | undefined, month: number | undefined, day: number | undefined;

        // Try to parse common formats
        // YYYY/MM/DD or YYYY-MM-DD (already handled but good to keep)
        let match = datePart.match(/^(\d{4})[/\-.](\d{1,2})[/\-.](\d{1,2})$/);
        if (match) {
            year = parseInt(match[1], 10);
            month = parseInt(match[2], 10);
            day = parseInt(match[3], 10);
        } else {
            // DD/MM/YYYY
            match = datePart.match(/^(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{4})$/);
            if (match) {
                day = parseInt(match[1], 10);
                month = parseInt(match[2], 10);
                year = parseInt(match[3], 10);
            } else {
                // DD/MM/YY
                match = datePart.match(/^(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{2})$/);
                if (match) {
                    day = parseInt(match[1], 10);
                    month = parseInt(match[2], 10);
                    let tempYear = parseInt(match[3], 10);
                    year = tempYear > 50 ? 1900 + tempYear : 2000 + tempYear;
                }
            }
        }

        if (year !== undefined && month !== undefined && day !== undefined) {
             const date = new Date(Date.UTC(year, month - 1, day));
             // Validate that the created date is what we expect (handles invalid dates like month 13)
             if (date.getUTCFullYear() === year && date.getUTCMonth() === month - 1 && date.getUTCDate() === day) {
                return date.toISOString().split('T')[0];
             }
        }
        
        // Fallback for formats that `new Date()` can parse (e.g. "Jul 12 2024", or locale-dependent "MM/DD/YYYY")
        const fallbackDate = new Date(datePart);
        if (!isNaN(fallbackDate.getTime())) {
            // To prevent timezone shift, we extract year/month/day from the local date
            // and then build a UTC date string from them.
            const date = new Date(Date.UTC(fallbackDate.getFullYear(), fallbackDate.getMonth(), fallbackDate.getDate()));
            return date.toISOString().split('T')[0];
        }

        return ''; // Return empty if parsing fails
    } catch (e) {
        console.error("Could not parse date string:", dateString, e);
        return '';
    }
};

export const parseDate = (dateString?: string): Date | null => {
    if (!dateString) return null;
    
    // Clean the string and handle potential "pm/am"
    const cleanStr = dateString.trim();
    
    // Try ISO format first (YYYY-MM-DD...)
    let d = new Date(cleanStr);
    if (!isNaN(d.getTime()) && cleanStr.includes('-') && cleanStr.indexOf('-') === 4) {
        return d;
    }

    // Handle DD/MM/YYYY HH:mm:ss
    const parts = cleanStr.split(/[\s,]+/);
    const datePart = parts[0];
    const timePart = parts[1] || '';
    const ampm = parts[2] || '';

    let day: number, month: number, year: number;
    let match = datePart.match(/^(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{4})$/);
    if (match) {
        day = parseInt(match[1], 10);
        month = parseInt(match[2], 10);
        year = parseInt(match[3], 10);
    } else {
        // Try YYYY/MM/DD
        match = datePart.match(/^(\d{4})[/\-.](\d{1,2})[/\-.](\d{1,2})$/);
        if (match) {
            year = parseInt(match[1], 10);
            month = parseInt(match[2], 10);
            day = parseInt(match[3], 10);
        } else {
            // Fallback to native Date for other formats
            d = new Date(cleanStr);
            return isNaN(d.getTime()) ? null : d;
        }
    }

    let hours = 0, minutes = 0, seconds = 0;
    if (timePart) {
        const tParts = timePart.split(':');
        hours = parseInt(tParts[0], 10) || 0;
        minutes = parseInt(tParts[1], 10) || 0;
        seconds = parseInt(tParts[2], 10) || 0;

        if (ampm.toLowerCase() === 'pm' && hours < 12) hours += 12;
        if (ampm.toLowerCase() === 'am' && hours === 12) hours = 0;
    }

    const finalDate = new Date(year, month - 1, day, hours, minutes, seconds);
    return isNaN(finalDate.getTime()) ? null : finalDate;
};

export const robustCsvParser = (csvText: string): string[][] => {
    const rows: string[][] = [];
    if (!csvText) return rows;

    const text = csvText.trim().replace(/\r\n/g, '\n');
    let pos = 0;
    let currentRow: string[] = [];
    let currentField = '';
    let inQuotes = false;

    while (pos < text.length) {
        const char = text[pos];

        if (inQuotes) {
            if (char === '"') {
                if (pos + 1 < text.length && text[pos + 1] === '"') { // Escaped double quote
                    currentField += '"';
                    pos++;
                } else {
                    inQuotes = false;
                }
            } else {
                currentField += char;
            }
        } else {
            if (char === '"') {
                inQuotes = true;
            } else if (char === ',') {
                currentRow.push(currentField);
                currentField = '';
            } else if (char === '\n') {
                currentRow.push(currentField);
                rows.push(currentRow);
                currentRow = [];
                currentField = '';
            } else {
                currentField += char;
            }
        }
        pos++;
    }

    // Add the very last field and row
    currentRow.push(currentField);
    rows.push(currentRow);

    // Remove header row and clean up fields
    return rows.slice(1).map(row => row.map(field => field.trim().replace(/^"|"$/g, '')));
};

export const calculateWorkingDaysDelay = (plannedDate: Date, actualDate: Date, holidays: Holiday[], options?: { isSaturdayWorkday?: boolean }): number => {
    // Create copies and normalize to midnight to compare dates only
    const pDate = new Date(plannedDate);
    pDate.setHours(0, 0, 0, 0);
    const aDate = new Date(actualDate);
    aDate.setHours(0, 0, 0, 0);

    // If actual is on or before planned, there's no delay.
    if (aDate.getTime() <= pDate.getTime()) {
        return 0;
    }

    const holidayDateStrings = new Set(
        holidays.map(h => {
            const d = parseDate(h.date);
            return d ? d.toDateString() : '';
        }).filter(Boolean)
    );

    let workingDays = 0;
    const currentDate = new Date(pDate);

    // Loop from the day *after* plannedDate up to and including actualDate
    while (currentDate.getTime() < aDate.getTime()) {
        currentDate.setDate(currentDate.getDate() + 1);
        
        const dayOfWeek = currentDate.getDay(); // Sunday: 0, Saturday: 6
        const isWeekend = options?.isSaturdayWorkday
            ? dayOfWeek === 0 // Only Sunday is a weekend
            : dayOfWeek === 0 || dayOfWeek === 6; // Sunday and Saturday are weekends
        const isHoliday = holidayDateStrings.has(currentDate.toDateString());

        if (!isWeekend && !isHoliday) {
            workingDays++;
        }
    }

    // If the task is late (aDate > pDate) but the calculated delay is 0 
    // (meaning it was completed over a weekend/holiday), count it as a 1-day delay.
    if (workingDays === 0) {
        return 1;
    }

    return workingDays;
};

// Calculates total working days between two dates, INCLUSIVE of both start and end.
export const calculateWorkingDaysPassed = (startDate: Date, endDate: Date, holidays: Holiday[], options?: { isSaturdayWorkday?: boolean }): number => {
    // This function is inclusive of start and end date
    const pDate = new Date(startDate);
    pDate.setHours(0, 0, 0, 0);
    const aDate = new Date(endDate);
    aDate.setHours(0, 0, 0, 0);

    if (aDate.getTime() < pDate.getTime()) return 0;

    const holidayDateStrings = new Set(holidays.map(h => { const d = parseDate(h.date); return d ? d.toDateString() : ''; }).filter(Boolean));
    let workingDays = 0;
    const currentDate = new Date(pDate);

    while (currentDate.getTime() <= aDate.getTime()) {
        const dayOfWeek = currentDate.getDay();
        const isWeekend = options?.isSaturdayWorkday ? dayOfWeek === 0 : dayOfWeek === 0 || dayOfWeek === 6;
        const isHoliday = holidayDateStrings.has(currentDate.toDateString());
        if (!isWeekend && !isHoliday) {
            workingDays++;
        }
        currentDate.setDate(currentDate.getDate() + 1);
    }
    return workingDays;
};
