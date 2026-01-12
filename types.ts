

export interface Task {
  id: string;
  description: string;
  assignee: string;
  buddy?: string;
  secondBuddy?: string;
  completed: boolean;
  createdAt: number;
  sourceChecklistId?: string;
}

export interface Checklist {
  id:string;
  task: string;
  doer: string;
  frequency: string;
  date: string;
  buddy: string;
  secondBuddy?: string;
}

export interface MasterTask {
  id: string;
  taskId: string;
  plannedDate: string;
  actualDate: string;
  taskDescription: string;
  doer: string;
  originalDoer: string;
  frequency: string;
  pc: string;
  status?: string;
}

export interface DelegationTask {
  id: string;
  timestamp: string;
  assignee: string;
  task: string;
  plannedDate: string;
  actualDate?: string;
  assignerEmail: string;
  assigner: string;
  delegateEmail?: string;
  taskId?: string;
}

export interface DashboardTask {
  id: string;
  timestamp: string;
  taskId: string;
  task: string;
  stepCode: string;
  planned: string;
  actual: string;
  name: string;
  link: string;
  forPc: string;
  systemType: string;
  photoUrl?: string;
  userEmail?: string;
  userName?: string;
  attachmentUrl?: string;
}

// Fix: Define UserRole type exported for use in index.tsx
export type UserRole = 'Admin' | 'Manager' | 'User';

// Fix: Add optional teamEmails property to support Manager accounts
export interface AuthenticatedUser {
    mailId: string;
    role: string;
    teamEmails?: string[];
}

export interface UserAuth extends AuthenticatedUser {
    password?: string;
}

export interface Person {
    name: string;
    email?: string;
    photoUrl?: string;
    status?: string;
}

export interface AttendanceData {
    email: string;
    daysPresent: number;
}

export interface Holiday {
    name: string;
    date: string;
}

export interface DailyAttendance {
    date: string;
    status: string;
    name: string;
    email?: string;
}

export interface TaskHistory {
  timestamp: string;
  systemType: string;
  task: string;
  changedBy: string;
  change: string;
}

export type AppMode = 'checklist' | 'delegation' | 'dashboard';
export type ChecklistSubMode = 'templates' | 'master';
