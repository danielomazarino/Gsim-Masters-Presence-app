// Declaration for the Google Apps Script client-side RPC API (google.script.run)
declare global {
  interface Window {
    google?: {
      script: {
        run: {
          withSuccessHandler: (callback: (result: unknown) => void) => {
            withFailureHandler: (callback: (error: Error | { message?: string }) => void) => {
              getAttendanceData: () => void;
              [key: string]: unknown;
            };
          };
          withFailureHandler: (callback: (error: Error | { message?: string }) => void) => {
            withSuccessHandler: (callback: (result: unknown) => void) => {
              getAttendanceData: () => void;
              [key: string]: unknown;
            };
          };
          [key: string]: unknown;
        };
      };
    };
  }
}

export interface MemberRow {
  row: number;
  name: string;
  paidFee?: string | boolean;
  totalTrainings?: number | string;
  [key: string]: unknown;
}

export interface TrainingColumnInfo {
  index: number;
  colLetter?: string;
  header: string;
  date?: string;
  location?: string;
}

export interface BackendAttendancePayload {
  spreadsheetName?: string;
  sheetName?: string;
  members?: MemberRow[];
  headers?: string[];
  trainingColumns?: TrainingColumnInfo[];
  attendance?: Record<string, unknown> | Array<unknown>;
  today?: string;
  todayStatus?: unknown;
  todayTrainings?: unknown;
  duplicateTrainings?: unknown;
  date?: string;
  rows?: unknown[][];
  [key: string]: unknown;
}

export type LocationOption = 'Askim' | 'Valhalla';

export interface DateOption {
  id: 'today' | 'yesterday' | 'two_days_ago';
  label: string; // e.g. "Idag"
  dateFormatted: string; // e.g. "4/9"
  subText: string; // e.g. "Fredag 4 sep"
}
