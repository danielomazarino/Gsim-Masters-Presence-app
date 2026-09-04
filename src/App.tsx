import { useState, useEffect, useMemo, useCallback } from 'react';
import {
  BackendAttendancePayload,
  DateOption,
  LocationOption,
  MemberRow,
  TrainingColumnInfo,
} from './types';
import { calculateDateOptions } from './dateUtils';
import { Check, Settings, ArrowLeft, AlertCircle } from 'lucide-react';

export default function App() {
  const dateOptions = useMemo(() => calculateDateOptions(), []);

  // View state: 'attendance' or 'admin'
  const [currentView, setCurrentView] = useState<'attendance' | 'admin'>('attendance');

  // Selection states
  const [selectedDateOption, setSelectedDateOption] = useState<DateOption>(dateOptions[0]);
  const [selectedLocation, setSelectedLocation] = useState<LocationOption>('Askim');

  // Backend state
  const [backendPayload, setBackendPayload] = useState<BackendAttendancePayload | null>(null);
  const [isLoading, setIsLoading] = useState(false);
  const [errorMessage, setErrorMessage] = useState<string | null>(null);

  // Local attendance state: key is member.name, value is boolean
  // Strictly LOCAL UI STATE ONLY - no saveAttendance() calls
  const [localAttendance, setLocalAttendance] = useState<Record<string, boolean>>({});

  // Request data using google.script.run
  // Strictly NO fetch(), NO XMLHttpRequest, NO REST API, NO OAuth tokens
  const loadAttendanceFromBackend = useCallback(() => {
    if (isLoading) return;

    setIsLoading(true);
    setErrorMessage(null);

    const googleScript = typeof window !== 'undefined' ? window.google?.script?.run : undefined;

    if (!googleScript) {
      setIsLoading(false);
      setErrorMessage(
        'google.script.run är inte tillgängligt i denna miljö. Kör sidan inuti Google Apps Script Web App (HtmlService) för att läsa direkt från Google Sheet.'
      );
      return;
    }

    googleScript
      .withSuccessHandler((response: unknown) => {
        setIsLoading(false);
        const data = response as BackendAttendancePayload | null;
        if (!data) {
          setErrorMessage('Inget data mottogs från servern.');
          return;
        }
        setBackendPayload(data);
      })
      .withFailureHandler((error: Error | { message?: string }) => {
        setIsLoading(false);
        const msg = error && typeof error.message === 'string' ? error.message : String(error);
        setErrorMessage(msg || 'Ett fel uppstod vid anrop till getAttendanceData().');
      })
      .getAttendanceData();
  }, [isLoading]);

  // Load automatically if running inside Apps Script environment
  useEffect(() => {
    if (typeof window !== 'undefined' && window.google?.script?.run) {
      loadAttendanceFromBackend();
    }
  }, [loadAttendanceFromBackend]);

  // Match training column: training.location === selectedLocation && training.date === selectedDateFormatted
  const matchedTraining = useMemo(() => {
    if (!backendPayload?.trainingColumns) return null;
    const cols: TrainingColumnInfo[] = backendPayload.trainingColumns;
    return cols.find(
      (c) => c.location === selectedLocation && c.date === selectedDateOption.dateFormatted
    ) || null;
  }, [backendPayload, selectedLocation, selectedDateOption]);

  // Recompute local attendance when backend data, date, or location changes
  useEffect(() => {
    if (!backendPayload || !backendPayload.members) {
      setLocalAttendance({});
      return;
    }

    const members: MemberRow[] = backendPayload.members || [];
    const newAttendance: Record<string, boolean> = {};

    if (!matchedTraining) {
      members.forEach((m) => {
        newAttendance[m.name] = false;
      });
    } else {
      const attData = backendPayload.attendance as Record<string, Record<string, boolean>> | undefined;
      members.forEach((m) => {
        const isPresent = Boolean(
          attData && attData[m.name] && attData[m.name][matchedTraining.header] === true
        );
        newAttendance[m.name] = isPresent;
      });
    }

    setLocalAttendance(newAttendance);
  }, [backendPayload, matchedTraining]);

  // Toggle single member attendance (LOCAL UI STATE ONLY)
  const toggleAttendance = (memberName: string) => {
    setLocalAttendance((prev) => ({
      ...prev,
      [memberName]: !prev[memberName],
    }));
  };

  const membersList: MemberRow[] = backendPayload?.members || [];

  // Discrete tally of present members
  const presentCount = useMemo(() => {
    return Object.values(localAttendance).filter(Boolean).length;
  }, [localAttendance]);

  return (
    <div className="min-h-screen bg-slate-100 text-slate-900 font-sans pb-10 select-none">
      {/* Sticky top control panel for iPhone quick access */}
      <header className="sticky top-0 z-30 bg-white border-b border-slate-200 shadow-xs">
        {/* Subtle preview notice */}
        <div className="bg-amber-100/90 text-amber-900 text-[11px] font-medium py-1 px-3 text-center border-b border-amber-200">
          Förhandsvisning – inga ändringar sparas ännu
        </div>

        {currentView === 'attendance' ? (
          <div className="max-w-md mx-auto p-2.5 space-y-2">
            {/* Row 1: Compact Date Choice */}
            <div className="flex items-center gap-1.5">
              {dateOptions.map((option) => {
                const isSelected = selectedDateOption.id === option.id;
                return (
                  <button
                    key={option.id}
                    type="button"
                    onClick={() => setSelectedDateOption(option)}
                    disabled={isLoading}
                    className={`flex-1 min-h-[44px] py-1.5 px-2 rounded-lg text-xs font-medium transition-colors cursor-pointer text-center flex flex-col justify-center items-center ${
                      isSelected
                        ? 'bg-blue-600 text-white font-semibold shadow-xs'
                        : 'bg-slate-100 text-slate-700 hover:bg-slate-200 active:bg-slate-300'
                    }`}
                  >
                    <span className="leading-tight">{option.label}</span>
                    <span className={`text-[10px] ${isSelected ? 'text-blue-100' : 'text-slate-400'}`}>
                      {option.dateFormatted}
                    </span>
                  </button>
                );
              })}
            </div>

            {/* Row 2: Compact Location Choice + Discrete Count Badge + Admin Cog */}
            <div className="flex items-center gap-2">
              <div className="flex flex-1 gap-1.5">
                {(['Askim', 'Valhalla'] as LocationOption[]).map((loc) => {
                  const isSelected = selectedLocation === loc;
                  return (
                    <button
                      key={loc}
                      type="button"
                      onClick={() => setSelectedLocation(loc)}
                      disabled={isLoading}
                      className={`flex-1 min-h-[44px] py-2 px-3 rounded-lg text-xs font-semibold transition-colors cursor-pointer text-center ${
                        isSelected
                          ? 'bg-blue-600 text-white shadow-xs'
                          : 'bg-slate-100 text-slate-700 hover:bg-slate-200 active:bg-slate-300'
                      }`}
                    >
                      {loc}
                    </button>
                  );
                })}
              </div>

              {/* Discrete verification checksum */}
              <div
                className="min-h-[44px] px-3 bg-slate-100 border border-slate-200 rounded-lg flex items-center justify-center shrink-0"
                title="Kontrollsumma för närvarande"
              >
                <span className="text-xs font-semibold text-slate-800">
                  Närvarande: <span className="font-bold font-mono text-blue-700 text-sm ml-0.5">{presentCount}</span>
                </span>
              </div>

              {/* Discrete Admin Gear Icon (replaces old refresh button) */}
              <button
                type="button"
                onClick={() => setCurrentView('admin')}
                aria-label="Öppna administration"
                title="Administration"
                className="min-h-[44px] min-w-[44px] rounded-lg bg-slate-100 hover:bg-slate-200 active:bg-slate-300 text-slate-600 flex items-center justify-center cursor-pointer shrink-0"
              >
                <Settings className="w-4 h-4" />
              </button>
            </div>
          </div>
        ) : (
          /* Admin Header Bar */
          <div className="max-w-md mx-auto p-2.5 flex items-center justify-between">
            <button
              type="button"
              onClick={() => setCurrentView('attendance')}
              className="flex items-center gap-1.5 text-xs font-semibold text-blue-600 hover:text-blue-800 active:text-blue-900 py-2 px-2.5 rounded-lg hover:bg-blue-50 cursor-pointer min-h-[44px]"
            >
              <ArrowLeft className="w-4 h-4" />
              <span>Tillbaka</span>
            </button>
            <h1 className="text-sm font-bold text-slate-800">Administration</h1>
            <div className="w-16"></div>
          </div>
        )}
      </header>

      {/* Main Content Area */}
      <main className="max-w-md mx-auto p-2.5 space-y-2">
        {currentView === 'attendance' ? (
          <>
            {/* Error notification if Apps Script call fails */}
            {errorMessage && (
              <div className="bg-rose-50 border border-rose-200 rounded-lg p-3 text-rose-900 text-xs space-y-1">
                <div className="flex items-center gap-1.5 font-semibold text-rose-800">
                  <AlertCircle className="w-4 h-4 text-rose-600 shrink-0" />
                  <span>Anslutningsfel</span>
                </div>
                <p className="text-[11px] font-mono leading-relaxed">{errorMessage}</p>
              </div>
            )}

            {/* Notice if training column does not exist yet */}
            {backendPayload && !matchedTraining && (
              <div className="bg-slate-50 border border-dashed border-slate-300 rounded-lg p-2.5 text-slate-500 text-xs text-center">
                Ingen registrerad kolumn för {selectedLocation} {selectedDateOption.dateFormatted} i kalkylarket.
              </div>
            )}

            {/* Empty state while outside Apps Script environment */}
            {membersList.length === 0 && !isLoading && (
              <div className="bg-white rounded-lg border border-slate-200 p-5 text-center space-y-2 text-xs text-slate-600">
                <p>
                  Medlemslistan läses in automatiskt i Google Apps Script Web App.
                </p>
                <button
                  type="button"
                  onClick={loadAttendanceFromBackend}
                  className="px-4 py-2 bg-blue-600 text-white rounded-md font-semibold cursor-pointer text-xs"
                >
                  Hämta från Google Sheet
                </button>
              </div>
            )}

            {/* Dense, fast, touch-friendly Member List */}
            <div className="space-y-1 bg-white rounded-xl border border-slate-200 overflow-hidden shadow-xs divide-y divide-slate-100">
              {membersList.map((member) => {
                const isPresent = Boolean(localAttendance[member.name]);

                return (
                  <button
                    key={member.name}
                    type="button"
                    onClick={() => toggleAttendance(member.name)}
                    className={`w-full min-h-[50px] px-3.5 py-2.5 flex items-center justify-between text-left transition-colors cursor-pointer select-none active:bg-slate-100 touch-manipulation ${
                      isPresent ? 'bg-emerald-50/70' : 'bg-white'
                    }`}
                  >
                    {/* Member Name */}
                    <span className={`text-sm leading-snug truncate pr-3 ${isPresent ? 'font-semibold text-emerald-950' : 'font-normal text-slate-800'}`}>
                      {member.name}
                    </span>

                    {/* Minimalist Checkbox Indicator */}
                    <div
                      className={`w-6 h-6 rounded border flex items-center justify-center shrink-0 transition-colors ${
                        isPresent
                          ? 'bg-emerald-600 border-emerald-600 text-white'
                          : 'bg-white border-slate-300'
                      }`}
                    >
                      {isPresent && <Check className="w-4 h-4 stroke-[3]" />}
                    </div>
                  </button>
                );
              })}
            </div>
          </>
        ) : (
          /* Step 2: Separate Admin View Structure (Read-only / Placeholder) */
          <div className="space-y-4 pt-2">
            <div className="bg-white rounded-xl border border-slate-200 p-4 shadow-xs space-y-3">
              <h2 className="text-xs font-bold uppercase tracking-wider text-slate-500">
                Träningspass
              </h2>
              <p className="text-xs text-slate-600 leading-relaxed">
                Här administreras träningspass och närvarokolumner i Google Sheet.
              </p>

              {/* Non-functional placeholder item */}
              <button
                type="button"
                disabled
                className="w-full min-h-[46px] px-4 py-2.5 rounded-lg border border-slate-200 bg-slate-50 text-slate-400 text-xs font-medium text-left flex items-center justify-between cursor-not-allowed"
              >
                <span>Hantera träningskolumner</span>
                <span className="text-[10px] bg-slate-200 text-slate-600 px-2 py-0.5 rounded font-mono">
                  Kommer snart
                </span>
              </button>
            </div>

            <div className="text-center text-[11px] text-slate-400">
              Inga ändringar görs i Google Sheet från denna vy ännu.
            </div>
          </div>
        )}
      </main>
    </div>
  );
}
