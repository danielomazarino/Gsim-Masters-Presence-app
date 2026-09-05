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
  const [isSaving, setIsSaving] = useState(false);
  const [saveSuccess, setSaveSuccess] = useState(false);
  const [errorMessage, setErrorMessage] = useState<string | null>(null);
  const [adminMessage, setAdminMessage] = useState<{ type: 'success' | 'error' | 'loading'; text: string } | null>(null);
  const [isDeleting, setIsDeleting] = useState(false);
  const [pendingDeleteTarget, setPendingDeleteTarget] = useState<{ date?: string; location?: string; header: string } | null>(null);

  // Local attendance state: key is member.name, value is boolean
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

  // Toggle single member attendance (LOCAL UI STATE)
  const toggleAttendance = (memberName: string) => {
    if (isSaving) return;
    setLocalAttendance((prev) => ({
      ...prev,
      [memberName]: !prev[memberName],
    }));
  };

  // Save current local attendance to Google Sheet via saveAttendance
  const handleSaveAttendance = () => {
    if (isSaving || isLoading) return;

    const googleScript = typeof window !== 'undefined' ? window.google?.script?.run : undefined;
    if (!googleScript) {
      setErrorMessage('google.script.run är inte tillgängligt i denna miljö.');
      return;
    }

    setIsSaving(true);
    setSaveSuccess(false);
    setErrorMessage(null);

    const members: MemberRow[] = backendPayload?.members || [];
    const attendancePayload = members.map((m) => ({
      row: m.row,
      member: m.name,
      present: Boolean(localAttendance[m.name]),
    }));

    googleScript
      .withSuccessHandler((response: unknown) => {
        setIsSaving(false);
        const res = response as { success?: boolean; error?: string; message?: string } | null;
        if (res && res.success === false) {
          setErrorMessage(`Kunde inte spara: ${res.error || res.message || 'Servern rapporterade ett fel'}. Försök igen.`);
          return;
        }

        setSaveSuccess(true);
        setTimeout(() => {
          setSaveSuccess(false);
        }, 3000);
        // Refresh backend data to verify synchronisation
        loadAttendanceFromBackend();
      })
      .withFailureHandler((error: Error | { message?: string }) => {
        setIsSaving(false);
        const msg = error && typeof error.message === 'string' ? error.message : String(error);
        setErrorMessage(`Kunde inte spara: ${msg || 'Okänt fel'}. Försök igen.`);
      })
      .saveAttendance(selectedDateOption.dateFormatted, selectedLocation, attendancePayload);
  };

  // Delete training column via deleteTrainingColumn (Step 4)
  const handleDeleteTrainingColumn = (target: { date?: string; location?: string; header: string }) => {
    const googleScript = typeof window !== 'undefined' ? window.google?.script?.run : undefined;
    if (!googleScript) {
      setAdminMessage({ type: 'error', text: 'google.script.run är inte tillgängligt i denna miljö.' });
      setPendingDeleteTarget(null);
      return;
    }

    let date = target.date;
    let location = target.location;

    if (!date || !location) {
      const parts = target.header.trim().split(/\s+/);
      if (parts.length >= 2) {
        location = parts[0];
        date = parts.slice(1).join(' ');
      }
    }

    if (!date || !location) {
      setAdminMessage({ type: 'error', text: `Kunde inte härleda datum och plats från "${target.header}".` });
      setPendingDeleteTarget(null);
      return;
    }

    setIsDeleting(true);
    setAdminMessage({ type: 'loading', text: `Tar bort ${target.header} från Google Sheet…` });

    googleScript
      .withSuccessHandler((response: unknown) => {
        setIsDeleting(false);
        setPendingDeleteTarget(null);
        const res = response as { success?: boolean; error?: string; message?: string } | null;

        if (res && res.success === false) {
          setAdminMessage({ type: 'error', text: `Kunde inte ta bort: ${res.error || res.message || 'Servern rapporterade ett fel.'}` });
          return;
        }

        setAdminMessage({ type: 'success', text: `${target.header} har tagits bort.` });
        loadAttendanceFromBackend();

        setTimeout(() => {
          setAdminMessage((prev) => (prev?.type === 'success' ? null : prev));
        }, 4000);
      })
      .withFailureHandler((error: Error | { message?: string }) => {
        setIsDeleting(false);
        setPendingDeleteTarget(null);
        const msg = error && typeof error.message === 'string' ? error.message : String(error);
        setAdminMessage({ type: 'error', text: `Kommunikationsfel vid radering: ${msg || 'Okänt fel'}.` });
      })
      .deleteTrainingColumn(date, location);
  };

  const trainingColumns: TrainingColumnInfo[] = backendPayload?.trainingColumns || [];

  const membersList: MemberRow[] = backendPayload?.members || [];

  // Discrete tally of present members
  const presentCount = useMemo(() => {
    return Object.values(localAttendance).filter(Boolean).length;
  }, [localAttendance]);

  return (
    <div className="min-h-screen bg-slate-100 text-slate-900 font-sans pb-10 select-none">
      {/* Sticky top control panel for iPhone quick access */}
      <header className="sticky top-0 z-30 bg-white border-b border-slate-200 shadow-xs">
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
                    disabled={isSaving}
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

            {/* Row 2: Compact Location Choice + Discrete Count Badge + Save Button + Admin Cog */}
            <div className="flex items-center gap-2">
              <div className="flex flex-1 gap-1.5">
                {(['Askim', 'Valhalla'] as LocationOption[]).map((loc) => {
                  const isSelected = selectedLocation === loc;
                  return (
                    <button
                      key={loc}
                      type="button"
                      onClick={() => setSelectedLocation(loc)}
                      disabled={isSaving}
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

              {/* Compact Save Button */}
              <button
                type="button"
                onClick={handleSaveAttendance}
                disabled={isSaving || isLoading}
                className={`min-h-[44px] px-3.5 rounded-lg text-xs font-semibold transition-colors cursor-pointer text-center shrink-0 flex items-center justify-center ${
                  saveSuccess
                    ? 'bg-emerald-600 text-white'
                    : isSaving
                    ? 'bg-slate-300 text-slate-500 cursor-not-allowed'
                    : 'bg-blue-600 text-white hover:bg-blue-700 active:bg-blue-800 shadow-xs'
                }`}
              >
                {isSaving ? 'Sparar…' : saveSuccess ? 'Sparat' : 'Spara'}
              </button>

              {/* Discrete Admin Gear Icon */}
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
            {/* Status message */}
            {isSaving && (
              <div className="bg-blue-50 border border-blue-200 rounded-lg p-2.5 text-blue-800 text-xs text-center font-medium">
                Sparar närvaro i Google Sheet…
              </div>
            )}
            {saveSuccess && (
              <div className="bg-emerald-50 border border-emerald-200 rounded-lg p-2.5 text-emerald-800 text-xs text-center font-medium">
                Närvaron sparades i Google Sheet.
              </div>
            )}
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

            {/* Missing column notice (informational, not blocking) */}
            {!matchedTraining && backendPayload && (
              <div className="bg-slate-50 border border-dashed border-slate-300 rounded-lg p-2.5 text-center text-xs text-slate-500">
                Ingen registrerad kolumn för {selectedLocation} {selectedDateOption.dateFormatted} i kalkylarket.
              </div>
            )}

            {/* List of members - Main screen focus with large 50px touch targets */}
            <div className="bg-white rounded-xl border border-slate-200 overflow-hidden shadow-xs">
              {membersList.length === 0 ? (
                <div className="p-8 text-center text-xs text-slate-500">
                  Inga medlemmar inlästa ännu. Kontrollera anslutningen till kalkylarket.
                </div>
              ) : (
                membersList.map((member) => {
                  const isPresent = Boolean(localAttendance[member.name]);
                  return (
                    <button
                      key={member.name}
                      type="button"
                      onClick={() => toggleAttendance(member.name)}
                      className={`w-full min-h-[50px] py-2.5 px-3.5 flex items-center justify-between border-b border-slate-100 last:border-b-0 cursor-pointer text-left transition-colors touch-manipulation ${
                        isPresent
                          ? 'bg-emerald-50/80 hover:bg-emerald-100/60 active:bg-emerald-100'
                          : 'bg-white hover:bg-slate-50 active:bg-slate-100'
                      }`}
                    >
                      <span
                        className={`text-sm tracking-tight pr-3 truncate select-none ${
                          isPresent ? 'font-semibold text-emerald-950' : 'font-normal text-slate-900'
                        }`}
                      >
                        {member.name}
                      </span>

                      {/* Visual-only check indicator */}
                      <div
                        className={`w-6 h-6 rounded border flex items-center justify-center shrink-0 transition-colors pointer-events-none ${
                          isPresent
                            ? 'bg-emerald-600 border-emerald-600 text-white'
                            : 'bg-white border-slate-300'
                        }`}
                      >
                        {isPresent && <Check className="w-4 h-4 stroke-[3]" />}
                      </div>
                    </button>
                  );
                })
              )}
            </div>
          </>
        ) : (
          /* Administration View - Training Column Management (Step 4) */
          <div className="space-y-3">
            {adminMessage && (
              <div
                className={`p-3 rounded-lg text-xs font-medium text-center ${
                  adminMessage.type === 'loading'
                    ? 'bg-blue-50 border border-blue-200 text-blue-800'
                    : adminMessage.type === 'success'
                    ? 'bg-emerald-50 border border-emerald-200 text-emerald-800'
                    : 'bg-rose-50 border border-rose-200 text-rose-800'
                }`}
              >
                {adminMessage.text}
              </div>
            )}

            <div className="bg-white rounded-xl border border-slate-200 p-4 shadow-xs space-y-3">
              <h2 className="text-xs font-bold uppercase tracking-wider text-slate-500">
                Träningspass
              </h2>
              <p className="text-xs text-slate-600 leading-relaxed">
                Befintliga träningskolumner i Google Sheet. Felaktigt skapade kolumner kan tas bort permanent härifrån.
              </p>

              <div className="space-y-2">
                {trainingColumns.length === 0 ? (
                  <div className="p-4 text-center text-xs text-slate-400 bg-slate-50 rounded-lg border border-slate-100">
                    Inga registrerade träningskolumner hittades.
                  </div>
                ) : (
                  trainingColumns.map((col) => (
                    <div
                      key={col.header}
                      className="flex items-center justify-between p-2.5 bg-slate-50 border border-slate-200 rounded-lg"
                    >
                      <div className="flex items-center gap-2">
                        <span className="text-xs font-semibold text-slate-800">{col.header}</span>
                        <span className="text-[10px] text-slate-500 bg-slate-200 py-0.5 px-1.5 rounded font-mono">
                          {col.colLetter ? `Kolumn ${col.colLetter}` : `Index ${col.index}`}
                        </span>
                      </div>
                      <button
                        type="button"
                        disabled={isDeleting}
                        onClick={() =>
                          setPendingDeleteTarget({
                            date: col.date,
                            location: col.location,
                            header: col.header,
                          })
                        }
                        className="text-xs font-semibold text-rose-600 bg-white border border-rose-200 hover:bg-rose-50 active:bg-rose-100 py-1.5 px-2.5 rounded-md cursor-pointer transition-colors disabled:opacity-50"
                      >
                        Ta bort
                      </button>
                    </div>
                  ))
                )}
              </div>
            </div>

            <p className="text-center text-[11px] text-slate-400 pt-2">
              Radering tar bort hela kolumnen i Google Sheet efter bekräftelse.
            </p>

            {/* Confirmation Modal for Delete */}
            {pendingDeleteTarget && (
              <div className="fixed inset-0 z-50 bg-slate-900/50 flex items-center justify-center p-4">
                <div className="bg-white rounded-xl border border-slate-200 max-w-xs w-full p-4 shadow-lg space-y-3">
                  <h3 className="text-sm font-bold text-slate-900">Ta bort träningspass?</h3>
                  <div className="inline-block text-xs font-semibold text-rose-700 bg-rose-50 border border-rose-200 px-2 py-1 rounded">
                    {pendingDeleteTarget.header}
                  </div>
                  <p className="text-xs text-slate-600 leading-relaxed">
                    Detta tar bort hela kolumnen och all registrerad närvaro i den.
                  </p>
                  <div className="flex justify-end gap-2 pt-2">
                    <button
                      type="button"
                      disabled={isDeleting}
                      onClick={() => setPendingDeleteTarget(null)}
                      className="text-xs font-semibold text-slate-700 bg-white border border-slate-300 py-2 px-3 rounded-lg hover:bg-slate-50 active:bg-slate-100 cursor-pointer disabled:opacity-50"
                    >
                      Avbryt
                    </button>
                    <button
                      type="button"
                      disabled={isDeleting}
                      onClick={() => handleDeleteTrainingColumn(pendingDeleteTarget)}
                      className="text-xs font-semibold text-white bg-rose-600 hover:bg-rose-700 active:bg-rose-800 py-2 px-3 rounded-lg cursor-pointer transition-colors disabled:bg-rose-400"
                    >
                      {isDeleting ? 'Tar bort…' : 'Ta bort'}
                    </button>
                  </div>
                </div>
              </div>
            )}
          </div>
        )}
      </main>
    </div>
  );
}
