import { useEffect, useMemo, useRef, useState } from "react";

interface StoreRecord {
  id: string;
  name: string;
}

interface MonthRecord {
  id: string;
  label: string;
  stores: StoreRecord[];
}

const nowId = () => `${Date.now()}-${Math.random().toString(16).slice(2)}`;
const normalize = (value: string) => value.trim().toLowerCase().replace(/\s+/g, "");

const normalizeStore = (raw: unknown): StoreRecord => {
  const s = raw as Record<string, unknown>;
  return {
    id: typeof s.id === "string" ? s.id : nowId(),
    name: typeof s.name === "string" ? s.name : "",
  };
};

const migrateMonthRecord = (raw: Record<string, unknown>): MonthRecord => {
  const stores = Array.isArray(raw.stores) ? (raw.stores as unknown[]).map(normalizeStore) : [];
  const label = typeof raw.label === "string" ? raw.label : "";
  const id = typeof raw.id === "string" ? raw.id : nowId();
  return { id, label, stores };
};

const sanitizeMonths = (value: unknown): MonthRecord[] => {
  if (!Array.isArray(value)) return [];
  return value
    .filter((item) => item && typeof item === "object" && Array.isArray((item as { stores?: unknown }).stores))
    .map((item) => migrateMonthRecord(item as Record<string, unknown>));
};

const loadState = async (): Promise<MonthRecord[]> => {
  const response = await fetch("/api/state");
  if (!response.ok) {
    const details = await response.text();
    throw new Error(`상태 불러오기 실패 (${response.status}): ${details || response.statusText}`);
  }
  const payload = (await response.json()) as { data?: MonthRecord[] };
  return sanitizeMonths(payload.data);
};

const saveState = async (months: MonthRecord[]) => {
  const response = await fetch("/api/state", {
    method: "PUT",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ data: months }),
  });
  if (!response.ok) {
    const details = await response.text();
    throw new Error(`API save failed: ${response.status} ${details}`);
  }
};

const emptyStore = (name: string): StoreRecord => ({
  id: nowId(),
  name,
});

const App = () => {
  const [months, setMonths] = useState<MonthRecord[]>([]);
  const monthsRef = useRef(months);
  monthsRef.current = months;
  const [saveStatus, setSaveStatus] = useState<"idle" | "saving" | "saved" | "error">("idle");
  const [syncError, setSyncError] = useState("");
  const [monthLabel, setMonthLabel] = useState("");
  const [activeMonthId, setActiveMonthId] = useState<string | null>(null);
  const [storeTemplateName, setStoreTemplateName] = useState("");
  const [customStoreName, setCustomStoreName] = useState("");
  const [activeStoreId, setActiveStoreId] = useState<string | null>(null);
  useEffect(() => {
    const run = async () => {
      try {
        const parsed = await loadState();
        setMonths(parsed);
        setActiveMonthId(parsed[0]?.id ?? null);
        setActiveStoreId(parsed[0]?.stores[0]?.id ?? null);
      } catch (error) {
        const message = error instanceof Error ? error.message : "Unknown load error";
        setSyncError(message);
        setMonths([]);
        setActiveMonthId(null);
        setActiveStoreId(null);
      }
    };
    void run();
  }, []);

  const saveNow = async () => {
    setSaveStatus("saving");
    try {
      await saveState(monthsRef.current);
      setSaveStatus("saved");
      setSyncError("");
    } catch (error) {
      const message = error instanceof Error ? error.message : "Unknown save error";
      setSaveStatus("error");
      setSyncError(message);
    }
  };

  const activeMonth = useMemo(
    () => months.find((month) => month.id === activeMonthId) ?? null,
    [months, activeMonthId],
  );
  const updateMonth = (mutator: (month: MonthRecord) => MonthRecord) => {
    if (!activeMonthId) return;
    setMonths((prev) => prev.map((month) => (month.id === activeMonthId ? mutator(month) : month)));
  };

  const createMonth = () => {
    if (!monthLabel.trim()) return;
    const month: MonthRecord = {
      id: nowId(),
      label: monthLabel.trim(),
      stores: [],
    };
    setMonths((prev) => [...prev, month]);
    setActiveMonthId(month.id);
    setActiveStoreId(null);
    setMonthLabel("");
  };

  const createStoreInMonth = () => {
    if (!activeMonth) return;
    const pickedName =
      storeTemplateName === "__new__" ? customStoreName.trim() : storeTemplateName.trim();
    if (!pickedName) return;

    const existing = activeMonth.stores.find(
      (store) => normalize(store.name) === normalize(pickedName),
    );
    if (existing) {
      setActiveStoreId(existing.id);
      return;
    }

    const nextStore = emptyStore(pickedName);
    updateMonth((month) => ({ ...month, stores: [...month.stores, nextStore] }));
    setActiveStoreId(nextStore.id);
    setCustomStoreName("");
    setStoreTemplateName("");
  };

  const deleteMonth = (monthId: string) => {
    const month = months.find((item) => item.id === monthId);
    if (!month) return;
    if (month.stores.length > 0) {
      window.alert("포함된 매장이 남아있어 월을 삭제할 수 없습니다.");
      return;
    }
    if (!window.confirm("월데이터를 삭제하시겠습니까?")) return;

    const next = months.filter((item) => item.id !== monthId);
    setMonths(next);
    if (activeMonthId === monthId) {
      setActiveMonthId(next[0]?.id ?? null);
      setActiveStoreId(next[0]?.stores[0]?.id ?? null);
    }
  };

  const deleteStore = (storeId: string) => {
    if (!activeMonth) return;
    if (!window.confirm("매장을 삭제하시겠습니까?")) return;

    updateMonth((month) => {
      const stores = month.stores.filter((store) => store.id !== storeId);
      if (activeStoreId === storeId) {
        setActiveStoreId(stores[0]?.id ?? null);
      }
      return { ...month, stores };
    });
  };

  const previousMonthStoreOptions = useMemo(() => {
    if (!activeMonth) return [];
    const index = months.findIndex((month) => month.id === activeMonth.id);
    if (index <= 0) return [];
    const previousMonth = months[index - 1];
    return [...new Set(previousMonth.stores.map((store) => store.name))]
      .sort((a, b) => a.localeCompare(b, "en"));
  }, [months, activeMonth]);

  return (
    <main className="layout">
      <section className="panel">
        <h1>월·매장 구성</h1>
        <p className="muted">월을 만들고, 그 아래 매장을 추가합니다.</p>
        <p className="muted">
          저장: 저장 버튼으로 서버(DB)에 반영 · 상태:{" "}
          {saveStatus === "saving" ? "저장 중" : saveStatus === "saved" ? "저장 완료" : saveStatus === "error" ? "저장 실패" : "대기"}
        </p>
        {syncError && <p className="error">{syncError}</p>}

        <div className="row">
          <input placeholder="예: 2026-04" value={monthLabel} onChange={(e) => setMonthLabel(e.target.value)} />
          <button onClick={createMonth}>월 생성</button>
        </div>

        <div className="chip-wrap">
          {months.map((month) => (
            <div key={month.id} className="chip-item">
              <button
                className={month.id === activeMonthId ? "chip active" : "chip"}
                onClick={() => {
                  setActiveMonthId(month.id);
                  setActiveStoreId(month.stores[0]?.id ?? null);
                }}
              >
                {month.label}
              </button>
              <button
                className="chip-delete"
                onClick={(e) => {
                  e.stopPropagation();
                  deleteMonth(month.id);
                }}
              >
                x
              </button>
            </div>
          ))}
        </div>

        {activeMonth && (
          <>
            <h3>매장 생성</h3>
            <div className="row">
              <select value={storeTemplateName} onChange={(e) => setStoreTemplateName(e.target.value)}>
                <option value="">전월 매장 선택</option>
                {previousMonthStoreOptions.map((name) => (
                  <option key={name} value={name}>
                    {name}
                  </option>
                ))}
                <option value="__new__">신규 입력</option>
              </select>
              {storeTemplateName === "__new__" && (
                <input
                  placeholder="신규 매장명"
                  value={customStoreName}
                  onChange={(e) => setCustomStoreName(e.target.value)}
                />
              )}
              <button type="button" onClick={createStoreInMonth}>
                신규 생성
              </button>
            </div>
            <div className="chip-wrap">
              {activeMonth.stores.map((store) => (
                <div key={store.id} className="chip-item">
                  <button
                    className={store.id === activeStoreId ? "chip active" : "chip"}
                    onClick={() => setActiveStoreId(store.id)}
                  >
                    {store.name}
                  </button>
                  <button
                    className="chip-delete"
                    onClick={(e) => {
                      e.stopPropagation();
                      deleteStore(store.id);
                    }}
                  >
                    x
                  </button>
                </div>
              ))}
            </div>

            <div className="save-row">
              <button type="button" className="btn-save" onClick={() => void saveNow()}>
                저장
              </button>
              <span className="muted">월·매장 구성 변경을 서버에 저장합니다.</span>
            </div>
          </>
        )}
      </section>
    </main>
  );
};

export default App;
