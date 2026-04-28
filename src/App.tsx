import { useEffect, useMemo, useRef, useState } from "react";

interface CategorySales {
  category: string;
  amount: number;
}

type ManualRevenueChannel = "cash" | "card" | "other";

interface ManualRevenueEntry {
  id: string;
  date: string;
  channel: ManualRevenueChannel;
  amount: number;
  note: string;
}

interface StoreRecord {
  id: string;
  name: string;
  salesSummary: {
    cashSales: number;
    cardSales: number;
  };
  categorySales: CategorySales[];
  manualRevenueEntries: ManualRevenueEntry[];
  uploadedSalesSummaryFileName?: string;
  uploadedCategorySalesFileName?: string;
  inventory: {
    menu: string;
    beverage: string;
  };
}

interface CardStatementRow {
  id: string;
  date: string;
  vendor: string;
  amount: number;
  rawCategory: string;
  assignedStoreId: string;
  assignedAccount: string;
}

interface CardStatementFile {
  id: string;
  fileName: string;
  rows: CardStatementRow[];
}

interface MonthRecord {
  id: string;
  label: string;
  stores: StoreRecord[];
  cardStatementFiles: CardStatementFile[];
  purchaseTaxInvoiceFileName?: string;
  purchaseInvoiceFileName?: string;
  purchaseEvidenceFileName?: string;
}

const nowId = () => `${Date.now()}-${Math.random().toString(16).slice(2)}`;
const normalize = (value: string) => value.trim().toLowerCase().replace(/\s+/g, "");

const normalizeStore = (raw: unknown): StoreRecord => {
  const s = raw as Partial<StoreRecord> & Record<string, unknown>;
  const ss = s.salesSummary as { cashSales?: unknown; cardSales?: unknown } | undefined;
  const inv = s.inventory as { menu?: unknown; beverage?: unknown } | undefined;
  return {
    id: typeof s.id === "string" ? s.id : nowId(),
    name: typeof s.name === "string" ? s.name : "",
    salesSummary: {
      cashSales: Number(ss?.cashSales) || 0,
      cardSales: Number(ss?.cardSales) || 0,
    },
    categorySales: Array.isArray(s.categorySales) ? (s.categorySales as CategorySales[]) : [],
    manualRevenueEntries: Array.isArray(s.manualRevenueEntries) ? (s.manualRevenueEntries as ManualRevenueEntry[]) : [],
    uploadedSalesSummaryFileName:
      typeof s.uploadedSalesSummaryFileName === "string" ? s.uploadedSalesSummaryFileName : undefined,
    uploadedCategorySalesFileName:
      typeof s.uploadedCategorySalesFileName === "string" ? s.uploadedCategorySalesFileName : undefined,
    inventory: {
      menu: typeof inv?.menu === "string" ? inv.menu : "",
      beverage: typeof inv?.beverage === "string" ? inv.beverage : "",
    },
  };
};

const migrateMonthRecord = (raw: Record<string, unknown>): MonthRecord => {
  const stores = Array.isArray(raw.stores) ? (raw.stores as unknown[]).map(normalizeStore) : [];
  const label = typeof raw.label === "string" ? raw.label : "";
  const id = typeof raw.id === "string" ? raw.id : nowId();

  if (Array.isArray(raw.cardStatementFiles)) {
    const cardStatementFiles = (raw.cardStatementFiles as Partial<CardStatementFile>[]).map((f) => ({
      id: typeof f.id === "string" ? f.id : nowId(),
      fileName: typeof f.fileName === "string" ? f.fileName : "파일",
      rows: Array.isArray(f.rows) ? (f.rows as CardStatementRow[]) : [],
    }));
    return {
      id,
      label,
      stores,
      cardStatementFiles,
      purchaseTaxInvoiceFileName: typeof raw.purchaseTaxInvoiceFileName === "string" ? raw.purchaseTaxInvoiceFileName : undefined,
      purchaseInvoiceFileName: typeof raw.purchaseInvoiceFileName === "string" ? raw.purchaseInvoiceFileName : undefined,
      purchaseEvidenceFileName: typeof raw.purchaseEvidenceFileName === "string" ? raw.purchaseEvidenceFileName : undefined,
    };
  }

  const legacyRows = Array.isArray(raw.cardStatements) ? (raw.cardStatements as CardStatementRow[]) : [];
  return {
    id,
    label,
    stores,
    cardStatementFiles:
      legacyRows.length > 0
        ? [{ id: nowId(), fileName: "카드내역서(이전)", rows: legacyRows }]
        : [],
    purchaseTaxInvoiceFileName: typeof raw.purchaseTaxInvoiceFileName === "string" ? raw.purchaseTaxInvoiceFileName : undefined,
    purchaseInvoiceFileName: typeof raw.purchaseInvoiceFileName === "string" ? raw.purchaseInvoiceFileName : undefined,
    purchaseEvidenceFileName: typeof raw.purchaseEvidenceFileName === "string" ? raw.purchaseEvidenceFileName : undefined,
  };
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
  salesSummary: { cashSales: 0, cardSales: 0 },
  categorySales: [],
  manualRevenueEntries: [],
  uploadedSalesSummaryFileName: undefined,
  uploadedCategorySalesFileName: undefined,
  inventory: { menu: "", beverage: "" },
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
      cardStatementFiles: [],
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
      return {
        ...month,
        stores,
        cardStatementFiles: month.cardStatementFiles.map((file) => ({
          ...file,
          rows: file.rows.map((row) =>
            row.assignedStoreId === storeId ? { ...row, assignedStoreId: "" } : row,
          ),
        })),
      };
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
        <h1>월별 손익 리포트 (공급가액 기준)</h1>
        <p className="muted">구조: 월 생성 · 월 하위 매장 생성</p>
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
