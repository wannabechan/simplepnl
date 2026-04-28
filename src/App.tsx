import { useEffect, useMemo, useRef, useState } from "react";

type EvidenceType = "taxInvoice" | "invoice" | "otherEvidence" | "cardSlip";
type ExpenseSource = "manual" | "cardStatement" | "purchaseAuto";

interface CategorySales {
  category: string;
  amount: number;
}

interface Expense {
  id: string;
  date: string;
  vendor: string;
  category: string;
  amount: number;
  note: string;
  source: ExpenseSource;
  evidence: Record<EvidenceType, boolean>;
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
  expenses: Expense[];
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

const APP_STATE_ID = "simplepnl-main-v2";
const money = new Intl.NumberFormat("ko-KR");
const nowId = () => `${Date.now()}-${Math.random().toString(16).slice(2)}`;
const normalize = (value: string) => value.trim().toLowerCase().replace(/\s+/g, "");

const readLocalState = (): MonthRecord[] => {
  const raw = localStorage.getItem(APP_STATE_ID);
  if (!raw) return [];
  return JSON.parse(raw) as MonthRecord[];
};

const normalizeStore = (store: StoreRecord): StoreRecord => ({
  ...store,
  manualRevenueEntries: Array.isArray(store.manualRevenueEntries) ? store.manualRevenueEntries : [],
});

const migrateMonthRecord = (raw: Record<string, unknown>): MonthRecord => {
  const stores = Array.isArray(raw.stores) ? (raw.stores as StoreRecord[]).map((s) => normalizeStore(s)) : [];
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
  try {
    const response = await fetch("/api/state");
    if (!response.ok) throw new Error("load failed");
    const payload = (await response.json()) as { data?: MonthRecord[] };
    return sanitizeMonths(payload.data);
  } catch {
    return sanitizeMonths(readLocalState());
  }
};

const saveState = async (months: MonthRecord[]) => {
  localStorage.setItem(APP_STATE_ID, JSON.stringify(months));
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
  expenses: [],
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
          저장: 각 입력 영역 하단의 저장 버튼으로 반영 · API + localStorage / 상태:{" "}
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

      {activeMonth && (
        <section className="panel">
          <h2>{activeMonth.label}</h2>

          <h3>매장별 비용 항목</h3>
          <table>
            <thead>
              <tr>
                <th>매장</th>
                <th>일자</th>
                <th>거래처</th>
                <th>분류</th>
                <th>공급가액</th>
                <th>증빙</th>
              </tr>
            </thead>
            <tbody>
              {activeMonth.stores.flatMap((store) =>
                store.expenses.map((expense) => (
                  <tr key={`${store.id}-${expense.id}`}>
                    <td>{store.name}</td>
                    <td>{expense.date}</td>
                    <td>{expense.vendor}</td>
                    <td>{expense.category}</td>
                    <td>{money.format(expense.amount)}</td>
                    <td>
                      {expense.evidence.taxInvoice ? "세금계산서 " : ""}
                      {expense.evidence.invoice ? "계산서 " : ""}
                      {expense.evidence.otherEvidence ? "기타증빙 " : ""}
                      {expense.evidence.cardSlip ? "카드매출전표 " : ""}
                    </td>
                  </tr>
                )),
              )}
            </tbody>
          </table>
        </section>
      )}
    </main>
  );
};

export default App;
