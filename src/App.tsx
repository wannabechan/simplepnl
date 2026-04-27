import { useEffect, useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";

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

interface MonthRecord {
  id: string;
  label: string;
  stores: StoreRecord[];
  cardStatements: CardStatementRow[];
}

const APP_STATE_ID = "simplepnl-main-v2";
const money = new Intl.NumberFormat("ko-KR");
const nowId = () => `${Date.now()}-${Math.random().toString(16).slice(2)}`;
const today = () => new Date().toISOString().slice(0, 10);
const normalize = (value: string) => value.trim().toLowerCase().replace(/\s+/g, "");

const EXPENSE_CATEGORIES = [
  "광고홍보",
  "교통비",
  "기타",
  "매장유지비",
  "보험",
  "부자재",
  "세금",
  "소모품",
  "수도광열비",
  "식사비",
  "식자재",
  "외주용역비",
  "운송비",
  "음료",
  "인건비",
  "임관리비",
  "주류",
  "주방기물",
  "홀기물",
  "saas",
];

const inferAccount = (text: string) => {
  const n = normalize(text);
  if (n.includes("택시") || n.includes("교통")) return "교통비";
  if (n.includes("식자재")) return "식자재";
  if (n.includes("광고")) return "광고홍보";
  if (n.includes("인건")) return "인건비";
  if (n.includes("술") || n.includes("주류")) return "주류";
  if (n.includes("음료")) return "음료";
  if (n.includes("전기") || n.includes("가스") || n.includes("수도")) return "수도광열비";
  return "기타";
};

const toNumber = (value: unknown): number => {
  if (typeof value === "number") return Number.isFinite(value) ? value : 0;
  if (typeof value !== "string") return 0;
  const cleaned = value.replace(/[,원\s]/g, "");
  const parsed = Number(cleaned);
  return Number.isFinite(parsed) ? parsed : 0;
};

const getCell = (row: Record<string, unknown>, aliases: string[]) => {
  const entries = Object.entries(row);
  const found = entries.find(([key]) => aliases.some((alias) => normalize(key).includes(alias)));
  return found?.[1];
};

const parseFileRows = async (file: File): Promise<Record<string, unknown>[]> => {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: "array" });
  const firstSheet = workbook.SheetNames[0];
  if (!firstSheet) return [];
  const sheet = workbook.Sheets[firstSheet];
  return XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet, { defval: "" });
};

const parsePurchaseRows = (rows: Record<string, unknown>[]) =>
  rows
    .map((row) => ({
      vendor: String(getCell(row, ["거래처", "상호", "vendor", "공급자"]) ?? "").trim(),
      amount: toNumber(getCell(row, ["공급가액", "금액", "amount"])),
      category: String(getCell(row, ["품목", "구분", "category"]) ?? "미분류").trim(),
      date: String(getCell(row, ["일자", "작성일", "date"]) ?? "").trim(),
    }))
    .filter((item) => item.vendor || item.amount > 0);

const parseCardRows = (rows: Record<string, unknown>[]) =>
  rows
    .map((row) => {
      const vendor = String(getCell(row, ["가맹점", "사용처", "상호", "vendor"]) ?? "").trim();
      const rawCategory = String(getCell(row, ["업종", "분류", "category"]) ?? "").trim();
      const amount = toNumber(getCell(row, ["공급가액", "이용금액", "결제금액", "amount"]));
      const date = String(getCell(row, ["이용일", "승인일", "date"]) ?? "").trim();
      return { vendor, rawCategory, amount, date };
    })
    .filter((item) => item.vendor || item.amount > 0);

const readLocalState = (): MonthRecord[] => {
  const raw = localStorage.getItem(APP_STATE_ID);
  if (!raw) return [];
  return JSON.parse(raw) as MonthRecord[];
};

const normalizeStore = (store: StoreRecord): StoreRecord => ({
  ...store,
  manualRevenueEntries: Array.isArray(store.manualRevenueEntries) ? store.manualRevenueEntries : [],
});

const sanitizeMonths = (value: unknown): MonthRecord[] => {
  if (!Array.isArray(value)) return [];
  const valid = value.filter(
    (item) =>
      item &&
      typeof item === "object" &&
      Array.isArray((item as { stores?: unknown }).stores) &&
      Array.isArray((item as { cardStatements?: unknown }).cardStatements),
  ) as MonthRecord[];
  return valid.map((month) => ({
    ...month,
    stores: month.stores.map((store) => normalizeStore(store)),
  }));
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
  const [manualExpense, setManualExpense] = useState({
    date: today(),
    vendor: "",
    category: "기타",
    amount: "",
    note: "",
  });
  const [manualRevenue, setManualRevenue] = useState({
    date: today(),
    channel: "cash" as ManualRevenueChannel,
    amount: "",
    note: "",
  });

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
  const activeStore = useMemo(
    () => activeMonth?.stores.find((store) => store.id === activeStoreId) ?? null,
    [activeMonth, activeStoreId],
  );

  const updateMonth = (mutator: (month: MonthRecord) => MonthRecord) => {
    if (!activeMonthId) return;
    setMonths((prev) => prev.map((month) => (month.id === activeMonthId ? mutator(month) : month)));
  };

  const updateStore = (mutator: (store: StoreRecord) => StoreRecord) => {
    if (!activeMonthId || !activeStoreId) return;
    setMonths((prev) =>
      prev.map((month) => {
        if (month.id !== activeMonthId) return month;
        return {
          ...month,
          stores: month.stores.map((store) => (store.id === activeStoreId ? mutator(store) : store)),
        };
      }),
    );
  };

  const createMonth = () => {
    if (!monthLabel.trim()) return;
    const month: MonthRecord = { id: nowId(), label: monthLabel.trim(), stores: [], cardStatements: [] };
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
        cardStatements: month.cardStatements.map((row) =>
          row.assignedStoreId === storeId ? { ...row, assignedStoreId: "" } : row,
        ),
      };
    });
  };

  const rebuildCardExpenses = (month: MonthRecord): MonthRecord => {
    const stores = month.stores.map((store) => ({
      ...store,
      expenses: store.expenses.filter((expense) => expense.source !== "cardStatement"),
    }));

    month.cardStatements.forEach((row) => {
      if (!row.assignedStoreId) return;
      const target = stores.find((store) => store.id === row.assignedStoreId);
      if (!target) return;
      target.expenses.push({
        id: row.id,
        date: row.date,
        vendor: row.vendor,
        category: row.assignedAccount,
        amount: row.amount,
        note: "당월 카드내역서 자동 등록",
        source: "cardStatement",
        evidence: { taxInvoice: false, invoice: false, otherEvidence: false, cardSlip: true },
      });
    });

    return { ...month, stores };
  };

  const applyPurchaseEvidence = (month: MonthRecord, evidenceType: EvidenceType, rows: ReturnType<typeof parsePurchaseRows>) => {
    const stores = month.stores.map((store) => ({ ...store, expenses: [...store.expenses] }));
    rows.forEach((row) => {
      let matched = false;
      stores.forEach((store) => {
        const expense = store.expenses.find((item) => {
          const sameAmount = Math.abs(item.amount - row.amount) < 1;
          const sameVendor = row.vendor && item.vendor ? normalize(item.vendor) === normalize(row.vendor) : false;
          return sameAmount && (sameVendor || !row.vendor);
        });
        if (expense) {
          expense.evidence[evidenceType] = true;
          matched = true;
        }
      });
      if (!matched && stores[0]) {
        stores[0].expenses.push({
          id: nowId(),
          date: row.date,
          vendor: row.vendor || "자동등록",
          category: row.category || "기타",
          amount: row.amount,
          note: "매입증빙 자동등록",
          source: "purchaseAuto",
          evidence: {
            taxInvoice: evidenceType === "taxInvoice",
            invoice: evidenceType === "invoice",
            otherEvidence: evidenceType === "otherEvidence",
            cardSlip: false,
          },
        });
      }
    });
    return { ...month, stores };
  };

  const onUploadCardStatements = async (files: FileList | null) => {
    if (!files) return;
    if (activeMonth?.cardStatements.length && !window.confirm("기존 카드내역서 데이터와 설정을 삭제하고 다시 업로드하시겠습니까?")) {
      return;
    }
    const allRows: CardStatementRow[] = [];
    const fileList = Array.from(files);
    for (const file of fileList) {
      const rows = await parseFileRows(file);
      parseCardRows(rows).forEach((row) => {
        allRows.push({
          id: nowId(),
          date: row.date,
          vendor: row.vendor,
          amount: row.amount,
          rawCategory: row.rawCategory,
          assignedStoreId: activeMonth?.stores[0]?.id ?? "",
          assignedAccount: inferAccount(`${row.vendor} ${row.rawCategory}`),
        });
      });
    }
    updateMonth((month) => rebuildCardExpenses({ ...month, cardStatements: allRows }));
  };

  const onUploadPurchaseDocs = async (file: File, evidenceType: EvidenceType) => {
    const rows = await parseFileRows(file);
    const parsed = parsePurchaseRows(rows);
    updateMonth((month) => applyPurchaseEvidence(month, evidenceType, parsed));
  };

  const updateCardStatementRow = (rowId: string, field: "assignedStoreId" | "assignedAccount", value: string) => {
    updateMonth((month) => {
      const next = {
        ...month,
        cardStatements: month.cardStatements.map((row) => (row.id === rowId ? { ...row, [field]: value } : row)),
      };
      return rebuildCardExpenses(next);
    });
  };

  const onUploadSalesSummary = async (file: File) => {
    const rows = await parseFileRows(file);
    const cashSales = rows.reduce((sum, row) => sum + toNumber(getCell(row, ["현금", "cash"])), 0);
    const cardSales = rows.reduce((sum, row) => sum + toNumber(getCell(row, ["카드", "card"])), 0);
    updateStore((store) => ({ ...store, salesSummary: { cashSales, cardSales } }));
  };

  const onUploadCategorySales = async (file: File) => {
    const rows = await parseFileRows(file);
    const categorySales = rows
      .map((row) => ({
        category: String(getCell(row, ["카테고리", "분류", "category"]) ?? "").trim(),
        amount: toNumber(getCell(row, ["매출", "금액", "amount"])),
      }))
      .filter((item) => item.category && item.amount > 0);
    updateStore((store) => ({ ...store, categorySales }));
  };

  const addManualExpense = () => {
    if (!manualExpense.vendor.trim() || !manualExpense.amount.trim()) return;
    updateStore((store) => ({
      ...store,
      expenses: [
        ...store.expenses,
        {
          id: nowId(),
          date: manualExpense.date,
          vendor: manualExpense.vendor.trim(),
          category: manualExpense.category,
          amount: toNumber(manualExpense.amount),
          note: manualExpense.note.trim(),
          source: "manual",
          evidence: { taxInvoice: false, invoice: false, otherEvidence: false, cardSlip: false },
        },
      ],
    }));
    setManualExpense({ date: today(), vendor: "", category: "기타", amount: "", note: "" });
  };

  const addManualRevenue = () => {
    if (!manualRevenue.amount.trim()) return;
    const amount = toNumber(manualRevenue.amount);
    if (amount <= 0) return;
    updateStore((store) => ({
      ...store,
      manualRevenueEntries: [
        ...store.manualRevenueEntries,
        {
          id: nowId(),
          date: manualRevenue.date,
          channel: manualRevenue.channel,
          amount,
          note: manualRevenue.note.trim(),
        },
      ],
    }));
    setManualRevenue({ date: today(), channel: "cash", amount: "", note: "" });
  };

  const storeTotalRevenue = (store: StoreRecord) => {
    let cash = store.salesSummary.cashSales;
    let card = store.salesSummary.cardSales;
    let other = 0;
    store.manualRevenueEntries.forEach((entry) => {
      if (entry.channel === "cash") cash += entry.amount;
      else if (entry.channel === "card") card += entry.amount;
      else other += entry.amount;
    });
    return cash + card + other;
  };

  const allMonthExpenses = activeMonth?.stores.flatMap((store) => store.expenses) ?? [];
  const previousMonthStoreOptions = useMemo(() => {
    if (!activeMonth) return [];
    const index = months.findIndex((month) => month.id === activeMonth.id);
    if (index <= 0) return [];
    const previousMonth = months[index - 1];
    return [...new Set(previousMonth.stores.map((store) => store.name))]
      .sort((a, b) => a.localeCompare(b, "en"));
  }, [months, activeMonth]);
  const totalSales = activeMonth?.stores.reduce((sum, store) => sum + storeTotalRevenue(store), 0) ?? 0;
  const totalExpense = allMonthExpenses.reduce((sum, item) => sum + item.amount, 0);
  const operatingProfit = totalSales - totalExpense;

  return (
    <main className="layout">
      <section className="panel">
        <h1>월별 손익 리포트 (공급가액 기준)</h1>
        <p className="muted">구조: 월 생성 · 월 하위 매장 생성 · 월 공통 입력/매장별 입력</p>
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
      </section>

      {activeMonth && (
        <section className="panel">
          <h2>{activeMonth.label} 데이터 입력</h2>

          <h3>생성월 전체 영향 입력</h3>
          <div className="line line-2">
            <label className="uploader">
              당월 카드내역서 업로드 (복수 가능)
              <input type="file" multiple accept=".xlsx,.xls,.csv" onChange={(e) => void onUploadCardStatements(e.target.files)} />
            </label>
            <div className="uploader">
              카드내역서 행별 설정은 아래 테이블에서 처리
              <p className="muted">재업로드 시 기존 카드내역서 데이터/분배/계정 설정이 삭제됩니다.</p>
            </div>
          </div>

          <div className="line line-3">
            <label className="uploader">
              매입세금계산서 업로드
              <input type="file" accept=".xlsx,.xls,.csv" onChange={(e) => e.target.files?.[0] && void onUploadPurchaseDocs(e.target.files[0], "taxInvoice")} />
            </label>
            <label className="uploader">
              매입계산서 업로드
              <input type="file" accept=".xlsx,.xls,.csv" onChange={(e) => e.target.files?.[0] && void onUploadPurchaseDocs(e.target.files[0], "invoice")} />
            </label>
            <label className="uploader">
              매입증빙내역 업로드
              <input type="file" accept=".xlsx,.xls,.csv" onChange={(e) => e.target.files?.[0] && void onUploadPurchaseDocs(e.target.files[0], "otherEvidence")} />
            </label>
          </div>

          {activeMonth.cardStatements.length > 0 && (
            <>
              <h3>당월 카드내역서 행별 비용 분배/계정 설정</h3>
              <table>
                <thead>
                  <tr>
                    <th>일자</th>
                    <th>가맹점</th>
                    <th>공급가액</th>
                    <th>매장 할당</th>
                    <th>계정</th>
                  </tr>
                </thead>
                <tbody>
                  {activeMonth.cardStatements.map((row) => (
                    <tr key={row.id}>
                      <td>{row.date}</td>
                      <td>{row.vendor}</td>
                      <td>{money.format(row.amount)}</td>
                      <td>
                        <select
                          value={row.assignedStoreId}
                          onChange={(e) => updateCardStatementRow(row.id, "assignedStoreId", e.target.value)}
                        >
                          <option value="">미할당</option>
                          {activeMonth.stores.map((store) => (
                            <option key={store.id} value={store.id}>
                              {store.name}
                            </option>
                          ))}
                        </select>
                      </td>
                      <td>
                        <select
                          value={row.assignedAccount}
                          onChange={(e) => updateCardStatementRow(row.id, "assignedAccount", e.target.value)}
                        >
                          {EXPENSE_CATEGORIES.map((category) => (
                            <option key={category} value={category}>
                              {category}
                            </option>
                          ))}
                        </select>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </>
          )}

          <div className="save-row">
            <button type="button" className="btn-save" onClick={() => void saveNow()}>
              저장
            </button>
            <span className="muted">생성월 전체 영향 입력 내용을 서버에 저장합니다.</span>
          </div>

          <hr className="report-divider" />

          <h3>매장별 입력 영역</h3>
          <div className="row">
            <select
              value={storeTemplateName}
              onChange={(e) => setStoreTemplateName(e.target.value)}
            >
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
            <button onClick={createStoreInMonth}>신규 생성</button>
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

          {activeStore && (
            <>
              <h3>{activeStore.name} 입력</h3>
              <div className="line line-2">
                <label className="uploader">
                  매출표 업로드 (현금/카드 요약)
                  <input type="file" accept=".xlsx,.xls,.csv" onChange={(e) => e.target.files?.[0] && void onUploadSalesSummary(e.target.files[0])} />
                </label>
                <label className="uploader">
                  매출상품분석표 업로드 (카테고리 매출)
                  <input type="file" accept=".xlsx,.xls,.csv" onChange={(e) => e.target.files?.[0] && void onUploadCategorySales(e.target.files[0])} />
                </label>
              </div>

              <div className="line line-2">
                <div className="uploader">
                  기타 매출 수동 입력
                  <input type="date" value={manualRevenue.date} onChange={(e) => setManualRevenue((prev) => ({ ...prev, date: e.target.value }))} />
                  <select
                    value={manualRevenue.channel}
                    onChange={(e) =>
                      setManualRevenue((prev) => ({ ...prev, channel: e.target.value as ManualRevenueChannel }))
                    }
                  >
                    <option value="cash">현금매출</option>
                    <option value="card">카드매출</option>
                    <option value="other">기타매출</option>
                  </select>
                  <input placeholder="공급가액" value={manualRevenue.amount} onChange={(e) => setManualRevenue((prev) => ({ ...prev, amount: e.target.value }))} />
                  <input placeholder="비고" value={manualRevenue.note} onChange={(e) => setManualRevenue((prev) => ({ ...prev, note: e.target.value }))} />
                  <button type="button" onClick={addManualRevenue}>
                    매출 반영
                  </button>
                </div>
                <div className="uploader">
                  현금 결제 비용 수동 입력
                  <input type="date" value={manualExpense.date} onChange={(e) => setManualExpense((prev) => ({ ...prev, date: e.target.value }))} />
                  <input placeholder="거래처" value={manualExpense.vendor} onChange={(e) => setManualExpense((prev) => ({ ...prev, vendor: e.target.value }))} />
                  <select value={manualExpense.category} onChange={(e) => setManualExpense((prev) => ({ ...prev, category: e.target.value }))}>
                    {EXPENSE_CATEGORIES.map((category) => (
                      <option key={category} value={category}>
                        {category}
                      </option>
                    ))}
                  </select>
                  <input placeholder="공급가액" value={manualExpense.amount} onChange={(e) => setManualExpense((prev) => ({ ...prev, amount: e.target.value }))} />
                  <input placeholder="비고" value={manualExpense.note} onChange={(e) => setManualExpense((prev) => ({ ...prev, note: e.target.value }))} />
                  <button type="button" onClick={addManualExpense}>
                    비용 추가
                  </button>
                </div>
              </div>

              <div className="line line-1">
                <div className="uploader">
                  재고 입력
                  <input
                    placeholder="메뉴 재고"
                    value={activeStore.inventory.menu}
                    onChange={(e) => updateStore((store) => ({ ...store, inventory: { ...store.inventory, menu: e.target.value } }))}
                  />
                  <input
                    placeholder="음료 재고"
                    value={activeStore.inventory.beverage}
                    onChange={(e) => updateStore((store) => ({ ...store, inventory: { ...store.inventory, beverage: e.target.value } }))}
                  />
                </div>
              </div>
            </>
          )}

          <div className="save-row">
            <button type="button" className="btn-save" onClick={() => void saveNow()}>
              저장
            </button>
            <span className="muted">매장별 입력 내용을 서버에 저장합니다.</span>
          </div>

          <hr className="report-divider" />
          <h3>월 전체 손익 정리</h3>
          <div className="summary-grid">
            <div>월 총 매출: {money.format(totalSales)}원</div>
            <div>월 총 비용: {money.format(totalExpense)}원</div>
            <div className={operatingProfit >= 0 ? "profit" : "loss"}>월 영업손익: {money.format(operatingProfit)}원</div>
          </div>

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
