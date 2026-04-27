import { useEffect, useMemo, useState } from "react";
import * as XLSX from "xlsx";

type EvidenceType = "taxInvoice" | "invoice" | "otherEvidence" | "cardSlip";
type UploadedDocType =
  | "cardPayments"
  | "purchaseTaxInvoice"
  | "purchaseInvoice"
  | "purchaseEvidence"
  | "cardStatement";

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
  evidence: Record<EvidenceType, boolean>;
}

interface MonthRecord {
  id: string;
  label: string;
  salesSummary: {
    cashSales: number;
    cardSales: number;
  };
  categorySales: CategorySales[];
  inventory: {
    menu: string;
    beverage: string;
  };
  expenses: Expense[];
  uploadedDocs: Record<UploadedDocType, number>;
}

interface Store {
  id: string;
  name: string;
  months: MonthRecord[];
}

const APP_STATE_ID = "simplepnl-main";
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

const money = new Intl.NumberFormat("ko-KR");

const nowId = () => `${Date.now()}-${Math.random().toString(16).slice(2)}`;
const today = () => new Date().toISOString().slice(0, 10);

const normalize = (value: string) => value.trim().toLowerCase().replace(/\s+/g, "");

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
    .map((row) => {
      const vendor = String(getCell(row, ["거래처", "상호", "vendor", "공급자"]) ?? "").trim();
      const amount = toNumber(getCell(row, ["공급가액", "금액", "amount"]));
      const category = String(getCell(row, ["품목", "구분", "category"]) ?? "미분류").trim();
      const date = String(getCell(row, ["일자", "작성일", "date"]) ?? "").trim();
      return { vendor, amount, category, date };
    })
    .filter((item) => item.vendor || item.amount > 0);

const parseCardStatementRows = (rows: Record<string, unknown>[]) =>
  rows
    .map((row) => {
      const vendor = String(getCell(row, ["가맹점", "사용처", "상호", "vendor"]) ?? "").trim();
      const amount = toNumber(getCell(row, ["이용금액", "결제금액", "공급가액", "amount"]));
      const category = String(getCell(row, ["업종", "분류", "category"]) ?? "카드비용").trim();
      const date = String(getCell(row, ["이용일", "승인일", "date"]) ?? "").trim();
      return { vendor, amount, category, date };
    })
    .filter((item) => item.vendor || item.amount > 0);

const readLocalState = (): Store[] => {
  const raw = localStorage.getItem(APP_STATE_ID);
  if (!raw) return [];
  return JSON.parse(raw) as Store[];
};

const loadState = async (): Promise<Store[]> => {
  try {
    const response = await fetch("/api/state");
    if (!response.ok) {
      throw new Error(`API load failed: ${response.status}`);
    }
    const payload = (await response.json()) as { data?: Store[] };
    return payload.data ?? [];
  } catch {
    return readLocalState();
  }
};

const saveState = async (stores: Store[]) => {
  localStorage.setItem(APP_STATE_ID, JSON.stringify(stores));

  try {
    const response = await fetch("/api/state", {
      method: "PUT",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ data: stores }),
    });
    if (!response.ok) {
      throw new Error(`API save failed: ${response.status}`);
    }
  } catch {
    // localStorage fallback is already saved above.
  }
};

const App = () => {
  const [stores, setStores] = useState<Store[]>([]);
  const [bootstrapped, setBootstrapped] = useState(false);
  const [saveStatus, setSaveStatus] = useState<"idle" | "saving" | "saved" | "error">("idle");
  const [syncError, setSyncError] = useState("");
  const [storeName, setStoreName] = useState("");
  const [activeStoreId, setActiveStoreId] = useState<string | null>(null);
  const [monthLabel, setMonthLabel] = useState("");
  const [activeMonthId, setActiveMonthId] = useState<string | null>(null);
  const [manualExpense, setManualExpense] = useState({
    date: today(),
    vendor: "",
    category: "기타",
    amount: "",
    note: "",
  });

  useEffect(() => {
    const run = async () => {
      try {
        const parsed = await loadState();
        setStores(parsed);
        setActiveStoreId(parsed[0]?.id ?? null);
        setActiveMonthId(parsed[0]?.months[0]?.id ?? null);
      } catch (error) {
        const message = error instanceof Error ? error.message : "Unknown load error";
        setSyncError(message);
      } finally {
        setBootstrapped(true);
      }
    };
    void run();
  }, []);

  useEffect(() => {
    if (!bootstrapped) return;
    const timer = window.setTimeout(() => {
      setSaveStatus("saving");
      saveState(stores)
        .then(() => {
          setSaveStatus("saved");
          setSyncError("");
        })
        .catch((error) => {
          const message = error instanceof Error ? error.message : "Unknown save error";
          setSaveStatus("error");
          setSyncError(message);
        });
    }, 350);
    return () => window.clearTimeout(timer);
  }, [stores, bootstrapped]);

  const activeStore = useMemo(
    () => stores.find((store) => store.id === activeStoreId) ?? null,
    [stores, activeStoreId],
  );
  const activeMonth = useMemo(
    () => activeStore?.months.find((month) => month.id === activeMonthId) ?? null,
    [activeStore, activeMonthId],
  );

  const upsertMonth = (mutator: (month: MonthRecord) => MonthRecord) => {
    if (!activeStoreId || !activeMonthId) return;
    setStores((prev) =>
      prev.map((store) => {
        if (store.id !== activeStoreId) return store;
        return {
          ...store,
          months: store.months.map((month) => (month.id === activeMonthId ? mutator(month) : month)),
        };
      }),
    );
  };

  const createStore = () => {
    if (!storeName.trim()) return;
    const newStore: Store = { id: nowId(), name: storeName.trim(), months: [] };
    setStores((prev) => [...prev, newStore]);
    setActiveStoreId(newStore.id);
    setActiveMonthId(null);
    setStoreName("");
  };

  const createMonth = () => {
    if (!activeStoreId || !monthLabel.trim()) return;
    const newMonth: MonthRecord = {
      id: nowId(),
      label: monthLabel.trim(),
      salesSummary: { cashSales: 0, cardSales: 0 },
      categorySales: [],
      inventory: { menu: "", beverage: "" },
      expenses: [],
      uploadedDocs: {
        cardPayments: 0,
        purchaseTaxInvoice: 0,
        purchaseInvoice: 0,
        purchaseEvidence: 0,
        cardStatement: 0,
      },
    };
    setStores((prev) =>
      prev.map((store) =>
        store.id === activeStoreId ? { ...store, months: [...store.months, newMonth] } : store,
      ),
    );
    setActiveMonthId(newMonth.id);
    setMonthLabel("");
  };

  const addManualExpense = () => {
    if (!manualExpense.vendor.trim() || !manualExpense.amount.trim()) return;
    upsertMonth((month) => ({
      ...month,
      expenses: [
        ...month.expenses,
        {
          id: nowId(),
          date: manualExpense.date,
          vendor: manualExpense.vendor.trim(),
          category: manualExpense.category.trim() || "기타",
          amount: toNumber(manualExpense.amount),
          note: manualExpense.note.trim(),
          evidence: { taxInvoice: false, invoice: false, otherEvidence: false, cardSlip: false },
        },
      ],
    }));
    setManualExpense({ date: today(), vendor: "", category: "기타", amount: "", note: "" });
  };

  const applyEvidenceRows = (
    month: MonthRecord,
    rows: { vendor: string; amount: number; category: string; date: string }[],
    type: EvidenceType,
    docType: UploadedDocType,
  ) => {
    const nextExpenses = [...month.expenses];

    rows.forEach((row) => {
      const match = nextExpenses.find((expense) => {
        const sameVendor =
          row.vendor && expense.vendor ? normalize(expense.vendor) === normalize(row.vendor) : false;
        const sameAmount = row.amount > 0 && Math.abs(expense.amount - row.amount) < 1;
        return sameAmount && (sameVendor || !row.vendor);
      });

      if (match) {
        match.evidence[type] = true;
        if (!match.date && row.date) match.date = row.date;
        if (match.category === "기타" && row.category) match.category = row.category;
      } else {
        nextExpenses.push({
          id: nowId(),
          date: row.date,
          vendor: row.vendor || `${docType}-자동등록`,
          category: row.category || "자동등록",
          amount: row.amount,
          note: `${docType} 자동 등록`,
          evidence: {
            taxInvoice: type === "taxInvoice",
            invoice: type === "invoice",
            otherEvidence: type === "otherEvidence",
            cardSlip: type === "cardSlip",
          },
        });
      }
    });

    return {
      ...month,
      expenses: nextExpenses,
      uploadedDocs: { ...month.uploadedDocs, [docType]: month.uploadedDocs[docType] + 1 },
    };
  };

  const onUploadSalesSummary = async (file: File) => {
    const rows = await parseFileRows(file);
    const cashSales = rows.reduce((sum, row) => sum + toNumber(getCell(row, ["현금", "cash"])), 0);
    const cardSales = rows.reduce((sum, row) => sum + toNumber(getCell(row, ["카드", "card"])), 0);
    upsertMonth((month) => ({
      ...month,
      salesSummary: { cashSales, cardSales },
    }));
  };

  const onUploadCategorySales = async (file: File) => {
    const rows = await parseFileRows(file);
    const categorySales = rows
      .map((row) => ({
        category: String(getCell(row, ["카테고리", "분류", "category"]) ?? "").trim(),
        amount: toNumber(getCell(row, ["매출", "금액", "amount"])),
      }))
      .filter((row) => row.category && row.amount > 0);
    upsertMonth((month) => ({ ...month, categorySales }));
  };

  const onUploadPurchaseDocs = async (file: File, type: EvidenceType, docType: UploadedDocType) => {
    const rows = await parseFileRows(file);
    const data = parsePurchaseRows(rows);
    upsertMonth((month) => applyEvidenceRows(month, data, type, docType));
  };

  const onUploadCardStatements = async (files: FileList | null) => {
    if (!files) return;
    const list = Array.from(files);
    for (const file of list) {
      const rows = await parseFileRows(file);
      const data = parseCardStatementRows(rows);
      upsertMonth((month) => applyEvidenceRows(month, data, "cardSlip", "cardStatement"));
    }
  };

  const totalSales = (activeMonth?.salesSummary.cashSales ?? 0) + (activeMonth?.salesSummary.cardSales ?? 0);
  const totalExpense = activeMonth?.expenses.reduce((sum, item) => sum + item.amount, 0) ?? 0;
  const operatingProfit = totalSales - totalExpense;

  return (
    <main className="layout">
      <section className="panel">
        <h1>매장 월별 손익 리포트 (공급가액 기준)</h1>
        <p className="muted">매장과 월을 생성하고 업로드/입력하면 매출, 비용, 손익이 자동 집계됩니다.</p>
        <p className="muted">
          저장 방식: API 영구 저장 + localStorage 캐시 / 상태:{" "}
          {saveStatus === "saving"
            ? "저장 중"
            : saveStatus === "saved"
              ? "저장 완료"
              : saveStatus === "error"
                ? "저장 실패"
                : "대기"}
        </p>
        {syncError && <p className="error">{syncError}</p>}

        <div className="row">
          <input
            placeholder="신규 매장명"
            value={storeName}
            onChange={(e) => setStoreName(e.target.value)}
          />
          <button onClick={createStore}>매장 생성</button>
        </div>

        <div className="chip-wrap">
          {stores.map((store) => (
            <button
              key={store.id}
              className={store.id === activeStoreId ? "chip active" : "chip"}
              onClick={() => {
                setActiveStoreId(store.id);
                setActiveMonthId(store.months[0]?.id ?? null);
              }}
            >
              {store.name}
            </button>
          ))}
        </div>
      </section>

      {activeStore && (
        <section className="panel">
          <h2>{activeStore.name} - 월 관리</h2>
          <div className="row">
            <input
              placeholder="예: 2026-04"
              value={monthLabel}
              onChange={(e) => setMonthLabel(e.target.value)}
            />
            <button onClick={createMonth}>월 생성</button>
          </div>

          <div className="chip-wrap">
            {activeStore.months.map((month) => (
              <button
                key={month.id}
                className={month.id === activeMonthId ? "chip active" : "chip"}
                onClick={() => setActiveMonthId(month.id)}
              >
                {month.label}
              </button>
            ))}
          </div>
        </section>
      )}

      {activeMonth && (
        <section className="panel">
          <h2>{activeMonth.label} 데이터 입력</h2>

          <div className="line line-2">
            <label className="uploader">
              매출표 업로드 (현금/카드 요약)
              <input type="file" accept=".xlsx,.xls,.csv" onChange={(e) => e.target.files?.[0] && onUploadSalesSummary(e.target.files[0])} />
            </label>

            <label className="uploader">
              매출상품분석표 업로드 (카테고리 매출)
              <input type="file" accept=".xlsx,.xls,.csv" onChange={(e) => e.target.files?.[0] && onUploadCategorySales(e.target.files[0])} />
            </label>
          </div>

          <div className="line line-2">
            <div className="uploader">
              재고 입력
              <input
                placeholder="메뉴 재고"
                value={activeMonth.inventory.menu}
                onChange={(e) =>
                  upsertMonth((month) => ({ ...month, inventory: { ...month.inventory, menu: e.target.value } }))
                }
              />
              <input
                placeholder="음료 재고"
                value={activeMonth.inventory.beverage}
                onChange={(e) =>
                  upsertMonth((month) => ({ ...month, inventory: { ...month.inventory, beverage: e.target.value } }))
                }
              />
            </div>

            <label className="uploader">
              당월 카드내역서 업로드 (복수 가능)
              <input type="file" multiple accept=".xlsx,.xls,.csv" onChange={(e) => onUploadCardStatements(e.target.files)} />
            </label>
          </div>

          <div className="line line-3">
            <label className="uploader">
              매입세금계산서 업로드
              <input type="file" accept=".xlsx,.xls,.csv" onChange={(e) => e.target.files?.[0] && onUploadPurchaseDocs(e.target.files[0], "taxInvoice", "purchaseTaxInvoice")} />
            </label>

            <label className="uploader">
              매입계산서 업로드
              <input type="file" accept=".xlsx,.xls,.csv" onChange={(e) => e.target.files?.[0] && onUploadPurchaseDocs(e.target.files[0], "invoice", "purchaseInvoice")} />
            </label>

            <label className="uploader">
              매입증빙내역 업로드
              <input type="file" accept=".xlsx,.xls,.csv" onChange={(e) => e.target.files?.[0] && onUploadPurchaseDocs(e.target.files[0], "otherEvidence", "purchaseEvidence")} />
            </label>
          </div>

          <div className="line line-1">
            <div className="uploader">
              현금 결제 비용 수동 입력
              <input
                type="date"
                placeholder="일자"
                value={manualExpense.date}
                onChange={(e) => setManualExpense((prev) => ({ ...prev, date: e.target.value }))}
              />
              <input
                placeholder="거래처"
                value={manualExpense.vendor}
                onChange={(e) => setManualExpense((prev) => ({ ...prev, vendor: e.target.value }))}
              />
              <select
                value={manualExpense.category}
                onChange={(e) => setManualExpense((prev) => ({ ...prev, category: e.target.value }))}
              >
                {EXPENSE_CATEGORIES.map((category) => (
                  <option key={category} value={category}>
                    {category}
                  </option>
                ))}
              </select>
              <input
                placeholder="공급가액"
                value={manualExpense.amount}
                onChange={(e) => setManualExpense((prev) => ({ ...prev, amount: e.target.value }))}
              />
              <input
                placeholder="비고"
                value={manualExpense.note}
                onChange={(e) => setManualExpense((prev) => ({ ...prev, note: e.target.value }))}
              />
              <button onClick={addManualExpense}>비용 추가</button>
            </div>
          </div>

          <h3>매출 항목 정리</h3>
          <div className="summary-grid">
            <div>현금 매출 합: {money.format(activeMonth.salesSummary.cashSales)}원</div>
            <div>카드 매출 합: {money.format(activeMonth.salesSummary.cardSales)}원</div>
            <div>총 매출: {money.format(totalSales)}원</div>
          </div>

          <h3>상품 카테고리별 매출</h3>
          <ul>
            {activeMonth.categorySales.map((item) => (
              <li key={`${item.category}-${item.amount}`}>
                {item.category}: {money.format(item.amount)}원
              </li>
            ))}
          </ul>

          <h3>비용 항목 정리</h3>
          <table>
            <thead>
              <tr>
                <th>일자</th>
                <th>거래처</th>
                <th>분류</th>
                <th>공급가액</th>
                <th>증빙</th>
                <th>비고</th>
              </tr>
            </thead>
            <tbody>
              {activeMonth.expenses.map((expense) => (
                <tr key={expense.id}>
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
                  <td>{expense.note}</td>
                </tr>
              ))}
            </tbody>
          </table>

          <h3>결과 손익 정리</h3>
          <div className="summary-grid">
            <div>총 매출: {money.format(totalSales)}원</div>
            <div>총 비용: {money.format(totalExpense)}원</div>
            <div className={operatingProfit >= 0 ? "profit" : "loss"}>
              영업손익: {money.format(operatingProfit)}원
            </div>
          </div>
        </section>
      )}
    </main>
  );
};

export default App;
