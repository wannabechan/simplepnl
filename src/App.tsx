import { useEffect, useMemo, useRef, useState, type ChangeEvent } from "react";
import * as XLSX from "xlsx";

interface SalesSummaryRow {
  id: string;
  businessDay: string;
  total: number;
  paymentAmount: number;
  supplyAmount: number;
  vat: number;
  discount: number;
  paymentMethod: string;
}

interface StoreRecord {
  id: string;
  name: string;
  salesSummaryRows?: SalesSummaryRow[];
}

interface MonthRecord {
  id: string;
  label: string;
  stores: StoreRecord[];
}

const nowId = () => `${Date.now()}-${Math.random().toString(16).slice(2)}`;
const normalize = (value: string) => value.trim().toLowerCase().replace(/\s+/g, "");

const toNumber = (value: unknown): number => {
  if (typeof value === "number") return Number.isFinite(value) ? value : 0;
  if (typeof value !== "string") return 0;
  const cleaned = value.replace(/[,원\s]/g, "");
  const parsed = Number(cleaned);
  return Number.isFinite(parsed) ? parsed : 0;
};

const findCell = (row: Record<string, unknown>, aliases: readonly string[]): unknown => {
  const keys = Object.keys(row);
  for (const alias of aliases) {
    const target = normalize(alias);
    const exact = keys.find((k) => normalize(k) === target);
    if (exact !== undefined) return row[exact];
  }
  for (const alias of aliases) {
    const target = normalize(alias);
    const hit = keys.find((k) => {
      const nk = normalize(k);
      return nk.includes(target) || target.includes(nk);
    });
    if (hit !== undefined) return row[hit];
  }
  return undefined;
};

const parseSalesFile = async (file: File): Promise<SalesSummaryRow[]> => {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: "array" });
  const firstSheet = workbook.SheetNames[0];
  if (!firstSheet) return [];
  const sheet = workbook.Sheets[firstSheet];
  const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet, { defval: "" });

  return rows
    .map((row) => {
      const businessDay = String(findCell(row, ["영업일", "일자", "날짜"]) ?? "").trim();
      const total = toNumber(findCell(row, ["합계"]));
      const paymentAmount = toNumber(findCell(row, ["결제금액"]));
      const supplyAmount = toNumber(findCell(row, ["공급가액"]));
      const vat = toNumber(findCell(row, ["부가세"]));
      const discount = toNumber(findCell(row, ["할인"]));
      const paymentMethod = String(findCell(row, ["결제수단"]) ?? "").trim();

      return {
        id: nowId(),
        businessDay,
        total,
        paymentAmount,
        supplyAmount,
        vat,
        discount,
        paymentMethod,
      };
    })
    .filter((r) => {
      return (
        r.businessDay !== "" ||
        r.paymentMethod !== "" ||
        r.total !== 0 ||
        r.paymentAmount !== 0 ||
        r.supplyAmount !== 0 ||
        r.vat !== 0 ||
        r.discount !== 0
      );
    });
};

const normalizeSalesRow = (raw: Partial<SalesSummaryRow>): SalesSummaryRow => ({
  id: typeof raw.id === "string" ? raw.id : nowId(),
  businessDay: typeof raw.businessDay === "string" ? raw.businessDay : String(raw.businessDay ?? ""),
  total: toNumber(raw.total),
  paymentAmount: toNumber(raw.paymentAmount),
  supplyAmount: toNumber(raw.supplyAmount),
  vat: toNumber(raw.vat),
  discount: toNumber(raw.discount),
  paymentMethod: typeof raw.paymentMethod === "string" ? raw.paymentMethod : String(raw.paymentMethod ?? ""),
});

const normalizeStore = (raw: unknown): StoreRecord => {
  const s = raw as Record<string, unknown>;
  let salesSummaryRows: SalesSummaryRow[] | undefined;
  if (Array.isArray(s.salesSummaryRows)) {
    salesSummaryRows = (s.salesSummaryRows as Partial<SalesSummaryRow>[]).map(normalizeSalesRow);
  }
  return {
    id: typeof s.id === "string" ? s.id : nowId(),
    name: typeof s.name === "string" ? s.name : "",
    salesSummaryRows,
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

const money = new Intl.NumberFormat("ko-KR");

const App = () => {
  const [months, setMonths] = useState<MonthRecord[]>([]);
  const monthsRef = useRef(months);
  monthsRef.current = months;
  const [syncError, setSyncError] = useState("");
  const [monthLabel, setMonthLabel] = useState("");
  const [activeMonthId, setActiveMonthId] = useState<string | null>(null);
  const [storeTemplateName, setStoreTemplateName] = useState("");
  const [customStoreName, setCustomStoreName] = useState("");
  const [activeStoreId, setActiveStoreId] = useState<string | null>(null);
  const [salesUploadBusy, setSalesUploadBusy] = useState(false);
  const salesFileInputRef = useRef<HTMLInputElement>(null);

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
    try {
      await saveState(monthsRef.current);
      setSyncError("");
    } catch (error) {
      const message = error instanceof Error ? error.message : "Unknown save error";
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

  const activeStore = useMemo(
    () => activeMonth?.stores.find((store) => store.id === activeStoreId) ?? null,
    [activeMonth, activeStoreId],
  );

  const persistSalesRows = async (rows: SalesSummaryRow[]) => {
    if (!activeMonthId || !activeStoreId) return;
    const next = monthsRef.current.map((month) => {
      if (month.id !== activeMonthId) return month;
      return {
        ...month,
        stores: month.stores.map((s) =>
          s.id === activeStoreId ? { ...s, salesSummaryRows: rows } : s,
        ),
      };
    });
    monthsRef.current = next;
    setMonths(next);
    setSyncError("");
    try {
      await saveState(next);
    } catch (error) {
      const message = error instanceof Error ? error.message : "Unknown save error";
      setSyncError(message);
    }
  };

  const onSalesFileChange = async (e: ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0] ?? null;
    e.target.value = "";
    if (!file || !activeMonthId || !activeStoreId || !activeStore) return;

    const ext = file.name.includes(".") ? file.name.split(".").pop()?.toLowerCase() : "";
    if (!ext || !["xls", "xlsx", "csv"].includes(ext)) {
      window.alert("xls, xlsx, csv 파일만 업로드할 수 있습니다.");
      return;
    }

    const hasSaved = (activeStore.salesSummaryRows?.length ?? 0) > 0;
    if (hasSaved && !window.confirm("저장 데이터가 있습니다. 재업로드 하시겠습니까?")) {
      return;
    }

    setSalesUploadBusy(true);
    try {
      const rows = await parseSalesFile(file);
      await persistSalesRows(rows);
    } catch (err) {
      window.alert(err instanceof Error ? err.message : "파일 분석에 실패했습니다.");
    } finally {
      setSalesUploadBusy(false);
    }
  };

  return (
    <main className="layout">
      <section className="panel">
        <h1>월·매장 구성</h1>
        <p className="muted">월을 만들고, 그 아래 매장을 추가합니다.</p>
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
            <br />

            <div className="save-row">
              <button type="button" className="btn-save" onClick={() => void saveNow()}>
                저장
              </button>
            </div>
          </>
        )}
      </section>

      {activeMonth && (
        <>
          <section className="panel">
            <div className="panel-heading">
              <h2>월 공통 데이터</h2>
              <span className="panel-heading-spacer" aria-hidden={true}>
                {"\u00A0\u00A0"}
              </span>
              <p className="muted card-meta">{activeMonth.label}</p>
            </div>
          </section>

          <section className="panel">
            <div className="panel-heading">
              <h2>매장별 데이터</h2>
              <span className="panel-heading-spacer" aria-hidden={true}>
                {"\u00A0\u00A0"}
              </span>
              <p className="muted card-meta">
                {activeStore ? `${activeMonth.label} / ${activeStore.name}` : activeMonth.label}
              </p>
            </div>

            <div className="sales-block">
              <br />
              <div className="sales-heading">
                <h3>매출 요약 데이터</h3>
                {activeStore && (
                  <>
                    <span className="panel-heading-spacer" aria-hidden={true}>
                      {"\u00A0\u00A0"}
                    </span>
                    <button
                      type="button"
                      disabled={salesUploadBusy}
                      onClick={() => salesFileInputRef.current?.click()}
                    >
                      {salesUploadBusy ? "…" : "upload"}
                    </button>
                  </>
                )}
              </div>
              {!activeStore ? (
                <p className="muted">매장을 선택한 뒤 파일을 업로드할 수 있습니다.</p>
              ) : (
                <>
                  <input
                    ref={salesFileInputRef}
                    type="file"
                    accept=".xls,.xlsx,.csv,application/vnd.ms-excel,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,text/csv"
                    className="visually-hidden"
                    onChange={(e) => void onSalesFileChange(e)}
                  />
                  {(activeStore.salesSummaryRows?.length ?? 0) > 0 && (
                    <table className="data-table">
                      <thead>
                        <tr>
                          <th>영업일</th>
                          <th>합계</th>
                          <th>결제금액</th>
                          <th>공급가액</th>
                          <th>부가세</th>
                          <th>할인</th>
                          <th>결제수단</th>
                        </tr>
                      </thead>
                      <tbody>
                        {activeStore.salesSummaryRows!.map((row) => (
                          <tr key={row.id}>
                            <td>{row.businessDay}</td>
                            <td>{money.format(row.total)}</td>
                            <td>{money.format(row.paymentAmount)}</td>
                            <td>{money.format(row.supplyAmount)}</td>
                            <td>{money.format(row.vat)}</td>
                            <td>{money.format(row.discount)}</td>
                            <td>{row.paymentMethod}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  )}
                </>
              )}
            </div>
          </section>
        </>
      )}
    </main>
  );
};

export default App;
