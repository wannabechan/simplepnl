import { useCallback, useEffect, useMemo, useRef, useState, type ChangeEvent } from "react";
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

interface ProductSummaryRow {
  id: string;
  category: string;
  quantity: number;
  totalSales: number;
  actualSales: number;
  discount: number;
}

interface StoreRecord {
  id: string;
  name: string;
  salesSummaryRows?: SalesSummaryRow[];
  productSummaryRows?: ProductSummaryRow[];
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

/** 상단에 제목·기간 등이 있어 1행이 헤더가 아닌 엑셀 대비: 영업일(또는 일자) 열이 있는 행을 헤더로 찾습니다. */
const findSalesHeaderRowIndex = (matrix: unknown[][]): number => {
  for (let i = 0; i < Math.min(matrix.length, 60); i++) {
    const row = matrix[i] ?? [];
    for (const cell of row) {
      const n = normalize(String(cell));
      if (n === "영업일" || n === "일자" || n === "날짜") return i;
    }
  }
  return 0;
};

const findProductHeaderRowIndex = (matrix: unknown[][]): number => {
  for (let i = 0; i < Math.min(matrix.length, 60); i++) {
    const row = matrix[i] ?? [];
    for (const cell of row) {
      const n = normalize(String(cell));
      if (n === "카테고리") return i;
    }
  }
  for (let i = 0; i < Math.min(matrix.length, 60); i++) {
    const row = matrix[i] ?? [];
    const cells = row.map((c) => normalize(String(c)));
    if (cells.includes("총매출") && (cells.includes("실매출") || cells.includes("수량"))) return i;
  }
  return 0;
};

const matrixRowsToObjects = (
  matrix: unknown[][],
  headerRowIdx: number,
  opts?: { skipRowsAfterHeader?: number },
): Record<string, unknown>[] => {
  const skipRowsAfterHeader = opts?.skipRowsAfterHeader ?? 0;
  const headerRow = matrix[headerRowIdx] ?? [];
  const headers = headerRow.map((h, c) => {
    const s = String(h ?? "").trim();
    return s !== "" ? s : `__empty_${c}`;
  });
  const objects: Record<string, unknown>[] = [];
  const dataStart = headerRowIdx + 1 + skipRowsAfterHeader;
  for (let r = dataStart; r < matrix.length; r++) {
    const row = matrix[r] ?? [];
    const obj: Record<string, unknown> = {};
    for (let c = 0; c < headers.length; c++) {
      obj[headers[c]!] = row[c] ?? "";
    }
    objects.push(obj);
  }
  return objects;
};

const formatBusinessDay = (value: unknown): string => {
  if (value === null || value === undefined || value === "") return "";
  if (typeof value === "number" && Number.isFinite(value)) {
    const epoch = Date.UTC(1899, 11, 30);
    const ms = epoch + value * 86400000;
    const d = new Date(ms);
    if (!Number.isNaN(d.getTime()) && value > 20000 && value < 120000) {
      const y = d.getUTCFullYear();
      const m = String(d.getUTCMonth() + 1).padStart(2, "0");
      const day = String(d.getUTCDate()).padStart(2, "0");
      return `${y}-${m}-${day}`;
    }
  }
  return String(value).trim();
};

const parseSalesFile = async (file: File): Promise<SalesSummaryRow[]> => {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: "array" });
  const firstSheet = workbook.SheetNames[0];
  if (!firstSheet) return [];
  const sheet = workbook.Sheets[firstSheet];
  const matrix = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, defval: "", raw: false });
  if (!matrix.length) return [];
  const headerRowIdx = findSalesHeaderRowIndex(matrix);
  const rows = matrixRowsToObjects(matrix, headerRowIdx);

  return rows
    .map((row) => {
      const businessDay = formatBusinessDay(findCell(row, ["영업일", "일자", "날짜"]));
      const total = toNumber(findCell(row, ["합계"]));
      const paymentAmount = toNumber(findCell(row, ["결제금액", "결제 금액"]));
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

const parseProductFile = async (file: File): Promise<ProductSummaryRow[]> => {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: "array" });
  const firstSheet = workbook.SheetNames[0];
  if (!firstSheet) return [];
  const sheet = workbook.Sheets[firstSheet];
  const matrix = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, defval: "", raw: false });
  if (!matrix.length) return [];
  const headerRowIdx = findProductHeaderRowIndex(matrix);
  /** 헤더 다음 2행(예: 소계·안내)은 제외하고 그 다음부터가 본문 데이터 */
  const rows = matrixRowsToObjects(matrix, headerRowIdx, { skipRowsAfterHeader: 2 });

  return rows
    .map((row) => {
      const category = String(findCell(row, ["카테고리", "분류", "품목군"]) ?? "").trim();
      const quantity = toNumber(findCell(row, ["수량", "판매수량", "판매 수량"]));
      const totalSales = toNumber(findCell(row, ["총매출", "총 매출"]));
      const actualSales = toNumber(findCell(row, ["실매출", "실 매출"]));
      const discount = toNumber(findCell(row, ["할인"]));

      return {
        id: nowId(),
        category,
        quantity,
        totalSales,
        actualSales,
        discount,
      };
    })
    .filter((r) => {
      return (
        r.category !== "" ||
        r.quantity !== 0 ||
        r.totalSales !== 0 ||
        r.actualSales !== 0 ||
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

const normalizeProductRow = (raw: Partial<ProductSummaryRow>): ProductSummaryRow => ({
  id: typeof raw.id === "string" ? raw.id : nowId(),
  category: typeof raw.category === "string" ? raw.category : String(raw.category ?? ""),
  quantity: toNumber(raw.quantity),
  totalSales: toNumber(raw.totalSales),
  actualSales: toNumber(raw.actualSales),
  discount: toNumber(raw.discount),
});

const normalizeStore = (raw: unknown): StoreRecord => {
  const s = raw as Record<string, unknown>;
  let salesSummaryRows: SalesSummaryRow[] | undefined;
  if (Array.isArray(s.salesSummaryRows)) {
    salesSummaryRows = (s.salesSummaryRows as Partial<SalesSummaryRow>[]).map(normalizeSalesRow);
  }
  let productSummaryRows: ProductSummaryRow[] | undefined;
  if (Array.isArray(s.productSummaryRows)) {
    productSummaryRows = (s.productSummaryRows as Partial<ProductSummaryRow>[]).map(normalizeProductRow);
  }
  return {
    id: typeof s.id === "string" ? s.id : nowId(),
    name: typeof s.name === "string" ? s.name : "",
    salesSummaryRows,
    productSummaryRows,
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

/** 월·매장 id·label·이름·목록만 로컬 기준으로 두고, 매출·상품 요약 행은 서버에 있던 매장 id 기준으로 유지합니다. */
const mergeStructureWithSalesFromServer = (
  localMonths: MonthRecord[],
  serverMonths: MonthRecord[],
): MonthRecord[] => {
  const salesByStoreId = new Map<string, SalesSummaryRow[] | undefined>();
  const productsByStoreId = new Map<string, ProductSummaryRow[] | undefined>();
  for (const m of serverMonths) {
    for (const s of m.stores) {
      salesByStoreId.set(s.id, s.salesSummaryRows);
      productsByStoreId.set(s.id, s.productSummaryRows);
    }
  }
  return localMonths.map((m) => ({
    id: m.id,
    label: m.label,
    stores: m.stores.map((s) => ({
      id: s.id,
      name: s.name,
      salesSummaryRows: salesByStoreId.get(s.id),
      productSummaryRows: productsByStoreId.get(s.id),
    })),
  }));
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
  /** 매출 업로드 모달: pick → reading → ready(저장 버튼) → saving → 닫힘 */
  const [salesModalOpen, setSalesModalOpen] = useState(false);
  const [salesModalPhase, setSalesModalPhase] = useState<"pick" | "reading" | "ready" | "saving" | "error">("pick");
  const [salesModalProgress, setSalesModalProgress] = useState("");
  const [salesParsedRows, setSalesParsedRows] = useState<SalesSummaryRow[] | null>(null);
  const salesModalFileInputRef = useRef<HTMLInputElement>(null);
  /** 상품 요약 업로드 모달 (매출 요약과 동일 단계) */
  const [productModalOpen, setProductModalOpen] = useState(false);
  const [productModalPhase, setProductModalPhase] = useState<"pick" | "reading" | "ready" | "saving" | "error">("pick");
  const [productModalProgress, setProductModalProgress] = useState("");
  const [productParsedRows, setProductParsedRows] = useState<ProductSummaryRow[] | null>(null);
  const productModalFileInputRef = useRef<HTMLInputElement>(null);
  /** 매장별 데이터 패널 탭 */
  const [storeDataTab, setStoreDataTab] = useState<"sales" | "products">("sales");

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
      const local = monthsRef.current;
      const server = await loadState();
      const merged = mergeStructureWithSalesFromServer(local, server);
      await saveState(merged);
      setMonths(merged);
      monthsRef.current = merged;
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

  const persistProductRows = async (
    rows: ProductSummaryRow[],
  ): Promise<{ ok: true } | { ok: false; message: string }> => {
    if (!activeMonthId || !activeStoreId) {
      return { ok: false, message: "월 또는 매장이 선택되지 않았습니다." };
    }
    const next = monthsRef.current.map((month) => {
      if (month.id !== activeMonthId) return month;
      return {
        ...month,
        stores: month.stores.map((s) =>
          s.id === activeStoreId ? { ...s, productSummaryRows: rows } : s,
        ),
      };
    });
    monthsRef.current = next;
    setMonths(next);
    setSyncError("");
    try {
      await saveState(next);
      return { ok: true };
    } catch (error) {
      const message = error instanceof Error ? error.message : "Unknown save error";
      setSyncError(message);
      return { ok: false, message };
    }
  };

  const persistSalesRows = async (
    rows: SalesSummaryRow[],
  ): Promise<{ ok: true } | { ok: false; message: string }> => {
    if (!activeMonthId || !activeStoreId) {
      return { ok: false, message: "월 또는 매장이 선택되지 않았습니다." };
    }
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
      return { ok: true };
    } catch (error) {
      const message = error instanceof Error ? error.message : "Unknown save error";
      setSyncError(message);
      return { ok: false, message };
    }
  };

  const resetSalesModal = useCallback(() => {
    setSalesModalOpen(false);
    setSalesModalPhase("pick");
    setSalesModalProgress("");
    setSalesParsedRows(null);
    if (salesModalFileInputRef.current) salesModalFileInputRef.current.value = "";
  }, []);

  const resetProductModal = useCallback(() => {
    setProductModalOpen(false);
    setProductModalPhase("pick");
    setProductModalProgress("");
    setProductParsedRows(null);
    if (productModalFileInputRef.current) productModalFileInputRef.current.value = "";
  }, []);

  const openProductUploadModal = () => {
    if (!activeMonthId || !activeStoreId || !activeStore) return;
    if (productModalOpen || salesModalOpen) return;
    setProductModalOpen(true);
    setProductModalPhase("pick");
    setProductModalProgress("파일을 선택해 주세요.");
    setProductParsedRows(null);
    if (productModalFileInputRef.current) productModalFileInputRef.current.value = "";
  };

  const closeProductModal = useCallback(() => {
    if (productModalPhase === "reading" || productModalPhase === "saving") return;
    if (productModalPhase === "ready" && productParsedRows !== null) {
      if (!window.confirm("저장하지 않고 닫으시겠습니까?")) return;
    }
    resetProductModal();
  }, [productModalPhase, productParsedRows, resetProductModal]);

  useEffect(() => {
    if (!productModalOpen) return;
    const onKey = (e: KeyboardEvent) => {
      if (e.key !== "Escape") return;
      closeProductModal();
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [productModalOpen, closeProductModal]);

  const onProductModalFileChange = async (e: ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0] ?? null;
    e.target.value = "";
    if (!file || !activeMonthId || !activeStoreId || !activeStore) return;

    const ext = file.name.includes(".") ? file.name.split(".").pop()?.toLowerCase() : "";
    if (!ext || !["xls", "xlsx", "csv"].includes(ext)) {
      window.alert("xls, xlsx, csv 파일만 업로드할 수 있습니다.");
      return;
    }

    const hasSaved = (activeStore.productSummaryRows?.length ?? 0) > 0;
    if (hasSaved && !window.confirm("저장 데이터가 있습니다. 재업로드 하시겠습니까?")) {
      return;
    }

    setProductModalPhase("reading");
    setProductModalProgress("파일을 읽는 중…");
    try {
      await new Promise((r) => requestAnimationFrame(() => r(undefined)));
      setProductModalProgress("시트에서 행을 불러오는 중…");
      const rows = await parseProductFile(file);
      setProductModalProgress("카테고리·수량·총매출·실매출·할인 컬럼을 확인하는 중…");
      await new Promise((r) => setTimeout(r, 200));
      setProductParsedRows(rows);
      setProductModalPhase("ready");
      setProductModalProgress(
        rows.length === 0
          ? "분석 결과 행이 없습니다. 저장하면 기존 상품 요약이 비워질 수 있습니다."
          : `${rows.length}건을 불러왔습니다. 저장을 누르면 서버에 반영됩니다.`,
      );
    } catch (err) {
      setProductModalPhase("error");
      setProductModalProgress(err instanceof Error ? err.message : "파일 분석에 실패했습니다.");
    }
  };

  const onProductModalSave = async () => {
    if (!productParsedRows || !activeMonthId || !activeStoreId) return;
    setProductModalPhase("saving");
    setProductModalProgress("추출한 컬럼으로 매장 데이터를 정리하는 중…");
    await new Promise((r) => setTimeout(r, 180));
    setProductModalProgress("데이터베이스에 저장하는 중…");
    const result = await persistProductRows(productParsedRows);
    if (!result.ok) {
      setProductModalPhase("ready");
      setProductModalProgress(`저장에 실패했습니다: ${result.message} (저장을 다시 눌러 재시도할 수 있습니다.)`);
      return;
    }
    setProductModalProgress("모든 작업이 완료되었습니다.");
    await new Promise((r) => setTimeout(r, 450));
    resetProductModal();
  };

  const openSalesUploadModal = () => {
    if (!activeMonthId || !activeStoreId || !activeStore) return;
    if (salesModalOpen || productModalOpen) return;
    setSalesModalOpen(true);
    setSalesModalPhase("pick");
    setSalesModalProgress("파일을 선택해 주세요.");
    setSalesParsedRows(null);
    if (salesModalFileInputRef.current) salesModalFileInputRef.current.value = "";
  };

  const closeSalesModal = useCallback(() => {
    if (salesModalPhase === "reading" || salesModalPhase === "saving") return;
    if (salesModalPhase === "ready" && salesParsedRows !== null) {
      if (!window.confirm("저장하지 않고 닫으시겠습니까?")) return;
    }
    resetSalesModal();
  }, [salesModalPhase, salesParsedRows, resetSalesModal]);

  useEffect(() => {
    if (!salesModalOpen) return;
    const onKey = (e: KeyboardEvent) => {
      if (e.key !== "Escape") return;
      closeSalesModal();
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [salesModalOpen, closeSalesModal]);

  const onSalesModalFileChange = async (e: ChangeEvent<HTMLInputElement>) => {
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

    setSalesModalPhase("reading");
    setSalesModalProgress("파일을 읽는 중…");
    try {
      await new Promise((r) => requestAnimationFrame(() => r(undefined)));
      setSalesModalProgress("시트에서 행을 불러오는 중…");
      const rows = await parseSalesFile(file);
      setSalesModalProgress("영업일·합계·결제금액 등 컬럼을 확인하는 중…");
      await new Promise((r) => setTimeout(r, 200));
      setSalesParsedRows(rows);
      setSalesModalPhase("ready");
      setSalesModalProgress(
        rows.length === 0
          ? "분석 결과 행이 없습니다. 저장하면 기존 매출 요약이 비워질 수 있습니다."
          : `${rows.length}건을 불러왔습니다. 저장을 누르면 서버에 반영됩니다.`,
      );
    } catch (err) {
      setSalesModalPhase("error");
      setSalesModalProgress(err instanceof Error ? err.message : "파일 분석에 실패했습니다.");
    }
  };

  const onSalesModalSave = async () => {
    if (!salesParsedRows || !activeMonthId || !activeStoreId) return;
    setSalesModalPhase("saving");
    setSalesModalProgress("추출한 컬럼으로 매장 데이터를 정리하는 중…");
    await new Promise((r) => setTimeout(r, 180));
    setSalesModalProgress("데이터베이스에 저장하는 중…");
    const result = await persistSalesRows(salesParsedRows);
    if (!result.ok) {
      setSalesModalPhase("ready");
      setSalesModalProgress(`저장에 실패했습니다: ${result.message} (저장을 다시 눌러 재시도할 수 있습니다.)`);
      return;
    }
    setSalesModalProgress("모든 작업이 완료되었습니다.");
    await new Promise((r) => setTimeout(r, 450));
    resetSalesModal();
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

            <div className="store-data-tabbed">
              <br />
              <div className="tab-bar" role="tablist" aria-label="매장별 데이터 구분">
                <button
                  type="button"
                  id="store-tab-sales"
                  role="tab"
                  aria-selected={storeDataTab === "sales"}
                  aria-controls="store-panel-sales"
                  className={`tab-trigger${storeDataTab === "sales" ? " tab-trigger-active" : ""}`}
                  onClick={() => setStoreDataTab("sales")}
                >
                  매출 요약
                </button>
                <button
                  type="button"
                  id="store-tab-products"
                  role="tab"
                  aria-selected={storeDataTab === "products"}
                  aria-controls="store-panel-products"
                  className={`tab-trigger${storeDataTab === "products" ? " tab-trigger-active" : ""}`}
                  onClick={() => setStoreDataTab("products")}
                >
                  상품 요약
                </button>
              </div>

              <div
                id="store-panel-sales"
                role="tabpanel"
                aria-labelledby="store-tab-sales"
                hidden={storeDataTab !== "sales"}
                className="tab-panel"
              >
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
                          disabled={salesModalOpen || productModalOpen}
                          onClick={openSalesUploadModal}
                        >
                          upload
                        </button>
                        <span aria-hidden={true}>{"\u00A0\u00A0"}</span>
                        {(activeStore.salesSummaryRows?.length ?? 0) > 0 && (
                          <span className="sales-sum-inline">
                            합계sum: {money.format(activeStore.salesSummaryRows!.reduce((a, r) => a + r.total, 0))}
                          </span>
                        )}
                      </>
                    )}
                  </div>
                  {!activeStore ? (
                    <p className="muted">매장을 선택한 뒤 파일을 업로드할 수 있습니다.</p>
                  ) : (
                    <>
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
              </div>

              <div
                id="store-panel-products"
                role="tabpanel"
                aria-labelledby="store-tab-products"
                hidden={storeDataTab !== "products"}
                className="tab-panel"
              >
                <div className="sales-block">
                  <br />
                  <div className="sales-heading">
                    <h3>상품 요약 데이터</h3>
                    {activeStore && (
                      <>
                        <span className="panel-heading-spacer" aria-hidden={true}>
                          {"\u00A0\u00A0"}
                        </span>
                        <button
                          type="button"
                          disabled={salesModalOpen || productModalOpen}
                          onClick={openProductUploadModal}
                        >
                          upload
                        </button>
                        <span aria-hidden={true}>{"\u00A0\u00A0"}</span>
                        {(activeStore.productSummaryRows?.length ?? 0) > 0 && (
                          <span className="sales-sum-inline">
                            총매출sum:{" "}
                            {money.format(
                              activeStore.productSummaryRows!.reduce((a, r) => a + r.totalSales, 0),
                            )}
                          </span>
                        )}
                      </>
                    )}
                  </div>
                  {!activeStore ? (
                    <p className="muted">매장을 선택한 뒤 파일을 업로드할 수 있습니다.</p>
                  ) : (
                    <>
                      {(activeStore.productSummaryRows?.length ?? 0) > 0 && (
                        <table className="data-table">
                          <thead>
                            <tr>
                              <th>카테고리</th>
                              <th>수량</th>
                              <th>총매출</th>
                              <th>실매출</th>
                              <th>할인</th>
                            </tr>
                          </thead>
                          <tbody>
                            {activeStore.productSummaryRows!.map((row) => (
                              <tr key={row.id}>
                                <td>{row.category}</td>
                                <td>{row.quantity.toLocaleString("ko-KR")}</td>
                                <td>{money.format(row.totalSales)}</td>
                                <td>{money.format(row.actualSales)}</td>
                                <td>{money.format(row.discount)}</td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      )}
                    </>
                  )}
                </div>
              </div>
            </div>
          </section>
        </>
      )}

      {productModalOpen && activeStore && (
        <div
          className="modal-backdrop"
          role="presentation"
          onMouseDown={(e) => {
            if (e.target === e.currentTarget) closeProductModal();
          }}
        >
          <div
            className="modal-panel"
            role="dialog"
            aria-modal="true"
            aria-labelledby="product-modal-title"
            onMouseDown={(e) => e.stopPropagation()}
          >
            <div className="modal-header">
              <h3 id="product-modal-title">상품 요약 업로드</h3>
              <button
                type="button"
                className="modal-close"
                aria-label="닫기"
                disabled={productModalPhase === "reading" || productModalPhase === "saving"}
                onClick={closeProductModal}
              >
                ×
              </button>
            </div>
            <p className="modal-progress" role="status" aria-live="polite">
              {productModalProgress}
            </p>
            {(productModalPhase === "pick" || productModalPhase === "error") && (
              <div className="modal-actions">
                <button type="button" onClick={() => productModalFileInputRef.current?.click()}>
                  파일 선택
                </button>
                <button type="button" className="btn-secondary" onClick={closeProductModal}>
                  취소
                </button>
              </div>
            )}
            {productModalPhase === "reading" && (
              <div className="modal-progress-bar" aria-hidden={true}>
                <span className="modal-progress-bar-fill modal-progress-bar-indeterminate" />
              </div>
            )}
            {productModalPhase === "ready" && (
              <div className="modal-actions">
                <button type="button" className="btn-save" onClick={() => void onProductModalSave()}>
                  저장
                </button>
                <button
                  type="button"
                  className="btn-secondary"
                  onClick={() => {
                    setProductModalPhase("pick");
                    setProductParsedRows(null);
                    setProductModalProgress("파일을 선택해 주세요.");
                  }}
                >
                  다른 파일
                </button>
              </div>
            )}
            {productModalPhase === "saving" && (
              <div className="modal-progress-bar" aria-hidden={true}>
                <span className="modal-progress-bar-fill modal-progress-bar-indeterminate" />
              </div>
            )}
            {productModalPhase === "error" && (
              <p className="error modal-error">파일을 다시 선택하거나 취소할 수 있습니다.</p>
            )}
            <input
              ref={productModalFileInputRef}
              type="file"
              accept=".xls,.xlsx,.csv,application/vnd.ms-excel,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,text/csv"
              className="visually-hidden"
              onChange={(e) => void onProductModalFileChange(e)}
            />
          </div>
        </div>
      )}

      {salesModalOpen && activeStore && (
        <div
          className="modal-backdrop"
          role="presentation"
          onMouseDown={(e) => {
            if (e.target === e.currentTarget) closeSalesModal();
          }}
        >
          <div
            className="modal-panel"
            role="dialog"
            aria-modal="true"
            aria-labelledby="sales-modal-title"
            onMouseDown={(e) => e.stopPropagation()}
          >
            <div className="modal-header">
              <h3 id="sales-modal-title">매출 요약 업로드</h3>
              <button
                type="button"
                className="modal-close"
                aria-label="닫기"
                disabled={salesModalPhase === "reading" || salesModalPhase === "saving"}
                onClick={closeSalesModal}
              >
                ×
              </button>
            </div>
            <p className="modal-progress" role="status" aria-live="polite">
              {salesModalProgress}
            </p>
            {(salesModalPhase === "pick" || salesModalPhase === "error") && (
              <div className="modal-actions">
                <button type="button" onClick={() => salesModalFileInputRef.current?.click()}>
                  파일 선택
                </button>
                <button type="button" className="btn-secondary" onClick={closeSalesModal}>
                  취소
                </button>
              </div>
            )}
            {salesModalPhase === "reading" && (
              <div className="modal-progress-bar" aria-hidden={true}>
                <span className="modal-progress-bar-fill modal-progress-bar-indeterminate" />
              </div>
            )}
            {salesModalPhase === "ready" && (
              <div className="modal-actions">
                <button type="button" className="btn-save" onClick={() => void onSalesModalSave()}>
                  저장
                </button>
                <button
                  type="button"
                  className="btn-secondary"
                  onClick={() => {
                    setSalesModalPhase("pick");
                    setSalesParsedRows(null);
                    setSalesModalProgress("파일을 선택해 주세요.");
                  }}
                >
                  다른 파일
                </button>
              </div>
            )}
            {salesModalPhase === "saving" && (
              <div className="modal-progress-bar" aria-hidden={true}>
                <span className="modal-progress-bar-fill modal-progress-bar-indeterminate" />
              </div>
            )}
            {salesModalPhase === "error" && (
              <p className="error modal-error">파일을 다시 선택하거나 취소할 수 있습니다.</p>
            )}
            <input
              ref={salesModalFileInputRef}
              type="file"
              accept=".xls,.xlsx,.csv,application/vnd.ms-excel,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,text/csv"
              className="visually-hidden"
              onChange={(e) => void onSalesModalFileChange(e)}
            />
          </div>
        </div>
      )}
    </main>
  );
};

export default App;
