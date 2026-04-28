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
  entryType?: "upload" | "manual";
  manualOrder?: number;
}

interface ProductSummaryRow {
  id: string;
  category: string;
  quantity: number;
  totalSales: number;
  actualSales: number;
  discount: number;
}

/** 비용 등록: 헤더(결제일·비용계정·…) 다음 행부터 데이터 (9컬럼) */
interface CostEntryRow {
  id: string;
  paymentDate: string;
  expenseKind: string;
  vendorName: string;
  totalAmount: number;
  supplyAmount: number;
  vat: number;
  taxMode: string;
  payStatus: string;
  memo: string;
  entryType?: "upload" | "manual";
  manualOrder?: number;
}

interface StoreRecord {
  id: string;
  name: string;
  salesSummaryRows?: SalesSummaryRow[];
  productSummaryRows?: ProductSummaryRow[];
  costEntryRows?: CostEntryRow[];
  menuInventory?: number;
  beverageInventory?: number;
}

interface MonthRecord {
  id: string;
  label: string;
  stores: StoreRecord[];
}

const nowId = () => `${Date.now()}-${Math.random().toString(16).slice(2)}`;
const normalize = (value: string) =>
  value
    .trim()
    .replace(/\uFEFF/g, "")
    .toLowerCase()
    .replace(/\s+/g, "")
    .replace(/[()[\]{}\-_/\\:.,]/g, "");

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

/** 비용 등록: 결제일 + 비용계정(또는 구 헤더) 행을 헤더로 인식 */
const findCostHeaderRowIndex = (matrix: unknown[][]): number => {
  for (let i = 0; i < Math.min(20, matrix.length); i++) {
    const row = matrix[i] ?? [];
    const cells = row.map((c) => normalize(String(c)));
    if (cells.includes("결제일") && (cells.includes("비용계정") || cells.includes("비용종류"))) return i;
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
  const numericValue =
    typeof value === "number"
      ? value
      : typeof value === "string" && value.trim() !== "" && Number.isFinite(Number(value))
        ? Number(value)
        : null;
  if (numericValue !== null && Number.isFinite(numericValue)) {
    const epoch = Date.UTC(1899, 11, 30);
    const ms = epoch + numericValue * 86400000;
    const d = new Date(ms);
    if (!Number.isNaN(d.getTime()) && numericValue > 20000 && numericValue < 120000) {
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
  const matrix = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, defval: "", raw: true });
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
        entryType: "upload" as const,
        manualOrder: 0,
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

/** 비용 파일: 헤더 행(기본 1행) 다음부터 본문 */
const parseCostFile = async (file: File): Promise<CostEntryRow[]> => {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: "array" });
  const firstSheet = workbook.SheetNames[0];
  if (!firstSheet) return [];
  const sheet = workbook.Sheets[firstSheet];
  const matrix = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, defval: "", raw: true });
  if (!matrix.length) return [];
  const headerRowIdx = findCostHeaderRowIndex(matrix);
  const rows = matrixRowsToObjects(matrix, headerRowIdx);

  return rows
    .map((row) => {
      const paymentDate = formatBusinessDay(findCell(row, ["결제일", "일자", "날짜"]));
      const expenseKind = String(
        findCell(row, [
          "비용계정",
          "비용 계정",
          "비용종류",
          "비용 종류",
          "비용 종류 콤보박스",
          "비용종류콤보박스",
        ]) ?? "",
      ).trim();
      const vendorName = String(findCell(row, ["업체명", "업체 명", "거래처"]) ?? "").trim();
      const totalAmount = toNumber(findCell(row, ["합계금액", "합계 금액"]));
      const supplyAmount = toNumber(findCell(row, ["공급가액"]));
      const vat = toNumber(findCell(row, ["부가세"]));
      const taxMode = String(
        findCell(row, [
          "과세여부",
          "과세 여부",
          "과세/면세",
          "과세면세",
          "과세 부과세",
          "과세/부과세",
          "과세부과세",
          "과세/부과세 콤보박스",
        ]) ?? "",
      ).trim();
      const payStatus = String(
        findCell(row, [
          "결제여부",
          "결제 여부",
          "결제/미결제",
          "결제미결제",
          "지급여부",
          "결제/미결제 콤보박스",
        ]) ?? "",
      ).trim();
      const memo = String(findCell(row, ["메모", "비고"]) ?? "").trim();

      return {
        id: nowId(),
        paymentDate,
        expenseKind,
        vendorName,
        totalAmount,
        supplyAmount,
        vat,
        taxMode,
        payStatus,
        memo,
        entryType: "upload" as const,
        manualOrder: 0,
      };
    })
    .filter((r) => {
      return (
        r.paymentDate !== "" ||
        r.expenseKind !== "" ||
        r.vendorName !== "" ||
        r.totalAmount !== 0 ||
        r.supplyAmount !== 0 ||
        r.vat !== 0 ||
        r.taxMode !== "" ||
        r.payStatus !== "" ||
        r.memo !== ""
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
  entryType: raw.entryType === "manual" ? "manual" : "upload",
  manualOrder: typeof raw.manualOrder === "number" && Number.isFinite(raw.manualOrder) ? raw.manualOrder : 0,
});

const normalizeProductRow = (raw: Partial<ProductSummaryRow>): ProductSummaryRow => ({
  id: typeof raw.id === "string" ? raw.id : nowId(),
  category: typeof raw.category === "string" ? raw.category : String(raw.category ?? ""),
  quantity: toNumber(raw.quantity),
  totalSales: toNumber(raw.totalSales),
  actualSales: toNumber(raw.actualSales),
  discount: toNumber(raw.discount),
});

const normalizeCostEntryRow = (raw: Partial<CostEntryRow>): CostEntryRow => ({
  id: typeof raw.id === "string" ? raw.id : nowId(),
  paymentDate:
    typeof raw.paymentDate === "string" ? raw.paymentDate : String(raw.paymentDate ?? ""),
  expenseKind:
    typeof raw.expenseKind === "string" ? raw.expenseKind : String(raw.expenseKind ?? ""),
  vendorName: typeof raw.vendorName === "string" ? raw.vendorName : String(raw.vendorName ?? ""),
  totalAmount: toNumber(raw.totalAmount),
  supplyAmount: toNumber(raw.supplyAmount),
  vat: toNumber(raw.vat),
  taxMode: typeof raw.taxMode === "string" ? raw.taxMode : String(raw.taxMode ?? ""),
  payStatus: typeof raw.payStatus === "string" ? raw.payStatus : String(raw.payStatus ?? ""),
  memo: typeof raw.memo === "string" ? raw.memo : String(raw.memo ?? ""),
  entryType: raw.entryType === "manual" ? "manual" : "upload",
  manualOrder: typeof raw.manualOrder === "number" && Number.isFinite(raw.manualOrder) ? raw.manualOrder : 0,
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
  let costEntryRows: CostEntryRow[] | undefined;
  if (Array.isArray(s.costEntryRows)) {
    costEntryRows = (s.costEntryRows as Partial<CostEntryRow>[]).map(normalizeCostEntryRow);
  }
  return {
    id: typeof s.id === "string" ? s.id : nowId(),
    name: typeof s.name === "string" ? s.name : "",
    salesSummaryRows,
    productSummaryRows,
    costEntryRows,
    menuInventory: toNumber(s.menuInventory),
    beverageInventory: toNumber(s.beverageInventory),
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

/** 월·매장 id·label·이름·목록만 로컬 기준으로 두고, 매출·상품·비용 행은 서버에 있던 매장 id 기준으로 유지합니다. */
const mergeStructureWithSalesFromServer = (
  localMonths: MonthRecord[],
  serverMonths: MonthRecord[],
): MonthRecord[] => {
  const salesByStoreId = new Map<string, SalesSummaryRow[] | undefined>();
  const productsByStoreId = new Map<string, ProductSummaryRow[] | undefined>();
  const costsByStoreId = new Map<string, CostEntryRow[] | undefined>();
  const menuInventoryByStoreId = new Map<string, number | undefined>();
  const beverageInventoryByStoreId = new Map<string, number | undefined>();
  for (const m of serverMonths) {
    for (const s of m.stores) {
      salesByStoreId.set(s.id, s.salesSummaryRows);
      productsByStoreId.set(s.id, s.productSummaryRows);
      costsByStoreId.set(s.id, s.costEntryRows);
      menuInventoryByStoreId.set(s.id, s.menuInventory);
      beverageInventoryByStoreId.set(s.id, s.beverageInventory);
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
      costEntryRows: costsByStoreId.get(s.id),
      menuInventory: menuInventoryByStoreId.get(s.id),
      beverageInventory: beverageInventoryByStoreId.get(s.id),
    })),
  }));
};

const emptyStore = (name: string): StoreRecord => ({
  id: nowId(),
  name,
});

const money = new Intl.NumberFormat("ko-KR");
const sortSalesRows = (rows: SalesSummaryRow[]): SalesSummaryRow[] => {
  const manual = rows
    .filter((r) => r.entryType === "manual")
    .sort((a, b) => (a.manualOrder ?? 0) - (b.manualOrder ?? 0));
  const uploaded = rows.filter((r) => r.entryType !== "manual");
  return [...manual, ...uploaded];
};
const sortCostRows = (rows: CostEntryRow[]): CostEntryRow[] => {
  const manual = rows
    .filter((r) => r.entryType === "manual")
    .sort((a, b) => (a.manualOrder ?? 0) - (b.manualOrder ?? 0));
  const uploaded = rows.filter((r) => r.entryType !== "manual");
  return [...manual, ...uploaded];
};
const todayYmd = () => {
  const now = new Date();
  const y = now.getFullYear();
  const m = String(now.getMonth() + 1).padStart(2, "0");
  const d = String(now.getDate()).padStart(2, "0");
  return `${y}-${m}-${d}`;
};
const COST_ACCOUNT_OPTIONS = [
  "광고홍보",
  "교통비",
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
  "_기타",
] as const;
const calcSupplyAndVat = (totalAmountRaw: string, taxModeRaw: string) => {
  const total = toNumber(totalAmountRaw);
  const roundedTotal = Math.round(total);
  if (taxModeRaw.trim() === "과세") {
    const supplyAmount = Math.round(total / 1.1);
    return { supplyAmount, vat: roundedTotal - supplyAmount };
  }
  return { supplyAmount: roundedTotal, vat: 0 };
};

interface SalesEntryDraft {
  businessDay: string;
  total: string;
  paymentAmount: string;
  supplyAmount: string;
  vat: string;
  discount: string;
  paymentMethod: "카드" | "현금" | "기타";
}

interface InventoryDraft {
  menuInventory: string;
  beverageInventory: string;
}

const calcSalesDerived = (paymentAmountRaw: string, discountRaw: string) => {
  const paymentAmount = toNumber(paymentAmountRaw);
  const discount = toNumber(discountRaw);
  const supplyAmount = Math.round(paymentAmount / 1.1);
  const vat = paymentAmount - supplyAmount;
  const total = paymentAmount + discount;
  return { paymentAmount, discount, supplyAmount, vat, total };
};

const emptySalesEntryDraft = (): SalesEntryDraft => {
  const base = calcSalesDerived("", "");
  return {
    businessDay: todayYmd(),
    total: String(base.total),
    paymentAmount: "",
    supplyAmount: String(base.supplyAmount),
    vat: String(base.vat),
    discount: "",
    paymentMethod: "카드",
  };
};

const emptyInventoryDraft = (): InventoryDraft => ({
  menuInventory: "",
  beverageInventory: "",
});

interface CostEntryDraft {
  paymentDate: string;
  expenseKind: string;
  vendorName: string;
  totalAmount: string;
  supplyAmount: string;
  vat: string;
  taxMode: string;
  payStatus: string;
  memo: string;
}

const emptyCostEntryDraft = (): CostEntryDraft => ({
  paymentDate: todayYmd(),
  expenseKind: "",
  vendorName: "",
  totalAmount: "",
  supplyAmount: "0",
  vat: "0",
  taxMode: "과세",
  payStatus: "결제",
  memo: "",
});

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
  const [salesEntryModalOpen, setSalesEntryModalOpen] = useState(false);
  const [salesEntryEditingId, setSalesEntryEditingId] = useState<string | null>(null);
  const [salesEntryDraft, setSalesEntryDraft] = useState<SalesEntryDraft>(() => emptySalesEntryDraft());
  const [salesEntryBusy, setSalesEntryBusy] = useState(false);
  const [inventoryModalOpen, setInventoryModalOpen] = useState(false);
  const [inventoryDraft, setInventoryDraft] = useState<InventoryDraft>(() => emptyInventoryDraft());
  const [inventoryBusy, setInventoryBusy] = useState(false);
  /** 상품 요약 업로드 모달 (매출 요약과 동일 단계) */
  const [productModalOpen, setProductModalOpen] = useState(false);
  const [productModalPhase, setProductModalPhase] = useState<"pick" | "reading" | "ready" | "saving" | "error">("pick");
  const [productModalProgress, setProductModalProgress] = useState("");
  const [productParsedRows, setProductParsedRows] = useState<ProductSummaryRow[] | null>(null);
  const productModalFileInputRef = useRef<HTMLInputElement>(null);
  /** 비용 등록 업로드 모달 */
  const [costModalOpen, setCostModalOpen] = useState(false);
  const [costModalPhase, setCostModalPhase] = useState<"pick" | "reading" | "ready" | "saving" | "error">("pick");
  const [costModalProgress, setCostModalProgress] = useState("");
  const [costParsedRows, setCostParsedRows] = useState<CostEntryRow[] | null>(null);
  const costModalFileInputRef = useRef<HTMLInputElement>(null);
  const [costEntryModalOpen, setCostEntryModalOpen] = useState(false);
  const [costEntryEditingId, setCostEntryEditingId] = useState<string | null>(null);
  const [costEntryDraft, setCostEntryDraft] = useState<CostEntryDraft>(() => emptyCostEntryDraft());
  const [costEntryBusy, setCostEntryBusy] = useState(false);
  /** 매장별 데이터 패널 탭 */
  const [storeDataTab, setStoreDataTab] = useState<"sales" | "products" | "costs">("sales");
  /** 첫 loadState 완료 여부 (로컬/배포 최초 접속 로딩) */
  const [remoteDataReady, setRemoteDataReady] = useState(false);
  /** 월·매장 저장(서버 병합) 진행 중 */
  const [saveMergeBusy, setSaveMergeBusy] = useState(false);

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
      } finally {
        setRemoteDataReady(true);
      }
    };
    void run();
  }, []);

  const saveNow = async () => {
    setSaveMergeBusy(true);
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
    } finally {
      setSaveMergeBusy(false);
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
  const salesRowsForDisplay = useMemo(
    () => sortSalesRows([...(activeStore?.salesSummaryRows ?? [])]),
    [activeStore?.salesSummaryRows],
  );
  const costRowsForDisplay = useMemo(
    () => sortCostRows([...(activeStore?.costEntryRows ?? [])]),
    [activeStore?.costEntryRows],
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

  const persistCostRows = async (
    rows: CostEntryRow[],
  ): Promise<{ ok: true } | { ok: false; message: string }> => {
    if (!activeMonthId || !activeStoreId) {
      return { ok: false, message: "월 또는 매장이 선택되지 않았습니다." };
    }
    const activeStoreCurrent =
      monthsRef.current
        .find((m) => m.id === activeMonthId)
        ?.stores.find((s) => s.id === activeStoreId) ?? null;
    const manualRows = (activeStoreCurrent?.costEntryRows ?? []).filter((r) => r.entryType === "manual");
    const mergedRows = sortCostRows([...manualRows, ...rows.map((r) => ({ ...r, entryType: "upload" as const }))]);
    const next = monthsRef.current.map((month) => {
      if (month.id !== activeMonthId) return month;
      return {
        ...month,
        stores: month.stores.map((s) =>
          s.id === activeStoreId ? { ...s, costEntryRows: mergedRows } : s,
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

  const persistAllCostRows = async (
    rows: CostEntryRow[],
  ): Promise<{ ok: true } | { ok: false; message: string }> => {
    if (!activeMonthId || !activeStoreId) {
      return { ok: false, message: "월 또는 매장이 선택되지 않았습니다." };
    }
    const sorted = sortCostRows(rows);
    const next = monthsRef.current.map((month) => {
      if (month.id !== activeMonthId) return month;
      return {
        ...month,
        stores: month.stores.map((s) =>
          s.id === activeStoreId ? { ...s, costEntryRows: sorted } : s,
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
    const activeStoreCurrent =
      monthsRef.current
        .find((m) => m.id === activeMonthId)
        ?.stores.find((s) => s.id === activeStoreId) ?? null;
    const manualRows = (activeStoreCurrent?.salesSummaryRows ?? []).filter((r) => r.entryType === "manual");
    const mergedRows = sortSalesRows([
      ...manualRows,
      ...rows.map((r) => ({ ...r, entryType: "upload" as const })),
    ]);
    const next = monthsRef.current.map((month) => {
      if (month.id !== activeMonthId) return month;
      return {
        ...month,
        stores: month.stores.map((s) =>
          s.id === activeStoreId ? { ...s, salesSummaryRows: mergedRows } : s,
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

  const persistAllSalesRows = async (
    rows: SalesSummaryRow[],
  ): Promise<{ ok: true } | { ok: false; message: string }> => {
    if (!activeMonthId || !activeStoreId) {
      return { ok: false, message: "월 또는 매장이 선택되지 않았습니다." };
    }
    const sorted = sortSalesRows(rows);
    const next = monthsRef.current.map((month) => {
      if (month.id !== activeMonthId) return month;
      return {
        ...month,
        stores: month.stores.map((s) =>
          s.id === activeStoreId ? { ...s, salesSummaryRows: sorted } : s,
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

  const persistStoreInventory = async (
    menuInventory: number,
    beverageInventory: number,
  ): Promise<{ ok: true } | { ok: false; message: string }> => {
    if (!activeMonthId || !activeStoreId) {
      return { ok: false, message: "월 또는 매장이 선택되지 않았습니다." };
    }
    const next = monthsRef.current.map((month) => {
      if (month.id !== activeMonthId) return month;
      return {
        ...month,
        stores: month.stores.map((s) =>
          s.id === activeStoreId ? { ...s, menuInventory, beverageInventory } : s,
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

  const resetCostModal = useCallback(() => {
    setCostModalOpen(false);
    setCostModalPhase("pick");
    setCostModalProgress("");
    setCostParsedRows(null);
    if (costModalFileInputRef.current) costModalFileInputRef.current.value = "";
  }, []);

  const openProductUploadModal = () => {
    if (!activeMonthId || !activeStoreId || !activeStore) return;
    if (productModalOpen || salesModalOpen || costModalOpen || salesEntryModalOpen) return;
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

  const openCostUploadModal = () => {
    if (!activeMonthId || !activeStoreId || !activeStore) return;
    if (costModalOpen || salesModalOpen || productModalOpen || salesEntryModalOpen) return;
    setCostModalOpen(true);
    setCostModalPhase("pick");
    setCostModalProgress("파일을 선택해 주세요.");
    setCostParsedRows(null);
    if (costModalFileInputRef.current) costModalFileInputRef.current.value = "";
  };

  const closeCostModal = useCallback(() => {
    if (costModalPhase === "reading" || costModalPhase === "saving") return;
    if (costModalPhase === "ready" && costParsedRows !== null) {
      if (!window.confirm("저장하지 않고 닫으시겠습니까?")) return;
    }
    resetCostModal();
  }, [costModalPhase, costParsedRows, resetCostModal]);

  useEffect(() => {
    if (!costModalOpen) return;
    const onKey = (e: KeyboardEvent) => {
      if (e.key !== "Escape") return;
      closeCostModal();
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [costModalOpen, closeCostModal]);

  const onCostModalFileChange = async (e: ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0] ?? null;
    e.target.value = "";
    if (!file || !activeMonthId || !activeStoreId || !activeStore) return;

    const ext = file.name.includes(".") ? file.name.split(".").pop()?.toLowerCase() : "";
    if (!ext || !["xls", "xlsx", "csv"].includes(ext)) {
      window.alert("xls, xlsx, csv 파일만 업로드할 수 있습니다.");
      return;
    }

    const hasSaved = (activeStore.costEntryRows?.length ?? 0) > 0;
    if (hasSaved && !window.confirm("저장 데이터가 있습니다. 재업로드 하시겠습니까?")) {
      return;
    }

    setCostModalPhase("reading");
    setCostModalProgress("파일을 읽는 중…");
    try {
      await new Promise((r) => requestAnimationFrame(() => r(undefined)));
      setCostModalProgress("시트에서 행을 불러오는 중…");
      const rows = await parseCostFile(file);
      setCostModalProgress("결제일·비용계정·업체명·합계금액 등 컬럼을 확인하는 중…");
      await new Promise((r) => setTimeout(r, 200));
      setCostParsedRows(rows);
      setCostModalPhase("ready");
      setCostModalProgress(
        rows.length === 0
          ? "분석 결과 행이 없습니다. 저장하면 기존 비용 등록이 비워질 수 있습니다."
          : `${rows.length}건을 불러왔습니다. 저장을 누르면 서버에 반영됩니다.`,
      );
    } catch (err) {
      setCostModalPhase("error");
      setCostModalProgress(err instanceof Error ? err.message : "파일 분석에 실패했습니다.");
    }
  };

  const onCostModalSave = async () => {
    if (!costParsedRows || !activeMonthId || !activeStoreId) return;
    setCostModalPhase("saving");
    setCostModalProgress("추출한 컬럼으로 매장 데이터를 정리하는 중…");
    await new Promise((r) => setTimeout(r, 180));
    setCostModalProgress("데이터베이스에 저장하는 중…");
    const result = await persistCostRows(costParsedRows);
    if (!result.ok) {
      setCostModalPhase("ready");
      setCostModalProgress(`저장에 실패했습니다: ${result.message} (저장을 다시 눌러 재시도할 수 있습니다.)`);
      return;
    }
    setCostModalProgress("모든 작업이 완료되었습니다.");
    await new Promise((r) => setTimeout(r, 450));
    resetCostModal();
  };

  const openCreateCostEntryModal = () => {
    if (!activeStore || costEntryBusy) return;
    setCostEntryEditingId(null);
    setCostEntryDraft(emptyCostEntryDraft());
    setCostEntryModalOpen(true);
  };

  const openEditCostEntryModal = (row: CostEntryRow) => {
    setCostEntryEditingId(row.id);
    setCostEntryDraft({
      paymentDate: row.paymentDate,
      expenseKind: row.expenseKind,
      vendorName: row.vendorName,
      totalAmount: String(row.totalAmount),
      supplyAmount: String(row.supplyAmount),
      vat: String(row.vat),
      taxMode: row.taxMode,
      payStatus: row.payStatus,
      memo: row.memo,
    });
    setCostEntryModalOpen(true);
  };

  const closeCostEntryModal = () => {
    if (costEntryBusy) return;
    setCostEntryModalOpen(false);
    setCostEntryEditingId(null);
    setCostEntryDraft(emptyCostEntryDraft());
  };
  const onCostEntryDraftChange = (patch: Partial<CostEntryDraft>) => {
    setCostEntryDraft((prev) => {
      const next: CostEntryDraft = { ...prev, ...patch };
      const derived = calcSupplyAndVat(next.totalAmount, next.taxMode);
      next.supplyAmount = String(derived.supplyAmount);
      next.vat = String(derived.vat);
      return next;
    });
  };

  const onCostEntrySave = async () => {
    if (!activeStore) return;
    setCostEntryBusy(true);
    const existing = [...(activeStore.costEntryRows ?? [])];
    const derived = calcSupplyAndVat(costEntryDraft.totalAmount, costEntryDraft.taxMode);
    const asRow: CostEntryRow = {
      id: costEntryEditingId ?? nowId(),
      paymentDate: formatBusinessDay(costEntryDraft.paymentDate),
      expenseKind: costEntryDraft.expenseKind.trim(),
      vendorName: costEntryDraft.vendorName.trim(),
      totalAmount: toNumber(costEntryDraft.totalAmount),
      supplyAmount: derived.supplyAmount,
      vat: derived.vat,
      taxMode: costEntryDraft.taxMode.trim(),
      payStatus: costEntryDraft.payStatus.trim(),
      memo: costEntryDraft.memo.trim(),
      entryType: "manual",
      manualOrder:
        costEntryEditingId === null
          ? existing
              .filter((r) => r.entryType === "manual")
              .reduce((max, r) => Math.max(max, r.manualOrder ?? 0), 0) + 1
          : existing.find((r) => r.id === costEntryEditingId)?.manualOrder ?? 0,
    };

    const next =
      costEntryEditingId === null
        ? [asRow, ...existing]
        : existing.map((row) => (row.id === costEntryEditingId ? asRow : row));
    const result = await persistAllCostRows(next);
    setCostEntryBusy(false);
    if (!result.ok) {
      window.alert(result.message);
      return;
    }
    closeCostEntryModal();
  };

  const onCostEntryDelete = async () => {
    if (!activeStore || !costEntryEditingId) return;
    if (!window.confirm("해당 비용을 삭제하사겠습니까?")) return;
    setCostEntryBusy(true);
    const existing = [...(activeStore.costEntryRows ?? [])];
    const next = existing.filter((row) => row.id !== costEntryEditingId);
    const result = await persistAllCostRows(next);
    setCostEntryBusy(false);
    if (!result.ok) {
      window.alert(result.message);
      return;
    }
    closeCostEntryModal();
  };

  const openSalesUploadModal = () => {
    if (!activeMonthId || !activeStoreId || !activeStore) return;
    if (salesModalOpen || productModalOpen || costModalOpen || salesEntryModalOpen) return;
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

  const openCreateSalesEntryModal = () => {
    if (!activeStore || salesEntryBusy) return;
    setSalesEntryEditingId(null);
    setSalesEntryDraft(emptySalesEntryDraft());
    setSalesEntryModalOpen(true);
  };

  const openEditSalesEntryModal = (row: SalesSummaryRow) => {
    const derived = calcSalesDerived(String(row.paymentAmount), String(row.discount));
    setSalesEntryEditingId(row.id);
    setSalesEntryDraft({
      businessDay: row.businessDay,
      total: String(derived.total),
      paymentAmount: String(row.paymentAmount),
      supplyAmount: String(derived.supplyAmount),
      vat: String(derived.vat),
      discount: String(row.discount),
      paymentMethod:
        row.paymentMethod === "현금" || row.paymentMethod === "기타" ? row.paymentMethod : "카드",
    });
    setSalesEntryModalOpen(true);
  };

  const closeSalesEntryModal = () => {
    if (salesEntryBusy) return;
    setSalesEntryModalOpen(false);
    setSalesEntryEditingId(null);
    setSalesEntryDraft(emptySalesEntryDraft());
  };

  const onSalesEntryDraftChange = (patch: Partial<SalesEntryDraft>) => {
    setSalesEntryDraft((prev) => {
      const next = { ...prev, ...patch };
      const derived = calcSalesDerived(next.paymentAmount, next.discount);
      next.total = String(derived.total);
      next.supplyAmount = String(derived.supplyAmount);
      next.vat = String(derived.vat);
      return next;
    });
  };

  const onSalesEntrySave = async () => {
    if (!activeStore) return;
    setSalesEntryBusy(true);
    const existing = [...(activeStore.salesSummaryRows ?? [])];
    const derived = calcSalesDerived(salesEntryDraft.paymentAmount, salesEntryDraft.discount);
    const asRow: SalesSummaryRow = {
      id: salesEntryEditingId ?? nowId(),
      businessDay: formatBusinessDay(salesEntryDraft.businessDay),
      total: derived.total,
      paymentAmount: derived.paymentAmount,
      supplyAmount: derived.supplyAmount,
      vat: derived.vat,
      discount: derived.discount,
      paymentMethod: salesEntryDraft.paymentMethod,
      entryType: "manual",
      manualOrder:
        salesEntryEditingId === null
          ? existing
              .filter((r) => r.entryType === "manual")
              .reduce((max, r) => Math.max(max, r.manualOrder ?? 0), 0) + 1
          : existing.find((r) => r.id === salesEntryEditingId)?.manualOrder ?? 0,
    };
    const next =
      salesEntryEditingId === null
        ? [asRow, ...existing]
        : existing.map((row) => (row.id === salesEntryEditingId ? asRow : row));
    const result = await persistAllSalesRows(next);
    setSalesEntryBusy(false);
    if (!result.ok) {
      window.alert(result.message);
      return;
    }
    closeSalesEntryModal();
  };

  const onSalesEntryDelete = async () => {
    if (!activeStore || !salesEntryEditingId) return;
    if (!window.confirm("해당 매출을 삭제하시겠습니까?")) return;
    setSalesEntryBusy(true);
    const next = [...(activeStore.salesSummaryRows ?? [])].filter((row) => row.id !== salesEntryEditingId);
    const result = await persistAllSalesRows(next);
    setSalesEntryBusy(false);
    if (!result.ok) {
      window.alert(result.message);
      return;
    }
    closeSalesEntryModal();
  };

  const openInventoryModal = () => {
    if (!activeStore) return;
    setInventoryDraft({
      menuInventory:
        activeStore.menuInventory !== undefined ? String(Math.round(activeStore.menuInventory)) : "",
      beverageInventory:
        activeStore.beverageInventory !== undefined ? String(Math.round(activeStore.beverageInventory)) : "",
    });
    setInventoryModalOpen(true);
  };

  const closeInventoryModal = () => {
    if (inventoryBusy) return;
    setInventoryModalOpen(false);
    setInventoryDraft(emptyInventoryDraft());
  };

  const onInventorySave = async () => {
    setInventoryBusy(true);
    const menuInventory = Math.round(toNumber(inventoryDraft.menuInventory));
    const beverageInventory = Math.round(toNumber(inventoryDraft.beverageInventory));
    const result = await persistStoreInventory(menuInventory, beverageInventory);
    setInventoryBusy(false);
    if (!result.ok) {
      window.alert(result.message);
      return;
    }
    closeInventoryModal();
  };

  const salesTabLoading = salesModalOpen && (salesModalPhase === "reading" || salesModalPhase === "saving");
  const productTabLoading =
    productModalOpen && (productModalPhase === "reading" || productModalPhase === "saving");
  const costTabLoading = costModalOpen && (costModalPhase === "reading" || costModalPhase === "saving");

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
              <button
                type="button"
                className="btn-save"
                disabled={saveMergeBusy}
                onClick={() => void saveNow()}
              >
                저장
              </button>
            </div>
          </>
        )}
      </section>

      {!remoteDataReady && (
        <section className="panel">
          <div className="panel-heading">
            <h2>매장별 데이터</h2>
            <span className="panel-heading-spacer" aria-hidden={true}>
              {"\u00A0\u00A0"}
            </span>
            <p className="muted card-meta">서버에서 불러오는 중…</p>
          </div>
          <div
            className="store-data-boot tab-panel-with-loader"
            role="status"
            aria-live="polite"
            aria-busy={true}
          >
            <div className="tab-loading-overlay boot-loading-overlay">
              <span className="tab-loading-spinner" />
            </div>
          </div>
        </section>
      )}

      {remoteDataReady && activeMonth && (
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

            <div className="store-data-shell tab-panel-with-loader">
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
                <button
                  type="button"
                  id="store-tab-costs"
                  role="tab"
                  aria-selected={storeDataTab === "costs"}
                  aria-controls="store-panel-costs"
                  className={`tab-trigger${storeDataTab === "costs" ? " tab-trigger-active" : ""}`}
                  onClick={() => setStoreDataTab("costs")}
                >
                  비용 등록
                </button>
              </div>

              <div
                id="store-panel-sales"
                role="tabpanel"
                aria-labelledby="store-tab-sales"
                hidden={storeDataTab !== "sales"}
                className="tab-panel tab-panel-with-loader"
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
                          disabled={
                            salesModalOpen || productModalOpen || costModalOpen || salesEntryModalOpen || inventoryModalOpen
                          }
                          onClick={openSalesUploadModal}
                        >
                          upload
                        </button>
                        <span aria-hidden={true}>{"\u00A0\u00A0"}</span>
                        {(activeStore.salesSummaryRows?.length ?? 0) > 0 && (
                          <span className="sales-sum-inline">
                            합계sum:{" "}
                            {money.format(
                              Math.round(
                                activeStore.salesSummaryRows!.reduce((a, r) => a + r.supplyAmount, 0),
                              ),
                            )}
                          </span>
                        )}
                        {(activeStore.menuInventory !== undefined ||
                          activeStore.beverageInventory !== undefined) && (
                          <span className="sales-sum-inline">
                            {"\u00A0\u00A0"}메뉴재고: {money.format(Math.round(activeStore.menuInventory ?? 0))}
                            {"\u00A0|\u00A0"}음료주류재고 : {money.format(Math.round(activeStore.beverageInventory ?? 0))}
                          </span>
                        )}
                        <span className="sales-heading-grow" />
                        <button
                          type="button"
                          className="btn-secondary"
                          disabled={
                            salesModalOpen ||
                            productModalOpen ||
                            costModalOpen ||
                            salesEntryModalOpen ||
                            inventoryModalOpen ||
                            salesEntryBusy ||
                            inventoryBusy
                          }
                          onClick={openInventoryModal}
                        >
                          재고 입력
                        </button>
                        <span aria-hidden={true}>{"\u00A0\u00A0"}</span>
                        <button
                          type="button"
                          className="btn-secondary"
                          disabled={
                            salesModalOpen ||
                            productModalOpen ||
                            costModalOpen ||
                            salesEntryModalOpen ||
                            inventoryModalOpen ||
                            salesEntryBusy ||
                            inventoryBusy
                          }
                          onClick={openCreateSalesEntryModal}
                        >
                          단건 등록
                        </button>
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
                            {salesRowsForDisplay.map((row) => (
                              <tr key={row.id}>
                                <td>
                                  {row.entryType === "manual" ? (
                                    <button
                                      type="button"
                                      className="cost-date-link"
                                      onClick={() => openEditSalesEntryModal(row)}
                                    >
                                      {row.businessDay}
                                    </button>
                                  ) : (
                                    row.businessDay
                                  )}
                                </td>
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
                {salesTabLoading && (
                  <div className="tab-loading-overlay" role="status" aria-live="polite">
                    <span className="tab-loading-spinner" />
                  </div>
                )}
              </div>

              <div
                id="store-panel-products"
                role="tabpanel"
                aria-labelledby="store-tab-products"
                hidden={storeDataTab !== "products"}
                className="tab-panel tab-panel-with-loader"
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
                          disabled={
                            salesModalOpen || productModalOpen || costModalOpen || salesEntryModalOpen || inventoryModalOpen
                          }
                          onClick={openProductUploadModal}
                        >
                          upload
                        </button>
                        <span aria-hidden={true}>{"\u00A0\u00A0"}</span>
                        {(activeStore.productSummaryRows?.length ?? 0) > 0 && (
                          <span className="sales-sum-inline">
                            총매출sum:{" "}
                            {money.format(
                              Math.round(
                                activeStore.productSummaryRows!.reduce((a, r) => a + r.actualSales, 0) / 1.1,
                              ),
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
                {productTabLoading && (
                  <div className="tab-loading-overlay" role="status" aria-live="polite">
                    <span className="tab-loading-spinner" />
                  </div>
                )}
              </div>

              <div
                id="store-panel-costs"
                role="tabpanel"
                aria-labelledby="store-tab-costs"
                hidden={storeDataTab !== "costs"}
                className="tab-panel tab-panel-with-loader"
              >
                <div className="sales-block">
                  <br />
                  <div className="sales-heading">
                    <h3>비용 등록 데이터</h3>
                    {activeStore && (
                      <>
                        <span className="panel-heading-spacer" aria-hidden={true}>
                          {"\u00A0\u00A0"}
                        </span>
                        <button
                          type="button"
                          disabled={
                            salesModalOpen || productModalOpen || costModalOpen || salesEntryModalOpen || inventoryModalOpen
                          }
                          onClick={openCostUploadModal}
                        >
                          upload
                        </button>
                        <span aria-hidden={true}>{"\u00A0\u00A0"}</span>
                        {(activeStore.costEntryRows?.length ?? 0) > 0 && (
                          <span className="sales-sum-inline">
                            공급가액sum:{" "}
                            {money.format(
                              Math.round(activeStore.costEntryRows!.reduce((acc, row) => acc + row.supplyAmount, 0)),
                            )}
                          </span>
                        )}
                        <span className="sales-heading-grow" />
                        <button
                          type="button"
                          className="btn-secondary"
                          disabled={salesModalOpen || productModalOpen || costModalOpen || costEntryBusy}
                          onClick={openCreateCostEntryModal}
                        >
                          단건 등록
                        </button>
                      </>
                    )}
                  </div>
                  {!activeStore ? (
                    <p className="muted">매장을 선택한 뒤 파일을 업로드할 수 있습니다.</p>
                  ) : (
                    <div className="data-table-scroll">
                      {(activeStore.costEntryRows?.length ?? 0) > 0 && (
                        <table className="data-table data-table-costs">
                          <thead>
                            <tr>
                              <th>결제일</th>
                              <th>비용계정</th>
                              <th>업체명</th>
                              <th>합계금액</th>
                              <th>공급가액</th>
                              <th>부가세</th>
                              <th>과세여부</th>
                              <th>결제여부</th>
                              <th>메모</th>
                            </tr>
                          </thead>
                          <tbody>
                            {costRowsForDisplay.map((row) => (
                              <tr key={row.id}>
                                <td>
                                  {row.entryType === "manual" ? (
                                    <button
                                      type="button"
                                      className="cost-date-link"
                                      onClick={() => openEditCostEntryModal(row)}
                                    >
                                      {row.paymentDate}
                                    </button>
                                  ) : (
                                    row.paymentDate
                                  )}
                                </td>
                                <td>{row.expenseKind}</td>
                                <td>{row.vendorName}</td>
                                <td>{money.format(Math.round(row.totalAmount))}</td>
                                <td>{money.format(Math.round(row.supplyAmount))}</td>
                                <td>{money.format(Math.round(row.vat))}</td>
                                <td>{row.taxMode}</td>
                                <td>{row.payStatus}</td>
                                <td>{row.memo}</td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      )}
                    </div>
                  )}
                </div>
                {costTabLoading && (
                  <div className="tab-loading-overlay" role="status" aria-live="polite">
                    <span className="tab-loading-spinner" />
                  </div>
                )}
              </div>
            </div>
            {saveMergeBusy && (
              <div className="tab-loading-overlay" role="status" aria-live="polite">
                <span className="tab-loading-spinner" />
              </div>
            )}
            </div>
          </section>
        </>
      )}

      {salesEntryModalOpen && activeStore && (
        <div
          className="modal-backdrop"
          role="presentation"
          onMouseDown={(e) => {
            if (e.target === e.currentTarget) closeSalesEntryModal();
          }}
        >
          <div
            className="modal-panel"
            role="dialog"
            aria-modal="true"
            aria-labelledby="sales-entry-modal-title"
            onMouseDown={(e) => e.stopPropagation()}
          >
            <div className="modal-header">
              <h3 id="sales-entry-modal-title">{salesEntryEditingId ? "매출 수정" : "매출 단건 등록"}</h3>
              <button type="button" className="modal-close" aria-label="닫기" onClick={closeSalesEntryModal}>
                ×
              </button>
            </div>
            <div className="cost-form-grid">
              <label>
                영업일
                <input
                  type="date"
                  value={salesEntryDraft.businessDay}
                  onChange={(e) => onSalesEntryDraftChange({ businessDay: e.target.value })}
                />
              </label>
              <label>
                결제수단
                <select
                  value={salesEntryDraft.paymentMethod}
                  onChange={(e) =>
                    onSalesEntryDraftChange({
                      paymentMethod: e.target.value as SalesEntryDraft["paymentMethod"],
                    })
                  }
                >
                  <option value="카드">카드</option>
                  <option value="현금">현금</option>
                  <option value="기타">기타</option>
                </select>
              </label>
              <label>
                결제금액
                <input
                  value={salesEntryDraft.paymentAmount}
                  onChange={(e) => onSalesEntryDraftChange({ paymentAmount: e.target.value })}
                />
              </label>
              <label>
                할인
                <input
                  value={salesEntryDraft.discount}
                  onChange={(e) => onSalesEntryDraftChange({ discount: e.target.value })}
                />
              </label>
              <label>
                합계
                <input value={salesEntryDraft.total} disabled />
              </label>
              <label>
                공급가액
                <input value={salesEntryDraft.supplyAmount} disabled />
              </label>
              <label>
                부가세
                <input value={salesEntryDraft.vat} disabled />
              </label>
            </div>
            <div className="modal-actions">
              {salesEntryEditingId && (
                <button
                  type="button"
                  className="btn-danger"
                  disabled={salesEntryBusy}
                  onClick={() => void onSalesEntryDelete()}
                >
                  삭제
                </button>
              )}
              <button type="button" className="btn-save" disabled={salesEntryBusy} onClick={() => void onSalesEntrySave()}>
                저장
              </button>
              <button type="button" className="btn-secondary" disabled={salesEntryBusy} onClick={closeSalesEntryModal}>
                취소
              </button>
            </div>
          </div>
        </div>
      )}

      {inventoryModalOpen && activeStore && (
        <div
          className="modal-backdrop"
          role="presentation"
          onMouseDown={(e) => {
            if (e.target === e.currentTarget) closeInventoryModal();
          }}
        >
          <div
            className="modal-panel"
            role="dialog"
            aria-modal="true"
            aria-labelledby="inventory-modal-title"
            onMouseDown={(e) => e.stopPropagation()}
          >
            <div className="modal-header">
              <h3 id="inventory-modal-title">재고 입력</h3>
              <button type="button" className="modal-close" aria-label="닫기" onClick={closeInventoryModal}>
                ×
              </button>
            </div>
            <div className="cost-form-grid">
              <label>
                메뉴 재고 (공급가액)
                <input
                  value={inventoryDraft.menuInventory}
                  onChange={(e) => setInventoryDraft((p) => ({ ...p, menuInventory: e.target.value }))}
                />
              </label>
              <label>
                음료주류 재고 (공급가액)
                <input
                  value={inventoryDraft.beverageInventory}
                  onChange={(e) => setInventoryDraft((p) => ({ ...p, beverageInventory: e.target.value }))}
                />
              </label>
            </div>
            <div className="modal-actions">
              <button type="button" className="btn-save" disabled={inventoryBusy} onClick={() => void onInventorySave()}>
                저장
              </button>
              <button type="button" className="btn-secondary" disabled={inventoryBusy} onClick={closeInventoryModal}>
                취소
              </button>
            </div>
            {inventoryBusy && (
              <div className="modal-progress-bar" aria-hidden={true}>
                <span className="modal-progress-bar-fill modal-progress-bar-indeterminate" />
              </div>
            )}
          </div>
        </div>
      )}

      {costEntryModalOpen && activeStore && (
        <div
          className="modal-backdrop"
          role="presentation"
          onMouseDown={(e) => {
            if (e.target === e.currentTarget) closeCostEntryModal();
          }}
        >
          <div
            className="modal-panel"
            role="dialog"
            aria-modal="true"
            aria-labelledby="cost-entry-modal-title"
            onMouseDown={(e) => e.stopPropagation()}
          >
            <div className="modal-header">
              <h3 id="cost-entry-modal-title">{costEntryEditingId ? "비용 수정" : "비용 단건 등록"}</h3>
              <button type="button" className="modal-close" aria-label="닫기" onClick={closeCostEntryModal}>
                ×
              </button>
            </div>
            <div className="cost-form-grid">
              <label>
                결제일
                <input
                  type="date"
                  value={costEntryDraft.paymentDate}
                  onChange={(e) => onCostEntryDraftChange({ paymentDate: e.target.value })}
                />
              </label>
              <label>
                비용계정
                <select
                  value={costEntryDraft.expenseKind}
                  onChange={(e) => onCostEntryDraftChange({ expenseKind: e.target.value })}
                >
                  <option value="">선택</option>
                  {COST_ACCOUNT_OPTIONS.map((opt) => (
                    <option key={opt} value={opt}>
                      {opt}
                    </option>
                  ))}
                </select>
              </label>
              <label>
                업체명
                <input
                  value={costEntryDraft.vendorName}
                  onChange={(e) => onCostEntryDraftChange({ vendorName: e.target.value })}
                />
              </label>
              <label>
                합계금액
                <input
                  value={costEntryDraft.totalAmount}
                  onChange={(e) => onCostEntryDraftChange({ totalAmount: e.target.value })}
                />
              </label>
              <label>
                공급가액
                <input value={costEntryDraft.supplyAmount} disabled />
              </label>
              <label>
                부가세
                <input value={costEntryDraft.vat} disabled />
              </label>
              <label>
                과세여부
                <select
                  value={costEntryDraft.taxMode}
                  onChange={(e) => onCostEntryDraftChange({ taxMode: e.target.value })}
                >
                  <option value="과세">과세</option>
                  <option value="비과세">비과세</option>
                </select>
              </label>
              <label>
                결제여부
                <select
                  value={costEntryDraft.payStatus}
                  onChange={(e) => onCostEntryDraftChange({ payStatus: e.target.value })}
                >
                  <option value="결제">결제</option>
                  <option value="미결제">미결제</option>
                </select>
              </label>
              <label className="cost-form-full">
                메모
                <input
                  value={costEntryDraft.memo}
                  onChange={(e) => onCostEntryDraftChange({ memo: e.target.value })}
                />
              </label>
            </div>
            <div className="modal-actions">
              {costEntryEditingId && (
                <button type="button" className="btn-danger" disabled={costEntryBusy} onClick={() => void onCostEntryDelete()}>
                  삭제
                </button>
              )}
              <button type="button" className="btn-save" disabled={costEntryBusy} onClick={() => void onCostEntrySave()}>
                저장
              </button>
              <button type="button" className="btn-secondary" disabled={costEntryBusy} onClick={closeCostEntryModal}>
                취소
              </button>
            </div>
            {costEntryBusy && (
              <>
                <p className="modal-progress" role="status" aria-live="polite">
                  저장 처리 중…
                </p>
                <div className="modal-progress-bar" aria-hidden={true}>
                  <span className="modal-progress-bar-fill modal-progress-bar-indeterminate" />
                </div>
              </>
            )}
          </div>
        </div>
      )}

      {costModalOpen && activeStore && (
        <div
          className="modal-backdrop"
          role="presentation"
          onMouseDown={(e) => {
            if (e.target === e.currentTarget) closeCostModal();
          }}
        >
          <div
            className="modal-panel"
            role="dialog"
            aria-modal="true"
            aria-labelledby="cost-modal-title"
            onMouseDown={(e) => e.stopPropagation()}
          >
            <div className="modal-header">
              <h3 id="cost-modal-title">비용 등록 업로드</h3>
              <button
                type="button"
                className="modal-close"
                aria-label="닫기"
                disabled={costModalPhase === "reading" || costModalPhase === "saving"}
                onClick={closeCostModal}
              >
                ×
              </button>
            </div>
            <p className="modal-progress" role="status" aria-live="polite">
              {costModalProgress}
            </p>
            {(costModalPhase === "pick" || costModalPhase === "error") && (
              <div className="modal-actions">
                <button type="button" onClick={() => costModalFileInputRef.current?.click()}>
                  파일 선택
                </button>
                <button type="button" className="btn-secondary" onClick={closeCostModal}>
                  취소
                </button>
              </div>
            )}
            {costModalPhase === "reading" && (
              <div className="modal-progress-bar" aria-hidden={true}>
                <span className="modal-progress-bar-fill modal-progress-bar-indeterminate" />
              </div>
            )}
            {costModalPhase === "ready" && (
              <div className="modal-actions">
                <button type="button" className="btn-save" onClick={() => void onCostModalSave()}>
                  저장
                </button>
                <button
                  type="button"
                  className="btn-secondary"
                  onClick={() => {
                    setCostModalPhase("pick");
                    setCostParsedRows(null);
                    setCostModalProgress("파일을 선택해 주세요.");
                  }}
                >
                  다른 파일
                </button>
              </div>
            )}
            {costModalPhase === "saving" && (
              <div className="modal-progress-bar" aria-hidden={true}>
                <span className="modal-progress-bar-fill modal-progress-bar-indeterminate" />
              </div>
            )}
            {costModalPhase === "error" && (
              <p className="error modal-error">파일을 다시 선택하거나 취소할 수 있습니다.</p>
            )}
            <input
              ref={costModalFileInputRef}
              type="file"
              accept=".xls,.xlsx,.csv,application/vnd.ms-excel,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,text/csv"
              className="visually-hidden"
              onChange={(e) => void onCostModalFileChange(e)}
            />
          </div>
        </div>
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
