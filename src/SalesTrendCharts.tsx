export type SalesChartSegment = {
  storeKey: string;
  storeName: string;
  sales: number;
};

export type StackedMonthPoint = {
  label: string;
  segments: SalesChartSegment[];
  total: number;
};

export type SingleMonthPoint = {
  label: string;
  sales: number;
};

export type StoreCatalogEntry = {
  key: string;
  storeName: string;
};

const STORE_COLORS = [
  "#3b82f6",
  "#10b981",
  "#f59e0b",
  "#ef4444",
  "#8b5cf6",
  "#06b6d4",
  "#ec4899",
  "#84cc16",
  "#f97316",
  "#6366f1",
  "#14b8a6",
  "#a855f7",
];

const CHART_W = 720;
const CHART_H = 260;
const PAD = { top: 16, right: 16, bottom: 52, left: 56 };

const STACKED_Y_MAX = 150_000_000;
const STACKED_Y_STEPS = 5;
const SINGLE_Y_MAX = 100_000_000;
const SINGLE_Y_STEPS = 4;

function formatAxisValue(n: number): string {
  if (n >= 100_000_000) return `${(n / 100_000_000).toFixed(1)}억`;
  if (n >= 10000) return `${Math.round(n / 10000)}만`;
  return String(Math.round(n));
}

function ChartTitleRow({ title, meta }: { title: string; meta?: string }) {
  return (
    <div className="sales-chart-heading">
      <h3 className="sales-chart-title">{title}</h3>
      {meta ? (
        <>
          <span className="panel-heading-spacer" aria-hidden={true}>
            {"\u00A0\u00A0"}
          </span>
          <span className="muted card-meta">{meta}</span>
        </>
      ) : null}
    </div>
  );
}

function YAxisGrid({ yMax, ySteps }: { yMax: number; ySteps: number }) {
  const innerH = CHART_H - PAD.top - PAD.bottom;
  return (
    <>
      {Array.from({ length: ySteps + 1 }, (_, i) => {
        const ratio = i / ySteps;
        const y = PAD.top + innerH - ratio * innerH;
        const val = Math.round(yMax * ratio);
        return (
          <g key={`y-${i}`}>
            <line x1={PAD.left} y1={y} x2={CHART_W - PAD.right} y2={y} className="sales-chart-grid" />
            <text x={PAD.left - 8} y={y + 4} textAnchor="end" className="sales-chart-axis-label">
              {formatAxisValue(val)}
            </text>
          </g>
        );
      })}
    </>
  );
}

type StackedProps = {
  title: string;
  titleMeta?: string;
  storeCatalog: StoreCatalogEntry[];
  months: StackedMonthPoint[];
  formatMoney: (n: number) => string;
};

export function StackedSalesBarChart({ title, titleMeta, storeCatalog, months, formatMoney }: StackedProps) {
  if (months.length === 0) {
    return (
      <div className="sales-chart-block">
        <br />
        <ChartTitleRow title={title} meta={titleMeta} />
        <p className="muted">표시할 월 데이터가 없습니다.</p>
      </div>
    );
  }

  const colorByStoreKey = new Map<string, string>();
  storeCatalog.forEach((s, i) => colorByStoreKey.set(s.key, STORE_COLORS[i % STORE_COLORS.length]!));

  const innerW = CHART_W - PAD.left - PAD.right;
  const innerH = CHART_H - PAD.top - PAD.bottom;
  const slotW = innerW / months.length;
  const barW = Math.max(12, Math.min(48, slotW * 0.65));

  return (
    <div className="sales-chart-block">
      <br />
      <ChartTitleRow title={title} meta={titleMeta} />
      <div className="sales-chart-svg-wrap">
        <svg
          className="sales-chart-svg"
          viewBox={`0 0 ${CHART_W} ${CHART_H}`}
          role="img"
          aria-label={title}
        >
          <YAxisGrid yMax={STACKED_Y_MAX} ySteps={STACKED_Y_STEPS} />
          {months.map((month, mi) => {
            const cx = PAD.left + slotW * mi + slotW / 2;
            const x = cx - barW / 2;
            let yBottom = PAD.top + innerH;
            return (
              <g key={`${month.label}-${mi}`}>
                {month.segments.map((seg) => {
                  const h = seg.sales > 0 ? (seg.sales / STACKED_Y_MAX) * innerH : 0;
                  const y = yBottom - h;
                  const rect = (
                    <rect
                      key={seg.storeKey}
                      x={x}
                      y={y}
                      width={barW}
                      height={Math.max(h, 0)}
                      fill={colorByStoreKey.get(seg.storeKey) ?? "#94a3b8"}
                      className="sales-chart-bar-seg"
                    >
                      <title>
                        {month.label} · {seg.storeName}: {formatMoney(seg.sales)}
                      </title>
                    </rect>
                  );
                  yBottom = y;
                  return rect;
                })}
                <text
                  x={cx}
                  y={CHART_H - PAD.bottom + 28}
                  textAnchor="middle"
                  className="sales-chart-month-label"
                >
                  {month.label.length > 10 ? `${month.label.slice(0, 9)}…` : month.label}
                </text>
                <title>
                  {month.label} 합계: {formatMoney(month.total)}
                </title>
              </g>
            );
          })}
        </svg>
      </div>
      {storeCatalog.length > 0 && (
        <ul className="sales-chart-legend">
          {storeCatalog.map((s) => (
            <li key={s.key}>
              <span className="sales-chart-legend-swatch" style={{ background: colorByStoreKey.get(s.key) }} />
              {s.storeName}
            </li>
          ))}
        </ul>
      )}
    </div>
  );
}

type SingleProps = {
  title: string;
  titleMeta?: string;
  storeName: string;
  months: SingleMonthPoint[];
  formatMoney: (n: number) => string;
};

export function SingleStoreSalesBarChart({ title, titleMeta, storeName, months, formatMoney }: SingleProps) {
  if (!storeName) {
    return (
      <div className="sales-chart-block">
        <br />
        <ChartTitleRow title={title} meta={titleMeta} />
        <p className="muted">매장을 선택하면 그래프를 표시합니다.</p>
      </div>
    );
  }
  if (months.length === 0) {
    return (
      <div className="sales-chart-block">
        <br />
        <ChartTitleRow title={title} meta={titleMeta} />
        <p className="muted">표시할 월 데이터가 없습니다.</p>
      </div>
    );
  }

  const innerW = CHART_W - PAD.left - PAD.right;
  const innerH = CHART_H - PAD.top - PAD.bottom;
  const slotW = innerW / months.length;
  const barW = Math.max(12, Math.min(48, slotW * 0.65));
  const barColor = "#3b82f6";

  return (
    <div className="sales-chart-block">
      <br />
      <ChartTitleRow title={title} meta={titleMeta} />
      <div className="sales-chart-svg-wrap">
        <svg className="sales-chart-svg" viewBox={`0 0 ${CHART_W} ${CHART_H}`} role="img" aria-label={title}>
          <YAxisGrid yMax={SINGLE_Y_MAX} ySteps={SINGLE_Y_STEPS} />
          {months.map((month, mi) => {
            const cx = PAD.left + slotW * mi + slotW / 2;
            const x = cx - barW / 2;
            const h = month.sales > 0 ? (month.sales / SINGLE_Y_MAX) * innerH : 0;
            const y = PAD.top + innerH - h;
            return (
              <g key={`${month.label}-${mi}`}>
                <rect
                  x={x}
                  y={y}
                  width={barW}
                  height={Math.max(h, 0)}
                  fill={barColor}
                  className="sales-chart-bar-seg"
                >
                  <title>
                    {month.label}: {formatMoney(month.sales)}
                  </title>
                </rect>
                <text
                  x={cx}
                  y={CHART_H - PAD.bottom + 28}
                  textAnchor="middle"
                  className="sales-chart-month-label"
                >
                  {month.label.length > 10 ? `${month.label.slice(0, 9)}…` : month.label}
                </text>
              </g>
            );
          })}
        </svg>
      </div>
    </div>
  );
}
