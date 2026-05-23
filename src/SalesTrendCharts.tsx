export type SalesChartSegment = {
  storeId: string;
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

function formatAxisValue(n: number): string {
  if (n >= 100_000_000) return `${(n / 100_000_000).toFixed(1)}억`;
  if (n >= 10000) return `${Math.round(n / 10000)}만`;
  return String(Math.round(n));
}

type StackedProps = {
  title: string;
  months: StackedMonthPoint[];
  formatMoney: (n: number) => string;
};

export function StackedSalesBarChart({ title, months, formatMoney }: StackedProps) {
  if (months.length === 0) {
    return (
      <div className="sales-chart-block">
        <h3 className="sales-chart-title">{title}</h3>
        <p className="muted">표시할 월 데이터가 없습니다.</p>
      </div>
    );
  }

  const storeOrder: { storeId: string; storeName: string }[] = [];
  const seen = new Set<string>();
  for (const month of months) {
    for (const seg of month.segments) {
      if (seen.has(seg.storeId)) continue;
      seen.add(seg.storeId);
      storeOrder.push({ storeId: seg.storeId, storeName: seg.storeName });
    }
  }
  storeOrder.sort((a, b) => a.storeName.localeCompare(b.storeName, "ko"));

  const colorByStoreId = new Map<string, string>();
  storeOrder.forEach((s, i) => colorByStoreId.set(s.storeId, STORE_COLORS[i % STORE_COLORS.length]!));

  const maxTotal = Math.max(...months.map((m) => m.total), 1);
  const innerW = CHART_W - PAD.left - PAD.right;
  const innerH = CHART_H - PAD.top - PAD.bottom;
  const slotW = innerW / months.length;
  const barW = Math.max(12, Math.min(48, slotW * 0.65));
  const yTicks = 4;

  return (
    <div className="sales-chart-block">
      <h3 className="sales-chart-title">{title}</h3>
      <div className="sales-chart-svg-wrap">
        <svg
          className="sales-chart-svg"
          viewBox={`0 0 ${CHART_W} ${CHART_H}`}
          role="img"
          aria-label={title}
        >
          {Array.from({ length: yTicks + 1 }, (_, i) => {
            const ratio = i / yTicks;
            const y = PAD.top + innerH - ratio * innerH;
            const val = Math.round(maxTotal * ratio);
            return (
              <g key={`y-${i}`}>
                <line x1={PAD.left} y1={y} x2={CHART_W - PAD.right} y2={y} className="sales-chart-grid" />
                <text x={PAD.left - 8} y={y + 4} textAnchor="end" className="sales-chart-axis-label">
                  {formatAxisValue(val)}
                </text>
              </g>
            );
          })}
          {months.map((month, mi) => {
            const cx = PAD.left + slotW * mi + slotW / 2;
            const x = cx - barW / 2;
            let yBottom = PAD.top + innerH;
            return (
              <g key={`${month.label}-${mi}`}>
                {month.segments.map((seg) => {
                  const h = seg.sales > 0 ? (seg.sales / maxTotal) * innerH : 0;
                  const y = yBottom - h;
                  const rect = (
                    <rect
                      key={seg.storeId}
                      x={x}
                      y={y}
                      width={barW}
                      height={Math.max(h, 0)}
                      fill={colorByStoreId.get(seg.storeId) ?? "#94a3b8"}
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
      {storeOrder.length > 0 && (
        <ul className="sales-chart-legend">
          {storeOrder.map((s) => (
            <li key={s.storeId}>
              <span className="sales-chart-legend-swatch" style={{ background: colorByStoreId.get(s.storeId) }} />
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
  storeName: string;
  months: SingleMonthPoint[];
  formatMoney: (n: number) => string;
};

export function SingleStoreSalesBarChart({ title, storeName, months, formatMoney }: SingleProps) {
  if (!storeName) {
    return (
      <div className="sales-chart-block">
        <h3 className="sales-chart-title">{title}</h3>
        <p className="muted">매장을 선택하면 그래프를 표시합니다.</p>
      </div>
    );
  }
  if (months.length === 0) {
    return (
      <div className="sales-chart-block">
        <h3 className="sales-chart-title">{title}</h3>
        <p className="muted">표시할 월 데이터가 없습니다.</p>
      </div>
    );
  }

  const maxSales = Math.max(...months.map((m) => m.sales), 1);
  const innerW = CHART_W - PAD.left - PAD.right;
  const innerH = CHART_H - PAD.top - PAD.bottom;
  const slotW = innerW / months.length;
  const barW = Math.max(12, Math.min(48, slotW * 0.65));
  const yTicks = 4;
  const barColor = "#3b82f6";

  return (
    <div className="sales-chart-block">
      <h3 className="sales-chart-title">{title}</h3>
      <p className="muted sales-chart-subtitle">{storeName}</p>
      <div className="sales-chart-svg-wrap">
        <svg className="sales-chart-svg" viewBox={`0 0 ${CHART_W} ${CHART_H}`} role="img" aria-label={title}>
          {Array.from({ length: yTicks + 1 }, (_, i) => {
            const ratio = i / yTicks;
            const y = PAD.top + innerH - ratio * innerH;
            const val = Math.round(maxSales * ratio);
            return (
              <g key={`y-${i}`}>
                <line x1={PAD.left} y1={y} x2={CHART_W - PAD.right} y2={y} className="sales-chart-grid" />
                <text x={PAD.left - 8} y={y + 4} textAnchor="end" className="sales-chart-axis-label">
                  {formatAxisValue(val)}
                </text>
              </g>
            );
          })}
          {months.map((month, mi) => {
            const cx = PAD.left + slotW * mi + slotW / 2;
            const x = cx - barW / 2;
            const h = month.sales > 0 ? (month.sales / maxSales) * innerH : 0;
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
