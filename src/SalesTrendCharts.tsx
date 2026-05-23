import { useEffect, useRef, useState } from "react";

export type SalesChartSegment = {
  storeKey: string;
  storeName: string;
  sales: number;
};

export type StackedMonthPoint = {
  label: string;
  segments: SalesChartSegment[];
  total: number;
  profitTotal: number;
};

export type SingleMonthPoint = {
  label: string;
  sales: number;
  profit: number;
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

const PROFIT_POSITIVE_COLOR = "#3b82f6";
const PROFIT_NEGATIVE_COLOR = "#ef4444";

/** 누적 막대 그래프와 동일한 매장별 색상 */
export function getStoreChartColor(catalog: StoreCatalogEntry[], storeKey: string): string {
  const idx = catalog.findIndex((s) => s.key === storeKey);
  const colorIndex = idx >= 0 ? idx : 0;
  return STORE_COLORS[colorIndex % STORE_COLORS.length]!;
}

const CHART_H = 260;
const CHART_W_MIN = 320;
const CHART_W_DEFAULT = 720;
const PAD = { top: 16, right: 48, bottom: 52, left: 56 };

const SALES_Y_MAX = 120_000_000;
const SALES_Y_STEPS = 4;

const PROFIT_Y_MIN = -20_000_000;
const PROFIT_Y_MAX = 60_000_000;
const PROFIT_Y_STEPS = 4;
const PROFIT_Y_RANGE = PROFIT_Y_MAX - PROFIT_Y_MIN;

const PROFIT_DOT_OFFSET = 10;
const BAR_VALUE_LABEL_OFFSET = 10;

function useChartWidth() {
  const wrapRef = useRef<HTMLDivElement>(null);
  const [chartW, setChartW] = useState(CHART_W_DEFAULT);

  useEffect(() => {
    const el = wrapRef.current;
    if (!el) return;
    const update = () => {
      const w = el.clientWidth;
      if (w > 0) setChartW(Math.max(CHART_W_MIN, Math.floor(w)));
    };
    update();
    const ro = new ResizeObserver(update);
    ro.observe(el);
    return () => ro.disconnect();
  }, []);

  return { wrapRef, chartW };
}

/** Y축·막대 라벨: 백만 원 단위 (1,000,000원 미만 반올림) */
function formatMillionsLabel(n: number): string {
  if (n === 0) return "0";
  const sign = n < 0 ? "-" : "";
  const millions = Math.round(Math.abs(n) / 1_000_000);
  return `${sign}${millions.toLocaleString("ko-KR")}백만`;
}

function profitToY(profit: number, innerH: number): number {
  const clamped = Math.max(PROFIT_Y_MIN, Math.min(PROFIT_Y_MAX, profit));
  const ratio = (clamped - PROFIT_Y_MIN) / PROFIT_Y_RANGE;
  return PAD.top + innerH * (1 - ratio);
}

function profitDotColor(profit: number): string {
  return profit < 0 ? PROFIT_NEGATIVE_COLOR : PROFIT_POSITIVE_COLOR;
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

function SalesYAxisGrid({ chartW, innerH }: { chartW: number; innerH: number }) {
  const plotRight = chartW - PAD.right;
  return (
    <>
      {Array.from({ length: SALES_Y_STEPS + 1 }, (_, i) => {
        const ratio = i / SALES_Y_STEPS;
        const y = PAD.top + innerH - ratio * innerH;
        const val = Math.round(SALES_Y_MAX * ratio);
        return (
          <g key={`sales-y-${i}`}>
            <line x1={PAD.left} y1={y} x2={plotRight} y2={y} className="sales-chart-grid" />
            <text x={PAD.left - 8} y={y + 4} textAnchor="end" className="sales-chart-label">
              {formatMillionsLabel(val)}
            </text>
          </g>
        );
      })}
    </>
  );
}

function ProfitYAxisGrid({ chartW, innerH }: { chartW: number; innerH: number }) {
  const plotRight = chartW - PAD.right;
  const zeroY = profitToY(0, innerH);
  return (
    <>
      {Array.from({ length: PROFIT_Y_STEPS + 1 }, (_, i) => {
        const ratio = i / PROFIT_Y_STEPS;
        const y = PAD.top + innerH - ratio * innerH;
        const val = Math.round(PROFIT_Y_MIN + PROFIT_Y_RANGE * ratio);
        return (
          <g key={`profit-y-${i}`}>
            <text x={plotRight + 6} y={y + 4} textAnchor="start" className="sales-chart-label">
              {formatMillionsLabel(val)}
            </text>
          </g>
        );
      })}
      <line
        x1={PAD.left}
        y1={zeroY}
        x2={plotRight}
        y2={zeroY}
        className="sales-chart-profit-zero-line"
      />
    </>
  );
}

function ProfitMarker({
  x,
  profit,
  innerH,
  tooltip,
}: {
  x: number;
  profit: number;
  innerH: number;
  tooltip: string;
}) {
  if (profit < PROFIT_Y_MIN) {
    const floorY = profitToY(PROFIT_Y_MIN, innerH);
    const arrowTipY = floorY + 9;
    const labelY = floorY - 6;
    return (
      <g className="sales-chart-profit-below-min">
        <text x={x} y={labelY} textAnchor="middle" className="sales-chart-label sales-chart-label-negative">
          {formatMillionsLabel(profit)}
        </text>
        <polygon
          points={`${x},${arrowTipY} ${x - 6},${floorY} ${x + 6},${floorY}`}
          fill={PROFIT_NEGATIVE_COLOR}
        />
        <title>{tooltip}</title>
      </g>
    );
  }

  const cy = profitToY(profit, innerH);
  return (
    <circle cx={x} cy={cy} r={5} fill={profitDotColor(profit)} className="sales-chart-profit-dot">
      <title>{tooltip}</title>
    </circle>
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
  const { wrapRef, chartW } = useChartWidth();

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
  storeCatalog.forEach((s) => colorByStoreKey.set(s.key, getStoreChartColor(storeCatalog, s.key)));

  const innerW = chartW - PAD.left - PAD.right;
  const innerH = CHART_H - PAD.top - PAD.bottom;
  const slotW = innerW / months.length;
  const barW = Math.max(12, Math.min(48, slotW * 0.65));

  return (
    <div className="sales-chart-block">
      <br />
      <ChartTitleRow title={title} meta={titleMeta} />
      <div ref={wrapRef} className="sales-chart-svg-wrap">
        <svg
          className="sales-chart-svg"
          viewBox={`0 0 ${chartW} ${CHART_H}`}
          preserveAspectRatio="xMidYMid meet"
          role="img"
          aria-label={title}
        >
          <SalesYAxisGrid chartW={chartW} innerH={innerH} />
          <ProfitYAxisGrid chartW={chartW} innerH={innerH} />
          {months.map((month, mi) => {
            const cx = PAD.left + slotW * mi + slotW / 2;
            const x = cx - barW / 2;
            const dotX = cx + barW / 2 + PROFIT_DOT_OFFSET;
            let yBottom = PAD.top + innerH;
            let barTop = yBottom;
            return (
              <g key={`${month.label}-${mi}`}>
                {month.segments.map((seg) => {
                  const h = seg.sales > 0 ? (seg.sales / SALES_Y_MAX) * innerH : 0;
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
                  barTop = y;
                  yBottom = y;
                  return rect;
                })}
                <text
                  x={cx}
                  y={barTop - BAR_VALUE_LABEL_OFFSET}
                  textAnchor="middle"
                  className="sales-chart-label sales-chart-label-strong"
                >
                  {formatMillionsLabel(month.total)}
                </text>
                <ProfitMarker
                  x={dotX}
                  profit={month.profitTotal}
                  innerH={innerH}
                  tooltip={`${month.label} 손익: ${formatMoney(month.profitTotal)}`}
                />
                <text
                  x={cx}
                  y={CHART_H - PAD.bottom + 28}
                  textAnchor="middle"
                  className="sales-chart-label"
                >
                  {month.label.length > 10 ? `${month.label.slice(0, 9)}…` : month.label}
                </text>
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
  barColor: string;
  months: SingleMonthPoint[];
  formatMoney: (n: number) => string;
};

export function SingleStoreSalesBarChart({ title, titleMeta, storeName, barColor, months, formatMoney }: SingleProps) {
  const { wrapRef, chartW } = useChartWidth();

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

  const innerW = chartW - PAD.left - PAD.right;
  const innerH = CHART_H - PAD.top - PAD.bottom;
  const slotW = innerW / months.length;
  const barW = Math.max(12, Math.min(48, slotW * 0.65));

  return (
    <div className="sales-chart-block">
      <br />
      <ChartTitleRow title={title} meta={titleMeta} />
      <div ref={wrapRef} className="sales-chart-svg-wrap">
        <svg
          className="sales-chart-svg"
          viewBox={`0 0 ${chartW} ${CHART_H}`}
          preserveAspectRatio="xMidYMid meet"
          role="img"
          aria-label={title}
        >
          <SalesYAxisGrid chartW={chartW} innerH={innerH} />
          <ProfitYAxisGrid chartW={chartW} innerH={innerH} />
          {months.map((month, mi) => {
            const cx = PAD.left + slotW * mi + slotW / 2;
            const x = cx - barW / 2;
            const dotX = cx + barW / 2 + PROFIT_DOT_OFFSET;
            const h = month.sales > 0 ? (month.sales / SALES_Y_MAX) * innerH : 0;
            const y = PAD.top + innerH - h;
            const barTop = y;
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
                  y={barTop - BAR_VALUE_LABEL_OFFSET}
                  textAnchor="middle"
                  className="sales-chart-label sales-chart-label-strong"
                >
                  {formatMillionsLabel(month.sales)}
                </text>
                <ProfitMarker
                  x={dotX}
                  profit={month.profit}
                  innerH={innerH}
                  tooltip={`${month.label} 손익: ${formatMoney(month.profit)}`}
                />
                <text
                  x={cx}
                  y={CHART_H - PAD.bottom + 28}
                  textAnchor="middle"
                  className="sales-chart-label"
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
