import * as XLSX from "xlsx";
import {
  DATE_COLUMN_CANDIDATES,
  MAX_ROWS,
  MIN_VALID_ROWS,
  PER_UNIT_REGEX,
  REQUIRED_COLUMNS,
  SERVICE_REGEX,
  TEXT_SIGNALS,
} from "./constants";
import { classify, type Segment } from "./segmenter";

export interface PriceTiers {
  economy: number;
  middle: number;
  premium: number;
}

export interface SegmentStats {
  count: number;
  min: number;
  max: number;
  avg: number;
  median: number;
  p25: number;
  p75: number;
  outliers_high: number;
  tiers: PriceTiers;
}

export interface CategoryShare {
  value: string;
  count: number;
}

export interface TopTitle {
  title: string;
  price: number;
  segment: Segment;
}

export interface Sample {
  title: string;
  price: number;
  seller: string;
  segment: Segment;
  position: number | null;
  views_today: number;
  views_total: number;
  description_excerpt: string;
}

export interface RowBrief {
  title: string;
  price: number;
  segment: Segment;
  seller: string;
  position: number | null;
  views_today: number;
  views_total: number;
  age_days: number | null;
  description_excerpt: string;
}

export interface TextSignalStat {
  count: number;
  share: number;
  examples: string[];
}

export interface SegmentRule {
  segment: Segment;
  rule: string;
  examples: string[];
}

export interface ViewsPerDayStats {
  n: number;
  min: number;
  p25: number;
  median: number;
  p75: number;
  max: number;
  basis: "views_total/max(age_days,1)";
}

export interface AgeStats {
  n: number;
  median_days: number;
}

export interface AggregatedMetrics {
  file: {
    total_rows: number;
    valid_rows: number;
    skipped_rows: number;
  };
  parse: {
    parsed_at: string;
    hours_since_midnight: number;
  };
  market: {
    unique_sellers: number;
    verified_seller_share: number;
    seller_rating_median: number | null;
    seller_reviews_median: number | null;
    sellers_with_multiple_listings_share: number;
  };
  demand: {
    views_today_sum: number;
    share_with_views_today: number;
    views_total_sum: number;
    views_per_day: ViewsPerDayStats | null;
    age: AgeStats | null;
  };
  prices: Record<Segment, SegmentStats | null>;
  unclassified_count: number;
  promotion: {
    paid_share: number;
    xl_share: number;
  };
  positions: {
    avg_position: number | null;
    top10_share: number | null;
  };
  categories: {
    category_3: CategoryShare[];
    category_4: CategoryShare[];
  };
  top_titles: TopTitle[];
  samples: Sample[];
  text_signals: Record<string, TextSignalStat>;
  segment_rules: SegmentRule[];
  rows: RowBrief[];
}

type Row = Record<string, unknown>;

interface ValidRow {
  title: string;
  seller: string;
  price: number;
  segment: Segment;
  description_excerpt: string;
  position: number | null;
  views_today: number;
  views_total: number;
  age_days: number | null;
  paid: boolean;
  xl: boolean;
  verified: boolean;
  rating: number | null;
  reviews: number | null;
  category_3: string;
  category_4: string;
}

const cleanDesc = (v: unknown): string =>
  String(v ?? "")
    .replace(/<[^>]*>/g, " ")
    .replace(/\s+/g, " ")
    .trim();

function toNum(v: unknown): number | null {
  const n = parseFloat(
    String(v ?? "")
      .replace(/[^\d.,-]/g, "")
      .replace(",", ".")
  );
  return Number.isFinite(n) ? n : null;
}

function normKey(s: string): string {
  return s.replace(/\s+/g, " ").trim().toLowerCase();
}

function pickColumn(headers: string[], candidates: string[]): string | null {
  const normalized = headers.map((h) => ({ raw: h, n: normKey(h) }));
  for (const candidate of candidates) {
    const c = normKey(candidate);
    const exact = normalized.find((h) => h.n === c);
    if (exact) return exact.raw;
  }
  for (const candidate of candidates) {
    const c = normKey(candidate);
    const partial = normalized.find((h) => h.n.includes(c));
    if (partial) return partial.raw;
  }
  return null;
}

function value(row: Row, col: string | null): string {
  if (!col) return "";
  return String(row[col] ?? "").trim();
}

const EXCEL_EPOCH_OFFSET_MS = 25569 * 86400e3;

function parseDateMs(v: unknown): number | null {
  if (v instanceof Date) {
    const t = v.getTime();
    return Number.isFinite(t) ? t : null;
  }
  if (typeof v === "number" && Number.isFinite(v) && v > 0) {
    return Math.round(v * 86400e3 - EXCEL_EPOCH_OFFSET_MS);
  }
  const s = String(v ?? "").trim();
  if (!s) return null;

  const ru =
    /^(\d{1,2})[./](\d{1,2})[./](\d{4})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?/.exec(
      s
    );
  if (ru) {
    return Date.UTC(
      Number(ru[3]),
      Number(ru[2]) - 1,
      Number(ru[1]),
      Number(ru[4] ?? 0),
      Number(ru[5] ?? 0),
      Number(ru[6] ?? 0)
    );
  }

  const iso =
    /^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?(Z|[+-]\d{2}:?\d{2})?$/.exec(
      s
    );
  if (iso) {
    if (iso[7]) {
      const t = Date.parse(s);
      return Number.isFinite(t) ? t : null;
    }
    return Date.UTC(
      Number(iso[1]),
      Number(iso[2]) - 1,
      Number(iso[3]),
      Number(iso[4] ?? 0),
      Number(iso[5] ?? 0),
      Number(iso[6] ?? 0)
    );
  }

  const t = Date.parse(s);
  return Number.isFinite(t) ? t : null;
}

function buildExportTs(parsedAt: string): number {
  const match = /^([01]\d|2[0-3]):([0-5]\d)$/.exec(parsedAt);
  const hours = match ? Number(match[1]) : 0;
  const minutes = match ? Number(match[2]) : 0;
  const now = new Date();
  return Date.UTC(
    now.getUTCFullYear(),
    now.getUTCMonth(),
    now.getUTCDate(),
    hours,
    minutes
  );
}

function statsOf(sorted: number[]): Omit<ViewsPerDayStats, "basis"> | null {
  if (sorted.length === 0) return null;
  return {
    n: sorted.length,
    min: round2(sorted[0]),
    p25: round2(percentile(sorted, 0.25)),
    median: round2(percentile(sorted, 0.5)),
    p75: round2(percentile(sorted, 0.75)),
    max: round2(sorted[sorted.length - 1]),
  };
}

function percentile(sorted: number[], p: number): number {
  if (sorted.length === 0) return 0;
  const idx = (sorted.length - 1) * p;
  const lo = Math.floor(idx);
  const hi = Math.ceil(idx);
  if (lo === hi) return sorted[lo];
  return sorted[lo] + (sorted[hi] - sorted[lo]) * (idx - lo);
}

function median(values: number[]): number {
  if (values.length === 0) return 0;
  const sorted = [...values].sort((a, b) => a - b);
  return percentile(sorted, 0.5);
}

function round2(n: number): number {
  return Math.round(n * 100) / 100;
}

function ratio(part: number, total: number): number {
  return total === 0 ? 0 : round2(part / total);
}

function topCategories(
  rows: ValidRow[],
  key: "category_3" | "category_4",
  limit: number
): CategoryShare[] {
  const counts = new Map<string, number>();
  for (const row of rows) {
    const v = row[key];
    if (!v) continue;
    counts.set(v, (counts.get(v) ?? 0) + 1);
  }
  return [...counts.entries()]
    .map(([value, count]) => ({ value, count }))
    .sort((a, b) => b.count - a.count || a.value.localeCompare(b.value, "ru"))
    .slice(0, limit);
}

function segmentStats(prices: number[]): SegmentStats | null {
  if (prices.length === 0) return null;
  const sorted = [...prices].sort((a, b) => a - b);
  const avg = sorted.reduce((sum, v) => sum + v, 0) / sorted.length;
  const p25 = percentile(sorted, 0.25);
  const p75 = percentile(sorted, 0.75);
  const outliersHigh = p75 > 0 ? sorted.filter((v) => v > p75 * 5).length : 0;
  const tiers: PriceTiers = { economy: 0, middle: 0, premium: 0 };
  for (const v of sorted) {
    if (v <= p25) tiers.economy += 1;
    else if (v <= p75) tiers.middle += 1;
    else tiers.premium += 1;
  }
  return {
    count: sorted.length,
    min: round2(sorted[0]),
    max: round2(sorted[sorted.length - 1]),
    avg: round2(avg),
    median: round2(percentile(sorted, 0.5)),
    p25: round2(p25),
    p75: round2(p75),
    outliers_high: outliersHigh,
    tiers,
  };
}

function parseTimeToHours(parsedAt: string): number {
  const match = /^([01]\d|2[0-3]):([0-5]\d)$/.exec(parsedAt);
  if (!match) {
    throw new Error("Некорректное время выгрузки, ожидается формат ЧЧ:ММ");
  }
  return Number(match[1]) + Number(match[2]) / 60;
}

export function parseAndAggregate(
  buffer: Buffer,
  parsedAt: string
): AggregatedMetrics {
  const hoursSinceMidnight = parseTimeToHours(parsedAt);
  const exportTs = buildExportTs(parsedAt);
  const workbook = XLSX.read(buffer, { type: "buffer", cellDates: true });
  const sheetName = workbook.SheetNames[0];
  if (!sheetName) throw new Error("Файл пустой или не содержит листов");

  const sheet = workbook.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json<Row>(sheet, { defval: "", raw: true });

  if (rows.length === 0) {
    throw new Error("Файл пустой или не содержит строк");
  }
  if (rows.length > MAX_ROWS) {
    throw new Error(
      `Слишком много строк: ${rows.length}. Максимум ${MAX_ROWS}.`
    );
  }

  const headers = Object.keys(rows[0]);
  const missing = REQUIRED_COLUMNS.filter(
    (required) => !pickColumn(headers, [required])
  );
  if (missing.length > 0) {
    throw new Error(`В файле нет обязательных колонок: ${missing.join(", ")}`);
  }

  const colTitle = pickColumn(headers, ["Заголовок"]);
  const colSeller = pickColumn(headers, ["Название продавца"]);
  const colPrice = pickColumn(headers, ["Цена"]);
  const colDesc = pickColumn(headers, ["Описание"]);
  const colPosition = pickColumn(headers, ["Позиция объявлений"]);
  const colViewsToday = pickColumn(headers, ["Просмотров сегодня"]);
  const colViewsTotal = pickColumn(headers, ["Всего просмотров"]);
  const colPaid = pickColumn(headers, [
    "Платные услуги",
    "Платное продвижение",
  ]);
  const colVerified = pickColumn(headers, ["Документы проверены", "Документы"]);
  const colRating = pickColumn(headers, [
    "Рейтинг продавца",
    "Средний рейтинг",
    "Рейтинг",
  ]);
  const colReviews = pickColumn(headers, [
    "Отзывы продавца",
    "Количество отзывов",
    "Отзывы",
  ]);
  const colCat3 = pickColumn(headers, ["Категория 3", "категория 3"]);
  const colCat4 = pickColumn(headers, ["Категория 4", "категория 4"]);
  const colDate = pickColumn(headers, [...DATE_COLUMN_CANDIDATES]);

  const valid: ValidRow[] = [];
  let skipped = 0;

  for (const row of rows) {
    const title = value(row, colTitle);
    const seller = value(row, colSeller);
    const price = toNum(value(row, colPrice));

    if (!title || !seller || price === null) {
      skipped += 1;
      continue;
    }

    const description = cleanDesc(value(row, colDesc));
    const paidRaw = value(row, colPaid);
    const verifiedRaw = value(row, colVerified);

    let ageDays: number | null = null;
    if (colDate) {
      const publishedTs = parseDateMs(row[colDate]);
      if (publishedTs !== null) {
        const age = (exportTs - publishedTs) / 86400e3;
        if (age > 0 && age <= 3650) ageDays = round2(age);
      }
    }

    valid.push({
      title,
      seller,
      price,
      segment: classify(title, description),
      description_excerpt: description.slice(0, 300),
      position: toNum(value(row, colPosition)),
      views_today: toNum(value(row, colViewsToday)) ?? 0,
      views_total: toNum(value(row, colViewsTotal)) ?? 0,
      age_days: ageDays,
      paid: paidRaw.length > 0,
      xl: paidRaw.toLowerCase().includes("xl"),
      verified: verifiedRaw.toLowerCase().includes("проверен"),
      rating: toNum(value(row, colRating)),
      reviews: toNum(value(row, colReviews)),
      category_3: value(row, colCat3),
      category_4: value(row, colCat4),
    });
  }

  if (valid.length < MIN_VALID_ROWS) {
    throw new Error("Недостаточно данных для анализа");
  }

  const sellers = new Map<string, ValidRow>();
  for (const row of valid) {
    if (!sellers.has(row.seller)) sellers.set(row.seller, row);
  }
  const uniqueSellers = sellers.size;

  const multiListingSellers = new Map<string, number>();
  for (const row of valid) {
    multiListingSellers.set(
      row.seller,
      (multiListingSellers.get(row.seller) ?? 0) + 1
    );
  }
  const sellersWithMultiple = [...multiListingSellers.values()].filter(
    (count) => count > 1
  ).length;

  const verifiedSellers = [...sellers.values()].filter(
    (row) => row.verified
  ).length;

  const ratings = [...sellers.values()]
    .map((row) => row.rating)
    .filter((v): v is number => v !== null);
  const reviews = [...sellers.values()]
    .map((row) => row.reviews)
    .filter((v): v is number => v !== null);

  const viewsToday = valid.map((row) => row.views_today);
  const viewsTotal = valid.map((row) => row.views_total);

  const pricesBySeg: Record<Segment, number[]> = {
    per_unit: [],
    service: [],
    product: [],
  };
  for (const row of valid) {
    if (row.price > 1) pricesBySeg[row.segment].push(row.price);
  }

  const positions = valid
    .map((row) => row.position)
    .filter((v): v is number => v !== null);
  const top10 = positions.filter((p) => p <= 10).length;
  const avgPosition =
    positions.length > 0
      ? round2(
          positions.reduce((sum, v) => sum + v, 0) / positions.length
        )
      : null;

  const ages = valid
    .map((row) => row.age_days)
    .filter((v): v is number => v !== null);
  const minDates = Math.max(3, valid.length / 2);
  const datesOk = ages.length >= minDates;

  let viewsPerDay: ViewsPerDayStats | null = null;
  let ageStats: AgeStats | null = null;
  if (datesOk) {
    const speeds = valid
      .filter((row) => row.age_days !== null)
      .map((row) => row.views_total / Math.max(row.age_days as number, 1))
      .sort((a, b) => a - b);
    const speedStats = statsOf(speeds);
    if (speedStats) {
      viewsPerDay = { ...speedStats, basis: "views_total/max(age_days,1)" };
    }
    ageStats = { n: ages.length, median_days: round2(median(ages)) };
  }

  const statsBySeg: Record<Segment, SegmentStats | null> = {
    per_unit: null,
    service: null,
    product: null,
  };
  statsBySeg.per_unit = segmentStats(pricesBySeg.per_unit);
  statsBySeg.service = segmentStats(pricesBySeg.service);
  statsBySeg.product = segmentStats(pricesBySeg.product);

  const tierSum = (["per_unit", "service", "product"] as const).reduce(
    (sum, seg) => {
      const stats = statsBySeg[seg];
      if (!stats) return sum;
      return (
        sum + stats.tiers.economy + stats.tiers.middle + stats.tiers.premium
      );
    },
    0
  );
  const unclassifiedCount = valid.filter((row) => row.price <= 1).length;
  if (tierSum + unclassifiedCount !== valid.length) {
    throw new Error(
      "Внутренняя ошибка агрегации: ярусы и Unclassified не сходятся к числу строк"
    );
  }

  const topTitles: TopTitle[] = [...valid]
    .sort((a, b) => b.price - a.price)
    .slice(0, 20)
    .map((row) => ({ title: row.title, price: row.price, segment: row.segment }));

  const step = Math.max(1, Math.floor(valid.length / 8));
  const samples: Sample[] = [];
  for (let i = 0; i < valid.length && samples.length < 8; i += step) {
    const row = valid[i];
    samples.push({
      title: row.title,
      price: row.price,
      seller: row.seller,
      segment: row.segment,
      position: row.position,
      views_today: row.views_today,
      views_total: row.views_total,
      description_excerpt: row.description_excerpt,
    });
  }

  const textSignals: Record<string, TextSignalStat> = {};
  for (const signal of TEXT_SIGNALS) {
    let count = 0;
    const examples: string[] = [];
    for (const row of valid) {
      const text = `${row.title} ${row.description_excerpt}`;
      if (!signal.regex.test(text)) continue;
      count += 1;
      if (examples.length < 3 && !examples.includes(row.title)) {
        examples.push(row.title.slice(0, 80));
      }
    }
    textSignals[signal.key] = {
      count,
      share: ratio(count, valid.length),
      examples,
    };
  }

  const SEGMENT_RULE_TEXT: Record<Segment, string> = {
    per_unit: `цена за единицу измерения в тексте: ${PER_UNIT_REGEX}`,
    service: `услуга/работа в тексте: ${SERVICE_REGEX}`,
    product: "не подходит под правила per_unit и service — готовые товары",
  };
  const segmentRules: SegmentRule[] = (
    ["per_unit", "service", "product"] as const
  ).map((segment) => {
    const examples: string[] = [];
    for (const row of valid) {
      if (row.segment !== segment) continue;
      if (examples.length < 3 && !examples.includes(row.title)) {
        examples.push(row.title.slice(0, 80));
      }
      if (examples.length >= 3) break;
    }
    return { segment, rule: SEGMENT_RULE_TEXT[segment], examples };
  });

  const rowsBrief: RowBrief[] =
    valid.length > 200
      ? []
      : valid.map((row) => ({
          title: row.title,
          price: row.price,
          segment: row.segment,
          seller: row.seller,
          position: row.position,
          views_today: row.views_today,
          views_total: row.views_total,
          age_days: row.age_days,
          description_excerpt: row.description_excerpt.slice(0, 200),
        }));

  return {
    file: {
      total_rows: rows.length,
      valid_rows: valid.length,
      skipped_rows: skipped,
    },
    parse: {
      parsed_at: parsedAt,
      hours_since_midnight: round2(hoursSinceMidnight),
    },
    market: {
      unique_sellers: uniqueSellers,
      verified_seller_share: ratio(verifiedSellers, uniqueSellers),
      seller_rating_median: ratings.length > 0 ? round2(median(ratings)) : null,
      seller_reviews_median:
        reviews.length > 0 ? round2(median(reviews)) : null,
      sellers_with_multiple_listings_share: ratio(
        sellersWithMultiple,
        uniqueSellers
      ),
    },
    demand: {
      views_today_sum: viewsToday.reduce((sum, v) => sum + v, 0),
      share_with_views_today: ratio(
        viewsToday.filter((v) => v > 0).length,
        valid.length
      ),
      views_total_sum: viewsTotal.reduce((sum, v) => sum + v, 0),
      views_per_day: viewsPerDay,
      age: ageStats,
    },
    prices: statsBySeg,
    unclassified_count: unclassifiedCount,
    promotion: {
      paid_share: ratio(valid.filter((row) => row.paid).length, valid.length),
      xl_share: ratio(valid.filter((row) => row.xl).length, valid.length),
    },
    positions: {
      avg_position: avgPosition,
      top10_share:
        positions.length > 0 ? ratio(top10, positions.length) : null,
    },
    categories: {
      category_3: topCategories(valid, "category_3", 5),
      category_4: topCategories(valid, "category_4", 5),
    },
    top_titles: topTitles,
    samples,
    text_signals: textSignals,
    segment_rules: segmentRules,
    rows: rowsBrief,
  };
}
