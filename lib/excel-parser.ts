import * as XLSX from "xlsx";
import {
  MAX_ROWS,
  MIN_VALID_ROWS,
  REQUIRED_COLUMNS,
} from "./constants";
import { classify, type Segment } from "./segmenter";

export interface SegmentStats {
  count: number;
  min: number;
  max: number;
  median: number;
  p25: number;
  p75: number;
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

export interface AggregatedMetrics {
  file: {
    total_rows: number;
    valid_rows: number;
    skipped_rows: number;
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
    views_today_median: number;
    share_with_views_today: number;
    views_total_median: number;
  };
  prices: Record<Segment, SegmentStats | null>;
  promotion: {
    paid_share: number;
    xl_share: number;
  };
  positions: {
    median_position: number | null;
    top10_share: number | null;
  };
  categories: {
    category_3: CategoryShare[];
    category_4: CategoryShare[];
  };
  top_titles: TopTitle[];
  samples: Sample[];
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
  return {
    count: sorted.length,
    min: round2(sorted[0]),
    max: round2(sorted[sorted.length - 1]),
    median: round2(percentile(sorted, 0.5)),
    p25: round2(percentile(sorted, 0.25)),
    p75: round2(percentile(sorted, 0.75)),
  };
}

export function parseAndAggregate(buffer: Buffer): AggregatedMetrics {
  const workbook = XLSX.read(buffer, { type: "buffer" });
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

    valid.push({
      title,
      seller,
      price,
      segment: classify(title, description),
      description_excerpt: description.slice(0, 300),
      position: toNum(value(row, colPosition)),
      views_today: toNum(value(row, colViewsToday)) ?? 0,
      views_total: toNum(value(row, colViewsTotal)) ?? 0,
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

  return {
    file: {
      total_rows: rows.length,
      valid_rows: valid.length,
      skipped_rows: skipped,
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
      views_today_median: round2(median(viewsToday)),
      share_with_views_today: ratio(
        viewsToday.filter((v) => v > 0).length,
        valid.length
      ),
      views_total_median: round2(median(viewsTotal)),
    },
    prices: {
      per_unit: segmentStats(pricesBySeg.per_unit),
      service: segmentStats(pricesBySeg.service),
      product: segmentStats(pricesBySeg.product),
    },
    promotion: {
      paid_share: ratio(valid.filter((row) => row.paid).length, valid.length),
      xl_share: ratio(valid.filter((row) => row.xl).length, valid.length),
    },
    positions: {
      median_position: positions.length > 0 ? round2(median(positions)) : null,
      top10_share:
        positions.length > 0 ? ratio(top10, positions.length) : null,
    },
    categories: {
      category_3: topCategories(valid, "category_3", 5),
      category_4: topCategories(valid, "category_4", 5),
    },
    top_titles: topTitles,
    samples,
  };
}
