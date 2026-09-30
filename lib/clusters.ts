import { z } from "zod";

export interface ClusterKey {
  key: string;
  label: string;
  keywords: string[];
}

export interface ClusterStat {
  key: string;
  label: string;
  count: number;
  share: number;
  examples: string[];
}

export interface ClusterSourceRow {
  title: string;
  description: string;
}

export const MAX_CLUSTERS = 12;
const MAX_KEYWORDS = 5;
const KEYWORD_MIN = 3;
const KEYWORD_MAX = 40;
const EXAMPLE_MAX = 80;

const clusterSchema = z.object({
  key: z
    .string()
    .trim()
    .min(1)
    .max(64)
    .regex(/^[a-z0-9_-]+$/i, "ключ должен быть латиницей"),
  label: z
    .string()
    .trim()
    .min(1)
    .max(80)
    .refine((s) => !/[\n\r`|*#]/.test(s), "метка содержит служебные символы"),
  keywords: z
    .array(
      z
        .string()
        .trim()
        .min(KEYWORD_MIN, "keyword короче 3 символов")
        .max(KEYWORD_MAX)
    )
    .min(1)
    .max(MAX_KEYWORDS),
});

const proposalSchema = z.array(clusterSchema).max(MAX_CLUSTERS);

function asRawString(v: unknown): string {
  return typeof v === "string" ? v.trim() : "";
}

function asCleanString(v: unknown): string {
  return typeof v === "string" ? v.replace(/\s+/g, " ").trim() : "";
}

export function sanitizeClusters(raw: unknown): ClusterKey[] {
  if (!Array.isArray(raw)) {
    throw new Error("ответ proposal не является массивом");
  }
  if (raw.length === 0) return [];

  const cleaned: ClusterKey[] = [];
  const seenKeys = new Set<string>();
  const seenLabels = new Set<string>();

  for (const item of raw) {
    if (cleaned.length >= MAX_CLUSTERS) break;
    if (typeof item !== "object" || item === null) continue;
    const record = item as Record<string, unknown>;

    const key = asRawString(record.key).toLowerCase();
    const label = asRawString(record.label);
    if (!key || !label) continue;
    if (seenKeys.has(key) || seenLabels.has(label)) continue;

    const keywords: string[] = [];
    const rawKeywords = Array.isArray(record.keywords) ? record.keywords : [];
    for (const keyword of rawKeywords) {
      const value = asCleanString(keyword);
      if (value.length < KEYWORD_MIN || value.length > KEYWORD_MAX) continue;
      if (keywords.length >= MAX_KEYWORDS) break;
      if (!keywords.some((k) => k.toLowerCase() === value.toLowerCase())) {
        keywords.push(value);
      }
    }
    if (keywords.length === 0) continue;

    const parsed = clusterSchema.safeParse({ key, label, keywords });
    if (!parsed.success) continue;

    seenKeys.add(parsed.data.key);
    seenLabels.add(parsed.data.label);
    cleaned.push(parsed.data);
  }

  if (cleaned.length === 0) {
    throw new Error("все кластеры отброшены (мусор в ответе proposal)");
  }

  const validated = proposalSchema.safeParse(cleaned);
  if (!validated.success) {
    throw new Error("ответ proposal не прошёл валидацию схемы");
  }
  return validated.data;
}

function clusterFragment(text: string, keyword: string): string | null {
  const index = text.toLowerCase().indexOf(keyword.toLowerCase());
  if (index === -1) return null;
  const start = Math.max(0, index - 25);
  const end = Math.min(text.length, start + EXAMPLE_MAX);
  const fragment = text.slice(start, end).replace(/\s+/g, " ").trim();
  if (!fragment.toLowerCase().includes(keyword.toLowerCase())) return null;
  return `${start > 0 ? "…" : ""}${fragment}${end < text.length ? "…" : ""}`;
}

export function countClusters(
  source: ClusterSourceRow[],
  clusters: ClusterKey[]
): ClusterStat[] {
  const total = source.length;
  if (total === 0) return [];

  const stats: ClusterStat[] = [];
  for (const cluster of clusters) {
    let count = 0;
    const examples: string[] = [];
    for (const row of source) {
      const text = `${row.title} ${row.description}`;
      const lower = text.toLowerCase();
      if (!cluster.keywords.some((keyword) => lower.includes(keyword.toLowerCase()))) {
        continue;
      }
      count += 1;
      if (examples.length >= 3) continue;
      for (const keyword of cluster.keywords) {
        const fragment = clusterFragment(text, keyword);
        if (fragment && !examples.includes(fragment)) {
          examples.push(fragment);
          break;
        }
      }
    }
    if (count === 0) continue;
    if (count > total) {
      throw new Error(
        `Внутренняя ошибка кластеризации: count=${count} больше числа строк ${total}`
      );
    }
    stats.push({
      key: cluster.key,
      label: cluster.label,
      count,
      share: Math.round((count / total) * 100) / 100,
      examples,
    });
  }
  return stats;
}
