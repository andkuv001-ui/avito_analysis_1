import { z } from "zod";
import { MAX_FILE_SIZE } from "./constants";
import type { AggregatedMetrics } from "./excel-parser";

export const reportFileSchema = z
  .instanceof(File)
  .refine((f) => f.size > 0, "Файл пустой")
  .refine((f) => f.size <= MAX_FILE_SIZE, "Максимальный размер файла 3 МБ")
  .refine(
    (f) => f.name.toLowerCase().endsWith(".xlsx"),
    "Разрешены только .xlsx файлы"
  );

export type ReportFile = z.infer<typeof reportFileSchema>;

export const parseTimeSchema = z
  .string()
  .regex(/^([01]\d|2[0-3]):[0-5]\d$/, "Укажите время выгрузки в формате ЧЧ:ММ");

const WHY_BLOCK_PATTERN = /\*\*Почему:\*\*/;
const HYPOTHESIS_BLOCK_PATTERN = /\*\*Гипотеза:\*\*/;

export function validateReport(
  report: string,
  metrics: AggregatedMetrics
): string[] {
  const warnings: string[] = [];
  const lines = report.split("\n");
  const clusters = metrics.clusters ?? [];

  for (const cluster of clusters) {
    const row = lines.find(
      (line) => line.trimStart().startsWith("|") && line.includes(cluster.label)
    );
    if (!row) {
      warnings.push(
        `Кластер «${cluster.label}» в карте спроса без count=${cluster.count}`
      );
      continue;
    }
    if (!new RegExp(`\\b${cluster.count}\\b`).test(row)) {
      warnings.push(
        `Кластер «${cluster.label}» в карте спроса без count=${cluster.count}`
      );
    }
  }

  const whyWithoutNumber = lines.filter(
    (line) => WHY_BLOCK_PATTERN.test(line) && !/\d/.test(line)
  );
  for (const line of whyWithoutNumber) {
    warnings.push(
      `Гипотеза без числового подтверждения: «${line.trim().slice(0, 100)}»`
    );
  }

  const hypothesisCount = lines.filter((line) =>
    HYPOTHESIS_BLOCK_PATTERN.test(line)
  ).length;
  if (hypothesisCount < 5) {
    warnings.push(
      `Менее 5 гипотез в разделе 8: найдено блоков «Гипотеза» — ${hypothesisCount}`
    );
  } else if (hypothesisCount > 15) {
    warnings.push(
      `Более 15 гипотез в разделе 8: найдено блоков «Гипотеза» — ${hypothesisCount}`
    );
  }

  return warnings;
}
