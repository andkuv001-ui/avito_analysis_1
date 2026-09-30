import { z } from "zod";
import {
  DEMAND_LINE_PREFIX,
  FORBIDDEN_PHRASES,
  HEADLINE_LINE_REGEX,
  MAX_FILE_SIZE,
  OPT_WORD_REGEX,
} from "./constants";
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

const DEMAND_STEM = DEMAND_LINE_PREFIX.replace(/:$/, "");
const DAILY_SPEED_PATTERN = /\d+\s*просмотр\w*\s*(?:в день|\/\s*день)/i;
const MEDIAN_IN_LINE_PATTERN = /медиана\s+(\d+(?:[.,]\d+)?)/i;
const MARKER_LINE_PATTERN = /(?:Наблюдение|Паттерн):/;

export function validateReport(
  report: string,
  metrics: AggregatedMetrics
): string[] {
  const warnings: string[] = [];

  for (const phrase of FORBIDDEN_PHRASES) {
    if (report.toLowerCase().includes(phrase.toLowerCase())) {
      warnings.push(`В отчёте найдена запрещённая фраза: «${phrase}»`);
    }
  }

  const hasOptHeadline = report
    .split("\n")
    .some(
      (line) => HEADLINE_LINE_REGEX.test(line) && OPT_WORD_REGEX.test(line)
    );
  if (hasOptHeadline) {
    warnings.push('В заголовке запрещено слово «опт» (включая словоформы)');
  }

  const demandLines = report
    .split("\n")
    .filter((line) => line.includes(DEMAND_STEM));
  if (demandLines.length !== 1) {
    warnings.push(
      `В отчёте должно быть ровно 1 строка со стемом «${DEMAND_STEM}», найдено: ${demandLines.length}`
    );
  } else {
    const line = demandLines[0];
    const viewsPerDay = metrics.demand.views_per_day;
    if (!viewsPerDay) {
      if (DAILY_SPEED_PATTERN.test(line)) {
        warnings.push(
          `Строка «${DEMAND_STEM}» содержит число дневной скорости при отсутствии данных о возрасте объявлений`
        );
      }
    } else {
      const medianMatch = MEDIAN_IN_LINE_PATTERN.exec(line);
      if (!medianMatch) {
        warnings.push(
          `В строке «${DEMAND_STEM}» не найдена медиана дневной скорости`
        );
      } else {
        const value = Number(medianMatch[1].replace(",", "."));
        if (!Number.isFinite(value)) {
          warnings.push(
            `Медиана дневной скорости в строке «${DEMAND_STEM}» не распознана как число`
          );
        } else if (Math.abs(value - viewsPerDay.median) > 0.011) {
          warnings.push(
            `Медиана дневной скорости в отчёте (${value}) не совпадает с расчётной (${viewsPerDay.median}): сигнал бага сборки строки`
          );
        }
      }
    }
  }

  for (const line of report.split("\n")) {
    if (MARKER_LINE_PATTERN.test(line) && !/\d/.test(line)) {
      warnings.push(
        `Утверждение без числового подтверждения: «${line.trim().slice(0, 100)}»`
      );
    }
  }

  const unclassified = metrics.unclassified_count;
  if (unclassified > 0) {
    const hasLabel = report.includes("Unclassified");
    const hasCount = report.includes(`n=${unclassified}`);
    if (!hasLabel || !hasCount) {
      warnings.push(
        `В отчёте не указана строка Unclassified (n=${unclassified})`
      );
    }
  }

  return warnings;
}
