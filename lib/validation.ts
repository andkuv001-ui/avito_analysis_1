import { z } from "zod";
import {
  DAILY_TRAFFIC_PREFIX,
  FORBIDDEN_PHRASES,
  HEADLINE_LINE_REGEX,
  INSUFFICIENT_DAILY_NOTE,
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

const DAILY_TRAFFIC_LINE_REGEX = new RegExp(
  `${DAILY_TRAFFIC_PREFIX}\\s*(\\d+(?:[.,]\\d+)?)\\s*просмотр[а-яё]*`,
  "gi"
);

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

  const trafficLines = [...report.matchAll(DAILY_TRAFFIC_LINE_REGEX)];
  const status = metrics.demand.views_per_day_status;

  if (status === "ok") {
    if (trafficLines.length === 0) {
      warnings.push(
        `В отчёте отсутствует строка «${DAILY_TRAFFIC_PREFIX} X просмотров»`
      );
    } else {
      const cap = metrics.demand.views_per_day_projected_max ?? 0;
      const expected = Math.round(
        metrics.demand.views_per_day_projected_median ?? 0
      );
      for (const match of trafficLines) {
        const value = Number(match[1].replace(",", "."));
        if (!Number.isFinite(value)) {
          warnings.push(
            `Дневной трафик в отчёте не распознан как число: «${match[0]}»`
          );
        } else if (value > cap) {
          warnings.push(
            `Дневной трафик в отчёте (${value}) превышает расчётный максимум (${cap})`
          );
        } else if (value !== expected) {
          warnings.push(
            `Число дневного трафика (${value}) не совпадает с расчётной медианой проекции (${expected}): сигнал бага сборки строки`
          );
        }
      }
    }
  } else {
    if (!report.includes(INSUFFICIENT_DAILY_NOTE)) {
      warnings.push(
        `В отчёте отсутствует фраза «${INSUFFICIENT_DAILY_NOTE}»`
      );
    }
    if (trafficLines.length > 0) {
      warnings.push(
        "В отчёте есть строка дневного трафика при недостаточных данных для прогноза"
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
