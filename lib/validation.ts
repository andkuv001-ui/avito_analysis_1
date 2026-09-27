import { z } from "zod";
import { MAX_FILE_SIZE } from "./constants";

export const reportFileSchema = z
  .instanceof(File)
  .refine((f) => f.size > 0, "Файл пустой")
  .refine((f) => f.size <= MAX_FILE_SIZE, "Максимальный размер файла 3 МБ")
  .refine(
    (f) => f.name.toLowerCase().endsWith(".xlsx"),
    "Разрешены только .xlsx файлы"
  );

export type ReportFile = z.infer<typeof reportFileSchema>;
