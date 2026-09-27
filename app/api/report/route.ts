import { NextRequest, NextResponse } from "next/server";
import { generateReport } from "@/lib/ai";
import { parseAndAggregate } from "@/lib/excel-parser";
import { checkRateLimit } from "@/lib/rate-limit";
import { reportFileSchema } from "@/lib/validation";

export const runtime = "nodejs";
export const maxDuration = 120;

export async function POST(req: NextRequest) {
  const ip =
    req.headers.get("x-forwarded-for")?.split(",")[0]?.trim() || "unknown";

  if (!checkRateLimit(ip)) {
    return NextResponse.json(
      { error: "Слишком много запросов. Лимит: 5 в час." },
      { status: 429 }
    );
  }

  let parsedFile: ReturnType<typeof reportFileSchema.safeParse>;
  try {
    const form = await req.formData();
    parsedFile = reportFileSchema.safeParse(form.get("file"));
  } catch {
    return NextResponse.json(
      { error: "Не удалось прочитать запрос. Ожидается FormData с файлом." },
      { status: 400 }
    );
  }

  if (!parsedFile.success) {
    return NextResponse.json(
      { error: parsedFile.error.errors[0]?.message ?? "Некорректный файл" },
      { status: 400 }
    );
  }

  try {
    const buffer = Buffer.from(await parsedFile.data.arrayBuffer());
    const metrics = parseAndAggregate(buffer);
    const report = await generateReport(metrics);

    return NextResponse.json({
      success: true,
      report,
      meta: {
        total_rows: metrics.file.total_rows,
        valid_rows: metrics.file.valid_rows,
        skipped_rows: metrics.file.skipped_rows,
        unique_sellers: metrics.market.unique_sellers,
      },
    });
  } catch (error) {
    const message = error instanceof Error ? error.message : "Ошибка анализа";
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
