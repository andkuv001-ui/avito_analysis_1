"use client";

import { useRef, useState } from "react";
import ReactMarkdown from "react-markdown";
import { Download, FileSpreadsheet, Loader2, Upload } from "lucide-react";
import { toast } from "sonner";
import { Button } from "@/components/ui/button";
import {
  Card,
  CardContent,
  CardDescription,
  CardHeader,
  CardTitle,
} from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { Skeleton } from "@/components/ui/skeleton";
import { reportFileSchema } from "@/lib/validation";

interface ReportMeta {
  total_rows: number;
  valid_rows: number;
  skipped_rows: number;
  unique_sellers: number;
}

export function ReportForm() {
  const [file, setFile] = useState<File | null>(null);
  const [loading, setLoading] = useState(false);
  const [report, setReport] = useState<string | null>(null);
  const [meta, setMeta] = useState<ReportMeta | null>(null);
  const inputRef = useRef<HTMLInputElement>(null);

  function handleFileChange(event: React.ChangeEvent<HTMLInputElement>) {
    const next = event.target.files?.[0] ?? null;
    if (!next) {
      setFile(null);
      return;
    }
    const parsed = reportFileSchema.safeParse(next);
    if (!parsed.success) {
      setFile(null);
      if (inputRef.current) inputRef.current.value = "";
      toast.error(parsed.error.errors[0]?.message ?? "Некорректный файл");
      return;
    }
    setFile(next);
  }

  async function handleSubmit(event: React.FormEvent) {
    event.preventDefault();
    if (!file) {
      toast.error("Выберите .xlsx файл");
      return;
    }

    setLoading(true);
    setReport(null);
    setMeta(null);
    const toastId = toast.loading("Анализируем нишу... до 2 минут");

    try {
      const formData = new FormData();
      formData.append("file", file);

      const response = await fetch("/api/report", {
        method: "POST",
        body: formData,
      });
      const data = await response.json().catch(() => null);

      if (!response.ok || !data?.success) {
        throw new Error(data?.error ?? `Ошибка сервера (${response.status})`);
      }

      setReport(data.report);
      setMeta(data.meta ?? null);
      toast.success("Отчёт готов", { id: toastId });
    } catch (error) {
      toast.error(
        error instanceof Error ? error.message : "Не удалось построить отчёт",
        { id: toastId }
      );
    } finally {
      setLoading(false);
    }
  }

  function handleDownload() {
    if (!report) return;
    const blob = new Blob([report], {
      type: "text/markdown;charset=utf-8",
    });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = `avito-niche-report-${new Date().toISOString().slice(0, 10)}.md`;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
  }

  return (
    <div className="mx-auto flex w-full max-w-3xl flex-col gap-6 px-4">
      <div className="flex flex-col gap-2">
        <h1 className="text-3xl font-bold tracking-tight">
          Анализ ниши Авито
        </h1>
        <p className="text-muted-foreground">
          Загрузите .xlsx выгрузку объявлений и получите отчёт о спросе,
          конкуренции, ценах и рисках входа в нишу.
        </p>
      </div>

      <Card>
        <CardHeader>
          <CardTitle className="flex items-center gap-2 text-xl">
            <FileSpreadsheet className="h-5 w-5" aria-hidden />
            Загрузка выгрузки
          </CardTitle>
          <CardDescription>
            Отчёт строится только по агрегированным метрикам файла.
          </CardDescription>
        </CardHeader>
        <CardContent>
          <form onSubmit={handleSubmit} className="flex flex-col gap-4">
            <Input
              ref={inputRef}
              type="file"
              accept=".xlsx"
              onChange={handleFileChange}
              disabled={loading}
            />
            <Button type="submit" disabled={loading || !file}>
              {loading ? (
                <Loader2 className="mr-2 h-4 w-4" aria-hidden />
              ) : (
                <Upload className="mr-2 h-4 w-4" aria-hidden />
              )}
              {loading ? "Анализируем..." : "Построить отчёт"}
            </Button>
            <p className="text-sm text-muted-foreground">
              Только .xlsx, до 3 МБ, до 10 000 строк. Лимит: 5 запросов в час.
            </p>
            {file ? (
              <p className="text-sm">
                Выбран файл: <span className="font-medium">{file.name}</span>
              </p>
            ) : null}
          </form>
        </CardContent>
      </Card>

      {loading ? (
        <div className="flex flex-col gap-3">
          <Skeleton className="h-8 w-2/3" />
          <Skeleton className="h-40 w-full" />
          <Skeleton className="h-40 w-full" />
        </div>
      ) : null}

      {report ? (
        <Card className="animate-in fade-in slide-in-from-bottom-4 duration-500">
          <CardHeader>
            <CardTitle className="text-xl">Отчёт по нише</CardTitle>
            {meta ? (
              <CardDescription>
                Объявлений: {meta.total_rows} · Пригодных: {meta.valid_rows} ·
                Пропущено: {meta.skipped_rows} · Продавцов:{" "}
                {meta.unique_sellers}
              </CardDescription>
            ) : null}
          </CardHeader>
          <CardContent className="flex flex-col gap-4">
            <div className="prose prose-sm max-w-none dark:prose-invert">
              <ReactMarkdown>{report}</ReactMarkdown>
            </div>
            <div>
              <Button type="button" variant="outline" onClick={handleDownload}>
                <Download className="mr-2 h-4 w-4" aria-hidden />
                Скачать .md
              </Button>
            </div>
          </CardContent>
        </Card>
      ) : null}
    </div>
  );
}
