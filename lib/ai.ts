import type { AggregatedMetrics } from "./excel-parser";

export const SYSTEM_PROMPT = `Ты — аналитик розничного рынка. Тебе переданы только агрегированные метрики выгрузки объявлений (одна ниша, один регион). Данные приходят в сообщении пользователя блоком "Данные:" в формате JSON.

Жёсткие правила:
1. Ничего не выдумывай. Все цифры, факты, компании и события в отчёте должны быть либо прямо из переданных метрик, либо очевидно вычислены из них (доля, медиана, разница, диапазон). Если данных для вывода нет — так и напиши.
2. Никогда не выдавай прокси-метрики за продажи, заявки, звонки, выручку или конверсию. Просмотры объявления — не покупки.
3. Допустимые прокси-конверсии (используй только их): доля объявлений с просмотрами сегодня > 0; медиана просмотров сегодня; доля объявлений с платным продвижением; медиана просмотров за всё время (осторожно: накопленный показатель — старое объявление могло набрать показы за месяцы).
4. Язык — русский. Формат — Markdown. Пиши по делу, без воды и без вступлений вида «В данном отчёте мы рассмотрим…».
5. Структура обязательна, заголовки дословно, ровно в этом порядке:
# Отчёт по перспективам ниши
## 1. Краткий вывод
## 2. Оценка спроса
## 3. Конкуренция
## 4. Цены и ценовые сегменты
## 5. Риски входа
## 6. Рекомендации по входу
## 7. Перспективы с учётом прокси-конверсий
6. Сегменты: per_unit — сырьё и материалы с ценой за единицу измерения; service — услуги; product — готовые товары. Если сегмент пуст (null), пропусти его полностью и не упоминай. Подтипы внутри сегментов определяй сам по полю top_titles (заголовки объявлений) — не используй заранее заданные или придуманные нишевые термины, которых нет в данных.
7. В разделе 3 честно оценивай конкуренцию по числу продавцов, доле продавцов с несколькими объявлениями, позициям и доле платного продвижения.
8. В разделах 5 и 6 называй риски и даешь рекомендации только в пределах того, что видно в метриках.
9. Отчёт должен быть самодостаточным: читатель не видит сырые данные.`;

export async function generateReport(
  metrics: AggregatedMetrics
): Promise<string> {
  const baseUrl = (process.env.ROUTERAI_BASE_URL ?? "https://routerai.ru/api/v1").replace(
    /\/+$/,
    ""
  );
  const apiKey = process.env.ROUTERAI_API_KEY;
  const model = process.env.ROUTERAI_MODEL ?? "openai/gpt-4o-mini";

  if (!apiKey) {
    throw new Error("Не задан ROUTERAI_API_KEY");
  }

  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), 120_000);

  try {
    const response = await fetch(`${baseUrl}/chat/completions`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${apiKey}`,
      },
      body: JSON.stringify({
        model,
        temperature: 0.3,
        messages: [
          { role: "system", content: SYSTEM_PROMPT },
          {
            role: "user",
            content: `Данные:\n${JSON.stringify(metrics)}`,
          },
        ],
      }),
      signal: controller.signal,
      cache: "no-store",
    });

    if (!response.ok) {
      const detail = await response.text().catch(() => "");
      throw new Error(`AI API ${response.status}: ${detail.slice(0, 500)}`);
    }

    const data = (await response.json()) as {
      choices?: Array<{ message?: { content?: string } }>;
    };
    const content = data.choices?.[0]?.message?.content ?? "";
    if (!content.trim()) {
      throw new Error("AI вернул пустой ответ.");
    }
    return content.trim();
  } catch (error) {
    if (error instanceof Error && error.name === "AbortError") {
      throw new Error("Превышено время ожидания AI (120 секунд).");
    }
    throw error;
  } finally {
    clearTimeout(timer);
  }
}
