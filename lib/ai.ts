import type { AggregatedMetrics, RowBrief } from "./excel-parser";
import { DEMAND_LINE_PREFIX, FORBIDDEN_PHRASES } from "./constants";
import { validateReport } from "./validation";

const FORBIDDEN_LIST = FORBIDDEN_PHRASES.map((p) => `«${p}»`).join(", ");

const SECTION_HEADINGS = [
  "#### 1. Данные (факты)",
  "#### 2. Измеримые закономерности",
  "#### 3. Паттерны и конкурентные смыслы",
  "#### 4. Спрос и конверсии",
  "#### 5. Риски входа",
  "#### 6. Стратегии входа",
  "#### 7. Гипотезы и что проверить",
  "#### 8. Условный вывод",
] as const;

export const SYSTEM_PROMPT = `Ты — Senior Data Analyst и стратег по маркетингу на Авито. Твоя задача — по выгрузке объявлений построить экспертный отчёт, где каждое утверждение опирается на данные и помечено уровнем доказательства.

Формат данных: сырой Excel/CSV в LLM не передаётся, файла и путей тоже нет. В сообщении пользователя два блока. Блок «Данные:» — агрегированный JSON: file (total_rows, valid_rows, skipped_rows); parse (parsed_at — время выгрузки в формате ЧЧ:ММ); market (unique_sellers, verified_seller_share, seller_rating_median, seller_reviews_median, sellers_with_multiple_listings_share); demand (views_today_sum — суммарные просмотры сегодня, share_with_views_today — доля объявлений с ненулевыми просмотрами сегодня, значение от 0 до 1, views_total_sum — суммарные просмотры всех объявлений, views_per_day {n, min, p25, median, p75, max, basis: "views_total/max(age_days,1)"} или null — скорость набора просмотров = всего просмотров / возраст объявления в днях с полом в 1 день, age {n, median_days} или null — возраст объявлений в днях); prices (count, min, max, avg, медиана, p25, p75, outliers_high — число позиций дороже p75×5 — и tiers со счётчиками эконом/средний/премиум по сегментам); unclassified_count (строки без цены или с условной ценой); promotion (paid_share, xl_share); positions (avg_position, top10_share); categories; top_titles (20 заголовков с ценами и сегментом); samples (8 объявлений: заголовок, цена, продавец, сегмент, позиция, просмотры, фрагмент описания до 300 символов); text_signals (по каждому ключу: count — число объявлений с упоминанием, share — доля от 0 до 1, examples — примеры заголовков); segment_rules (правило сегментации и примеры заголовков по каждому сегменту). Блок «Строки:» — markdown-таблица всех строк (колонки: заголовок, цена, сегмент, продавец, позиция, просмотры сегодня, всего просмотров, возраст_дн, фрагмент описания), если valid_rows ≤ 200; иначе в этом блоке пометка, что таблица не включена, — тогда анализ веди по счётчикам text_signals, top_titles и samples и обязательно укажи это в разделе 1.

Анализируй строго по этим данным. Запрещено выдумывать факты, цифры, названия брендов, которых нет в данных. Столбцов региона/адреса в JSON нет — регион не выдумывай.

Ниша может быть любой (товарка или услуги). Тип ниши определяй по заголовкам, категориям и структуре цен из данных.

Уровни доказательств (пометка ставится в начале строки-утверждения, каждое бизнес-утверждение обязано её иметь):
- «Факт:» — число или поле из JSON, обязательно с n= или указанием метрики.
- «Наблюдение:» — частотный вывод о текстах, только со счётчиком text_signals или формулировкой «n=X из N» и хотя бы одним примером-заголовком (из rows или samples).
- «Гипотеза:» — предположение вместе со способом проверки.
- «Условный вывод:» — вывод с условиями и рисками.
Непомеченные бизнес-выводы запрещены. «У единиц», «у большинства», «многие» и подобные оценки без числа запрещены. Заявления продавцов — только как «конкуренты заявляют/используют в коммуникации…», не как факт о продукте. Абсолютные утверждения («гарантий нет», «никто не предлагает») запрещены — вместо них «представлено неравномерно, n=X из N».

Сегменты и кластеры:
- Сегменты (per_unit, service, product) бери строго из JSON и segment_rules; тексты правил сегментации выведи в разделе 2; самодельные группы запрещены.
- Кластеры покупателей определяй по явным признакам в данных; их несколько (B2B, B2C, смешанный) — не своди к одному.

Ценовые правила:
- Уровни (эконом/средний/премиум) и n= только из prices[*].tiers: эконом — цена ≤ p25, средний — p25 < цена ≤ p75, премиум — цена > p75. Собственная нарезка цен на уровни запрещена.
- Диапазон цены печатай только вместе с count и медианой, например «750–380400 (n=42, медиана 1400)».
- outliers_high — несопоставимая позиция (оптовая партия, продажа в другой единице измерения, скрытая/фейковая цена или ошибка данных); не выдавай такой максимум за уровень сегмента.
- Широкий разброс цен не равен демпингу: единицы измерения не нормализованы, упоминание демпинга как вывода запрещено — допустима только оговорка о несопоставимости единиц.
- top10_share используй только как «доля объявлений выборки в ТОП-10 на момент выгрузки».
- Если unclassified_count > 0 — добавь строку «Unclassified (цена не указана/условная) (n=X)», где X = unclassified_count. Сумма всех n (ярусы всех сегментов + Unclassified) равна file.valid_rows.

Контекст площадки:
- Просмотры — прокси-метрика, не продажи и не заявки.
- Новые объявления первые 2-3 дня получают пик трафика (не всегда целевого): всплеск просмотров не равен устойчивому спросу.
- Цены в выгрузке якорные: продавцы часто ставят заниженную базовую цену ради трафика, а настоящую называют в диалоге. Аномально низкая цена не равна свободному сегменту: чаще фейк/кликбейт, скрытая часть цены, неликвид или ошибка.

СТРУКТУРА ИТОГОВОГО ОТЧЁТА (Markdown, заголовки дословно, ровно в этом порядке):

#### 1. Данные (факты)
- Объём выборки строго file.valid_rows («Объём выборки: N строк»), рядом total_rows и skipped_rows; время выгрузки parse.parsed_at («Время выгрузки (ЧЧ:ММ): 14:00»).
- Тип ниши (товар/услуга) и что представлено в данных: заголовки, цены, категории, сегменты — только из JSON.
- Если таблица строк не включена (valid_rows > 200) — прямо укажи это здесь.
- Каждая строка этого раздела начинается с пометки «Факт:».

#### 2. Измеримые закономерности
- Закономерности только со счётчиками: цены, позиции, просмотры — числа из JSON с n=.
- Выведи тексты правил сегментации из segment_rules (правило + примеры заголовков).
- Частоты text_signals: счётчик, доля, примеры заголовков (можно из rows).
- Каждая строка-утверждение начинается с пометки: «Факт:» — число из JSON, «Наблюдение:» — частотный вывод со счётчиком «n=X из N» и примером.

#### 3. Паттерны и конкурентные смыслы
- Паттерны заголовков и описаний — только со счётчиком «n=X из N» (text_signals) и хотя бы одним примером-заголовком.
- Смыслы и обещания конкурентов — с пометкой «конкуренты заявляют/используют в коммуникации…», не как факт о продукте.
- Кластеры покупателей — по явным признакам в данных (rows, text_signals, цены); их несколько (B2B, B2C, смешанный), не своди к одному.
- Каждая строка этого раздела начинается с пометки «Наблюдение:» и содержит число; формулировки без чисел запрещены.

#### 4. Спрос и конверсии
- Воронка объявления: просмотры (прокси) → звонки/заявки (этих данных в выгрузке нет — скажи прямо) → сделки. Обещания заявок, звонков или выручки запрещены.
- Рычаги конверсии (заголовок, цена-«от», CTA, скорость ответа, фото) и раздел «что замерить самому» (контрольные объявления, сквозная аналитика) — как стратегии и рекомендации, не как факты.
- Строка «${DEMAND_LINE_PREFIX}» вставляется сервером автоматически — её содержимое не придумывай.

#### 5. Риски входа
- Минимум: фейк-цены/кликбейт-якоря; пик новизны объявлений 2-3 дня; прокси-метрики не равны продажам; сезонность неизвестна (нет данных); сравнение цен без нормализации единиц измерения.
- Каждый риск — с пометкой уровня либо с опорой на число из JSON.

#### 6. Стратегии входа
- Стратегии входа и рычаги (заголовок, цена-«от», CTA, скорость ответа, фото, продвижение) — с пометками уровней доказательств; что замерить самому.
- B2B-заголовок: запрещено слово «Опт» во всех словоформах (опт, опта, оптовая, оптовый, оптом); формула: [Товар/Услуга] + [Сфера применения] + [Главная выгода для бизнеса].
- Маркетинговые клише запрещены, не используй дословно: ${FORBIDDEN_LIST}. Этот запрет распространяется на весь отчёт: любое вхождение заменяй на факт из данных.
- Пары BAD → GOOD (подставляй свои значения из данных, нишевые слова из примера не переноси):
  BAD: «Лучший [Изделие] на рынке — индивидуальный подход и гарантия» → GOOD: «[Товар/Услуга] из [Материал] для [Сфера применения] — [Главная выгода для бизнеса] за [Срок]».
  BAD: «Оптовая закупка [Изделия] — выгодные условия для бизнеса» → GOOD: «[Товар/Услуга] из [Материал] для [Сфера применения] — [Главная выгода для бизнеса] за [Срок]».
  BAD: «Уникальное [Изделие] ручной работы» → GOOD: «[Товар/Услуга] из [Материал] — [Фактура/Характеристика] для [Сфера применения]».

#### 7. Гипотезы и что проверить
- Каждая гипотеза — отдельной строкой, начинающейся с пометки «Гипотеза:», вместе с явным способом проверки: что замерить, на каких контрольных объявлениях, какой сигнал подтвердит или опровергнет.

#### 8. Условный вывод
- Каждая строка этого раздела начинается с пометки «Условный вывод:» и использует только два допустимых формата: «вход возможен при условиях X при рисках Y» либо «данных для решения недостаточно — проверьте Z».
- Безусловные вердикты и размытые оценки интереса запрещены; фразы из списка запрещённых клише (в том числе о «уникальности») не используй — заменяй на условия с числами из данных.

ЖЕСТКИЕ ПРАВИЛА:
1. Если данных для какого-то параметра нет — не выдумывай их. Напиши: «Недостаточно данных для анализа [параметр], предоставьте столбец [название]».
2. Маркетинговые клише (${FORBIDDEN_LIST}) запрещены во всём отчёте; перед выдачей просканируй его — ни одна фраза из этого списка не должна встретиться нигде.
3. Прокси-метрики (просмотры, доли продвижения) никогда не выдавай за продажи, заявки или звонки.
4. Числа — только из JSON, таблицы строк и text_signals; собственных прогнозов дневного трафика не давай — блок спроса сервер собирает сам.
5. Язык — русский, Markdown. Без вступлений вида «В данном отчёте мы рассмотрим…».
6. Утверждения о текстах конкурентов — только со счётчиком и примером; «у единиц / большинство» без числа запрещено.
7. Отчёт лаконичный и структурированный; каждый раздел заполнен по правилам своего уровня доказательств.

Отчёт должен быть самодостаточным инструментом для принятия решений.`;

export const CRITIC_SYSTEM_PROMPT = `Ты — ревизор экспертного отчёта по выгрузке объявлений. На вход ты получаешь блок «Данные:» (агрегированный JSON, включая text_signals и segment_rules), блок «Строки:» (markdown-таблица строк, если включена) и draft_report (черновик отчёта в Markdown).

Твоя задача — сверить утверждения черновика с данными и исправить противоречия. Правила:
1. Сверяй строки с пометками «Наблюдение», «Гипотеза», «Инсайт» и любые утверждения о текстах конкурентов с text_signals и таблицей строк: противоречие счётчику (например, «у единиц» при count=31) или оценка без числа → переформулируй с корректным «n=X из N» (можно добавить пример-заголовок) либо удали строку.
2. Уровень пометки (Факт / Наблюдение / Гипотеза / Условный вывод) сохраняй: не повышай уровень утверждения без данных.
3. Цифры — только из JSON и таблицы строк; свои числа, проценты и прогнозы не вводи.
4. Новые факты, бренды, нишевые термины, которых нет во входных данных, не добавляй.
5. Всё вне исправляемых строк не переписывай: верни полный отчёт в Markdown дословно, изменив только строки с противоречиями (переформулировал с n= либо удалил).
6. Если в черновике есть запрещённая фраза (например, маркетинговый клише или безусловный вердикт) — замени её в ответе на нейтральную формулировку с условием или числом из данных.

Формат ответа — только полный исправленный отчёт в Markdown, без пояснений и без кодовых блоков, со всеми разделами «#### 1 … #### 8». Если противоречий нет — верни черновик без изменений.

Пример (противоречие → исправление; фрагмент, остальной текст отчёта без изменений), вход:
#### 3. Паттерны и конкурентные смыслы
- Наблюдение: о гарантии пишут лишь единицы продавцов.
text_signals.guarantee_lifespan.count = 31, file.valid_rows = 50.
Ответ (фрагмент):
#### 3. Паттерны и конкурентные смыслы
- Наблюдение: о гарантии пишут 31 из 50 продавцов (text_signals.guarantee_lifespan).`;

export interface GeneratedReport {
  report: string;
  warnings: string[];
}

const REQUEST_TIMEOUT_MS = 90_000;

interface ChatMessage {
  role: "system" | "user";
  content: string;
}

interface AiConfig {
  baseUrl: string;
  apiKey: string;
  model: string;
}

function getConfig(): AiConfig {
  const baseUrl = (
    process.env.ROUTERAI_BASE_URL ?? "https://routerai.ru/api/v1"
  ).replace(/\/+$/, "");
  const apiKey = process.env.ROUTERAI_API_KEY?.trim();
  const model = process.env.ROUTERAI_MODEL ?? "openai/gpt-4o-mini";

  if (!apiKey) {
    throw new Error("Не задан ROUTERAI_API_KEY");
  }
  return { baseUrl, apiKey, model };
}

async function requestCompletion(
  config: AiConfig,
  messages: ChatMessage[],
  temperature: number
): Promise<string> {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), REQUEST_TIMEOUT_MS);

  try {
    const response = await fetch(`${config.baseUrl}/chat/completions`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${config.apiKey}`,
      },
      body: JSON.stringify({
        model: config.model,
        temperature,
        messages,
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
      throw new Error(
        `Превышено время ожидания AI (${REQUEST_TIMEOUT_MS / 1000} секунд).`
      );
    }
    throw error;
  } finally {
    clearTimeout(timer);
  }
}

function escapeCell(v: string | number | null): string {
  return String(v ?? "")
    .replace(/\|/g, "\\|")
    .replace(/\s+/g, " ")
    .trim();
}

function buildRowsTable(rows: RowBrief[]): string {
  const header =
    "| Заголовок | Цена | Сегмент | Продавец | Позиция | Просмотры сегодня | Всего просмотров | Возраст_дн | Фрагмент описания |";
  const separator = "|---|---|---|---|---|---|---|---|---|";
  const body = rows
    .map(
      (r) =>
        `| ${escapeCell(r.title)} | ${r.price} | ${r.segment} | ${escapeCell(
          r.seller
        )} | ${r.position ?? "—"} | ${r.views_today} | ${r.views_total} | ${
          r.age_days ?? "—"
        } | ${escapeCell(r.description_excerpt)} |`
    )
    .join("\n");
  return `${header}\n${separator}\n${body}`;
}

function buildUserMessage(metrics: AggregatedMetrics): string {
  const { rows, ...jsonMetrics } = metrics;
  const rowsBlock =
    rows.length > 0
      ? buildRowsTable(rows)
      : "Таблица строк не включена: valid_rows > 200. Анализируй по счётчикам text_signals, top_titles и samples и отметь это в разделе 1.";
  return `Данные:\n${JSON.stringify(jsonMetrics)}\n\nСтроки:\n${rowsBlock}`;
}

function buildDemandLine(metrics: AggregatedMetrics): string {
  const demand = metrics.demand;
  const pct = Math.round(demand.share_with_views_today * 100);
  const head = `${DEMAND_LINE_PREFIX} просмотров сегодня ${demand.views_today_sum} суммарно (доля объявлений с просмотрами сегодня ${pct}%)`;
  const viewsPerDay = demand.views_per_day;
  if (!viewsPerDay) {
    return `${head}; дневная скорость не оценивается — возраст объявлений неизвестен`;
  }
  return `${head}; скорость набора ${viewsPerDay.p25}–${viewsPerDay.p75} просмотров/день (медиана ${viewsPerDay.median}, n=${viewsPerDay.n}, Всего/возраст)`;
}

const DEMAND_STEM = DEMAND_LINE_PREFIX.replace(/:$/, "");

function applyDemandLine(report: string, metrics: AggregatedMetrics): string {
  const kept: string[] = [];
  let firstAt = -1;
  let prefix = "- ";

  for (const line of report.split("\n")) {
    if (line.includes(DEMAND_STEM)) {
      if (firstAt === -1) {
        firstAt = kept.length;
        prefix = line.match(/^(\s*(?:[-*—]\s*)?)/)?.[1] || "- ";
      }
      continue;
    }
    kept.push(line);
  }

  const canonical = buildDemandLine(metrics);
  if (firstAt !== -1) {
    kept.splice(firstAt, 0, `${prefix}${canonical}`);
    return kept.join("\n");
  }

  const sectionAt = kept.findIndex((line) => line.trim() === SECTION_HEADINGS[3]);
  const anchor = sectionAt === -1 ? kept.length : sectionAt + 1;
  kept.splice(anchor, 0, `- ${canonical}`);
  return kept.join("\n");
}

function hasAllSections(text: string): boolean {
  return SECTION_HEADINGS.every((heading) => text.includes(heading));
}

export async function generateReport(
  metrics: AggregatedMetrics
): Promise<GeneratedReport> {
  const config = getConfig();
  const warnings: string[] = [];
  const userMessage = buildUserMessage(metrics);

  const draft = await requestCompletion(
    config,
    [
      { role: "system", content: SYSTEM_PROMPT },
      { role: "user", content: userMessage },
    ],
    0.3
  );

  let report = draft;
  try {
    const criticReply = await requestCompletion(
      config,
      [
        { role: "system", content: CRITIC_SYSTEM_PROMPT },
        {
          role: "user",
          content: `${userMessage}\n\ndraft_report:\n${draft}`,
        },
      ],
      0.2
    );
    if (!hasAllSections(criticReply)) {
      warnings.push(
        "Self-correction не выполнен: в ответе критика нет обязательных разделов отчёта."
      );
    } else {
      report = criticReply;
    }
  } catch (error) {
    const reason =
      error instanceof Error ? error.message : "неизвестная ошибка";
    warnings.push(`Self-correction не выполнен: ${reason}`);
  }

  report = applyDemandLine(report, metrics);
  warnings.push(...validateReport(report, metrics));
  return { report, warnings };
}
