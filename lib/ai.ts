import type { AggregatedMetrics, RowBrief } from "./excel-parser";
import { DEMAND_MAP_TABLE_HEADER, REPORT_HEADINGS } from "./constants";
import {
  countClusters,
  sanitizeClusters,
  type ClusterSourceRow,
} from "./clusters";
import { validateReport } from "./validation";

export const PROPOSAL_PROMPT = `Ты аналитик поисковых запросов. Перед тобой список записей «заголовок | описание». Выдели от 1 до 12 поисковых кластеров — тематических групп запросов покупателей, которые реально видны в текстах.

Правила:
1. Ответ — только JSON-массив вида [{"key":"...","label":"...","keywords":["..."]}], без пояснений, без markdown и без кодового блока.
2. key — латиница и цифры, без пробелов, уникален (например "scenario_office").
3. label — короткая метка по-русски (до 80 символов), одна строка, без markdown-символов (*, #, |).
4. keywords — от 1 до 5 подстрок длиной 3–40 символов, которые дословно встречаются в текстах записей (регистр не важен). Только содержательные слова и словосочетания; служебные частицы («из», «для», «и», «в») не используй.
5. Не придумывай кластеры, которых нет в текстах. Если тексты однородны — верни меньше кластеров.

Пример формата (плейсхолдеры, содержимое подставь своё по данным):
[{"key":"product","label":"[Товар]","keywords":["[Товар]"]},{"key":"material","label":"[Материал]","keywords":["[Материал]"]},{"key":"place","label":"[Помещение]","keywords":["[Помещение]"]}]`;

export const REPORT_SYSTEM_PROMPT = `Ты — аналитик продвижения на Avito. Проанализируй предоставленную поисковую выдачу по нише **[НИША / ТОВАР]**, регион **[РЕГИОН]**.

Плейсхолдеры подставь сам: нишу определи по данным выгрузки (заголовки, категории, цены) и назови её в разделе 1; регион бери из блока «Контекст:», а если его там нет — напиши «регион не указан в выгрузке» и не выдумывай его.

Твоя задача — не переписывать объявления, а определить структуру коммерческого спроса и рекламные возможности. Главный принцип: не «как написать объявление», а «какой коммерческий сценарий спроса стоит занять».

Проанализируй объявления и выдели:
1. Поисковые кластеры
* основные товарные запросы;
* характеристики;
* сценарии покупки;
* проблемы/задачи клиента;
* B2C и B2B-запросы;
* явно коммерческие запросы.
2. Ситуации покупки. Для каких конкретных ситуаций люди ищут товар: замена; ремонт; переезд; запуск бизнеса; нестандартный размер; срочная покупка; экономия; решение специфической задачи и т.д.
3. Предложения конкурентов. Определи: какие товары и услуги предлагают; какие УТП используют; какие цены; какие аргументы повторяются; какие поисковые сценарии уже заняты большинством; какие сегменты представлены слабо.
4. Незакрытые возможности. Найди: коммерческие сценарии с низкой конкуренцией; запросы/кластеры, которые конкуренты используют редко; проблемы клиентов, которые почти не раскрываются; потенциальные B2B-сегменты; отличающиеся причины открыть объявление.
5. Карта спроса — таблица со строкой-заголовком: «${DEMAND_MAP_TABLE_HEADER}».
6. Рекламные гипотезы. Для каждой перспективной возможности сформулируй три строки подряд:
**Гипотеза:** Если сделать объявление для [сегмент/ситуация] с оффером [решение], то можно получить дополнительный целевой спрос.
**Почему:** конкретное наблюдение из выдачи, с числом из данных.
**Что тестировать:** сценарий / проблема / оффер / ключевой запрос.

ВАЖНО:
* Не оценивай качество объявления только по его тексту.
* Не считай количество объявлений доказательством объёма спроса.
* Не путай частотность запроса с конверсией.
* Не придумывай данные, которых нет в выдаче.
* Отделяй факты от своих гипотез.
* Не занимайся искусственной уникализацией текстов.
* Ищи разные причины, по которым клиент может открыть объявление.

Техническая часть — как читать данные и оформить ответ:

Формат данных: сырой Excel/CSV в LLM не передаётся, файла и путей тоже нет. В сообщении пользователя до четырёх блоков:
- «Данные:» — агрегированный JSON: file (total_rows, valid_rows, skipped_rows); parse (parsed_at — время выгрузки в формате ЧЧ:ММ); market (unique_sellers, verified_seller_share, seller_rating_median, seller_reviews_median, sellers_with_multiple_listings_share); demand (views_today_sum, views_today_median, share_with_views_today, views_total_sum, views_total_median, views_per_day {n, min, p25, median, p75, max, basis} или null, age {n, median_days} или null); quadrants (threshold_views_today и threshold_views_total — пороги = медианы выборки; ячейки strong_stable / stable_weak_now / rising_fast / new_or_weak с count и examples); prices (count, min, max, avg, медиана, p25, p75, outliers_high и tiers со счётчиками эконом/средний/премиум по сегментам); unclassified_count; promotion (paid_share, xl_share); forecast (views_today_avg — просмотров сегодня в среднем по выгрузке, views_today_avg_paid — в среднем среди объявлений с платными сервисами, views_today_avg_organic — в среднем среди объявлений без платных, views_today_min и views_today_max — минимум и максимум просмотров сегодня, paid_listings и organic_listings — сколько объявлений с платными сервисами и сколько без; age_2_days и age_5_days — статистика свежести: {listings, views_today_avg, views_today_min, views_today_max} для объявлений, опубликованных ровно 2 и ровно 5 дней назад по колонке «Дата публикации»); positions (top_count, top_unique_sellers, top_repeat_share); categories; top_titles; samples; text_signals (count, share, examples); segment_rules.
- «Контекст:» — включается только при наличии опциональных колонок выгрузки: строки «- Регион: …» и/или «- Запрос: …». Если блока нет — региона и запроса в данных нет.
- «Кластеры:» — серверный подсчёт кластеров (метка, count, share, примеры-фрагменты).
- «Строки:» — markdown-таблица строк; при valid_rows > 200 её нет.

Кластеры:
- Кластеры, их count, share и examples заданы только серверным блоком «Кластеры:». Самодельные кластеры, собственные названия кластеров и собственные счётчики запрещены.
- Пересечения кластеров допустимы: сумма count может быть больше или меньше N — это нормально, не выравнивай её.
- Если блок «Кластеры:» сообщает о пустом списке — в разделе 2 напиши честно, что кластерный анализ не дал результатов по этой выгрузке, и не выдумывай кластеры.

Анализируй строго по этим данным: цифры, факты, бренды и ассортимент — только из выгрузки; обещания продавцов передавай как «конкуренты заявляют…», а не как факт о продукте. Отчёт универсален: товарные позиции и услуги анализируй одинаково.

Контекст площадки (внешние закономерности, не факты этой выгрузки):
- Просмотры — прокси-метрика: они не равны продажам, звонкам и заявкам; данных о звонках и сделках в выгрузке нет.
- Первые 2–3 дня — пик новизны: новое объявление получает всплеск трафика (не всегда целевого), который не равен устойчивому спросу.
- Цены в выгрузке якорные: продавцы часто ставят заниженную базовую цену ради трафика; аномально низкая цена — чаще фейк/кликбейт, скрытая часть цены, неликвид или ошибка данных, а не свободный сегмент; outliers_high — несопоставимая позиция (оптовая партия, другая единица измерения, скрытая цена), не выдавай её за уровень.
- Доли сегментов (per_unit/service/product) — это структура предложения, не спроса; о спросе конкретного сегмента говори только по его просмотрам, если их нет — не утверждай.

СТРУКТУРА ИТОГОВОГО ОТЧЁТА (Markdown, заголовки дословно, ровно в этом порядке; каждый заголовок — строго с четырьмя решётками «####», не «###» и не «##»). «Итог» из задачи = разделы 1–8 ниже: 1) данные и оговорки; 2) карта поисковых кластеров; 3) сценарии покупки и проблемы клиентов; 4) карта предложений конкурентов; 5) карта спроса; 6) перегруженные и слабо занятые сегменты; 7) незакрытые коммерческие сценарии; 8) 5–15 рекламных гипотез.

${REPORT_HEADINGS[0]}
- Объём выборки: file.valid_rows («Объём выборки: N строк»), рядом total_rows и skipped_rows; время выгрузки parse.parsed_at («Время выгрузки (ЧЧ:ММ): 14:00»).
- Ниша: что именно представлено в данных — название и тип ниши (товар/услуга) выводи из заголовков, категорий, цен и сегментов.
- Регион — только из блока «Контекст:»; если его там нет — «регион не указан в выгрузке», не выдумывай.
- Прогноз для нового объявления — обязательный блок в этом разделе, все числа дословно из полей forecast JSON (без собственных вычислений; если по группе данных нет — «нет данных по группе»):
  **Прогноз для нового объявления (просмотры сегодня):**
  - в среднем по выгрузке: views_today_avg;
  - в среднем с платными сервисами: views_today_avg_paid (paid_listings объявлений);
  - в среднем без платных сервисов: views_today_avg_organic (organic_listings объявлений);
  - минимальное количество просмотров сегодня: views_today_min; максимальное: views_today_max.
  **Свежесть объявлений (по колонке «Дата публикации»):**
  - объявлениям 2 дня: age_2_days.listings объявлений, в среднем age_2_days.views_today_avg просмотров сегодня (минимум …, максимум …);
  - объявлениям 5 дней: то же по age_5_days.
  Если возрастных данных нет (нет колонки «Дата публикации») или в группе 0 объявлений — так и напиши, не выдумывай. Если в группе меньше 3 объявлений — добавь оговорку «мало данных для вывода». Сравни средние у 2 дней, 5 дней и по всей выгрузке — это проверка пика новизны (первые дни обычно дают больше просмотров). Пик подтверждается, только если среднее у 2 дней выше и чем у 5 дней, и чем среднее по выгрузке: тогда назови разницу с этими двумя числами. Если среднее у 2 дней не выше — напиши прямо: «пик новизны на этой выгрузке не подтвердился». Разницу считай только между этими уже готовыми числами, собственные метрики не вводи.
  Сразу за цифрами — оговорки: разница между группами «с платными» и «без платных» — наблюдаемая корреляция в этой выгрузке, а не доказательство, что платные сервисы дают прирост; группы без платных — базовый ориентир, к которому ближе новое объявление без продвижения; первые 2–3 дня возможен пик новизны; просмотры — прокси, не обращения и не сделки.
- Обязательные оговорки — дословно: «Количество объявлений ≠ объём спроса» и «Просмотры ≠ спрос (разный возраст/продвижение)».
- Если в блоке «Кластеры:» есть пометка о выборке — упомяни, что предложения кластеров строились по выборке, а счёт выполнялся по всей выгрузке.
- Если таблица строк не включена (valid_rows > 200) — прямо укажи это здесь.

${REPORT_HEADINGS[1]}
- Только серверные данные блока «Кластеры:»: для каждого кластера — метка, «n=count из N», доля share и хотя бы один пример-фрагмент из examples (фрагмент содержит keyword кластера).
- Свои кластеры, собственные count и выдуманные метки запрещены; пересечения допустимы (сумма count ≠ N — нормально).
- Если список кластеров пуст — честная отписка без выдуманных кластеров.

${REPORT_HEADINGS[2]}
- Сценарии покупки и проблемы клиента — по явным признакам в данных: метки и примеры кластеров, text_signals, строки. Если ситуация видна в данных — приводи типовые причины покупки: замена, ремонт, переезд, запуск бизнеса, нестандартный размер, срочная покупка, экономия, решение задачи.
- Опирайся на число из данных (count кластера, счётчик text_signals, пример из rows) и на конкретный пример; то, что не подтверждено числом, подавай как гипотезу для проверки.
- Кластеры покупателей могут быть B2B, B2C и смешанными — не своди их к одному.

${REPORT_HEADINGS[3]}
- Товары и услуги, УТП, цены, повторяющиеся аргументы, занятые и слабо представленные сегменты — из JSON (prices, top_titles, samples, text_signals) и примеров кластеров; обещания — в формулировке «конкуренты заявляют/используют в коммуникации…», не как факт о продукте.
- Цены в любом читаемом виде, обязательный формат «n=, медиана» не требуется, но выдуманных чисел быть не должно; outliers_high — несопоставимая позиция, не уровень сегмента.
- Тексты правил сегментации выведи из segment_rules (правило + примеры заголовков).

${REPORT_HEADINGS[4]}
- Таблица строго с заголовком: «${DEMAND_MAP_TABLE_HEADER}».
- Ровно одна строка на каждый кластер из блока «Кластеры:». Колонка «Кол-во объявлений» — только серверный count, без изменений.
- Качественные колонки («Тип спроса», «Как предлагают конкуренты», «Конкуренция», «Коммерческий потенциал», «Возможность для теста») — кратко и по данным: опора на счётчик, пример или цену из JSON.
- Если кластеров нет — раздел заполни честной отпиской, без таблицы и без выдуманных строк.

${REPORT_HEADINGS[5]}
- «Перегруженный» и «слабо занятый» сегменты — через count кластеров из блока «Кластеры:» («n=count из N»); порог называй конкретным числом, оценки вроде «очень мало» без числа не используй.
- Если данных для оценки занятости мало — так и скажи: сегмент под вопросом, нужна проверка.

${REPORT_HEADINGS[6]}
- Незакрытые коммерческие сценарии — это предположения, пока не проверенные данными: каждое оформляй как гипотезу со способом проверки (что замерить, на каких контрольных объявлениях, какой сигнал подтвердит или опровергнет).

${REPORT_HEADINGS[7]}
- От 5 до 15 гипотез. Каждая — ровно тремя строками подряд в формате из задачи 6:
  **Гипотеза:** Если сделать объявление для [сегмент/ситуация] с оффером [решение], то можно получить дополнительный целевой спрос.
  **Почему:** <наблюдение из данных: число, count кластера, счётчик text_signals, пример из выдачи>
  **Что тестировать:** <сценарий / проблема / оффер / ключевой запрос>
- Строка «**Почему:**» обязана опираться на конкретное наблюдение из данных — иначе гипотеза недоказуема.
- Редкие сигналы (count ≤ 2 или share ≤ 0.05) не выдавай за подтверждённый приём: только с оговоркой «данных недостаточно — экспериментальная гипотеза».

Язык — русский, Markdown. Без вступлений вида «В данном отчёте мы рассмотрим…». Отчёт лаконичный и структурированный: каждый раздел — по правилам своего блока, каждый факт — из данных. Отчёт должен быть самодостаточным инструментом для принятия решений.`;

export const REPORT_CRITIC = `Ты — ревизор отчёта по выгрузке объявлений. На вход ты получаешь блок «Данные:» (агрегированный JSON), блок «Контекст:» (регион/запрос, если есть), блок «Кластеры:» (серверный подсчёт: метка, count, share, примеры), блок «Строки:» (markdown-таблица строк, если включена) и draft_report (черновик отчёта в Markdown).

Твоя единственная задача — анти-галлюцинации: убрать из черновика всё, чего нет во входных данных. Правила:
1. Выдуманные цифры, факты, бренды и ассортимент, которых нет во входных данных, → исправь по данным либо удали утверждение.
2. Count каждого кластера в разделах 2, 5 и 6 сверяй с блоком «Кластеры:»: несовпадение, выдуманный кластер или счётчик без числа → переформулируй с корректным «n=count из N» (можно добавить фрагмент из examples) либо удали строку.
3. Строки «**Почему:**» сверяй с JSON, блоком «Кластеры:» и таблицей строк: в них должно быть конкретное наблюдение из входных данных; без числа или с противоречием → добавь корректное число либо удали блок гипотезы целиком.
4. Гипотез в разделе 8 должно быть от 5 до 15: меньше — добавь по данным, больше — удали самые слабые.
5. Регион и запрос — только из блока «Контекст:»; если блока нет, а регион назван — убери выдуманный регион.
6. Обещания продавцов не должны подаваться как факт о продукте — формулируй как «конкуренты заявляют…».
7. Всё вне исправляемых строк не переписывай: верни полный отчёт в Markdown дословно, изменив только строки с противоречиями.

Заголовки разделов — дословно, ровно эти восемь: ${REPORT_HEADINGS.join(" | ")}. Пропущенный или переименованный заголовок → восстанови исходный.

Формат ответа — только полный исправленный отчёт в Markdown, без пояснений и без кодового блока, со всеми разделами «#### 1 … #### 8». Если противоречий нет — верни черновик без изменений.`;

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

const PROPOSAL_SAMPLE_CAP = 1000;

function buildUserMessage(metrics: AggregatedMetrics): string {
  const {
    rows,
    cluster_source: _clusterSource,
    clusters,
    context,
    ...jsonMetrics
  } = metrics;
  void _clusterSource;
  const rowsBlock =
    rows.length > 0
      ? buildRowsTable(rows)
      : "Таблица строк не включена: valid_rows > 200. Анализируй по счётчикам text_signals, top_titles и samples и отметь это в разделе 1.";

  const contextLines: string[] = [];
  if (context.region) contextLines.push(`- Регион: ${context.region}`);
  if (context.query) contextLines.push(`- Запрос: ${context.query}`);
  const contextBlock =
    contextLines.length > 0 ? `Контекст:\n${contextLines.join("\n")}` : null;

  const clusterStats = clusters ?? [];
  const validRows = metrics.file.valid_rows;
  const sampled = metrics.cluster_source.length > PROPOSAL_SAMPLE_CAP;
  const clustersBlock =
    clusterStats.length === 0
      ? "Список кластеров пуст: кластерный анализ не дал результатов по этой выгрузке. В разделе 2 укажи это честно, кластеры не выдумывай."
      : [
          sampled
            ? `Подсчёт по всем ${validRows} строкам; предложение кластеров строилось по выборке из ${PROPOSAL_SAMPLE_CAP} записей.`
            : `Подсчёт по всем ${validRows} строкам.`,
          ...clusterStats.map((c) => {
            const pct = Math.round(c.share * 100);
            const examples =
              c.examples.length > 0
                ? `; примеры: ${c.examples.map((e) => `«${e}»`).join(" ")}`
                : "";
            return `- «${c.label}» (key=${c.key}): n=${c.count} из ${validRows} (${pct}%)${examples}`;
          }),
        ].join("\n");

  const blocks = [`Данные:\n${JSON.stringify(jsonMetrics)}`];
  if (contextBlock) blocks.push(contextBlock);
  blocks.push(`Кластеры:\n${clustersBlock}`);
  blocks.push(`Строки:\n${rowsBlock}`);
  return blocks.join("\n\n");
}

function hasAllSections(text: string): boolean {
  return REPORT_HEADINGS.every((heading) => text.includes(heading));
}

function normalizeHeadings(text: string): string {
  return text.replace(/^#{1,6}(?=[ ]*\d+\.\s)/gm, "####");
}

function buildProposalInput(source: ClusterSourceRow[]): string {
  let items = source;
  let sampled = false;
  if (source.length > PROPOSAL_SAMPLE_CAP) {
    sampled = true;
    items = [];
    for (let k = 0; k < PROPOSAL_SAMPLE_CAP; k += 1) {
      items.push(source[Math.floor((k * source.length) / PROPOSAL_SAMPLE_CAP)]);
    }
  }
  const head = sampled
    ? `Записи: выборка из ${PROPOSAL_SAMPLE_CAP} из ${source.length}. Предложи кластеры по выборке; счёт по всей выборке выполняет сервер.`
    : `Записи: ${source.length}.`;
  const lines = items.map(
    (row, i) =>
      `${i + 1}. ${row.title.replace(/\s+/g, " ").trim()} | ${row.description
        .replace(/\s+/g, " ")
        .trim()
        .slice(0, 120)}`
  );
  return `${head}\n${lines.join("\n")}`;
}

function parseProposalJson(raw: string): unknown {
  const text = raw.trim();
  const fenced = /^```(?:json)?\s*([\s\S]*?)\s*```$/.exec(text);
  let payload = fenced ? fenced[1].trim() : text;
  if (!payload.startsWith("[")) {
    const start = payload.indexOf("[");
    const end = payload.lastIndexOf("]");
    if (start === -1 || end <= start) {
      throw new Error("в ответе proposal нет JSON-массива");
    }
    payload = payload.slice(start, end + 1);
  }
  try {
    return JSON.parse(payload);
  } catch {
    throw new Error("невалидный JSON в ответе proposal");
  }
}

export async function generateReport(
  metrics: AggregatedMetrics
): Promise<GeneratedReport> {
  const config = getConfig();
  const warnings: string[] = [];
  let effectiveMetrics = metrics;

  try {
    const proposalReply = await requestCompletion(
      config,
      [
        { role: "system", content: PROPOSAL_PROMPT },
        { role: "user", content: buildProposalInput(metrics.cluster_source) },
      ],
      0
    );
    const clusterKeys = sanitizeClusters(parseProposalJson(proposalReply));
    effectiveMetrics = {
      ...metrics,
      clusters: countClusters(metrics.cluster_source, clusterKeys),
    };
  } catch (error) {
    const reason = error instanceof Error ? error.message : "неизвестная ошибка";
    warnings.push(`Кластерный анализ не выполнен: ${reason}`);
    effectiveMetrics = { ...metrics, clusters: [] };
  }

  const userMessage = buildUserMessage(effectiveMetrics);

  const draft = await requestCompletion(
    config,
    [
      { role: "system", content: REPORT_SYSTEM_PROMPT },
      { role: "user", content: userMessage },
    ],
    0.3
  );

  let report = normalizeHeadings(draft);
  try {
    const criticReply = await requestCompletion(
      config,
      [
        { role: "system", content: REPORT_CRITIC },
        {
          role: "user",
          content: `${userMessage}\n\ndraft_report:\n${draft}`,
        },
      ],
      0.2
    );
    const criticReport = normalizeHeadings(criticReply);
    if (!hasAllSections(criticReport)) {
      warnings.push(
        "Self-correction не выполнен: в ответе критика нет обязательных разделов отчёта."
      );
    } else {
      report = criticReport;
    }
  } catch (error) {
    const reason = error instanceof Error ? error.message : "неизвестная ошибка";
    warnings.push(`Self-correction не выполнен: ${reason}`);
  }

  warnings.push(...validateReport(report, effectiveMetrics));
  return { report, warnings };
}
