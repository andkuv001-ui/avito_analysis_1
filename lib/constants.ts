export const MAX_FILE_SIZE = 3 * 1024 * 1024;
export const MAX_ROWS = 10_000;
export const MIN_VALID_ROWS = 5;
export const DEV_PORT = 8090;

export const REQUIRED_COLUMNS = [
  "Название продавца",
  "Заголовок",
  "Цена",
  "Позиция объявлений",
  "Просмотров сегодня",
  "Всего просмотров",
] as const;

export const PER_UNIT_REGEX =
  /(за|\/|\s)(кг|грамм|г|м2|м²|кв\.?\s*м|метр|погонн|литр|л|час|смен|сутк|100\s*г|0[,.]5\s*кг|0[,.]5\s*л)(?![a-zа-яё0-9])/i;

export const SERVICE_REGEX =
  /услуг|работ|монтаж|установк|под\s*ключ|ремонт|пошив|изготовлени|на\s*заказ|аренд|прокат|обслуживани|диагностик/i;

export const FORBIDDEN_PHRASES = [
  "Создайте уют",
  "Оживите интерьер",
  "индивидуальный подход",
  "высокое качество",
] as const;

export const DAILY_TRAFFIC_PREFIX = "Ожидаемый дневной трафик:";
export const INSUFFICIENT_DAILY_NOTE =
  "Недостаточно данных для суточного прогноза";
export const MIN_HOURS_FOR_DAILY_FORECAST = 8;
