import { PER_UNIT_REGEX, SERVICE_REGEX } from "./constants";

export type Segment = "per_unit" | "service" | "product";

export function classify(title: string, description: string): Segment {
  const text = `${title} ${description}`;
  if (PER_UNIT_REGEX.test(text)) return "per_unit";
  if (SERVICE_REGEX.test(text)) return "service";
  return "product";
}

export const SEGMENT_LABELS: Record<Segment, string> = {
  per_unit: "Сырьё / материалы (цена за единицу измерения)",
  service: "Услуги",
  product: "Готовые товары",
};
