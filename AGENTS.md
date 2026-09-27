# AGENTS.md

- Порт: 8090 везде (`npm run dev`, `npm run start`, docker-compose).
- Даты и время — только UTC (в т.ч. имена файлов отчётов).
- Запрещено упоминать конкретные ниши/бренды в коде и константах (никаких «мох», «панно», «фитостена», BMW, iPhone и т.п.). Сегментация — только нишезависимыми regex из `lib/constants.ts`.
- Стек фиксирован: Next.js 14 App Router (TypeScript strict, без `src/`), Tailwind CSS v3 + `tailwind.config.ts` (`@tailwindcss/typography`, `tailwindcss-animate`), локальные shadcn-компоненты в `components/ui`, zod v3, xlsx (SheetJS), react-markdown, sonner, lucide-react.
- Запрещено: MUI/Chakra/Mantine/Ant, framer-motion, собственные `@keyframes`, Supabase/n8n/Redis/Postgres, авторизация, история, хранение файлов, экспорт .docx/.pdf, эмодзи в UI, `alert`/`confirm`, отправка сырого Excel в LLM.
- Команды проверки: `npm run typecheck && npm run lint && npm run build`.
