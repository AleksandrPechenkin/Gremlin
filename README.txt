Структура проекта:
- main.gs: основная синхронизация заказов с МойСклад
- payments_sync.gs: перенос оплат в таблицу "Закуплено"
- moysklad_api.gs: авторизация, HTTP-методы, обработка ошибок API
- helpers.gs: общие функции (даты, числа, строки, атрибуты)

Обязательные Script Properties:
- MS_TOKEN
- MS_ORGANIZATION_ID
- EXTERNAL_SPREADSHEET_ID

Опциональные Script Properties:
- MS_STORE_ID
- EXTERNAL_SHEET_NAME (по умолчанию: Закуплено)

---
Синхронизация с облаком

1) OneDrive (уже)
Проект лежит в папке OneDrive (Документы). При включённом клиенте OneDrive и
зелёной галочке у файлов копии уходят в облако Microsoft автоматически.
Перед выключением ПК: сохраните файлы в Cursor (Ctrl+S) и дождитесь синхронизации
(иконка OneDrive в трее — без «в ожидании»).

2) Git + GitHub (рекомендуется для истории версий и копии вне OneDrive)
Git установлен, в папке проекта уже выполнены: git init и первый коммит (ветка master).
Локальная история версий доступна командой: git log --oneline

На github.com создайте пустой репозиторий (без README), в папке проекта в PowerShell:

  git remote add origin https://github.com/ВАШ_ЛОГИН/ИМЯ_РЕПО.git
  git branch -M main
  git push -u origin main

Дальше после правок: git add -A && git commit -m "кратко что сделано" && git push
В репозиторий не кладите токены МойСклад — только код из .gs и PROJECT_CONTEXT.md.
Файл .gitignore уже в проекте (в т.ч. .clasp.json не попадёт в коммит).
