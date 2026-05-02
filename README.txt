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
Удалённый репозиторий: https://github.com/AleksandrPechenkin/Gremlin
Ветка по умолчанию: main. Локальная история: git log --oneline

После правок в папке проекта:
  git add -A
  git commit -m "кратко что сделано"
  git push
В репозиторий не кладите токены МойСклад — только код из .gs и PROJECT_CONTEXT.md.
Файл .gitignore уже в проекте (в т.ч. .clasp.json не попадёт в коммит).
