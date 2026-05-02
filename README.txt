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
Папка в браузере (OneDrive в облаке):
https://onedrive.live.com/?id=%2Fpersonal%2Fd0b2710d08b281b9%2FDocuments%2FGremlin%2F%D0%90%D0%B2%D1%82%D0%BE%D0%BC%D0%B0%D1%82%D0%B8%D0%B7%D0%B0%D1%86%D0%B8%D1%8F%20Gremlin&viewid=37fd5091%2D1e00%2D48a9%2Dadea%2D0ae2cd53bd3d&view=0

2) Git + GitHub (рекомендуется для истории версий и копии вне OneDrive)
Удалённый репозиторий: https://github.com/AleksandrPechenkin/Gremlin
Ветка по умолчанию: main. Локальная история: git log --oneline

После правок в папке проекта:
  git add -A
  git commit -m "кратко что сделано"
  git push
В репозиторий не кладите токены МойСклад — только код из .gs и PROJECT_CONTEXT.md.
Файл .gitignore уже в проекте (в т.ч. .clasp.json не попадёт в коммит).
