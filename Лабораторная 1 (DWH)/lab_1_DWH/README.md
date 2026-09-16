# Лабораторная работа 1. Витрина данных «звезда» для банка «Синица»

- `sql/` — заполненные `staging_template.sql`, `marts_template.sql` и запросы Q1–Q3 (`queries.sql`)
- `Отчет_лаб1.docx` — отчёт по частям 1–5
- `screenshots/` — скриншоты к отчёту

## Запуск (Windows)

```powershell
py -m venv .venv
.venv\Scripts\python -m pip install -r requirements.txt
.venv\Scripts\python generator\generate.py
.venv\Scripts\python -c "import duckdb,pathlib; con=duckdb.connect('bank.duckdb'); [con.execute(pathlib.Path(f).read_text(encoding='utf-8')) for f in ['scripts/raw_views.sql','sql/staging_template.sql','sql/marts_template.sql']]"
```

Самопроверка: в `marts.fct_transactions` 2 312 606 строк, из них 11 557 с `client_sk = -1`.
