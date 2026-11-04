-- Часть 3, задание 3.2. Планы запроса Q1 к CSV и к Parquet.
-- Запуск: из каталога lab_2 (пути к файлам относительные).

explain analyze
select federal_district, sum(amount_rub)
from read_csv('data/txn.csv', types={'mcc': 'VARCHAR', 'txn_date': 'DATE'})
where status = 'approved' group by 1;

explain analyze
select federal_district, sum(amount_rub)
from read_parquet('data/txn_zstd.parquet')
where status = 'approved' group by 1;
