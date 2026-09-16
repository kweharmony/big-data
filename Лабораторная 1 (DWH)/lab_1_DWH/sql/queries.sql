-- Запросы к витрине (часть 5). Выполнять после staging_template.sql и marts_template.sql.

-- Q1. Оборот одобренных покупок в рублях по федеральным округам, июль 2026
select dc.federal_district, round(sum(f.amount_rub)) as turnover_rub
from marts.fct_transactions f
join marts.dim_client dc using (client_sk)
join marts.dim_date d    using (date_sk)
where d.month = 7 and f.status = 'approved' and f.channel in ('pos', 'ecom')
group by 1
order by 2 desc;

-- Q1, «сырой» запрос для сравнения
select r.federal_district, round(sum(t.amount)) as turnover
from raw.transactions t
join raw.cards k    using (card_id)
join raw.accounts a using (account_id)
join raw.clients c  using (client_id)
join raw.regions r  using (city)
where t.txn_ts >= '2026-07-01' and t.txn_ts < '2026-08-01' and t.status = 'approved'
group by 1
order by 2 desc;

-- Q2. Топ-5 мерчантов по обороту среди клиентов сегмента premium за всё лето
select dm.merchant_name, round(sum(f.amount_rub)) as turnover_rub
from marts.fct_transactions f
join marts.dim_client dc   using (client_sk)
join marts.dim_merchant dm using (merchant_sk)
where dc.segment = 'premium' and f.status = 'approved' and f.channel in ('pos', 'ecom')
group by 1
order by 2 desc
limit 5;

-- Q3 (бонус). Доля отклонённых операций по каналам по неделям августа
select d.week_start, f.channel,
       round(100.0 * count(*) filter (where f.status = 'declined') / count(*), 2) as decline_pct
from marts.fct_transactions f
join marts.dim_date d using (date_sk)
where d.month = 8 and f.status in ('approved', 'declined')
group by 1, 2
order by 1, 2;
