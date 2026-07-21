# НТ пререквизиты — некредитное страхование (ncins)

Пакет документов для специалиста НТ по воркерам UMP (табличный формат по образцу `ump-onboarding-pa`).

| Документ | Процесс |
|---|---|
| [01-ump-main-ma-ncins-pa.md](01-ump-main-ma-ncins-pa.md) | Управление процессом мультизаявки |
| [02-ump-prepare-documents-ncins-pa.md](02-ump-prepare-documents-ncins-pa.md) | Формирование документов |
| [03-ump-signing-documents-ncins-pa.md](03-ump-signing-documents-ncins-pa.md) | Подписание документов |
| [04-ump-payment-ncins-pa.md](04-ump-payment-ncins-pa.md) | Оплата страховки |
| [05-ump-finalisation-ncins-pa.md](05-ump-finalisation-ncins-pa.md) | Финализация заявки |
| [06-ump-delete-documents-ncins-pa.md](06-ump-delete-documents-ncins-pa.md) | Удаление документов |

## Общий профиль нагрузки

| Параметр | Значение |
|---|---|
| База | 15000 заявок / месяц |
| Канал | SFA, рабочие часы ЮЛ |
| Среднее | ~89 заявок/час (21 р.д. × 8 ч) |
| Пик (×2) | ~180 заявок/час |
| Timeout-ветки | ~5% (~750/мес, ~4–5/час) |
| Из timeout: signing | ~95% |

Пункты **TBD** — финализация на встрече с НТ / у аналитики.
