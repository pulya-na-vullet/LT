# НТ пререквизиты :: Worker ump-main-ma-ncins-pa

**UMP** · 21 июля 2026 · актуализировано по BPMN PR (NCINS)

## 1. Согласование

| Аналитик | Согласование |
|---|---|
| TBD | TBD |

---

## 2. Минимальная информация для определения необходимости проведения НТ

| # | Пререквизит | Значение | Пример | Где взять | Роль |
|---|---|---|---|---|---|
| 1 | Описание алгоритма работы тестируемого метода | `[ump-main-ma-ncins-pa]` оркестратор: prepare → stage DOCS_COMPLETED → signing → XOR(TIMEOUT→delete+BANK_REJECT / SIGN_COMPLETED) → payment → XOR(TIMEOUT→delete / PAYMENT_COMPLETED) → finalisation → COMPLETED → SUCCESS. Event SubProcess **PT30M** → Call Activity `ump-delete-documents-ncins-pa` → BANK_REJECT | BPMN `ump-main-ma-ncins-pa.bpmn` | Confluence / BPMN PR | Разработчик |
| 2 | SLA по времени отклика | E2E SLA main TBD с НТ. Sync дочерних: несколько секунд. Таймеры: signing PT25M, payment PT5M, main PT30M | TBD числом | Заказчик / НТ | Аналитик |
| 3 | Частота планируемой нагрузки (ЧПН) | ~89/час среднее, пик ~180/час; ~95% happy / ~5% timeout | 15000 заявок/мес SFA | Расчёт | Аналитик |
| 4 | Прогнозируемая через полгода ЧПН | TBD | — | Аналитик | Аналитик |
| 5 | ЧПН на вызываемые внутри методы | prepare/signing ~89–180; payment/finalisation ~85–170; delete ~4–5; kafka stages ~89–180; kafka status reject ~4–5 | — | Аналитик | Аналитик |
| 6 | Прогнозируемая через полгода ЧПН на зависимости | TBD | — | Аналитик | Аналитик |

---

## 3. Полная информация для проведения НТ

| # | Пререквизит | Значение | Пример | Где взять | Роль |
|---|---|---|---|---|---|
| 6 | Пример кода вызова | Zeebe create instance: `businessKey` + `productCode=NON_CREDIT_INSURANCE` | bpmnProcessId `ump-main-ma-ncins-pa` | BPMN | Тестировщик |
| 7 | Примеры кода ответа | End `result=SUCCESS`; reject `BANK_REJECT` + `forcefullyTerminate=true` | SUCCESS / BANK_REJECT | BPMN | Разработчик |
| 8 | Изменяемые данные | UMP DB напрямую не меняет; стадии через Kafka connectors | DOCS/SIGN/PAYMENT/COMPLETED | BPMN | Разработчик |
| 9 | Скрипт миграции БД | Нет необходимости | — | — | Разработчик |
| 10 | Заглушки | Моки в дочерних; на НТ подкладывать SigningDocs / paymentFinished с `correlationKey=businessKey+".NON_CREDIT_INSURANCE"` | Kafka publish | НТ | Разработчик |
| 11 | Время отклика замоканных зависимостей | Sync несколько сек; async signing на НТ ≤~3 сек; payment почти мгновенно | — | НТ | Разработчик |
| 12 | Сценарий тестирования | Happy-path ~95%; signing timeout ~4.75% → delete; payment timeout редко; main PT30M аварийный | timeoutMessage=TIMEOUT | Аналитик | Аналитик |
