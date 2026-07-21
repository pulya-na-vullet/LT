# НТ пререквизиты :: Worker ump-prepare-documents-ncins-pa

**UMP** · 21 июля 2026 · актуализировано по BPMN PR (NCINS)

## 1. Согласование

| Аналитик | Согласование |
|---|---|
| TBD | TBD |

---

## 2. Минимальная информация

| # | Пререквизит | Значение | Роль |
|---|---|---|---|
| 1 | Алгоритм | Start → `get-application-data` (`processType=PREPARE_DOCUMENTS`) → `get-report-data` → Call Activity `ump-generate-and-save-document-pa` (`serviceCode="APPLICATION"`) → `update-product` → End. MultiInstance на generate **нет**. | Разработчик |
| 2 | SLA | Sync, несколько секунд/вызов (TBD) | Аналитик |
| 3 | ЧПН | ~89/час, пик ~180/час; 1 документ; reportData ≈283 КБ | Аналитик |
| 4 | ЧПН через 6 мес | TBD | Аналитик |
| 5 | ЧПН зависимостей | get / get-report-data / generate-and-save / update-product — каждый ~89–180 | Аналитик |
| 6 | ЧПН зависимостей через 6 мес | TBD | Аналитик |

## 3. Полная информация

| # | Пререквизит | Значение |
|---|---|---|
| 6 | Вызов | Zeebe / Call Activity из main; `businessKey` DYNAMIC |
| 7 | Ответ | SUCCESS; outs Call Activity: `docsResult`, `acRequestId`, `acDocuments`; ошибки ПФ 400/500 TBD |
| 8 | Изменяемые данные | `update-product` → фактически `contractLink` |
| 9 | Миграция | Нет |
| 10 | Заглушки | Call Activity generate/save AC — уточнять мок на НТ |
| 11 | Latency моков | Несколько секунд |
| 12 | Сценарий | Последовательные 4 шага + негативы get/ПФ + пик 180/час |
