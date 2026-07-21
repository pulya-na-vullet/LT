# НТ пререквизиты — пакет ncins

Актуализировано **21 июля 2026** по финальному BPMN из PR (NCINS, commits 1–7) + ответы СА/бэкенд.

## Excel (основной артефакт для НТ)

Скачать zip: [`nt-prerequisites-ncins-xlsx.zip`](../nt-prerequisites-ncins-xlsx.zip)

Папка: [`nt-prerequisites-xlsx/`](../nt-prerequisites-xlsx/)

Структура каждого файла = образец `ump-onboarding-pa`:

1. Согласование  
2. Минимальная информация (пререквизиты 1–6)  
3. Полная информация (пререквизиты 6–12)  
4. BPMN методы (источник) — STATIC/DYNAMIC

| Файл | Процесс |
|---|---|
| `... ump-main-ma-ncins-pa.xlsx` | Оркестратор |
| `... ump-prepare-documents-ncins-pa.xlsx` | Подготовка документов |
| `... ump-signing-documents-ncins-pa.xlsx` | Подписание (PT25M) |
| `... ump-payment-ncins-pa.xlsx` | Оплата (PT5M) |
| `... ump-finalisation-ncins-pa.xlsx` | Финализация (3× parallel EA) |
| `... ump-delete-documents-ncins-pa.xlsx` | Удаление (multiInstance) |

## Ключевые правки по BPMN PR

- Signing/Payment **correlationKey** = `businessKey + "." + "NON_CREDIT_INSURANCE"`
- Timers: signing **PT25M**, payment **PT5M**, main **PT30M**
- Finalisation: **3×** `ea-send-documents.v1` (agreement/policy/contract), `stopInIncident=false`
- Delete: BPMN id `ump-delete-documents-ncins-pa`, `processType=DELETE_DOCS`, multiInstance `documents`
- Main на timeout: Call Activity delete → BANK_REJECT

Markdown-копии в этой папке — краткие шпаргалки; для заполнения НТ использовать Excel.
