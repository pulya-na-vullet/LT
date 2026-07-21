# НТ пререквизиты — пакет ncins

Актуализировано по коду **develop** `ump-ncins-pa@4a921a1653e` (воркеры + IT + BPMN).

## Download
- Excel: [`nt-prerequisites-ncins-xlsx.zip`](../nt-prerequisites-ncins-xlsx.zip)
- Исходник: [`ump-ncins-pa-develop@4a921a1653e (1).zip`](../ump-ncins-pa-develop@4a921a1653e%20(1).zip)

## JobWorker → HTTP (из кода)

| JobWorker type | Класс | HTTP / поведение |
|---|---|---|
| `get-application-data` | ApplicationDataWorker | `GET /applications/{id}?include=PARTICIPANT&include=PRODUCT` |
| `...get-report-data` | PrepareDataForPrintFormWorker | локально XML→Base64 |
| `update-product` | ProductDataUpdateWorker | `PUT /products/{id}` (policyLink=acId) |
| `...set-hold` / `...create-payment` | InsurancePaymentWorker | **STUB** `{code:SUCCESS}` |
| `...create-contract` | FinalisationWorker | `POST /v1/ins-contracts` |
| `...delete-documents` | DeleteDocumentsWorker | **STUB** `{status:"1"}` |

Внешние connectors (не в этом сервисе): `ea-send-documents.v1`, `transfer-control.v1`.

**BPMN `ump-main-ma-ncins-pa` в develop отсутствует.** Payment timer в develop BPMN = **PT15M**.
