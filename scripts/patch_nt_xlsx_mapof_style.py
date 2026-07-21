#!/usr/bin/env python3
"""Rewrite NT Excel п.6/п.7 in onboarding-sample style: Map.of + createInstance + HTTP."""

from __future__ import annotations

from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment

OUT = Path("/workspace/nt-prerequisites-xlsx")
WRAP = Alignment(wrap_text=True, vertical="top")

CAMUNDA_INSTR = (
    "Camunda на базе Java скрипта, инструкция — "
    "https://confluence.moscow.alfaintra.net/pages/viewpage.action?pageId=2136656813"
)

WHERE_CALL = (
    "Сначала нужно добыть живой пример вызова сервиса (локально/стенд) и параметризовать.\n"
    "Параметризовать businessKey (≥ ЧПН). productCode — const NON_CREDIT_INSURANCE.\n"
    "Источник примеров: develop ump-ncins-pa *WorkerIT + stubs."
)
WHERE_RESP = (
    "Нужны примеры ответов, которые можно считать ожидаемыми.\n"
    "Например: getProcessParams=SUCCESS; HTTP 200/201; result=SUCCESS.\n"
    "Источник: *WorkerIT + stubs/ump/ump_application_response_200.json."
)

BK = "550e8400-e29b-41d4-a716-446655110003"
BK2 = "550e8400-e29b-41d4-a716-446655110000"
PROD = "3d9fe175-2bc5-442a-b900-48522551de3f"
AC_ID = "770e8400-e29b-41d4-a716-446655110002"
APP_ID = "6dc5daf3-da12-4aae-9766-572da7b6b003"


def start_block(bpmn_id: str, extra_vars: str = "") -> str:
    extra = (",\n" + extra_vars) if extra_vars else ""
    return f"""У воркера отсутствует HTTP-эндпоинт.
Вызов происходит напрямую в Camunda.
{CAMUNDA_INSTR}

Пример запроса:

Map<String, Object> variables = Map.of(
  "businessKey", "{BK}", // id заявки (UMP application)
  "productCode", "NON_CREDIT_INSURANCE" // const{extra}
);

final ProcessInstanceEvent processInstanceEvent = client
  .newCreateInstanceCommand()
  .bpmnProcessId("{bpmn_id}")
  .latestVersion()
  .variables(variables)
  .send()
  .join();

Параметры для запроса:
1. businessKey — UUID заявки (параметризовать на объём ≥ ЧПН)
2. productCode — const NON_CREDIT_INSURANCE (параметризация не нужна)"""


def patch_full_sheet(path: Path, call_val: str, call_ex: str, resp_val: str, resp_ex: str) -> None:
    wb = openpyxl.load_workbook(path)
    ws = wb["3. Полная информация"]
    # find rows by # in col A
    for row in ws.iter_rows(min_row=6, max_row=20):
        num = str(row[0].value or "")
        if num == "6":
            row[2].value = call_val
            row[3].value = call_ex
            row[4].value = WHERE_CALL
            for c in row:
                c.alignment = WRAP
            ws.row_dimensions[row[0].row].height = 220
        elif num == "7":
            row[2].value = resp_val
            row[3].value = resp_ex
            row[4].value = WHERE_RESP
            for c in row:
                c.alignment = WRAP
            ws.row_dimensions[row[0].row].height = 180
    # bump date line
    if ws["A2"].value:
        base = str(ws["A2"].value).split("·")[0].strip()
        ws["A2"] = f"{base} · п.6/п.7 в формате образца onboarding (Map.of)"
    wb.save(path)
    print("patched", path.name)


# ---- MAIN ----
patch_full_sheet(
    OUT / "НТ пререквизиты __ Worker ump-main-ma-ncins-pa.xlsx",
    start_block("ump-main-ma-ncins-pa")
    + "\n\nВНИМАНИЕ: BPMN main-ma в develop-архиве отсутствует — bpmnProcessId уточнить по ветке с оркестратором.",
    'Map.of("businessKey","550e8400-...","productCode","NON_CREDIT_INSURANCE")',
    """Ответ как у HTTP-эндпоинта у воркера отсутствует (как в образце onboarding п.7).

Ожидаемые исходы процесса (по дочерним + СА):
- Happy-path End: result = SUCCESS
- Timeout/reject: statusCode = BANK_REJECT, statusMessage = "Превышено время ожидания", forcefullyTerminate = true

Дочерние HTTP (см. Excel prepare/finalisation):
- GET /applications/{id} → 200 + ApplicationResponseDto
- PUT /products/{id} → SUCCESS
- POST /v1/ins-contracts → 201 {}""",
    "SUCCESS / BANK_REJECT; дочерние 200/201",
)

# ---- PREPARE ----
prepare_call = (
    start_block("ump-prepare-documents-ncins-pa")
    + f"""

--- Внутренние вызовы воркеров (из кода/IT) ---

1) ApplicationDataWorker type=get-application-data
Camunda job vars:
Map.of(
  "businessKey", "{BK}",
  "processType", "PREPARE_DOCUMENTS" // STATIC в BPMN
);
HTTP:
GET {{applicationFacade.url}}/applications/{BK}?include=PARTICIPANT&include=PRODUCT
Host example: http://ump-application-facade.ump.svc.cluster.local

2) PrepareDataForPrintFormWorker type=ump-prepare-documents-ncins-pa.get-report-data
Вход (IT PrepareDataForPrintFormWorkerIT) — поля заявки:
programId=1, fullName="Иванов Иван Иванович", inn="123456789012",
email="ivanov@example.com", contractNumber="NC-2025-001",
beginDate/endDate, currentDate="2025-01-15",
insuranceSum=1000000.0, insurancePremium=50000.0,
paymentAccount="40817810000000000001"

3) Call Activity ump-generate-and-save-document-pa (внешний)
serviceId=businessKey, serviceCode="APPLICATION", productCode, documents

4) ProductDataUpdateWorker type=update-product
Map.of(
  "businessKey", "{BK2}",
  "productId", "{PROD}",
  "acDocuments", "[{{ \"type\": \"CONTRACT_ACCOUNT_BLOCK\", \"acId\": \"{AC_ID}\" }}]",
  "productInfo", "{{}}"
);
HTTP: PUT {{products.url}}/products/{PROD}
"""
)
prepare_resp = f"""1) get-application-data
Camunda out: getProcessParams = SUCCESS (+ поля PREPARE_DOCUMENTS).
HTTP 200 stub stubs/ump/ump_application_response_200.json (фрагмент):
{{
  "id": "{APP_ID}",
  "participants": [{{ "fullName": "ООО Звезда", "inn": "7826688577", ... }}],
  "products": [{{
    "id": "{PROD}",
    "code": "NON_CREDIT_INSURANCE",
    "productProperties": {{
      "contractNumber": "1423423/sdf/123",
      "insuranceSum": 20000.0,
      "agreementLink": "7ca85f64-5717-4562-b3fc-6c163f65aba9"
    }}
  }}]
}}
Ошибка: getProcessParams = ERROR

2) get-report-data out:
{{
  "serviceId": "{BK}",
  "serviceCode": "APPLICATION",
  "productCode": "NON_CREDIT_INSURANCE",
  "documents": [{{
    "documentType": "CONTRACT_ACCOUNT_BLOCK",
    "reportData": "<Base64(XML datasource)>",
    "isWriteInEa": false
  }}]
}}

3) update-product out: {{ "getProcessParams": "SUCCESS" }}
   при exception putProduct: {{ "getProcessParams": "ERROR" }}"""

patch_full_sheet(
    OUT / "НТ пререквизиты __ Worker ump-prepare-documents-ncins-pa.xlsx",
    prepare_call,
    "Map.of start + GET applications + PUT products (IT)",
    prepare_resp,
    "getProcessParams=SUCCESS; HTTP 200; documents[].reportData",
)

# ---- SIGNING ----
signing_call = (
    start_block("ump-signing-documents-ncins-pa")
    + f"""

--- Шаги ---

1) ApplicationDataWorker
Map.of("businessKey", "{BK}", "processType", "SIGNING");
HTTP: GET .../applications/{BK}?include=PARTICIPANT&include=PRODUCT

2) Connector ump-process-transfer-connector.transfer-control.v1 (вне репы)
messageName=SigningDocs, useProcessInstanceKey=true, systemCode=productCode

3) Receive message SigningDocs
correlationKey = processInstanceKey (develop BPMN)
Полный JSON Kafka/message — TBD (в репе нет контракта; как в образце — добыть прогоном на стенде).
"""
)
signing_resp = """Ответ HTTP у воркера signing (кроме get) отсутствует.

get-application-data: getProcessParams = SUCCESS | ERROR; для SIGNING маппер отдаёт processInstanceKey.

После message SigningDocs процесс идёт в End.
Ветка systemCode=\"AlfaOffice\" — End без connector.

Ожидаемые коды/статусы для НТ:
- SUCCESS (дошёл message)
- зависание без message (в develop BPMN timer нет — сверить с PR)"""

patch_full_sheet(
    OUT / "НТ пререквизиты __ Worker ump-signing-documents-ncins-pa.xlsx",
    signing_call,
    "Map.of start + GET applications; Kafka JSON TBD",
    signing_resp,
    "getProcessParams=SUCCESS; message TBD",
)

# ---- PAYMENT ----
payment_call = (
    start_block("ump-payment-ncins-pa")
    + f"""

--- Шаги (InsurancePaymentWorkerIT) ---

1) get-application-data processType=PAYMENT
Map.of("businessKey", "{BK}", "processType", "PAYMENT");
HTTP: GET .../applications/{BK}?include=PARTICIPANT&include=PRODUCT

2) set-hold type=ump-payment-ncins-pa.set-hold
Map.of(
  "businessKey", "{BK}",
  "productCode", "NON_CREDIT_INSURANCE",
  "systemCode", "SFA"
);
HTTP: отсутствует (PaymentService TODO stub)

3) create-payment type=ump-payment-ncins-pa.create-payment
те же variables; HTTP отсутствует (stub)

4) Receive paymentFinished, correlationKey=businessKey (develop)
JSON payload — TBD (добыть на стенде / подложить пустой/минимальный message)
"""
)
payment_resp = """set-hold / create-payment (stub):
{ "code": "SUCCESS" }
// message = null

При ошибке маппера:
{ "code": "ERROR", "message": "<text>" }

get-application-data: getProcessParams = SUCCESS | ERROR
(+ fullName, productId, inn, contractNumber, insurancePremium, debitAccount)

Timeout event sub-process (PT15M): timeoutMessage = "TIMEOUT"

HTTP EQ в develop нет — негативы 4xx/5xx пока не применимы."""

patch_full_sheet(
    OUT / "НТ пререквизиты __ Worker ump-payment-ncins-pa.xlsx",
    payment_call,
    'Map.of(...); stub без HTTP; message paymentFinished TBD',
    payment_resp,
    '{"code":"SUCCESS"} / TIMEOUT',
)

# ---- FINALISATION ----
final_call = (
    start_block("ump-finalisation-ncins-pa")
    + f"""

--- Шаги (FinalisationWorkerIT) ---

1) get-application-data processType=FINALISATION
Map.of("businessKey", "{BK2}", "processType", "FINALISATION");
HTTP: GET .../applications/{BK2}?include=PARTICIPANT&include=PRODUCT

2) create-contract type=ump-finalisation-ncins-pa.create-contract
Camunda vars (createValidVariables из IT) — эквивалент Map:
Map.of(
  "programId", 3,
  "inn", "7707083893",
  "email", "romashka@rambler.ru",
  "contractNumber", "Z6922/888/ABR14880/10", // параметризовать на каждый прогон
  "beginDate", "2027-07-01T00:00:00Z",
  "endDate", "2028-07-01T00:00:00Z",
  "insuranceSum", 150000.93,
  "insurancePremium", 50000.55,
  "signDate", "2027-07-01T00:00:00Z",
  "debitAccount", "40802810000000000000",
  "duration", 12,
  "paymentType", "payment_account",
  "contractLink", "http://contract.link",
  "policyLink", "http://policy.link",
  "agreementLink", "http://agreement.link",
  "ownerId", "AAAX22",
  "phoneNumber", "+79001234567",
  "legalAddress", "г. Пушкино ул. Колотушкина д. 14/88",
  "sellerId", "ASASAS",
  "sellerChannel", "SFA"
  // + insuranceObjects[] (2 объекта, см. FinalisationWorkerIT)
);

HTTP:
POST {{integration.nib.corp-ncins-acc-api.url}}/v1/ins-contracts
Headers:
  A-userId: AAAX22
  A-customerId: AAAX22
  A-clientType: BACKEND
  A-channelId: ump
Body: ContractRequestDto (из маппера ContractCreateInputVariables)

3) 3× connector ea-send-documents.v1 (вне репы)
acId = agreementLink | policyLink | contractLink; stopInIncident=false
"""
)
final_resp = """HTTP create-contract: 201 Created, body {{}}

Camunda worker out success:
{ "result": "SUCCESS", "errorMessage": null }

HTTP 503 → 3 retry →:
{ "result": "ERROR", "errorMessage": "<log>" }

HTTP 401 → refresh token + retry (IT shouldRetryWithNewTokenOn401)

get-application-data: getProcessParams = SUCCESS | ERROR

EA connector: response необязателен; при отсутствии документов не обязан инцидент."""

patch_full_sheet(
    OUT / "НТ пререквизиты __ Worker ump-finalisation-ncins-pa.xlsx",
    final_call,
    "Map.of start + POST /v1/ins-contracts (IT)",
    final_resp,
    '201 {}; {"result":"SUCCESS","errorMessage":null}',
)

# ---- DELETE ----
delete_call = (
    start_block("ump-delete-documents-ncins-pa")
    + f"""

Обычно стартует Call Activity из main timeout (не напрямую).

--- Шаги (DeleteDocumentsWorkerIT) ---

1) get-application-data processType=DELETE_DOCS
Map.of("businessKey", "{BK2}", "processType", "DELETE_DOCS");
HTTP: GET .../applications/{BK2}?include=PARTICIPANT&include=PRODUCT
→ documents = [{{"documentLink": "<contract|policy|agreement link>"}}]

2) delete-documents (multiInstance, каждый document)
Map.of(
  "businessKey", "{BK2}",
  "document", "{{ \"documentLink\": \"https://example.com/doc/123\" }}"
);
HTTP: отсутствует (DeleteDocumentsService TODO AC 1.0 stub)
"""
)
delete_resp = """Camunda out (IT):
{ "documentResponse": { "status": "1" } }

При exception в сервисе:
{ "documentResponse": { "status": "2" } }

get-application-data: getProcessParams = SUCCESS | ERROR

Реальный AC deleteDocsh request/response — TBD после снятия stub
(как в образце п.7: если нет — прогнать локально/стенд и сохранить ответ)."""

patch_full_sheet(
    OUT / "НТ пререквизиты __ Worker ump-delete-documents-ncins-pa.xlsx",
    delete_call,
    "Map.of start/Call Activity + stub delete",
    delete_resp,
    '{"documentResponse":{"status":"1"}}',
)

print("done")
