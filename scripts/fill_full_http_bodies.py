#!/usr/bin/env python3
"""Replace 'Методы и примеры' sheets with FULL HTTP request/response bodies from develop IT."""

from __future__ import annotations

import json
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

OUT = Path("/workspace/nt-prerequisites-xlsx")
STUB = Path("/tmp/ump-dev/src/test/resources/stubs/ump/ump_application_response_200.json")

HEADER_FILL = PatternFill("solid", fgColor="D6EAF8")
TITLE_FONT = Font(name="Calibri", size=14, bold=True)
SUB_FONT = Font(name="Calibri", size=11, italic=True)
SECTION_FONT = Font(name="Calibri", size=12, bold=True)
HEADER_FONT = Font(name="Calibri", size=11, bold=True)
CELL_FONT = Font(name="Calibri", size=10)
THIN = Border(
    left=Side(style="thin", color="BFBFBF"),
    right=Side(style="thin", color="BFBFBF"),
    top=Side(style="thin", color="BFBFBF"),
    bottom=Side(style="thin", color="BFBFBF"),
)
WRAP = Alignment(wrap_text=True, vertical="top")

APP_JSON = json.loads(STUB.read_text())
APP_JSON_PRETTY = json.dumps(APP_JSON, ensure_ascii=False, indent=2)

# Compact but complete ContractRequestDto reconstructed from ContractMapper + FinalisationWorkerIT
CONTRACT_REQUEST = {
    "programId": 3,
    "contractNumber": "Z6922/888/ABR14880/10",
    "beginDate": "2027-07-01T00:00:00Z",
    "endDate": "2028-07-01T00:00:00Z",
    "signDate": "2027-07-01T00:00:00Z",
    "insuranceSum": 150000.93,
    "insurancePremium": 50000.55,
    "debitAccount": "40802810000000000000",
    "duration": 12,
    "paymentType": "payment_account",
    "contractLink": "http://contract.link",
    "policyLink": "http://policy.link",
    "agreementLink": "http://agreement.link",
    "sellerId": "ASASAS",
    "sellerChannel": "SFA",
    "owner": {
        "inn": "7707083893",
        "email": "romashka@rambler.ru",
        "ownerId": "AAAX22",
        "ogrn": None,
        "phoneNumber": "+79001234567",
        "legalAddress": "г. Пушкино ул. Колотушкина д. 14/88",
    },
    "insuranceObjects": [
        {
            "employeeFIO": "Иванов Иван Иванович",
            "employeeEmail": "ivanov.ii@example.com",
            "employeePhoneNumber": "+7 (999) 123-45-67",
            "employeeBirthDate": "1985-05-15",
            "paymentAccount": "40817810099910004312",
            "cadastralNumber": "77:01:0001045:32",
            "area": 45.5,
            "realEstateAddress": "г. Москва, ул. Тверская, д. 15, кв. 78",
            "realEstateType": "COMMERCIAL",
        },
        {
            "employeeFIO": "Иванов Иван Иванович",
            "employeeEmail": "ivanov.ii@example.com",
            "employeePhoneNumber": "+7 (999) 123-45-67",
            "employeeBirthDate": "1985-05-15",
            "paymentAccount": "40817810099910004312",
            "cadastralNumber": "77:01:0001045:32",
            "area": 45.5,
            "realEstateAddress": "г. Москва, ул. Тверская, д. 15, кв. 78",
            "realEstateType": "RESIDENTIAL",
        },
    ],
}
CONTRACT_REQ_PRETTY = json.dumps(CONTRACT_REQUEST, ensure_ascii=False, indent=2)

# PUT product — reconstructed from ProductDataUpdateWorker (policyLink = acId)
PUT_PRODUCT_REQ = {
    "code": "NON_CREDIT_INSURANCE",
    "name": "Некредитное страхование ЮЛ",
    "productProperties": {
        "code": "NON_CREDIT_INSURANCE",
        "programId": 1073741824,
        "contractNumber": "1423423/sdf/123",
        "beginDate": "2026-06-25T11:47:13.57Z",
        "endDate": "2026-06-25T11:47:13.57Z",
        "signDate": "2026-06-25T11:47:13.57Z",
        "duration": 12,
        "insuranceSum": 20000.0,
        "insurancePremium": 20000.0,
        "paymentType": "payment_account",
        "agreementLink": "7ca85f64-5717-4562-b3fc-6c163f65aba9",
        "policyLink": "770e8400-e29b-41d4-a716-446655110002",
        "insuranceObjects": [{"paymentAccount": "123523464567347"}],
    },
}
PUT_PRODUCT_REQ_PRETTY = json.dumps(PUT_PRODUCT_REQ, ensure_ascii=False, indent=2)

REPORT_XML = """<?xml version="1.0" encoding="UTF-8"?>
<datasource>
  <contractNumber>NC-2025-001</contractNumber>
  <fullName>Иванов Иван Иванович</fullName>
  <inn>123456789012</inn>
  <email>ivanov@example.com</email>
  <paymentAccount>40817810000000000001</paymentAccount>
  <insuranceSum>1000000.00</insuranceSum>
  <insurancePremium>50000.00</insurancePremium>
  <beginDate>...</beginDate>
  <endDate>...</endDate>
  <currentDate>2025-01-15</currentDate>
</datasource>"""

GET_APP_REQ = """GET /applications/6dc5daf3-da12-4aae-9766-572da7b6b003?include=PARTICIPANT&include=PRODUCT HTTP/1.1
Host: ump-application-facade.ump.svc.cluster.local
Accept: application/json

(без body)

Camunda job input:
{
  "businessKey": "6dc5daf3-da12-4aae-9766-572da7b6b003",
  "processType": "<PREPARE_DOCUMENTS|SIGNING|PAYMENT|FINALISATION|DELETE_DOCS>"
}"""

GET_APP_RESP = f"""HTTP/1.1 200 OK
Content-Type: application/json

{APP_JSON_PRETTY}

Camunda job output (успех):
{{ "getProcessParams": "SUCCESS", ...поля по processType из ApplicationMapper }}

Ошибка:
{{ "getProcessParams": "ERROR" }}"""

REPORT_REQ = """JobWorker: ump-prepare-documents-ncins-pa.get-report-data
HTTP: нет (локальная сборка XML + Base64)

Camunda input (PrepareDataForPrintFormWorkerIT):
{
  "businessKey": "550e8400-e29b-41d4-a716-446655110003",
  "programId": 1,
  "fullName": "Иванов Иван Иванович",
  "inn": "123456789012",
  "email": "ivanov@example.com",
  "contractNumber": "NC-2025-001",
  "beginDate": "2025-01-01T00:00:00+03:00",
  "endDate": "2026-01-01T00:00:00+03:00",
  "currentDate": "2025-01-15",
  "insuranceSum": 1000000.0,
  "insurancePremium": 50000.0,
  "paymentAccount": "40817810000000000001"
}"""

REPORT_RESP = f"""Camunda output:
{{
  "serviceId": "550e8400-e29b-41d4-a716-446655110003",
  "serviceCode": "APPLICATION",
  "productCode": "NON_CREDIT_INSURANCE",
  "documents": [
    {{
      "documentType": "CONTRACT_ACCOUNT_BLOCK",
      "reportData": "<Base64 от XML ниже>",
      "isWriteInEa": false
    }}
  ]
}}

XML до Base64 (IT assert):
{REPORT_XML}"""

PUT_REQ = f"""PUT /products/3d9fe175-2bc5-442a-b900-48522551de3f HTTP/1.1
Host: ump-products.ump.svc.cluster.local
Content-Type: application/json

{PUT_PRODUCT_REQ_PRETTY}

Примечание: body = ProductCommonParamDto из productInfo;
worker ДО вызова ставит productProperties.policyLink = acDocuments[0].acId.
Живого JSON в IT нет (Mockito) — тело собрано по stub заявки + логике ProductDataUpdateWorker."""

PUT_RESP = """HTTP/1.1 200 OK
Content-Type: application/json

ProductDto (возврат Feign; в IT не зафиксирован literal — обычно эхо обновлённого продукта).

Camunda job output:
{ "getProcessParams": "SUCCESS" }

При exception:
{ "getProcessParams": "ERROR" }"""

HOLD_REQ = """JobWorker: ump-payment-ncins-pa.set-hold
HTTP: нет (PaymentService TODO stub)

Camunda input (InsurancePaymentWorkerIT):
{
  "businessKey": "550e8400-e29b-41d4-a716-446655110003",
  "productCode": "NON_CREDIT_INSURANCE",
  "systemCode": "SFA"
}"""

HOLD_RESP = """Camunda output:
{
  "code": "SUCCESS"
}
// message = null

Ошибка:
{ "code": "ERROR", "message": "<text>" }"""

PAY_REQ = """JobWorker: ump-payment-ncins-pa.create-payment
HTTP: нет (PaymentService TODO stub)

Camunda input: тот же Map, что для set-hold (см. выше)."""

PAY_RESP = HOLD_RESP

CONTRACT_HTTP_REQ = f"""POST /v1/ins-contracts HTTP/1.1
Host: corp-gateway-test.moscow.alfaintra.net
  /corp-ncins-acc-gateway/secure/corp-ncins-acc-corp-ncins-acc-api
Content-Type: application/json
A-userId: AAAX22
A-customerId: AAAX22
A-clientType: BACKEND
A-channelId: ump
A-projectId: ncins
Authorization: Bearer <keycloak access_token>

{CONTRACT_REQ_PRETTY}

Источник: FinalisationWorkerIT.createValidVariables + ContractMapper → ContractRequestDto."""

CONTRACT_HTTP_RESP = """HTTP/1.1 201 Created
Content-Type: application/json

{}

Camunda job output:
{
  "result": "SUCCESS",
  "errorMessage": null
}

HTTP 503 (после 3 retry):
{
  "result": "ERROR",
  "errorMessage": "<exception log>"
}"""

DELETE_REQ = """JobWorker: ump-delete-documents-ncins-pa.delete-documents
HTTP: нет (DeleteDocumentsService TODO AC 1.0 stub)

Camunda input (DeleteDocumentsWorkerIT / multiInstance element):
{
  "businessKey": "550e8400-e29b-41d4-a716-446655110000",
  "document": {
    "documentLink": "https://example.com/doc/123"
  }
}"""

DELETE_RESP = """Camunda output:
{
  "documentResponse": {
    "status": "1"
  }
}

При exception:
{
  "documentResponse": {
    "status": "2"
  }
}

Реальный POST AlfaCapture deleteDocsh — TBD (stub)."""

EA_REQ = """Connector (вне ump-ncins-pa): ump-document-connectors.ea-send-documents.v1
×3 параллельно после create-contract

Input (BPMN):
{
  "acId": "<agreementLink | policyLink | contractLink>",
  "stopInIncident": false
}

Полный HTTP body connector — в репе ncins нет (внешний модуль)."""

EA_RESP = """Response необязателен.
stopInIncident=false — отсутствие документов не обязано ронять в инцидент.
Полный response TBD (модуль document-connectors)."""

TRANSFER_REQ = """Connector (вне ump-ncins-pa): ump-process-transfer-connector.transfer-control.v1

BPMN inputs (develop):
{
  "stopInIncident": true,
  "useProcessInstanceKey": true,
  "systemCode": "<productCode>",
  "messageName": "SigningDocs",
  "applicationId": "<businessKey>"
}

Полный Kafka JSON to.system — TBD (внешний connector / аналитика)."""

TRANSFER_RESP = """Далее receive message SigningDocs, correlationKey=processInstanceKey (develop).
Payload from.system — TBD."""

HEADERS = [
    "#",
    "Метод / JobWorker",
    "Где",
    "Полный REQUEST (headers + body)",
    "Полный RESPONSE (status + body)",
    "Источник",
]


def write_methods_sheet(wb: openpyxl.Workbook, title: str, rows: list[tuple]) -> None:
    name = "Методы и примеры (код)"
    if name in wb.sheetnames:
        del wb[name]
    ws = wb.create_sheet(name)
    ws["A1"] = title
    ws["A1"].font = TITLE_FONT
    ws["A2"] = (
        "Полные request/response (не только имя метода). "
        "Источник: develop ump-ncins-pa@4a921a1653e + *WorkerIT + stubs."
    )
    ws["A2"].font = SUB_FONT
    ws["A4"] = "HTTP / JobWorker — полные тела запросов и ответов"
    ws["A4"].font = SECTION_FONT
    for i, h in enumerate(HEADERS, 1):
        cell = ws.cell(row=5, column=i, value=h)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = WRAP
        cell.border = THIN
    for r_i, row in enumerate(rows, 6):
        for c_i, val in enumerate(row, 1):
            cell = ws.cell(row=r_i, column=c_i, value=val)
            cell.font = CELL_FONT
            cell.alignment = WRAP
            cell.border = THIN
        # tall rows for JSON
        text = str(row[3]) + "\n" + str(row[4])
        ws.row_dimensions[r_i].height = min(420, max(120, 12 + text.count("\n") * 8))
    widths = [5, 28, 22, 55, 55, 22]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


# Catalog of all methods
ALL_METHODS = [
    (
        "1",
        "get-application-data\nGET /applications/{id}",
        "ApplicationDataWorker\nвсе процессы",
        GET_APP_REQ,
        GET_APP_RESP,
        "ApplicationDataWorkerIT\nstubs/ump/ump_application_response_200.json",
    ),
    (
        "2",
        "ump-prepare-documents-ncins-pa.get-report-data\n(локально XML→Base64)",
        "PrepareDataForPrintFormWorker\nprepare",
        REPORT_REQ,
        REPORT_RESP,
        "PrepareDataForPrintFormWorkerIT",
    ),
    (
        "3",
        "update-product\nPUT /products/{id}",
        "ProductDataUpdateWorker\nprepare",
        PUT_REQ,
        PUT_RESP,
        "ProductDataUpdateWorker + stub заявки\n(IT без literal PUT body)",
    ),
    (
        "4",
        "ump-payment-ncins-pa.set-hold\nSTUB (нет HTTP)",
        "InsurancePaymentWorker\npayment",
        HOLD_REQ,
        HOLD_RESP,
        "InsurancePaymentWorkerIT\nPaymentService stub",
    ),
    (
        "5",
        "ump-payment-ncins-pa.create-payment\nSTUB (нет HTTP)",
        "InsurancePaymentWorker\npayment",
        PAY_REQ,
        PAY_RESP,
        "InsurancePaymentWorkerIT\nPaymentService stub",
    ),
    (
        "6",
        "ump-finalisation-ncins-pa.create-contract\nPOST /v1/ins-contracts",
        "FinalisationWorker\nfinalisation",
        CONTRACT_HTTP_REQ,
        CONTRACT_HTTP_RESP,
        "FinalisationWorkerIT\nContractMapper",
    ),
    (
        "7",
        "ump-delete-documents-ncins-pa.delete-documents\nSTUB AC 1.0",
        "DeleteDocumentsWorker\ndelete",
        DELETE_REQ,
        DELETE_RESP,
        "DeleteDocumentsWorkerIT\nDeleteDocumentsService stub",
    ),
    (
        "8",
        "ump-document-connectors.ea-send-documents.v1\n×3 parallel",
        "внешний connector\nfinalisation",
        EA_REQ,
        EA_RESP,
        "BPMN finalisation\n(код connector вне репы)",
    ),
    (
        "9",
        "ump-process-transfer-connector.transfer-control.v1",
        "внешний connector\nsigning",
        TRANSFER_REQ,
        TRANSFER_RESP,
        "BPMN signing\n(Kafka JSON TBD)",
    ),
]

PER_FILE = {
    "НТ пререквизиты __ Worker ump-main-ma-ncins-pa.xlsx": ALL_METHODS,
    "НТ пререквизиты __ Worker ump-prepare-documents-ncins-pa.xlsx": [
        ALL_METHODS[0],
        ALL_METHODS[1],
        ALL_METHODS[2],
        (
            "4",
            "Call Activity ump-generate-and-save-document-pa",
            "внешний процесс\nprepare",
            """calledElement inputs:
{
  "serviceId": "<businessKey>",
  "serviceCode": "APPLICATION",
  "productCode": "<productCode>",
  "documents": [ { "documentType": "CONTRACT_ACCOUNT_BLOCK", "reportData": "<Base64>", "isWriteInEa": false } ]
}
Полный HTTP ПФ/AC — во внешнем процессе (не в ump-ncins-pa).""",
            """outputs:
{
  "docsResult": "<result>",
  "acRequestId": "<...>",
  "acDocuments": [ { "type": "CONTRACT_ACCOUNT_BLOCK", "acId": "<uuid>" } ]
}""",
            "BPMN prepare Call Activity",
        ),
    ],
    "НТ пререквизиты __ Worker ump-signing-documents-ncins-pa.xlsx": [
        ALL_METHODS[0],
        ALL_METHODS[8],
    ],
    "НТ пререквизиты __ Worker ump-payment-ncins-pa.xlsx": [
        ALL_METHODS[0],
        ALL_METHODS[3],
        ALL_METHODS[4],
        (
            "4",
            "message paymentFinished",
            "Zeebe receive\npayment",
            """correlationKey = businessKey (develop BPMN)
Полный JSON payload сообщения — TBD (в репе нет; добыть на стенде).
Минимально для НТ: опубликовать completion с corr = businessKey.""",
            "Продолжение процесса / End.\nTimeout PT15M → timeoutMessage=TIMEOUT",
            "BPMN ump-payment-ncins-pa",
        ),
    ],
    "НТ пререквизиты __ Worker ump-finalisation-ncins-pa.xlsx": [
        ALL_METHODS[0],
        ALL_METHODS[5],
        ALL_METHODS[7],
    ],
    "НТ пререквизиты __ Worker ump-delete-documents-ncins-pa.xlsx": [
        ALL_METHODS[0],
        ALL_METHODS[6],
    ],
}


def main() -> None:
    for fname, rows in PER_FILE.items():
        path = OUT / fname
        wb = openpyxl.load_workbook(path)
        title = wb["Содержание"]["A1"].value or fname
        write_methods_sheet(wb, str(title), rows)
        # also bump date
        for sn in wb.sheetnames:
            if wb[sn]["A2"].value:
                s = str(wb[sn]["A2"].value)
                if "полные body" not in s:
                    wb[sn]["A2"] = s + " · полные HTTP body"
        wb.save(path)
        print("updated methods sheet:", fname)


if __name__ == "__main__":
    main()
