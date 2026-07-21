# НТ пререквизиты :: Worker ump-signing-documents-ncins-pa

**UMP** · 21 июля 2026 · актуализировано по BPMN PR (NCINS)

## Ключевые факты BPMN

- Timer: **PT25M**
- Message: `SigningDocs`
- **correlationKey** = `businessKey + "." + "NON_CREDIT_INSURANCE"`
- Transfer: `ump-process-transfer-connector` → `ump.process.to.system` (`systemCode=B2B`, `messageName=GenerateDocs`, `stopInIncident=true`)
- Stage: `SIGN_WAITING`
- После сообщения: `update-product`
- Delete **не** в signing — cleanup в main по `timeoutMessage`

## Пререквизиты 1–12

| # | Значение (кратко) |
|---|---|
| 1 | get → to.system → SIGN_WAITING → Receive SigningDocs → update-product; EventSubProcess PT25M → BANK_REJECT |
| 2 | Sync несколько сек; async на НТ ≤~3 сек (прод до PT25M) |
| 3 | ~89 / пик ~180 в час |
| 4–6 | TBD / зависимости = get, to.system, stage, from.system, update-product |
| 6 (full) | Publish SigningDocs с corr=`<bk>.NON_CREDIT_INSURANCE` |
| 7 | SUCCESS → оплата; timeout → BANK_REJECT (+ main delete) |
| 8 | SIGN_WAITING + agreementLink |
| 9 | Нет миграции |
| 10 | Kafka completion подкладываем |
| 11 | ≤~3 сек на НТ |
| 12 | Success + timeout + late-message ignore |
