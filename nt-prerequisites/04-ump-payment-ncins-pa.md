# НТ пререквизиты :: Worker ump-payment-ncins-pa

**UMP** · 21 июля 2026 · актуализировано по BPMN PR (NCINS)

## Ключевые факты BPMN

- Timer: **PT5M**
- Message: `paymentFinished` (`ump.process.from.system`)
- **correlationKey** = `businessKey + "." + "NON_CREDIT_INSURANCE"`
- Jobs: `ump-payment-ncins-pa.set-hold`, `ump-payment-ncins-pa.create-payment` — **MOCK/DRAFT**
- Delete при timeout — в **main**

## Пререквизиты 1–12

| # | Значение (кратко) |
|---|---|
| 1 | get → mock hold → mock create → Receive paymentFinished → End; EventSubProcess PT5M → BANK_REJECT |
| 2 | Sync несколько сек; message почти мгновенно (+30с); timer PT5M |
| 3 | ~85 / пик ~170 (после signing-timeout) |
| 4–6 | TBD / get + EQ mock + Kafka |
| 6 (full) | Подложить paymentFinished с corr=`<bk>.NON_CREDIT_INSURANCE` |
| 7 | Мок SUCCESS; timeout BANK_REJECT |
| 8 | Существенных изменений UMP DB нет |
| 9 | Нет миграции |
| 10 | EQ мок + Kafka publish |
| 11 | Мок несколько сек; message 0–30с |
| 12 | Success + timeout |
