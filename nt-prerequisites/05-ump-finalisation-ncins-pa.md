# НТ пререквизиты :: Worker ump-finalisation-ncins-pa

**UMP** · 21 июля 2026 · актуализировано по BPMN PR (NCINS)

## Ключевые факты BPMN (финал PR)

```
Start → get-application-data (FINALISATION)
     → ump-finalisation-ncins-pa.create-contract
     → Parallel Gateway
          ├─ ea-send-documents.v1  acId=agreementLink  stopInIncident=false
          ├─ ea-send-documents.v1  acId=policyLink     stopInIncident=false
          └─ ea-send-documents.v1  acId=contractLink   stopInIncident=false
     → Join → End
```

MultiInstance на EA **удалён**. Job create-contract: `ump-finalisation-ncins-pa.create-contract`.

## Пререквизиты 1–12

| # | Значение (кратко) |
|---|---|
| 1 | См. цепочку выше |
| 2 | Sync несколько сек; 3 EA параллельно |
| 3 | ~85/170 заявок; EA connector ×3 → ~255–510/час на тип |
| 4–6 | TBD / get + ins-contracts + 3× ea-send |
| 6 (full) | Уникальный `contractNumber` на прогон; интеграция стенда договоров |
| 7 | get/create SUCCESS\|ERROR; EA response необязателен |
| 8 | UMP DB не меняется |
| 9 | Нет миграции |
| 10 | Мок 3× EA; АБ живой/стенд |
| 11 | Несколько сек |
| 12 | Success / get err / create err / EA без docs без инцидента |
