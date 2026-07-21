# НТ пререквизиты :: Worker ump-delete-documents-ncins-pa

**UMP** · 21 июля 2026 · актуализировано по BPMN PR (NCINS)

**BPMN process id:** `ump-delete-documents-ncins-pa`  
**Алиас в архиве/Confluence:** `ump-delete-documents-ncins-pa-ncins-pa`

## Ключевые факты BPMN

```
Start → get-application-data (processType=DELETE_DOCS)
     → delete-documents (multiInstance sequential over documents)
     → End
```

- multiInstance: `inputCollection=documents`, `inputElement=document`, `outputCollection=acDocuments`, `outputElement=documentResponse`
- Вызов только Call Activity из **main** (signing/payment XOR timeout + PT30M)
- Таска сейчас **моковая**

## Пререквизиты 1–12

| # | Значение (кратко) |
|---|---|
| 1 | См. цепочку; источники: PT25M signing / PT5M payment / PT30M main |
| 2 | Sync несколько сек × N documents |
| 3 | ~5% потока ≈4–5/час (~95% из timeout — signing) |
| 4–6 | TBD / get + AC deleteDocsh |
| 6 (full) | Call Activity; documents из заявки; syscode=UMP |
| 7 | `acDocuments` ≈ `[{"status":"1"}]` |
| 8 | UMP DB без изменений |
| 9 | Нет миграции |
| 10 | Мок на НТ (нет папки AC) |
| 11 | Несколько сек × N |
| 12 | Timeout-ветки → delete mock → BANK_REJECT на оркестраторе |
