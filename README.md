# ИИ-помощник дознавателя: граф материалов дела

`graph_importer.py` читает OCR-документы из `OCR_Результаты`, включая DOCX-файлы с расширением `.doc`, отправляет текст в локальный Qwen через Ollama и записывает доказуемые факты в Neo4j.

## Графовая модель

Создаются узлы `Person`, `Case`, `Event`, `Location`, `Evidence`, `Document`, `Organization`, `Time`. Каждый извлечённый узел связан с исходным `Document` отношением `MENTIONED_IN` и с `Case` отношением `BELONGS_TO_CASE`. Qwen может добавить связи `PARTICIPATED_IN`, `OCCURRED_AT`, `OCCURRED_AT_TIME`, `HAS_EVIDENCE`, `WORKS_AT` и `RELATED_TO`; в свойства связи сохраняются `confidence` и короткая цитата-опора.

`Document.text` сохраняет полный текст для проверки оператором. Результат Qwen не считается установленным фактом без проверки первоисточника.

## Запуск

```bash
python -m pip install -r requirements.txt
docker compose up -d
ollama pull qwen2.5:7b
python graph_importer.py OCR_Результаты --password password123
```

Номера дел извлекаются из текста (`Дело № …`); если номера нет, используется имя файла. Идентификаторы детерминированы, поэтому повторный импорт файла не создаёт дублей.

## Проверка

```bash
python -m unittest discover -s tests -v
```
