"""Импорт судебных документов в граф Neo4j с извлечением сущностей Qwen/Ollama."""

from __future__ import annotations

import argparse
import hashlib
import json
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable
from xml.etree import ElementTree
from zipfile import BadZipFile, ZipFile

import requests
from neo4j import GraphDatabase


NODE_TYPES = {"person": "Person", "case": "Case", "event": "Event", "location": "Location",
              "evidence": "Evidence", "document": "Document", "organization": "Organization", "time": "Time"}
ENTITY_TYPE_ALIASES = {
    "человек": "person", "персона": "person", "лицо": "person",
    "событие": "event", "действие": "event",
    "место": "location", "адрес": "location",
    "организация": "organization", "компания": "organization",
    "время": "time", "дата": "time", "доказательство": "evidence",
}
EVENT_KEYWORDS = {
    "допрос": "допрос", "изби": "избиение", "удар": "нанесение удара",
    "звон": "звонок", "разговор": "разговор", "встреч": "встреча",
    "конфликт": "конфликт", "напад": "нападение", "оскорб": "оскорбление",
    "удерж": "удержание", "задерж": "задержание", "предъявлен": "предъявление",
    "поезд": "поездка", "прибыл": "прибытие", "приех": "прибытие",
    "вызов": "вызов", "обращен": "обращение",
}
RELATION_TYPES = {
    "PARTICIPATED_IN": "УЧАСТВОВАЛ_В",
    "OCCURRED_AT": "ПРОИЗОШЛО_В",
    "OCCURRED_AT_TIME": "ПРОИЗОШЛО_ВО_ВРЕМЯ",
    "HAS_EVIDENCE": "ИМЕЕТ_ДОКАЗАТЕЛЬСТВО",
    "WORKS_AT": "РАБОТАЕТ_В",
    "WORKED_AT": "РАБОТАЛ_В",
    "MARRIED_TO": "СУПРУГ_СУПРУГА",
    "PARENT_OF": "РОДИТЕЛЬ",
    "CHILD_OF": "РЕБЕНОК",
    "RELATIVE_OF": "РОДСТВЕННИК",
    "LIVES_AT": "ПРОЖИВАЕТ_В",
    "KNOWS": "ЗНАКОМ_С",
    "RELATED_TO": "СВЯЗАН_С",
}
RELATION_TYPE_ALIASES = {
    "УЧАСТВОВАЛ_В": "PARTICIPATED_IN", "УЧАСТВОВАЛА_В": "PARTICIPATED_IN",
    "УЧАСТНИК_СОБЫТИЯ": "PARTICIPATED_IN", "ПРОИЗОШЛО_В": "OCCURRED_AT",
    "ПРОИЗОШЛО_ВО_ВРЕМЯ": "OCCURRED_AT_TIME", "РАБОТАЕТ_В": "WORKS_AT",
    "РАБОТАЛ_В": "WORKED_AT", "СУПРУГ_СУПРУГА": "MARRIED_TO",
    "РОДИТЕЛЬ": "PARENT_OF", "РЕБЕНОК": "CHILD_OF", "РОДСТВЕННИК": "RELATIVE_OF",
    "ПРОЖИВАЕТ_В": "LIVES_AT", "ЗНАКОМ_С": "KNOWS", "СВЯЗАН_С": "RELATED_TO",
}
RELATION_ENDPOINTS = {
    "PARTICIPATED_IN": ("person", "event"),
    "OCCURRED_AT": ("event", "location"),
    "OCCURRED_AT_TIME": ("event", "time"),
    "HAS_EVIDENCE": ("event", "evidence"),
    "WORKS_AT": ("person", "organization"),
    "WORKED_AT": ("person", "organization"),
    "MARRIED_TO": ("person", "person"),
    "PARENT_OF": ("person", "person"),
    "CHILD_OF": ("person", "person"),
    "RELATIVE_OF": ("person", "person"),
    "LIVES_AT": ("person", "location"),
    "KNOWS": ("person", "person"),
}
EXTRACTION_SCHEMA = {
    "type": "object",
    "properties": {
        "entities": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "type": {"type": "string"},
                    "name": {"type": "string"},
                    "properties": {"type": "object"},
                },
                "required": ["type", "name", "properties"],
            },
        },
        "relations": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "from": {"type": "string"},
                    "type": {"type": "string"},
                    "to": {"type": "string"},
                    "properties": {"type": "object"},
                },
                "required": ["from", "type", "to", "properties"],
            },
        },
    },
    "required": ["entities", "relations"],
}
W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def stable_id(prefix: str, *parts: str) -> str:
    """Return a reproducible id, so re-importing a file does not duplicate it."""
    value = "|".join(part.strip().casefold() for part in parts if part)
    return f"{prefix}_{hashlib.sha256(value.encode('utf-8')).hexdigest()[:20]}"


def read_document(path: Path) -> str:
    """Read DOCX data even when OCR gave the file a .doc extension; also supports RTF."""
    try:
        with ZipFile(path) as archive:
            root = ElementTree.fromstring(archive.read("word/document.xml"))
        paragraphs = []
        for paragraph in root.iter(W_NS + "p"):
            value = "".join(node.text or "" for node in paragraph.iter(W_NS + "t")).strip()
            if value:
                paragraphs.append(value)
        return "\n".join(paragraphs)
    except (BadZipFile, KeyError, ElementTree.ParseError):
        pass

    raw = path.read_bytes()
    for encoding in ("utf-8", "cp1251", "koi8-r"):
        text = raw.decode(encoding, errors="ignore")
        if text.lstrip().startswith("{\\rtf"):
            text = re.sub(r"\\'[0-9a-fA-F]{2}|\\[a-zA-Z]+-?\d* ?|[{}]", " ", text)
        if re.search(r"[А-Яа-яЁё]", text):
            return re.sub(r"\s+", " ", text).strip()
    raise ValueError(f"Не удалось извлечь текст из {path}")


def extract_case_number(text: str, fallback: str) -> str:
    match = re.search(r"(?:уголовное\s+)?дело\s*№\s*([\w./-]+)", text, re.IGNORECASE)
    # OCR templates sometimes leave the number blank and put "ПРИГОВОР" on the
    # next line.  A case number must contain at least one digit.
    if match and re.search(r"\d", match.group(1)):
        return match.group(1)
    # Имена OCR-файлов обычно начинаются с номера дела. Не превращаем всё имя
    # документа в номер дела, если номер не удалось найти в тексте.
    fallback_match = re.match(r"\s*(\d[\w./-]*)", fallback)
    return fallback_match.group(1) if fallback_match else fallback


def parse_llm_json(response: str) -> dict[str, Any]:
    """Accept Qwen JSON surrounded by Markdown fences, but reject non-object responses."""
    cleaned = re.sub(r"^\s*```(?:json)?\s*|\s*```\s*$", "", response.strip(), flags=re.IGNORECASE)
    start, end = cleaned.find("{"), cleaned.rfind("}")
    if start < 0 or end < start:
        raise ValueError("Qwen did not return a JSON object")
    data = json.loads(cleaned[start:end + 1])
    if not isinstance(data, dict):
        raise ValueError("Qwen response must be a JSON object")
    return data


@dataclass
class QwenExtractor:
    model: str
    url: str

    def _generate(self, prompt: str) -> dict[str, Any]:
        last_error: Exception | None = None
        retry_prompt = prompt
        for attempt in range(2):
            response = requests.post(f"{self.url.rstrip('/')}/api/generate", json={
                "model": self.model, "prompt": retry_prompt, "stream": False,
                "format": EXTRACTION_SCHEMA,
                "options": {"temperature": 0, "num_ctx": 32768, "num_predict": 8192},
            }, timeout=300)
            response.raise_for_status()
            raw_response = response.json()["response"]
            try:
                return parse_llm_json(raw_response)
            except (json.JSONDecodeError, ValueError) as error:
                last_error = error
                retry_prompt = (prompt + "\n\nПРЕДЫДУЩИЙ ОТВЕТ СОДЕРЖАЛ ПОВРЕЖДЕННЫЙ JSON. "
                                "Повтори извлечение. Верни строго один JSON-объект по заданной схеме: "
                                "без Markdown, комментариев и неэкранированных кавычек внутри строк.")
        raise ValueError(f"Qwen дважды вернул некорректный JSON: {last_error}") from last_error

    @staticmethod
    def _merge_extractions(*extractions: dict[str, Any]) -> dict[str, Any]:
        entities, relations, seen_entities, seen_relations = [], [], set(), set()
        for extraction in extractions:
            for entity in extraction.get("entities", []):
                if not isinstance(entity, dict):
                    continue
                key = (str(entity.get("type", "")).casefold(), str(entity.get("name", "")).strip().casefold())
                if key not in seen_entities:
                    seen_entities.add(key)
                    entities.append(entity)
            for relation in extraction.get("relations", []):
                if not isinstance(relation, dict):
                    continue
                key = (str(relation.get("from", "")).strip().casefold(),
                       str(relation.get("type", "")).upper(),
                       str(relation.get("to", "")).strip().casefold())
                if key not in seen_relations:
                    seen_relations.add(key)
                    relations.append(relation)
        return {"entities": entities, "relations": relations}

    @staticmethod
    def _expand_combined_event_locations(extraction: dict[str, Any]) -> dict[str, Any]:
        """Recover events that Qwen incorrectly embeds into people or locations."""
        entities = []
        for value in extraction.get("entities", []):
            if not isinstance(value, dict):
                entities.append(value)
                continue
            copied = dict(value)
            copied["properties"] = dict(value.get("properties", {})) if isinstance(value.get("properties"), dict) else {}
            entities.append(copied)
        relations = list(extraction.get("relations", []))
        additions: list[dict[str, Any]] = []
        recovered_events: dict[tuple[str, str], str] = {}

        # Reuse an event already returned by the focused extraction pass.
        for entity in entities:
            if not isinstance(entity, dict) or str(entity.get("type", "")).casefold() not in {"event", "событие", "действие"}:
                continue
            properties = entity.get("properties", {})
            context = " ".join(str(properties.get(key, "")) for key in
                               ("event_type", "тип", "description", "описание", "topic", "тема"))
            context += " " + str(entity.get("name", ""))
            inferred = next((name for keyword, name in EVENT_KEYWORDS.items() if keyword in context.casefold()), "")
            if inferred:
                recovered_events[(inferred, str(properties.get("date", properties.get("дата", ""))))] = str(entity.get("name"))

        for entity in entities:
            if not isinstance(entity, dict):
                continue
            entity_type = ENTITY_TYPE_ALIASES.get(str(entity.get("type", "")).casefold(),
                                                  str(entity.get("type", "")).casefold())
            if entity_type not in {"person", "location"}:
                continue
            properties = entity.get("properties") if isinstance(entity.get("properties"), dict) else {}
            event_context = " ".join(str(properties.get(key, "")) for key in
                                     ("event_type", "тип", "description", "описание", "topic", "тема")).casefold()
            inferred_type = next((name for keyword, name in EVENT_KEYWORDS.items() if keyword in event_context), "")
            if not inferred_type:
                continue
            entity_name = str(entity.get("name", "")).strip()
            description = str(properties.get("description", properties.get("описание", ""))).strip()
            date = str(properties.get("date", properties.get("дата", ""))).strip()
            event_key = (inferred_type, date)
            event_name = recovered_events.get(event_key)
            if not event_name:
                suffix = entity_name if entity_type == "location" else date
                event_name = f"{inferred_type.capitalize()} — {suffix}" if suffix else inferred_type.capitalize()
                recovered_events[event_key] = event_name
            event_properties = {
                key: properties[key] for key in ("event_type", "date", "time", "description", "topic", "result")
                if key in properties
            }
            event_properties["event_type"] = inferred_type
            additions.append({"type": "event", "name": event_name, "properties": event_properties})
            if entity_type == "location":
                relations.append({"from": event_name, "type": "OCCURRED_AT", "to": entity_name,
                                  "properties": {"confidence": 1.0, "quote": description}})
            else:
                role = "участник"
                if re.search(r"\bизбил[аи]?\b|нан[её]с(?:ла)?\s+удар", description.casefold()):
                    role = "наносил удары"
                elif re.search(r"свидетел|видел|наблюдал", description.casefold()):
                    role = "свидетель"
                elif re.search(r"присутств", description.casefold()):
                    role = "присутствовал"
                relations.append({"from": entity_name, "type": "PARTICIPATED_IN", "to": event_name,
                                  "properties": {"event_role": role, "confidence": 1.0,
                                                 "quote": description}})

            # Эти поля описывают событие, а не человека или место.
            for key in ("event_type", "date", "time", "description", "topic", "result"):
                properties.pop(key, None)
        return QwenExtractor._merge_extractions({"entities": entities + additions, "relations": relations})

    @staticmethod
    def _chunk_text(text: str, limit: int = 4500) -> list[str]:
        chunks, current = [], []
        current_size = 0
        for paragraph in text.splitlines():
            paragraph = paragraph.strip()
            if not paragraph:
                continue
            if current and current_size + len(paragraph) + 1 > limit:
                chunks.append("\n".join(current))
                current, current_size = [], 0
            current.append(paragraph)
            current_size += len(paragraph) + 1
        if current:
            chunks.append("\n".join(current))
        return chunks

    @staticmethod
    def _has_event_participants(extraction: dict[str, Any]) -> bool:
        return any(isinstance(relation, dict) and
                   RELATION_TYPE_ALIASES.get(str(relation.get("type", "")).upper(),
                                             str(relation.get("type", "")).upper()) == "PARTICIPATED_IN"
                   for relation in extraction.get("relations", []))

    def extract(self, text: str, case_number: str, document_name: str) -> dict[str, Any]:
        subject_prompt = f"""Ты извлекаешь факты из материалов уголовного дела. Не выдумывай факты.\n
Дело: {case_number}\nДокумент: {document_name}\nТекст:\n{text}\n\nВерни только JSON следующей структуры:\n{{
  "entities": [{{"type":"person|event|location|evidence|organization|time", "name":"краткое уникальное название", "properties":{{"role":"процессуальная роль person", "event_type":"тип event", "date":"дата event", "time":"время event", "description":"что произошло", "topic":"тема разговора или звонка"}}}}],
  "relations": [{{"from":"точное name первой сущности", "type":"PARTICIPATED_IN|OCCURRED_AT|OCCURRED_AT_TIME|HAS_EVIDENCE|WORKS_AT|WORKED_AT|MARRIED_TO|PARENT_OF|CHILD_OF|RELATIVE_OF|LIVES_AT|KNOWS|RELATED_TO", "to":"точное name второй сущности", "properties":{{"event_role":"роль участника в событии", "confidence":0.0, "quote":"короткая дословная опора"}}}}]
}}\n
Построй предметный граф, а не список упоминаний. Обязательно извлеки ВСЕ явно названные персоны,
организации, места и значимые действия из всего текста. Создай отдельную сущность event для каждого
значимого действия: звонок, разговор, встреча, конфликт, удар, избиение, поездка, задержание, допрос,
предъявление фотографии. Название event должно быть коротким, например «Звонок Пановой Татаринцеву».
В properties события запиши тип, дату, время, описание, тему разговора и результат, когда они известны.
Если звонок прямо указан, звонивший и получатель должны оба вести к одному event через PARTICIPATED_IN,
а event_role должен быть «инициатор звонка» и «получатель звонка». Не создавай звонок по догадке.

Связь PARTICIPATED_IN всегда направляй person -> event. OCCURRED_AT: event -> location.
OCCURRED_AT_TIME: event -> time. Не связывай человека напрямую с датой события. Родство, знакомство,
место жительства и работу связывай напрямую: MARRIED_TO, PARENT_OF/CHILD_OF, RELATIVE_OF, KNOWS,
LIVES_AT, WORKS_AT/WORKED_AT. Работодатель всегда должен быть отдельной organization.
Используй для персоны наиболее полное ФИО из документа во всех entities и relations. Инициалы,
краткое имя и полное ФИО одного человека не должны становиться разными сущностями; варианты имени
запиши в properties.aliases.

Для каждой персоны укажи properties.role, если роль прямо следует
из документа: свидетель, потерпевший, подозреваемый, обвиняемый, подсудимый, следователь, адвокат,
эксперт, переводчик, понятый или иное точное обозначение. Не делай юридический вывод по поведению:
например, человек, которого кто-то ударил, не становится автоматически потерпевшим без прямого
указания в документе. Извлекай также людей с неполным именем или кличкой, помечая это в properties.
Для каждого события укажи участников, место/время и доказательства, если это прямо следует из текста.
Не включай сам документ и дело: они добавляются программой."""
        def build_event_prompt(fragment: str) -> str:
            return f"""Проанализируй текст, но сосредоточься только на действиях,
местах и времени. Не выдумывай события. Для каждого явно описанного действия создай event, даже если
точное время неизвестно. Для каждого явно указанного адреса или места создай location. Свяжи event с
location через OCCURRED_AT, а с датой/временем через OCCURRED_AT_TIME. Всех названных участников
свяжи с event через PARTICIPATED_IN и укажи properties.event_role. Используй наиболее полное ФИО.

Дело: {case_number}
Документ: {document_name}
Текст:
{fragment}

Верни только JSON:
{{
  "entities": [{{"type":"person|event|location|time", "name":"краткое уникальное название", "properties":{{"event_type":"тип события", "date":"дата", "time":"время", "description":"что произошло", "topic":"тема разговора", "address":"полный адрес места"}}}}],
  "relations": [{{"from":"точное name", "type":"PARTICIPATED_IN|OCCURRED_AT|OCCURRED_AT_TIME", "to":"точное name", "properties":{{"event_role":"роль участника", "confidence":0.0, "quote":"опора в тексте"}}}}]
}}
Каждое событие и каждое место обязаны быть ОТДЕЛЬНЫМИ элементами массива entities. Не помещай
event_type, date, time, description или topic в properties сущности location.
Направления обязательны: person -> event, event -> location, event -> time."""

        subject = self._generate(subject_prompt)
        focused = self._generate(build_event_prompt(text))
        merged = self._expand_combined_event_locations(self._merge_extractions(subject, focused))
        if not self._has_event_participants(merged):
            chunk_results = [self._generate(build_event_prompt(chunk)) for chunk in self._chunk_text(text)]
            merged = self._expand_combined_event_locations(self._merge_extractions(merged, *chunk_results))
        if not self._has_event_participants(merged):
            people = [str(entity.get("name")) for entity in merged.get("entities", [])
                      if ENTITY_TYPE_ALIASES.get(str(entity.get("type", "")).casefold(),
                                                 str(entity.get("type", "")).casefold()) == "person"]
            events = [str(entity.get("name")) for entity in merged.get("entities", [])
                      if ENTITY_TYPE_ALIASES.get(str(entity.get("type", "")).casefold(),
                                                 str(entity.get("type", "")).casefold()) == "event"]
            if people and events:
                participant_prompt = f"""Определи участников уже найденных событий по тексту документа.
Не создавай новые сущности. Используй ИСКЛЮЧИТЕЛЬНО точные имена из списков ниже.

Персоны: {json.dumps(people, ensure_ascii=False)}
События: {json.dumps(events, ensure_ascii=False)}
Текст: {text}

Верни JSON {{"entities": [], "relations": [{{"from":"точное имя персоны",
"type":"PARTICIPATED_IN", "to":"точное имя события", "properties":{{"event_role":"конкретная роль",
"confidence":0.0, "quote":"короткая опора в тексте"}}}}]}}.
Не связывай человека с событием, если участие не следует из текста."""
                merged = self._merge_extractions(merged, self._generate(participant_prompt))
        return merged


class Neo4jGraphWriter:
    def __init__(self, uri: str, user: str, password: str):
        self.driver = GraphDatabase.driver(uri, auth=(user, password))

    def close(self) -> None:
        self.driver.close()

    def create_schema(self) -> None:
        with self.driver.session() as session:
            for label in NODE_TYPES.values():
                session.run(f"CREATE CONSTRAINT {label.lower()}_id IF NOT EXISTS FOR (n:{label}) REQUIRE n.id IS UNIQUE")
            session.run("CREATE CONSTRAINT case_number IF NOT EXISTS FOR (c:Case) REQUIRE c.number IS UNIQUE")

    def write(self, source: Path, text: str, case_number: str, extraction: dict[str, Any]) -> None:
        document_id = stable_id("document", str(source.resolve()))
        case_id = stable_id("case", case_number)
        entities = self._normalise_entities(extraction.get("entities", []), case_id, document_id)
        by_name = {entity["name"].casefold(): entity for entity in entities}
        with self.driver.session() as session:
            session.execute_write(self._write_transaction, document_id, case_id, source, text, case_number, entities,
                                  extraction.get("relations", []), by_name)

    @staticmethod
    def _normalised_name(value: str) -> list[str]:
        value = re.sub(r"\([^)]*\)", " ", value.casefold().replace("ё", "е"))
        return re.findall(r"[а-яa-z0-9]+", value)

    @staticmethod
    def _resolve_entity(value: Any, by_name: dict[str, dict[str, Any]]) -> dict[str, Any] | None:
        raw = str(value or "").strip()
        exact = by_name.get(raw.casefold())
        if exact:
            return exact
        query = Neo4jGraphWriter._normalised_name(raw)
        if not query:
            return None
        candidates = []
        for entity in {item["id"]: item for item in by_name.values()}.values():
            candidate = Neo4jGraphWriter._normalised_name(entity["name"])
            if query == candidate:
                candidates.append(entity)
                continue
            entity_type = entity["properties"].get("entity_type")
            if entity_type == "person" and query[0] == candidate[0]:
                initials_match = all(index < len(candidate) and token[0] == candidate[index][0]
                                     for index, token in enumerate(query[1:], start=1))
                if initials_match:
                    candidates.append(entity)
            elif entity_type == "event" and (candidate[:len(query)] == query or query[:len(candidate)] == candidate):
                candidates.append(entity)
        unique = {item["id"]: item for item in candidates}
        return next(iter(unique.values())) if len(unique) == 1 else None

    @staticmethod
    def _write_transaction(tx: Any, document_id: str, case_id: str, source: Path, text: str, case_number: str,
                           entities: list[dict[str, Any]], relations: list[dict[str, Any]], by_name: dict[str, dict[str, Any]]) -> None:
        tx.run("MERGE (c:Case {number:$number}) ON CREATE SET c.id=$id", id=case_id, number=case_number)
        tx.run("MERGE (d:Document {id:$id}) SET d.name=$name, d.path=$path, d.text=$text, d.imported_at=datetime()",
               id=document_id, name=source.name, path=str(source), text=text)
        tx.run("MATCH (d:Document {id:$document_id}),(c:Case {number:$case_number}) MERGE (d)-[:ОТНОСИТСЯ_К_ДЕЛУ]->(c)",
               document_id=document_id, case_number=case_number)
        for entity in entities:
            tx.run(f"MERGE (n:{entity['label']} {{id:$id}}) SET n += $properties", id=entity["id"], properties=entity["properties"])
            tx.run("MATCH (n {id:$entity_id}),(d:Document {id:$document_id}) MERGE (n)-[:УПОМИНАЕТСЯ_В]->(d)",
                   entity_id=entity["id"], document_id=document_id)
            if entity["label"] == "Person":
                tx.run("MATCH (n {id:$entity_id}),(c:Case {number:$case_number}) "
                       "MERGE (n)-[r:УЧАСТВУЕТ_В_ДЕЛЕ]->(c) "
                       "SET r.роль=$role",
                       entity_id=entity["id"], case_number=case_number,
                       role=entity["properties"].get("роль", "не установлена"))
            else:
                tx.run("MATCH (n {id:$entity_id}),(c:Case {number:$case_number}) "
                       "MERGE (n)-[:ОТНОСИТСЯ_К_ДЕЛУ]->(c)",
                       entity_id=entity["id"], case_number=case_number)
        for relation in Neo4jGraphWriter._normalise_relations(relations, by_name):
            tx.run(f"MATCH (a {{id:$from_id}}),(b {{id:$to_id}}) MERGE (a)-[r:{relation['type']}]->(b) SET r += $properties",
                   from_id=relation["from_id"], to_id=relation["to_id"], properties=relation["properties"])

    @staticmethod
    def _normalise_relations(raw: Iterable[Any], by_name: dict[str, dict[str, Any]]) -> list[dict[str, Any]]:
        result = []
        for value in raw:
            if not isinstance(value, dict):
                continue
            source = Neo4jGraphWriter._resolve_entity(value.get("from"), by_name)
            target = Neo4jGraphWriter._resolve_entity(value.get("to"), by_name)
            source_type = source and source["properties"]["entity_type"]
            target_type = target and target["properties"]["entity_type"]
            raw_type = str(value.get("type", "RELATED_TO")).upper()
            raw_type = RELATION_TYPE_ALIASES.get(raw_type, raw_type)
            expected = RELATION_ENDPOINTS.get(raw_type)
            if not source or not target or raw_type not in RELATION_TYPES:
                continue
            if expected and (source_type, target_type) != expected:
                # Qwen иногда разворачивает «событие -> человек». Для участия
                # безопасно исправляем направление; прочие неверные пары отбрасываем.
                if raw_type == "PARTICIPATED_IN" and (source_type, target_type) == ("event", "person"):
                    source, target = target, source
                else:
                    continue
            properties = value.get("properties") if isinstance(value.get("properties"), dict) else {}
            properties = {key: item for key, item in properties.items() if isinstance(item, (str, int, float, bool, list))}
            translated = {
                "роль_в_событии": properties.pop("event_role", properties.pop("role", "")),
                "цитата": properties.pop("quote", ""),
                "уверенность": properties.pop("confidence", 0.0),
            }
            translated.update(properties)
            result.append({"from_id": source["id"], "to_id": target["id"],
                           "type": RELATION_TYPES[raw_type], "properties": translated})
        return result

    @staticmethod
    def _normalise_entities(raw: Iterable[Any], case_id: str, document_id: str) -> list[dict[str, Any]]:
        result, seen = [], set()
        for value in raw:
            if not isinstance(value, dict):
                continue
            entity_type, name = str(value.get("type", "")).lower(), str(value.get("name", "")).strip()
            entity_type = ENTITY_TYPE_ALIASES.get(entity_type, entity_type)
            if entity_type not in NODE_TYPES or not name or (entity_type, name.casefold()) in seen:
                continue
            # Названия вроде ConsNonform — служебные токены Word/LLM, а не
            # события русскоязычного документа.
            if entity_type == "event" and not re.search(r"[А-Яа-яЁё]", name):
                continue
            seen.add((entity_type, name.casefold()))
            properties = value.get("properties") if isinstance(value.get("properties"), dict) else {}
            properties = {key: item for key, item in properties.items() if isinstance(item, (str, int, float, bool, list))}
            if entity_type == "person":
                role = str(properties.pop("role", properties.get("роль", "не установлена"))).strip().lower()
                properties["роль"] = role or "не установлена"
            elif entity_type == "event":
                key_names = {"event_type": "тип", "date": "дата", "time": "время",
                             "description": "описание", "topic": "тема", "result": "результат"}
                properties = {key_names.get(key, key): item for key, item in properties.items()}
            properties.update({"name": name, "entity_type": entity_type})
            # Устойчивые сущности объединяются между документами одного дела.
            # События и доказательства остаются привязаны к документу-источнику.
            id_parts = (case_id, document_id, name) if entity_type in {"event", "evidence"} else (case_id, name)
            result.append({"id": stable_id(entity_type, *id_parts), "name": name,
                           "label": NODE_TYPES[entity_type], "properties": properties})
        return result


def main() -> None:
    parser = argparse.ArgumentParser(description="Импорт документов судебного дела в Neo4j через Qwen/Ollama")
    parser.add_argument("folder", type=Path, nargs="?", default=Path("OCR_Результаты"))
    parser.add_argument("--uri", default="bolt://localhost:7687")
    parser.add_argument("--user", default="neo4j")
    parser.add_argument("--password", required=True, help="Пароль Neo4j (не храните его в коде).")
    parser.add_argument("--model", default="qwen2.5:7b")
    parser.add_argument("--ollama-url", default="http://localhost:11434")
    args = parser.parse_args()
    files = sorted(path for path in args.folder.rglob("*") if path.suffix.lower() in {".doc", ".docx", ".rtf"})
    if not files:
        raise SystemExit(f"Документы в папке {args.folder} не найдены")
    writer, extractor = Neo4jGraphWriter(args.uri, args.user, args.password), QwenExtractor(args.model, args.ollama_url)
    try:
        writer.create_schema()
        for path in files:
            text = read_document(path)
            case_number = extract_case_number(text, path.stem)
            writer.write(path, text, case_number, extractor.extract(text, case_number, path.name))
            print(f"Импортировано: {path.name} (дело {case_number})")
    finally:
        writer.close()


if __name__ == "__main__":
    main()
