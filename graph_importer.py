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
        response = requests.post(f"{self.url.rstrip('/')}/api/generate", json={
            "model": self.model, "prompt": prompt, "stream": False,
            "format": "json", "options": {"temperature": 0, "num_ctx": 32768, "num_predict": 8192},
        }, timeout=300)
        response.raise_for_status()
        return parse_llm_json(response.json()["response"])

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
        event_prompt = f"""Повторно проанализируй весь документ, но сосредоточься только на действиях,
местах и времени. Не выдумывай события. Для каждого явно описанного действия создай event, даже если
точное время неизвестно. Для каждого явно указанного адреса или места создай location. Свяжи event с
location через OCCURRED_AT, а с датой/временем через OCCURRED_AT_TIME. Всех названных участников
свяжи с event через PARTICIPATED_IN и укажи properties.event_role. Используй наиболее полное ФИО.

Дело: {case_number}
Документ: {document_name}
Текст:
{text}

Верни только JSON:
{{
  "entities": [{{"type":"person|event|location|time", "name":"краткое уникальное название", "properties":{{"event_type":"тип события", "date":"дата", "time":"время", "description":"что произошло", "topic":"тема разговора", "address":"полный адрес места"}}}}],
  "relations": [{{"from":"точное name", "type":"PARTICIPATED_IN|OCCURRED_AT|OCCURRED_AT_TIME", "to":"точное name", "properties":{{"event_role":"роль участника", "confidence":0.0, "quote":"опора в тексте"}}}}]
}}
Направления обязательны: person -> event, event -> location, event -> time."""
        return self._merge_extractions(self._generate(subject_prompt), self._generate(event_prompt))


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
            source = by_name.get(str(value.get("from", "")).strip().casefold())
            target = by_name.get(str(value.get("to", "")).strip().casefold())
            source_type = source and source["properties"]["entity_type"]
            target_type = target and target["properties"]["entity_type"]
            raw_type = str(value.get("type", "RELATED_TO")).upper()
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
            if entity_type not in NODE_TYPES or not name or (entity_type, name.casefold()) in seen:
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
