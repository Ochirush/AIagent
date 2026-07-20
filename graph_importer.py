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
    "RELATED_TO": "СВЯЗАН_С",
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

    def extract(self, text: str, case_number: str, document_name: str) -> dict[str, Any]:
        prompt = f"""Ты извлекаешь факты из материалов уголовного дела. Не выдумывай факты.\n
Дело: {case_number}\nДокумент: {document_name}\nТекст:\n{text}\n\nВерни только JSON следующей структуры:\n{{
  "entities": [{{"type":"person|event|location|evidence|organization|time", "name":"строка", "properties":{{"role":"роль лица в деле, только для person"}}}}],
  "relations": [{{"from":"точное name первой сущности", "type":"PARTICIPATED_IN|OCCURRED_AT|OCCURRED_AT_TIME|HAS_EVIDENCE|WORKS_AT|RELATED_TO", "to":"точное name второй сущности", "properties":{{"confidence":0.0, "quote":"короткая опора в документе"}}}}]
}}\n
Обязательно создай отдельную сущность event для каждого значимого действия. Связи PARTICIPATED_IN,
OCCURRED_AT и OCCURRED_AT_TIME должны начинаться или заканчиваться на event: не связывай человека
напрямую с местом или временем. Для каждой персоны укажи properties.role, если роль прямо следует
из документа: свидетель, потерпевший, подозреваемый, обвиняемый, подсудимый, следователь, адвокат,
эксперт, переводчик, понятый или иное точное обозначение. Не делай юридический вывод по поведению:
например, человек, которого кто-то ударил, не становится автоматически потерпевшим без прямого
указания в документе. Извлекай также людей с неполным именем или кличкой, помечая это в properties.
Для каждого события укажи участников, место/время и доказательства, если это прямо следует из текста.
Не включай сам документ и дело: они добавляются программой."""
        response = requests.post(f"{self.url.rstrip('/')}/api/generate", json={
            "model": self.model, "prompt": prompt, "stream": False,
            "format": "json", "options": {"temperature": 0},
        }, timeout=300)
        response.raise_for_status()
        return parse_llm_json(response.json()["response"])


class Neo4jGraphWriter:
    def __init__(self, uri: str, user: str, password: str):
        self.driver = GraphDatabase.driver(uri, auth=(user, password))

    def close(self) -> None:
        self.driver.close()

    def create_schema(self) -> None:
        with self.driver.session() as session:
            for label in NODE_TYPES.values():
                session.run(f"CREATE CONSTRAINT {label.lower()}_id IF NOT EXISTS FOR (n:{label}) REQUIRE n.id IS UNIQUE")

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
        tx.run("MERGE (c:Case {id:$id}) SET c.number=$number", id=case_id, number=case_number)
        tx.run("MERGE (d:Document {id:$id}) SET d.name=$name, d.path=$path, d.text=$text, d.imported_at=datetime()",
               id=document_id, name=source.name, path=str(source), text=text)
        tx.run("MATCH (d:Document {id:$document_id}),(c:Case {id:$case_id}) MERGE (d)-[:ОТНОСИТСЯ_К_ДЕЛУ]->(c)",
               document_id=document_id, case_id=case_id)
        for entity in entities:
            tx.run(f"MERGE (n:{entity['label']} {{id:$id}}) SET n += $properties", id=entity["id"], properties=entity["properties"])
            tx.run("MATCH (n {id:$entity_id}),(d:Document {id:$document_id}) MERGE (n)-[:УПОМИНАЕТСЯ_В]->(d)",
                   entity_id=entity["id"], document_id=document_id)
            if entity["label"] == "Person":
                tx.run("MATCH (n {id:$entity_id}),(c:Case {id:$case_id}) "
                       "MERGE (n)-[r:УЧАСТВУЕТ_В_ДЕЛЕ]->(c) "
                       "SET r.роль=$role",
                       entity_id=entity["id"], case_id=case_id,
                       role=entity["properties"].get("роль", "не установлена"))
            else:
                tx.run("MATCH (n {id:$entity_id}),(c:Case {id:$case_id}) "
                       "MERGE (n)-[:ОТНОСИТСЯ_К_ДЕЛУ]->(c)",
                       entity_id=entity["id"], case_id=case_id)
        for relation in relations:
            rel_type = RELATION_TYPES.get(str(relation.get("type", "RELATED_TO")).upper())
            first, second = by_name.get(str(relation.get("from", "")).casefold()), by_name.get(str(relation.get("to", "")).casefold())
            if first and second and rel_type:
                tx.run(f"MATCH (a {{id:$from_id}}),(b {{id:$to_id}}) MERGE (a)-[r:{rel_type}]->(b) SET r += $properties",
                       from_id=first["id"], to_id=second["id"], properties=relation.get("properties") or {})

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
            properties.update({"name": name, "entity_type": entity_type})
            result.append({"id": stable_id(entity_type, case_id, document_id, name), "name": name,
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
