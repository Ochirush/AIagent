import unittest
from pathlib import Path
from unittest.mock import Mock, patch

from graph_importer import Neo4jGraphWriter, QwenExtractor, extract_case_number, parse_llm_json, read_document


class GraphImporterTests(unittest.TestCase):
    def test_reads_docx_payload_with_doc_extension(self):
        text = read_document(Path("OCR_Результаты/3.doc"))
        self.assertIn("Дело № 1-108/2025", text)

    def test_extracts_case_number(self):
        self.assertEqual(extract_case_number("Дело № 1-108/2025", "fallback"), "1-108/2025")
        self.assertEqual(extract_case_number("Дело №\nПРИГОВОР", "fallback"), "fallback")
        self.assertEqual(extract_case_number("номер в тексте отсутствует", "152320033 Протокол допроса"), "152320033")

    def test_parses_json_in_markdown_fence(self):
        self.assertEqual(parse_llm_json("```json\n{\"entities\": []}\n```"), {"entities": []})

    def test_normalises_only_supported_entities(self):
        entities = Neo4jGraphWriter._normalise_entities([
            {"type": "person", "name": "Панова Н.А.", "properties": {"role": "свидетель"}},
            {"type": "unknown", "name": "не попадет"},
        ], "case", "document")
        self.assertEqual(entities[0]["label"], "Person")
        self.assertEqual(entities[0]["properties"]["роль"], "свидетель")
        self.assertEqual(len(entities), 1)

    def test_normalises_subject_relations_and_event_direction(self):
        entities = Neo4jGraphWriter._normalise_entities([
            {"type": "person", "name": "Панова"},
            {"type": "event", "name": "Звонок", "properties": {"topic": "происшествие"}},
            {"type": "organization", "name": "АД Пластик"},
        ], "case", "document")
        by_name = {entity["name"].casefold(): entity for entity in entities}
        relations = Neo4jGraphWriter._normalise_relations([
            {"from": "Звонок", "type": "PARTICIPATED_IN", "to": "Панова",
             "properties": {"event_role": "инициатор звонка"}},
            {"from": "Панова", "type": "WORKED_AT", "to": "АД Пластик"},
            {"from": "Панова", "type": "OCCURRED_AT_TIME", "to": "Звонок"},
        ], by_name)
        self.assertEqual([item["type"] for item in relations], ["УЧАСТВОВАЛ_В", "РАБОТАЛ_В"])
        self.assertEqual(relations[0]["from_id"], by_name["панова"]["id"])
        self.assertEqual(relations[0]["properties"]["роль_в_событии"], "инициатор звонка")

    def test_people_are_merged_across_documents_but_events_are_not(self):
        raw = [{"type": "person", "name": "Панова Наталья Александровна"},
               {"type": "event", "name": "Допрос Пановой"}]
        first = Neo4jGraphWriter._normalise_entities(raw, "case", "document-1")
        second = Neo4jGraphWriter._normalise_entities(raw, "case", "document-2")
        self.assertEqual(first[0]["id"], second[0]["id"])
        self.assertNotEqual(first[1]["id"], second[1]["id"])

    def test_merges_focused_event_pass_without_duplicates(self):
        merged = QwenExtractor._merge_extractions(
            {"entities": [{"type": "person", "name": "Панова"}], "relations": []},
            {"entities": [{"type": "person", "name": "Панова"},
                          {"type": "event", "name": "Избиение"},
                          {"type": "location", "name": "ул. Плеханова"}],
             "relations": [{"from": "Избиение", "type": "OCCURRED_AT", "to": "ул. Плеханова"}]},
        )
        self.assertEqual(len(merged["entities"]), 3)
        self.assertEqual(len(merged["relations"]), 1)

    @patch("graph_importer.requests.post")
    def test_retries_malformed_llm_json(self, post):
        malformed = Mock()
        malformed.raise_for_status.return_value = None
        malformed.json.return_value = {"response": '{"entities": [}' }
        valid = Mock()
        valid.raise_for_status.return_value = None
        valid.json.return_value = {"response": '{"entities": [], "relations": []}' }
        post.side_effect = [malformed, valid]
        result = QwenExtractor("qwen", "http://ollama")._generate("prompt")
        self.assertEqual(result, {"entities": [], "relations": []})
        self.assertEqual(post.call_count, 2)
        self.assertIsInstance(post.call_args.kwargs["json"]["format"], dict)

    def test_splits_event_properties_out_of_location(self):
        extraction = {"entities": [{
            "type": "location", "name": "кабинет №301",
            "properties": {"event_type": "допрос", "date": "05.03.2015",
                           "description": "место проведения допроса"},
        }], "relations": []}
        expanded = QwenExtractor._expand_combined_event_locations(extraction)
        self.assertEqual([item["type"] for item in expanded["entities"]], ["location", "event"])
        self.assertEqual(expanded["relations"][0]["type"], "OCCURRED_AT")
        self.assertEqual(expanded["relations"][0]["to"], "кабинет №301")

    def test_does_not_turn_static_location_category_into_event(self):
        extraction = {"entities": [{
            "type": "location", "name": "квартира №12",
            "properties": {"event_type": "жилище", "description": "место проживания Пановой"},
        }], "relations": []}
        expanded = QwenExtractor._expand_combined_event_locations(extraction)
        self.assertEqual(len(expanded["entities"]), 1)
        self.assertEqual(expanded["relations"], [])

    def test_recovers_action_name_from_location_description(self):
        extraction = {"entities": [{
            "type": "location", "name": "дом №44",
            "properties": {"event_type": "место событий", "description": "место нахождения избитого мужчины"},
        }], "relations": []}
        expanded = QwenExtractor._expand_combined_event_locations(extraction)
        event = next(item for item in expanded["entities"] if item["type"] == "event")
        self.assertEqual(event["properties"]["event_type"], "избиение")

    def test_recovers_one_shared_event_from_person_properties(self):
        extraction = {"entities": [
            {"type": "person", "name": "Татаринцев Дмитрий", "properties": {
                "event_type": "избиение", "date": "2015-02-11", "description": "Избил мужчину."}},
            {"type": "person", "name": "Леонов Юрий", "properties": {
                "event_type": "избиение", "date": "2015-02-11", "description": "Присутствовал при избиении."}},
        ], "relations": []}
        expanded = QwenExtractor._expand_combined_event_locations(extraction)
        events = [item for item in expanded["entities"] if item["type"] == "event"]
        participation = [item for item in expanded["relations"] if item["type"] == "PARTICIPATED_IN"]
        self.assertEqual(len(events), 1)
        self.assertEqual(len(participation), 2)
        self.assertEqual({item["to"] for item in participation}, {events[0]["name"]})
        self.assertNotIn("event_type", expanded["entities"][0]["properties"])

    def test_rejects_non_cyrillic_event_artifact(self):
        entities = Neo4jGraphWriter._normalise_entities(
            [{"type": "event", "name": "ConsNonform", "properties": {}}], "case", "document")
        self.assertEqual(entities, [])

    def test_accepts_russian_participation_relation_and_resolves_initials(self):
        entities = Neo4jGraphWriter._normalise_entities([
            {"type": "person", "name": "Панова Наталья Александровна"},
            {"type": "event", "name": "Допрос — кабинет №301"},
        ], "case", "document")
        by_name = {entity["name"].casefold(): entity for entity in entities}
        relations = Neo4jGraphWriter._normalise_relations([{
            "from": "Панова Н.А.", "type": "УЧАСТВОВАЛА_В", "to": "Допрос",
            "properties": {"event_role": "свидетель"},
        }], by_name)
        self.assertEqual(len(relations), 1)
        self.assertEqual(relations[0]["type"], "УЧАСТВОВАЛ_В")
        self.assertEqual(relations[0]["properties"]["роль_в_событии"], "свидетель")

    def test_accepts_russian_entity_type_alias(self):
        entities = Neo4jGraphWriter._normalise_entities(
            [{"type": "событие", "name": "Допрос", "properties": {}}], "case", "document")
        self.assertEqual(entities[0]["label"], "Event")

    def test_chunks_long_text_on_paragraph_boundaries(self):
        chunks = QwenExtractor._chunk_text("первый абзац\nвторой абзац\nтретий абзац", limit=30)
        self.assertEqual(chunks, ["первый абзац\nвторой абзац", "третий абзац"])


if __name__ == "__main__":
    unittest.main()
