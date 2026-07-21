import unittest
from pathlib import Path

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


if __name__ == "__main__":
    unittest.main()
