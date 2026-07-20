import unittest
from pathlib import Path

from graph_importer import Neo4jGraphWriter, extract_case_number, parse_llm_json, read_document


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


if __name__ == "__main__":
    unittest.main()
