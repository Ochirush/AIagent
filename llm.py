import json
from pathlib import Path
from neo4j import GraphDatabase
import docx
import re
from datetime import datetime
import requests
import time
import subprocess
import tempfile

class QwenDocImporter:
    def __init__(self, uri="bolt://localhost:7687", user="neo4j", password="neo4j123", 
                 model="qwen2.5:7b", ollama_url="http://localhost:11434"):
        self.driver = GraphDatabase.driver(uri, auth=(user, password))
        self.model = model
        self.ollama_url = ollama_url
        
    def close(self):
        self.driver.close()
    
    def clear_database(self):
        with self.driver.session() as session:
            session.run("MATCH (n) DETACH DELETE n")
            print("✓ База данных очищена")
    
    def read_docx_file(self, file_path):
        try:
            doc = docx.Document(file_path)
            full_text = "\n".join([p.text for p in doc.paragraphs if p.text.strip()])
            return full_text if full_text.strip() else None
        except:
            return None
    
    def read_rtf_with_powerpoint(self, file_path):
        """Попытка прочитать RTF через PowerShell (Windows)"""
        try:
            # Используем PowerShell для конвертации RTF в текст
            ps_script = f'''
            Add-Type -AssemblyName System.Windows.Forms
            $rtf = Get-Content -Path "{file_path}" -Raw
            $tb = New-Object System.Windows.Forms.RichTextBox
            $tb.Rtf = $rtf
            $tb.Text
            '''
            result = subprocess.run(
                ['powershell', '-Command', ps_script],
                capture_output=True,
                text=True,
                encoding='utf-8'
            )
            if result.stdout and len(result.stdout) > 100:
                return result.stdout
        except:
            pass
        return None
    
    def read_rtf_manual(self, file_path):
        """Ручная очистка RTF от мусора"""
        try:
            with open(file_path, 'r', encoding='cp1251', errors='ignore') as f:
                content = f.read()
            
            # Удаляем управляющие последовательности
            content = re.sub(r'\\[a-z]+(?:\-?\d+)?', ' ', content)
            content = re.sub(r'\\\'[0-9a-f]{2}', ' ', content)
            content = re.sub(r'\\u\d+', ' ', content)
            content = re.sub(r'\{', ' ', content)
            content = re.sub(r'\}', ' ', content)
            content = re.sub(r'\\', ' ', content)
            
            # Удаляем маркеры шрифтов и таблиц
            content = re.sub(r'\*\\[a-z]+', ' ', content)
            content = re.sub(r'\\[a-z]+\s+[A-Za-z\s]+;', ' ', content)
            
            # Удаляем теги маскировки
            content = re.sub(r'<[^>]+>', ' ', content)
            
            # Удаляем лишние пробелы
            content = re.sub(r'\s+', ' ', content)
            
            # Ищем русский текст (обычно после очистки остается нормальный текст)
            russian_text = re.findall(r'[А-ЯЁ][а-яё\s\.\,\;\:\!\?\-]{20,}', content)
            if russian_text:
                return ' '.join(russian_text)
            
            return None
        except Exception as e:
            return None
    
    def read_file_content(self, file_path):
        """Чтение файла с улучшенной обработкой RTF"""
        
        # Сначала пробуем как DOCX
        text = self.read_docx_file(file_path)
        if text:
            return text
        
        # Пробуем через PowerShell (Windows)
        text = self.read_rtf_with_powerpoint(file_path)
        if text:
            print(f"   ✓ RTF прочитан через PowerShell")
            return text
        
        # Ручная очистка RTF
        text = self.read_rtf_manual(file_path)
        if text:
            print(f"   ✓ RTF прочитан ручной очисткой")
            return text
        
        return None
    
    def check_ollama(self):
        try:
            response = requests.get(f"{self.ollama_url}/api/tags", timeout=5)
            return response.status_code == 200
        except:
            return False
    
    def safe_int(self, value, default=0):
        if value is None:
            return default
        if isinstance(value, (int, float)):
            return int(value)
        if isinstance(value, str):
            digits = re.findall(r'\d+', value)
            return int(digits[0]) if digits else default
        return default
    
    def safe_str(self, value, default=""):
        return default if value is None else str(value)
    
    def extract_with_qwen(self, text: str, doc_name: str = ""):
        """Извлечение сущностей через Qwen 2.5"""
        
        # Ограничиваем текст
        text_preview = text[:8000] if len(text) > 8000 else text
        
        prompt = f"""Ты судебный аналитик. Проанализируй приговор и извлеки информацию в JSON.

Текст приговора:
{text_preview}

Верни ТОЛЬКО JSON (без пояснений) в таком формате:
{{
    "defendant": {{
        "full_name": "ФИО подсудимого или null",
        "birth_date": null,
        "characteristics": []
    }},
    "victims": [
        {{
            "full_name": "ФИО потерпевшего или null",
            "injuries": "повреждения или null"
        }}
    ],
    "crime": {{
        "article": "номер статьи",
        "part": "часть",
        "paragraphs": [],
        "attempt": false,
        "description": null
    }},
    "sentence": {{
        "years": 0,
        "months": 0,
        "regime": null,
        "additional": null
    }},
    "court": {{
        "full_name": null,
        "date": null,
        "judge": null
    }},
    "case_details": {{
        "date_of_crime": null,
        "location": null,
        "method": null
    }},
    "witnesses": [],
    "evidence": []
}}

Используй null для отсутствующих значений, 0 для чисел, [] для массивов.
"""
        
        try:
            response = requests.post(
                f"{self.ollama_url}/api/generate",
                json={
                    "model": self.model,
                    "prompt": prompt,
                    "stream": False,
                    "options": {
                        "temperature": 0.1,
                        "num_predict": 2048
                    }
                },
                timeout=180
            )
            
            if response.status_code == 200:
                result = response.json()
                text_response = result['response']
                
                json_match = re.search(r'\{.*\}', text_response, re.DOTALL)
                if json_match:
                    return json.loads(json_match.group())
            return None
        except Exception as e:
            print(f"   Ошибка AI: {e}")
            return None
    
    def import_document(self, file_path: Path):
        print(f"\n{'='*60}")
        print(f"📄 Обработка: {file_path.name}")
        
        try:
            if not self.check_ollama():
                print(f"   ✗ Ollama не запущен!")
                return
            
            full_text = self.read_file_content(file_path)
            
            # Если не удалось прочитать - пропускаем
            if not full_text:
                print(f"   ✗ Не удалось прочитать файл (возможно поврежденный RTF)")
                print(f"   Создаю минимальную запись в БД")
                
                # Создаем минимальную запись
                doc_id = f"doc_{datetime.now().strftime('%Y%m%d_%H%M%S')}_{file_path.stem}"
                with self.driver.session() as session:
                    session.run("""
                        CREATE (d:Document {
                            id: $id,
                            title: $title,
                            type: 'Приговор',
                            status: 'error_reading',
                            import_date: datetime(),
                            file_name: $file_name
                        })
                    """, 
                        id=doc_id,
                        title=file_path.stem,
                        file_name=file_path.name
                    )
                return
            
            print(f"   ✓ Текст загружен ({len(full_text)} символов)")
            
            # Показываем начало текста для отладки
            preview = full_text[:300].replace('\n', ' ')
            print(f"   📝 Начало текста: {preview}...")
            
            print(f"   🤖 Анализирую с помощью Qwen...")
            
            start_time = time.time()
            entities = self.extract_with_qwen(full_text, file_path.stem)
            elapsed = time.time() - start_time
            
            if not entities:
                print(f"   ✗ AI не смог извлечь данные")
                return
            
            print(f"   ✓ Обработано за {elapsed:.1f} секунд")
            
            # Выводим результаты
            print(f"\n   📋 Результаты:")
            
            defendant = entities.get('defendant', {})
            defendant_name = self.safe_str(defendant.get('full_name'))
            print(f"   👤 Подсудимый: {defendant_name if defendant_name else 'не найден'}")
            
            victims = entities.get('victims', [])
            if victims:
                victim_names = [v.get('full_name') for v in victims if v.get('full_name')]
                print(f"   🎯 Потерпевшие: {', '.join(victim_names) if victim_names else 'не найдены'}")
            
            crime = entities.get('crime', {})
            article = crime.get('article')
            if article:
                print(f"   ⚖️ Статья: ст.{article} УК РФ")
            
            sentence = entities.get('sentence', {})
            years = self.safe_int(sentence.get('years'))
            months = self.safe_int(sentence.get('months'))
            if years > 0:
                print(f"   🔒 Наказание: {years} лет {months} месяцев")
            
            court = entities.get('court', {})
            court_name = self.safe_str(court.get('full_name'))
            if court_name:
                print(f"   🏛️ Суд: {court_name}")
            
            # Сохраняем в Neo4j
            doc_id = f"doc_{datetime.now().strftime('%Y%m%d_%H%M%S')}_{file_path.stem}"
            
            with self.driver.session() as session:
                # Создаем документ
                session.run("""
                    CREATE (d:Document {
                        id: $id,
                        title: $title,
                        type: 'Приговор',
                        court: $court,
                        date: $date,
                        judge: $judge,
                        content_preview: $content,
                        ai_model: $model,
                        import_date: datetime(),
                        file_name: $file_name
                    })
                """, 
                    id=doc_id,
                    title=file_path.stem,
                    court=self.safe_str(court.get('full_name'), 'unknown'),
                    date=self.safe_str(court.get('date'), 'unknown'),
                    judge=self.safe_str(court.get('judge'), 'unknown'),
                    content=full_text[:2000],
                    model=self.model,
                    file_name=file_path.name
                )
                
                # Подсудимый
                if defendant_name:
                    session.run("""
                        MATCH (d:Document {id: $doc_id})
                        CREATE (p:Person {
                            full_name: $name,
                            role: 'Подсудимый',
                            birth_date: $birth_date,
                            characteristics: $chars
                        })
                        CREATE (p)-[:DEFENDANT_IN]->(d)
                    """,
                        doc_id=doc_id,
                        name=defendant_name,
                        birth_date=self.safe_str(defendant.get('birth_date')),
                        chars=defendant.get('characteristics', [])
                    )
                
                # Потерпевшие
                for victim in victims:
                    victim_name = self.safe_str(victim.get('full_name'))
                    if victim_name:
                        session.run("""
                            MATCH (d:Document {id: $doc_id})
                            CREATE (v:Person {
                                full_name: $name,
                                role: 'Потерпевший',
                                injuries: $injuries
                            })
                            CREATE (v)-[:VICTIM_IN]->(d)
                        """,
                            doc_id=doc_id,
                            name=victim_name,
                            injuries=self.safe_str(victim.get('injuries'))
                        )
                
                # Преступление
                if article:
                    session.run("""
                        MATCH (d:Document {id: $doc_id})
                        CREATE (c:Crime {
                            article: $article,
                            part: $part,
                            paragraphs: $paragraphs,
                            attempt: $attempt,
                            description: $description
                        })
                        CREATE (d)-[:CHARGE_WITH]->(c)
                    """,
                        doc_id=doc_id,
                        article=self.safe_str(article),
                        part=self.safe_str(crime.get('part')),
                        paragraphs=crime.get('paragraphs', []),
                        attempt=crime.get('attempt', False),
                        description=self.safe_str(crime.get('description'))
                    )
                
                # Наказание
                if years > 0 or months > 0:
                    total_months = years * 12 + months
                    session.run("""
                        MATCH (d:Document {id: $doc_id})
                        CREATE (s:Sentence {
                            years: $years,
                            months: $months,
                            total_months: $total_months,
                            regime: $regime,
                            text: $text
                        })
                        CREATE (d)-[:HAS_SENTENCE]->(s)
                    """,
                        doc_id=doc_id,
                        years=years,
                        months=months,
                        total_months=total_months,
                        regime=self.safe_str(sentence.get('regime')),
                        text=f"{years} лет {months} месяцев"
                    )
                
                print(f"\n   ✅ Сохранено в Neo4j")
                
                # Сохраняем JSON для отладки
                debug_file = Path(f"debug_{file_path.stem}.json")
                with open(debug_file, 'w', encoding='utf-8') as f:
                    json.dump(entities, f, ensure_ascii=False, indent=2)
                print(f"   📁 Отладка: {debug_file}")
                
        except Exception as e:
            print(f"   ✗ Ошибка: {e}")


if __name__ == "__main__":
    print("="*60)
    print("⚖️ Импорт судебных приговоров с Qwen 2.5")
    print("="*60)
    
    importer = QwenDocImporter(
        uri="bolt://localhost:7687",
        user="neo4j",
        password="password123",
        model="qwen2.5:7b"
    )
    
    importer.clear_database()
    
    folder = Path(r"C:\EB\testik\OCR_Результаты")
    doc_files = list(folder.glob("**/*.doc")) + list(folder.glob("**/*.docx"))
    
    print(f"\n📁 Найдено документов: {len(doc_files)}")
    print(f"🤖 Модель: {importer.model}\n")
    
    for idx, f in enumerate(doc_files, 1):
        print(f"\n[{idx}/{len(doc_files)}]")
        importer.import_document(f)
    
    importer.close()
    print("\n" + "="*60)
    print("✅ Импорт завершён!")
    print("="*60)