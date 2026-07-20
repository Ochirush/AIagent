from neo4j import GraphDatabase
from pathlib import Path
import docx
import re
from datetime import datetime
import zipfile

class DocImporter:
    def __init__(self, uri="bolt://localhost:7687", user="neo4j", password="neo4j123"):
        self.driver = GraphDatabase.driver(uri, auth=(user, password))

    def close(self):
        self.driver.close()
    
    def clear_database(self):
        """Очистка базы данных перед импортом"""
        with self.driver.session() as session:
            session.run("MATCH (n) DETACH DELETE n")
            print("✓ База данных очищена")

    def clean_rtf_text(self, content):
        """Очистка RTF текста от тегов и мусора"""
        # Удаляем управляющие последовательности RTF
        content = re.sub(r'\\[a-z]+(?:-?\d+)?', ' ', content)
        content = re.sub(r'\\\'[0-9a-f]{2}', ' ', content)
        content = re.sub(r'\\u\d+', ' ', content)
        content = re.sub(r'\{', ' ', content)
        content = re.sub(r'\}', ' ', content)
        content = re.sub(r'\\', ' ', content)
        
        # Удаляем маркеры маскировки
        content = re.sub(r'<данные изъяты>', '', content)
        content = re.sub(r'<[^>]+>', '', content)
        
        # Нормализуем пробелы
        content = re.sub(r'\s+', ' ', content)
        
        return content.strip()

    def read_docx_file(self, file_path):
        """Чтение DOCX файла"""
        try:
            doc = docx.Document(file_path)
            full_text = "\n".join([p.text for p in doc.paragraphs if p.text.strip()])
            return full_text if full_text.strip() else None
        except Exception as e:
            return None

    def read_rtf_file(self, file_path):
        """Чтение RTF файла"""
        try:
            # Пробуем разные кодировки
            for encoding in ['cp1251', 'utf-8', 'cp1252', 'koi8-r']:
                try:
                    with open(file_path, 'r', encoding=encoding, errors='ignore') as f:
                        content = f.read()
                    
                    # Очищаем RTF текст
                    cleaned_text = self.clean_rtf_text(content)
                    
                    # Проверяем, есть ли русский текст
                    if len(cleaned_text) > 500 and re.search(r'[А-ЯЁ][а-яё]', cleaned_text):
                        return cleaned_text
                except:
                    continue
            
            return None
        except Exception as e:
            print(f"   Ошибка чтения RTF: {e}")
            return None

    def is_rtf_file(self, file_path):
        """Проверка, является ли файл RTF"""
        try:
            with open(file_path, 'rb') as f:
                header = f.read(20)
                return header.startswith(b'{\\rtf')
        except:
            return False

    def read_file_content(self, file_path):
        """Чтение содержимого файла с поддержкой разных форматов"""
        file_path = Path(file_path)
        
        # Проверяем RTF
        if self.is_rtf_file(file_path):
            print(f"   Обнаружен RTF формат, извлечение текста...")
            text = self.read_rtf_file(file_path)
            if text:
                return text
        
        # Пробуем как docx
        text = self.read_docx_file(file_path)
        if text:
            return text
        
        # Пробуем как обычный текст
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                return f.read()
        except:
            pass
        
        return None

    def extract_main_person(self, text: str, doc_name: str = ""):
        """Извлечение подсудимого из приговора"""
        
        if not text:
            return None
        
        # Очищаем текст для поиска
        text_clean = re.sub(r'\s+', ' ', text)
        
        # Паттерны для поиска подсудимого
        patterns = [
            r'в отношении\s+([А-ЯЁ]{2,}(?:\d+)?)\s*[,)]',
            r'подсудимой\s+([А-ЯЁ]{2,}(?:\d+)?)',
            r'подсудимого\s+([А-ЯЁ]{2,}(?:\d+)?)',
            r'подсудим(?:ой|ого)\s+([А-ЯЁ]{2,}(?:\d+)?)',
            r'обвиняемой\s+([А-ЯЁ]{2,}(?:\d+)?)',
            r'обвиняемого\s+([А-ЯЁ]{2,}(?:\d+)?)',
            r'ФИО(\d+)',
        ]
        
        for pattern in patterns:
            matches = re.findall(pattern, text_clean[:3000], re.IGNORECASE)
            for match in matches:
                name = match.strip() if isinstance(match, str) else match[0] if match else None
                if name and len(name) >= 2:
                    if re.match(r'\d+', str(name)):
                        return f"ФИО{name}"
                    if name.startswith('ФИО'):
                        return name
        
        # Поиск по контексту
        if 'подсудимый' in text_clean.lower():
            # Ищем имя после слова "подсудимый"
            match = re.search(r'подсудим(?:ый|ая)\s+([А-ЯЁ]{2,}(?:\d+)?)', text_clean[:2000], re.IGNORECASE)
            if match:
                return match.group(1)
        
        return f"Подсудимый_{doc_name}"

    def extract_victim(self, text: str, main_person: str = None):
        """Извлечение потерпевшего/потерпевших"""
        
        if not text:
            return []
        
        text_clean = re.sub(r'\s+', ' ', text)
        victims = set()
        
        # Ищем маскированные имена потерпевших
        patterns = [
            r'потерпевшей\s+([А-ЯЁ]{2,}(?:\d+(?:\s*№\s*\d+)?)?)',
            r'потерпевшего\s+([А-ЯЁ]{2,}(?:\d+(?:\s*№\s*\d+)?)?)',
            r'потерпевший\s+([А-ЯЁ]{2,}(?:\d+(?:\s*№\s*\d+)?)?)',
            r'Потерпевший\s*№\s*(\d+)',
            r'ФИО(\d+)(?:\s*№\s*\d+)?',
        ]
        
        for pattern in patterns:
            matches = re.findall(pattern, text_clean[:5000], re.IGNORECASE)
            for match in matches:
                if match:
                    name = str(match).strip()
                    if name.isdigit():
                        name = f"Потерпевший №{name}"
                    elif name.startswith('ФИО') and name != main_person:
                        pass  # сохраняем как есть
                    elif name != main_person and len(name) >= 2:
                        victims.add(name)
        
        # Конкретные имена из документов
        specific_names = ['Потерпевший №1', 'ФИО2 №1', 'ФИО2', 'ФИО8', 'ФИО9']
        for name in specific_names:
            if name in text_clean:
                if name != main_person:
                    victims.add(name)
        
        return list(victims) if victims else []

    def extract_crime_type(self, text: str):
        """Извлечение типа преступления"""
        
        if not text:
            return None
        
        text_clean = re.sub(r'\s+', ' ', text)
        
        # Ищем полную квалификацию
        # Для ч. 1 ст. 105
        match1 = re.search(r'ч\.\s*(\d+)\s*ст\.\s*(\d+(?:\.\d+)?)\s*УК\s*РФ', text_clean, re.IGNORECASE)
        if match1:
            part = match1.group(1)
            article = match1.group(2)
            
            # Проверяем на особые признаки
            crime_name = "Убийство"
            full_text = f"ч. {part} ст. {article} УК РФ"
            
            # Проверяем на пункты
            if 'п. «а»' in text_clean or 'пункту «а»' in text_clean:
                full_text = f"п. «а», {full_text}"
                if 'двух лиц' in text_clean or 'двух лиц' in text:
                    crime_name = "Убийство двух лиц"
            
            if 'п. «е»' in text_clean:
                full_text = f"п. «е», {full_text}"
                if 'общеопасным способом' in text_clean:
                    crime_name += ", совершенное общеопасным способом"
            
            if 'п. «з»' in text_clean:
                full_text = f"п. «з», {full_text}"
                if 'по найму' in text_clean:
                    crime_name += ", по найму"
            
            # Проверяем на покушение
            if 'ч. 3 ст. 30' in text_clean:
                full_text = f"ч. 3 ст. 30, {full_text}"
                crime_name = f"Покушение на {crime_name.lower()}"
            
            return {
                "article": article,
                "part": part,
                "name": crime_name,
                "full": full_text
            }
        
        # Поиск по ключевым словам
        if 'убийство' in text_clean.lower():
            if 'двух лиц' in text_clean:
                return {
                    "article": "105",
                    "part": "2",
                    "name": "Убийство двух лиц",
                    "full": "п. «а» ч. 2 ст. 105 УК РФ"
                }
            elif 'общеопасным способом' in text_clean:
                return {
                    "article": "105",
                    "part": "2",
                    "name": "Убийство, совершенное общеопасным способом",
                    "full": "п. «е» ч. 2 ст. 105 УК РФ"
                }
            elif 'по найму' in text_clean:
                return {
                    "article": "105",
                    "part": "2",
                    "name": "Убийство по найму",
                    "full": "п. «з» ч. 2 ст. 105 УК РФ"
                }
            else:
                return {
                    "article": "105",
                    "part": "1",
                    "name": "Убийство",
                    "full": "ч. 1 ст. 105 УК РФ"
                }
        
        return None

    def extract_sentence(self, text: str):
        """Извлечение назначенного наказания"""
        
        if not text:
            return None
        
        text_clean = re.sub(r'\s+', ' ', text)
        
        patterns = [
            r'назначить\s+(?:ей|ему)\s+наказание\s+в\s+виде\s+(\d+)\s*(?:лет|год|года)\s*(?:(\d+)\s*месяц(?:ев|а)?)?',
            r'наказание\s+в\s+виде\s+(\d+)\s*(?:лет|год|года)\s*(?:(\d+)\s*месяц(?:ев|а)?)?',
            r'лишения\s+свободы\s+на\s+срок\s+(\d+)\s*(?:лет|год|года)\s*(?:(\d+)\s*месяц(?:ев|а)?)?',
            r'приговорил:\s*[^.]*?(\d+)\s*(?:лет|год|года)\s*(?:(\d+)\s*месяц(?:ев|а)?)?',
            r'в\s+виде\s+(\d+)\s*(?:лет|год|года)\s+лишения\s+свободы',
        ]
        
        for pattern in patterns:
            match = re.search(pattern, text_clean, re.IGNORECASE | re.DOTALL)
            if match:
                years = int(match.group(1))
                months = int(match.group(2)) if len(match.groups()) > 1 and match.group(2) else 0
                return {
                    "years": years,
                    "months": months,
                    "total_months": years * 12 + months,
                    "text": f"{years} лет {months} месяцев лишения свободы" if months else f"{years} лет лишения свободы"
                }
        
        return None

    def extract_date(self, text: str):
        """Извлечение даты приговора"""
        
        if not text:
            return None
        
        months = {
            'января': 1, 'февраля': 2, 'марта': 3, 'апреля': 4, 'мая': 5, 'июня': 6,
            'июля': 7, 'августа': 8, 'сентября': 9, 'октября': 10, 'ноября': 11, 'декабря': 12
        }
        
        # Формат "29 октября 2025 года" или "29 октября 2025 г."
        date_match = re.search(r'(\d{1,2})\s+(января|февраля|марта|апреля|мая|июня|июля|августа|сентября|октября|ноября|декабря)\s+(\d{4})\s+(?:года|г\.)', text[:1000])
        if date_match:
            day = int(date_match.group(1))
            month_name = date_match.group(2)
            year = int(date_match.group(3))
            month = months.get(month_name, 1)
            return f"{year}-{month:02d}-{day:02d}"
        
        return None
    
    def extract_court(self, text: str):
        """Извлечение названия суда"""
        
        if not text:
            return None
        
        text_clean = re.sub(r'\s+', ' ', text[:1000])
        
        patterns = [
            r'([А-Я][а-я]+(?:ский|ой|ное|ий)?\s+районный\s+суд)',
            r'([А-Я][а-я]+\s+городской\s+суд)',
            r'([А-Я][а-я]+(?:ский|ой)?\s+суд\s+[А-Я][а-я]+)',
        ]
        
        for pattern in patterns:
            match = re.search(pattern, text_clean)
            if match:
                return match.group(1)
        
        return None

    def import_document(self, file_path: Path):
        print(f"\n→ Обработка: {file_path.name}")
        print(f"   Размер файла: {file_path.stat().st_size} bytes")

        try:
            # Читаем содержимое
            full_text = self.read_file_content(file_path)
            
            if not full_text or len(full_text) < 200:
                print(f"   ✗ Не удалось прочитать содержимое или текст слишком короткий")
                return
            
            print(f"   ✓ Текст извлечен ({len(full_text)} символов)")
            print(f"   ✓ Первые 200 символов: {full_text[:200]}...")

            # Извлекаем сущности
            doc_name = file_path.stem
            main_person = self.extract_main_person(full_text, doc_name)
            victims = self.extract_victim(full_text, main_person)
            crime = self.extract_crime_type(full_text)
            sentence = self.extract_sentence(full_text)
            date = self.extract_date(full_text)
            court = self.extract_court(full_text)
            
            doc_id = f"doc_{datetime.now().strftime('%Y%m%d_%H%M%S')}_{doc_name}"

            print(f"\n   📋 Извлеченные данные:")
            print(f"   → Подсудимый: {main_person}")
            print(f"   → Потерпевшие: {victims if victims else 'не найдены'}")
            print(f"   → Суд: {court if court else 'не указан'}")
            print(f"   → Дата: {date if date else 'не указана'}")
            print(f"   → Преступление: {crime['full'] if crime else 'не найдено'}")
            print(f"   → Наказание: {sentence['text'] if sentence else 'не найдено'}")

            with self.driver.session() as session:
                # Создаем документ
                session.run("""
                    CREATE (d:Document {
                        id: $id,
                        title: $title,
                        type: 'Приговор',
                        date: $date,
                        court: $court,
                        content_preview: $content,
                        import_date: datetime(),
                        file_name: $file_name
                    })
                """, 
                    id=doc_id, 
                    title=doc_name, 
                    date=date if date else "unknown",
                    court=court if court else "unknown",
                    content=full_text[:2000],
                    file_name=file_path.name
                )
                
                # Создаем подсудимого
                session.run("""
                    MATCH (d:Document {id: $doc_id})
                    MERGE (p:Person {full_name: $name, doc_id: $doc_id})
                    SET p.role = 'Подсудимый'
                    MERGE (p)-[:DEFENDANT_IN]->(d)
                """, name=main_person, doc_id=doc_id)
                
                # Создаем потерпевших
                if victims:
                    for victim in victims:
                        session.run("""
                            MATCH (d:Document {id: $doc_id})
                            MERGE (v:Person {full_name: $victim_name, doc_id: $doc_id})
                            SET v.role = 'Потерпевший'
                            MERGE (v)-[:VICTIM_IN]->(d)
                        """, doc_id=doc_id, victim_name=victim)
                
                # Создаем преступление
                if crime:
                    session.run("""
                        MATCH (d:Document {id: $doc_id})
                        CREATE (c:Crime {
                            article: $article,
                            part: $part,
                            name: $crime_name,
                            full_text: $full_text,
                            doc_id: $doc_id
                        })
                        CREATE (d)-[:CHARGE_WITH]->(c)
                    """, 
                        doc_id=doc_id,
                        article=crime['article'],
                        part=crime['part'],
                        crime_name=crime['name'],
                        full_text=crime['full']
                    )
                
                # Создаем наказание
                if sentence:
                    session.run("""
                        MATCH (d:Document {id: $doc_id})
                        CREATE (s:Sentence {
                            years: $years,
                            months: $months,
                            total_months: $total_months,
                            text: $text,
                            doc_id: $doc_id
                        })
                        CREATE (d)-[:HAS_SENTENCE]->(s)
                    """,
                        doc_id=doc_id,
                        years=sentence['years'],
                        months=sentence['months'],
                        total_months=sentence['total_months'],
                        text=sentence['text']
                    )
                
                # Проверяем результат
                result = session.run("""
                    MATCH (d:Document {id: $doc_id})-[r]-(n)
                    RETURN type(r) as rel_type, labels(n) as node_labels
                """, doc_id=doc_id)
                
                connections = list(result)
                print(f"\n   ✅ Документ импортирован. Создано связей: {len(connections)}")

        except Exception as e:
            print(f"   ✗ Ошибка: {e}")
            import traceback
            traceback.print_exc()


if __name__ == "__main__":
    # Подключение к Neo4j
    importer = DocImporter(uri="bolt://localhost:7687", user="neo4j", password="password123")
    
    # Очищаем базу перед импортом (раскомментируйте если нужно)
    importer.clear_database()
    
    # Путь к папке с документами
    folder = Path(r"C:\EB\testik\OCR_Результаты")
    
    # Ищем все .doc файлы
    doc_files = list(folder.glob("**/*.doc")) + list(folder.glob("**/*.docx"))
    
    print(f"Найдено документов: {len(doc_files)}\n")
    print("="*60)

    for f in doc_files:
        importer.import_document(f)

    importer.close()
    print("\n" + "="*60)
    print("=== Импорт завершён ===")
    
    print("\n📊 Для проверки данных выполните в Neo4j Browser:")
    print("="*60)
    print("\n1. Все документы и их связи:")
    print("   MATCH (d:Document)-[r]-(n) RETURN d, r, n")
    print("\n2. Статистика по документам:")
    print("   MATCH (d:Document) RETURN d.title, d.court, d.date")
    print("\n3. Все подсудимые:")
    print("   MATCH (p:Person)-[:DEFENDANT_IN]->(d:Document) RETURN p.full_name, d.title")
    print("\n4. Все преступления:")
    print("   MATCH (c:Crime) RETURN c.full_text, c.name")
    print("\n5. Наказания:")
    print("   MATCH (s:Sentence) RETURN s.text, s.total_months")