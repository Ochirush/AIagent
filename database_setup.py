from neo4j import GraphDatabase
from datetime import datetime

class IIDoznavatelDB:
    def __init__(self, uri="bolt://localhost:7687", user="neo4j", password="neo4j123"):
        self.driver = GraphDatabase.driver(uri, auth=(user, password))
        print("✅ Подключение к Neo4j установлено")

    def close(self):
        self.driver.close()

    def create_constraints(self):
        with self.driver.session() as session:
            # Основные уникальные идентификаторы
            session.run("CREATE CONSTRAINT person_id IF NOT EXISTS FOR (p:Person) REQUIRE p.id IS UNIQUE")
            session.run("CREATE CONSTRAINT document_id IF NOT EXISTS FOR (d:Document) REQUIRE d.id IS UNIQUE")
            session.run("CREATE CONSTRAINT event_id IF NOT EXISTS FOR (e:Event) REQUIRE e.id IS UNIQUE")
            
            # Индексы для быстрого поиска
            session.run("CREATE INDEX person_name IF NOT EXISTS FOR (p:Person) ON p.full_name")
            session.run("CREATE INDEX document_title IF NOT EXISTS FOR (d:Document) ON d.title")
            session.run("CREATE INDEX event_date IF NOT EXISTS FOR (e:Event) ON e.date")
        print("✅ Ограничения и индексы созданы")

    # === Основные методы ===

    def add_person(self, full_name, role="Person", birth_date=None, address=None, phone=None):
        person_id = f"pers_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        with self.driver.session() as session:
            result = session.run("""
                MERGE (p:Person {id: $id})
                SET p.full_name = $full_name,
                    p.role = $role,
                    p.birth_date = $birth_date,
                    p.address = $address,
                    p.phone = $phone,
                    p.created_at = datetime()
                RETURN p
            """, id=person_id, full_name=full_name, role=role, 
                 birth_date=birth_date, address=address, phone=phone)
            return result.single()

    def add_document(self, title, content, doc_type="Protocol", date=None):
        doc_id = f"doc_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        with self.driver.session() as session:
            result = session.run("""
                CREATE (d:Document {
                    id: $id,
                    title: $title,
                    content: $content,
                    type: $doc_type,
                    date: $date,
                    created_at: datetime()
                })
                RETURN d
            """, id=doc_id, title=title, content=content, doc_type=doc_type, date=date)
            return result.single()

    def add_relationship(self, person1_name, person2_name, rel_type, properties=None):
        if properties is None:
            properties = {}
        with self.driver.session() as session:
            session.run(f"""
                MATCH (p1:Person {{full_name: $p1}})
                MATCH (p2:Person {{full_name: $p2}})
                MERGE (p1)-[r:{rel_type}]->(p2)
                SET r += $props, r.created_at = datetime()
            """, p1=person1_name, p2=person2_name, props=properties)

if __name__ == "__main__":
    db = IIDoznavatelDB(password="password123")  # ← измени, если у тебя другой пароль
    
    db.create_constraints()
    
    # Пример заполнения
    db.add_person("Иванов Иван Иванович", role="Подозреваемый", birth_date="1992-05-15", address="г. Воронеж, ул. Ленина, 17")
    db.add_person("Петров Сергей Александрович", role="Следователь")
    db.add_person("Сидоров Алексей Петрович", role="Потерпевший")

    print("\n🎉 База данных успешно инициализирована!")
    db.close()