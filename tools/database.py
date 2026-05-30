import sqlite3
from pathlib import Path
from contextlib import contextmanager

DB_PATH = str(Path(__file__).parent.parent / "autocad_data.db")


@contextmanager
def db_connect():
    conn = sqlite3.connect(DB_PATH)
    try:
        yield conn
    finally:
        conn.close()


def init_db():
    try:
        with db_connect() as conn:
            cursor = conn.cursor()
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS cad_elements (
                    id INTEGER PRIMARY KEY,
                    handle TEXT UNIQUE,
                    name TEXT NOT NULL,
                    type TEXT NOT NULL,
                    layer TEXT,
                    properties TEXT
                )
            ''')
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS text_patterns (
                    id INTEGER PRIMARY KEY,
                    pattern TEXT UNIQUE,
                    count INTEGER DEFAULT 0,
                    drawing TEXT
                )
            ''')
            conn.commit()
        return True
    except Exception as e:
        print(f"数据库初始化失败: {str(e)}")
        return False
