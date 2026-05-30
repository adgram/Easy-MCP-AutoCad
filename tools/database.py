import os
import sqlite3
from pathlib import Path
from contextlib import contextmanager

_DEFAULT_DB = str(Path(__file__).parent.parent / "autocad_data.db")
_CWD_DB = str(Path(os.getcwd()) / "autocad_data.db")

def _get_db_path():
    """项目级配置（有 .opencode 目录）时用 CWD，否则用包目录"""
    if (Path(os.getcwd()) / ".opencode").exists():
        return _CWD_DB
    return _DEFAULT_DB

DB_PATH = _get_db_path()


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
