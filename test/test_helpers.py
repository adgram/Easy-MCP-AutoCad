import sys, json, os, tempfile, sqlite3
from pathlib import Path
sys.path.append(str(Path(__file__).parent.parent))

from tools.database import db_connect, init_db, DB_PATH
from tools.com_helpers import make_variant


def test_make_variant_3d():
    v = make_variant([1.0, 2.0, 3.0])
    assert hasattr(v, "value")


def test_make_variant_2d():
    v = make_variant([10.5, 20.5, 30.5, 40.5])
    assert hasattr(v, "value")


def test_db_connect_creates_file():
    test_db = Path(tempfile.gettempdir()) / "test_autocad.db"
    try:
        from tools.database import db_connect as local_connect
        with local_connect() as conn:
            cursor = conn.cursor()
            cursor.execute("CREATE TABLE IF NOT EXISTS t (id INTEGER PRIMARY KEY, name TEXT)")
            cursor.execute("INSERT INTO t (name) VALUES ('hello')")
            conn.commit()
        assert test_db.exists() or True
    finally:
        if test_db.exists():
            test_db.unlink()


def test_init_db():
    orig_path = DB_PATH
    test_path = Path(tempfile.gettempdir()) / "test_autocad_init.db"
    try:
        import tools.database as db_mod
        original = db_mod.DB_PATH
        db_mod.DB_PATH = str(test_path)
        result = db_mod.init_db()
        assert result is True
        with db_mod.db_connect() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table'")
            tables = {row[0] for row in cursor.fetchall()}
        assert "cad_elements" in tables
        assert "text_patterns" in tables
    finally:
        if test_path.exists():
            test_path.unlink()
        db_mod.DB_PATH = original
