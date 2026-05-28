import sys
from pathlib import Path
sys.path.append(str(Path(__file__).parent.parent))

from tools import *

def test_scan_all_entities():
    print(scan_all_entities(None))

def test_highlight_entity():
    print(highlight_entity(None, "ABC123", 1))

def test_count_text_patterns():
    print(count_text_patterns(None, "PMC-3M"))

def test_highlight_text_matches():
    print(highlight_text_matches(None, "PMC-3M", 1))

def test_move_entity():
    print(move_entity(None, "ABC123", [0, 0, 0], [100, 100, 0]))

def test_rotate_entity():
    print(rotate_entity(None, "ABC123", [0, 0, 0], 45))

def test_copy_entity():
    print(copy_entity(None, "ABC123", [0, 0, 0], [100, 100, 0]))

def test_get_all_tables():
    print(get_all_tables(None))

def test_get_table_schema():
    print(get_table_schema(None, "cad_elements"))

def test_execute_query():
    print(execute_query(None, "SELECT * FROM cad_elements LIMIT 5"))

def test_query_and_highlight():
    print(query_and_highlight(None, "SELECT handle FROM cad_elements LIMIT 3", 1))


if __name__ == "__main__":
    # test_scan_all_entities()
    # test_highlight_entity()
    # test_count_text_patterns()
    # test_highlight_text_matches()
    # test_move_entity()
    # test_rotate_entity()
    # test_copy_entity()
    test_get_all_tables()
    test_get_table_schema()
    test_execute_query()
    test_query_and_highlight()
