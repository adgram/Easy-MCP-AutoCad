import sys
from pathlib import Path
sys.path.append(str(Path(__file__).parent.parent))

try:
    import win32com.client
    HAS_AUTOCAD = True
except ImportError:
    HAS_AUTOCAD = False

if HAS_AUTOCAD:
    try:
        acad = win32com.client.Dispatch("AutoCAD.Application.25")
        acad.Documents.Count
        AUTOCAD_RUNNING = True
    except:
        AUTOCAD_RUNNING = False
else:
    AUTOCAD_RUNNING = False


def test_import_server():
    from server import mcp
    assert mcp is not None
    tools = mcp._tool_manager._tools
    assert len(tools) == 28

def test_all_tool_names():
    from server import mcp
    names = set(mcp._tool_manager._tools.keys())
    expected = {
        "create_new_drawing", "draw_line", "draw_circle", "draw_polyline",
        "draw_rectangle", "draw_text", "draw_device_connection", "create_layer",
        "autocad_get_layers", "autocad_get_blocks", "autocad_get_linetypes",
        "scan_all_entities", "highlight_entity", "count_text_patterns",
        "highlight_text_matches", "move_entity", "rotate_entity", "copy_entity",
        "get_all_tables", "get_table_schema", "execute_query", "query_and_highlight",
        "execute_command", "create_block", "autocad_select_objects",
        "autocad_eval_lisp", "autocad_get_sysvar", "autocad_set_sysvar",
    }
    assert names == expected


def test_import_drawing():
    from tools.drawing import (
        create_new_drawing, draw_line, draw_circle, draw_polyline,
        draw_rectangle, draw_text, draw_device_connection, create_layer,
        autocad_get_layers, autocad_get_blocks, autocad_get_linetypes,
    )
    assert all(callable(f) for f in [
        create_new_drawing, draw_line, draw_circle, draw_polyline,
        draw_rectangle, draw_text, draw_device_connection, create_layer,
        autocad_get_layers, autocad_get_blocks, autocad_get_linetypes,
    ])


def test_import_entities():
    from tools.entities import (
        scan_all_entities, highlight_entity, count_text_patterns,
        highlight_text_matches, move_entity, rotate_entity, copy_entity,
        get_all_tables, get_table_schema, execute_query, query_and_highlight,
    )
    assert all(callable(f) for f in [
        scan_all_entities, highlight_entity, count_text_patterns,
        highlight_text_matches, move_entity, rotate_entity, copy_entity,
        get_all_tables, get_table_schema, execute_query, query_and_highlight,
    ])


def test_import_lisp():
    from tools.lisp import (
        execute_command, create_block, autocad_select_objects,
        autocad_eval_lisp, autocad_get_sysvar, autocad_set_sysvar,
    )
    assert all(callable(f) for f in [
        execute_command, create_block, autocad_select_objects,
        autocad_eval_lisp, autocad_get_sysvar, autocad_set_sysvar,
    ])


def test_import_com_helpers():
    from tools.com_helpers import get_acad, make_variant, wait_cmd, set_active_layer
    assert callable(get_acad)
    assert callable(make_variant)
    assert callable(wait_cmd)
    assert callable(set_active_layer)


def test_import_database():
    from tools.database import db_connect, init_db
    assert callable(db_connect)
    assert callable(init_db)
