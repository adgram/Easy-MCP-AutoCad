from mcp.server.fastmcp import FastMCP
from tools.database import init_db

mcp = FastMCP("AutoCAD-DB-Server")
init_db()

from .tools import *

for func in [
    create_new_drawing, draw_line, draw_circle, draw_polyline,
    draw_rectangle, draw_text, create_layer,
    autocad_get_layers, autocad_get_blocks, autocad_get_linetypes,
    scan_all_entities, highlight_entity, count_text_patterns,
    highlight_text_matches, move_entity, rotate_entity, copy_entity,
    get_all_tables, get_table_schema, execute_query, query_and_highlight,
    execute_command, create_block, autocad_select_objects,
    autocad_eval_lisp, autocad_get_sysvar, autocad_set_sysvar,
]:
    mcp.tool()(func)

if __name__ == "__main__":
    mcp.run()
