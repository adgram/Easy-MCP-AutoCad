from mcp.server.fastmcp import FastMCP
from tools.database import init_db

mcp = FastMCP("AutoCAD-DB-Server")
init_db()

from tools import drawing, entities, lisp
drawing.register_tools(mcp)
entities.register_tools(mcp)
lisp.register_tools(mcp)

if __name__ == "__main__":
    mcp.run()
