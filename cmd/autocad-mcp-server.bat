@echo off
REM AutoCAD MCP Server Launcher
REM Run as a module so relative imports in tools/ work

python -m easy_mcp_autocad.server %*