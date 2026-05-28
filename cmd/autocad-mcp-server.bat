@echo off
REM AutoCAD MCP Server Launcher
REM This script launches the AutoCAD MCP Server

cd /d "%~dp0"
python "[.\Lib\site-packages]\easy_mcp_autocad\server.py" %*