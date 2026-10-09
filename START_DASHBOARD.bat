@echo off
cd /d "%~dp0"
start "" "http://localhost:8765/index.html"
py -m http.server 8765
