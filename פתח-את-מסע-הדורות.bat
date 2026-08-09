@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo פותח את מסע הדורות...
start "" http://localhost:4173
npx -y serve . -l 4173
