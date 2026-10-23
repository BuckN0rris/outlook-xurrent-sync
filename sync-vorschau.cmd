@echo off
rem Dry-run: zeigt nur an, was geaendert WUERDE. Schreibt nichts.
rem %~dp0 ist der Ordner dieser Datei, das Projekt darf also umziehen.
chcp 65001 >nul
cd /d "%~dp0"
title Outlook-Xurrent-Sync - Vorschau
python sync.py
echo.
echo ------------------------------------------------------------
echo Vorschau beendet, Exit-Code %ERRORLEVEL%  (0 = ok, 1 = Teilfehler, 2 = Konfiguration)
echo Es wurde nichts geschrieben.
echo ------------------------------------------------------------
pause
