@echo off
rem Schreibt die Kategorien tatsaechlich. Siehe sync-vorschau.cmd fuer den Trockenlauf.
rem %~dp0 ist der Ordner dieser Datei, das Projekt darf also umziehen.
chcp 65001 >nul
cd /d "%~dp0"
title Outlook-Xurrent-Sync - anwenden
python sync.py --apply
echo.
echo ------------------------------------------------------------
echo Beendet, Exit-Code %ERRORLEVEL%  (0 = ok, 1 = Teilfehler, 2 = Konfiguration)
echo Protokoll: logs\daily\
echo ------------------------------------------------------------
pause
