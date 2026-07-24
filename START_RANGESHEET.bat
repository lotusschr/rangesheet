@echo off
rem ==========================================================================
rem  RangeSheet - double-click this file to start the program.
rem
rem  This file is intentionally ASCII-only and does almost nothing.
rem  All real logic (and all Thai on-screen text) lives in launcher.ps1.
rem
rem  Why the split: cmd.exe re-reads a .bat file by byte offset. If the file
rem  mixes UTF-8 Thai text with a "chcp" call, those offsets stop matching and
rem  cmd starts slicing commands in half ("SystemExit" -> "ystemExit").
rem  PowerShell reads UTF-8 correctly, so the Thai messages go there.
rem ==========================================================================

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0launcher.ps1"

if errorlevel 1 (
    echo.
    pause
)
