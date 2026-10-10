@echo off
rem Double click this to run the splitter. It just starts the PowerShell script
rem sitting next to it, without changing any settings on the computer.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Split video for chat.ps1" %*
