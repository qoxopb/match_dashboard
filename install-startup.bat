@echo off
powershell.exe -ExecutionPolicy Bypass -Command "Add-Type -AssemblyName System.Windows.Forms; & '%~dp0install-startup.ps1'"
