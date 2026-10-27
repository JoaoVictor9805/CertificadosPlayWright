@echo off
title RoboSocietario - Atualizar Certificados
cd /d "%~dp0"

if exist "RoboSocietario.exe" (
    RoboSocietario.exe --atualizar-certificados
) else (
    python main.py --atualizar-certificados
)
