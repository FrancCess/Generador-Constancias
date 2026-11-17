@echo off
cd /d "%~dp0"
if not exist node_modules (
  echo Instalando por primera vez, espera unos minutos...
  call npm install
)
call npm start
