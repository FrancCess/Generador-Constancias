@echo off
cd /d "%~dp0"
echo Creando el instalador de Generador de Constancias...
if not exist node_modules (
  echo Instalando lo necesario por primera vez...
  call npm install
)
call npm run build
echo.
if exist dist (
  echo Listo. El instalador esta en la carpeta "dist".
  start "" "%~dp0dist"
) else (
  echo Algo salio mal. Revisa los mensajes de arriba.
)
pause
