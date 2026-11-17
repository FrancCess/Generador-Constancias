; Instalación limpia: si existen datos de una instalación anterior,
; pregunta si se deben borrar (plantillas, historial, correo guardado y ajustes).
; En instalaciones silenciosas (/S, p. ej. actualizaciones) nunca se borra nada.
!ifndef BUILD_UNINSTALLER
Var GcBorrarDatos
Var GcDirDatos

!macro customInit
  StrCpy $GcBorrarDatos "0"
  ReadEnvStr $GcDirDatos APPDATA
  StrCmp $GcDirDatos "" gc_fin_init
  StrCpy $GcDirDatos "$GcDirDatos\${PRODUCT_NAME}"
  IfSilent gc_fin_init
  IfFileExists "$GcDirDatos\*.*" 0 gc_fin_init
    MessageBox MB_YESNO|MB_ICONQUESTION|MB_DEFBUTTON2 "Se encontraron datos de una instalación anterior de ${PRODUCT_NAME}:$\r$\nplantillas, historial, correo guardado y ajustes.$\r$\n$\r$\n¿Quieres una instalación limpia y borrarlos?$\r$\n$\r$\nSí: se borran y el programa abre como nuevo.$\r$\nNo: se conservan (lo mejor si solo estás actualizando)." IDNO gc_fin_init
    StrCpy $GcBorrarDatos "1"
  gc_fin_init:
!macroend

!macro customInstall
  StrCmp $GcBorrarDatos "1" 0 gc_fin_install
  StrCmp $GcDirDatos "" gc_fin_install
    RMDir /r "$GcDirDatos"
  gc_fin_install:
!macroend
!endif
