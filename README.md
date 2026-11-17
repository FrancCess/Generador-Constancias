# 🎓 Generador de Constancias / Diplomas en PDF — v3.1

Programa de escritorio (Electron) para **generar constancias en PDF de forma masiva** a partir de una lista de alumnos en CSV y una imagen de fondo.

## 🛠️ v3.1 — Revisión de errores y validación de datos

**Lectura de listas**
- Excel con columnas de título repetido ya no pierde datos («Nombre», «Nombre 2»); celdas con error (#N/A) se ignoran;
  códigos con formato «#,##0» salen sin comas; rangos gigantes declarados por algunos programas ya no traban el programa.
- La columna «ID» de Microsoft Forms ya no se toma como código de estudiante (se respeta la prioridad Código › Matrícula › Clave…).
- CSV con comillas sin cerrar o comas de más: ahora avisa en qué filas.
- Si no hay columna «Nombre», se elige mejor (nunca evento, folio, fecha, apellidos o una columna de puros números).
- Detección de duplicados mucho más rápida (miles de filas) y completa el código si faltaba.

**Folios, historial y verificación**
- El folio guardado en el historial siempre es el que se imprimió (antes podía diferir si cambiaba la columna Folio).
- Una persona registrada sin código y después con código conserva su folio y su código de verificación.
- «Inicia en 0» ya funciona; los dígitos del folio se limitan a 1–10.
- Restaurar un respaldo recalcula las claves (no se repiten folios).
- En ZIP, el historial se guarda solo cuando el archivo ya se descargó; en carpeta, aunque se cancele o falle a la mitad.
- Aviso si el historial no se pudo guardar; el programa ya no se abre dos veces al mismo tiempo.

**Antes de generar se revisa**
- QR en modo página web sin dirección, QR con demasiado texto, firmas que no cargaron y
  párrafos con {datos} que no existen en la lista (saldrían en blanco).

**Plantillas, fondo y proyectos**
- Una plantilla dañada ya no deja el diseño a medias; se ignoran datos desconocidos, máximo 100 textos,
  y el nombre del alumno nunca puede quedar reemplazado por otro campo.
- Fondos gigantes (más de 6000 px) se reducen; el fondo se guarda en la base de datos local (ya no se pierde por tamaño).
- Con un proyecto abierto, «Plantilla nueva» y «Limpiar lista» preguntan si es otro evento (el proyecto queda intacto).
- Abrir un proyecto avisa si se va a reemplazar trabajo sin guardar; nombres de proyecto repetidos piden confirmación.

**Correo**
- El límite diario de Gmail ahora sí pausa el envío (antes se confundía con «destinatario rechazado»).
- Enviar a una sola persona avisa si ya se le había enviado y no permite doble clic.
- Se marcan correos bien escritos pero con dominio sospechoso (gmial.com, hotmail.con…) con su arreglo sugerido.
- La prueba de correo de una constancia escrita a mano ya lleva su PDF; el reporte incluye envíos de días anteriores.

**Otros**
- Nombres de archivo: «François» → «Francois» (antes «Franois»); nombres reservados de Windows (CON, NUL…) se corrigen.
- En la lista, Enter/↑/↓ ya pasan a la siguiente fila; no se permite que dos filas sean la misma persona.
- Atajos Ctrl+G / Ctrl+O no actúan mientras hay una ventana abierta.
- Varias correcciones de fuentes importadas (nombres repetidos, medidas al cambiar el archivo, familias sin estilo normal).
- El PDF temporal de impresión se borra al cerrar la ventana.

## ✨ Novedades de la v3.0 — Una sola interfaz completa + Express

- Ahora hay **dos modos**: **Completo** y **Express** (botones arriba a la derecha).
- El modo **Completo** une la **cinta única** con la **Mesa de trabajo**:
  - Arriba: plantilla (Nueva · Abrir · Guardar), **Proyecto**, **buscador Ctrl+F**, Completo/Express y Tutorial.
  - La **cinta** con todas las herramientas (Alumnos, Hoja, Fuente, Párrafo y posición, Generar, Correo, Más).
  - A la izquierda: **tu flujo** (Diseño → Personas → Revisar → Entregar con ✓ / !) y los textos de la constancia.
  - Al centro: la hoja con la persona actual (◀ ▶ y botones rápidos) y la **lista siempre visible** abajo, editable,
    con pestañas *Personas · Revisar · Envíos*. «↕ Tamaño» la agranda, la oculta o la deja normal.
- Se quitaron la cinta con 4 pestañas y la Mesa de trabajo como pantallas separadas (todo quedó en el modo Completo).
- El **Modo Express** se conserva igual.

## ✨ Novedades de la v2.9 — Mesa de trabajo

- **Mesa de trabajo** (cuarta interfaz, botón «Mesa de trabajo» arriba): todo en una sola pantalla.
  - **Tu flujo** a la izquierda: Diseño → Personas → Revisar → Entregar, cada paso con ✓ o !; plantillas recientes.
  - La **hoja** con una persona real de la lista y botones rápidos para esa persona (enviar, PDF).
  - La **lista siempre visible** abajo, como Excel: editar en la celda (Enter / ↑ ↓), pestañas *Personas · Revisar · Envíos*,
    estado de cada persona (lista, revisar, emitida, enviada) y folio.
  - El **texto elegido** a la derecha, con colores tomados de tu fondo.
  - **Proyectos**: cada evento guarda su diseño, su lista y su avance; se guarda solo mientras trabajas
    y se recupera al volver a abrir el programa.
- En **todas** las interfaces:
  - **Buscar acciones y personas con Ctrl+F** («imprimir», «correo», «QR», o el nombre de alguien). (Ctrl+K sigue siendo cursiva.)
  - **Colores de tu fondo**: muestras de color tomadas de la imagen para el texto con un clic.
  - **Arreglos sugeridos** en la lista: «Corregir a juan@alumnos.udg.mx» (coma por punto, gmial → gmail…) y
    «→ Juan Pérez» para nombres en mayúsculas; un clic y queda.

## ✨ Novedades de la v2.8 — Asistente de diseño y más

- **Asistente de diseño** (junto a «Imagen de fondo»): crea el fondo sin otro programa. Elige hoja, uno de 8 marcos
  (clásico, ornamental, banda lateral, ondas, moderno, geométrico, minimalista, con medalla), colores (6 paletas o los tuyos),
  hasta 3 logos, institución, título con 7 tipos de letra, subtítulo y hasta 3 firmas con cargo. Genera la imagen a 200 ppp.
- **Imágenes encima del fondo** (Texto → Insertar → «Imagen (firma/sello)»): firmas escaneadas, sellos o logos (PNG transparente)
  que se arrastran y cambian de tamaño. Se guardan en la plantilla.
- **Guardar como imagen**: en Generar, «🖼️ Imágenes» crea un PNG o JPG por persona (100, 150 o 220 ppp) para WhatsApp o redes.
- **Envíos que se retoman**: si Gmail llega a su límite o se va el internet, el envío se pausa y te explica cómo seguir.
  «Solo a quienes aún no se les ha enviado» también omite a quien ya lo recibió en días anteriores (historial).
- **Nombres casi iguales**: la revisión de la lista avisa de «Jose Perez» / «José Pérez», una letra de diferencia o
  palabras en otro orden dentro del mismo evento («¿es la misma persona?»).
- **Aviso de versión nueva**: con tu repositorio de GitHub en `package.json → actualizaciones.github` (o clic en el
  número de versión, arriba a la derecha) el programa revisa una vez al día si hay un Release más nuevo.
  Para publicar: en GitHub → Releases → «Draft a new release», etiqueta `v2.8.0`, adjunta el instalador .exe.

## ✨ Novedades de la v2.7 — Código QR e historial

- **Código QR de verificación** (Texto → Insertar → «Código QR»): cada constancia lleva un QR y un
  **código de verificación** único (ej. `3D9Z-DEJL`) que no se puede adivinar. Dos modos:
  - **Datos de la constancia**: al escanearlo, cualquier celular muestra folio, nombre, evento, fecha y código (sin internet).
  - **Página web de verificación**: el QR abre tu página (GitHub Pages) que dice si es válida.
    La página está en `pagina_verificacion/` (instrucciones en LEEME.txt).
- **Historial de constancias emitidas** (botón «Historial» junto a Imprimir): cada constancia generada,
  impresa o enviada queda registrada con su folio y código. Desde ahí:
  buscar, **verificar un código**, volver a generar el PDF (mismo folio), reenviar por correo, agregar a la lista,
  exportar a Excel, exportar `registro.json` para la página web, respaldo y restaurar.
- **Folios que no se repiten**: quien ya recibió constancia de un evento conserva su folio; las personas nuevas
  siguen después del último folio emitido (casilla en el campo Folio, activada por defecto).

## ✨ Novedades de la v2.6 — Plantillas, párrafos y nombres

- **Plantillas a la mano:** botones **Nueva · Abrir · Guardar** en la barra de título (y en Diseño → Plantilla; en Express: «Diseño nuevo»).
  - **Nueva** limpia el diseño (fondo y textos) sin borrar tu lista; antes pregunta si quieres guardar.
  - **Guardar** pide un nombre, descarga el archivo .json y la deja en **Recientes**.
  - **Abrir** muestra tus plantillas recientes con miniatura (o busca un archivo .json). `Ctrl+O` también.
- **Párrafo con datos** (Texto → Insertar): escribe un texto como
  *«por su participación en el {Taller}, con duración de {Horas} horas»*. Cada persona recibe sus datos,
  los renglones se acomodan solos en el ancho elegido y los datos pueden ir en **negrita**.
  Chips para insertar datos, aviso si una {variable} no existe, interlineado y máximo de renglones (reduce la letra si se pasa).
- **Corregir nombres** (Ver / editar lista → «Aa Corregir nombres»): «JUAN PEREZ DE LA CRUZ» → «Juan Perez de la Cruz»;
  respeta partículas (de, del, la…) y nombres con mezcla intencional (McDonald). Sugiere acentos comunes
  (Perez → Pérez) **solo si los marcas**. Se puede deshacer.

## ✨ Novedades de la v2.5 — Cargar datos más fácil

- **Signo ❓ junto a "Cargar CSV / Excel"** (y en el paso 2 del Modo Express): al pasar el mouse muestra
  **cómo debe ser la tabla** según los textos que tiene la constancia en ese momento (si activas Folio o
  Fecha, aparecen esas columnas). Indica cuáles son obligatorias, recomendadas u opcionales, y trae
  **Descargar plantilla con estas columnas** (CSV listo para llenar).
- **Pegar tabla** (opción extra, el archivo sigue funcionando igual): copia las celdas en Excel o Google Sheets
  (`Ctrl+C`) y pégalas (`Ctrl+V`) en la tabla del programa, o escribe directo en las celdas.
  - Si copias los títulos, la tabla toma esas columnas; bajo cada título se ve qué dato reconoció (✓ Nombre, ✓ Correo…).
  - Marca en color las filas sin nombre y los correos con error antes de agregarlos.
  - `Enter`/`↑`/`↓` para moverse como en Excel; ＋ Fila, ＋ Columna, quitar filas/columnas.
  - Al agregar, sigue el mismo camino que un archivo: duplicados, revisión de la lista, etc.
- El tutorial incluye un paso nuevo sobre **Pegar tabla** y el signo ❓.

## ✨ Novedades de la v2.4 — Generación masiva

Botón **Generar** → ventana con las opciones de salida (se recuerdan):

| Opción | Elecciones |
|---|---|
| **Formato** | Un PDF por constancia · Un solo PDF con todas (una por página) |
| **Guardar en** | Una carpeta que tú eliges (se escribe archivo por archivo, ideal para cientos) · Archivo ZIP |
| **Organizar por** | Sin organizar · Evento/taller · **Grupo** · **Carrera** · cualquier otra columna del CSV/Excel. Con “un PDF por constancia” crea una carpeta por valor; con “un solo PDF” crea un PDF por valor |
| **Imagen de fondo** | JPEG (más rápido) · PNG (imagen exacta). Muestra el peso aproximado de cada una para elegir |

- **Excel**: además de CSV se pueden cargar `.xlsx`, `.xls` y `.ods`. Si el libro tiene varias hojas sin columna Evento, el nombre de cada hoja se usa como evento. La lista también se exporta a Excel.
- **Imprimir…**: la constancia actual, todas las de la lista (o del evento filtrado) o un rango (de la 5 a la 20). Abre la vista de impresión para elegir impresora y copias.
- La barra de progreso muestra el **tiempo restante**.
- Para organizar por grupo o carrera, agrega columnas `Grupo` y/o `Carrera` a tu archivo (el CSV de ejemplo ya las trae).
- Lector de Excel: SheetJS (licencia Apache 2.0) en `librerias/`.

### ⚡ Modo Express (versión rápida)

Para cuando hay prisa o se necesita algo sencillo. Botón **Modo Express** en la barra de arriba (y **Modo avanzado** para regresar; el modo avanzado siempre está disponible y el programa abre en el último modo usado).

1. **Diseño:** elegir la imagen (o abrir un diseño guardado).
2. **Personas:** cargar Excel/CSV o escribir una a mano; ver y corregir la lista.
3. **Textos:** marcar los textos, arrastrarlos y cambiar letra (8 fuentes), tamaño, color y negrita.
4. **¡Listo!:** Generar (un PDF por persona o todos en uno, en una carpeta; si hay varios talleres, se separan solos), Imprimir o Correo.

- Usa el mismo motor que el modo avanzado: el diseño, la lista y la cuenta de correo son los mismos al cambiar de modo.
- Tiene su propio tutorial de 7 pasos (botón Tutorial o F1).
- Archivo: `index_express.html`.

### Tutorial de primer uso

- La primera vez que se abre el programa aparece un **recorrido guiado** de 14 pasos que resalta cada botón y explica con palabras sencillas qué hace (pensado para personas con poca experiencia en computadora).
- Se puede ver **cuantas veces se quiera** con el botón **Tutorial** de la barra de arriba o con la tecla **F1**.
- Se maneja con los botones, con las flechas ← → del teclado, y se cierra con **Esc**. Abre solo la pestaña que corresponde a cada paso.

### Correcciones de seguridad y errores (revisión de la v2.4)

- La ventana ya no puede abrir otras páginas ni archivos soltados encima (solo las del programa); la ventana de impresión no tiene acceso a la computadora.
- Al exportar a CSV/Excel, los textos que empiezan con `=`, `+`, `-` o `@` se guardan con apóstrofo para que Excel no los ejecute como fórmula.
- Las plantillas `.json` se validan (números, colores, ids) y solo aceptan imágenes incrustadas.
- Excel: los códigos largos ya no salen como `2.19E+13` y las fechas se leen como “5 de octubre de 2026”.
- La contraseña solo pierde los espacios con Gmail.
- El instalador incluye la cinta única.
- El folio automático empieza desde el número inicial en cada evento.
- Nombres de archivo limitados a 120 caracteres; un correo con coma o punto y coma se marca como inválido; “Deshacer” reinserta en su lugar.

## ✨ Novedades de la v2.3 — Interfaz tipo Word

- **Cinta con pestañas** como Word: **Inicio** (alumnos, evento, revisar, generar), **Diseño** (fondo, hoja, calibración, plantilla), **Texto** (fuente, estilos, párrafo, auto-ajuste, posición, insertar) y **Correo**.
- Botones grandes para lo principal y botones pequeños **apilados** para lo secundario, con iconos.
- Al hacer clic en un texto de la constancia se abre la pestaña **Texto**.
- **Panel lateral** con los campos y **barra de estado** con navegación entre alumnos, totales y tamaño de hoja.
- Atajos: `Ctrl+N` negrita, `Ctrl+K` cursiva, `Ctrl+S` subrayado, `Ctrl+G` guardar plantilla, `Ctrl+O` abrir plantilla.
- **Dos estilos de cinta**, se cambian con el botón de la barra de título (el programa recuerda cuál usaste):
  - **Cinta con pestañas** (`index.html`): Inicio, Diseño, Texto y Correo.
  - **Cinta única** (`index_una_pestana.html`): todo en una sola pestaña *General* (Alumnos, Hoja, Fuente, Párrafo y posición, Generar, Correo y Más). La calibración se abre desde **Más → Calibración** y el filtro de evento está en la barra de estado.
- La interfaz anterior (barras laterales) sigue en la carpeta `Generador_constancias_v2.2`.
- **Una constancia** (Inicio → Alumnos): escribe nombre, código, evento, folio y fecha a mano; la hoja se actualiza mientras escribes y puedes generar el PDF o agregarla a la lista. No necesita CSV.
- **Revisión al cargar un CSV**: se abre la lista y se marcan con ⚠️ las filas con posibles errores (nombre de una sola palabra, números o símbolos en el nombre, sin código, correo mal escrito, mismo nombre con otro código). Se corrigen ahí mismo; se puede desactivar.
- Iconos: Lucide (licencia ISC).

## ✨ Novedades de la v2.2

### Hoja horizontal o vertical
En **🖼️ Plantilla** eliges:

- **Tamaño de la hoja**: *Según la imagen* (como antes), **Carta**, **Oficio** o **A4**.
- **Orientación**: ▭ Horizontal o ▯ Vertical.
- **Imagen dentro de la hoja** (solo con tamaño fijo): *Estirar*, *Rellenar* (recorta lo que sobre) o *Ajustar sin deformar*.

Con *Según la imagen*, una imagen vertical produce una hoja vertical (21.6 cm de ancho) automáticamente.

### Varios talleres en un solo CSV
Cada fila es **un alumno en un evento**. Columnas que se reconocen:

| Columna | Nombres aceptados | Qué hace |
|---|---|---|
| Evento | Evento, Taller, Curso, Actividad, Diplomado, Seminario… | Se imprime en el campo **Evento / Taller** de cada constancia |
| Folio | Folio, No. folio | Se usa ese folio; si está vacío, se numera solo |
| Fecha | Fecha, Periodo | Se usa la fecha de cada fila; si está vacía, la fija |
| Código | Código, Matrícula, Clave… | Igual que antes |
| Otras | Horas, Instructor, Calificación… | Se pueden imprimir con **＋ Columna del CSV** y usar en el correo como `{Horas}` |

- Un alumno puede repetirse en varios eventos; solo se omite si se repite **en el mismo evento**.
- El selector **Evento / taller** permite trabajar con un solo evento (vista previa, ZIP y correo).
- El ZIP puede llevar **una carpeta por evento**.
- En el correo, con *“Un solo correo por alumno”*, a quien tomó varios talleres le llega **un correo con todas sus constancias**.
- Botón **📄 Descargar CSV de ejemplo** para ver el formato.

```csv
Codigo,Nombre,Correo,Evento,Folio,Fecha,Horas
219512345,Juan Pérez García,juan@alumnos.udg.mx,Taller de Git,GIT-001,5 de octubre de 2026,10
219512345,Juan Pérez García,juan@alumnos.udg.mx,Curso de Python,PY-001,12 de octubre de 2026,20
```

## ✨ Novedades de la v2.1

| Función | Descripción |
|---|---|
| **Envío por correo** | Botón **✉️ Enviar por correo**: manda a cada alumno su constancia en PDF como adjunto, al correo de la columna `Correo` / `Email` del CSV (o escrito a mano en la lista). Asunto y mensaje personalizables con `{nombre}`, `{codigo}`, `{correo}`, `{folio}`, `{evento}`, `{fecha}` y cualquier columna como `{Horas}`. |
| **Control del envío** | Probar conexión, enviarte una prueba, pausa entre correos, cancelar, estado ✔/✖ por alumno, reintentar solo los que fallaron y reporte CSV. |
| **25 fuentes** | Las 7 originales + Helvetica y Courier + 18 de Google Fonts (Montserrat, Poppins, Roboto, Open Sans, Lato, Oswald, Playfair Display, Lora, Merriweather, EB Garamond, Cormorant Garamond, Cinzel, Great Vibes, Pinyon Script, Allura, Alex Brush, Parisienne, Dancing Script). |
| **Estilos** | **Negrita**, *cursiva*, subrayado, sombra (color, distancia y opacidad) y espaciado entre letras, por campo. |
| **Importar fuentes** | Botón **🔤 Fuentes** → importa cualquier `.ttf` de tu computadora (por ejemplo `C:\Windows\Fonts\calibrib.ttf` para tener Calibri en negrita real). |

### Configurar el correo

1. **Gmail** (recomendado): activa la verificación en 2 pasos y crea una **contraseña de aplicación** en <https://myaccount.google.com/apppasswords>. Escribe tu correo y esa contraseña (16 letras). Límite aproximado: 500 correos al día.
2. **Outlook / Office 365**: algunas instituciones bloquean el SMTP; si la prueba falla, pide a sistemas que habiliten “SMTP autenticado” o usa Gmail.
3. **Otro servidor**: escribe servidor, puerto y seguridad.
4. Presiona **🔌 Probar conexión** y luego **🧪 Enviarme una prueba** antes de enviar a todos.

La contraseña solo se guarda si marcas “Recordar”, y se guarda **cifrada** con la protección de tu usuario de Windows.

### Negrita y cursiva “reales” o “simuladas”

Si la fuente tiene el archivo de negrita/cursiva se usa el real; si no, el programa lo **simula** (contorno para negrita, inclinación para cursiva). En **🔤 Fuentes** se ve qué estilos reales tiene cada familia (N, K, NK).

## ✨ Novedades de la v2.0

| Función | Descripción |
|---|---|
| **Código de estudiante** | Se lee del CSV (columna `Código`, `Matrícula`, `Clave`, `No. control`, `Boleta`…) y se puede imprimir en la constancia con un texto antes (ej. `Código: `). Si un alumno no tiene código, ese campo no se imprime. |
| **Varios campos de texto** | Nombre, Código, Curso/Evento, Fecha, Folio consecutivo y **textos libres** ilimitados. Cada uno con su fuente, tamaño, color, alineación, mayúsculas y posición. |
| **Vista previa real** | Botones ◀ ▶ (o `Re Pág` / `Av Pág`) para ver la constancia de cada alumno. El botón **🔍 Nombre más largo** muestra el caso más difícil. |
| **Editor de la lista** | Ver, buscar, editar, agregar, eliminar (con *Deshacer*), ordenar A→Z y exportar la lista a CSV. |
| **Plantillas** | 💾 Guardar / 📂 Abrir el diseño completo (imagen + campos + calibración) en un archivo `.json`. El último diseño se recuerda automáticamente. |
| **Auto-ajuste visible** | “Ancho máximo (mm)” por campo: si el texto no cabe, la letra se reduce sola. La guía roja muestra el área y la vista previa refleja el tamaño real. |
| **Alineación** | Izquierda, centro o derecha por campo. |
| **PDF de la vista actual** | Genera solo un PDF para hacer pruebas antes de generar todos. |
| **Nombre de los archivos** | Nombre, Código + Nombre, solo Código o Folio + Nombre. |
| **Notificaciones** | Avisos integrados en la interfaz en lugar de ventanas `alert()`; barra de progreso con botón **Cancelar**. |
| **CSV más flexibles** | Carga varios CSV a la vez, combina `Nombre` + `Apellido Paterno` + `Apellido Materno`, y lee archivos guardados desde Excel (acentos en ANSI/UTF-8). |
| **Duplicados más inteligentes** | Si ambos alumnos tienen código se compara el código; si no, el nombre (ignorando mayúsculas y acentos). |

## 🚀 Cómo usar

1. **📊 Cargar CSV** con los alumnos. Ejemplo:

   ```csv
   Código,Nombre
   219512345,Juan Pérez García
   219512346,María López Hernández
   ```

   También sirve: `Matrícula,Nombre,Apellido Paterno,Apellido Materno`. Las demás columnas se ignoran.

2. **🖼️ Subir imagen de fondo** (JPG o PNG, de preferencia horizontal, 1920 px de ancho o más).
3. En **🧩 Campos** activa los textos que quieras y selecciona uno para editar sus **propiedades**.
4. **Arrastra** cada texto sobre la imagen (o usa las flechas del teclado; `Shift` = más rápido). Al acercarse al centro se “imanta” (mantén `Alt` para evitarlo).
5. Revisa algunos alumnos con ◀ ▶ y prueba con **📄 PDF de la vista actual**.
6. **📦 Generar constancias (ZIP)**.
7. Opcional: **💾 Guardar** la plantilla para reutilizarla.

## 🔧 Calibración

`Calibración (mm)` suma un desplazamiento a **todos** los campos en el PDF (si sale 2 mm a la derecha, escribe −2 en X). Se guarda en la computadora y conserva los valores de la v1.

Fórmulas (iguales a la v1):

- Ancho de la hoja = 279.4 mm; el alto se calcula con la proporción de la imagen.
- Tamaño en PDF (pt) = tamaño elegido × 3.5
- Posición final (mm) = posición arrastrada + calibración

## 📁 Estructura

```
Generador_constancias/
├── index.html               # Interfaz
├── renderer.js              # Lógica del editor, PDFs, fuentes y envío
├── configuracion_fuentes.js # Fuentes originales en Base64 (pantalla y PDF)
├── fuentes/                 # 18 familias de Google Fonts (.ttf) + catalogo.js
├── librerias/               # Lector de Excel (SheetJS)
├── correo.js                # Envío SMTP con nodemailer (proceso principal)
├── main.js                  # Proceso principal de Electron (ventana, correo, contraseña cifrada)
├── package.json
└── _respaldo_v1/            # Archivos de la versión anterior
```

## ▶️ Ejecutar / compilar

```bash
npm install        # instala también nodemailer (necesario para el correo)
npm start          # ejecutar
npm run build      # crear el instalador .exe (ver COMO_CREAR_EXE.md)
```
