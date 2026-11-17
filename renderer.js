/**
 * ==================================================================================
 * GENERADOR DE CONSTANCIAS / DIPLOMAS EN PDF — v3.1
 * ==================================================================================
 *
 * ¿QUÉ HACE ESTE ARCHIVO?
 * Controla todo el editor y la generación masiva de constancias:
 *
 *  1. ALUMNOS: lee uno o varios CSV (nombre + código de estudiante), evita
 *     duplicados y permite ver, buscar, editar, agregar y borrar alumnos.
 *  2. CAMPOS: la constancia puede llevar varios textos (Nombre, Código, Folio,
 *     Fecha y textos libres como el nombre del curso). Cada uno tiene su propia
 *     fuente, tamaño, color, alineación, mayúsculas y ancho máximo.
 *  3. VISTA PREVIA REAL: se navega alumno por alumno para ver exactamente cómo
 *     quedará cada constancia (incluido el auto-ajuste de nombres largos).
 *  4. PLANTILLAS: se guarda/abre todo el diseño (imagen + campos) en un .json.
 *     Además el último diseño se recuerda automáticamente al reabrir.
 *  5. PDF: genera un PDF por alumno y los empaqueta en un ZIP, o solo el PDF
 *     de la vista actual para hacer pruebas.
 *  6. FUENTES Y ESTILOS: 25 familias (7 originales + 18 de Google Fonts en /fuentes),
 *     importar .ttf propios, negrita, cursiva, subrayado, sombra y espaciado.
 *  7. CORREO: envía a cada alumno su constancia (columna Correo del CSV) usando
 *     el proceso principal de Electron (main.js + correo.js con nodemailer).
 *
 * DEPENDENCIAS (se cargan en index.html):
 *  - jsPDF (window.jspdf), JSZip, PapaParse y configuracion_fuentes.js
 *
 * IMPORTANTE SOBRE LA CALIBRACIÓN (igual que en la v1):
 *  - Ancho de la hoja PDF: 279.4 mm (carta). El alto se calcula con la imagen.
 *  - Tamaño en PDF (pt) = tamaño elegido × 3.5
 *  - Posición final (mm) = posición arrastrada + ajuste manual X/Y
 * ==================================================================================
 */

// ==================================================================================
// 1. CONSTANTES
// ==================================================================================

const PDF_WIDTH_MM = 279.4;           // Ancho de carta horizontal (hoja "según la imagen")
const PT_POR_MM = 72 / 25.4;          // 1 mm = 2.8346 pt
// Tamaños de hoja en mm (ancho × alto en vertical)
const TAMANOS_HOJA = {
    carta: [215.9, 279.4],
    oficio: [216, 340],
    a4: [210, 297]
};
const FONT_SIZE_CORRECTION = 3.5;     // tamaño elegido × 3.5 = puntos en el PDF
const MIN_AUTOFIT_PT = 8;             // El auto-ajuste nunca baja de este tamaño
const STORAGE_STATE_KEY = 'gc_state_v2';
const STORAGE_BG_KEY = 'gc_background_v2';

// Nombre CSS de cada fuente para la vista previa (las de Base64 se inyectan
// con su mismo nombre desde configuracion_fuentes.js)
const CSS_FONT = {
    CalibriLight: "'CalibriLight', 'Calibri Light', sans-serif",
    Calibri: "'Calibri', sans-serif",
    times: "'Times New Roman', Times, serif",
    Arial: "'Arial', sans-serif",
    ArialNarrow: "'ArialNarrow', 'Arial Narrow', sans-serif",
    ArialRounded: "'ArialRounded', 'Arial Rounded MT Bold', sans-serif",
    Bahnschrift: "'Bahnschrift', sans-serif",
    helvetica: "Helvetica, Arial, sans-serif",
    courier: "'Courier New', Courier, monospace"
};

const FIELD_TYPE_LABEL = {
    nombre: 'de la lista',
    codigo: 'de la lista',
    evento: 'lista o fijo',
    folio: 'lista o consecutivo',
    fecha: 'lista o fijo',
    columna: 'columna CSV',
    texto: 'texto fijo',
    parrafo: 'párrafo con datos',
    qr: 'verificación',
    imagen: 'imagen'
};

// ==================================================================================
// 2. UTILIDADES
// ==================================================================================

const $ = (id) => document.getElementById(id);

/** Quita acentos, pasa a minúsculas y limpia espacios (para comparar textos) */
function normalizar(texto) {
    return String(texto || '')
        .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
        .toLowerCase()
        .replace(/,/g, ' ')
        .replace(/\s+/g, ' ')
        .trim();
}

/** Normaliza un encabezado de CSV: "Código del Alumno" → "codigodelalumno" */
function normalizarEncabezado(h) {
    return normalizar(h).replace(/[^a-z0-9]/g, '');
}

/** Escapa texto para meterlo en HTML sin riesgos */
function escapeHtml(s) {
    return String(s ?? '').replace(/[&<>"']/g, c => ({
        '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
    }[c]));
}

/** Convierte texto en un nombre de archivo seguro (conserva acentos y ñ) */
function nombreArchivoSeguro(texto) {
    let t = String(texto || '')
        // Letras con acentos que no son del español: se quita el acento en vez de borrar la letra (ç→c, è→e)
        .replace(/[^\x00-\x7FÑñÁÉÍÓÚáéíóúÜü]/g, ch => ch.normalize('NFD').replace(/[\u0300-\u036f]/g, ''))
        .replace(/[^a-zA-Z0-9ÑñÁÉÍÓÚáéíóúÜü_\- ]/g, '')
        .trim()
        .replace(/\s+/g, '_');
    // Nombres reservados de Windows (CON, PRN, AUX, NUL, COM1…, LPT1…)
    if (/^(con|prn|aux|nul|com\d|lpt\d)$/i.test(t)) t += '_';
    return t;
}

/** "5 de octubre de 2026" */
function fechaLarga(fecha = new Date()) {
    const meses = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio',
        'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];
    return `${fecha.getDate()} de ${meses[fecha.getMonth()]} de ${fecha.getFullYear()}`;
}

/**
 * Seguridad al exportar a CSV/Excel: un texto que empieza con = + - @ se
 * tomaría como fórmula al abrirlo en Excel. Se le antepone un apóstrofo.
 */
function celdaSegura(v) {
    const t = String(v ?? '');
    return /^[=+\-@\t\r]/.test(t) ? "'" + t : t;
}

/** Aplica celdaSegura a todas las propiedades de una fila */
function filaSegura(fila) {
    const r = {};
    for (const [k, v] of Object.entries(fila)) r[k] = celdaSegura(v);
    return r;
}

/** '#336699' + 0.5 → 'rgba(51,102,153,0.5)' */
function hexARgba(hex, alfa) {
    const h = String(hex || '#000000').replace('#', '');
    const n = parseInt(h.length === 3 ? h.split('').map(c => c + c).join('') : h, 16) || 0;
    return `rgba(${(n >> 16) & 255},${(n >> 8) & 255},${n & 255},${alfa})`;
}

/** Aplica MAYÚSCULAS o Tipo Título */
function aplicarMayusculas(texto, modo) {
    if (modo === 'upper') return texto.toLocaleUpperCase('es-MX');
    if (modo === 'title') {
        const menores = ['de', 'del', 'la', 'las', 'los', 'y', 'e', 'el', 'en', 'a'];
        return texto.toLocaleLowerCase('es-MX').split(' ').map((p, i) =>
            (i > 0 && menores.includes(p)) ? p : p.charAt(0).toLocaleUpperCase('es-MX') + p.slice(1)
        ).join(' ');
    }
    return texto;
}

/** Descarga un Blob como archivo */
// ----------------------------------------------------------------------------------
// Nombres propios: "JUAN PEREZ DE LA CRUZ" → "Juan Perez de la Cruz"
// ----------------------------------------------------------------------------------
const PARTICULAS = new Set(['de', 'del', 'la', 'las', 'los', 'y', 'e', 'da', 'das', 'do', 'dos', 'van', 'von', 'der', 'di']);

/** Corrige mayúsculas palabra por palabra. Solo toca palabras escritas TODO en mayúsculas o
 *  TODO en minúsculas; respeta las que ya tienen una mezcla intencional (McDonald, DeLuca). */
function nombrePropio(nombre) {
    const palabras = String(nombre || '').replace(/\s+/g, ' ').trim().split(' ');
    return palabras.map((p, i) => {
        const letras = p.replace(/[^A-Za-zÀ-ÿ]/g, '');
        if (!letras) return p;
        const todoMay = letras === letras.toUpperCase(), todoMin = letras === letras.toLowerCase();
        if (!todoMay && !todoMin) return p;
        const min = p.toLowerCase();
        if (i > 0 && PARTICULAS.has(min)) return min;
        // Mayúscula después de inicio, guion o apóstrofo (María-José, O'Higgins)
        return min.replace(/(^|[-'’.])([a-zà-ÿ])/g, (t, sep, c) => sep + c.toUpperCase());
    }).join(' ');
}

/** Palabras que casi siempre llevan acento en nombres y apellidos mexicanos */
const ACENTOS_NOMBRES = Object.fromEntries([
    'Pérez', 'López', 'González', 'Hernández', 'Martínez', 'Rodríguez', 'Sánchez', 'Ramírez', 'Gómez', 'Díaz',
    'Jiménez', 'Gutiérrez', 'Álvarez', 'Fernández', 'Vázquez', 'Domínguez', 'Méndez', 'Chávez', 'García', 'Juárez',
    'Suárez', 'Benítez', 'Velázquez', 'Márquez', 'Ávila', 'Ángeles', 'Ramón', 'José', 'María', 'Jesús', 'Martín',
    'Julián', 'Adrián', 'Sebastián', 'Andrés', 'Héctor', 'Rubén', 'Raúl', 'Iván', 'Joaquín', 'Agustín', 'Germán',
    'Nicolás', 'Tomás', 'Lucía', 'Sofía', 'Mónica', 'Verónica', 'Ángel', 'Efraín', 'Inés', 'Belén', 'Rocío',
    'Concepción', 'Asunción', 'Guzmán', 'León', 'Calderón', 'Rincón', 'Beltrán', 'Galván', 'Durán', 'Román',
    'Cortés', 'Muñoz', 'Núñez', 'Ibáñez', 'Ríos', 'Peña', 'Valdés', 'Báez', 'Sáenz', 'Cázares', 'Íñiguez',
    'Ruíz', 'Ochoa', 'Hernán', 'Simón', 'Fabián', 'Damián', 'Cristián', 'Darío', 'Mario', 'Óscar', 'Martha',
    'Ramírez', 'Jaramillo', 'Zúñiga', 'Montaño', 'Patiño', 'Toño', 'Bolaños', 'Avendaño', 'Treviño', 'Saldaña',
    'Ibarra', 'Arámbula', 'Cárdenas', 'Bárcenas', 'Sánchez', 'Elías', 'Matías', 'Isaías', 'Tobías', 'Josué',
    'Noé', 'Moisés', 'Abraham', 'Rafael', 'Efrén', 'Jonás', 'Saúl', 'Ismael', 'Ezequiel', 'Begoña', 'Nazaret'
].filter(p => /[áéíóúñÁÉÍÓÚÑü]/.test(p) && p !== 'Ruíz').map(p => [normalizar(p), p]));

/** Sugerencias de acento para un nombre: [{ de: 'Perez', a: 'Pérez' }] */
function acentosSugeridos(nombre) {
    const r = [];
    for (const p of String(nombre || '').split(/\s+/)) {
        const sug = ACENTOS_NOMBRES[normalizar(p)];
        if (sug && p !== sug && p.toLowerCase() === normalizar(p) && !/[áéíóúñ]/i.test(p)) r.push({ de: p, a: sug });
    }
    return r;
}

/** Títulos de columna repetidos o vacíos → "Nombre", "Nombre 2"… (para no perder datos) */
function encabezadosUnicos(lista) {
    const vistos = new Set();
    return lista.map((c, j) => {
        const base = String(c ?? '').trim() || `Columna ${j + 1}`;
        let n = base, k = 2;
        while (vistos.has(normalizar(n))) n = `${base} ${k++}`;
        vistos.add(normalizar(n));
        return n;
    });
}

/** Distancia de edición (cuántas letras hay que cambiar), con tope para ir rápido */
function distanciaEdicion(a, b, tope = 3) {
    if (Math.abs(a.length - b.length) > tope) return tope + 1;
    let prev = Array.from({ length: b.length + 1 }, (_, j) => j);
    for (let i = 1; i <= a.length; i++) {
        const cur = [i];
        let minFila = i;
        for (let j = 1; j <= b.length; j++) {
            cur[j] = Math.min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
            if (cur[j] < minFila) minFila = cur[j];
        }
        if (minFila > tope) return tope + 1;
        prev = cur;
    }
    return prev[b.length];
}

/**
 * Pares de filas [i, j] del mismo evento con nombres casi iguales pero NO idénticos:
 * mismo nombre sin acentos/mayúsculas, mismas palabras en otro orden, o 1-2 letras de diferencia.
 * Si ambos tienen código y es distinto, solo se avisa cuando los nombres son iguales sin acentos.
 */
let _parecidosCache = { firma: '', pares: [] };
function nombresParecidos(lista) {
    const firma = lista.map(a => (a.evento || '') + '\u0001' + (a.nombre || '') + '\u0001' + (a.codigo || '')).join('\u0002');
    if (firma === _parecidosCache.firma) return _parecidosCache.pares;
    const pares = [];
    const vistos = new Set();
    const porEvento = new Map();
    lista.forEach((a, i) => {
        const ev = normalizar(a.evento);
        if (!porEvento.has(ev)) porEvento.set(ev, []);
        const n = normalizar(a.nombre).replace(/[^a-z0-9 ]/g, '').replace(/\s+/g, ' ').trim();
        const pal = n.split(' ');
        porEvento.get(ev).push({ i, n, pal, orden: [...pal].sort().join(' '), cod: normalizar(a.codigo) });
    });
    for (const grupo of porEvento.values()) {
        // Solo se comparan nombres que comparten el inicio de alguna palabra (rápido aunque haya miles)
        const cubetas = new Map();
        for (const A of grupo) {
            if (!A.n) continue;
            for (const p of new Set(A.pal.filter(w => w.length >= 3).map(w => w.slice(0, 3)))) {
                if (!cubetas.has(p)) cubetas.set(p, []);
                cubetas.get(p).push(A);
            }
        }
        for (const cub of cubetas.values()) {
            if (cub.length > 400) continue; // apellido muy común: se omite esa cubeta
            for (let x = 0; x < cub.length; x++) {
                for (let y = x + 1; y < cub.length; y++) {
                    const A = cub[x], B = cub[y];
                    const clave = A.i < B.i ? A.i + ',' + B.i : B.i + ',' + A.i;
                    if (vistos.has(clave)) continue;
                    vistos.add(clave);
                    if (A.n === B.n) continue; // iguales sin acentos: ya lo avisa otra revisión
                    if (A.cod && B.cod && A.cod !== B.cod) continue; // códigos distintos: son personas distintas
                    let parecido = A.orden === B.orden;
                    if (!parecido && A.n.length >= 10 && Math.abs(A.n.length - B.n.length) <= 2) {
                        parecido = distanciaEdicion(A.n, B.n, 2) <= (A.n.length >= 18 ? 2 : 1);
                    }
                    if (parecido) pares.push(A.i < B.i ? [A.i, B.i] : [B.i, A.i]);
                }
            }
        }
    }
    _parecidosCache = { firma, pares };
    return pares;
}

/** Comprueba que una imagen se pueda abrir y no sea absurdamente grande */
function probarImagen(dataUrl) {
    return new Promise((resolve, reject) => {
        const im = new Image();
        im.onload = () => {
            if (!im.naturalWidth || !im.naturalHeight) return reject(new Error('La imagen está vacía.'));
            if (im.naturalWidth * im.naturalHeight > 150e6) return reject(new Error('La imagen es demasiado grande (más de 150 megapíxeles).'));
            resolve({ ancho: im.naturalWidth, alto: im.naturalHeight });
        };
        im.onerror = () => reject(new Error('La imagen de fondo está dañada o no es PNG/JPG.'));
        im.src = dataUrl;
    });
}

/** Lee una imagen y la reduce si es muy grande (conserva la transparencia si es PNG) */
function reducirImagen(file, maxLado) {
    return new Promise((resolve, reject) => {
        const r = new FileReader();
        r.onerror = () => reject(new Error('No se pudo leer el archivo.'));
        r.onload = () => {
            const img = new Image();
            img.onerror = () => reject(new Error('La imagen está dañada o no es PNG/JPG.'));
            img.onload = () => {
                const esc = Math.min(1, maxLado / Math.max(img.naturalWidth, img.naturalHeight));
                const proporcion = img.naturalHeight / img.naturalWidth;
                if (esc >= 1) return resolve({ dataUrl: r.result, proporcion });
                const c = document.createElement('canvas');
                c.width = Math.round(img.naturalWidth * esc);
                c.height = Math.round(img.naturalHeight * esc);
                c.getContext('2d').drawImage(img, 0, 0, c.width, c.height);
                const png = /png/i.test(file.type);
                resolve({ dataUrl: c.toDataURL(png ? 'image/png' : 'image/jpeg', 0.9), proporcion });
            };
            img.src = r.result;
        };
        r.readAsDataURL(file);
    });
}

function descargarBlob(blob, nombre) {
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = nombre;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    setTimeout(() => URL.revokeObjectURL(link.href), 2000);
}

/** Lee un archivo de texto detectando si viene en UTF-8 o en ANSI (Excel en Windows) */
async function leerTextoConAcentos(file) {
    const leer = (encoding) => new Promise((resolve, reject) => {
        const r = new FileReader();
        r.onload = () => resolve(r.result);
        r.onerror = () => reject(r.error);
        r.readAsText(file, encoding);
    });
    let texto = await leer('UTF-8');
    // El carácter de reemplazo indica que no era UTF-8 → reintentar como Windows-1252
    if (texto.includes('\uFFFD')) {
        texto = await leer('windows-1252');
    }
    return texto.replace(/^\uFEFF/, '');
}

// ==================================================================================
// 3. NOTIFICACIONES Y DIÁLOGOS (reemplazan a alert() y confirm())
// ==================================================================================

const UI = {
    /**
     * Muestra una notificación flotante.
     * @param {string} mensaje
     * @param {'info'|'success'|'warning'|'error'} tipo
     * @param {{duracion?:number, accion?:{texto:string, fn:Function}}} opciones
     */
    toast(mensaje, tipo = 'info', opciones = {}) {
        const iconos = { info: 'ℹ️', success: '✅', warning: '⚠️', error: '❌' };
        const el = document.createElement('div');
        el.className = `toast ${tipo}`;
        el.innerHTML = `<span>${iconos[tipo] || ''}</span><span class="toast-msg"></span>`;
        el.querySelector('.toast-msg').textContent = mensaje;

        const cerrar = () => {
            el.classList.add('hide');
            setTimeout(() => el.remove(), 300);
        };

        if (opciones.accion) {
            const b = document.createElement('button');
            b.className = 'btn small';
            b.textContent = opciones.accion.texto;
            b.onclick = () => { opciones.accion.fn(); cerrar(); };
            el.appendChild(b);
        }
        const x = document.createElement('button');
        x.className = 'close-x';
        x.style.fontSize = '0.9rem';
        x.textContent = '✕';
        x.onclick = cerrar;
        el.appendChild(x);

        $('toast-container').appendChild(el);
        setTimeout(cerrar, opciones.duracion || (tipo === 'error' ? 8000 : 4500));
    },

    /** Diálogo de confirmación. Devuelve una promesa con true/false */
    confirmar(titulo, mensaje, textoOk = 'Aceptar', peligro = false) {
        return new Promise(resolve => {
            $('confirm-title').textContent = titulo;
            $('confirm-message').textContent = mensaje;
            const ok = $('confirm-ok');
            const cancel = $('confirm-cancel');
            ok.textContent = textoOk;
            ok.className = peligro ? 'btn danger' : 'btn primary';
            ok.style.padding = '0.55rem 1rem';
            const modal = $('confirm-modal');
            modal.classList.add('open');
            const terminar = (valor) => {
                modal.classList.remove('open');
                ok.onclick = cancel.onclick = null;
                resolve(valor);
            };
            ok.onclick = () => terminar(true);
            cancel.onclick = () => terminar(false);
            setTimeout(() => ok.focus(), 50);
        });
    },

    /** Ventana con varios botones. opciones: [{ v, t, clase }]. Devuelve el v elegido (null al cerrar). */
    elegir(titulo, mensaje, opciones) {
        return new Promise(resolve => {
            let m = $('choice-modal');
            if (!m) {
                m = document.createElement('div');
                m.className = 'modal-backdrop';
                m.id = 'choice-modal';
                m.style.zIndex = '9450';
                m.innerHTML = `<div class="modal" style="width:min(480px,92vw)"><div class="modal-header"><h3 id="choice-title"></h3></div>
                    <div class="modal-body"><p id="choice-msg" style="white-space:pre-line;line-height:1.5"></p></div>
                    <div class="modal-footer" id="choice-btns"></div></div>`;
                document.body.appendChild(m);
            }
            $('choice-title').textContent = titulo;
            $('choice-msg').textContent = mensaje;
            const cont = $('choice-btns');
            cont.innerHTML = '';
            const terminar = (v) => { m.classList.remove('open'); document.removeEventListener('keydown', esc); resolve(v); };
            const esc = (e) => { if (e.key === 'Escape') terminar(null); };
            document.addEventListener('keydown', esc);
            [...opciones].reverse().forEach(o => {
                const b = document.createElement('button');
                b.className = 'btn ' + (o.clase || '');
                b.textContent = o.t;
                b.onclick = () => terminar(o.v);
                cont.prepend(b);
            });
            m.classList.add('open');
            setTimeout(() => cont.querySelector('.primary')?.focus(), 50);
        });
    },

    /** Pide un texto corto. Devuelve el texto o null si se cancela. */
    pedirTexto(titulo, mensaje, inicial = '', textoOk = 'Aceptar') {
        return new Promise(resolve => {
            let m = $('prompt-modal');
            if (!m) {
                m = document.createElement('div');
                m.className = 'modal-backdrop';
                m.id = 'prompt-modal';
                m.style.zIndex = '9450';
                m.innerHTML = `<div class="modal" style="width:min(460px,92vw)"><div class="modal-header"><h3 id="prompt-title"></h3></div>
                    <div class="modal-body"><p id="prompt-msg" style="line-height:1.5"></p><input id="prompt-input" class="input" maxlength="80"></div>
                    <div class="modal-footer"><button class="btn" id="prompt-cancel">Cancelar</button><button class="btn primary" id="prompt-ok"></button></div></div>`;
                document.body.appendChild(m);
            }
            $('prompt-title').textContent = titulo;
            $('prompt-msg').textContent = mensaje;
            $('prompt-ok').textContent = textoOk;
            const inp = $('prompt-input');
            inp.value = inicial;
            const terminar = (v) => { m.classList.remove('open'); inp.onkeydown = null; resolve(v); };
            $('prompt-ok').onclick = () => { const v = inp.value.trim(); if (v) terminar(v); else inp.focus(); };
            $('prompt-cancel').onclick = () => terminar(null);
            inp.onkeydown = (e) => {
                if (e.key === 'Enter') $('prompt-ok').click();
                if (e.key === 'Escape') terminar(null);
            };
            m.classList.add('open');
            setTimeout(() => { inp.focus(); inp.select(); }, 50);
        });
    },

    abrirModal(id) { $(id)?.classList.add('open'); },
    cerrarModal(id) { $(id)?.classList.remove('open'); }
};

// ==================================================================================
// 3b. PUENTE CON ELECTRON (archivos y proceso principal)
// ==================================================================================
// En la app de escritorio el HTML tiene acceso a Node.js (nodeIntegration).
// Si el archivo se abre en un navegador normal, estas funciones quedan
// desactivadas y el programa sigue funcionando (sin correo ni fuentes importadas).

const Escritorio = (() => {
    let fs = null, path = null, ipc = null, carpetaApp = '';
    try {
        if (typeof require === 'function') {
            fs = require('fs');
            path = require('path');
            ipc = require('electron').ipcRenderer;
            // Carpeta donde está index.html (también funciona dentro de app.asar)
            carpetaApp = (typeof __dirname !== 'undefined' && __dirname)
                ? __dirname
                : decodeURIComponent(new URL('.', location.href).pathname).replace(/^\/([A-Za-z]:)/, '$1');
        }
    } catch (e) {
        console.warn('Funciones de escritorio no disponibles:', e.message);
    }
    // Para pruebas automáticas fuera de Electron
    if (!ipc && window.gcIpcPrueba) ipc = window.gcIpcPrueba;

    return {
        disponible: !!ipc,
        tieneArchivos: !!fs,
        fs, path, carpetaApp,
        invocar(canal, ...args) {
            if (!ipc) return Promise.reject(new Error('Disponible solo en la aplicación de escritorio.'));
            return ipc.invoke(canal, ...args);
        }
    };
})();

// ==================================================================================
// 3b-2. BASE DE DATOS LOCAL (IndexedDB): plantillas recientes, historial…
// ==================================================================================
// Se guarda en el perfil del programa en esta computadora. Si no está disponible
// (por ejemplo, en una ventana privada) las funciones fallan en silencio.
const BD = {
    _bd: null,
    abrir() {
        if (this._bd) return this._bd;
        this._bd = new Promise((resolve, reject) => {
            if (!window.indexedDB) return reject(new Error('Sin base de datos local'));
            const r = indexedDB.open('generador_constancias', 3);
            r.onupgradeneeded = () => {
                const db = r.result;
                if (!db.objectStoreNames.contains('plantillas')) db.createObjectStore('plantillas', { keyPath: 'id' });
                if (!db.objectStoreNames.contains('historial')) {
                    const h = db.createObjectStore('historial', { keyPath: 'clave' });
                    h.createIndex('evento', 'eventoClave');
                }
                if (!db.objectStoreNames.contains('ajustes')) db.createObjectStore('ajustes', { keyPath: 'id' });
                if (!db.objectStoreNames.contains('proyectos')) db.createObjectStore('proyectos', { keyPath: 'id' });
            };
            r.onsuccess = () => {
                const db = r.result;
                db.onversionchange = () => db.close();
                resolve(db);
            };
            r.onerror = () => reject(r.error);
            r.onblocked = () => reject(new Error('La base de datos está en uso por otra ventana del programa.'));
        });
        this._bd.catch(() => { this._bd = null; });
        return this._bd;
    },
    async _tx(store, modo, fn) {
        const db = await this.abrir();
        return new Promise((resolve, reject) => {
            const tx = db.transaction(store, modo);
            const res = fn(tx.objectStore(store));
            tx.oncomplete = () => resolve(res && 'result' in res ? res.result : res);
            tx.onerror = () => reject(tx.error);
            tx.onabort = () => reject(tx.error);
        });
    },
    poner(store, obj) { return this._tx(store, 'readwrite', s => s.put(obj)); },
    ponerVarios(store, objs) { return this._tx(store, 'readwrite', s => { for (const o of objs) s.put(o); }); },
    leer(store, clave) { return this._tx(store, 'readonly', s => s.get(clave)); },
    todos(store) { return this._tx(store, 'readonly', s => s.getAll()); },
    borrar(store, clave) { return this._tx(store, 'readwrite', s => s.delete(clave)); },
    vaciar(store) { return this._tx(store, 'readwrite', s => s.clear()); }
};

// ==================================================================================
// 3b-3. HISTORIAL DE CONSTANCIAS EMITIDAS Y CÓDIGOS DE VERIFICACIÓN
// ==================================================================================

/** SHA-256 (síncrono) de un texto. Devuelve un arreglo de 32 bytes. */
function sha256Bytes(texto) {
    const bytes = Array.from(new TextEncoder().encode(String(texto)));
    const K = [0x428a2f98, 0x71374491, 0xb5c0fbcf, 0xe9b5dba5, 0x3956c25b, 0x59f111f1, 0x923f82a4, 0xab1c5ed5,
        0xd807aa98, 0x12835b01, 0x243185be, 0x550c7dc3, 0x72be5d74, 0x80deb1fe, 0x9bdc06a7, 0xc19bf174,
        0xe49b69c1, 0xefbe4786, 0x0fc19dc6, 0x240ca1cc, 0x2de92c6f, 0x4a7484aa, 0x5cb0a9dc, 0x76f988da,
        0x983e5152, 0xa831c66d, 0xb00327c8, 0xbf597fc7, 0xc6e00bf3, 0xd5a79147, 0x06ca6351, 0x14292967,
        0x27b70a85, 0x2e1b2138, 0x4d2c6dfc, 0x53380d13, 0x650a7354, 0x766a0abb, 0x81c2c92e, 0x92722c85,
        0xa2bfe8a1, 0xa81a664b, 0xc24b8b70, 0xc76c51a3, 0xd192e819, 0xd6990624, 0xf40e3585, 0x106aa070,
        0x19a4c116, 0x1e376c08, 0x2748774c, 0x34b0bcb5, 0x391c0cb3, 0x4ed8aa4a, 0x5b9cca4f, 0x682e6ff3,
        0x748f82ee, 0x78a5636f, 0x84c87814, 0x8cc70208, 0x90befffa, 0xa4506ceb, 0xbef9a3f7, 0xc67178f2];
    let H = [0x6a09e667, 0xbb67ae85, 0x3c6ef372, 0xa54ff53a, 0x510e527f, 0x9b05688c, 0x1f83d9ab, 0x5be0cd19];
    const l = bytes.length * 8;
    bytes.push(0x80);
    while (bytes.length % 64 !== 56) bytes.push(0);
    for (let i = 7; i >= 0; i--) bytes.push(i >= 4 ? 0 : (l >>> (i * 8)) & 0xff);
    const rotr = (x, n) => (x >>> n) | (x << (32 - n));
    const W = new Array(64);
    for (let o = 0; o < bytes.length; o += 64) {
        for (let i = 0; i < 16; i++) W[i] = (bytes[o + i * 4] << 24) | (bytes[o + i * 4 + 1] << 16) | (bytes[o + i * 4 + 2] << 8) | bytes[o + i * 4 + 3];
        for (let i = 16; i < 64; i++) {
            const s0 = rotr(W[i - 15], 7) ^ rotr(W[i - 15], 18) ^ (W[i - 15] >>> 3);
            const s1 = rotr(W[i - 2], 17) ^ rotr(W[i - 2], 19) ^ (W[i - 2] >>> 10);
            W[i] = (W[i - 16] + s0 + W[i - 7] + s1) | 0;
        }
        let [a, b, c, d, e, f, g, h] = H;
        for (let i = 0; i < 64; i++) {
            const S1 = rotr(e, 6) ^ rotr(e, 11) ^ rotr(e, 25);
            const t1 = (h + S1 + ((e & f) ^ (~e & g)) + K[i] + W[i]) | 0;
            const S0 = rotr(a, 2) ^ rotr(a, 13) ^ rotr(a, 22);
            const t2 = (S0 + ((a & b) ^ (a & c) ^ (b & c))) | 0;
            h = g; g = f; f = e; e = (d + t1) | 0; d = c; c = b; b = a; a = (t1 + t2) | 0;
        }
        H = [H[0] + a, H[1] + b, H[2] + c, H[3] + d, H[4] + e, H[5] + f, H[6] + g, H[7] + h].map(x => x | 0);
    }
    const out = [];
    for (const x of H) out.push((x >>> 24) & 0xff, (x >>> 16) & 0xff, (x >>> 8) & 0xff, x & 0xff);
    return out;
}

/** Código corto legible (sin letras que se confunden): ABCD-2345 */
function codigoLegible(bytes) {
    const ALF = '23456789ABCDEFGHJKLMNPQRSTUVWXYZ'; // sin 0/O, 1/I
    let bits = '', t = '';
    for (const b of bytes.slice(0, 5)) bits += b.toString(2).padStart(8, '0');
    for (let i = 0; i < 40; i += 5) t += ALF[parseInt(bits.slice(i, i + 5), 2)];
    return t.slice(0, 4) + '-' + t.slice(4);
}

const limpiarCodigo = (c) => String(c || '').toUpperCase().replace(/[^0-9A-Z]/g, '').replace(/O/g, '0').replace(/I/g, '1');

const Historial = {
    registros: new Map(),       // clave → registro
    porCodigo: new Map(),       // código de verificación (limpio) → clave
    secreto: '',
    disponible: false,

    async init() {
        // Clave secreta de esta instalación: hace que los códigos no se puedan adivinar
        try {
            this.secreto = localStorage.getItem('gc_secreto') || '';
            if (!this.secreto) {
                const r = new Uint8Array(16);
                crypto.getRandomValues(r);
                this.secreto = Array.from(r, b => b.toString(16).padStart(2, '0')).join('');
                localStorage.setItem('gc_secreto', this.secreto);
            }
        } catch (e) { this.secreto = this.secreto || 'sin-secreto'; }
        try {
            const ajuste = await BD.leer('ajustes', 'secreto');
            if (ajuste && ajuste.valor && ajuste.valor !== this.secreto) {
                // La base de datos manda (por si se borró el almacenamiento local)
                this.secreto = ajuste.valor;
                try { localStorage.setItem('gc_secreto', this.secreto); } catch (e) { }
            } else if (!ajuste) {
                await BD.poner('ajustes', { id: 'secreto', valor: this.secreto });
            }
            const todos = await BD.todos('historial');
            for (const r of todos) this._indexar(r);
            this.disponible = true;
        } catch (e) {
            console.warn('Historial no disponible', e);
        }
    },

    _version: 0,
    _indexar(r) {
        this._version++;
        if (typeof Editor !== 'undefined') Editor._folioCache = null;
        this.registros.set(r.clave, r);
        if (r.verificacion) this.porCodigo.set(limpiarCodigo(r.verificacion), r.clave);
    },

    clave(alumno) {
        const ev = normalizar(alumno.evento);
        const quien = alumno.codigo ? 'c:' + normalizar(alumno.codigo) : 'n:' + normalizar(alumno.nombre);
        return ev + '|' + quien;
    },

    de(alumno) {
        if (!alumno || !this.registros.size) return null;
        const r = this.registros.get(this.clave(alumno));
        if (r) return r;
        // La misma persona pudo registrarse antes sin código (por nombre): se reconoce igual
        if (alumno.codigo) {
            const porNombre = this.registros.get(normalizar(alumno.evento) + '|n:' + normalizar(alumno.nombre));
            if (porNombre && !porNombre.codigo) return porNombre;
        }
        return null;
    },

    /** Folio consecutivo más alto ya emitido en un evento */
    maxFolio(evento) {
        const ev = normalizar(evento);
        if (this._maxCache && this._maxCache.n === this.registros.size && this._maxCache.v === this._version && this._maxCache.m.has(ev)) return this._maxCache.m.get(ev);
        if (!this._maxCache || this._maxCache.n !== this.registros.size || this._maxCache.v !== this._version) this._maxCache = { n: this.registros.size, v: this._version, m: new Map() };
        let max = null; // null = aún no hay folios numéricos de este evento
        for (const r of this.registros.values()) {
            if (r.eventoClave !== ev || !/^\d+$/.test(String(r.folio || ''))) continue;
            const n = parseInt(r.folio, 10);
            if (max === null || n > max) max = n;
        }
        this._maxCache.m.set(ev, max);
        return max;
    },

    /** Código de verificación de una persona (el mismo siempre para la misma persona y evento) */
    codigo(alumno, folio) {
        const r = this.de(alumno);
        if (r && r.verificacion) return r.verificacion;
        const base = [this.secreto, normalizar(alumno ? alumno.evento : ''), alumno ? this.clave(alumno) : 'ejemplo', folio].join('|');
        return codigoLegible(sha256Bytes(base));
    },

    /**
     * Guarda en el historial las constancias que se generaron, imprimieron o enviaron.
     * @param items [{ alumno, indice }]
     * @param accion 'generada' | 'impresa' | 'enviada'
     */
    async registrar(items, accion, extra = {}) {
        items = items.filter(it => it && it.alumno && it.alumno.nombre);
        if (!items.length) return;
        const ahora = Date.now();
        // Primero se calculan TODOS los folios y códigos (con el historial como estaba)
        const datos = items.map(({ alumno, indice }) => {
            const folio = Editor.folioDe(alumno, indice);
            return { alumno, folio, verificacion: this.codigo(alumno, folio) };
        });
        const nuevos = [], viejas = [];
        for (const { alumno, folio, verificacion } of datos) {
            const clave = this.clave(alumno);
            const previo = this.de(alumno);
            if (previo && previo.clave !== clave) {
                // Antes estaba registrada por nombre; ahora tiene código: se mueve el registro
                this.registros.delete(previo.clave);
                viejas.push(previo.clave);
                previo.clave = clave;
            }
            const r = Object.assign({
                clave, eventoClave: normalizar(alumno.evento), primera: ahora,
                acciones: { generada: 0, impresa: 0, enviada: 0 }
            }, previo || {});
            Object.assign(r, {
                nombre: alumno.nombre, codigo: alumno.codigo || '', correo: alumno.correo || r.correo || '',
                evento: alumno.evento || '', fecha: alumno.fecha || r.fecha || '',
                // El folio guardado es el que se imprimió (si cambió, p. ej. por la columna Folio, se actualiza)
                folio, folioNum: /^\d+$/.test(folio) ? parseInt(folio, 10) : 0,
                verificacion: r.verificacion || verificacion,
                extra: Object.assign({}, r.extra || {}, alumno.extra || {}),
                plantilla: Editor.nombrePlantilla || r.plantilla || '',
                ultima: ahora
            });
            r.acciones = Object.assign({ generada: 0, impresa: 0, enviada: 0 }, r.acciones);
            r.acciones[accion] = (r.acciones[accion] || 0) + 1;
            if (accion === 'enviada') { r.enviadoA = extra.para || alumno.correo || ''; r.enviadoEn = ahora; }
            this._indexar(r);
            nuevos.push(r);
        }
        try {
            for (const k of viejas) await BD.borrar('historial', k);
            await BD.ponerVarios('historial', nuevos);
        } catch (e) {
            console.warn('No se pudo guardar el historial', e);
            if (!this._avisoFallo) {
                this._avisoFallo = true;
                UI.toast('⚠️ No se pudo guardar el historial de constancias en esta computadora (¿disco lleno o el programa abierto dos veces?). ' +
                    'Los folios podrían repetirse después. Cierra las otras ventanas del programa y vuelve a intentar.', 'error', { duracion: 15000 });
            }
        }
    },

    /** El registro convertido en "alumno" para volver a generar/enviar su constancia */
    comoAlumno(r) {
        return {
            codigo: r.codigo, nombre: r.nombre, correo: r.correo, evento: r.evento,
            folio: r.folio, fecha: r.fecha, extra: Object.assign({}, r.extra || {})
        };
    },

    buscarCodigo(codigo) {
        const clave = this.porCodigo.get(limpiarCodigo(codigo));
        return clave ? this.registros.get(clave) : null;
    },

    lista() {
        return [...this.registros.values()].sort((a, b) => (b.ultima || 0) - (a.ultima || 0));
    },

    async borrarTodo() {
        this.registros.clear();
        this.porCodigo.clear();
        await BD.vaciar('historial');
    },

    async borrar(clave) {
        const r = this.registros.get(clave);
        if (r && r.verificacion) this.porCodigo.delete(limpiarCodigo(r.verificacion));
        this.registros.delete(clave);
        await BD.borrar('historial', clave);
    },

    filas() {
        const f = (t) => t ? new Date(t).toLocaleString('es-MX') : '';
        return this.lista().map(r => filaSegura(Object.assign({
            Folio: r.folio, Nombre: r.nombre, Codigo: r.codigo, Correo: r.correo, Evento: r.evento, Fecha: r.fecha,
            Verificacion: r.verificacion, Primera_emision: f(r.primera), Ultima_actividad: f(r.ultima),
            Generada: r.acciones?.generada || 0, Impresa: r.acciones?.impresa || 0, Enviada: r.acciones?.enviada || 0,
            Enviada_a: r.enviadoA || '', Enviada_el: f(r.enviadoEn), Plantilla: r.plantilla || ''
        }, r.extra || {})));
    },

    exportarExcel() {
        const filas = this.filas();
        if (!filas.length) return UI.toast('El historial está vacío.', 'warning');
        if (typeof XLSX === 'undefined') {
            return descargarBlob(new Blob(['﻿' + Papa.unparse(filas)], { type: 'text/csv;charset=utf-8' }), 'historial_constancias.csv');
        }
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(filas), 'Historial');
        descargarBlob(new Blob([XLSX.write(wb, { type: 'array', bookType: 'xlsx' })],
            { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), 'historial_constancias.xlsx');
    },

    /** Respaldo completo (para pasarlo a otra computadora) */
    exportarRespaldo() {
        const datos = { app: 'generador-constancias', tipo: 'historial', version: 1, exportado: new Date().toISOString(),
            secreto: this.secreto, registros: this.lista() };
        descargarBlob(new Blob([JSON.stringify(datos)], { type: 'application/json' }), 'respaldo_historial_constancias.json');
    },

    async importarRespaldo(texto) {
        const d = JSON.parse(texto);
        if (!d || d.app !== 'generador-constancias' || d.tipo !== 'historial' || !Array.isArray(d.registros)) {
            throw new Error('El archivo no es un respaldo del historial.');
        }
        let n = 0;
        const nuevos = [];
        for (const r0 of d.registros) {
            if (!r0 || typeof r0.clave !== 'string' || typeof r0.nombre !== 'string') continue;
            const r = {
                clave: r0.clave.substring(0, 300), eventoClave: String(r0.eventoClave || '').substring(0, 200),
                nombre: r0.nombre.substring(0, 200), codigo: String(r0.codigo || '').substring(0, 60),
                correo: String(r0.correo || '').substring(0, 200), evento: String(r0.evento || '').substring(0, 200),
                fecha: String(r0.fecha || '').substring(0, 100), folio: String(r0.folio || '').substring(0, 60),
                folioNum: parseInt(r0.folioNum, 10) || 0, verificacion: String(r0.verificacion || '').substring(0, 20),
                primera: +r0.primera || Date.now(), ultima: +r0.ultima || Date.now(),
                acciones: { generada: +r0.acciones?.generada || 0, impresa: +r0.acciones?.impresa || 0, enviada: +r0.acciones?.enviada || 0 },
                enviadoA: String(r0.enviadoA || '').substring(0, 200), enviadoEn: +r0.enviadoEn || 0,
                plantilla: String(r0.plantilla || '').substring(0, 80),
                extra: {}
            };
            if (r0.extra && typeof r0.extra === 'object') {
                for (const [k, v] of Object.entries(r0.extra).slice(0, 40)) r.extra[String(k).substring(0, 60)] = String(v ?? '').substring(0, 300);
            }
            // Las claves se calculan de nuevo (no se confía en las del archivo)
            r.eventoClave = normalizar(r.evento);
            r.clave = this.clave(r);
            if (this.registros.has(r.clave) || nuevos.some(x => x.clave === r.clave)) continue; // no se pisan los que ya hay
            this._indexar(r);
            nuevos.push(r);
            n++;
        }
        await BD.ponerVarios('historial', nuevos);
        return n;
    },

    /** Archivo para la página web de verificación (GitHub Pages) */
    exportarVerificacion(nombreCompleto, emisor) {
        const enmascarar = (n) => {
            const p = String(n).trim().split(/\s+/);
            return p.length <= 1 ? p[0] : p[0] + ' ' + p.slice(1).map(x => x[0].toUpperCase() + '.').join(' ');
        };
        const registros = {};
        for (const r of this.registros.values()) {
            if (!r.verificacion) continue;
            registros[limpiarCodigo(r.verificacion)] = {
                f: r.folio, n: nombreCompleto ? r.nombre : enmascarar(r.nombre), e: r.evento, d: r.fecha,
                t: new Date(r.primera).toLocaleDateString('es-MX', { day: 'numeric', month: 'long', year: 'numeric' })
            };
        }
        const datos = { emisor: emisor || '', actualizado: new Date().toISOString(), total: Object.keys(registros).length, registros };
        descargarBlob(new Blob([JSON.stringify(datos)], { type: 'application/json' }), 'registro.json');
        return datos.total;
    }
};

// ==================================================================================
// 3b-4. AVISO DE VERSIÓN NUEVA (GitHub Releases)
// ==================================================================================
// El repositorio se configura en package.json → "actualizaciones": { "github": "usuario/repositorio" }
// o con clic en el número de versión (arriba a la derecha). Se revisa una vez al día.
const Actualizaciones = {
    version: '',
    repo: '',

    async init() {
        try {
            const info = await Escritorio.invocar('app:info');
            this.version = info.version || '';
            this.repo = info.repo || '';
        } catch (e) { /* navegador */ }
        try { this.repo = localStorage.getItem('gc_repo') || this.repo; } catch (e) { }
        const ver = document.querySelector('.titlebar .ver');
        if (ver) {
            ver.style.cursor = 'pointer';
            ver.title = 'Buscar una versión nueva';
            ver.addEventListener('click', () => this.revisar(true));
        }
        let ultima = 0;
        try { ultima = +localStorage.getItem('gc_ultima_revision') || 0; } catch (e) { }
        if (this.repo && Date.now() - ultima > 24 * 3600 * 1000) setTimeout(() => this.revisar(false), 4000);
    },

    comparar(a, b) {
        const pa = String(a).replace(/^v/i, '').split('.').map(n => parseInt(n, 10) || 0);
        const pb = String(b).replace(/^v/i, '').split('.').map(n => parseInt(n, 10) || 0);
        for (let i = 0; i < 3; i++) if ((pa[i] || 0) !== (pb[i] || 0)) return (pa[i] || 0) - (pb[i] || 0);
        return 0;
    },

    async revisar(manual) {
        if (!this.repo) {
            if (!manual) return;
            const r = await UI.pedirTexto('Buscar versiones nuevas',
                'Escribe el repositorio de GitHub donde publicas el instalador, por ejemplo «javi/generador-constancias». ' +
                'El programa revisará ahí una vez al día.', '', 'Guardar');
            if (!r) return;
            if (!/^[A-Za-z0-9_.-]+\/[A-Za-z0-9_.-]+$/.test(r.trim())) return UI.toast('Escribe usuario/repositorio, sin espacios.', 'warning');
            this.repo = r.trim();
            try { localStorage.setItem('gc_repo', this.repo); } catch (e) { }
        }
        if (!/^[A-Za-z0-9_.-]+\/[A-Za-z0-9_.-]+$/.test(this.repo)) return;
        try { localStorage.setItem('gc_ultima_revision', String(Date.now())); } catch (e) { }
        const actual = this.version || (document.querySelector('.titlebar .ver')?.textContent || '').replace(/^v/i, '');
        try {
            const resp = await fetch(`https://api.github.com/repos/${this.repo}/releases/latest`, { headers: { Accept: 'application/vnd.github+json' } });
            if (resp.status === 404) {
                if (manual) UI.toast('Ese repositorio aún no tiene versiones publicadas (Releases).', 'info');
                return;
            }
            if (!resp.ok) throw new Error('GitHub respondió ' + resp.status);
            const rel = await resp.json();
            const nueva = String(rel.tag_name || '').replace(/^v/i, '');
            const url = String(rel.html_url || '');
            if (nueva && this.comparar(nueva, actual) > 0 && /^https:\/\/github\.com\//.test(url)) {
                UI.toast(`Hay una versión nueva: v${nueva} (tienes v${actual}).`, 'info', {
                    duracion: 15000,
                    accion: { texto: 'Descargar', fn: () => window.open(url, '_blank') }
                });
            } else if (manual) {
                UI.toast(`Tienes la versión más reciente (v${actual}).`, 'success');
            }
        } catch (e) {
            if (manual) UI.toast('No se pudo revisar (¿sin internet?):\n' + e.message, 'warning');
        }
    }
};

// ==================================================================================
// 3c. FUENTES (originales, adicionales de /fuentes e importadas por el usuario)
// ==================================================================================

const ESTILOS = ['normal', 'bold', 'italic', 'bolditalic'];
const FAUX_BOLD = 0.035;      // Grosor del contorno de la negrita simulada (fracción del tamaño)
const FAUX_ITALIC = 0.2;      // Inclinación de la cursiva simulada (≈ 11.3°)

const Fuentes = {
    familias: [],             // { id, nombre, categoria, base, archivos: {estilo: {ruta, origen}} }
    _cache: new Map(),        // 'id|estilo' → base64
    carpetaUsuario: '',

    async init() {
        // 1) Las 7 originales (Base64 en configuracion_fuentes.js, solo estilo normal)
        const originales = [
            ['CalibriLight', 'Calibri Light'], ['Calibri', 'Calibri'], ['times', 'Times New Roman'],
            ['Arial', 'Arial'], ['ArialNarrow', 'Arial Narrow'], ['ArialRounded', 'Arial Rounded MT Bold'],
            ['Bahnschrift', 'Bahnschrift Condensed']
        ];
        this.familias = originales.map(([id, nombre]) => ({
            id, nombre, categoria: 'originales', base: true, archivos: { normal: { origen: 'base' } }
        }));
        // Times, Helvetica y Courier vienen dentro de jsPDF con sus 4 estilos reales
        const completos = { origen: 'base' };
        this.familia('times').archivos = { normal: completos, bold: completos, italic: completos, bolditalic: completos };
        for (const [id, nombre] of [['helvetica', 'Helvetica'], ['courier', 'Courier']]) {
            this.familias.push({
                id, nombre, categoria: 'originales', base: true,
                archivos: { normal: completos, bold: completos, italic: completos, bolditalic: completos }
            });
        }

        // 2) Las adicionales de la carpeta /fuentes
        const extra = (typeof FUENTES_EXTRA !== 'undefined') ? FUENTES_EXTRA : [];
        for (const f of extra) {
            const archivos = {};
            for (const [estilo, archivo] of Object.entries(f.archivos)) {
                archivos[estilo] = { ruta: 'fuentes/' + archivo, origen: 'app' };
            }
            this.familias.push({ id: f.id, nombre: f.nombre, categoria: f.categoria, base: false, archivos });
        }

        // 3) Las importadas por el usuario (se guardan en la carpeta de datos del programa)
        await this.cargarImportadas();
        this.inyectarCSS();
    },

    async cargarImportadas() {
        if (!Escritorio.tieneArchivos || !Escritorio.disponible) return;
        try {
            const datos = await Escritorio.invocar('app:rutaDatos');
            this.carpetaUsuario = Escritorio.path.join(datos, 'fuentes_usuario');
            const catalogo = Escritorio.path.join(this.carpetaUsuario, 'catalogo.json');
            if (!Escritorio.fs.existsSync(catalogo)) return;
            const lista = JSON.parse(Escritorio.fs.readFileSync(catalogo, 'utf8'));
            for (const imp of lista) {
                let fam = this.familia(imp.id);
                if (!fam) {
                    fam = { id: imp.id, nombre: imp.nombre, categoria: 'importadas', base: false, archivos: {} };
                    this.familias.push(fam);
                }
                for (const [estilo, archivo] of Object.entries(imp.archivos)) {
                    const ruta = Escritorio.path.join(this.carpetaUsuario, archivo);
                    if (Escritorio.fs.existsSync(ruta)) fam.archivos[estilo] = { ruta, origen: 'usuario', archivo };
                }
            }
        } catch (e) {
            console.warn('No se pudieron leer las fuentes importadas', e);
        }
    },

    familia(id) {
        return this.familias.find(f => f.id === id);
    },

    tiene(id, estilo) {
        const f = this.familia(id);
        return !!(f && f.archivos[estilo]);
    },

    /** Nombre CSS de la familia para la vista previa */
    css(id) {
        const f = this.familia(id);
        if (!f) return 'sans-serif';
        if (f.base && !Object.values(f.archivos).some(a => a.origen === 'usuario')) {
            return CSS_FONT[id] || `'${id}'`;
        }
        return `'GC-${id}', ${CSS_FONT[id] || 'sans-serif'}`;
    },

    /** Crea las reglas @font-face para la vista previa */
    inyectarCSS() {
        let css = '';
        for (const f of this.familias) {
            for (const [estilo, a] of Object.entries(f.archivos)) {
                let url;
                if (a.origen === 'app') url = a.ruta;
                else if (a.origen === 'usuario') url = 'file:///' + a.ruta.replace(/\\/g, '/').replace(/^\//, '');
                else if (f.base && estilo === 'normal' && Object.values(f.archivos).some(x => x.origen === 'usuario')) {
                    // Familia original con estilos importados: la normal sale del Base64
                    const b64 = this.base64(f.id, 'normal');
                    if (b64) url = 'data:font/ttf;base64,' + b64;
                }
                if (!url) continue;
                css += `@font-face{font-family:'GC-${f.id}';src:url("${encodeURI(url).replace(/#/g, '%23')}") format('truetype');` +
                    `font-weight:${estilo.includes('bold') ? 700 : 400};font-style:${estilo.includes('italic') ? 'italic' : 'normal'};}\n`;
            }
        }
        let tag = document.getElementById('gc-font-faces');
        if (!tag) {
            tag = document.createElement('style');
            tag.id = 'gc-font-faces';
            document.head.appendChild(tag);
        }
        tag.textContent = css;
    },

    /** Lee el archivo .ttf y lo devuelve en Base64 (con caché) */
    base64(id, estilo) {
        const clave = id + '|' + estilo;
        if (this._cache.has(clave)) return this._cache.get(clave);
        const f = this.familia(id);
        const a = f && f.archivos[estilo];
        if (!a) return null;
        let b64 = null;
        if (a.origen === 'base') {
            // Fuentes originales guardadas en configuracion_fuentes.js
            const mapa = {
                CalibriLight: 'calibriLight', Calibri: 'calibri', Arial: 'arial',
                ArialNarrow: 'arialNarrow', ArialRounded: 'arialRounded', Bahnschrift: 'bahnschrift'
            };
            const crudo = (typeof fonts !== 'undefined' && mapa[id]) ? fonts[mapa[id]] : null;
            b64 = crudo ? crudo.trim().replace(/^data:font\/ttf;base64,/, '').replace(/^2/, '').replace(/\s/g, '') : null;
        } else if (Escritorio.tieneArchivos) {
            const ruta = a.origen === 'app' ? Escritorio.path.join(Escritorio.carpetaApp, a.ruta) : a.ruta;
            b64 = Escritorio.fs.readFileSync(ruta).toString('base64');
        } else {
            // Navegador: lectura síncrona (solo para pruebas)
            const xhr = new XMLHttpRequest();
            xhr.open('GET', a.ruta, false);
            xhr.overrideMimeType('text/plain; charset=x-user-defined');
            xhr.send();
            let bin = '';
            for (let i = 0; i < xhr.responseText.length; i++) bin += String.fromCharCode(xhr.responseText.charCodeAt(i) & 0xff);
            b64 = btoa(bin);
        }
        this._cache.set(clave, b64);
        return b64;
    },

    /** Registra la fuente en un documento jsPDF solo cuando se necesita */
    asegurarEnDoc(doc, id, estilo) {
        const f = this.familia(id);
        if (!f || !f.archivos[estilo]) return false;
        // Las originales en estilo normal ya las registra configuracion_fuentes.js
        if (f.archivos[estilo].origen === 'base') return true;
        doc.__gcFuentes = doc.__gcFuentes || new Set();
        const clave = id + '|' + estilo;
        if (doc.__gcFuentes.has(clave)) return true;
        const b64 = this.base64(id, estilo);
        if (!b64) return false;
        const vfs = `${id}-${estilo}.ttf`;
        doc.addFileToVFS(vfs, b64);
        doc.addFont(vfs, id, estilo);
        doc.__gcFuentes.add(clave);
        return true;
    },

    /**
     * Decide qué archivo usar y qué simular según Negrita/Cursiva.
     * @returns {{estilo:string, fauxBold:boolean, fauxItalic:boolean}}
     */
    resolver(campo) {
        const id = campo.fuente;
        const b = !!campo.negrita, i = !!campo.cursiva;
        const r = (estilo, fauxBold = false, fauxItalic = false) => ({ estilo, fauxBold, fauxItalic });
        if (b && i) {
            if (this.tiene(id, 'bolditalic')) return r('bolditalic');
            if (this.tiene(id, 'bold')) return r('bold', false, true);
            if (this.tiene(id, 'italic')) return r('italic', true, false);
            return r('normal', true, true);
        }
        if (b) return this.tiene(id, 'bold') ? r('bold') : r('normal', true);
        if (i) return this.tiene(id, 'italic') ? r('italic') : r('normal', false, true);
        return r('normal');
    },

    /** Llena un <select> con las familias agrupadas */
    llenarSelect(select, conNueva = false) {
        const grupos = [
            ['originales', 'Originales'], ['sans', 'Sans serif (modernas)'], ['serif', 'Serif (elegantes)'],
            ['script', 'Caligráficas'], ['importadas', 'Importadas por ti']
        ];
        const valor = select.value;
        select.innerHTML = '';
        for (const [cat, titulo] of grupos) {
            const lista = this.familias.filter(f => f.categoria === cat);
            if (!lista.length) continue;
            const og = document.createElement('optgroup');
            og.label = titulo;
            for (const f of lista) {
                const op = document.createElement('option');
                op.value = f.id;
                op.textContent = f.nombre;
                og.appendChild(op);
            }
            select.appendChild(og);
        }
        if (conNueva) {
            const op = document.createElement('option');
            op.value = '__nueva__';
            op.textContent = '➕ Nueva familia…';
            select.appendChild(op);
        }
        if (valor && [...select.options].some(o => o.value === valor)) select.value = valor;
    },

    /** Importa un archivo .ttf a la carpeta de datos del programa */
    async importar(file, familiaId, nombreNueva, estilo) {
        if (!Escritorio.tieneArchivos || !this.carpetaUsuario) {
            throw new Error('Importar fuentes solo funciona en la aplicación de escritorio.');
        }
        if (!ESTILOS.includes(estilo)) throw new Error('Estilo de fuente no válido.');
        if (file.size > 30 * 1024 * 1024) throw new Error('El archivo es demasiado grande para ser una fuente (más de 30 MB).');
        const buffer = new Uint8Array(await file.arrayBuffer());
        const firma = String.fromCharCode(...buffer.slice(0, 4));
        const esTTF = (buffer[0] === 0 && buffer[1] === 1 && buffer[2] === 0 && buffer[3] === 0) || firma === 'true';
        if (firma === 'OTTO') throw new Error('Ese archivo es OpenType con contornos CFF (.otf). Usa una versión .ttf.');
        if (firma === 'ttcf') throw new Error('Ese archivo es una colección (.ttc). Usa un archivo .ttf individual.');
        if (!esTTF) throw new Error('El archivo no parece ser una fuente TrueType (.ttf) válida.');

        let id = familiaId, nombre;
        if (familiaId === '__nueva__') {
            nombre = (nombreNueva || '').trim() || file.name.replace(/\.ttf$/i, '');
            const base = 'U_' + (nombre.normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^A-Za-z0-9]/g, '').substring(0, 40) || 'fuente' + Date.now());
            // Nunca se mezcla con otra familia: si el nombre ya existe se numera
            id = base;
            let k = 2;
            while (this.familia(id)) id = `${base}_${k++}`;
        } else {
            nombre = this.familia(id)?.nombre || id;
        }

        const { fs, path } = Escritorio;
        fs.mkdirSync(this.carpetaUsuario, { recursive: true });
        const archivo = `${id}-${estilo}.ttf`;
        fs.writeFileSync(path.join(this.carpetaUsuario, archivo), buffer);

        const rutaCat = path.join(this.carpetaUsuario, 'catalogo.json');
        const cat = fs.existsSync(rutaCat) ? JSON.parse(fs.readFileSync(rutaCat, 'utf8')) : [];
        let entrada = cat.find(c => c.id === id);
        if (!entrada) { entrada = { id, nombre, archivos: {} }; cat.push(entrada); }
        entrada.archivos[estilo] = archivo;
        fs.writeFileSync(rutaCat, JSON.stringify(cat, null, 2));

        // Recargar (y medir otra vez con el archivo nuevo)
        this._cache.delete(id + '|' + estilo);
        await this.init();
        if (typeof Editor !== 'undefined') Editor._docMedidas = null;
        return { id, nombre };
    },

    /** Quita una fuente importada (solo las del usuario) */
    eliminarImportada(id, estilo) {
        const { fs, path } = Escritorio;
        const rutaCat = path.join(this.carpetaUsuario, 'catalogo.json');
        let cat = [];
        try { cat = JSON.parse(fs.readFileSync(rutaCat, 'utf8')); } catch (e) { return; }
        if (!Array.isArray(cat)) return;
        const entrada = cat.find(c => c && c.id === id);
        if (!entrada || !entrada.archivos || !entrada.archivos[estilo]) return;
        // Solo se borra un archivo dentro de la carpeta de fuentes (nunca una ruta de afuera)
        try { fs.unlinkSync(path.join(this.carpetaUsuario, path.basename(String(entrada.archivos[estilo])))); } catch (e) { }
        delete entrada.archivos[estilo];
        const nuevo = cat.filter(c => Object.keys(c.archivos).length);
        fs.writeFileSync(rutaCat, JSON.stringify(nuevo, null, 2));
        this._cache.delete(id + '|' + estilo);
        if (typeof Editor !== 'undefined') Editor._docMedidas = null;
    }
};

// ==================================================================================
// 4. LISTA DE ALUMNOS
// ==================================================================================
// Cada registro (un alumno en un evento):
//   { codigo, nombre, correo, evento, folio, fecha,
//     extra: { "Horas": "20", "Instructor": "..." },   ← otras columnas del CSV
//     envio: { estado: 'ok'|'error', detalle, fecha } }
// Un mismo alumno puede aparecer varias veces si participó en varios talleres.

const Alumnos = {
    lista: [],
    archivos: [],         // Nombres de los CSV cargados (para nombrar el ZIP)
    columnasExtra: [],    // Otras columnas encontradas en los CSV (se pueden imprimir)

    /** Encabezados que se reconocen (se comparan sin acentos, espacios ni mayúsculas) */
    ALIAS: {
        codigo: ['codigo', 'code', 'matricula', 'clave', 'codigoestudiante',
            'codigodeestudiante', 'codigoalumno', 'codigodelalumno', 'numerodecontrol', 'nocontrol',
            'numcontrol', 'boleta', 'idalumno', 'expediente', 'numerodecuenta', 'cuenta'],
        nombre: ['nombrecompleto', 'nombre', 'name', 'fullname', 'alumno', 'nombredelalumno',
            'nombrealumno', 'estudiante', 'participante', 'nombres'],
        correo: ['correo', 'email', 'mail', 'correoelectronico', 'correoinstitucional',
            'emailinstitucional', 'correoalumno', 'correodelalumno', 'emailaddress', 'direcciondecorreo',
            'correoe', 'ecorreo', 'correopersonal', 'emailalumno'],
        evento: ['evento', 'taller', 'curso', 'actividad', 'nombredelevento', 'nombredeltaller',
            'nombredelcurso', 'eventoocurso', 'eventotaller', 'tallercurso', 'cursotaller', 'conferencia', 'diplomado',
            'seminario', 'ponencia', 'modulo'],
        folio: ['folio', 'nofolio', 'numerodefolio', 'numfolio', 'foliodelaconstancia', 'folioconstancia'],
        fecha: ['fecha', 'fechadelevento', 'fechadeltaller', 'fechadelcurso', 'fechas', 'periodo'],
        apPaterno: ['apellidopaterno', 'primerapellido', 'apellido1', 'paterno', 'appaterno'],
        apMaterno: ['apellidomaterno', 'segundoapellido', 'apellido2', 'materno', 'apmaterno'],
        apellidos: ['apellidos', 'apellido', 'lastname', 'surname']
    },

    /** ¿Qué dato es este encabezado? ('nombre', 'codigo', … o '' si es una columna extra) */
    tipoDeEncabezado(h) {
        const n = normalizarEncabezado(h);
        if (!n) return '';
        for (const [tipo, claves] of Object.entries(this.ALIAS)) if (claves.includes(n)) return tipo;
        return '';
    },

    /** Agrega datos pegados/escritos en la tabla del programa (sin archivo) */
    cargarPegado(campos, filas) {
        const datos = filas.map(f => Object.fromEntries(campos.map((c, j) => [c, f[j] ?? ''])));
        const r = this.procesarTabla(campos, datos);
        if (r.agregados && !this.archivos.includes('Datos pegados')) this.archivos.push('Datos pegados');
        return r;
    },

    /** ¿Es el mismo registro? Debe ser el MISMO evento y el mismo alumno
     *  (por código si ambos lo tienen; si no, por nombre sin importar acentos). */
    esDuplicado(a, b) {
        if (normalizar(a.evento) !== normalizar(b.evento)) return false;
        if (a.codigo && b.codigo) return normalizar(a.codigo) === normalizar(b.codigo);
        return normalizar(a.nombre) === normalizar(b.nombre);
    },

    /** Lista de eventos distintos con cuántos registros tiene cada uno */
    eventos() {
        const mapa = new Map();
        for (const a of this.lista) {
            const clave = normalizar(a.evento);
            if (!mapa.has(clave)) mapa.set(clave, { nombre: a.evento || '', total: 0 });
            mapa.get(clave).total++;
        }
        return [...mapa.values()];
    },

    existe(alumno, ignorarIndice = -1) {
        return this.lista.findIndex((x, i) => i !== ignorarIndice && this.esDuplicado(x, alumno));
    },

    agregar(alumno) {
        if (this.existe(alumno) !== -1) return false;
        this.lista.push(alumno);
        return true;
    },

    get conCodigo() {
        return this.lista.filter(a => a.codigo).length;
    },

    get conCorreo() {
        return this.lista.filter(a => esCorreoValido(a.correo)).length;
    },

    /** Lee un CSV y devuelve un resumen de lo que se agregó */
    /** Lee un CSV o un Excel (.xlsx/.xls/.ods). Un Excel puede traer varias hojas. */
    async cargarCSV(file) {
        const tablas = await this.leerTablas(file);
        if (!tablas.length) throw new Error('El archivo está vacío o no tiene encabezados.');
        const varias = tablas.length > 1;
        const suma = { agregados: 0, duplicados: 0, vacios: 0, total: 0, extras: [], hojas: tablas.length };
        for (const t of tablas) {
            // Si el Excel tiene varias hojas y no trae columna Evento, el nombre de la hoja es el evento
            const r = this.procesarTabla(t.campos, t.filas, varias ? t.hoja : '');
            for (const k of ['agregados', 'duplicados', 'vacios', 'total']) suma[k] += r[k];
            for (const c of r.extras) if (!suma.extras.includes(c)) suma.extras.push(c);
            for (const k of ['colNombre', 'colCodigo', 'colCorreo', 'colEvento', 'colFolio', 'colFecha']) {
                if (!suma[k] && r[k]) suma[k] = r[k];
            }
            if (varias && !r.colEvento) suma.eventoPorHoja = true;
        }
        this.archivos.push(file.name);
        return suma;
    },

    /** Convierte el archivo en tablas { hoja, campos, filas } */
    async leerTablas(file) {
        if (/\.(xlsx|xlsm|xls|ods)$/i.test(file.name)) {
            if (typeof XLSX === 'undefined') throw new Error('Falta el lector de Excel (librerias/xlsx.full.min.js).');
            const wb = XLSX.read(new Uint8Array(await file.arrayBuffer()), { type: 'array', cellNF: true });
            const tablas = [];
            for (const hoja of wb.SheetNames) {
                const filas = this.leerHojaExcel(wb.Sheets[hoja]);
                const iEnc = filas.findIndex(f => f.some(v => String(v).trim()));
                if (iEnc < 0) continue;
                const campos = encabezadosUnicos(filas[iEnc].map((h, j) => String(h).trim() || `Columna ${j + 1}`));
                const datos = filas.slice(iEnc + 1)
                    .filter(f => f.some(v => String(v).trim()))
                    .map(f => Object.fromEntries(campos.map((c, j) => [c, f[j] ?? ''])));
                if (datos.length) tablas.push({ hoja, campos, filas: datos });
            }
            return tablas;
        }
        const texto = await leerTextoConAcentos(file);
        const res = Papa.parse(texto, {
            header: true,
            skipEmptyLines: 'greedy',
            transformHeader: h => String(h).replace(/^\uFEFF/, '').trim()
        });
        const campos = res.meta.fields || [];
        // Errores del archivo (comillas sin cerrar, filas con columnas de más): se avisa, no se ignoran
        const errores = (res.errors || []).filter(e => e.code !== 'UndetectableDelimiter');
        const conDeMas = res.data.filter(f => f.__parsed_extra).length;
        if (errores.length || conDeMas) {
            const filasMal = [...new Set(errores.map(e => (e.row ?? -1) + 2).filter(n => n > 1))].slice(0, 5);
            UI.toast(`Revisa «${file.name}»: ${errores.length + conDeMas} fila(s) con un problema de formato` +
                (filasMal.length ? ` (por ejemplo la fila ${filasMal.join(', ')})` : '') +
                '. Puede haber comillas sin cerrar o comas de más dentro de un dato. Revisa la lista antes de generar.', 'warning', { duracion: 12000 });
        }
        res.data.forEach(f => { delete f.__parsed_extra; });
        return campos.length ? [{ hoja: '', campos, filas: res.data }] : [];
    },

    /**
     * Lee una hoja de Excel celda por celda:
     *  - Los números enteros se conservan completos (un código de 14 dígitos no se vuelve 2.19E+13).
     *  - Las fechas se escriben en español: "5 de octubre de 2026".
     *  - Lo demás se toma como el texto que se ve en Excel.
     */
    leerHojaExcel(ws) {
        if (!ws || !ws['!ref']) return [];
        const rango = XLSX.utils.decode_range(ws['!ref']);
        // Algunos programas declaran rangos enormes (A1:XFD1048576): se usan los límites reales de las celdas
        let maxR = rango.s.r, maxC = rango.s.c;
        for (const k of Object.keys(ws)) {
            if (k[0] === '!') continue;
            const d = XLSX.utils.decode_cell(k);
            if (d.r > maxR) maxR = d.r;
            if (d.c > maxC) maxC = d.c;
        }
        rango.e.r = Math.min(rango.e.r, maxR, rango.s.r + 100000);
        rango.e.c = Math.min(rango.e.c, maxC, rango.s.c + 300);
        const filas = [];
        for (let r = rango.s.r; r <= rango.e.r; r++) {
            const fila = [];
            for (let c = rango.s.c; c <= rango.e.c; c++) {
                const celda = ws[XLSX.utils.encode_cell({ r, c })];
                fila.push(this.valorCelda(celda));
            }
            if (fila.some(v => String(v).trim())) filas.push(fila);
        }
        return filas;
    },

    valorCelda(celda) {
        if (!celda || celda.v === undefined || celda.v === null) return '';
        if (celda.t === 'e') return ''; // #N/A, #REF!… no son datos
        if (celda.t === 'n') {
            const formato = celda.z || '';
            if (formato && XLSX.SSF.is_date(formato)) {
                const f = XLSX.SSF.parse_date_code(celda.v);
                if (f && f.y) return fechaLarga(new Date(f.y, f.m - 1, f.d));
            }
            // Enteros con formato General/0/@ (códigos, matrículas): completos, sin notación científica
            if (Number.isInteger(celda.v) && Math.abs(celda.v) < 1e21) {
                // Formato con ceros a la izquierda (00123): se respetan; cualquier otro (#,##0 · 0.00): número limpio
                if (/^0+$/.test(formato) && formato.length > 1) return String(celda.v).padStart(formato.length, '0');
                if (!formato || /^(General|0|@|#,##0|#,##0\.0+|0\.0+)$/i.test(formato)) return BigInt(celda.v).toString();
            }
            return celda.w ?? String(celda.v);
        }
        if (celda.t === 'd' && celda.v instanceof Date) return fechaLarga(celda.v);
        if (celda.t === 'b') return celda.v ? 'Sí' : 'No';
        return String(celda.w ?? celda.v);
    },

    /** Agrega a la lista los registros de una tabla */
    procesarTabla(campos, filas, eventoHoja = '') {
        // --- Detectar columnas por su encabezado ---
        const A = Alumnos.ALIAS;
        // Se respeta la prioridad de los nombres (Código antes que Clave, Nombre completo antes que Nombre)
        const buscar = (claves) => {
            for (const k of claves) {
                const c = campos.find(x => normalizarEncabezado(x) === k);
                if (c) return c;
            }
            return undefined;
        };

        const colCodigo = buscar(A.codigo);
        let colNombre = buscar(A.nombre);
        const colCorreo = buscar(A.correo);
        const colEvento = buscar(A.evento);
        const colFolio = buscar(A.folio);
        const colFecha = buscar(A.fecha);
        const colApPat = buscar(A.apPaterno);
        const colApMat = buscar(A.apMaterno);
        const colApellidos = buscar(A.apellidos);

        // Si no hay columna de nombre reconocida: una que diga "nombre" o, si no, la primera columna de
        // texto que no sea código, correo, evento, folio, fecha ni apellidos (y que no sean puros números)
        if (!colNombre) {
            const ocupadas = new Set([colCodigo, colCorreo, colEvento, colFolio, colFecha, colApPat, colApMat, colApellidos].filter(Boolean));
            const libres = campos.filter(c => !ocupadas.has(c));
            const soloNumeros = (c) => filas.slice(0, 30).every(f => /^[\d\s.,-]*$/.test(String(f[c] ?? '')));
            colNombre = libres.find(c => /nombre|participante|asistente|alumn|estudiante/.test(normalizarEncabezado(c)))
                || libres.find(c => !soloNumeros(c));
        }
        if (!colNombre) throw new Error('No se encontró una columna con nombres.');

        // Solo se combinan apellidos si la columna de nombre NO es ya el nombre completo
        const nombreEsCompleto = normalizarEncabezado(colNombre) === 'nombrecompleto';

        // Las columnas que no se reconocieron se guardan como "extra"
        const usadas = new Set([colCodigo, colNombre, colCorreo, colEvento, colFolio, colFecha,
            nombreEsCompleto ? null : colApPat, nombreEsCompleto ? null : colApMat,
            nombreEsCompleto ? null : colApellidos].filter(Boolean));
        const extras = campos.filter(c => c && !usadas.has(c));
        for (const c of extras) if (!this.columnasExtra.includes(c)) this.columnasExtra.push(c);

        // Índice para detectar duplicados rápido (miles de filas sin trabar el programa)
        const porCodigo = new Map(), porNombre = new Map();
        const indexar = (a, i) => {
            const ev = normalizar(a.evento);
            if (a.codigo) porCodigo.set(ev + '|' + normalizar(a.codigo), i);
            const kn = ev + '|' + normalizar(a.nombre);
            if (!porNombre.has(kn)) porNombre.set(kn, []);
            porNombre.get(kn).push(i);
        };
        this.lista.forEach(indexar);
        const duplicadoDe = (reg) => {
            const ev = normalizar(reg.evento);
            const delNombre = porNombre.get(ev + '|' + normalizar(reg.nombre)) || [];
            if (reg.codigo) {
                const k = porCodigo.get(ev + '|' + normalizar(reg.codigo));
                if (k !== undefined) return k;
                const sinCodigo = delNombre.find(i => !this.lista[i].codigo);
                return sinCodigo === undefined ? -1 : sinCodigo;
            }
            return delNombre.length ? delNombre[0] : -1;
        };

        let agregados = 0, duplicados = 0, vacios = 0;
        for (const fila of filas) {
            let nombre = String(fila[colNombre] ?? '').trim();
            if (!nombreEsCompleto) {
                const partes = [nombre];
                if (colApPat) partes.push(fila[colApPat]);
                if (colApMat) partes.push(fila[colApMat]);
                if (!colApPat && !colApMat && colApellidos) partes.push(fila[colApellidos]);
                nombre = partes.filter(Boolean).map(s => String(s).trim()).join(' ');
            }
            nombre = nombre.replace(/,/g, ' ').replace(/\s+/g, ' ').trim();
            if (!nombre) { vacios++; continue; }

            const codigo = colCodigo ? String(fila[colCodigo] ?? '').trim() : '';
            const correo = colCorreo ? String(fila[colCorreo] ?? '').replace(/\s/g, '').toLowerCase() : '';
            const leer = (col) => col ? String(fila[col] ?? '').replace(/\s+/g, ' ').trim() : '';
            const extra = {};
            for (const c of extras) {
                const v = leer(c);
                if (v) extra[c] = v;
            }
            const registro = {
                codigo, nombre, correo,
                evento: leer(colEvento) || eventoHoja, folio: leer(colFolio), fecha: leer(colFecha), extra
            };
            const dup = duplicadoDe(registro);
            if (dup === -1) {
                this.lista.push(registro);
                indexar(registro, this.lista.length - 1);
                agregados++;
            } else {
                duplicados++;
                // Si el duplicado trae datos que faltaban, se completan
                const existente = this.lista[dup];
                for (const k of ['codigo', 'correo', 'folio', 'fecha']) if (!existente[k] && registro[k]) existente[k] = registro[k];
                if (registro.codigo && existente.codigo === registro.codigo) indexar(existente, dup);
                existente.extra = Object.assign({}, extra, existente.extra || {});
            }
        }

        return {
            agregados, duplicados, vacios, colNombre, colCodigo, colCorreo, colEvento, colFolio, colFecha,
            extras, total: filas.length
        };
    },

    /** Exporta la lista actual como CSV (con BOM para que Excel respete acentos) */
    exportarCSV() {
        const csv = Papa.unparse(this.lista.map(a => filaSegura((() => {
            const fila = {
                Codigo: a.codigo, Nombre: a.nombre, Correo: a.correo || '',
                Evento: a.evento || '', Folio: a.folio || '', Fecha: a.fecha || ''
            };
            for (const c of this.columnasExtra) fila[c] = (a.extra && a.extra[c]) || '';
            return fila;
        })())));
        descargarBlob(new Blob(['\uFEFF' + csv], { type: 'text/csv;charset=utf-8' }), 'lista_alumnos.csv');
    },

    /** Exporta la lista a Excel (.xlsx) */
    exportarExcel() {
        if (typeof XLSX === 'undefined') return UI.toast('Falta el lector de Excel.', 'error');
        const filas = this.lista.map(a => {
            const fila = {
                Codigo: a.codigo, Nombre: a.nombre, Correo: a.correo || '',
                Evento: a.evento || '', Folio: a.folio || '', Fecha: a.fecha || ''
            };
            for (const c of this.columnasExtra) fila[c] = (a.extra && a.extra[c]) || '';
            return fila;
        });
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(filas.map(filaSegura)), 'Alumnos');
        const datos = XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
        descargarBlob(new Blob([datos], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), 'lista_alumnos.xlsx');
    },

    /** CSV de ejemplo para que el usuario vea el formato */
    descargarEjemplo() {
        const filas = [
            ['Codigo', 'Nombre', 'Correo', 'Evento', 'Folio', 'Fecha', 'Carrera', 'Grupo', 'Horas'],
            ['219512345', 'Juan Pérez García', 'juan.perez@alumnos.udg.mx', 'Taller de Git', 'GIT-001', '5 de octubre de 2026', 'Teleinformática', 'A', '10'],
            ['219512346', 'María López Hernández', 'maria.lopez@alumnos.udg.mx', 'Taller de Git', 'GIT-002', '5 de octubre de 2026', 'Agronegocios', 'B', '10'],
            ['219512345', 'Juan Pérez García', 'juan.perez@alumnos.udg.mx', 'Curso de Python', 'PY-001', '12 de octubre de 2026', 'Teleinformática', 'A', '20'],
            ['219512347', 'Ana Martínez Ruiz', 'ana.martinez@alumnos.udg.mx', 'Curso de Python', 'PY-002', '12 de octubre de 2026', 'Teleinformática', 'B', '20']
        ];
        descargarBlob(new Blob(['\uFEFF' + Papa.unparse(filas)], { type: 'text/csv;charset=utf-8' }), 'ejemplo_alumnos.csv');
    }
};

// ==================================================================================
// 5. CAMPOS DE TEXTO DE LA CONSTANCIA
// ==================================================================================

/** Crea un campo con valores por defecto */
function crearCampo(tipo, extra = {}) {
    const base = {
        id: tipo,
        tipo,
        etiqueta: tipo,
        activo: true,
        x: 50, y: 50,              // Posición en % de la hoja (punto de anclaje)
        tamano: 8,                 // × 3.5 = puntos en PDF
        fuente: 'CalibriLight',
        color: '#000000',
        alineacion: 'center',
        anchoMax: 0,               // mm (0 = sin límite)
        mayusculas: 'none',
        prefijo: '',
        valor: '',
        folioInicio: 1,
        folioDigitos: 4,
        columna: '',               // Para campos de tipo "columna" (otra columna del CSV)
        // Estilos
        negrita: false,
        cursiva: false,
        subrayado: false,
        sombra: false,
        sombraColor: '#000000',
        sombraDist: 6,             // % del tamaño de la letra
        sombraOpacidad: 45,        // %
        espaciado: 0,              // % del tamaño de la letra entre cada letra
        // Párrafos con datos
        resaltar: true,            // Los datos de la lista van en negrita
        interlineado: 1.3,         // Separación entre renglones (× tamaño de letra)
        maxRenglones: 0,           // 0 = sin límite; si se pasa, la letra se reduce
        // Folio
        continuarFolio: true,      // No repetir folios ya emitidos (historial)
        // Código QR
        lado: 24,                  // mm
        qrModo: 'texto',           // 'texto' | 'url'
        qrUrl: '',
        qrLeyenda: true,
        qrFondo: true,
        // Imagen encima del fondo (firma, sello)
        imagen: '',                // data:image/png|jpeg;base64,…
        anchoImg: 45,              // mm (el alto sale de la proporción)
        proporcion: 0.5            // alto / ancho
    };
    return Object.assign(base, extra);
}

function camposPorDefecto() {
    return [
        crearCampo('nombre', { etiqueta: 'Nombre del alumno', y: 50, tamano: 12, anchoMax: 180 }),
        crearCampo('codigo', { etiqueta: 'Código de estudiante', y: 60, tamano: 5, prefijo: 'Código: ' }),
        crearCampo('evento', {
            id: 'curso', etiqueta: 'Evento / Taller', activo: false, y: 40, tamano: 6,
            valor: 'Nombre del curso o evento', anchoMax: 220
        }),
        crearCampo('fecha', { etiqueta: 'Fecha', activo: false, y: 82, tamano: 4.5, valor: fechaLarga() }),
        crearCampo('folio', {
            etiqueta: 'Folio', activo: false, x: 92, y: 93, tamano: 3.5, alineacion: 'right',
            prefijo: 'Folio: '
        })
    ];
}

// ==================================================================================
// 6. EDITOR PRINCIPAL
// ==================================================================================

const Editor = {
    campos: camposPorDefecto(),
    seleccionado: 'nombre',
    indiceVista: -1,              // -1 = ejemplo, 0..n = alumno real
    fondo: null,                  // { dataUrl, ancho, alto }
    // Hoja: tamano 'imagen'|'carta'|'oficio'|'a4' · orientacion 'horizontal'|'vertical' · ajuste de la imagen
    pagina: { tamano: 'imagen', orientacion: 'horizontal', ajuste: 'estirar' },
    filtroEvento: '__todos__',
    manual: null,                 // Datos de una constancia capturada a mano (sin CSV)
    // Opciones de salida: un PDF por constancia o uno solo, carpeta o ZIP, organizar por…, formato del fondo
    opciones: { modo: 'individual', destino: 'carpeta', agruparPor: '', formatoFondo: 'jpeg', carpeta: '', formatoImagen: 'png', dpi: 150 },
    cancelarGeneracion: false,
    _docMedidas: null,
    _guardarTimer: null,

    // ------------------------------------------------------------------
    // ARRANQUE
    // ------------------------------------------------------------------
    init() {
        this.restaurarEstado();
        this.conectarAlumnos();
        this.conectarPlantilla();
        this.conectarCalibracion();
        this.conectarPropiedades();
        this.conectarVistaPrevia();
        this.conectarGeneracion();
        this.conectarModales();
        this.conectarTeclado();
        this.conectarFuentes();
        this.conectarPagina();
        this.conectarManual();
        Envio.init();

        // Re-dibujar cuando cambie el tamaño de la ventana (cambia la escala)
        new ResizeObserver(() => this.ajustarLienzo()).observe($('canvas-scroll'));

        this.renderListaCampos();
        this.renderPropiedades();
        this.actualizarResumen();
        this.actualizarVista();
        console.log('=== Generador de Constancias v3.1 listo ===');
    },

    campo(id) {
        return this.campos.find(c => c.id === id);
    },

    get campoSel() {
        return this.campo(this.seleccionado);
    },

    // ------------------------------------------------------------------
    // TEXTO QUE LLEVA CADA CAMPO PARA UN ALUMNO
    // ------------------------------------------------------------------
    /**
     * @param campo  definición del campo
     * @param alumno {codigo, nombre} o null para el ejemplo
     * @param indice posición del alumno en la lista (para el folio)
     * @returns texto final o '' si no se debe dibujar
     */
    textoDeCampo(campo, alumno, indice) {
        let t = '';
        switch (campo.tipo) {
            case 'nombre':
                t = alumno ? alumno.nombre : 'Nombre del Alumno de Ejemplo';
                t = t.replace(/,/g, '').replace(/\s+/g, ' ').trim();
                break;
            case 'codigo': {
                const cod = alumno ? alumno.codigo : '219000000';
                t = cod ? (campo.prefijo || '') + cod : '';
                break;
            }
            case 'folio':
                t = (campo.prefijo || '') + this.folioDe(alumno, indice, campo);
                break;
            case 'evento':
                // Si el CSV trae el evento de cada alumno se usa; si no, el texto fijo
                t = (alumno && alumno.evento) || campo.valor || '';
                break;
            case 'fecha':
                t = (alumno && alumno.fecha) || campo.valor || '';
                break;
            case 'columna': {
                const v = alumno ? (alumno.extra || {})[campo.columna] : `[${campo.columna || 'columna'}]`;
                t = v ? (campo.prefijo || '') + v : '';
                break;
            }
            case 'qr':
                return this.contenidoQR(campo, alumno, indice);
            case 'imagen':
                return campo.imagen && campo.imagen.startsWith('data:') ? '[imagen]' : '';
            case 'parrafo':
                // Ya viene con mayúsculas aplicadas por partes
                return this.partesParrafo(campo, alumno, indice).map(p => p.t).join('').replace(/\s+/g, ' ').trim();
            default: // texto libre
                t = campo.valor || '';
        }
        return aplicarMayusculas(t, campo.mayusculas);
    },

    /** Folio del CSV si existe; si no, consecutivo según la posición en la lista */
    folioDe(alumno, indice, campo = this.campo('folio') || crearCampo('folio')) {
        if (alumno && alumno.folio) return alumno.folio;
        // Si esta persona ya recibió constancia de este evento, conserva su folio (historial)
        const continuar = campo.continuarFolio !== false;
        if (continuar && alumno) {
            const r = Historial.de(alumno);
            if (r && r.folio) return r.folio;
        }
        // Consecutivo DENTRO de su evento: el 1.º alumno de cada taller recibe el número inicial.
        // Con historial, se sigue después del último folio ya emitido (no se repiten).
        let pos = 0;
        if (alumno && indice > 0 && Alumnos.lista[indice] === alumno) pos = this.posicionesFolio(continuar)[indice] || 0;
        let inicio = parseInt(campo.folioInicio, 10);
        if (!Number.isFinite(inicio) || inicio < 0) inicio = 1;
        const maxHist = continuar && alumno ? Historial.maxFolio(alumno.evento) : null;
        if (maxHist !== null) inicio = Math.max(inicio, maxHist + 1);
        return String(inicio + pos).padStart(Math.min(10, Math.max(1, parseInt(campo.folioDigitos, 10) || 1)), '0');
    },

    /**
     * Posición de cada persona dentro de su evento (para el consecutivo), calculada en UNA pasada.
     * Se guarda unos instantes para que generar miles de constancias no la recalcule cada vez.
     */
    posicionesFolio(continuar) {
        const lista = Alumnos.lista;
        const c = this._folioCache;
        const firma = lista.length + '|' + Historial.registros.size + '|' + continuar;
        if (c && c.lista === lista && c.firma === firma && Date.now() - c.t < 400) return c.pos;
        const usarHist = continuar && Historial.registros.size > 0;
        const cuenta = new Map(), pos = new Array(lista.length);
        for (let j = 0; j < lista.length; j++) {
            const a = lista[j];
            const ev = normalizar(a.evento);
            const n = cuenta.get(ev) || 0;
            pos[j] = n;
            if (usarHist && !a.folio && Historial.de(a)) continue; // ya tiene su folio guardado: no ocupa número
            cuenta.set(ev, n + 1);
        }
        this._folioCache = { lista, firma, pos, t: Date.now() };
        return pos;
    },

    // ------------------------------------------------------------------
    // AUTO-AJUSTE: tamaño final (pt) para que el texto quepa en el ancho máximo
    // Se usa EXACTAMENTE la misma medición en la vista previa y en el PDF.
    // ------------------------------------------------------------------
    docMedidas() {
        if (!this._docMedidas) {
            this._docMedidas = new window.jspdf.jsPDF({ unit: 'mm', format: [PDF_WIDTH_MM, 215.9] });
        }
        return this._docMedidas;
    },

    ponerFuente(doc, fuente, estilo = 'normal') {
        try {
            if (!Fuentes.asegurarEnDoc(doc, fuente, estilo)) {
                estilo = 'normal';
                if (!Fuentes.asegurarEnDoc(doc, fuente, estilo)) {
                    // La familia no tiene ese estilo ni el normal: se usa una que sí existe (igual en vista y PDF)
                    fuente = Fuentes.familia('CalibriLight') ? 'CalibriLight' : 'helvetica';
                    Fuentes.asegurarEnDoc(doc, fuente, estilo);
                }
            }
            doc.setFont(fuente, estilo);
        } catch (e) {
            console.warn('Fuente no disponible:', fuente, estilo, e.message);
            doc.setFont('helvetica', estilo);
        }
    },

    /**
     * Calcula el tamaño final (pt) y el ancho (mm) del texto con su fuente,
     * estilo y espaciado. El mismo cálculo sirve para la vista previa y el PDF.
     */
    medir(campo, texto, doc = this.docMedidas()) {
        const est = Fuentes.resolver(campo);
        this.ponerFuente(doc, campo.fuente, est.estilo);
        const k = doc.internal.scaleFactor;                       // pt por mm
        const esp = (parseFloat(campo.espaciado) || 0) / 100;     // fracción del tamaño
        const unidad = texto ? doc.getStringUnitWidth(texto) : 0; // ancho por cada pt de tamaño
        const n = texto.length;
        let pt = campo.tamano * FONT_SIZE_CORRECTION;
        const max = parseFloat(campo.anchoMax) || 0;
        const anchoDe = (p) => (p / k) * (unidad + n * esp);
        if (max > 0 && texto && anchoDe(pt) > max) {
            pt = Math.min(pt, Math.max(MIN_AUTOFIT_PT, (max * k) / (unidad + n * esp)));
            pt = Math.floor(pt * 2) / 2; // redondear hacia abajo a 0.5
        }
        return { pt, ancho: anchoDe(pt), alto: pt / k, espacio: (pt / k) * esp, est };
    },

    tamanoFinalPt(campo, texto, doc = this.docMedidas()) {
        return this.medir(campo, texto, doc).pt;
    },

    // ------------------------------------------------------------------
    // PÁRRAFOS CON DATOS DE LA LISTA ("por su participación en el {Taller}…")
    // ------------------------------------------------------------------
    /** Valor de una variable {Algo} para un alumno (o un ejemplo si no hay alumno) */
    valorVariable(clave, alumno, indice) {
        const n = normalizarEncabezado(clave);
        if (!n) return null;
        if (n === 'hoy') return fechaLarga();
        const tipo = Alumnos.tipoDeEncabezado(clave);
        const curso = this.campo('curso'), fecha = this.campos.find(c => c.tipo === 'fecha');
        switch (tipo) {
            case 'nombre': return alumno ? alumno.nombre.replace(/,/g, '').replace(/\s+/g, ' ').trim() : 'Nombre del Alumno';
            case 'codigo': return alumno ? (alumno.codigo || '') : '219000000';
            case 'correo': return alumno ? (alumno.correo || '') : 'alumno@correo.com';
            case 'evento': return (alumno && alumno.evento) || (curso ? curso.valor : '') || '';
            case 'folio': return this.folioDe(alumno, indice);
            case 'fecha': return (alumno && alumno.fecha) || (fecha ? fecha.valor : '') || fechaLarga();
        }
        const col = Alumnos.columnasExtra.find(c => normalizarEncabezado(c) === n);
        if (alumno) return col ? ((alumno.extra || {})[col] || '') : null;
        return col ? `[${col}]` : null;
    },

    /** Variables del párrafo que no corresponden a ningún dato conocido */
    variablesDesconocidas(campo) {
        const fuera = [];
        String(campo.valor || '').replace(/\{([^{}]+)\}/g, (t, k) => {
            if (this.valorVariable(k, null, 0) === null && !fuera.includes(k)) fuera.push(k);
            return t;
        });
        return fuera;
    },

    /** Partes del párrafo: [{ t, b }] (b = resaltado en negrita, viene de la lista) */
    partesParrafo(campo, alumno, indice) {
        const partes = [];
        const plantilla = String(campo.valor || '');
        let ultimo = 0;
        plantilla.replace(/\{([^{}]+)\}/g, (todo, k, pos) => {
            if (pos > ultimo) partes.push({ t: plantilla.slice(ultimo, pos), b: false });
            let v = this.valorVariable(k, alumno, indice);
            if (v === null) v = alumno ? '' : todo; // en la vista de ejemplo se ve la variable desconocida
            partes.push({ t: String(v), b: !!campo.resaltar });
            ultimo = pos + todo.length;
            return todo;
        });
        if (ultimo < plantilla.length) partes.push({ t: plantilla.slice(ultimo), b: false });
        for (const p of partes) p.t = aplicarMayusculas(p.t, campo.mayusculas);
        return partes;
    },

    /**
     * Acomoda el párrafo en renglones dentro del ancho (mm) del campo.
     * El mismo cálculo se usa en la vista previa y en el PDF.
     * @returns { pt, lineH (mm), lineas: [{ ancho, segs: [{ t, b, x }] }] }
     */
    layoutParrafo(campo, alumno, indice, doc = this.docMedidas()) {
        const partes = this.partesParrafo(campo, alumno, indice);
        const k = doc.internal.scaleFactor;
        const estilos = {
            n: Fuentes.resolver(campo),
            b: Fuentes.resolver(Object.assign({}, campo, { negrita: true }))
        };
        const unidades = {};
        const unidad = (t, b) => {
            const clave = (b ? 'b' : 'n') + t;
            if (!(clave in unidades)) {
                this.ponerFuente(doc, campo.fuente, estilos[b ? 'b' : 'n'].estilo);
                unidades[clave] = doc.getStringUnitWidth(t);
            }
            return unidades[clave];
        };
        // Palabras y saltos de línea, cada una con su estilo
        const fichas = [];
        for (const p of partes) {
            for (const trozo of p.t.split(/(\n)/)) {
                if (trozo === '\n') { fichas.push({ salto: true }); continue; }
                for (const pal of trozo.split(/(\s+)/)) {
                    if (!pal) continue;
                    fichas.push(/^\s+$/.test(pal) ? { espacio: true, b: p.b } : { t: pal, b: p.b });
                }
            }
        }
        const maxMm = parseFloat(campo.anchoMax) || 0;
        const armar = (pt) => {
            const mm = (u) => (u * pt) / k;
            const lineas = [];
            let actual = { segs: [], ancho: 0 }, espacioPendiente = null;
            const cerrar = () => { lineas.push(actual); actual = { segs: [], ancho: 0 }; espacioPendiente = null; };
            for (const f of fichas) {
                if (f.salto) { cerrar(); continue; }
                if (f.espacio) { if (actual.segs.length) espacioPendiente = f.b; continue; }
                const w = mm(unidad(f.t, f.b));
                const wEsp = espacioPendiente !== null ? mm(unidad(' ', espacioPendiente)) : 0;
                if (actual.segs.length && maxMm > 0 && actual.ancho + wEsp + w > maxMm + 0.01) cerrar();
                if (actual.segs.length && espacioPendiente !== null) {
                    actual.segs.push({ t: ' ', b: espacioPendiente, x: actual.ancho, w: wEsp });
                    actual.ancho += wEsp;
                }
                actual.segs.push({ t: f.t, b: f.b, x: actual.ancho, w });
                actual.ancho += w;
                espacioPendiente = null;
            }
            if (actual.segs.length || !lineas.length) lineas.push(actual);
            // Juntar trozos seguidos con el mismo estilo (menos llamadas al dibujar)
            for (const l of lineas) {
                const unidos = [];
                for (const s of l.segs) {
                    const u = unidos[unidos.length - 1];
                    if (u && u.b === s.b) { u.t += s.t; u.w += s.w; } else unidos.push(Object.assign({}, s));
                }
                l.segs = unidos;
            }
            return lineas;
        };
        let pt = campo.tamano * FONT_SIZE_CORRECTION;
        let lineas = armar(pt);
        const maxR = parseInt(campo.maxRenglones, 10) || 0;
        // Los saltos de línea escritos con Enter no se pueden quitar achicando la letra
        const minimo = Math.max(maxR, (String(campo.valor || '').match(/\n/g) || []).length + 1);
        while (maxR > 0 && lineas.length > minimo && pt > MIN_AUTOFIT_PT) {
            pt = Math.max(MIN_AUTOFIT_PT, pt - 0.5);
            lineas = armar(pt);
        }
        const inter = parseFloat(campo.interlineado) || 1.3;
        return { pt, lineH: (pt / k) * inter, alto: pt / k, lineas, estilos };
    },

    /** Dibuja el párrafo en el PDF. (x, y) = punto de anclaje; y es el borde de arriba del párrafo */
    dibujarParrafo(doc, campo, alumno, indice, x, y) {
        const L = this.layoutParrafo(campo, alumno, indice, doc);
        doc.setFontSize(L.pt);
        doc.setTextColor(campo.color);
        doc.setDrawColor(campo.color);
        const alturaPagina = doc.internal.pageSize.getHeight();
        L.lineas.forEach((linea, i) => {
            const yMid = y + L.lineH * i + L.lineH / 2;
            let x0 = x;
            if (campo.alineacion === 'center') x0 = x - linea.ancho / 2;
            else if (campo.alineacion === 'right') x0 = x - linea.ancho;
            for (const s of linea.segs) {
                if (!s.t.trim()) continue;
                const est = L.estilos[s.b ? 'b' : 'n'];
                this.ponerFuente(doc, campo.fuente, est.estilo);
                doc.setFontSize(L.pt);
                const op = { baseline: 'middle' };
                let comp = 0;
                if (est.fauxItalic) {
                    op.angle = new doc.Matrix(1, 0, FAUX_ITALIC, 1, 0, 0);
                    const lineaBase = yMid + L.alto / 2 - L.alto * (doc.getLineHeightFactor() - 1);
                    comp = FAUX_ITALIC * (alturaPagina - lineaBase);
                }
                if (est.fauxBold) {
                    op.renderingMode = 'fillThenStroke';
                    doc.setLineWidth(L.alto * FAUX_BOLD);
                } else {
                    op.renderingMode = 'fill';
                }
                doc.text(s.t, x0 + s.x - comp, yMid, op);
            }
        });
    },

    // ------------------------------------------------------------------
    // CÓDIGO QR DE VERIFICACIÓN
    // ------------------------------------------------------------------
    /** Lo que guarda el QR: un texto con los datos o una dirección de la página de verificación */
    contenidoQR(campo, alumno, indice) {
        const folio = this.folioDe(alumno, indice);
        const codigo = Historial.codigo(alumno, folio);
        if (campo.qrModo === 'url' && campo.qrUrl) {
            return campo.qrUrl + (campo.qrUrl.includes('?') ? '&' : '?') + 'c=' + limpiarCodigo(codigo);
        }
        const nombre = alumno ? alumno.nombre : 'Nombre del Alumno';
        const evento = (alumno && alumno.evento) || this.campo('curso')?.valor || '';
        const fecha = (alumno && alumno.fecha) || this.campos.find(c => c.tipo === 'fecha')?.valor || '';
        return ['CONSTANCIA' + (evento ? ' · ' + evento : ''), 'Folio: ' + folio, 'Nombre: ' + nombre,
            fecha ? 'Fecha: ' + fecha : '', 'Verificación: ' + codigo, campo.valor ? 'Emitida por: ' + campo.valor : '']
            .filter(Boolean).join('\n');
    },

    _qrCache: new Map(),
    /** Matriz del QR: { n, oscuro(r, c) } */
    matrizQR(texto) {
        if (this._qrCache.has(texto)) return this._qrCache.get(texto);
        if (typeof qrcode === 'undefined') throw new Error('Falta la librería del código QR (librerias/qrcode.js).');
        qrcode.stringToBytes = qrcode.stringToBytesFuncs['UTF-8'];
        const q = qrcode(0, 'M');
        q.addData(texto, 'Byte');
        q.make();
        const m = { n: q.getModuleCount(), oscuro: (r, c) => q.isDark(r, c) };
        if (this._qrCache.size > 400) this._qrCache.clear();
        this._qrCache.set(texto, m);
        return m;
    },

    leyendaQR(alumno, indice) {
        const folio = this.folioDe(alumno, indice);
        return `Folio ${folio} · Verificación ${Historial.codigo(alumno, folio)}`;
    },

    /** Dibuja el QR en el PDF. (x, y) = centro del cuadro */
    dibujarQR(doc, campo, alumno, indice, x, y) {
        const texto = this.contenidoQR(campo, alumno, indice);
        const q = this.matrizQR(texto);
        const lado = campo.lado;
        const mod = lado / q.n;
        const x0 = x - lado / 2, y0 = y - lado / 2;
        if (campo.qrFondo) {
            doc.setFillColor('#ffffff');
            doc.rect(x0 - mod * 1.5, y0 - mod * 1.5, lado + mod * 3, lado + mod * 3, 'F');
        }
        doc.setFillColor(campo.color || '#000000');
        for (let r = 0; r < q.n; r++) {
            let c = 0;
            while (c < q.n) {
                if (!q.oscuro(r, c)) { c++; continue; }
                let fin = c;
                while (fin < q.n && q.oscuro(r, fin)) fin++;
                doc.rect(x0 + c * mod, y0 + r * mod, (fin - c) * mod, mod + 0.01, 'F');
                c = fin;
            }
        }
        if (campo.qrLeyenda) {
            const pt = Math.max(5, lado * 0.3);
            this.ponerFuente(doc, 'helvetica', 'normal');
            doc.setFontSize(pt);
            doc.setTextColor(campo.color || '#000000');
            const k = doc.internal.scaleFactor;
            doc.text(this.leyendaQR(alumno, indice), x, y0 + lado + (campo.qrFondo ? mod * 1.5 : 0) + (pt / k) * 0.9,
                { align: 'center', baseline: 'middle' });
        }
    },

    elementoQR(campo, alumno, pxPorMm) {
        const texto = this.contenidoQR(campo, alumno, this.indiceVista);
        const q = this.matrizQR(texto);
        const ladoPx = campo.lado * pxPorMm;
        let d = '';
        for (let r = 0; r < q.n; r++) {
            let c = 0;
            while (c < q.n) {
                if (!q.oscuro(r, c)) { c++; continue; }
                let fin = c;
                while (fin < q.n && q.oscuro(r, fin)) fin++;
                d += `M${c} ${r}h${fin - c}v1h${-(fin - c)}z`;
                c = fin;
            }
        }
        const el = document.createElement('div');
        el.style.width = ladoPx + 'px';
        const pad = campo.qrFondo ? 1.5 : 0;
        el.innerHTML = `<svg viewBox="${-pad} ${-pad} ${q.n + pad * 2} ${q.n + pad * 2}" width="${ladoPx * (q.n + pad * 2) / q.n}" height="${ladoPx * (q.n + pad * 2) / q.n}"
            style="display:block;margin:${-ladoPx * pad / q.n}px" shape-rendering="crispEdges">
            ${campo.qrFondo ? `<rect x="${-pad}" y="${-pad}" width="${q.n + pad * 2}" height="${q.n + pad * 2}" fill="#fff"/>` : ''}
            <path d="${d}" fill="${escapeHtml(campo.color || '#000')}"/></svg>`;
        if (campo.qrLeyenda) {
            const ley = document.createElement('div');
            const ptPx = Math.max(5, campo.lado * 0.3) * (pxPorMm / PT_POR_MM);
            ley.style.cssText = `font:${ptPx}px Helvetica,Arial,sans-serif;color:${campo.color || '#000'};white-space:nowrap;text-align:center;` +
                `width:0;margin-left:${ladoPx / 2}px;display:flex;justify-content:center;margin-top:${ptPx * 0.4 + (campo.qrFondo ? ladoPx * 1.5 / q.n : 0)}px;line-height:1`;
            ley.innerHTML = `<span>${escapeHtml(this.leyendaQR(alumno, this.indiceVista))}</span>`;
            el.appendChild(ley);
        }
        return el;
    },

    /** Vista previa del párrafo (mismos renglones que el PDF) */
    elementoParrafo(campo, alumno, pxPorMm, pxPorPt) {
        const L = this.layoutParrafo(campo, alumno, this.indiceVista);
        const el = document.createElement('div');
        const hPx = L.pt * pxPorPt;
        const lineHPx = L.lineH * pxPorMm;
        el.style.fontSize = hPx + 'px';
        el.style.fontFamily = Fuentes.css(campo.fuente);
        el.style.color = campo.color;
        el.style.fontStyle = L.estilos.n.estilo.includes('italic') ? 'italic' : 'normal';
        const maxMm = parseFloat(campo.anchoMax) || 0;
        const anchoMm = maxMm > 0 ? maxMm : Math.max(...L.lineas.map(l => l.ancho), 1);
        el.style.width = (anchoMm * pxPorMm) + 'px';
        el.style.textAlign = { left: 'left', center: 'center', right: 'right' }[campo.alineacion] || 'center';
        for (const linea of L.lineas) {
            const d = document.createElement('div');
            d.style.height = lineHPx + 'px';
            d.style.lineHeight = lineHPx + 'px';
            d.style.whiteSpace = 'pre';
            for (const s of linea.segs) {
                const sp = document.createElement('span');
                sp.textContent = s.t;
                const est = L.estilos[s.b ? 'b' : 'n'];
                sp.style.fontWeight = est.estilo.includes('bold') ? '700' : '400';
                sp.style.fontStyle = est.estilo.includes('italic') ? 'italic' : 'normal';
                if (est.fauxBold) sp.style.webkitTextStroke = `${hPx * FAUX_BOLD}px ${campo.color}`;
                if (est.fauxItalic) {
                    sp.style.display = 'inline-block';
                    sp.style.transform = `skewX(${-Math.atan(FAUX_ITALIC) * 180 / Math.PI}deg)`;
                }
                d.appendChild(sp);
            }
            el.appendChild(d);
        }
        return { el, pt: L.pt, renglones: L.lineas.length };
    },

    // ------------------------------------------------------------------
    // VISTA PREVIA: dibujar los campos sobre la imagen
    // ------------------------------------------------------------------
    alumnoActual() {
        if (this.manual) return this.manual;
        return this.indiceVista >= 0 ? Alumnos.lista[this.indiceVista] : null;
    },

    /** Índices de los registros del evento elegido (o todos) */
    indicesActivos() {
        const idx = [];
        Alumnos.lista.forEach((a, i) => {
            if (this.filtroEvento === '__todos__' || normalizar(a.evento) === this.filtroEvento) idx.push(i);
        });
        return idx;
    },

    /** Llena el selector de eventos (solo se muestra si el CSV trae eventos) */
    renderFiltroEventos() {
        const eventos = Alumnos.eventos();
        const hayEventos = eventos.some(e => e.nombre);
        $('event-filter-wrap').style.display = hayEventos ? 'flex' : 'none';
        if (!hayEventos) this.filtroEvento = '__todos__';
        const sel = $('event-filter');
        sel.innerHTML = `<option value="__todos__">Todos los eventos (${Alumnos.lista.length})</option>` +
            eventos.map(e => `<option value="${escapeHtml(normalizar(e.nombre))}">${escapeHtml(e.nombre || '(sin evento)')} (${e.total})</option>`).join('');
        if (![...sel.options].some(o => o.value === this.filtroEvento)) this.filtroEvento = '__todos__';
        sel.value = this.filtroEvento;
    },

    renderCampos() {
        const capa = $('fields-layer');
        const anchoPx = $('certificate-canvas').clientWidth || 0;

        // Quitar los elementos anteriores (excepto las guías de centro)
        capa.querySelectorAll('.field-el, .maxwidth-guide').forEach(e => e.remove());
        if (!this.fondo || anchoPx === 0) return;

        const hoja = this.dimensionesPagina();
        const pxPorMm = anchoPx / hoja.w;
        const pxPorPt = pxPorMm / PT_POR_MM;
        const alumno = this.alumnoActual();

        for (const campo of this.campos) {
            if (!campo.activo) continue;
            let texto = this.textoDeCampo(campo, alumno, this.indiceVista);
            const vacio = !texto;
            if (vacio) texto = `[${campo.etiqueta} vacío]`;

            if (campo.tipo === 'imagen') {
                if (!campo.imagen.startsWith('data:')) continue;
                const el = document.createElement('div');
                const w = campo.anchoImg * pxPorMm, h = w * campo.proporcion;
                el.className = 'field-el imagen' + (campo.id === this.seleccionado ? ' selected' : '');
                el.dataset.id = campo.id;
                el.style.cssText += `left:${campo.x}%;top:${campo.y}%;width:${w}px;height:${h}px;transform:translate(-50%,-50%);` +
                    `background:url('${campo.imagen}') center/100% 100% no-repeat;`;
                el.title = campo.etiqueta;
                el.addEventListener('mousedown', (e) => this.iniciarArrastre(e, campo.id));
                capa.appendChild(el);
                continue;
            }
            if (campo.tipo === 'qr') {
                let el;
                try { el = this.elementoQR(campo, alumno, pxPorMm); } catch (e) { console.warn(e); continue; }
                el.className = 'field-el qr' + (campo.id === this.seleccionado ? ' selected' : '');
                el.dataset.id = campo.id;
                el.style.left = campo.x + '%';
                el.style.top = campo.y + '%';
                el.style.transform = `translate(-50%, ${-campo.lado * pxPorMm / 2}px)`;
                el.style.whiteSpace = 'normal';
                el.title = 'Código QR de verificación';
                el.addEventListener('mousedown', (e) => this.iniciarArrastre(e, campo.id));
                capa.appendChild(el);
                continue;
            }
            if (campo.tipo === 'parrafo' && !vacio) {
                const { el, renglones } = this.elementoParrafo(campo, alumno, pxPorMm, pxPorPt);
                el.className = 'field-el parrafo' + (campo.id === this.seleccionado ? ' selected' : '');
                el.dataset.id = campo.id;
                el.style.left = campo.x + '%';
                el.style.top = campo.y + '%';
                const tx = { left: '0', center: '-50%', right: '-100%' }[campo.alineacion] || '-50%';
                el.style.transform = `translate(${tx}, 0)`;
                el.style.whiteSpace = 'normal';
                el.title = `${campo.etiqueta} · ${renglones} renglón(es)`;
                el.addEventListener('mousedown', (e) => this.iniciarArrastre(e, campo.id));
                capa.appendChild(el);
                continue;
            }
            const m = this.medir(campo, vacio ? '' : texto);
            const pt = m.pt;
            const hPx = pt * pxPorPt;
            const el = document.createElement('div');
            el.className = 'field-el' + (campo.id === this.seleccionado ? ' selected' : '');
            el.dataset.id = campo.id;
            el.textContent = texto;
            el.style.left = campo.x + '%';
            el.style.top = campo.y + '%';
            el.style.fontSize = hPx + 'px';
            el.style.fontFamily = Fuentes.css(campo.fuente);
            el.style.fontWeight = m.est.estilo.includes('bold') ? '700' : '400';
            el.style.fontStyle = m.est.estilo.includes('italic') ? 'italic' : 'normal';
            el.style.color = campo.color;
            el.style.opacity = vacio ? '0.45' : '1';
            el.style.letterSpacing = ((parseFloat(campo.espaciado) || 0) / 100) + 'em';
            if (m.est.fauxBold) el.style.webkitTextStroke = `${hPx * FAUX_BOLD}px ${campo.color}`;
            if (campo.subrayado) {
                el.style.textDecoration = 'underline';
                el.style.textDecorationThickness = '0.05em';
                el.style.textUnderlineOffset = '0.1em';
            }
            if (campo.sombra) {
                const d = hPx * (campo.sombraDist || 0) / 100;
                el.style.textShadow = `${d}px ${d}px 0 ${hexARgba(campo.sombraColor, (campo.sombraOpacidad ?? 45) / 100)}`;
            }
            const tx = { left: '0', center: '-50%', right: '-100%' }[campo.alineacion] || '-50%';
            el.style.transformOrigin = '50% 80%';
            el.style.transform = `translate(${tx}, -50%)` + (m.est.fauxItalic ? ` skewX(${-Math.atan(FAUX_ITALIC) * 180 / Math.PI}deg)` : '');
            el.title = campo.etiqueta + (pt < campo.tamano * FONT_SIZE_CORRECTION
                ? ' (reducido automáticamente para que quepa)' : '');
            el.addEventListener('mousedown', (e) => this.iniciarArrastre(e, campo.id));
            capa.appendChild(el);

            // Guía roja del ancho máximo para el campo seleccionado
            if (campo.id === this.seleccionado && campo.anchoMax > 0 && campo.tipo !== 'parrafo') {
                const g = document.createElement('div');
                g.className = 'maxwidth-guide';
                const w = campo.anchoMax * pxPorMm;
                const xPx = (campo.x / 100) * anchoPx;
                const left = campo.alineacion === 'left' ? xPx
                    : campo.alineacion === 'right' ? xPx - w : xPx - w / 2;
                const h = campo.tamano * FONT_SIZE_CORRECTION * pxPorPt * 1.3;
                g.style.left = left + 'px';
                g.style.width = w + 'px';
                g.style.top = `calc(${campo.y}% - ${h / 2}px)`;
                g.style.height = h + 'px';
                capa.appendChild(g);
            }
        }
    },

    // ------------------------------------------------------------------
    // ARRASTRAR CAMPOS CON EL MOUSE
    // ------------------------------------------------------------------
    iniciarArrastre(e, id) {
        e.preventDefault();
        if (this.seleccionado !== id) this.seleccionar(id);

        const campo = this.campo(id);
        const rect = $('fields-layer').getBoundingClientRect();
        const inicio = { mx: e.clientX, my: e.clientY, x: campo.x, y: campo.y };
        const el = $('fields-layer').querySelector(`.field-el[data-id="${id}"]`);
        if (el) el.classList.add('dragging');
        const gv = $('guide-v'), gh = $('guide-h');

        const mover = (ev) => {
            let x = inicio.x + ((ev.clientX - inicio.mx) / rect.width) * 100;
            let y = inicio.y + ((ev.clientY - inicio.my) / rect.height) * 100;
            x = Math.min(100, Math.max(0, x));
            y = Math.min(100, Math.max(0, y));
            // "Imán" al centro de la hoja (desactivable con Alt)
            const iman = !ev.altKey;
            const enCentroX = iman && Math.abs(x - 50) < 0.8;
            const enCentroY = iman && Math.abs(y - 50) < 0.8;
            if (enCentroX) x = 50;
            if (enCentroY) y = 50;
            gv.style.display = enCentroX ? 'block' : 'none';
            gh.style.display = enCentroY ? 'block' : 'none';

            campo.x = Math.round(x * 100) / 100;
            campo.y = Math.round(y * 100) / 100;
            const actual = $('fields-layer').querySelector(`.field-el[data-id="${id}"]`);
            if (actual) {
                actual.style.left = campo.x + '%';
                actual.style.top = campo.y + '%';
            }
            this.actualizarInputsPosicion();
        };

        const soltar = () => {
            document.removeEventListener('mousemove', mover);
            document.removeEventListener('mouseup', soltar);
            gv.style.display = gh.style.display = 'none';
            this.renderCampos();
            this.guardarEstado();
        };

        document.addEventListener('mousemove', mover);
        document.addEventListener('mouseup', soltar);
    },

    // ------------------------------------------------------------------
    // PANEL DERECHO: LISTA DE CAMPOS
    // ------------------------------------------------------------------
    renderListaCampos() {
        const cont = $('fields-list');
        cont.innerHTML = '';
        for (const campo of this.campos) {
            const fila = document.createElement('div');
            fila.className = 'field-row' + (campo.id === this.seleccionado ? ' selected' : '')
                + (campo.activo ? '' : ' disabled');
            fila.innerHTML = `
                <input type="checkbox" ${campo.activo ? 'checked' : ''} title="Mostrar / ocultar en la constancia">
                <span class="field-row-label">${escapeHtml(campo.etiqueta)}</span>
                <span class="field-row-type">${FIELD_TYPE_LABEL[campo.tipo] || ''}</span>`;
            const check = fila.querySelector('input');
            check.addEventListener('click', ev => ev.stopPropagation());
            check.addEventListener('change', () => {
                campo.activo = check.checked;
                if (campo.activo) this.seleccionado = campo.id;
                this.refrescarTodo();
            });
            fila.addEventListener('click', () => this.seleccionar(campo.id));
            cont.appendChild(fila);
        }
    },

    seleccionar(id) {
        this.seleccionado = id;
        this.renderListaCampos();
        this.renderPropiedades();
        this.renderCampos();
    },

    refrescarTodo() {
        this.renderListaCampos();
        this.renderPropiedades();
        this.renderCampos();
        this.actualizarResumen();
        this.guardarEstado();
    },

    // ------------------------------------------------------------------
    // PANEL DERECHO: PROPIEDADES DEL CAMPO SELECCIONADO
    // ------------------------------------------------------------------
    renderPropiedades() {
        const campo = this.campoSel;
        const panel = $('props-panel');
        if (!campo) {
            panel.classList.add('disabled-panel');
            $('prop-title').textContent = '';
            $('prop-content').innerHTML = '';
            return;
        }
        panel.classList.toggle('disabled-panel', false);
        $('prop-title').textContent = campo.etiqueta;

        // Contenido específico según el tipo de campo
        const c = $('prop-content');
        let html = '';
        if (!campo.activo) {
            html += `<div class="info-box">Este campo está oculto. Actívalo con la casilla en la lista de campos.</div>`;
        }
        switch (campo.tipo) {
            case 'nombre':
                html += `<div class="info-box">Se toma el <b>nombre</b> de cada alumno de la lista.</div>`;
                break;
            case 'codigo':
                html += `<div class="info-box">Se toma el <b>código</b> de cada alumno. Si un alumno no tiene código, este campo no se imprime en su constancia.</div>
                    <div class="stack"><span class="label">Texto antes del código</span>
                    <input id="pc-prefijo" class="input" placeholder="Ej. Código: "></div>`;
                break;
            case 'folio':
                html += `<div class="info-box">Si el CSV tiene columna <b>Folio</b> se usa la de cada alumno. Si no, se pone un número consecutivo según el orden de la lista (si hay varios eventos, cada evento empieza desde el número inicial).</div>
                    <div class="stack"><span class="label">Texto antes del número</span>
                    <input id="pc-prefijo" class="input" placeholder="Ej. Folio: CUCSUR-2026-"></div>
                    <div class="row">
                        <div class="stack"><span class="label">Inicia en</span>
                        <input id="pc-inicio" type="number" min="0" class="input"></div>
                        <div class="stack"><span class="label">Dígitos</span>
                        <input id="pc-digitos" type="number" min="1" max="10" class="input"></div>
                    </div>
                    <label class="toolbar-check"><input type="checkbox" id="pc-continuar"> No repetir folios ya emitidos (historial)</label>
                    <span class="hint">Quien ya recibió constancia de ese evento conserva su folio; los nuevos siguen después del último.</span>`;
                break;
            case 'evento':
                html += `<div class="info-box">Si el CSV tiene columna <b>Evento / Taller / Curso</b> se usa la de cada alumno (útil cuando hay varios talleres en un archivo). Si no, se usa este texto:</div>
                    <div class="stack"><span class="label">Texto del evento</span>
                    <input id="pc-valor" class="input" placeholder="Nombre del curso o evento"></div>`;
                break;
            case 'columna': {
                const opciones = Alumnos.columnasExtra.map(c =>
                    `<option value="${escapeHtml(c)}">${escapeHtml(c)}</option>`).join('');
                html += `<div class="info-box">Imprime el valor de otra columna de tu CSV (por ejemplo horas, instructor o calificación).</div>
                    <div class="stack"><span class="label">Columna del CSV</span>
                    <select id="pc-columna" class="input">${opciones || ''}${campo.columna && !Alumnos.columnasExtra.includes(campo.columna)
                        ? `<option value="${escapeHtml(campo.columna)}">${escapeHtml(campo.columna)} (no cargada)</option>` : ''}</select></div>
                    <div class="stack"><span class="label">Texto antes del valor</span>
                    <input id="pc-prefijo" class="input" placeholder="Ej. Duración (horas): "></div>
                    <button id="pc-eliminar" class="btn small danger">🗑️ Eliminar este campo</button>`;
                break;
            }
            case 'fecha':
                html += `<div class="info-box">Si el CSV tiene columna <b>Fecha</b> se usa la de cada alumno; si no, este texto.</div>
                    <div class="stack"><span class="label">Texto de la fecha</span>
                    <input id="pc-valor" class="input">
                    <button id="pc-hoy" class="btn small ghost">📅 Poner la fecha de hoy</button></div>`;
                break;
            case 'imagen':
                html += `<div class="info-box">Imagen encima del fondo (firma escaneada, sello, logo). Arrástrala con el mouse.
                    Para firmas, usa un <b>PNG con fondo transparente</b>.</div>
                    <div class="stack"><span class="label">Nombre</span><input id="pc-etiqueta" class="input"></div>
                    <div class="row">
                        <div class="stack"><span class="label">Ancho (mm)</span><input id="pc-anchoimg" type="number" min="3" max="600" step="1" class="input"></div>
                        <div class="stack"><span class="label">Alto (mm)</span><input id="pc-altoimg" class="input" disabled></div>
                    </div>
                    <button id="pc-cambiarimg" class="btn small">🖼️ Cambiar imagen…</button>
                    <button id="pc-eliminar" class="btn small danger">🗑️ Quitar esta imagen</button>`;
                break;
            case 'qr':
                html += `<div class="info-box">El <b>código QR</b> sirve para comprobar que la constancia es auténtica.
                    Cada persona tiene un <b>código de verificación</b> único que queda en el <b>Historial</b>.</div>
                    <div class="stack"><span class="label">Al escanearlo muestra…</span>
                    <select id="pc-qrmodo" class="input">
                        <option value="texto">Los datos de la constancia (funciona sin internet)</option>
                        <option value="url">Una página web de verificación</option>
                    </select></div>
                    <div class="stack" id="pc-qrurl-wrap"><span class="label">Dirección de tu página (https://…)</span>
                    <input id="pc-qrurl" class="input" placeholder="https://tu-usuario.github.io/constancias/">
                    <span class="hint">La página está en la carpeta <b>pagina_verificacion</b> del programa. Súbela a GitHub Pages
                        junto con el archivo <b>registro.json</b> que exportas desde el Historial.</span></div>
                    <div class="stack"><span class="label">Emitida por (opcional)</span>
                    <input id="pc-valor" class="input" maxlength="120" placeholder="Ej. CUCSur · Universidad de Guadalajara"></div>
                    <div class="row">
                        <div class="stack"><span class="label">Tamaño (mm)</span><input id="pc-lado" type="number" min="8" max="120" step="1" class="input"></div>
                    </div>
                    <label class="toolbar-check"><input type="checkbox" id="pc-qrley"> Mostrar folio y código debajo</label>
                    <label class="toolbar-check"><input type="checkbox" id="pc-qrfondo"> Fondo blanco (se lee mejor)</label>
                    <button id="pc-historial" class="btn small">🗂 Abrir historial y verificar</button>
                    <button id="pc-eliminar" class="btn small danger">🗑️ Quitar el código QR</button>`;
                break;
            case 'parrafo': {
                const datos = ['Nombre', 'Código', 'Evento', 'Fecha', 'Folio', 'Correo', ...Alumnos.columnasExtra];
                html += `<div class="info-box">Escribe el texto y pon los <b>datos de la lista entre llaves</b>, por ejemplo
                    <b>{Taller}</b> o <b>{Horas}</b>. Cada persona recibe sus propios datos y los renglones se acomodan solos.</div>
                    <div class="stack"><span class="label">Nombre del campo</span>
                    <input id="pc-etiqueta" class="input"></div>
                    <div class="stack"><span class="label">Texto del párrafo</span>
                    <textarea id="pc-parrafo" class="input" rows="5" style="resize:vertical;min-height:90px;line-height:1.4;padding:6px 8px"></textarea></div>
                    <div class="stack"><span class="label">Insertar un dato (clic)</span>
                    <div class="chips" id="pc-chips">${datos.map(d => `<span class="chip" data-var="${escapeHtml(d)}">{${escapeHtml(d)}}</span>`).join('')}
                    <span class="chip" data-var="Hoy" title="Fecha del día en que se genera">{Hoy}</span></div></div>
                    <div id="pc-desconocidas" class="info-box" style="display:none;border-color:#f0c27a;background:#fff9ef"></div>
                    <label class="toolbar-check"><input type="checkbox" id="pc-resaltar"> Datos de la lista en <b>negrita</b></label>
                    <div class="row">
                        <div class="stack"><span class="label">Ancho (mm)</span><input id="pc-ancho" type="number" min="20" max="600" step="5" class="input"></div>
                        <div class="stack"><span class="label">Interlineado</span><input id="pc-inter" type="number" min="0.8" max="3" step="0.05" class="input"></div>
                        <div class="stack"><span class="label">Máx. renglones</span><input id="pc-maxr" type="number" min="0" max="30" class="input" title="0 = sin límite"></div>
                    </div>
                    <span class="hint">Si el texto pasa del máximo de renglones, la letra se reduce sola. En párrafos no se aplican subrayado, sombra ni espacio entre letras.</span>
                    <button id="pc-eliminar" class="btn small danger">🗑️ Eliminar este campo</button>`;
                break;
            }
            default:
                html += `<div class="stack"><span class="label">Nombre del campo</span>
                    <input id="pc-etiqueta" class="input"></div>
                    <div class="stack"><span class="label">Texto</span>
                    <input id="pc-valor" class="input" placeholder="Escribe el texto"></div>`;
                if (campo.id !== 'curso') {
                    html += `<button id="pc-eliminar" class="btn small danger">🗑️ Eliminar este campo</button>`;
                }
        }
        c.innerHTML = html;

        // Conectar los inputs específicos
        const enlazar = (idInput, prop, tipoNum = false) => {
            const inp = $(idInput);
            if (!inp) return;
            inp.value = campo[prop];
            inp.addEventListener('input', () => {
                if (tipoNum) {
                    const v = parseInt(inp.value, 10);
                    if (!Number.isFinite(v)) return;
                    campo[prop] = prop === 'folioDigitos' ? Math.min(10, Math.max(1, v)) : Math.min(1e9, Math.max(0, v));
                } else {
                    campo[prop] = String(inp.value).substring(0, prop === 'valor' ? 300 : 120);
                }
                if (prop === 'etiqueta') {
                    $('prop-title').textContent = campo.etiqueta;
                    this.renderListaCampos();
                }
                this.renderCampos();
                this.guardarEstado();
            });
        };
        enlazar('pc-prefijo', 'prefijo');
        enlazar('pc-inicio', 'folioInicio', true);
        enlazar('pc-digitos', 'folioDigitos', true);
        enlazar('pc-valor', 'valor');
        enlazar('pc-etiqueta', 'etiqueta');
        if ($('pc-columna')) {
            $('pc-columna').value = campo.columna;
            $('pc-columna').addEventListener('change', e => {
                campo.columna = e.target.value;
                campo.etiqueta = 'Columna: ' + campo.columna;
                $('prop-title').textContent = campo.etiqueta;
                this.renderListaCampos();
                this.renderCampos();
                this.guardarEstado();
            });
        }
        if ($('pc-parrafo')) {
            const ta = $('pc-parrafo');
            ta.value = campo.valor;
            const avisar = () => {
                const fuera = this.variablesDesconocidas(campo);
                const box = $('pc-desconocidas');
                box.style.display = fuera.length ? 'block' : 'none';
                box.innerHTML = fuera.length ? `⚠️ No hay una columna para ${fuera.map(v => `<b>{${escapeHtml(v)}}</b>`).join(', ')}. ` +
                    'Revisa que se escriba igual que el título de la columna en tu lista (o carga la lista primero).' : '';
            };
            avisar();
            ta.addEventListener('input', () => {
                campo.valor = ta.value;
                avisar();
                this.renderCampos();
                this.guardarEstado();
            });
            $('pc-chips').addEventListener('mousedown', e => e.preventDefault()); // no quitar el cursor del cuadro
            $('pc-chips').addEventListener('click', e => {
                const chip = e.target.closest('[data-var]');
                if (!chip) return;
                const ins = `{${chip.dataset.var}}`;
                const a = ta.selectionStart ?? ta.value.length, b = ta.selectionEnd ?? ta.value.length;
                ta.value = ta.value.slice(0, a) + ins + ta.value.slice(b);
                ta.focus();
                ta.selectionStart = ta.selectionEnd = a + ins.length;
                ta.dispatchEvent(new Event('input'));
            });
            const num = (id, prop, min, max, def) => {
                const inp = $(id);
                inp.value = campo[prop];
                inp.addEventListener('input', () => {
                    const v = parseFloat(inp.value);
                    if (isNaN(v)) return;
                    campo[prop] = Math.min(max, Math.max(min, v));
                    if (prop === 'anchoMax') $('prop-maxwidth').value = campo.anchoMax;
                    this.renderCampos();
                    this.guardarEstado();
                });
            };
            num('pc-ancho', 'anchoMax', 20, 600);
            num('pc-inter', 'interlineado', 0.8, 3);
            num('pc-maxr', 'maxRenglones', 0, 30);
            $('pc-resaltar').checked = !!campo.resaltar;
            $('pc-resaltar').addEventListener('change', e => {
                campo.resaltar = e.target.checked;
                this.renderCampos();
                this.guardarEstado();
            });
        }
        if ($('pc-anchoimg')) {
            const alto = () => { $('pc-altoimg').value = (campo.anchoImg * campo.proporcion).toFixed(1); };
            $('pc-anchoimg').value = campo.anchoImg;
            alto();
            $('pc-anchoimg').addEventListener('input', e => {
                const v = parseFloat(e.target.value);
                if (isNaN(v)) return;
                campo.anchoImg = Math.min(600, Math.max(3, v));
                alto();
                this.renderCampos();
                this.guardarEstado();
            });
            $('pc-cambiarimg').addEventListener('click', () => this.elegirImagenCampo(campo.id));
        }
        if ($('pc-continuar')) {
            $('pc-continuar').checked = campo.continuarFolio !== false;
            $('pc-continuar').addEventListener('change', e => {
                campo.continuarFolio = e.target.checked;
                this.renderCampos();
                this.guardarEstado();
            });
        }
        if ($('pc-qrmodo')) {
            const mostrarUrl = () => { $('pc-qrurl-wrap').style.display = campo.qrModo === 'url' ? 'flex' : 'none'; };
            $('pc-qrmodo').value = campo.qrModo;
            $('pc-qrurl').value = campo.qrUrl;
            mostrarUrl();
            $('pc-qrmodo').addEventListener('change', e => {
                campo.qrModo = e.target.value;
                mostrarUrl();
                if (campo.qrModo === 'url' && !campo.qrUrl) UI.toast('Escribe la dirección de tu página de verificación.', 'info');
                this.renderCampos();
                this.guardarEstado();
            });
            $('pc-qrurl').addEventListener('input', e => {
                const v = e.target.value.trim();
                const ok = /^https:\/\/[^\s"'<>]{3,300}$/i.test(v);
                e.target.style.borderColor = v && !ok ? '#c4314b' : '';
                campo.qrUrl = ok ? v : '';
                this.renderCampos();
                this.guardarEstado();
            });
            $('pc-lado').value = campo.lado;
            $('pc-lado').addEventListener('input', e => {
                const v = parseFloat(e.target.value);
                if (isNaN(v)) return;
                campo.lado = Math.min(120, Math.max(8, v));
                this.renderCampos();
                this.guardarEstado();
            });
            for (const [id, prop] of [['pc-qrley', 'qrLeyenda'], ['pc-qrfondo', 'qrFondo']]) {
                $(id).checked = !!campo[prop];
                $(id).addEventListener('change', e => { campo[prop] = e.target.checked; this.renderCampos(); this.guardarEstado(); });
            }
            $('pc-historial').addEventListener('click', () => this.abrirHistorial());
        }
        if ($('pc-hoy')) {
            $('pc-hoy').onclick = () => {
                campo.valor = fechaLarga();
                $('pc-valor').value = campo.valor;
                this.renderCampos();
                this.guardarEstado();
            };
        }
        if ($('pc-eliminar')) {
            $('pc-eliminar').onclick = async () => {
                const ok = await UI.confirmar('Eliminar campo', `¿Eliminar el campo "${campo.etiqueta}"?`, 'Eliminar', true);
                if (!ok) return;
                this.campos = this.campos.filter(x => x.id !== campo.id);
                this.seleccionado = 'nombre';
                this.refrescarTodo();
            };
        }

        // Propiedades comunes
        if (!Fuentes.familia(campo.fuente)) campo.fuente = 'CalibriLight';
        $('prop-font').value = campo.fuente;
        document.querySelectorAll('#prop-styles button').forEach(b =>
            b.classList.toggle('active', !!campo[b.dataset.style]));
        $('shadow-options').style.display = campo.sombra ? 'flex' : 'none';
        $('prop-shadow-color').value = campo.sombraColor;
        $('prop-shadow-dist').value = campo.sombraDist;
        $('prop-shadow-dist-val').textContent = campo.sombraDist + '%';
        $('prop-shadow-op').value = campo.sombraOpacidad;
        $('prop-shadow-op-val').textContent = campo.sombraOpacidad + '%';
        $('prop-spacing').value = campo.espaciado;
        $('prop-spacing-val').textContent = campo.espaciado ? campo.espaciado + '%' : 'normal';
        this.actualizarAyudaEstilo();
        $('prop-size').value = campo.tamano;
        $('prop-size-range').value = Math.min(60, campo.tamano);
        $('prop-color').value = campo.color;
        $('prop-color-hex').textContent = campo.color.toUpperCase();
        $('prop-case').value = campo.mayusculas;
        $('prop-maxwidth').value = campo.anchoMax;
        document.querySelectorAll('#prop-align button').forEach(b =>
            b.classList.toggle('active', b.dataset.align === campo.alineacion));
        this.actualizarInputsPosicion();
    },

    /** Indica si la negrita/cursiva es real o simulada para la fuente elegida */
    actualizarAyudaEstilo() {
        const campo = this.campoSel;
        if (!campo) return;
        const est = Fuentes.resolver(campo);
        const partes = [];
        if (campo.negrita) partes.push(est.fauxBold ? 'negrita simulada' : 'negrita real');
        if (campo.cursiva) partes.push(est.fauxItalic ? 'cursiva simulada' : 'cursiva real');
        let t = partes.length ? 'Esta fuente usa ' + partes.join(' y ') + '.' : '';
        if (est.fauxBold || est.fauxItalic) t += ' Puedes importar el archivo real en 🔤 Fuentes.';
        $('prop-style-hint').textContent = t;
    },

    actualizarInputsPosicion() {
        const campo = this.campoSel;
        if (!campo) return;
        if (document.activeElement !== $('prop-x')) $('prop-x').value = campo.x.toFixed(1);
        if (document.activeElement !== $('prop-y')) $('prop-y').value = campo.y.toFixed(1);
    },

    conectarPropiedades() {
        const cambiar = (fn) => {
            const campo = this.campoSel;
            if (!campo) return;
            fn(campo);
            this.renderCampos();
            this.guardarEstado();
        };

        $('prop-font').addEventListener('change', e => {
            cambiar(c => c.fuente = e.target.value);
            this.actualizarAyudaEstilo();
        });

        // Negrita, cursiva, subrayado y sombra
        document.querySelectorAll('#prop-styles button').forEach(b => b.addEventListener('click', () => {
            cambiar(c => c[b.dataset.style] = !c[b.dataset.style]);
            b.classList.toggle('active', !!this.campoSel[b.dataset.style]);
            $('shadow-options').style.display = this.campoSel.sombra ? 'flex' : 'none';
            this.actualizarAyudaEstilo();
        }));
        $('prop-shadow-color').addEventListener('input', e => cambiar(c => c.sombraColor = e.target.value));
        $('prop-shadow-dist').addEventListener('input', e => {
            cambiar(c => c.sombraDist = parseFloat(e.target.value));
            $('prop-shadow-dist-val').textContent = e.target.value + '%';
        });
        $('prop-shadow-op').addEventListener('input', e => {
            cambiar(c => c.sombraOpacidad = parseFloat(e.target.value));
            $('prop-shadow-op-val').textContent = e.target.value + '%';
        });
        $('prop-spacing').addEventListener('input', e => {
            cambiar(c => c.espaciado = parseFloat(e.target.value));
            $('prop-spacing-val').textContent = e.target.value > 0 ? e.target.value + '%' : 'normal';
        });

        const ponerTamano = (v) => {
            v = parseFloat(v);
            if (isNaN(v)) return;
            v = Math.min(150, Math.max(2, v));
            cambiar(c => c.tamano = v);
            $('prop-size-range').value = Math.min(60, v);
        };
        $('prop-size').addEventListener('input', e => ponerTamano(e.target.value));
        $('prop-size').addEventListener('blur', e => { e.target.value = this.campoSel?.tamano ?? ''; });
        $('prop-size-range').addEventListener('input', e => {
            $('prop-size').value = e.target.value;
            ponerTamano(e.target.value);
        });

        $('prop-color').addEventListener('input', e => {
            $('prop-color-hex').textContent = e.target.value.toUpperCase();
            cambiar(c => c.color = e.target.value);
        });
        $('prop-case').addEventListener('change', e => cambiar(c => c.mayusculas = e.target.value));
        $('prop-maxwidth').addEventListener('input', e => cambiar(c => {
            const v = Math.min(1000, Math.max(0, parseFloat(e.target.value) || 0));
            c.anchoMax = c.tipo === 'parrafo' ? Math.max(20, v || 20) : v; // un párrafo siempre necesita ancho
            if ($('pc-ancho')) $('pc-ancho').value = c.anchoMax;
        }));

        document.querySelectorAll('#prop-align button').forEach(b => b.addEventListener('click', () => {
            cambiar(c => c.alineacion = b.dataset.align);
            document.querySelectorAll('#prop-align button').forEach(x => x.classList.toggle('active', x === b));
        }));

        const ponerPos = (eje, v) => {
            v = parseFloat(v);
            if (isNaN(v)) return;
            cambiar(c => c[eje] = Math.min(100, Math.max(0, v)));
        };
        $('prop-x').addEventListener('input', e => ponerPos('x', e.target.value));
        $('prop-y').addEventListener('input', e => ponerPos('y', e.target.value));
        $('center-h-btn').addEventListener('click', () => {
            cambiar(c => { c.x = 50; c.alineacion = 'center'; });
            this.renderPropiedades();
        });

        $('add-column-field-btn').addEventListener('click', () => {
            if (!Alumnos.columnasExtra.length) {
                return UI.toast('Tu CSV no tiene columnas adicionales. Agrega columnas como "Horas" o "Instructor" y vuelve a cargarlo.', 'info', { duracion: 7000 });
            }
            const usadas = this.campos.filter(c => c.tipo === 'columna').map(c => c.columna);
            const col = Alumnos.columnasExtra.find(c => !usadas.includes(c)) || Alumnos.columnasExtra[0];
            const id = 'columna_' + Date.now();
            this.campos.push(crearCampo('columna', { id, etiqueta: 'Columna: ' + col, columna: col, tamano: 5, y: 72 }));
            this.seleccionado = id;
            this.refrescarTodo();
        });

        // Botón "Párrafo con datos" (se crea junto a "Texto libre" en todas las interfaces)
        const btnTexto = $('add-text-field-btn');
        if (btnTexto && !$('add-paragraph-field-btn')) {
            const b = btnTexto.cloneNode(false);
            b.id = 'add-paragraph-field-btn';
            b.title = 'Un texto de varios renglones con datos de la lista, por ejemplo: por su participación en el {Taller}';
            b.innerHTML = '<svg class="ico"><use href="#i-file-text" /></svg>Párrafo con datos';
            btnTexto.insertAdjacentElement('afterend', b);
        }
        $('add-paragraph-field-btn').addEventListener('click', () => this.agregarParrafo());
        if (!$('add-qr-field-btn')) {
            const b = btnTexto.cloneNode(false);
            b.id = 'add-qr-field-btn';
            b.title = 'Código QR para verificar que la constancia es auténtica';
            b.innerHTML = '<svg class="ico"><use href="#i-scan-search" /></svg>Código QR';
            $('add-paragraph-field-btn').insertAdjacentElement('afterend', b);
        }
        if (!$('add-image-field-btn')) {
            const b = btnTexto.cloneNode(false);
            b.id = 'add-image-field-btn';
            b.title = 'Firma escaneada, sello o logo encima del fondo (PNG o JPG)';
            b.innerHTML = '<svg class="ico"><use href="#i-image" /></svg>Imagen (firma/sello)';
            $('add-qr-field-btn').insertAdjacentElement('afterend', b);
        }
        $('add-image-field-btn').addEventListener('click', () => this.elegirImagenCampo(null));
        $('add-qr-field-btn').addEventListener('click', () => {
            let c = this.campo('qr');
            if (!c) {
                c = crearCampo('qr', { id: 'qr', etiqueta: 'Código QR', x: 90, y: 80, lado: 22, alineacion: 'center' });
                this.campos.push(c);
            }
            c.activo = true;
            this.seleccionado = 'qr';
            this.refrescarTodo();
            UI.toast('Código QR agregado. Arrástralo a una esquina libre de la hoja.', 'info');
        });

        $('add-text-field-btn').addEventListener('click', () => {
            const n = this.campos.filter(c => c.tipo === 'texto').length + 1;
            const id = 'texto_' + Date.now();
            this.campos.push(crearCampo('texto', {
                id, etiqueta: `Texto libre ${n}`, valor: 'Escribe aquí el texto', tamano: 5, y: 70
            }));
            this.seleccionado = id;
            this.refrescarTodo();
            setTimeout(() => $('pc-valor')?.select(), 50);
        });
    },

    /** Elige un PNG/JPG para un campo de imagen (nuevo si id es null) */
    elegirImagenCampo(id) {
        const inp = document.createElement('input');
        inp.type = 'file';
        inp.accept = 'image/png,image/jpeg';
        inp.onchange = async () => {
            const file = inp.files[0];
            if (!file) return;
            if (!/^image\/(png|jpe?g)$/i.test(file.type)) return UI.toast('Usa una imagen PNG o JPG.', 'warning');
            try {
                const { dataUrl, proporcion } = await reducirImagen(file, 1400);
                let campo = id ? this.campo(id) : null;
                if (!campo) {
                    const n = this.campos.filter(c => c.tipo === 'imagen').length + 1;
                    campo = crearCampo('imagen', {
                        id: 'imagen_' + Date.now(), etiqueta: n > 1 ? `Imagen ${n}` : 'Firma / sello', x: 30, y: 80, anchoImg: 45
                    });
                    this.campos.push(campo);
                }
                campo.imagen = dataUrl;
                campo.proporcion = proporcion;
                campo.activo = true;
                this.seleccionado = campo.id;
                this.refrescarTodo();
                UI.toast('Imagen agregada. Arrástrala a su lugar y ajusta el ancho.', 'success');
            } catch (e) {
                UI.toast('No se pudo abrir la imagen:\n' + e.message, 'error');
            }
        };
        inp.click();
    },

    agregarParrafo(valor) {
        const n = this.campos.filter(c => c.tipo === 'parrafo').length + 1;
        const id = 'parrafo_' + Date.now();
        const hayHoras = Alumnos.columnasExtra.find(c => normalizarEncabezado(c) === 'horas');
        const hoja = this.dimensionesPagina();
        this.campos.push(crearCampo('parrafo', {
            id, etiqueta: n > 1 ? `Párrafo ${n}` : 'Párrafo con datos',
            valor: valor || ('por su participación en el {Evento}' + (hayHoras ? ', con duración de {Horas} horas' : '') +
                ', realizado el {Fecha}.'),
            tamano: 4, y: 62, anchoMax: Math.round(hoja.w * 0.72), resaltar: true, interlineado: 1.35
        }));
        this.seleccionado = id;
        this.refrescarTodo();
        setTimeout(() => $('pc-parrafo')?.focus(), 50);
        return id;
    },

    // ------------------------------------------------------------------
    // VENTANA DE FUENTES (ver todas e importar .ttf)
    // ------------------------------------------------------------------
    conectarFuentes() {
        $('manage-fonts-btn').addEventListener('click', () => this.abrirFuentes());
        $('font-file').addEventListener('change', e => {
            const f = e.target.files[0];
            $('font-file-name').textContent = f ? f.name : 'Ningún archivo';
            if (f && !$('font-new-name').value) $('font-new-name').value = f.name.replace(/\.ttf$/i, '');
        });
        const mostrarNombreNuevo = () => {
            $('font-new-name-wrap').style.display = $('font-import-family').value === '__nueva__' ? 'flex' : 'none';
        };
        $('font-import-family').addEventListener('change', mostrarNombreNuevo);

        $('font-import-btn').addEventListener('click', async () => {
            const file = $('font-file').files[0];
            if (!file) return UI.toast('Primero elige un archivo .ttf.', 'warning');
            try {
                const r = await Fuentes.importar(file, $('font-import-family').value,
                    $('font-new-name').value, $('font-import-style').value);
                $('font-file').value = '';
                $('font-file-name').textContent = 'Ningún archivo';
                $('font-new-name').value = '';
                Fuentes.llenarSelect($('prop-font'));
                this.abrirFuentes(r.id);
                this.renderPropiedades();
                this.renderCampos();
                UI.toast(`Fuente importada en "${r.nombre}".`, 'success');
            } catch (err) {
                console.error(err);
                UI.toast('No se pudo importar:\n' + err.message, 'error');
            }
        });

        $('font-list').addEventListener('click', async e => {
            const usar = e.target.closest('[data-usar]');
            if (usar && this.campoSel) {
                this.campoSel.fuente = usar.dataset.usar;
                this.renderPropiedades();
                this.renderCampos();
                this.guardarEstado();
                UI.cerrarModal('fonts-modal');
                return;
            }
            const quitar = e.target.closest('[data-quitar]');
            if (quitar) {
                const [id, estilo] = quitar.dataset.quitar.split('|');
                if (!await UI.confirmar('Quitar fuente importada', '¿Quitar este estilo importado?', 'Quitar', true)) return;
                Fuentes.eliminarImportada(id, estilo);
                await Fuentes.init();
                // Los textos que usaban una familia que ya no existe pasan a Calibri Light
                this.campos.forEach(c => { if (!Fuentes.familia(c.fuente)) c.fuente = 'CalibriLight'; });
                this.guardarEstado();
                Fuentes.llenarSelect($('prop-font'));
                this.renderPropiedades();
                this.renderCampos();
                this.abrirFuentes();
            }
        });
        mostrarNombreNuevo();
    },

    abrirFuentes(idSeleccion) {
        const nombresEstilo = { bold: 'N', italic: 'K', bolditalic: 'NK' };
        const grupos = [['originales', 'Originales'], ['sans', 'Sans serif'], ['serif', 'Serif (elegantes)'],
            ['script', 'Caligráficas'], ['importadas', 'Importadas por ti']];
        const actual = this.campoSel?.fuente;
        let html = '';
        for (const [cat, titulo] of grupos) {
            const lista = Fuentes.familias.filter(f => f.categoria === cat);
            if (!lista.length) continue;
            html += `<div class="font-group-title">${titulo}</div>`;
            for (const f of lista) {
                const badges = ['bold', 'italic', 'bolditalic'].map(st =>
                    `<span class="badge ${f.archivos[st] ? 'on' : ''}" title="${f.archivos[st] ? 'Estilo real' : 'Se simula'}">${nombresEstilo[st]}</span>`).join(' ');
                const importados = Object.entries(f.archivos).filter(([, a]) => a.origen === 'usuario')
                    .map(([st]) => `<button class="icon-btn" data-quitar="${f.id}|${st}" title="Quitar estilo importado (${st})">🗑️${nombresEstilo[st] || ''}</button>`).join('');
                html += `<div class="font-item">
                    <span class="sample" style="font-family:${escapeHtml(Fuentes.css(f.id))}">${escapeHtml(f.nombre)} — María Hernández</span>
                    ${badges} ${importados}
                    <button class="btn small ${f.id === actual ? '' : 'ghost'}" data-usar="${f.id}">${f.id === actual ? '✓ En uso' : 'Usar'}</button>
                </div>`;
            }
        }
        $('font-list').innerHTML = html;

        Fuentes.llenarSelect($('font-import-family'), true);
        $('font-import-family').value = idSeleccion || actual || 'Calibri';
        $('font-import-family').dispatchEvent(new Event('change'));
        const puede = Escritorio.tieneArchivos && !!Fuentes.carpetaUsuario;
        $('font-import-box').style.opacity = puede ? '1' : '0.5';
        $('font-import-btn').disabled = !puede;
        UI.abrirModal('fonts-modal');
    },

    // ------------------------------------------------------------------
    // CARGA DE ALUMNOS (CSV) Y RESUMEN
    // ------------------------------------------------------------------
    conectarAlumnos() {
        $('csv-upload').addEventListener('change', async (e) => {
            const archivos = Array.from(e.target.files || []);
            e.target.value = ''; // permite volver a cargar el mismo archivo
            for (const file of archivos) {
                try {
                    const r = await Alumnos.cargarCSV(file);
                    this.avisoCarga(file.name, r);
                } catch (err) {
                    console.error(err);
                    UI.toast(`No se pudo leer ${file.name}:\n${err.message}`, 'error');
                }
            }
            this.trasCargar(archivos.length > 0);
        });

        this.conectarPegado();
        this.conectarAyudaFormato();

        $('clear-list-btn').addEventListener('click', async () => {
            if (!Alumnos.lista.length) return;
            const ok = await UI.confirmar('Limpiar lista',
                `¿Eliminar los ${Alumnos.lista.length} alumnos de la lista?`, 'Eliminar todo', true);
            if (!ok) return;
            if (Proyectos.activo && !(await Proyectos.preguntarSiNuevoEvento('vaciar la lista'))) return;
            const respaldo = { lista: Alumnos.lista, archivos: Alumnos.archivos, columnas: Alumnos.columnasExtra };
            Alumnos.lista = [];
            Alumnos.archivos = [];
            Alumnos.columnasExtra = [];
            this.indiceVista = -1;
            this.actualizarResumen();
            this.actualizarVista();
            UI.toast('Lista vacía. Puedes cargar nuevos CSV.', 'info', {
                accion: {
                    texto: 'Deshacer', fn: () => {
                        if (Alumnos.lista.length) return UI.toast('Ya cargaste otra lista; no se puede deshacer.', 'warning');
                        Alumnos.lista = respaldo.lista;
                        Alumnos.archivos = respaldo.archivos;
                        Alumnos.columnasExtra = respaldo.columnas;
                        this.indiceVista = 0;
                        this.actualizarResumen();
                        this.actualizarVista();
                    }
                }
            });
        });

        $('event-filter').addEventListener('change', e => {
            this.filtroEvento = e.target.value;
            this.actualizarResumen();
            this.actualizarVista();
        });
        $('csv-example-btn').addEventListener('click', () => Alumnos.descargarEjemplo());

        $('open-list-btn').addEventListener('click', () => {
            $('list-search').value = '';
            this.renderTablaAlumnos();
            UI.abrirModal('list-modal');
            setTimeout(() => $('list-search').focus(), 50);
        });
    },

    /** Notificación con lo que se leyó de un archivo o de la tabla pegada */
    avisoCarga(origen, r) {
        let msg = `${origen}\n${r.agregados} alumnos nuevos`;
        if (r.duplicados) msg += `, ${r.duplicados} duplicados omitidos`;
        if (r.vacios) msg += `, ${r.vacios} filas sin nombre`;
        msg += `.\nColumna de nombre: "${r.colNombre}"`;
        msg += r.colCodigo ? ` · Código: "${r.colCodigo}"` : ' · (sin columna de código)';
        msg += r.colCorreo ? ` · Correo: "${r.colCorreo}"` : ' · (sin columna de correo)';
        const otras = [['Evento', r.colEvento], ['Folio', r.colFolio], ['Fecha', r.colFecha]]
            .filter(([, c]) => c).map(([n, c]) => `${n}: "${c}"`);
        if (otras.length) msg += '\n' + otras.join(' · ');
        if (r.extras.length) msg += `\nOtras columnas disponibles: ${r.extras.join(', ')}`;
        if (r.hojas > 1) msg += `\n${r.hojas} hojas leídas` + (r.eventoPorHoja ? ' (el nombre de cada hoja se usó como evento)' : '');
        UI.toast(msg, r.agregados ? 'success' : 'warning', { duracion: 7000 });
    },

    /** Después de cargar datos: refresca la pantalla y (si está activado) abre la revisión */
    trasCargar(huboCarga) {
        if (Alumnos.lista.length && this.indiceVista < 0) this.indiceVista = 0;
        this.manual = null;
        this.actualizarResumen();
        this.actualizarVista();
        this.renderPropiedades();
        if (huboCarga && Alumnos.lista.length && $('review-after-load').checked) this.abrirRevision();
    },

    // ------------------------------------------------------------------
    // FORMATO ESPERADO DE LA TABLA (ayuda "?") — según los textos del diseño
    // ------------------------------------------------------------------
    /** Columnas que necesita el diseño actual: [{ col, rol, nota, ej: [v1, v2] }] */
    formatoEsperado() {
        const activo = (id) => { const c = this.campo(id); return c && c.activo ? c : null; };
        const cols = [];
        const cCod = activo('codigo');
        if (cCod) cols.push({ col: 'Código', rol: 'recomendada', nota: 'Aparece en la constancia', ej: ['219512345', '219512346'] });
        cols.push({ col: 'Nombre', rol: 'obligatoria', nota: 'Nombre completo de la persona', ej: ['Juan Pérez García', 'María López Hernández'] });
        cols.push({ col: 'Correo', rol: 'opcional', nota: 'Solo si vas a enviarlas por correo', ej: ['juan.perez@alumnos.udg.mx', 'maria.lopez@alumnos.udg.mx'] });
        const cEv = this.campos.find(c => c.tipo === 'evento' && c.activo);
        if (cEv) cols.push({
            col: 'Evento', rol: 'opcional',
            nota: cEv.valor ? `Si no la pones se usa «${cEv.valor}»` : 'Nombre del taller o curso',
            ej: ['Taller de Git', 'Curso de Python']
        });
        const cFolio = activo('folio');
        if (cFolio) cols.push({
            col: 'Folio', rol: 'opcional',
            nota: `Si no la pones se numera solo (${this.folioDe(null, 0, cFolio)}, …)`, ej: ['GIT-001', 'PY-001']
        });
        const cFecha = this.campos.find(c => c.tipo === 'fecha' && c.activo);
        if (cFecha) cols.push({
            col: 'Fecha', rol: 'opcional',
            nota: cFecha.valor ? `Si no la pones se usa «${cFecha.valor}»` : 'Fecha del evento',
            ej: ['5 de octubre de 2026', '12 de octubre de 2026']
        });
        for (const c of this.campos) {
            if (c.tipo !== 'columna' || !c.activo || !c.columna) continue;
            if (cols.some(x => normalizarEncabezado(x.col) === normalizarEncabezado(c.columna))) continue;
            cols.push({ col: c.columna, rol: 'obligatoria', nota: 'Aparece en la constancia', ej: ['…', '…'] });
        }
        return cols;
    },

    /** HTML del recuadro de ayuda con la tabla de ejemplo */
    htmlFormato() {
        const cols = this.formatoEsperado();
        const etiqueta = { obligatoria: 'Obligatoria', recomendada: 'Recomendada', opcional: 'Opcional' };
        const th = cols.map(c => `<th class="fa-${c.rol}">${escapeHtml(c.col)}</th>`).join('');
        const fila = (k) => '<tr>' + cols.map(c => `<td>${escapeHtml(c.ej[k])}</td>`).join('') + '</tr>';
        const notas = cols.map(c => `<li><b>${escapeHtml(c.col)}</b> <span class="fa-tag fa-${c.rol}">${etiqueta[c.rol]}</span> ${escapeHtml(c.nota)}</li>`).join('');
        return `<div class="fa-titulo">Así debe verse tu tabla</div>
            <div class="fa-sub">Según los textos que tiene tu constancia ahora. La <b>primera fila</b> lleva los títulos de las columnas; cada persona va en su propia fila.</div>
            <div class="fa-tabla"><table><thead><tr>${th}</tr></thead><tbody>${fila(0)}${fila(1)}</tbody></table></div>
            <ul class="fa-notas">${notas}</ul>
            <div class="fa-sub">El orden de las columnas no importa. También se reconocen títulos como <i>Matrícula</i>, <i>Email</i>, <i>Taller</i> o <i>Apellido paterno</i>. Si hay columnas de más (Carrera, Grupo…) se guardan y puedes usarlas.</div>
            <div class="fa-acciones">
                <button type="button" class="fa-btn" data-fa="plantilla">⬇️ Descargar plantilla con estas columnas</button>
                <button type="button" class="fa-btn" data-fa="pegar">📋 Pegar mis datos</button>
            </div>`;
    },

    /** Plantilla CSV con exactamente las columnas que pide el diseño */
    descargarPlantilla() {
        const cols = this.formatoEsperado();
        const filas = [cols.map(c => c.col), cols.map(c => c.ej[0]), cols.map(c => c.ej[1])];
        descargarBlob(new Blob(['﻿' + Papa.unparse(filas)], { type: 'text/csv;charset=utf-8' }), 'plantilla_alumnos.csv');
    },

    /** Signos "?" (clase .ayuda-formato): al pasar el mouse muestran cómo debe ser la tabla */
    conectarAyudaFormato() {
        let pop = $('formato-pop');
        if (!pop) {
            pop = document.createElement('div');
            pop.id = 'formato-pop';
            pop.setAttribute('role', 'tooltip');
            document.body.appendChild(pop);
        }
        let timer = null, ancla = null;
        const ocultar = () => { pop.classList.remove('open'); ancla = null; };
        const ocultarLuego = () => { clearTimeout(timer); timer = setTimeout(ocultar, 260); };
        const mostrar = (el) => {
            clearTimeout(timer);
            if (ancla === el && pop.classList.contains('open')) return;
            ancla = el;
            pop.innerHTML = this.htmlFormato();
            pop.classList.add('open');
            // Posición: debajo del "?" sin salirse de la ventana
            // (abajo; si no cabe, arriba; si tampoco, a un lado) sin tapar el propio "?"
            const r = el.getBoundingClientRect();
            const w = pop.offsetWidth, h = pop.offsetHeight;
            const W = window.innerWidth, H = window.innerHeight;
            let left = Math.min(Math.max(8, r.left + r.width / 2 - 40), W - w - 8);
            let top;
            if (r.bottom + 8 + h <= H - 8) top = r.bottom + 8;
            else if (r.top - 8 - h >= 8) top = r.top - h - 8;
            else {
                top = Math.min(Math.max(8, r.top - h / 2), H - h - 8);
                left = (r.right + 8 + w <= W - 8) ? r.right + 8 : Math.max(8, r.left - w - 8);
            }
            pop.style.left = left + 'px';
            pop.style.top = top + 'px';
        };
        document.querySelectorAll('.ayuda-formato').forEach(el => {
            el.addEventListener('mouseenter', () => mostrar(el));
            el.addEventListener('mouseleave', ocultarLuego);
            el.addEventListener('focus', () => mostrar(el));
            el.addEventListener('blur', ocultarLuego);
            // En pantallas táctiles (sin mouse) se abre/cierra tocando
            el.addEventListener('click', (e) => {
                e.preventDefault(); e.stopPropagation();
                if (ancla === el && pop.classList.contains('open')) ocultar(); else mostrar(el);
            });
        });
        pop.addEventListener('mouseenter', () => clearTimeout(timer));
        pop.addEventListener('mouseleave', ocultarLuego);
        pop.addEventListener('click', (e) => {
            const b = e.target.closest('[data-fa]');
            if (!b) return;
            ocultar();
            if (b.dataset.fa === 'plantilla') this.descargarPlantilla();
            else this.abrirPegado();
        });
        document.addEventListener('keydown', e => { if (e.key === 'Escape') ocultar(); });
        document.addEventListener('mousedown', e => {
            if (!pop.contains(e.target) && !e.target.closest('.ayuda-formato')) ocultar();
        });
    },

    // ------------------------------------------------------------------
    // PEGAR DATOS EN UNA TABLA (sin subir archivo)
    // ------------------------------------------------------------------
    pegado: { cols: [], filas: [] },

    /** Crea la ventana de "Pegar datos" (una sola vez, la comparten todas las interfaces) */
    crearModalPegado() {
        if ($('paste-modal')) return;
        if (!$('pg-estilos')) {
            const st = document.createElement('style');
            st.id = 'pg-estilos';
            st.textContent = `
            .ayuda-formato { display:inline-flex; align-items:center; justify-content:center; width:18px; height:18px;
                border-radius:50%; background:var(--oro, #b8860b); color:#fff; font:700 12px/1 system-ui, sans-serif;
                cursor:help; user-select:none; flex:0 0 auto; vertical-align:middle; margin-left:4px; }
            .ayuda-formato.esquina { margin:-3px -4px 0 -18px; position:relative; z-index:2; align-self:flex-start;
                box-shadow:0 0 0 2px #fff; }
            .ayuda-formato:hover, .ayuda-formato:focus { outline:2px solid var(--oro-medio, #e5c77a); outline-offset:1px; }
            #formato-pop { position:fixed; z-index:12500; display:none; width:min(560px, calc(100vw - 16px));
                background:#fff; color:#222; border:1px solid var(--borde, #ddd); border-radius:8px;
                box-shadow:0 14px 40px rgba(0,0,0,.22); padding:12px 14px; font-size:12.5px; line-height:1.4; }
            #formato-pop.open { display:block; }
            #formato-pop .fa-titulo { font-weight:700; font-size:14px; margin-bottom:2px; }
            #formato-pop .fa-sub { color:#555; margin:4px 0; }
            #formato-pop .fa-tabla { overflow-x:auto; margin:8px 0; border:1px solid #cfd5dc; border-radius:4px; }
            #formato-pop table { border-collapse:collapse; width:100%; font-size:12px; }
            #formato-pop th, #formato-pop td { border:1px solid #dfe3e8; padding:4px 7px; white-space:nowrap; text-align:left; }
            #formato-pop th { background:#e9f2ea; font-weight:700; }
            #formato-pop th.fa-obligatoria { background:#d8eedd; }
            #formato-pop th.fa-opcional { background:#f3f4f6; color:#444; }
            #formato-pop .fa-notas { margin:6px 0 6px 18px; padding:0; }
            #formato-pop .fa-notas li { margin:2px 0; }
            #formato-pop .fa-tag { display:inline-block; font-size:10.5px; padding:0 6px; border-radius:999px; margin-right:4px; }
            #formato-pop .fa-tag.fa-obligatoria { background:#d8eedd; color:#1d6b33; }
            #formato-pop .fa-tag.fa-recomendada { background:#fff1cf; color:#8a5a00; }
            #formato-pop .fa-tag.fa-opcional { background:#eef0f3; color:#555; }
            #formato-pop .fa-acciones { display:flex; gap:6px; flex-wrap:wrap; margin-top:8px; }
            #formato-pop .fa-btn { border:1px solid var(--borde, #ccc); background:#fafafa; border-radius:5px;
                padding:4px 9px; cursor:pointer; font-size:12px; }
            #formato-pop .fa-btn:hover { background:var(--oro-claro, #fff6dd); border-color:var(--oro-medio, #e5c77a); }
            #paste-modal .pg-instr { font-size:12.5px; line-height:1.5; }
            #paste-modal kbd { border:1px solid #ccc; border-bottom-width:2px; border-radius:3px; padding:0 4px; font-size:11px; background:#fff; }
            #paste-modal .pg-wrap { max-height:52vh; }
            #paste-modal table.pg-tabla th { vertical-align:top; min-width:140px; }
            #paste-modal table.pg-tabla th:first-child, #paste-modal table.pg-tabla th:last-child { min-width:0; }
            #paste-modal .pg-th { display:flex; align-items:center; gap:2px; }
            #paste-modal .pg-h { font-weight:600; background:#fff; border-color:var(--borde, #ddd); }
            #paste-modal .pg-tipo { display:block; font-size:10.5px; font-weight:500; color:#666; padding:2px 6px 0; }
            #paste-modal .pg-tipo.ok { color:#1d6b33; }
            #paste-modal .pg-tipo.mal { color:#b3261e; }
            #paste-modal td .cell-input { min-width:120px; }
            #paste-modal .modal-header h3 { display:flex; align-items:center; gap:4px; }
            `;
            document.head.appendChild(st);
        }
        const m = document.createElement('div');
        m.className = 'modal-backdrop';
        m.id = 'paste-modal';
        m.innerHTML = `
        <div class="modal">
            <div class="modal-header">
                <h3>📋 Pegar o escribir datos <span class="ayuda-formato" tabindex="0" title="">?</span></h3>
                <button class="close-x" data-close="paste-modal">✕</button>
            </div>
            <div class="modal-body">
                <div class="info-box pg-instr">
                    <b>1.</b> En Excel o Google Sheets selecciona las celdas (con los títulos) y copia con <kbd>Ctrl</kbd>+<kbd>C</kbd>.
                    &nbsp;<b>2.</b> Haz clic en la primera celda de abajo y pega con <kbd>Ctrl</kbd>+<kbd>V</kbd>.
                    &nbsp;<b>3.</b> Revisa y presiona <b>Agregar a la lista</b>. También puedes escribir directo en las celdas.
                </div>
                <div class="row" style="align-items:center">
                    <button type="button" class="btn small" id="pg-portapapeles" style="flex:0">📋 Pegar del portapapeles</button>
                    <button type="button" class="btn small" id="pg-fila" style="flex:0">＋ Fila</button>
                    <button type="button" class="btn small" id="pg-columna" style="flex:0">＋ Columna</button>
                    <button type="button" class="btn small" id="pg-limpiar" style="flex:0">🧹 Vaciar tabla</button>
                    <span class="hint" id="pg-estado" style="flex:1;text-align:right"></span>
                </div>
                <div class="table-wrap pg-wrap">
                    <table class="students pg-tabla">
                        <thead id="pg-head"></thead>
                        <tbody id="pg-body"></tbody>
                    </table>
                </div>
            </div>
            <div class="modal-footer">
                <button type="button" class="btn" data-close="paste-modal">Cerrar</button>
                <button type="button" class="btn primary" id="pg-agregar">Agregar a la lista</button>
            </div>
        </div>`;
        document.body.appendChild(m);
        m.querySelectorAll('[data-close]').forEach(b => b.addEventListener('click', () => UI.cerrarModal('paste-modal')));
    },

    conectarPegado() {
        this.crearModalPegado();
        document.querySelectorAll('.paste-btn').forEach(b => b.addEventListener('click', () => this.abrirPegado()));

        $('pg-fila').addEventListener('click', () => {
            this.leerTablaPegado();
            this.pegado.filas.push(this.pegado.cols.map(() => ''));
            this.renderPegado();
            this.focoPegado(this.pegado.filas.length - 1, 0);
        });
        $('pg-columna').addEventListener('click', () => {
            this.leerTablaPegado();
            this.pegado.cols.push(this.nombreColumnaLibre('Columna'));
            this.pegado.filas.forEach(f => f.push(''));
            this.renderPegado();
            const ins = $('pg-head').querySelectorAll('input');
            const ult = ins[ins.length - 1];
            if (ult) { ult.focus(); ult.select(); }
        });
        $('pg-limpiar').addEventListener('click', async () => {
            this.leerTablaPegado();
            if (this.filasConDatos().length) {
                const ok = await UI.confirmar('Vaciar tabla', '¿Borrar todo lo que está en la tabla?', 'Vaciar', true);
                if (!ok) return;
            }
            this.reiniciarPegado();
            this.renderPegado();
            this.focoPegado(0, 0);
        });
        $('pg-portapapeles').addEventListener('click', async () => {
            let texto = '';
            try { texto = await navigator.clipboard.readText(); } catch (e) { }
            if (!texto) return UI.toast('No se pudo leer el portapapeles. Haz clic en la primera celda y presiona Ctrl+V.', 'warning');
            this.leerTablaPegado();
            this.pegarTexto(texto, 0, 0, true);
        });
        $('pg-agregar').addEventListener('click', () => this.agregarPegado());

        const body = $('pg-body'), head = $('pg-head');
        // Pegar sobre una celda: el bloque copiado de Excel se reparte en las celdas
        const alPegar = (e) => {
            const inp = e.target.closest('input');
            if (!inp) return;
            const texto = (e.clipboardData || window.clipboardData).getData('text');
            if (!/[\t\n\r]/.test(texto)) return; // un solo valor: pegado normal
            e.preventDefault();
            e.stopPropagation(); // que no lo vuelva a pegar el detector general
            this.leerTablaPegado();
            if (inp.dataset.h !== undefined) this.pegarTexto(texto, -1, +inp.dataset.h, false);
            else this.pegarTexto(texto, +inp.dataset.r, +inp.dataset.c, false);
        };
        body.addEventListener('paste', alPegar);
        head.addEventListener('paste', alPegar);
        // Si el modal está abierto y no hay una celda seleccionada, Ctrl+V pega desde la primera
        document.addEventListener('paste', (e) => {
            if (e.defaultPrevented || !$('paste-modal').classList.contains('open')) return;
            if (e.target.closest && e.target.closest('#paste-modal input, #paste-modal textarea')) return;
            const texto = (e.clipboardData || window.clipboardData).getData('text');
            if (!texto) return;
            e.preventDefault();
            this.leerTablaPegado();
            this.pegarTexto(texto, 0, 0, true);
        });

        // Escribir: actualizar el contador y los avisos sin redibujar
        let tEstado = null;
        body.addEventListener('input', () => { clearTimeout(tEstado); tEstado = setTimeout(() => this.estadoPegado(), 250); });
        head.addEventListener('input', (e) => {
            const i = e.target.dataset.h;
            if (i === undefined) return;
            const tag = head.querySelector(`[data-tipo="${i}"]`);
            if (tag) tag.outerHTML = this.etiquetaTipo(e.target.value, i);
            this.estadoPegado();
        });
        // Enter / ↑ / ↓ para moverse como en Excel
        body.addEventListener('keydown', (e) => {
            const inp = e.target.closest('input');
            if (!inp) return;
            const r = +inp.dataset.r, c = +inp.dataset.c;
            let destino = null;
            if (e.key === 'Enter' || e.key === 'ArrowDown') destino = r + 1;
            else if (e.key === 'ArrowUp') destino = r - 1;
            if (destino === null) return;
            e.preventDefault();
            if (destino < 0) return;
            if (destino >= this.pegado.filas.length) {
                if (e.key !== 'Enter') return;
                this.leerTablaPegado();
                this.pegado.filas.push(this.pegado.cols.map(() => ''));
                this.renderPegado();
            }
            this.focoPegado(destino, c);
        });
        // Borrar fila / columna
        $('paste-modal').addEventListener('click', (e) => {
            const bf = e.target.closest('[data-borrar-fila]');
            const bc = e.target.closest('[data-borrar-col]');
            if (!bf && !bc) return;
            this.leerTablaPegado();
            if (bf) {
                this.pegado.filas.splice(+bf.dataset.borrarFila, 1);
                if (!this.pegado.filas.length) this.pegado.filas.push(this.pegado.cols.map(() => ''));
            } else {
                if (this.pegado.cols.length <= 1) return UI.toast('La tabla necesita al menos una columna.', 'warning');
                const j = +bc.dataset.borrarCol;
                this.pegado.cols.splice(j, 1);
                this.pegado.filas.forEach(f => f.splice(j, 1));
            }
            this.renderPegado();
        });
        document.addEventListener('keydown', e => {
            if (e.key === 'Escape' && $('paste-modal').classList.contains('open')
                && !document.querySelector('#confirm-modal.open')) UI.cerrarModal('paste-modal');
        });
    },

    reiniciarPegado() {
        const cols = this.formatoEsperado().map(c => c.col);
        this.pegado = { cols, filas: Array.from({ length: 8 }, () => cols.map(() => '')) };
    },

    abrirPegado() {
        // Si la tabla está vacía se arma con las columnas que pide el diseño actual
        if (!this.pegado.cols.length || !this.filasConDatos().length) this.reiniciarPegado();
        this.renderPegado();
        UI.abrirModal('paste-modal');
        setTimeout(() => this.focoPegado(0, 0), 50);
    },

    filasConDatos() {
        return this.pegado.filas.filter(f => f.some(v => String(v).trim()));
    },

    nombreColumnaLibre(base) {
        const usados = new Set(this.pegado.cols.map(c => normalizar(c)));
        let n = this.pegado.cols.length + 1, nombre = `${base} ${n}`;
        while (usados.has(normalizar(nombre))) nombre = `${base} ${++n}`;
        return nombre;
    },

    focoPegado(r, c) {
        const inp = $('pg-body').querySelector(`input[data-r="${r}"][data-c="${c}"]`);
        if (inp) { inp.focus(); inp.select(); }
    },

    /** Copia lo que hay escrito en las celdas al estado (antes de redibujar) */
    leerTablaPegado() {
        $('pg-head').querySelectorAll('input[data-h]').forEach(i => { this.pegado.cols[+i.dataset.h] = i.value; });
        $('pg-body').querySelectorAll('input[data-r]').forEach(i => {
            const f = this.pegado.filas[+i.dataset.r];
            if (f) f[+i.dataset.c] = i.value;
        });
    },

    /** Etiqueta bajo el título: qué dato reconoció el programa en esa columna */
    etiquetaTipo(titulo, i) {
        const nombres = {
            nombre: 'Nombre', codigo: 'Código', correo: 'Correo', evento: 'Evento', folio: 'Folio', fecha: 'Fecha',
            apPaterno: 'Apellido paterno', apMaterno: 'Apellido materno', apellidos: 'Apellidos'
        };
        const t = Alumnos.tipoDeEncabezado(titulo);
        if (t) return `<span class="pg-tipo ok" data-tipo="${i}">✓ ${nombres[t]}</span>`;
        if (!String(titulo).trim()) return `<span class="pg-tipo mal" data-tipo="${i}">Sin título</span>`;
        return `<span class="pg-tipo" data-tipo="${i}">Columna extra</span>`;
    },

    renderPegado() {
        const { cols, filas } = this.pegado;
        $('pg-head').innerHTML = '<tr><th style="width:34px"></th>' + cols.map((c, j) =>
            `<th><div class="pg-th"><input class="cell-input pg-h" data-h="${j}" value="${escapeHtml(c)}" placeholder="Título">
                <button type="button" class="icon-btn" data-borrar-col="${j}" title="Quitar columna">✕</button></div>
                ${this.etiquetaTipo(c, j)}</th>`).join('') + '<th style="width:34px"></th></tr>';
        // Muchas filas: se arma el HTML de una vez (mucho más rápido que crear celda por celda)
        $('pg-body').innerHTML = filas.map((f, r) =>
            `<tr><td class="num">${r + 1}</td>` + cols.map((_, c) =>
                `<td><input class="cell-input" data-r="${r}" data-c="${c}" value="${escapeHtml(f[c] ?? '')}"></td>`).join('') +
            `<td><button type="button" class="icon-btn" data-borrar-fila="${r}" title="Quitar fila">🗑</button></td></tr>`).join('');
        this.estadoPegado();
    },

    /** Contador y marcas de celdas con problemas (sin nombre, correo inválido) */
    estadoPegado() {
        this.leerTablaPegado();
        const { cols, filas } = this.pegado;
        const tipos = cols.map(c => Alumnos.tipoDeEncabezado(c));
        let jNombre = tipos.indexOf('nombre');
        const jApellidos = tipos.map((t, j) => (t === 'apPaterno' || t === 'apMaterno' || t === 'apellidos') ? j : -1).filter(j => j >= 0);
        const jCorreo = tipos.indexOf('correo');
        let conDatos = 0, sinNombre = 0, malCorreo = 0;
        const celdas = new Map();
        $('pg-body').querySelectorAll('input[data-r]').forEach(i => celdas.set(i.dataset.r + ',' + i.dataset.c, i));
        filas.forEach((f, r) => {
            const tiene = f.some(v => String(v).trim());
            if (tiene) conDatos++;
            const marca = (c, mal) => {
                const inp = celdas.get(r + ',' + c);
                if (inp) inp.classList.toggle('missing', mal);
            };
            if (jNombre >= 0) {
                const mal = tiene && !String(f[jNombre] || '').trim() && !jApellidos.some(j => String(f[j] || '').trim());
                if (mal) sinNombre++;
                marca(jNombre, mal);
            }
            if (jCorreo >= 0) {
                const v = String(f[jCorreo] || '').trim();
                const mal = !!v && !esCorreoValido(v.replace(/\s/g, ''));
                if (mal) malCorreo++;
                marca(jCorreo, mal);
            }
        });
        let txt = conDatos === 1 ? '1 fila con datos' : `${conDatos} filas con datos`;
        if (jNombre < 0) txt += ' · ⚠️ Falta una columna «Nombre»';
        if (sinNombre) txt += ` · ⚠️ ${sinNombre} sin nombre`;
        if (malCorreo) txt += ` · ⚠️ ${malCorreo} correo(s) con error`;
        $('pg-estado').textContent = txt;
        $('pg-agregar').textContent = conDatos ? `Agregar ${conDatos} a la lista` : 'Agregar a la lista';
        $('pg-agregar').disabled = !conDatos;
    },

    /**
     * Reparte un bloque copiado de Excel/Sheets (separado por tabuladores) en la tabla.
     * @param fila -1 = se pegó sobre los títulos
     * @param inicio true = pegado general (detecta si la primera fila son títulos)
     */
    pegarTexto(texto, fila, col, inicio) {
        texto = String(texto).replace(/^﻿/, '').replace(/\r\n?/g, '\n').replace(/\n+$/, '');
        if (!texto.trim()) return;
        // Excel y Google Sheets copian con tabuladores. Sin tabuladores solo se separa por comas o
        // punto y coma si la primera línea trae títulos (texto CSV); si no, cada línea es un valor
        // (así un nombre como "Pérez, Juan" no se parte en dos).
        let bloque;
        if (texto.includes('\t')) {
            bloque = Papa.parse(texto, { delimiter: '\t', skipEmptyLines: false }).data;
        } else {
            const csv = Papa.parse(texto, { skipEmptyLines: false }).data;
            const titulos = csv[0] || [];
            bloque = (titulos.length > 1 && titulos.some(v => Alumnos.tipoDeEncabezado(v)))
                ? csv : texto.split('\n').map(l => [l]);
        }
        bloque = bloque.map(f => f.map(v => String(v ?? '').trim()));
        bloque = bloque.filter(f => f.some(v => v));
        if (!bloque.length) return;

        const p = this.pegado;
        const pareceTitulos = (f) => f.filter(v => Alumnos.tipoDeEncabezado(v)).length >= 1
            && f.every(v => !/@|\d{4,}/.test(v));
        const vacia = !this.filasConDatos().length;

        let conTitulos = false;
        if (fila === -1) {
            // Se pegó sobre los títulos: la primera fila copiada son los títulos desde esa columna
            conTitulos = true;
            const titulos = bloque.shift();
            titulos.forEach((t, k) => { p.cols[col + k] = t; });
            if (vacia) p.filas = [];
            fila = 0;
        } else if ((inicio || (fila === 0 && col === 0)) && vacia && pareceTitulos(bloque[0])) {
            // La primera fila copiada trae los títulos: la tabla toma esas columnas
            conTitulos = true;
            p.cols = bloque.shift().slice();
            p.filas = [];
            fila = 0;
            col = 0;
        } else if (inicio) {
            // Pegado general sin títulos: se agrega debajo de lo que ya hay
            fila = vacia ? 0 : this.pegado.filas.length;
            if (!vacia) {
                // Quita filas vacías del final para no dejar huecos
                while (p.filas.length && !p.filas[p.filas.length - 1].some(v => String(v).trim())) p.filas.pop();
                fila = p.filas.length;
            }
            col = 0;
        }

        // Títulos repetidos o vacíos se renombran para no perder datos
        const vistos = new Set();
        p.cols = p.cols.map((c, j) => {
            let n = String(c ?? '').trim() || `Columna ${j + 1}`;
            let k = 2;
            while (vistos.has(normalizar(n))) n = `${String(c).trim() || 'Columna'} ${k++}`;
            vistos.add(normalizar(n));
            return n;
        });

        const anchoNecesario = col + bloque.reduce((m, f) => Math.max(m, f.length), 0);
        while (p.cols.length < anchoNecesario) p.cols.push(this.nombreColumnaLibre('Columna'));
        for (const f of p.filas) while (f.length < p.cols.length) f.push('');
        bloque.forEach((valores, k) => {
            const r = fila + k;
            while (p.filas.length <= r) p.filas.push(p.cols.map(() => ''));
            valores.forEach((v, j) => { p.filas[r][col + j] = v; });
        });
        if (!p.filas.length) p.filas.push(p.cols.map(() => ''));
        this.renderPegado();
        if (bloque.length === 1 && bloque[0].length === 1 && !conTitulos) return; // una sola celda: sin aviso
        UI.toast(conTitulos
            ? `Se pegaron ${bloque.length} fila(s) con sus títulos. Revisa las etiquetas ✓ de cada columna.`
            : `Se pegaron ${bloque.length} fila(s). Revisa que el título de cada columna corresponda a sus datos.`,
            'success', { duracion: 4000 });
    },

    agregarPegado() {
        this.leerTablaPegado();
        const { cols } = this.pegado;
        const filas = this.filasConDatos();
        if (!filas.length) return UI.toast('La tabla está vacía. Pega o escribe al menos una persona.', 'warning');
        const tipos = cols.map(c => Alumnos.tipoDeEncabezado(c));
        if (!tipos.includes('nombre')) {
            return UI.toast('Ninguna columna se llama «Nombre». Escribe «Nombre» en el título de la columna que tiene los nombres.', 'warning', { duracion: 6000 });
        }
        let r;
        try {
            r = Alumnos.cargarPegado(cols, filas);
        } catch (err) {
            return UI.toast('No se pudieron agregar los datos:\n' + err.message, 'error');
        }
        this.avisoCarga('Datos pegados', r);
        if (r.agregados) {
            this.reiniciarPegado();
            UI.cerrarModal('paste-modal');
        }
        this.trasCargar(r.agregados > 0);
    },

    // ------------------------------------------------------------------
    // HISTORIAL DE CONSTANCIAS EMITIDAS
    // ------------------------------------------------------------------
    crearModalHistorial() {
        if ($('history-modal')) return;
        const m = document.createElement('div');
        m.className = 'modal-backdrop';
        m.id = 'history-modal';
        m.innerHTML = `
        <div class="modal" style="width:min(1200px,96vw)">
            <div class="modal-header"><h3>🗂 Historial de constancias emitidas <span id="hi-count" class="hint"></span></h3><button class="close-x" data-cerrar>✕</button></div>
            <div class="modal-body">
                <div class="info-box" style="font-size:12.5px">Aquí queda cada constancia que generas, imprimes o envías: su <b>folio</b> y su
                    <b>código de verificación</b>. Sirve para reimprimir o reenviar después y para comprobar si una constancia es auténtica.
                    Se guarda en esta computadora; haz un <b>respaldo</b> de vez en cuando.</div>
                <div class="row" style="align-items:center;flex-wrap:wrap">
                    <input type="search" id="hi-buscar" class="input" placeholder="🔎 Buscar nombre, código, folio o verificación…" style="flex:2;min-width:220px">
                    <select id="hi-evento" class="input" style="flex:1;min-width:160px"></select>
                    <button class="btn small primary" id="hi-verificar" style="flex:0;white-space:nowrap">✔ Verificar un código</button>
                </div>
                <div id="hi-resultado" style="display:none"></div>
                <div class="table-wrap" style="max-height:46vh"><table class="students">
                    <thead><tr><th>Folio</th><th>Nombre</th><th>Código</th><th>Evento</th><th>Verificación</th><th>Emitida</th><th title="Generada · impresa · enviada">G · I · E</th><th>Enviada a</th><th></th></tr></thead>
                    <tbody id="hi-body"></tbody></table></div>
                <div id="hi-mas" class="hint"></div>
            </div>
            <div class="modal-footer" style="justify-content:space-between;flex-wrap:wrap">
                <div style="display:flex;gap:6px;flex-wrap:wrap">
                    <button class="btn small" id="hi-excel">⬇️ Excel</button>
                    <button class="btn small" id="hi-web" title="Archivo registro.json para tu página de verificación (GitHub Pages)">🌐 Exportar para página web</button>
                    <button class="btn small" id="hi-respaldo">💾 Respaldo</button>
                    <button class="btn small" id="hi-restaurar">📥 Restaurar respaldo</button>
                    <input type="file" id="hi-archivo" accept=".json,application/json" style="display:none">
                </div>
                <div style="display:flex;gap:6px"><button class="btn small danger" id="hi-borrar">Borrar historial</button><button class="btn" data-cerrar>Cerrar</button></div>
            </div>
        </div>`;
        document.body.appendChild(m);
        m.querySelectorAll('[data-cerrar]').forEach(b => b.addEventListener('click', () => UI.cerrarModal('history-modal')));
        $('hi-buscar').addEventListener('input', () => this.pintarHistorial());
        $('hi-evento').addEventListener('change', () => this.pintarHistorial());
        $('hi-excel').addEventListener('click', () => Historial.exportarExcel());
        $('hi-respaldo').addEventListener('click', () => {
            if (!Historial.registros.size) return UI.toast('El historial está vacío.', 'warning');
            Historial.exportarRespaldo();
        });
        $('hi-restaurar').addEventListener('click', () => $('hi-archivo').click());
        $('hi-archivo').addEventListener('change', async e => {
            const file = e.target.files[0];
            e.target.value = '';
            if (!file) return;
            try {
                const n = await Historial.importarRespaldo(await file.text());
                UI.toast(`Se agregaron ${n} registro(s) al historial.`, 'success');
                this.abrirHistorial();
                this.renderCampos();
            } catch (err) { UI.toast('No se pudo restaurar:\n' + err.message, 'error'); }
        });
        $('hi-web').addEventListener('click', async () => {
            if (!Historial.registros.size) return UI.toast('El historial está vacío.', 'warning');
            const r = await UI.elegir('Exportar para la página de verificación',
                'Se descargará «registro.json». Súbelo a tu página (carpeta pagina_verificacion, en GitHub Pages) para que el QR muestre si la constancia es válida.\n\n' +
                'Como la página es pública, se recomienda mostrar el nombre abreviado (Paola A. G. S.).',
                [{ v: 'corto', t: 'Nombre abreviado', clase: 'primary' }, { v: 'completo', t: 'Nombre completo' }, { v: null, t: 'Cancelar' }]);
            if (!r) return;
            const n = Historial.exportarVerificacion(r === 'completo', this.campo('qr')?.valor || '');
            UI.toast(`registro.json con ${n} constancias. Súbelo a tu página junto a index.html.`, 'success', { duracion: 8000 });
        });
        $('hi-borrar').addEventListener('click', async () => {
            if (!Historial.registros.size) return;
            const ok = await UI.confirmar('Borrar historial',
                `Se borrarán ${Historial.registros.size} registros. Ya no podrás verificar esas constancias ni reimprimirlas con el mismo folio desde aquí.\n\nTe recomendamos descargar un respaldo antes.`,
                'Borrar todo', true);
            if (!ok) return;
            await Historial.borrarTodo();
            this.pintarHistorial();
            this.renderCampos();
            UI.toast('Historial borrado.', 'info');
        });
        $('hi-verificar').addEventListener('click', async () => {
            const c = await UI.pedirTexto('Verificar constancia', 'Escribe el código de verificación que aparece en la constancia (por ejemplo ABCD-2345).', '', 'Verificar');
            if (!c) return;
            const r = Historial.buscarCodigo(c);
            const box = $('hi-resultado');
            box.style.display = 'block';
            box.className = 'info-box';
            box.style.borderColor = r ? '#2e7d4f' : '#c4314b';
            box.style.background = r ? '#e9f6ee' : '#fdecef';
            box.innerHTML = r
                ? `✅ <b>Constancia válida.</b> ${escapeHtml(r.nombre)} · ${escapeHtml(r.evento || 'sin evento')} · Folio <b>${escapeHtml(r.folio)}</b> · emitida el ${new Date(r.primera).toLocaleDateString('es-MX')}`
                : `❌ <b>No se encontró</b> el código «${escapeHtml(c)}» en el historial de esta computadora.`;
            if (r) { $('hi-buscar').value = r.verificacion; this.pintarHistorial(); }
        });
        $('hi-body').addEventListener('click', async e => {
            const b = e.target.closest('[data-hi]');
            if (!b) return;
            const r = Historial.registros.get(b.dataset.clave);
            if (!r) return;
            const alumno = Historial.comoAlumno(r);
            if (b.dataset.hi === 'pdf') {
                if (!this.validarAntesDeGenerar()) return;
                const doc = this.construirPdf(alumno, 0);
                descargarBlob(doc.output('blob'), (this.nombreDeArchivo(alumno, 0) || 'constancia') + '.pdf');
                await Historial.registrar([{ alumno, indice: 0 }], 'generada');
                this.pintarHistorial();
                UI.toast(`PDF de ${r.nombre} (folio ${r.folio}) con el diseño actual.`, 'success');
            } else if (b.dataset.hi === 'mail') {
                const para = await UI.pedirTexto('Reenviar constancia', `¿A qué correo se envía la constancia de ${r.nombre}?`, r.enviadoA || r.correo || '', 'Siguiente');
                if (para) await Envio.enviarUno(alumno, 0, para);
                this.pintarHistorial();
            } else if (b.dataset.hi === 'lista') {
                if (Alumnos.agregar(alumno)) {
                    this.indiceVista = Alumnos.lista.length - 1;
                    this.trasCargar(false);
                    UI.toast(`${r.nombre} se agregó a la lista.`, 'success');
                } else UI.toast('Esa persona ya está en la lista.', 'info');
            } else if (b.dataset.hi === 'borrar') {
                if (!await UI.confirmar('Quitar del historial', `¿Quitar a ${r.nombre} (folio ${r.folio}) del historial?`, 'Quitar', true)) return;
                await Historial.borrar(r.clave);
                this.pintarHistorial();
            }
        });
    },

    abrirHistorial() {
        this.crearModalHistorial();
        const eventos = [...new Set(Historial.lista().map(r => r.evento || ''))].sort();
        $('hi-evento').innerHTML = `<option value="__todos__">Todos los eventos</option>` +
            eventos.map(e => `<option value="${escapeHtml(normalizar(e))}">${escapeHtml(e || '(sin evento)')}</option>`).join('');
        $('hi-resultado').style.display = 'none';
        this.pintarHistorial();
        UI.abrirModal('history-modal');
    },

    pintarHistorial() {
        const q = normalizar($('hi-buscar').value);
        const qc = limpiarCodigo($('hi-buscar').value);
        const ev = $('hi-evento').value;
        const todos = Historial.lista();
        const filtrados = todos.filter(r => (ev === '__todos__' || r.eventoClave === ev) &&
            (!q || normalizar(`${r.nombre} ${r.codigo} ${r.folio} ${r.evento}`).includes(q) || (qc.length >= 4 && limpiarCodigo(r.verificacion).includes(qc))));
        $('hi-count').textContent = `(${todos.length})`;
        const fmt = (t) => t ? new Date(t).toLocaleDateString('es-MX', { day: 'numeric', month: 'short', year: 'numeric' }) : '';
        const max = 400;
        $('hi-body').innerHTML = filtrados.slice(0, max).map(r => `<tr>
            <td><b>${escapeHtml(r.folio)}</b></td><td>${escapeHtml(r.nombre)}</td><td>${escapeHtml(r.codigo)}</td>
            <td>${escapeHtml(r.evento)}</td><td style="font-family:monospace">${escapeHtml(r.verificacion)}</td>
            <td>${fmt(r.primera)}</td><td style="white-space:nowrap">${r.acciones?.generada || 0} · ${r.acciones?.impresa || 0} · ${r.acciones?.enviada || 0}</td>
            <td style="font-size:11.5px">${escapeHtml(r.enviadoA || '')}</td>
            <td class="actions" style="width:120px"><button class="icon-btn" data-hi="pdf" data-clave="${escapeHtml(r.clave)}" title="Volver a generar su PDF (mismo folio)">📄</button>
                <button class="icon-btn" data-hi="mail" data-clave="${escapeHtml(r.clave)}" title="Reenviar por correo">✉️</button>
                <button class="icon-btn" data-hi="lista" data-clave="${escapeHtml(r.clave)}" title="Agregar a la lista actual">＋</button>
                <button class="icon-btn" data-hi="borrar" data-clave="${escapeHtml(r.clave)}" title="Quitar del historial">🗑</button></td></tr>`).join('') ||
            `<tr><td colspan="9" style="text-align:center;color:#777;padding:24px">${todos.length ? 'Nada coincide con la búsqueda.' : 'Aún no hay constancias en el historial. Se agregan solas al generar, imprimir o enviar.'}</td></tr>`;
        $('hi-mas').textContent = filtrados.length > max ? `Se muestran ${max} de ${filtrados.length}. Usa la búsqueda para encontrar a alguien.` : '';
    },

    // ------------------------------------------------------------------
    // CORREGIR NOMBRES (mayúsculas y acentos sugeridos)
    // ------------------------------------------------------------------
    crearModalNombres() {
        if ($('names-modal')) return;
        const m = document.createElement('div');
        m.className = 'modal-backdrop';
        m.id = 'names-modal';
        m.style.zIndex = '9100';
        m.innerHTML = `
        <div class="modal" style="width:min(820px,96vw)">
            <div class="modal-header"><h3>Aa Corregir nombres</h3><button class="close-x" data-cerrar>✕</button></div>
            <div class="modal-body">
                <div class="info-box" style="font-size:12.5px">Los nombres escritos todo en MAYÚSCULAS o todo en minúsculas quedan como
                    <b>«Juan Pérez de la Cruz»</b>. Los acentos <b>solo se ponen si tú los marcas</b>.</div>
                <div id="nm-acentos-wrap" class="panel-box" style="display:none">
                    <h4>Posibles acentos <span class="hint" style="font-weight:400">(revisa: algunas personas escriben su apellido sin acento)</span></h4>
                    <div id="nm-acentos" style="display:flex;flex-wrap:wrap;gap:4px 14px;font-size:12.5px"></div>
                    <div><button type="button" class="btn small" id="nm-todos" style="flex:0">Marcar todos</button></div>
                </div>
                <div class="row" style="align-items:center"><b id="nm-resumen" style="flex:1"></b>
                    <label class="toolbar-check" style="flex:0;white-space:nowrap"><input type="checkbox" id="nm-sel" checked> Seleccionar todos</label></div>
                <div class="table-wrap" style="max-height:44vh"><table class="students"><thead><tr><th style="width:34px"></th><th>Como está</th><th>Cómo quedará</th></tr></thead><tbody id="nm-body"></tbody></table></div>
            </div>
            <div class="modal-footer"><button class="btn" data-cerrar>Cancelar</button><button class="btn primary" id="nm-aplicar">Aplicar cambios</button></div>
        </div>`;
        document.body.appendChild(m);
        m.querySelectorAll('[data-cerrar]').forEach(b => b.addEventListener('click', () => UI.cerrarModal('names-modal')));
        $('nm-acentos').addEventListener('change', () => this.pintarNombres());
        $('nm-todos').addEventListener('click', () => {
            $('nm-acentos').querySelectorAll('input').forEach(i => { i.checked = true; });
            this.pintarNombres();
        });
        $('nm-sel').addEventListener('change', e => {
            $('nm-body').querySelectorAll('input').forEach(i => { i.checked = e.target.checked; });
        });
        $('nm-aplicar').addEventListener('click', () => this.aplicarCorreccionNombres());
    },

    abrirCorreccionNombres() {
        if (!Alumnos.lista.length) return UI.toast('La lista está vacía.', 'warning');
        this.crearModalNombres();
        // Sugerencias de acento (únicas)
        const sug = new Map();
        for (const a of Alumnos.lista) {
            for (const s of acentosSugeridos(nombrePropio(a.nombre))) {
                const k = s.de;
                if (!sug.has(k)) sug.set(k, { de: s.de, a: s.a, n: 0 });
                sug.get(k).n++;
            }
        }
        const lista = [...sug.values()].sort((x, y) => y.n - x.n);
        $('nm-body').innerHTML = '';
        $('nm-sel').checked = true;
        $('nm-acentos-wrap').style.display = lista.length ? 'flex' : 'none';
        $('nm-acentos').innerHTML = lista.map(s => `<label class="toolbar-check"><input type="checkbox" data-de="${escapeHtml(s.de)}" data-a="${escapeHtml(s.a)}">
            ${escapeHtml(s.de)} → <b>${escapeHtml(s.a)}</b> <span class="hint">(${s.n})</span></label>`).join('');
        this.pintarNombres();
        UI.abrirModal('names-modal');
    },

    /** Nombre final según las casillas de acentos marcadas */
    nombreCorregido(nombre) {
        let r = nombrePropio(nombre);
        const marcados = [...$('nm-acentos').querySelectorAll('input:checked')].map(i => [i.dataset.de, i.dataset.a]);
        if (marcados.length) {
            const mapa = new Map(marcados);
            r = r.split(' ').map(p => mapa.get(p) || p).join(' ');
        }
        return r;
    },

    pintarNombres() {
        const cambios = [];
        Alumnos.lista.forEach((a, i) => {
            const nuevo = this.nombreCorregido(a.nombre);
            if (nuevo !== a.nombre) cambios.push({ i, antes: a.nombre, despues: nuevo });
        });
        this._cambiosNombres = cambios;
        $('nm-resumen').textContent = cambios.length
            ? `${cambios.length} de ${Alumnos.lista.length} nombres cambiarán`
            : 'Todos los nombres ya están bien escritos.';
        const marcadosAntes = new Set([...$('nm-body').querySelectorAll('input:not(:checked)')].map(x => +x.dataset.i));
        $('nm-body').innerHTML = cambios.map(c => `<tr><td><input type="checkbox" data-i="${c.i}" ${marcadosAntes.has(c.i) ? '' : 'checked'}></td>
            <td style="color:#777">${escapeHtml(c.antes)}</td><td><b>${escapeHtml(c.despues)}</b></td></tr>`).join('');
        $('nm-aplicar').disabled = !cambios.length;
    },

    aplicarCorreccionNombres() {
        const elegidos = new Set([...$('nm-body').querySelectorAll('input:checked')].map(x => +x.dataset.i));
        const respaldo = [];
        for (const c of this._cambiosNombres || []) {
            if (!elegidos.has(c.i)) continue;
            respaldo.push([Alumnos.lista[c.i], Alumnos.lista[c.i].nombre]);
            Alumnos.lista[c.i].nombre = c.despues;
        }
        UI.cerrarModal('names-modal');
        if (!respaldo.length) return;
        this.renderTablaAlumnos();
        this.actualizarVista();
        UI.toast(`Se corrigieron ${respaldo.length} nombres.`, 'success', {
            accion: {
                texto: 'Deshacer', fn: () => {
                    for (const [a, n] of respaldo) if (Alumnos.lista.includes(a)) a.nombre = n;
                    this.renderTablaAlumnos();
                    this.actualizarVista();
                }
            }
        });
    },

    actualizarResumen() {
        this._folioCache = null;
        Proyectos.tocar();
        const total = Alumnos.lista.length;
        const conCodigo = Alumnos.conCodigo;
        $('summary-count').textContent = total;
        $('summary-codes').textContent = `${conCodigo} con código`;
        $('summary-emails').textContent = `${Alumnos.conCorreo} con correo`;
        this.renderFiltroEventos();
        const activos = this.indicesActivos().length;
        $('summary-file').textContent = Alumnos.archivos.length === 0 ? 'Sin archivo'
            : Alumnos.archivos.length === 1 ? Alumnos.archivos[0] : `${Alumnos.archivos.length} archivos`;
        $('summary-file').title = Alumnos.archivos.join('\n');

        const campoCodigo = this.campo('codigo');
        const aviso = $('summary-warning');
        const sinCodigo = total - conCodigo;
        if (total && campoCodigo?.activo && sinCodigo > 0) {
            aviso.style.display = 'block';
            aviso.textContent = `⚠️ ${sinCodigo} alumno(s) sin código`;
        } else {
            aviso.style.display = 'none';
        }

        $('clear-list-btn').disabled = total === 0;
        $('generate-pdf-btn').disabled = activos === 0;
        // En la interfaz tipo cinta el texto va en una etiqueta aparte (para no borrar el icono)
        const etiquetaZip = $('generate-pdf-label');
        const textoZip = activos ? `Generar ${activos}` : 'Generar';
        if (etiquetaZip) etiquetaZip.textContent = textoZip;
        else $('generate-pdf-btn').textContent = '📦 ' + textoZip.replace('(ZIP)', 'constancias (ZIP)');
    },

    // ------------------------------------------------------------------
    // MODAL: TABLA DE ALUMNOS (ver, buscar, editar, eliminar, agregar)
    // ------------------------------------------------------------------
    renderTablaAlumnos() {
        const filtro = normalizar($('list-search').value);
        const tbody = $('list-tbody');
        const filas = [];
        const activos = new Set(this.indicesActivos());
        const soloProblemas = $('list-only-issues').checked;
        const problemas = this.problemasDeLista();
        let conProblemas = 0;
        Alumnos.lista.forEach((a, i) => {
            if (!activos.has(i)) return;
            const prob = problemas.get(i) || [];
            if (prob.length) conProblemas++;
            if (soloProblemas && !prob.length) return;
            if (filtro && ![a.nombre, a.codigo, a.correo, a.evento, a.folio].some(v => normalizar(v).includes(filtro))) return;
            const correoMal = a.correo && !esCorreoValido(a.correo);
            const estado = a.envio
                ? (a.envio.estado === 'ok'
                    ? `<span class="status-ok" title="Enviado ${escapeHtml(new Date(a.envio.fecha).toLocaleString('es-MX'))}">✔</span>`
                    : `<span class="status-err" title="${escapeHtml(a.envio.detalle)}">✖</span>`)
                : '';
            filas.push(`<tr data-i="${i}" class="${i === this.indiceVista ? 'current' : ''}">
                <td class="num">${prob.length ? `<span title="${escapeHtml(prob.join('\n'))}" style="cursor:help">⚠️</span> ` : ''}${i + 1}</td>
                <td class="code"><input class="cell-input ${a.codigo ? '' : 'missing'}" data-prop="codigo"
                    value="${escapeHtml(a.codigo)}" placeholder="sin código"></td>
                <td><input class="cell-input" data-prop="nombre" value="${escapeHtml(a.nombre)}">${(() => {
                    const np = nombrePropio(a.nombre);
                    return np !== a.nombre ? `<button class="fix-sug" data-fix-prop="nombre" data-fix="${escapeHtml(np)}" title="Corregir mayúsculas">→ ${escapeHtml(np)}</button>` : '';
                })()}</td>
                <td class="mail"><input class="cell-input ${correoMal ? 'missing' : ''}" data-prop="correo" type="email"
                    value="${escapeHtml(a.correo || '')}" placeholder="sin correo"
                    title="${correoMal ? 'Correo con formato inválido' : ''}">${(() => {
                        const sug = correoMal ? sugerirCorreo(a.correo) : '';
                        return sug ? `<button class="fix-sug" data-fix-prop="correo" data-fix="${escapeHtml(sug)}" title="Arreglo sugerido">→ ${escapeHtml(sug)}</button>` : '';
                    })()}</td>
                <td class="event"><input class="cell-input" data-prop="evento" value="${escapeHtml(a.evento || '')}" placeholder="—"></td>
                <td class="folio"><input class="cell-input" data-prop="folio" value="${escapeHtml(a.folio || '')}" placeholder="auto"></td>
                <td style="text-align:center;width:34px">${estado}</td>
                <td class="actions">
                    <button class="icon-btn" data-act="enviar" title="Enviar por correo solo a este alumno">✉️</button>
                    <button class="icon-btn" data-act="ver" title="Ver en la vista previa">👁</button>
                    <button class="icon-btn" data-act="borrar" title="Eliminar">🗑️</button>
                </td></tr>`);
        });
        tbody.innerHTML = filas.length ? filas.join('')
            : `<tr><td colspan="8" class="empty-row">${Alumnos.lista.length
                ? 'Ningún alumno coincide con la búsqueda.'
                : 'La lista está vacía. Carga un CSV o agrega alumnos abajo.'}</td></tr>`;
        // Aviso de revisión
        const aviso = $('list-review');
        if (Alumnos.lista.length) {
            aviso.style.display = 'block';
            aviso.innerHTML = conProblemas
                ? `<b>Revisa los datos:</b> ${conProblemas} fila(s) con posibles errores (marcadas con ⚠️; pasa el mouse para ver el motivo). Puedes corregirlas aquí mismo.`
                : '✔ No se detectaron problemas en los datos.';
        } else {
            aviso.style.display = 'none';
        }
        $('list-count').textContent = filas.length !== Alumnos.lista.length
            ? `(${filas.length} de ${Alumnos.lista.length})` : `(${Alumnos.lista.length})`;
    },

    /**
     * Posibles errores de captura en cada registro (para la revisión).
     * @returns Map(indice → [motivos])
     */
    problemasDeLista() {
        const mapa = new Map();
        const pedirCodigo = this.campo('codigo')?.activo;
        // Mismo nombre con distinto código dentro del mismo evento → posible error de captura
        const porNombre = new Map();
        Alumnos.lista.forEach((a, i) => {
            const k = normalizar(a.evento) + '|' + normalizar(a.nombre);
            if (!porNombre.has(k)) porNombre.set(k, []);
            porNombre.get(k).push(i);
        });
        Alumnos.lista.forEach((a, i) => {
            const p = [];
            const nombre = (a.nombre || '').trim();
            if (nombre.split(/\s+/).length < 2) p.push('Nombre incompleto (una sola palabra)');
            if (/[0-9@#$%&*_=+<>{}\[\]|\\/]/.test(nombre)) p.push('El nombre tiene números o símbolos');
            if (pedirCodigo && !a.codigo) p.push('Sin código de estudiante');
            if (a.correo && !esCorreoValido(a.correo)) p.push('Correo con formato inválido');
            else if (a.correo && sugerirCorreo(a.correo)) p.push(`¿El correo está mal escrito? Quizá es ${sugerirCorreo(a.correo)}`);
            const iguales = porNombre.get(normalizar(a.evento) + '|' + normalizar(a.nombre));
            if (iguales.length > 1) p.push(iguales.some(j => j !== i && normalizar(Alumnos.lista[j].codigo) === normalizar(a.codigo))
                ? 'Esta persona aparece repetida en este evento'
                : 'El mismo nombre aparece con otro código en este evento');
            if (p.length) mapa.set(i, p);
        });
        // Nombres casi iguales (acentos, espacios, una letra de diferencia o en otro orden)
        for (const [i, j] of nombresParecidos(Alumnos.lista)) {
            const txt = (k) => `Se parece a «${Alumnos.lista[k].nombre}» (fila ${k + 1}): ¿es la misma persona?`;
            if (!mapa.has(i)) mapa.set(i, []);
            if (!mapa.has(j)) mapa.set(j, []);
            mapa.get(i).push(txt(j));
            mapa.get(j).push(txt(i));
        }
        return mapa;
    },

    abrirRevision() {
        $('list-search').value = '';
        $('list-only-issues').checked = this.problemasDeLista().size > 0;
        this.renderTablaAlumnos();
        UI.abrirModal('list-modal');
    },

    // ------------------------------------------------------------------
    // UNA CONSTANCIA CAPTURADA A MANO (sin CSV)
    // ------------------------------------------------------------------
    conectarManual() {
        const ids = ['nombre', 'codigo', 'folio', 'evento', 'fecha', 'correo'];
        const leer = () => {
            const d = { extra: {} };
            for (const k of ids) d[k] = $('man-' + k).value.replace(/\s+/g, ' ').trim();
            d.correo = d.correo.replace(/\s/g, '').toLowerCase();
            return d;
        };
        const mostrar = () => {
            this.manual = leer();
            this.actualizarVista();
        };
        $('manual-btn').addEventListener('click', () => {
            UI.abrirModal('manual-modal');
            mostrar();
            setTimeout(() => $('man-nombre').focus(), 50);
        });
        ids.forEach(k => $('man-' + k).addEventListener('input', mostrar));
        $('man-ver').addEventListener('click', () => { mostrar(); UI.cerrarModal('manual-modal'); });
        $('man-pdf').addEventListener('click', () => {
            const d = leer();
            if (!d.nombre) { UI.toast('Escribe el nombre.', 'warning'); return $('man-nombre').focus(); }
            if (!this.validarAntesDeGenerar()) return;
            try {
                const doc = this.construirPdf(d, 0);
                descargarBlob(doc.output('blob'), (this.nombreDeArchivo(d, 0) || 'constancia') + '.pdf');
                Historial.registrar([{ alumno: d, indice: 0 }], 'generada');
                UI.toast(`PDF de ${d.nombre} generado.`, 'success');
            } catch (err) {
                console.error(err);
                UI.toast('Error al generar el PDF:\n' + err.message, 'error');
            }
        });
        $('man-enviar').addEventListener('click', () => {
            const d = leer();
            if (!d.nombre) { UI.toast('Escribe el nombre.', 'warning'); return $('man-nombre').focus(); }
            if (!esCorreoValido(d.correo)) {
                UI.toast('Escribe el correo del destinatario.', 'warning');
                return $('man-correo').focus();
            }
            Envio.enviarUno(d, 0, d.correo);
        });

        $('man-agregar').addEventListener('click', () => {
            const d = leer();
            if (!d.nombre) { UI.toast('Escribe el nombre.', 'warning'); return $('man-nombre').focus(); }
            if (d.correo && !esCorreoValido(d.correo)) { UI.toast('El correo no tiene un formato válido.', 'warning'); return; }
            if (!Alumnos.agregar(d)) return UI.toast('Ese alumno ya está en la lista (mismo evento).', 'warning');
            ids.forEach(k => { $('man-' + k).value = ''; });
            this.manual = null;
            this.indiceVista = Alumnos.lista.length - 1;
            this.actualizarResumen();
            this.actualizarVista();
            UI.toast(`${d.nombre} se agregó a la lista.`, 'success');
            $('man-nombre').focus();
        });
        // Preferencia: revisar al cargar CSV
        const lblPend = $('mail-only-pending')?.closest('label');
        if (lblPend && !lblPend.dataset.ok) {
            lblPend.dataset.ok = '1';
            lblPend.title = 'Omite a quien ya recibió su constancia en este mismo correo, también en días anteriores (historial). Así puedes retomar un envío que se interrumpió.';
        }
        try { $('review-after-load').checked = localStorage.getItem('gc_revisar') !== 'no'; } catch (e) { }
        $('review-after-load').addEventListener('change', e => {
            try { localStorage.setItem('gc_revisar', e.target.checked ? 'si' : 'no'); } catch (_) { }
        });
        $('list-only-issues').addEventListener('change', () => this.renderTablaAlumnos());
    },

    conectarModales() {
        // Botones ✕ y clic fuera del modal
        document.querySelectorAll('[data-close]').forEach(b =>
            b.addEventListener('click', () => UI.cerrarModal(b.dataset.close)));
        ['list-modal', 'fonts-modal', 'email-modal', 'manual-modal', 'gen-modal', 'print-modal'].forEach(id => $(id).addEventListener('mousedown', e => {
            if (e.target.id === id) UI.cerrarModal(id);
        }));

        $('list-search').addEventListener('input', () => this.renderTablaAlumnos());

        const tbody = $('list-tbody');
        // Edición en línea
        tbody.addEventListener('change', e => {
            const inp = e.target.closest('.cell-input');
            if (!inp) return;
            const i = parseInt(inp.closest('tr').dataset.i, 10);
            const alumno = Alumnos.lista[i];
            let nuevo = inp.value.replace(/\s+/g, ' ').trim();
            if (inp.dataset.prop === 'correo') {
                nuevo = nuevo.replace(/\s/g, '').toLowerCase();
                inp.value = nuevo;
                if (nuevo && !esCorreoValido(nuevo)) UI.toast('Revisa el correo: el formato no es válido.', 'warning');
                if (nuevo !== (alumno.correo || '')) delete alumno.envio; // correo nuevo → se puede volver a enviar
            }
            if (inp.dataset.prop === 'nombre' && !nuevo) {
                inp.value = alumno.nombre;
                UI.toast('El nombre no puede quedar vacío.', 'warning');
                return;
            }
            const prueba = { ...alumno, [inp.dataset.prop]: nuevo };
            if (['codigo', 'nombre', 'evento'].includes(inp.dataset.prop) && Alumnos.existe(prueba, i) !== -1) {
                // Dos filas de la misma persona en el mismo evento compartirían folio y código de verificación
                inp.value = alumno[inp.dataset.prop] || '';
                UI.toast('Ya hay otra fila con esa misma persona en este evento (mismo ' + inp.dataset.prop + '). Revisa los datos.', 'warning', { duracion: 7000 });
                return;
            }
            alumno[inp.dataset.prop] = nuevo;
            this._folioCache = null;
            if (inp.dataset.prop === 'evento') this.renderFiltroEventos();
            inp.classList.toggle('missing', (inp.dataset.prop === 'codigo' && !nuevo)
                || (inp.dataset.prop === 'correo' && nuevo && !esCorreoValido(nuevo)));
            // Volver a dibujar la tabla para actualizar los avisos de revisión
            this.renderTablaAlumnos();
            this.actualizarResumen();
            if (i === this.indiceVista) this.actualizarVista();
        });
        // Arreglos sugeridos (un clic)
        tbody.addEventListener('click', e => {
            const b = e.target.closest('.fix-sug');
            if (!b) return;
            const inp = b.closest('td').querySelector('.cell-input');
            inp.value = b.dataset.fix;
            inp.dispatchEvent(new Event('change', { bubbles: true }));
        });
        if (!$('fix-sug-css')) {
            const st = document.createElement('style');
            st.id = 'fix-sug-css';
            st.textContent = `.fix-sug{display:block;margin:1px 0 2px 4px;border:0;background:none;color:#8a4b00;font-size:11px;cursor:pointer;text-decoration:underline;text-align:left;padding:0;max-width:100%;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
                .fix-sug:hover{color:#5a3000}
                .bg-swatches{display:flex;gap:5px;flex-wrap:wrap;align-items:center}
                .bg-swatches button{width:22px;height:22px;border-radius:5px;border:2px solid #fff;box-shadow:0 0 0 1px #bbb;cursor:pointer;padding:0}
                .bg-swatches button:hover{box-shadow:0 0 0 2px var(--oro,#b8913f)}`;
            document.head.appendChild(st);
        }
        // Botones de cada fila
        tbody.addEventListener('click', e => {
            const btn = e.target.closest('[data-act]');
            if (!btn) return;
            const i = parseInt(btn.closest('tr').dataset.i, 10);
            if (btn.dataset.act === 'ver') {
                this.manual = null;
                this.indiceVista = i;
                this.actualizarVista();
                UI.cerrarModal('list-modal');
            } else if (btn.dataset.act === 'enviar') {
                Envio.enviarUno(Alumnos.lista[i], i, Alumnos.lista[i].correo).then(() => this.renderTablaAlumnos());
            } else if (btn.dataset.act === 'borrar') {
                const [quitado] = Alumnos.lista.splice(i, 1);
                // Vecinos, para reinsertarlo en su lugar aunque se borren otros
                const siguiente = Alumnos.lista[i] || null;
                const anterior = Alumnos.lista[i - 1] || null;
                if (this.indiceVista >= Alumnos.lista.length) this.indiceVista = Alumnos.lista.length - 1;
                this.renderTablaAlumnos();
                this.actualizarResumen();
                this.actualizarVista();
                UI.toast(`Se eliminó a ${quitado.nombre}`, 'info', {
                    accion: {
                        texto: 'Deshacer', fn: () => {
                            if (Alumnos.lista.includes(quitado)) return;
                            let pos = siguiente ? Alumnos.lista.indexOf(siguiente) : -1;
                            if (pos < 0 && anterior && Alumnos.lista.includes(anterior)) pos = Alumnos.lista.indexOf(anterior) + 1;
                            if (pos < 0) pos = Math.min(i, Alumnos.lista.length);
                            Alumnos.lista.splice(pos, 0, quitado);
                            this.renderTablaAlumnos();
                            this.actualizarResumen();
                            this.actualizarVista();
                        }
                    }
                });
            }
        });

        // Agregar alumno a mano
        const agregar = () => {
            const nombre = $('new-name').value.replace(/\s+/g, ' ').trim();
            const codigo = $('new-code').value.trim();
            const correo = $('new-email').value.replace(/\s/g, '').toLowerCase();
            const evento = $('new-event').value.replace(/\s+/g, ' ').trim()
                || (this.filtroEvento !== '__todos__' ? (Alumnos.eventos().find(e => normalizar(e.nombre) === this.filtroEvento)?.nombre || '') : '');
            if (correo && !esCorreoValido(correo)) {
                UI.toast('El correo no tiene un formato válido.', 'warning');
                $('new-email').focus();
                return;
            }
            if (!nombre) {
                UI.toast('Escribe el nombre del alumno.', 'warning');
                $('new-name').focus();
                return;
            }
            if (!Alumnos.agregar({ codigo, nombre, correo, evento, folio: '', fecha: '', extra: {} })) {
                UI.toast('Ese alumno ya está en la lista.', 'warning');
                return;
            }
            $('new-name').value = '';
            $('new-code').value = '';
            $('new-email').value = '';
            $('new-event').value = '';
            $('list-search').value = '';
            this.renderTablaAlumnos();
            this.actualizarResumen();
            if (this.indiceVista < 0) this.indiceVista = 0;
            this.actualizarVista();
            $('new-code').focus();
            const wrap = document.querySelector('.table-wrap');
            wrap.scrollTop = wrap.scrollHeight;
        };
        $('add-student-btn').addEventListener('click', agregar);
        $('new-name').addEventListener('keydown', e => { if (e.key === 'Enter') $('new-email').focus(); });
        $('new-email').addEventListener('keydown', e => { if (e.key === 'Enter') $('new-event').focus(); });
        $('new-event').addEventListener('keydown', e => { if (e.key === 'Enter') agregar(); });
        $('new-code').addEventListener('keydown', e => { if (e.key === 'Enter') $('new-name').focus(); });

        // Botón "Corregir nombres" junto a A→Z (mayúsculas y acentos)
        if (!$('fix-names-btn')) {
            const b = document.createElement('button');
            b.id = 'fix-names-btn';
            b.className = $('sort-list-btn').className;
            b.style.flex = '0';
            b.title = 'Deja los nombres como "Juan Pérez García" y sugiere acentos';
            b.textContent = 'Aa Corregir nombres';
            $('sort-list-btn').insertAdjacentElement('afterend', b);
        }
        $('fix-names-btn').addEventListener('click', () => this.abrirCorreccionNombres());

        $('sort-list-btn').addEventListener('click', () => {
            const actual = this.alumnoActual();
            Alumnos.lista.sort((a, b) => a.nombre.localeCompare(b.nombre, 'es', { sensitivity: 'base' }));
            if (actual) this.indiceVista = Alumnos.lista.indexOf(actual);
            this.renderTablaAlumnos();
            this.actualizarVista();
            UI.toast('Lista ordenada alfabéticamente. Recuerda que el folio sigue este orden.', 'info');
        });
        $('export-xlsx-btn').addEventListener('click', () => {
            if (!Alumnos.lista.length) return UI.toast('La lista está vacía.', 'warning');
            Alumnos.exportarExcel();
        });
        $('export-csv-btn').addEventListener('click', () => {
            if (!Alumnos.lista.length) return UI.toast('La lista está vacía.', 'warning');
            Alumnos.exportarCSV();
        });

        // Esc cierra el modal de lista
        document.addEventListener('keydown', e => {
            if (e.key === 'Escape') ['list-modal', 'fonts-modal', 'email-modal', 'manual-modal', 'gen-modal', 'print-modal'].forEach(m => UI.cerrarModal(m));
        });
    },

    // ------------------------------------------------------------------
    // NAVEGACIÓN DE LA VISTA PREVIA
    // ------------------------------------------------------------------
    conectarVistaPrevia() {
        $('prev-btn').addEventListener('click', () => this.moverVista(-1));
        $('next-btn').addEventListener('click', () => this.moverVista(1));
        $('longest-btn').addEventListener('click', () => {
            if (!Alumnos.lista.length) return UI.toast('Primero carga la lista de alumnos.', 'info');
            const activos = this.indicesActivos();
            if (!activos.length) return;
            let mejor = activos[0];
            for (const i of activos) {
                if (Alumnos.lista[i].nombre.length > Alumnos.lista[mejor].nombre.length) mejor = i;
            }
            this.manual = null;
            this.indiceVista = mejor;
            this.actualizarVista();
        });
        $('toggle-guides').addEventListener('change', e =>
            $('certificate-canvas').classList.toggle('show-guides', e.target.checked));
        $('certificate-canvas').classList.toggle('show-guides', $('toggle-guides').checked);
    },

    moverVista(delta) {
        this.manual = null;   // Al navegar se sale de la captura manual
        const activos = this.indicesActivos();
        const n = activos.length;
        if (!n) { this.indiceVista = -1; return this.actualizarVista(); }
        // Posiciones: -1 (ejemplo) … n-1 dentro de los registros activos
        let pos = activos.indexOf(this.indiceVista) + delta;
        if (pos < -1) pos = n - 1;
        if (pos > n - 1) pos = -1;
        this.indiceVista = pos < 0 ? -1 : activos[pos];
        this.actualizarVista();
    },

    actualizarVista() {
        const activos = this.indicesActivos();
        const n = activos.length;
        if (this.indiceVista >= 0 && !activos.includes(this.indiceVista)) this.indiceVista = n ? activos[0] : -1;
        const a = this.alumnoActual();
        if (this.manual) {
            $('preview-indicator').textContent = 'Manual';
            $('preview-name').textContent = (a.codigo ? `${a.codigo} · ` : '') + (a.nombre || 'Escribe el nombre…') +
                (a.evento ? ` · ${a.evento}` : '') + '  (captura manual)';
            $('prev-btn').disabled = $('next-btn').disabled = n === 0;
            return this.renderCampos();
        }
        $('preview-indicator').textContent = a ? `${activos.indexOf(this.indiceVista) + 1} / ${n}` : 'Ejemplo';
        $('preview-name').textContent = a
            ? (a.codigo ? `${a.codigo} · ` : '') + a.nombre + (a.evento ? ` · ${a.evento}` : '')
            : (n ? 'Datos de ejemplo — usa ▶ para ver a cada alumno' : 'Carga un CSV para ver las constancias reales');
        $('prev-btn').disabled = $('next-btn').disabled = n === 0;
        this.renderCampos();
    },

    conectarTeclado() {
        document.addEventListener('keydown', (e) => {
            const tag = (document.activeElement?.tagName || '').toLowerCase();
            if (['input', 'select', 'textarea'].includes(tag)) return;
            if (document.querySelector('.modal-backdrop.open')) return;

            if (e.key === 'PageDown') { e.preventDefault(); return this.moverVista(1); }
            if (e.key === 'PageUp') { e.preventDefault(); return this.moverVista(-1); }

            const campo = this.campoSel;
            if (!campo || !this.fondo) return;
            const paso = e.shiftKey ? 1 : 0.1;
            const mov = { ArrowLeft: ['x', -paso], ArrowRight: ['x', paso], ArrowUp: ['y', -paso], ArrowDown: ['y', paso] }[e.key];
            if (!mov) return;
            e.preventDefault();
            campo[mov[0]] = Math.round(Math.min(100, Math.max(0, campo[mov[0]] + mov[1])) * 100) / 100;
            this.renderCampos();
            this.actualizarInputsPosicion();
            this.guardarEstado();
        });
    },

    // ------------------------------------------------------------------
    // IMAGEN DE FONDO Y PLANTILLAS
    // ------------------------------------------------------------------
    ponerFondo(dataUrl) {
        return new Promise((resolve, reject) => {
            const img = $('certificate-background');
            img.onload = () => {
                // Imágenes gigantes (fotos de 40+ megapíxeles) hacen los PDF pesadísimos: se reducen a 6000 px
                const MAX = 6000;
                if (Math.max(img.naturalWidth, img.naturalHeight) > MAX) {
                    try {
                        const esc = MAX / Math.max(img.naturalWidth, img.naturalHeight);
                        const c = document.createElement('canvas');
                        c.width = Math.round(img.naturalWidth * esc);
                        c.height = Math.round(img.naturalHeight * esc);
                        c.getContext('2d').drawImage(img, 0, 0, c.width, c.height);
                        dataUrl = /^data:image\/png/i.test(dataUrl) ? c.toDataURL('image/png') : c.toDataURL('image/jpeg', 0.92);
                        UI.toast('La imagen de fondo era muy grande; se redujo a 6000 píxeles para que los PDF no pesen tanto.', 'info');
                        img.src = dataUrl;
                        return; // se vuelve a llamar onload con la imagen reducida
                    } catch (e) { console.warn('No se pudo reducir la imagen', e); }
                }
                this.fondo = { dataUrl, ancho: img.naturalWidth, alto: img.naturalHeight, jpeg: null };
                // Si la imagen es PNG se prepara también una copia JPEG (PDF más ligero y rápido)
                if (/^data:image\/png/i.test(dataUrl)) {
                    try {
                        const c = document.createElement('canvas');
                        c.width = img.naturalWidth;
                        c.height = img.naturalHeight;
                        const ctx = c.getContext('2d');
                        ctx.fillStyle = '#ffffff';
                        ctx.fillRect(0, 0, c.width, c.height);
                        ctx.drawImage(img, 0, 0);
                        this.fondo.jpeg = c.toDataURL('image/jpeg', 0.88);
                    } catch (e) {
                        console.warn('No se pudo preparar la copia JPEG del fondo', e);
                    }
                }
                img.style.display = 'block';
                $('certificate-canvas').classList.remove('empty');
                this.coloresFondo = coloresDeImagen(img);
                this.pintarColoresFondo();
                this.actualizarPagina();
                resolve();
            };
            img.onerror = () => reject(new Error('No se pudo abrir la imagen.'));
            img.src = dataUrl;
        });
    },

    // ------------------------------------------------------------------
    // HOJA: TAMAÑO, ORIENTACIÓN Y CÓMO SE ACOMODA LA IMAGEN
    // ------------------------------------------------------------------
    /** Medidas de la hoja del PDF en mm: { w, h, orientacion } */
    dimensionesPagina() {
        const img = this.fondo;
        const imgVertical = img ? img.alto > img.ancho : false;
        if (this.pagina.tamano === 'imagen' || !TAMANOS_HOJA[this.pagina.tamano]) {
            // Como en la v1: ancho de carta y alto según la proporción de la imagen.
            // Si la imagen es vertical se usa el ancho de carta vertical.
            const w = imgVertical ? 215.9 : PDF_WIDTH_MM;
            const h = img ? (w * img.alto) / img.ancho : 215.9;
            return { w, h, orientacion: imgVertical ? 'vertical' : 'horizontal' };
        }
        const [corto, largo] = TAMANOS_HOJA[this.pagina.tamano];
        return this.pagina.orientacion === 'vertical'
            ? { w: corto, h: largo, orientacion: 'vertical' }
            : { w: largo, h: corto, orientacion: 'horizontal' };
    },

    /** Rectángulo (mm) donde se dibuja la imagen dentro de la hoja */
    colocacionImagen(hoja) {
        const { ancho, alto } = this.fondo;
        if (this.pagina.tamano === 'imagen' || this.pagina.ajuste === 'estirar') {
            return { x: 0, y: 0, w: hoja.w, h: hoja.h };
        }
        const escala = this.pagina.ajuste === 'ajustar'
            ? Math.min(hoja.w / ancho, hoja.h / alto)
            : Math.max(hoja.w / ancho, hoja.h / alto);
        const w = ancho * escala, h = alto * escala;
        return { x: (hoja.w - w) / 2, y: (hoja.h - h) / 2, w, h };
    },

    /** Ajusta el lienzo de la vista previa a la proporción de la hoja */
    ajustarLienzo() {
        const canvas = $('certificate-canvas');
        if (!this.fondo) {
            canvas.style.width = canvas.style.height = '';
            return;
        }
        const zona = $('canvas-scroll');
        const disponibleW = Math.max(200, zona.clientWidth - 48);
        const disponibleH = Math.max(200, zona.clientHeight - 48);
        const hoja = this.dimensionesPagina();
        const proporcion = hoja.w / hoja.h;
        const w = Math.min(disponibleW, disponibleH * proporcion);
        canvas.style.width = Math.floor(w) + 'px';
        canvas.style.height = Math.floor(w / proporcion) + 'px';
        this.renderCampos();
    },

    /** Aplica los cambios de hoja a la interfaz y a la vista previa */
    actualizarPagina() {
        const p = this.pagina;
        const hoja = this.dimensionesPagina();
        $('page-size').value = p.tamano;
        $('page-fit').value = p.ajuste;
        $('page-fit-wrap').style.display = p.tamano === 'imagen' ? 'none' : 'flex';
        document.querySelectorAll('#page-orientation button').forEach(b =>
            b.classList.toggle('active', b.dataset.o === hoja.orientacion));

        const img = $('certificate-background');
        img.style.objectFit = p.tamano === 'imagen' || p.ajuste === 'estirar' ? 'fill'
            : p.ajuste === 'ajustar' ? 'contain' : 'cover';

        let info = `Hoja de ${(hoja.w / 10).toFixed(1)} × ${(hoja.h / 10).toFixed(1)} cm, ${hoja.orientacion}.`;
        if (this.fondo && p.tamano !== 'imagen') {
            const difer = Math.abs((this.fondo.ancho / this.fondo.alto) / (hoja.w / hoja.h) - 1);
            if (difer > 0.03) {
                info += p.ajuste === 'estirar'
                    ? ' ⚠️ La imagen no tiene la proporción de la hoja y se verá deformada; prueba “Rellenar” o “Ajustar”.'
                    : ' La imagen no tiene la misma proporción que la hoja.';
            }
        }
        $('page-info').textContent = this.fondo ? info : '';
        this.ajustarLienzo();
    },

    conectarPagina() {
        $('page-size').addEventListener('change', e => {
            this.pagina.tamano = e.target.value;
            this.actualizarPagina();
            this.guardarEstado();
        });
        $('page-fit').addEventListener('change', e => {
            this.pagina.ajuste = e.target.value;
            this.actualizarPagina();
            this.guardarEstado();
        });
        document.querySelectorAll('#page-orientation button').forEach(b => b.addEventListener('click', () => {
            const hoja = this.dimensionesPagina();
            if (b.dataset.o === hoja.orientacion) return;
            if (this.pagina.tamano === 'imagen') {
                // Con "según la imagen" la orientación la decide la imagen: se pasa a Carta
                this.pagina.tamano = 'carta';
                UI.toast('Se cambió el tamaño de la hoja a Carta para poder elegir la orientación.', 'info');
            }
            this.pagina.orientacion = b.dataset.o;
            this.actualizarPagina();
            this.guardarEstado();
        }));
        this.actualizarPagina();
    },


    conectarPlantilla() {
        $('image-upload').addEventListener('change', (e) => {
            const file = e.target.files[0];
            e.target.value = '';
            if (!file) return;
            if (!/^image\/(png|jpe?g)$/i.test(file.type)) {
                return UI.toast('Usa una imagen JPG o PNG.', 'warning');
            }
            const reader = new FileReader();
            reader.onload = async (ev) => {
                try {
                    await this.ponerFondo(ev.target.result);
                    this.guardarFondo();
                    if (this.fondo.ancho < this.fondo.alto) {
                        UI.toast('La imagen es vertical: el PDF saldrá en orientación vertical.', 'info');
                    }
                } catch (err) {
                    UI.toast(err.message, 'error');
                }
            };
            reader.readAsDataURL(file);
        });

        $('save-template-btn').addEventListener('click', () => this.guardarPlantillaUI());

        $('template-upload').addEventListener('change', async (e) => {
            const file = e.target.files[0];
            e.target.value = '';
            if (!file) return;
            try {
                const datos = JSON.parse(await file.text());
                const nombre = (typeof datos.nombre === 'string' && datos.nombre.trim()) || file.name.replace(/\.json$/i, '');
                await this.aplicarPlantilla(datos, nombre);
                UI.cerrarModal('templates-modal');
                this.recordarPlantilla(nombre);
            } catch (err) {
                console.error(err);
                UI.toast('No se pudo abrir la plantilla:\n' + err.message, 'error');
            }
        });

        // Botón "Asistente de diseño" junto a "Imagen de fondo" (todas las interfaces)
        const lblFondo = $('image-upload').closest('label');
        if (lblFondo && !$('wizard-btn')) {
            const b = document.createElement('button');
            b.id = 'wizard-btn';
            b.className = lblFondo.className.replace(/\b(primary|primario)\b/g, '').trim();
            b.title = 'Crea el fondo de la constancia aquí mismo: marco, colores, logos y firmas';
            b.innerHTML = `<svg class="ico${lblFondo.className.includes('big') ? ' lg' : ''}"><use href="#i-award" /></svg>${lblFondo.className.includes('big') ? 'Asistente de diseño' : 'Crear con el asistente'}`;
            lblFondo.insertAdjacentElement('afterend', b);
        }
        $('wizard-btn')?.addEventListener('click', () => Asistente.abrir());

        // Nueva / Abrir (botones de todas las interfaces y atajos)
        document.querySelectorAll('[data-accion="nueva-plantilla"]').forEach(b => b.addEventListener('click', () => this.nuevaPlantilla()));
        document.querySelectorAll('[data-accion="abrir-plantilla"]').forEach(b => b.addEventListener('click', (e) => {
            e.preventDefault();
            this.abrirPlantillaUI();
        }));
    },

    // ------------------------------------------------------------------
    // PLANTILLAS: nueva, abrir (con recientes) y guardar
    // ------------------------------------------------------------------
    nombrePlantilla: '',

    /** Todo lo que forma el diseño, listo para guardarse */
    datosPlantilla(nombre) {
        return {
            app: 'generador-constancias',
            version: 3,
            nombre: nombre || this.nombrePlantilla || 'Constancia',
            guardada: new Date().toISOString(),
            campos: this.campos,
            nombreArchivo: $('filename-mode').value,
            pagina: this.pagina,
            calibracion: this.calibracion(),
            fondo: this.fondo ? this.fondo.dataUrl : null
        };
    },

    /** Aplica una plantilla (de archivo o de recientes). Valida TODO antes de cambiar nada. */
    async aplicarPlantilla(datos, nombre) {
        if (!datos || datos.app !== 'generador-constancias' || !Array.isArray(datos.campos)) {
            throw new Error('El archivo no es una plantilla de este programa.');
        }
        // Solo se aceptan imágenes incrustadas (nada que se descargue de internet)
        if (datos.fondo && !/^data:image\/(png|jpe?g);base64,/i.test(String(datos.fondo))) {
            throw new Error('La imagen de fondo de la plantilla no es válida.');
        }
        // La imagen se prueba ANTES de cambiar nada (si está dañada, el diseño actual queda intacto)
        if (datos.fondo) await probarImagen(String(datos.fondo));
        const faltan = [...new Set(datos.campos.map(c => c && c.fuente).filter(f => typeof f === 'string' && !Fuentes.familia(f)))].slice(0, 10);
        this.campos = this.completarCampos(datos.campos);
        this.seleccionado = 'nombre';
        if (['nombre', 'codigo_nombre', 'codigo', 'folio_nombre', 'evento_nombre'].includes(datos.nombreArchivo)) {
            $('filename-mode').value = datos.nombreArchivo;
        }
        const pg = datos.pagina || {};
        this.pagina = {
            tamano: ['imagen', 'carta', 'oficio', 'a4'].includes(pg.tamano) ? pg.tamano : 'imagen',
            orientacion: pg.orientacion === 'vertical' ? 'vertical' : 'horizontal',
            ajuste: ['estirar', 'rellenar', 'ajustar'].includes(pg.ajuste) ? pg.ajuste : 'estirar'
        };
        this.actualizarPagina();
        if (datos.calibracion) {
            $('adjust-x-input').value = Math.max(-100, Math.min(100, parseFloat(datos.calibracion.x) || 0));
            $('adjust-y-input').value = Math.max(-100, Math.min(100, parseFloat(datos.calibracion.y) || 0));
            $('adjust-x-input').dispatchEvent(new Event('input'));
            $('adjust-y-input').dispatchEvent(new Event('input'));
        }
        if (datos.fondo) {
            await this.ponerFondo(datos.fondo);
        } else {
            this.quitarFondo();
        }
        this.guardarFondo();
        this.nombrePlantilla = String(nombre || datos.nombre || '').substring(0, 80);
        this.refrescarTodo();
        UI.toast(`Plantilla «${this.nombrePlantilla || 'sin nombre'}» abierta.`, 'success');
        if (faltan.length) {
            UI.toast(`La plantilla usa fuentes que no hay en esta computadora (${faltan.join(', ')}); ` +
                'se cambiaron por Calibri Light. Puedes importarlas en 🔤 Fuentes.', 'warning', { duracion: 10000 });
        }
    },

    coloresFondo: [],
    /** Muestras de color tomadas del fondo, para elegir el color del texto con un clic */
    pintarColoresFondo() {
        const panel = $('props-panel');
        if (!panel) return;
        let box = $('bg-colors');
        if (!box) {
            box = document.createElement('div');
            box.id = 'bg-colors';
            box.className = 'stack';
            box.innerHTML = '<span class="label">Colores de tu fondo</span><div class="bg-swatches" id="bg-swatches"></div>';
            const ancla = $('prop-content');
            ancla.insertAdjacentElement('afterend', box);
            $('bg-swatches').addEventListener('click', e => {
                const b = e.target.closest('[data-color]');
                if (!b || !this.campoSel) return;
                $('prop-color').value = b.dataset.color;
                $('prop-color').dispatchEvent(new Event('input', { bubbles: true }));
                $('prop-color').dispatchEvent(new Event('change', { bubbles: true }));
            });
        }
        const cols = this.coloresFondo || [];
        box.style.display = cols.length ? 'flex' : 'none';
        $('bg-swatches').innerHTML = cols.map(c => `<button data-color="${c}" style="background:${c}" title="${c.toUpperCase()}"></button>`).join('') +
            '<button data-color="#000000" style="background:#000" title="Negro"></button>';
    },

    quitarFondo() {
        this.fondo = null;
        const img = $('certificate-background');
        img.removeAttribute('src');
        img.style.display = 'none';
        $('certificate-canvas').classList.add('empty');
        this.actualizarPagina();
    },

    /** Empieza un diseño desde cero (la lista de personas NO se borra) */
    async nuevaPlantilla() {
        if (Proyectos.activo && !(await Proyectos.preguntarSiNuevoEvento('empezar una plantilla nueva'))) return;
        const hayDiseno = !!this.fondo || this.campos.some(c => c.tipo === 'texto' || c.tipo === 'parrafo' || c.tipo === 'columna');
        if (hayDiseno) {
            const r = await UI.elegir('Nueva plantilla',
                'Se quitará la imagen de fondo y los textos vuelven a su lugar inicial. Tu lista de personas no se borra.\n\n¿Quieres guardar la plantilla actual antes?',
                [{ v: 'guardar', t: 'Guardar y empezar nueva', clase: 'primary' },
                 { v: 'nueva', t: 'Empezar sin guardar' },
                 { v: null, t: 'Cancelar' }]);
            if (!r) return;
            if (r === 'guardar' && !(await this.guardarPlantillaUI())) return;
        }
        this.campos = camposPorDefecto();
        this.seleccionado = 'nombre';
        this.pagina = { tamano: 'imagen', orientacion: 'horizontal', ajuste: 'estirar' };
        this.nombrePlantilla = '';
        this.quitarFondo();
        this.guardarFondo();
        this.refrescarTodo();
        UI.toast('Plantilla nueva. Empieza eligiendo la imagen de fondo.', 'info', { duracion: 5000 });
    },

    /** Miniatura (JPEG pequeño) del fondo con la posición de los textos */
    miniaturaPlantilla() {
        if (!this.fondo) return '';
        try {
            const img = $('certificate-background');
            const hoja = this.dimensionesPagina();
            const W = 240, H = Math.round(W * hoja.h / hoja.w);
            const c = document.createElement('canvas');
            c.width = W; c.height = H;
            const ctx = c.getContext('2d');
            ctx.fillStyle = '#fff';
            ctx.fillRect(0, 0, W, H);
            ctx.drawImage(img, 0, 0, W, H);
            ctx.fillStyle = 'rgba(29,27,51,.55)';
            for (const campo of this.campos) {
                if (!campo.activo) continue;
                const h = Math.max(2, campo.tamano * 0.5 * (W / 280));
                const w = campo.tipo === 'parrafo' ? (campo.anchoMax / hoja.w) * W : Math.min(W * 0.6, h * 9);
                const x = (campo.x / 100) * W, y = (campo.y / 100) * H;
                const x0 = campo.alineacion === 'left' ? x : campo.alineacion === 'right' ? x - w : x - w / 2;
                ctx.fillRect(x0, campo.tipo === 'parrafo' ? y : y - h / 2, w, campo.tipo === 'parrafo' ? h * 2.4 : h);
            }
            return c.toDataURL('image/jpeg', 0.75);
        } catch (e) {
            return '';
        }
    },

    /** Guarda (o actualiza) la plantilla en "recientes" */
    async recordarPlantilla(nombre, datos) {
        try {
            datos = datos || this.datosPlantilla(nombre);
            const id = normalizar(nombre || datos.nombre || 'constancia') || 'constancia';
            await BD.poner('plantillas', { id, nombre: nombre || datos.nombre, usada: Date.now(), miniatura: this.miniaturaPlantilla(), datos });
            // Solo se conservan las 12 más recientes
            const todas = (await BD.todos('plantillas')).sort((a, b) => b.usada - a.usada);
            for (const vieja of todas.slice(12)) await BD.borrar('plantillas', vieja.id);
        } catch (e) {
            console.warn('No se pudo recordar la plantilla', e);
        }
    },

    async guardarPlantillaUI() {
        const nombre = await UI.pedirTexto('Guardar plantilla',
            'Ponle un nombre para encontrarla después (por ejemplo «Constancia taller» o «Reconocimiento ponente»).',
            this.nombrePlantilla || 'Constancia', 'Guardar');
        if (!nombre) return false;
        this.nombrePlantilla = nombre;
        const datos = this.datosPlantilla(nombre);
        descargarBlob(new Blob([JSON.stringify(datos)], { type: 'application/json' }), nombreArchivoSeguro(nombre) + '.json');
        await this.recordarPlantilla(nombre, datos);
        UI.toast(this.fondo ? `Plantilla «${nombre}» guardada (incluye la imagen de fondo). También queda en «Abrir plantilla → Recientes».`
            : `Plantilla «${nombre}» guardada (sin imagen de fondo).`, 'success', { duracion: 5000 });
        return true;
    },

    crearModalPlantillas() {
        if ($('templates-modal')) return;
        const m = document.createElement('div');
        m.className = 'modal-backdrop';
        m.id = 'templates-modal';
        m.innerHTML = `
        <div class="modal" style="width:min(860px,96vw)">
            <div class="modal-header"><h3>📂 Abrir plantilla</h3><button class="close-x" data-cerrar>✕</button></div>
            <div class="modal-body">
                <div class="row" style="align-items:center">
                    <span class="hint" style="flex:1">Tus plantillas recientes en esta computadora. Haz clic en una para abrirla.</span>
                    <button class="btn primary" id="tpl-archivo" style="flex:0;white-space:nowrap">Buscar un archivo…</button>
                </div>
                <div id="tpl-grid" style="display:grid;grid-template-columns:repeat(auto-fill,minmax(180px,1fr));gap:12px"></div>
            </div>
        </div>`;
        document.body.appendChild(m);
        if (!$('tpl-estilos')) {
            const st = document.createElement('style');
            st.id = 'tpl-estilos';
            st.textContent = `
            .tpl-card{border:1px solid var(--borde,#ddd);border-radius:8px;overflow:hidden;cursor:pointer;background:#fff;position:relative;display:flex;flex-direction:column}
            .tpl-card:hover{border-color:var(--oro,#b8913f);box-shadow:0 4px 14px rgba(0,0,0,.1)}
            .tpl-card .tpl-img{aspect-ratio:1.3;background:#f3f3f6 center/cover no-repeat;border-bottom:1px solid var(--borde,#ddd)}
            .tpl-card .tpl-info{padding:7px 9px}
            .tpl-card b{font-size:13px;display:block;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
            .tpl-card small{color:#777;font-size:11.5px}
            .tpl-card .tpl-del{position:absolute;top:5px;right:5px;background:#fff;border:1px solid #ddd;border-radius:5px;cursor:pointer;padding:1px 6px;opacity:0;font-size:12px}
            .tpl-card:hover .tpl-del{opacity:1}
            .tpl-vacio{grid-column:1/-1;text-align:center;color:#777;padding:30px;border:1px dashed #ccc;border-radius:8px}`;
            document.head.appendChild(st);
        }
        m.querySelectorAll('[data-cerrar]').forEach(b => b.addEventListener('click', () => UI.cerrarModal('templates-modal')));
        m.addEventListener('mousedown', e => { if (e.target === m) UI.cerrarModal('templates-modal'); });
        $('tpl-archivo').addEventListener('click', () => $('template-upload').click());
        $('tpl-grid').addEventListener('click', async e => {
            const del = e.target.closest('[data-borrar]');
            if (del) {
                e.stopPropagation();
                if (!await UI.confirmar('Quitar de recientes', `¿Quitar «${del.dataset.nombre}» de la lista? (El archivo .json que guardaste no se borra.)`, 'Quitar', true)) return;
                await BD.borrar('plantillas', del.dataset.borrar);
                return this.abrirPlantillaUI();
            }
            const card = e.target.closest('[data-tpl]');
            if (!card) return;
            try {
                const reg = await BD.leer('plantillas', card.dataset.tpl);
                if (!reg) return;
                await this.aplicarPlantilla(reg.datos, reg.nombre);
                reg.usada = Date.now();
                await BD.poner('plantillas', reg);
                UI.cerrarModal('templates-modal');
            } catch (err) {
                UI.toast('No se pudo abrir la plantilla:\n' + err.message, 'error');
            }
        });
    },

    async abrirPlantillaUI() {
        this.crearModalPlantillas();
        let lista = [];
        try { lista = (await BD.todos('plantillas')).sort((a, b) => b.usada - a.usada); } catch (e) { }
        if (!lista.length) {
            // Sin recientes: directo al explorador de archivos
            return $('template-upload').click();
        }
        const fmt = (t) => new Date(t).toLocaleDateString('es-MX', { day: 'numeric', month: 'short', year: 'numeric' });
        $('tpl-grid').innerHTML = lista.map(p => `
            <div class="tpl-card" data-tpl="${escapeHtml(p.id)}" title="Abrir «${escapeHtml(p.nombre)}»">
                <div class="tpl-img" style="${p.miniatura ? `background-image:url('${p.miniatura}')` : ''}"></div>
                <div class="tpl-info"><b>${escapeHtml(p.nombre)}</b><small>Usada: ${fmt(p.usada)} · ${(p.datos.campos || []).filter(c => c.activo !== false).length} textos</small></div>
                <button class="tpl-del" data-borrar="${escapeHtml(p.id)}" data-nombre="${escapeHtml(p.nombre)}" title="Quitar de recientes">✕</button>
            </div>`).join('');
        UI.abrirModal('templates-modal');
    },

    /** Asegura que existan los campos básicos y que todos tengan todas las propiedades */
    completarCampos(lista) {
        const TIPOS = ['nombre', 'codigo', 'evento', 'folio', 'fecha', 'columna', 'texto', 'parrafo', 'qr', 'imagen'];
        const num = (v, def, min, max) => {
            const n = parseFloat(v);
            return isNaN(n) ? def : Math.min(max, Math.max(min, n));
        };
        const texto = (v, def = '') => (typeof v === 'string' ? v : (v == null ? def : String(v))).substring(0, 500);
        const color = (v, def) => /^#[0-9a-f]{6}$/i.test(String(v)) ? v : def;
        const usados = new Set();
        // Ids reservados: solo pueden ser del tipo correspondiente (si no, el nombre del alumno podría desaparecer)
        const RESERVADOS = { nombre: 'nombre', codigo: 'codigo', curso: 'evento', fecha: 'fecha', folio: 'folio', qr: 'qr' };
        const resultado = (Array.isArray(lista) ? lista : []).filter(c => c && typeof c === 'object' && !Array.isArray(c)).slice(0, 100).map(c0 => {
            const tipo = TIPOS.includes(c0.tipo) ? c0.tipo : 'texto';
            // Solo se toman las propiedades conocidas (nada extra del archivo)
            const c = crearCampo(tipo);
            for (const k of ['negrita', 'cursiva', 'subrayado', 'sombra']) c[k] = c0[k] === true;
            // Id seguro y único (se usa en selectores de la página)
            let id = String(c0.id || tipo).replace(/[^A-Za-z0-9_-]/g, '').substring(0, 40) || tipo;
            const legado = id === 'curso' && tipo === 'texto'; // versiones viejas: "curso" era texto
            if (RESERVADOS[id] && RESERVADOS[id] !== tipo && !legado) id = id + '_' + tipo;
            while (usados.has(id)) id += '_';
            usados.add(id);
            Object.assign(c, {
                id, tipo,
                etiqueta: texto(c0.etiqueta, tipo),
                activo: c0.activo !== false,
                x: num(c0.x, 50, 0, 100), y: num(c0.y, 50, 0, 100),
                tamano: num(c0.tamano, 8, 2, 150),
                color: color(c0.color, '#000000'),
                sombraColor: color(c0.sombraColor, '#000000'),
                alineacion: ['left', 'center', 'right'].includes(c0.alineacion) ? c0.alineacion : 'center',
                anchoMax: num(c0.anchoMax, 0, 0, 1000),
                mayusculas: ['none', 'upper', 'title'].includes(c0.mayusculas) ? c0.mayusculas : 'none',
                prefijo: texto(c0.prefijo), columna: texto(c0.columna),
                valor: (typeof c0.valor === 'string' ? c0.valor : (c0.valor == null ? '' : String(c0.valor))).substring(0, 3000),
                resaltar: c0.resaltar !== false,
                interlineado: num(c0.interlineado, 1.3, 0.8, 3),
                maxRenglones: num(c0.maxRenglones, 0, 0, 30),
                continuarFolio: c0.continuarFolio !== false,
                lado: num(c0.lado, 24, 8, 120),
                qrModo: c0.qrModo === 'url' ? 'url' : 'texto',
                qrUrl: /^https:\/\/[^\s"'<>]{3,300}$/i.test(String(c0.qrUrl || '')) ? String(c0.qrUrl) : '',
                qrLeyenda: c0.qrLeyenda !== false,
                qrFondo: c0.qrFondo !== false,
                imagen: String(c0.imagen || '').length < 15e6 && /^data:image\/(png|jpe?g);base64,[A-Za-z0-9+/=]+$/i.test(String(c0.imagen || '')) ? String(c0.imagen)
                    : (/^idb:[A-Za-z0-9_-]+$/.test(String(c0.imagen || '')) ? String(c0.imagen) : ''),
                anchoImg: num(c0.anchoImg, 45, 3, 600),
                proporcion: num(c0.proporcion, 0.5, 0.02, 50),
                fuente: texto(c0.fuente, 'CalibriLight'),
                folioInicio: num(c0.folioInicio, 1, 0, 1e9), folioDigitos: num(c0.folioDigitos, 4, 1, 10),
                sombraDist: num(c0.sombraDist, 6, 0, 50), sombraOpacidad: num(c0.sombraOpacidad, 45, 0, 100),
                espaciado: num(c0.espaciado, 0, 0, 100)
            });
            return c;
        });
        // El campo "curso" de versiones anteriores ahora es de tipo evento
        resultado.forEach(c => {
            if (c.id === 'curso' && c.tipo === 'texto') {
                c.tipo = 'evento';
                if (c.etiqueta === 'Curso / Evento') c.etiqueta = 'Evento / Taller';
            }
        });
        // Fuentes que no existen en esta computadora → Calibri Light
        resultado.forEach(c => { if (!Fuentes.familia(c.fuente)) c.fuente = 'CalibriLight'; });
        for (const def of camposPorDefecto()) {
            if (!resultado.some(c => c.id === def.id)) {
                def.activo = false;
                resultado.push(def);
            }
        }
        return resultado;
    },

    // ------------------------------------------------------------------
    // CALIBRACIÓN X/Y (mismas claves que la v1 para conservar los ajustes)
    // ------------------------------------------------------------------
    conectarCalibracion() {
        const pares = [['adjust-x-input', 'certificate_adjust_x'], ['adjust-y-input', 'certificate_adjust_y']];
        for (const [idInput, clave] of pares) {
            const inp = $(idInput);
            try {
                const guardado = localStorage.getItem(clave);
                if (guardado !== null) inp.value = parseFloat(guardado) || 0;
            } catch (e) { /* sin almacenamiento */ }
            inp.addEventListener('input', () => {
                try { localStorage.setItem(clave, String(parseFloat(inp.value) || 0)); } catch (e) { }
            });
        }
    },

    calibracion() {
        return {
            x: parseFloat($('adjust-x-input').value) || 0,
            y: parseFloat($('adjust-y-input').value) || 0
        };
    },

    // ------------------------------------------------------------------
    // GUARDADO AUTOMÁTICO DEL DISEÑO
    // ------------------------------------------------------------------
    guardarEstado() {
        clearTimeout(this._guardarTimer);
        this._guardarTimer = setTimeout(() => this.guardarEstadoAhora(), 300);
        Proyectos.tocar();
    },

    guardarEstadoAhora() {
        clearTimeout(this._guardarTimer);
        try {
            // Las imágenes (firmas, sellos) se guardan aparte, en la base de datos local
            const campos = this.campos.map(c => {
                if (c.tipo !== 'imagen' || !c.imagen.startsWith('data:')) return c;
                if (this._imgGuardada?.[c.id] !== c.imagen) {
                    BD.poner('ajustes', { id: 'img_' + c.id, valor: c.imagen }).catch(() => { });
                    (this._imgGuardada = this._imgGuardada || {})[c.id] = c.imagen;
                }
                return Object.assign({}, c, { imagen: 'idb:' + c.id });
            });
            localStorage.setItem(STORAGE_STATE_KEY, JSON.stringify({
                campos,
                seleccionado: this.seleccionado,
                nombreArchivo: $('filename-mode').value,
                pagina: this.pagina,
                opciones: this.opciones,
                nombrePlantilla: this.nombrePlantilla
            }));
        } catch (e) {
            console.warn('No se pudo guardar el estado', e);
        }
    },

    /** Recupera de la base de datos local las imágenes de los campos (firmas, sellos) */
    async cargarImagenesGuardadas() {
        let cambio = false;
        for (const c of this.campos) {
            if (c.tipo !== 'imagen' || !c.imagen.startsWith('idb:')) continue;
            try {
                const r = await BD.leer('ajustes', 'img_' + c.imagen.slice(4));
                if (r && /^data:image\/(png|jpe?g);base64,/i.test(r.valor)) {
                    c.imagen = r.valor;
                    (this._imgGuardada = this._imgGuardada || {})[c.id] = r.valor;
                    cambio = true;
                }
            } catch (e) { }
        }
        if (cambio) this.renderCampos();
    },

    guardarFondo() {
        // El fondo se guarda en la base de datos local (no tiene el límite de ~5 MB de localStorage,
        // y así no le quita espacio al diseño)
        const valor = this.fondo ? this.fondo.dataUrl : '';
        BD.poner('ajustes', { id: 'fondo', valor }).catch(e => console.warn('No se pudo recordar el fondo', e));
        try { localStorage.removeItem(STORAGE_BG_KEY); } catch (_) { }
    },

    /** Antes de pasar a Express / modo completo: ¿la lista cabe para llevarla? (si no, se avisa) */
    async antesDeCambiarDeModo() {
        if (!Alumnos.lista.length) return true;
        try {
            sessionStorage.setItem('gc_prueba_lista', JSON.stringify({ lista: Alumnos.lista, columnas: Alumnos.columnasExtra }));
            sessionStorage.removeItem('gc_prueba_lista');
            return true;
        } catch (e) {
            try { sessionStorage.removeItem('gc_prueba_lista'); } catch (_) { }
            return UI.confirmar('Lista muy grande',
                'Tu lista es demasiado grande para pasarla al otro modo y se perdería. ' +
                'Guárdala antes como proyecto (botón «Proyecto») o vuelve a cargar el archivo después.\n\n¿Cambiar de modo de todos modos?', 'Cambiar', true);
        }
    },

    restaurarEstado() {
        try {
            const estado = JSON.parse(localStorage.getItem(STORAGE_STATE_KEY) || 'null');
            if (estado && Array.isArray(estado.campos)) {
                this.campos = this.completarCampos(estado.campos);
                if (this.campo(estado.seleccionado)) this.seleccionado = estado.seleccionado;
                if (estado.nombreArchivo) $('filename-mode').value = estado.nombreArchivo;
                if (estado.pagina) this.pagina = Object.assign({}, this.pagina, estado.pagina);
                if (estado.opciones) this.opciones = Object.assign({}, this.opciones, estado.opciones);
                if (typeof estado.nombrePlantilla === 'string') this.nombrePlantilla = estado.nombrePlantilla.substring(0, 80);
            }
            // Fondo: base de datos local (o localStorage de versiones anteriores)
            const fondoLS = localStorage.getItem(STORAGE_BG_KEY);
            this._cargaFondo = (async () => {
                let fondo = '';
                try { fondo = (await BD.leer('ajustes', 'fondo'))?.valor || ''; } catch (e) { }
                if (!fondo && fondoLS) fondo = fondoLS;
                if (fondo && /^data:image\/(png|jpe?g);base64,/i.test(fondo)) {
                    await this.ponerFondo(fondo).catch(() => { });
                    if (fondoLS) this.guardarFondo(); // se pasa a la base de datos
                }
            })();
            this.cargarImagenesGuardadas();
        } catch (e) {
            console.warn('No se pudo restaurar el estado anterior', e);
        }
        $('filename-mode').addEventListener('change', () => this.guardarEstado());
        // Guardar lo pendiente justo antes de cerrar el programa o de cambiar de modo
        window.addEventListener('beforeunload', () => {
            this.guardarEstadoAhora();
            // La lista se conserva al cambiar entre Express / cinta con pestañas / cinta única
            // (sessionStorage dura mientras la ventana esté abierta; al cerrar el programa se borra)
            try {
                sessionStorage.removeItem('gc_lista_sesion');
                sessionStorage.setItem('gc_lista_sesion', JSON.stringify({
                    lista: Alumnos.lista, archivos: Alumnos.archivos, columnas: Alumnos.columnasExtra,
                    filtro: this.filtroEvento, indice: this.indiceVista
                }));
            } catch (e) { console.warn('No se pudo conservar la lista', e); }
        });
        try {
            const ses = JSON.parse(sessionStorage.getItem('gc_lista_sesion') || 'null');
            if (ses && Array.isArray(ses.lista)) {
                Alumnos.lista = ses.lista;
                Alumnos.archivos = Array.isArray(ses.archivos) ? ses.archivos : [];
                Alumnos.columnasExtra = Array.isArray(ses.columnas) ? ses.columnas : [];
                this.filtroEvento = typeof ses.filtro === 'string' ? ses.filtro : '__todos__';
                this.indiceVista = Math.min(Number.isInteger(ses.indice) ? ses.indice : -1, Alumnos.lista.length - 1);
            }
        } catch (e) { /* sin lista previa */ }
    },

    // ------------------------------------------------------------------
    // ★★★ GENERACIÓN DE PDFs ★★★
    // ------------------------------------------------------------------
    /** Imagen de fondo que va dentro del PDF: copia JPEG (ligera) o la original PNG */
    imagenParaPdf() {
        const f = this.fondo;
        const esPng = /^data:image\/png/i.test(f.dataUrl);
        if (esPng && this.opciones.formatoFondo === 'jpeg' && f.jpeg) return { data: f.jpeg, fmt: 'JPEG' };
        return { data: f.dataUrl, fmt: esPng ? 'PNG' : 'JPEG' };
    },

    nuevoDoc(hoja) {
        const { jsPDF } = window.jspdf;
        return new jsPDF({
            orientation: hoja.w > hoja.h ? 'landscape' : 'portrait',
            unit: 'mm',
            format: [hoja.w, hoja.h]
        });
    },

    /** Dibuja una constancia completa (fondo + textos) en la página actual */
    dibujarPagina(doc, alumno, indice, hoja, img) {
        // Imagen de fondo (lo que sale de la hoja se recorta solo).
        // El alias 'fondo' hace que en un PDF de varias páginas la imagen se guarde UNA sola vez.
        const r = this.colocacionImagen(hoja);
        doc.addImage(img.data, img.fmt, r.x, r.y, r.w, r.h, 'fondo', 'FAST');

        const cal = this.calibracion();
        for (const campo of this.campos) {
            if (!campo.activo) continue;
            const texto = this.textoDeCampo(campo, alumno, indice);
            if (!texto) continue;
            // posiciónFinal = posición arrastrada + ajuste manual (igual que v1)
            const x = (hoja.w * campo.x) / 100 + cal.x;
            const y = (hoja.h * campo.y) / 100 + cal.y;
            if (campo.tipo === 'parrafo') this.dibujarParrafo(doc, campo, alumno, indice, x, y);
            else if (campo.tipo === 'qr') this.dibujarQR(doc, campo, alumno, indice, x, y);
            else if (campo.tipo === 'imagen') {
                const w = campo.anchoImg, h = campo.anchoImg * campo.proporcion;
                const fmt = /^data:image\/png/i.test(campo.imagen) ? 'PNG' : 'JPEG';
                // Mismo alias = la imagen se guarda una sola vez aunque el PDF tenga muchas páginas
                doc.addImage(campo.imagen, fmt, x - w / 2, y - h / 2, w, h, 'img_' + campo.id, 'FAST');
            }
            else this.dibujarCampo(doc, campo, texto, x, y);
        }
    },

    /** PDF de UNA constancia */
    construirPdf(alumno, indice) {
        const hoja = this.dimensionesPagina();
        const doc = this.nuevoDoc(hoja);
        this.dibujarPagina(doc, alumno, indice, hoja, this.imagenParaPdf());
        return doc;
    },

    /**
     * PDF con varias constancias, una por página.
     * @param items [{alumno, indice}]
     * @param alAvanzar callback(n) opcional para la barra de progreso (permite cancelar devolviendo false)
     */
    async construirPdfMultiple(items, alAvanzar) {
        const hoja = this.dimensionesPagina();
        const img = this.imagenParaPdf();
        const doc = this.nuevoDoc(hoja);
        for (let k = 0; k < items.length; k++) {
            if (k > 0) doc.addPage([hoja.w, hoja.h], hoja.w > hoja.h ? 'landscape' : 'portrait');
            this.dibujarPagina(doc, items[k].alumno, items[k].indice, hoja, img);
            if (alAvanzar && alAvanzar(k + 1) === false) return null;
            if (k % 5 === 4) await new Promise(r => setTimeout(r, 0));
        }
        return doc;
    },

    /**
     * Dibuja un texto con sus estilos en el PDF.
     * La alineación se calcula aquí (se dibuja siempre desde la izquierda) para
     * que el espaciado entre letras, el subrayado y la sombra queden exactos.
     */
    dibujarCampo(doc, campo, texto, x, y) {
        const m = this.medir(campo, texto, doc);
        const { est } = m;
        doc.setFontSize(m.pt);

        let xi = x;
        if (campo.alineacion === 'center') xi = x - m.ancho / 2;
        else if (campo.alineacion === 'right') xi = x - m.ancho;

        const opciones = { baseline: 'middle' };
        if (m.espacio) opciones.charSpace = m.espacio;
        // Cursiva simulada: se inclina el texto con una matriz. jsPDF también
        // inclina la posición (x + 0.2·altura desde abajo), así que se compensa.
        let compensar = 0;
        if (est.fauxItalic) {
            opciones.angle = new doc.Matrix(1, 0, FAUX_ITALIC, 1, 0, 0);
            const lineaBase = y + m.alto / 2 - m.alto * (doc.getLineHeightFactor() - 1);
            compensar = FAUX_ITALIC * (doc.internal.pageSize.getHeight() - lineaBase);
        }
        if (est.fauxBold) {
            opciones.renderingMode = 'fillThenStroke';
            doc.setLineWidth(m.alto * FAUX_BOLD);
        }

        // Subrayado: la línea base está en y + 0.35·alto (baseline "middle" de jsPDF)
        const subrayar = (color, dx, dy) => {
            if (!campo.subrayado) return;
            const yl = y + m.alto * 0.45 + dy;
            doc.setDrawColor(color);
            doc.setLineWidth(m.alto * 0.05);
            doc.line(xi + dx, yl, xi + dx + m.ancho - m.espacio, yl);
            if (est.fauxBold) doc.setLineWidth(m.alto * FAUX_BOLD);
        };

        // 1) Sombra (copia del texto desplazada y semitransparente)
        if (campo.sombra) {
            const d = m.alto * (campo.sombraDist || 0) / 100;
            const op = Math.min(1, Math.max(0, (campo.sombraOpacidad ?? 45) / 100));
            doc.setGState(new doc.GState({ opacity: op, 'stroke-opacity': op }));
            doc.setTextColor(campo.sombraColor);
            doc.setDrawColor(campo.sombraColor);
            doc.text(texto, xi + d - compensar + FAUX_ITALIC * d, y + d, opciones);
            subrayar(campo.sombraColor, d, d);
            doc.setGState(new doc.GState({ opacity: 1, 'stroke-opacity': 1 }));
        }

        // 2) Texto principal
        doc.setTextColor(campo.color);
        doc.setDrawColor(campo.color);
        doc.text(texto, xi - compensar, y, opciones);
        subrayar(campo.color, 0, 0);
    },

    // ------------------------------------------------------------------
    // CONSTANCIA COMO IMAGEN (PNG / JPG) — mismo acomodo que el PDF
    // ------------------------------------------------------------------
    async prepararFuentesImagen() {
        if (!document.fonts) return;
        const familias = new Set(this.campos.filter(c => c.activo).map(c => Fuentes.css(c.fuente)));
        const pedidos = [];
        for (const fam of familias) {
            for (const v of ['400', '700', 'italic 400', 'italic 700']) pedidos.push(document.fonts.load(`${v} 32px ${fam}`).catch(() => { }));
        }
        await Promise.all(pedidos);
    },

    /** Dibuja la constancia en un canvas a la resolución pedida (ppp) */
    dibujarCanvas(alumno, indice, dpi = 150) {
        const hoja = this.dimensionesPagina();
        const ppm = dpi / 25.4; // píxeles por mm
        const c = document.createElement('canvas');
        c.width = Math.round(hoja.w * ppm);
        c.height = Math.round(hoja.h * ppm);
        const ctx = c.getContext('2d');
        ctx.fillStyle = '#ffffff';
        ctx.fillRect(0, 0, c.width, c.height);
        const r = this.colocacionImagen(hoja);
        ctx.drawImage($('certificate-background'), r.x * ppm, r.y * ppm, r.w * ppm, r.h * ppm);
        const k = this.docMedidas().internal.scaleFactor;
        const fuenteCss = (campo, est, px) =>
            `${est.estilo.includes('italic') ? 'italic ' : ''}${est.estilo.includes('bold') ? '700' : '400'} ${px}px ${Fuentes.css(campo.fuente)}`;
        // Texto con negrita/cursiva simuladas
        const texto = (t, x, y, est, px, color, alto) => {
            ctx.save();
            ctx.translate(x, y);
            if (est.fauxItalic) ctx.transform(1, 0, -FAUX_ITALIC, 1, 0, 0);
            ctx.fillStyle = color;
            ctx.fillText(t, 0, 0);
            if (est.fauxBold) {
                ctx.strokeStyle = color;
                ctx.lineWidth = alto * FAUX_BOLD;
                ctx.lineJoin = 'round';
                ctx.strokeText(t, 0, 0);
            }
            ctx.restore();
        };
        ctx.textBaseline = 'middle';
        for (const campo of this.campos) {
            if (!campo.activo) continue;
            const x = (hoja.w * campo.x) / 100, y = (hoja.h * campo.y) / 100;
            if (campo.tipo === 'imagen') {
                if (!campo.imagen.startsWith('data:')) continue;
                const img = this._imgCache?.[campo.id];
                if (!img) continue;
                const w = campo.anchoImg, h = w * campo.proporcion;
                ctx.drawImage(img, (x - w / 2) * ppm, (y - h / 2) * ppm, w * ppm, h * ppm);
                continue;
            }
            if (campo.tipo === 'qr') {
                const q = this.matrizQR(this.contenidoQR(campo, alumno, indice));
                const lado = campo.lado, mod = lado / q.n, x0 = x - lado / 2, y0 = y - lado / 2;
                if (campo.qrFondo) {
                    ctx.fillStyle = '#fff';
                    ctx.fillRect((x0 - mod * 1.5) * ppm, (y0 - mod * 1.5) * ppm, (lado + mod * 3) * ppm, (lado + mod * 3) * ppm);
                }
                ctx.fillStyle = campo.color || '#000';
                for (let rr = 0; rr < q.n; rr++) for (let cc = 0; cc < q.n; cc++) {
                    if (q.oscuro(rr, cc)) ctx.fillRect(Math.floor((x0 + cc * mod) * ppm), Math.floor((y0 + rr * mod) * ppm), Math.ceil(mod * ppm), Math.ceil(mod * ppm));
                }
                if (campo.qrLeyenda) {
                    const pt = Math.max(5, lado * 0.3);
                    ctx.font = `${(pt / k) * ppm}px Helvetica, Arial, sans-serif`;
                    ctx.textAlign = 'center';
                    ctx.fillText(this.leyendaQR(alumno, indice), x * ppm, (y0 + lado + (campo.qrFondo ? mod * 1.5 : 0) + (pt / k) * 0.9) * ppm);
                    ctx.textAlign = 'left';
                }
                continue;
            }
            if (campo.tipo === 'parrafo') {
                const L = this.layoutParrafo(campo, alumno, indice);
                const px = (L.pt / k) * ppm;
                L.lineas.forEach((linea, i) => {
                    const yMid = y + L.lineH * i + L.lineH / 2;
                    let x0 = x;
                    if (campo.alineacion === 'center') x0 = x - linea.ancho / 2;
                    else if (campo.alineacion === 'right') x0 = x - linea.ancho;
                    for (const sg of linea.segs) {
                        if (!sg.t.trim()) continue;
                        const est = L.estilos[sg.b ? 'b' : 'n'];
                        ctx.font = fuenteCss(campo, est, px);
                        texto(sg.t, (x0 + sg.x) * ppm, yMid * ppm, est, px, campo.color, px);
                    }
                });
                continue;
            }
            const t = this.textoDeCampo(campo, alumno, indice);
            if (!t) continue;
            const m = this.medir(campo, t);
            const px = (m.pt / k) * ppm;
            let xi = x;
            if (campo.alineacion === 'center') xi = x - m.ancho / 2;
            else if (campo.alineacion === 'right') xi = x - m.ancho;
            ctx.font = fuenteCss(campo, m.est, px);
            try { ctx.letterSpacing = (m.espacio * ppm) + 'px'; } catch (e) { }
            if (campo.sombra) {
                const d = m.alto * (campo.sombraDist || 0) / 100;
                ctx.globalAlpha = Math.min(1, Math.max(0, (campo.sombraOpacidad ?? 45) / 100));
                texto(t, (xi + d) * ppm, (y + d) * ppm, m.est, px, campo.sombraColor, px);
                ctx.globalAlpha = 1;
            }
            texto(t, xi * ppm, y * ppm, m.est, px, campo.color, px);
            if (campo.subrayado) {
                ctx.fillStyle = campo.color;
                ctx.fillRect(xi * ppm, (y + m.alto * 0.45) * ppm, (m.ancho - m.espacio) * ppm, Math.max(1, m.alto * 0.05 * ppm));
            }
            try { ctx.letterSpacing = '0px'; } catch (e) { }
        }
        return c;
    },

    /** Imagen PNG/JPG de una constancia (bytes) */
    async construirImagen(alumno, indice, ext = 'png', dpi = 150) {
        // Imágenes de los campos (firmas) listas para dibujar
        this._imgCache = this._imgCache || {};
        for (const campo of this.campos) {
            if (campo.tipo !== 'imagen' || !campo.activo || !campo.imagen.startsWith('data:')) continue;
            const prev = this._imgCache[campo.id];
            if (prev && prev.src === campo.imagen) continue;
            this._imgCache[campo.id] = await new Promise((res, rej) => {
                const im = new Image();
                im.onload = () => res(im);
                im.onerror = () => rej(new Error('Imagen dañada'));
                im.src = campo.imagen;
            });
        }
        const canvas = this.dibujarCanvas(alumno, indice, dpi);
        const blob = await new Promise(res => canvas.toBlob(res, ext === 'jpg' ? 'image/jpeg' : 'image/png', 0.92));
        return new Uint8Array(await blob.arrayBuffer());
    },

    nombreDeArchivo(alumno, indice) {
        // Límite de largo para no rebasar el máximo de Windows (260 caracteres con la carpeta)
        return this.nombreDeArchivoCompleto(alumno, indice).substring(0, 120).replace(/[_\-\s]+$/, '');
    },

    nombreDeArchivoCompleto(alumno, indice) {
        const modo = $('filename-mode').value;
        const nombre = nombreArchivoSeguro(alumno ? alumno.nombre : 'Ejemplo');
        const codigo = nombreArchivoSeguro(alumno ? alumno.codigo : '');
        switch (modo) {
            case 'codigo_nombre': return codigo ? `${codigo}_${nombre}` : nombre;
            case 'codigo': return codigo || nombre;
            case 'folio_nombre': return `${nombreArchivoSeguro(this.folioDe(alumno, indice))}_${nombre}`;
            case 'evento_nombre': {
                const ev = nombreArchivoSeguro(alumno ? alumno.evento : '').substring(0, 40);
                return ev ? `${ev}_${nombre}` : nombre;
            }
            default: return nombre;
        }
    },

    validarAntesDeGenerar() {
        if (!this.fondo) {
            UI.toast('Primero sube la imagen de fondo de la constancia.', 'warning');
            return false;
        }
        if (!window.jspdf || !window.jspdf.jsPDF) {
            UI.toast('La librería jsPDF no está cargada. Reinicia el programa.', 'error');
            return false;
        }
        if (!this.campos.some(c => c.activo)) {
            UI.toast('No hay ningún campo activo para imprimir.', 'warning');
            return false;
        }
        const activos = this.campos.filter(c => c.activo);
        // Código QR
        const qr = activos.find(c => c.tipo === 'qr');
        if (qr) {
            if (qr.qrModo === 'url' && !qr.qrUrl) {
                UI.toast('El código QR está en modo «página web» pero falta la dirección (debe empezar con https://). Escríbela o cambia el modo.', 'warning', { duracion: 9000 });
                this.seleccionar(qr.id);
                return false;
            }
            try {
                const muestra = Alumnos.lista.reduce((m, a, i) => (a.nombre.length + (a.evento || '').length > m.l ? { l: a.nombre.length + (a.evento || '').length, a, i } : m), { l: -1, a: null, i: 0 });
                this.matrizQR(this.contenidoQR(qr, muestra.a, muestra.i));
            } catch (e) {
                UI.toast('El código QR tiene demasiado texto. Acorta «Emitida por» o usa el modo página web.', 'warning', { duracion: 9000 });
                this.seleccionar(qr.id);
                return false;
            }
        }
        // Firmas / sellos que aún no terminan de cargarse
        const img = activos.find(c => c.tipo === 'imagen' && !String(c.imagen).startsWith('data:'));
        if (img) {
            UI.toast(`La imagen «${img.etiqueta}» no está cargada. Espera un momento o vuelve a elegirla.`, 'warning');
            this.cargarImagenesGuardadas();
            return false;
        }
        // Párrafos con datos que no existen en la lista: saldrían en blanco
        if (Alumnos.lista.length) {
            for (const p of activos.filter(c => c.tipo === 'parrafo')) {
                const fuera = this.variablesDesconocidas(p);
                if (fuera.length) {
                    UI.toast(`El párrafo «${p.etiqueta}» usa ${fuera.map(v => '{' + v + '}').join(', ')}, pero tu lista no tiene esa columna; saldría en blanco. ` +
                        'Corrige el nombre entre llaves o agrega la columna.', 'warning', { duracion: 10000 });
                    this.seleccionar(p.id);
                    return false;
                }
            }
        }
        return true;
    },

    // ------------------------------------------------------------------
    // ORGANIZAR POR EVENTO, GRUPO, CARRERA… (cualquier columna del CSV/Excel)
    // ------------------------------------------------------------------
    opcionesAgrupar() {
        const ops = [{ v: '', t: 'Sin organizar (todo junto)' }];
        if (Alumnos.eventos().some(e => e.nombre)) ops.push({ v: 'evento', t: 'Evento / taller' });
        // Primero las columnas típicas (Grupo, Carrera…), luego las demás
        const prioridad = (c) => /grupo|carrera|programa|licenciatura|semestre|turno|seccion|escuela|plantel/.test(normalizar(c)) ? 0 : 1;
        [...Alumnos.columnasExtra].sort((a, b) => prioridad(a) - prioridad(b))
            .forEach(c => ops.push({ v: 'col:' + c, t: c }));
        return ops;
    },

    etiquetaAgrupar(clave) {
        if (clave === 'evento') return 'evento';
        if (clave && clave.startsWith('col:')) return clave.slice(4).toLowerCase();
        return '';
    },

    valorGrupo(alumno, clave) {
        if (!clave || !alumno) return '';
        if (clave === 'evento') return alumno.evento || '';
        if (clave.startsWith('col:')) return (alumno.extra || {})[clave.slice(4)] || '';
        return '';
    },

    /** Separa los registros en grupos según la opción "Organizar por" */
    agrupar(items, clave) {
        const grupos = new Map();
        const sinValor = 'Sin ' + (this.etiquetaAgrupar(clave) || 'grupo');
        for (const it of items) {
            const g = clave ? (this.valorGrupo(it.alumno, clave) || sinValor) : '';
            if (!grupos.has(g)) grupos.set(g, []);
            grupos.get(g).push(it);
        }
        return grupos;
    },

    // ------------------------------------------------------------------
    // VENTANA "GENERAR" (opciones de salida)
    // ------------------------------------------------------------------
    abrirGeneracion() {
        if (!this.indicesActivos().length) return UI.toast('No hay alumnos en la lista.', 'warning');
        if (!this.validarAntesDeGenerar()) return;
        if (!Escritorio.tieneArchivos && this.opciones.destino === 'carpeta') this.opciones.destino = 'zip';
        // Opciones de "Organizar por"
        const sel = $('gen-agrupar');
        const ops = this.opcionesAgrupar();
        sel.innerHTML = ops.map(o => `<option value="${escapeHtml(o.v)}">${escapeHtml(o.t)}</option>`).join('');
        if (!ops.some(o => o.v === this.opciones.agruparPor)) this.opciones.agruparPor = '';
        sel.value = this.opciones.agruparPor;
        this.pintarGeneracion();
        UI.abrirModal('gen-modal');
    },

    pintarGeneracion() {
        const op = this.opciones;
        const marcar = (id, v) => document.querySelectorAll(`#${id} button`).forEach(b =>
            b.classList.toggle('active', b.dataset.v === v));
        marcar('gen-modo', op.modo);
        marcar('gen-destino', op.destino);
        marcar('gen-fondo', op.formatoFondo);
        if (!['individual', 'unico', 'imagen'].includes(op.modo)) op.modo = 'individual';
        const esImagen = op.modo === 'imagen';
        if ($('gen-imagen-wrap')) {
            marcar('gen-imgfmt', op.formatoImagen || 'png');
            $('gen-dpi').value = String(op.dpi || 150);
            $('gen-imagen-wrap').style.display = esImagen ? 'flex' : 'none';
            const hoja = this.dimensionesPagina();
            const px = (mm) => Math.round(mm / 25.4 * (op.dpi || 150));
            $('gen-imagen-hint').textContent = `Cada imagen medirá ${px(hoja.w)} × ${px(hoja.h)} píxeles.`;
            $('gen-fondo').closest('.stack').style.display = esImagen ? 'none' : '';
        }

        // Carpeta (solo en la app de escritorio)
        const escritorio = Escritorio.tieneArchivos;
        document.querySelector('#gen-destino [data-v="carpeta"]').disabled = !escritorio;
        $('gen-carpeta-row').style.display = op.destino === 'carpeta' && escritorio ? 'flex' : 'none';
        $('gen-carpeta').value = op.carpeta || '';
        $('gen-destino-hint').textContent = !escritorio
            ? 'Guardar en carpeta solo funciona en la aplicación de escritorio.'
            : op.destino === 'carpeta'
                ? 'Se creará una carpeta nueva dentro de la que elijas. Ideal para muchas constancias.'
                : 'Se descarga un solo archivo .zip con todo adentro.';

        // Organizar
        const clave = op.agruparPor;
        const activos = this.indicesActivos();
        const items = activos.map(i => ({ alumno: Alumnos.lista[i], indice: i }));
        const grupos = this.agrupar(items, clave);
        const etiqueta = this.etiquetaAgrupar(clave);
        $('gen-agrupar-hint').textContent = !Alumnos.columnasExtra.length && !Alumnos.eventos().some(e => e.nombre)
            ? 'Para organizar por grupo o carrera agrega esas columnas a tu CSV o Excel.'
            : clave ? (op.modo === 'unico' ? `Se hará un PDF por cada ${etiqueta}.` : `Se hará una carpeta por cada ${etiqueta}.`) : '';

        // Nombre de archivo: no aplica para "un solo PDF"
        $('gen-nombre-wrap').style.display = op.modo !== 'unico' ? 'flex' : 'none';

        // Fondo
        const esPng = this.fondo && /^data:image\/png/i.test(this.fondo.dataUrl);
        document.querySelector('#gen-fondo [data-v="png"]').disabled = !esPng;
        if (!esPng) {
            $('gen-fondo-hint').textContent = 'Tu imagen ya es JPEG: se usa tal cual.';
        } else {
            // Peso aproximado de la imagen en cada PDF, para decidir con datos
            const kb = (u) => Math.round((u.length * 0.75) / 1024);
            const pngKb = kb(this.fondo.dataUrl), jpgKb = this.fondo.jpeg ? kb(this.fondo.jpeg) : pngKb;
            const ligero = jpgKb <= pngKb ? 'JPEG' : 'PNG';
            $('gen-fondo-hint').textContent = `Peso aproximado por constancia: JPEG ≈ ${jpgKb} KB · PNG ≈ ${pngKb} KB. ` +
                `Para esta imagen conviene ${ligero}. JPEG siempre genera más rápido; PNG conserva la imagen exacta.`;
        }

        // Resumen
        const n = items.length;
        const ev = Alumnos.eventos().find(e => normalizar(e.nombre) === this.filtroEvento);
        let res = `<b>${n} constancia${n === 1 ? '' : 's'}</b>${ev ? ` del evento “${escapeHtml(ev.nombre)}”` : ''}. `;
        if (op.modo === 'unico') {
            res += clave ? `${grupos.size} PDF${grupos.size === 1 ? '' : 's'} (uno por ${escapeHtml(etiqueta)}), una constancia por página.`
                : `1 PDF de ${n} página${n === 1 ? '' : 's'}.`;
        } else {
            const fmtImg = op.formatoImagen === 'jpg' ? 'JPG' : 'PNG';
            const tipoArch = esImagen ? (n === 1 ? `imagen ${fmtImg}` : `imágenes ${fmtImg}`) : `PDF${n === 1 ? '' : 's'}`;
            res += `${n} ${tipoArch}` + (clave ? ` en ${grupos.size} carpeta${grupos.size === 1 ? '' : 's'} por ${escapeHtml(etiqueta)}.` : '.');
        }
        res += op.destino === 'carpeta' && escritorio ? ' Se guardan en una carpeta.' : ' Se descargan en un ZIP.';
        if (op.modo === 'unico' && !clave && op.destino === 'zip') res = res.replace(' Se descargan en un ZIP.', ' Se descarga el PDF.');
        $('gen-resumen').innerHTML = res;
        $('gen-go').textContent = `Generar ${n}`;
    },

    conectarGeneracion() {
        // Botón "Historial" junto a "Imprimir…" (en todas las interfaces)
        const pb = $('print-btn');
        if (pb && !$('history-btn')) {
            const b = pb.cloneNode(false);
            b.id = 'history-btn';
            b.title = 'Constancias ya emitidas: folios, verificación, reimprimir y reenviar';
            b.innerHTML = '<svg class="ico"><use href="#i-file-text" /></svg>Historial';
            pb.insertAdjacentElement('afterend', b);
        }
        $('history-btn')?.addEventListener('click', () => this.abrirHistorial());

        // Solo la constancia que se ve en pantalla (sirve para pruebas)
        $('generate-current-btn').addEventListener('click', () => {
            if (!this.validarAntesDeGenerar()) return;
            try {
                const alumno = this.alumnoActual();
                const doc = this.construirPdf(alumno, Math.max(0, this.indiceVista));
                descargarBlob(doc.output('blob'), (this.nombreDeArchivo(alumno, Math.max(0, this.indiceVista)) || 'constancia') + '.pdf');
                if (alumno) Historial.registrar([{ alumno, indice: Math.max(0, this.indiceVista) }], 'generada');
                UI.toast(alumno ? `PDF de ${alumno.nombre} generado.` : 'PDF de ejemplo generado.', 'success');
            } catch (err) {
                console.error(err);
                UI.toast('Error al generar el PDF:\n' + err.message, 'error');
            }
        });

        $('generate-pdf-btn').addEventListener('click', () => this.abrirGeneracion());
        $('progress-cancel').addEventListener('click', () => {
            this.cancelarGeneracion = true;
            $('progress-info').textContent = 'Cancelando…';
        });

        // Opción "Imágenes (PNG/JPG)" en la ventana Generar (para WhatsApp y redes)
        if (!document.querySelector('#gen-modo [data-v="imagen"]')) {
            const b = document.createElement('button');
            b.dataset.v = 'imagen';
            b.textContent = '🖼️ Imágenes';
            b.title = 'Una imagen PNG o JPG por persona (para WhatsApp o redes sociales)';
            $('gen-modo').appendChild(b);
            const st = document.createElement('div');
            st.className = 'stack';
            st.id = 'gen-imagen-wrap';
            st.innerHTML = `<span class="label">Imagen</span>
                <div class="row" style="align-items:center">
                    <div class="segmented opt-seg" id="gen-imgfmt"><button data-v="png">PNG (nítida)</button><button data-v="jpg">JPG (más ligera)</button></div>
                    <select id="gen-dpi" class="input" style="flex:0;min-width:170px">
                        <option value="100">Calidad normal (100 ppp)</option>
                        <option value="150">Alta (150 ppp)</option>
                        <option value="220">Para imprimir (220 ppp)</option>
                    </select>
                </div>
                <span class="hint" id="gen-imagen-hint"></span>`;
            $('gen-modo').closest('.stack').insertAdjacentElement('afterend', st);
        }

        // Botones de opción de la ventana Generar
        const elegir = (id, prop) => document.querySelectorAll(`#${id} button`).forEach(b =>
            b.addEventListener('click', () => {
                if (b.disabled) return;
                this.opciones[prop] = b.dataset.v;
                this.pintarGeneracion();
                this.guardarEstado();
            }));
        elegir('gen-modo', 'modo');
        elegir('gen-destino', 'destino');
        elegir('gen-fondo', 'formatoFondo');
        elegir('gen-imgfmt', 'formatoImagen');
        $('gen-dpi').addEventListener('change', e => {
            this.opciones.dpi = parseInt(e.target.value, 10) || 150;
            this.pintarGeneracion();
            this.guardarEstado();
        });
        $('gen-agrupar').addEventListener('change', e => {
            this.opciones.agruparPor = e.target.value;
            this.pintarGeneracion();
            this.guardarEstado();
        });
        $('filename-mode').addEventListener('change', () => this.pintarGeneracion());
        $('gen-elegir').addEventListener('click', async () => {
            try {
                const ruta = await Escritorio.invocar('dialogo:carpeta');
                if (ruta) {
                    this.opciones.carpeta = ruta;
                    this.pintarGeneracion();
                    this.guardarEstado();
                }
            } catch (e) {
                UI.toast(e.message, 'error');
            }
        });
        $('gen-go').addEventListener('click', () => {
            const op = this.opciones;
            if (op.destino === 'carpeta' && Escritorio.tieneArchivos) {
                if (!op.carpeta || !Escritorio.fs.existsSync(op.carpeta)) {
                    UI.toast('Elige la carpeta donde se guardarán.', 'warning');
                    return $('gen-elegir').focus();
                }
            }
            UI.cerrarModal('gen-modal');
            this.generarTodas();
        });

        // Imprimir
        $('print-btn').addEventListener('click', () => this.abrirImpresion());
        $('print-go').addEventListener('click', () => this.imprimir());
        ['print-desde', 'print-hasta'].forEach(id => $(id).addEventListener('focus', () => {
            document.querySelector('input[name="print-que"][value="rango"]').checked = true;
        }));
    },

    async generarTodas() {
        const activos = this.indicesActivos();
        if (!activos.length) return UI.toast('No hay alumnos en la lista.', 'warning');
        if (!this.validarAntesDeGenerar()) return;
        if (typeof JSZip === 'undefined') return UI.toast('La librería JSZip no está cargada.', 'error');

        // Aviso si el código está activo y faltan códigos
        const campoCodigo = this.campo('codigo');
        const sinCodigo = activos.filter(i => !Alumnos.lista[i].codigo).length;
        if (campoCodigo?.activo && sinCodigo > 0) {
            const ok = await UI.confirmar('Alumnos sin código',
                `${sinCodigo} alumno(s) no tienen código de estudiante; su constancia saldrá sin código.\n\n¿Continuar de todos modos?`,
                'Continuar');
            if (!ok) return;
        }

        const op = this.opciones;
        const enCarpeta = op.destino === 'carpeta' && Escritorio.tieneArchivos;
        const items = activos.map(i => ({ alumno: Alumnos.lista[i], indice: i }));
        const grupos = this.agrupar(items, op.agruparPor);
        const total = items.length;

        // Nombre base: evento elegido, nombre del archivo cargado o "alumnos"
        const eventoSel = Alumnos.eventos().find(e => normalizar(e.nombre) === this.filtroEvento);
        const base = 'constancias_' + ((eventoSel ? nombreArchivoSeguro(eventoSel.nombre)
            : Alumnos.archivos.length === 1
                ? nombreArchivoSeguro(Alumnos.archivos[0].replace(/\.(csv|xlsx|xlsm|xls|ods)$/i, ''))
                : 'alumnos').substring(0, 50) || 'alumnos');

        // ---- Destino: carpeta (se escribe archivo por archivo) o ZIP ----
        const usados = new Set();
        const rutaUnica = (rel) => {
            const ext = (rel.match(/\.[a-z0-9]+$/i) || ['.pdf'])[0];
            const sin = rel.slice(0, -ext.length);
            let r = rel, k = 2;
            while (usados.has(r.toLowerCase())) r = `${sin}_${k++}${ext}`;
            usados.add(r.toLowerCase());
            return r;
        };
        let raiz = '', zip = null, escritos = 0;
        const escribir = (rel, bytes) => {
            rel = rutaUnica(rel);
            if (enCarpeta) {
                const { fs, path } = Escritorio;
                const ruta = path.join(raiz, ...rel.split('/'));
                fs.mkdirSync(path.dirname(ruta), { recursive: true });
                fs.writeFileSync(ruta, Buffer.from(bytes));
            } else {
                zip.file(rel, bytes);
            }
            escritos++;
        };
        if (enCarpeta) {
            const { fs, path } = Escritorio;
            try {
                raiz = path.join(op.carpeta, base);
                let k = 2;
                while (fs.existsSync(raiz)) raiz = path.join(op.carpeta, `${base}_${k++}`);
                fs.mkdirSync(raiz, { recursive: true });
            } catch (err) {
                UI.toast('No se pudo crear la carpeta para guardar las constancias:\n' + err.message +
                    '\n\nElige otra carpeta (por ejemplo Documentos o el Escritorio).', 'error', { duracion: 12000 });
                return;
            }
        } else {
            zip = new JSZip();
        }
        const carpetaDe = (g) => g ? (nombreArchivoSeguro(g).substring(0, 60) || 'Sin_grupo') + '/' : '';

        // ---- Progreso ----
        this.cancelarGeneracion = false;
        $('progress-title').textContent = 'Generando constancias…';
        UI.abrirModal('progress-modal');
        const inicio = Date.now();
        const pintar = (n, texto) => {
            const pct = Math.round((n / total) * 100);
            $('progress-fill').style.width = pct + '%';
            let t = texto;
            if (n > 3 && n < total) {
                const seg = Math.round(((Date.now() - inicio) / n) * (total - n) / 1000);
                t += seg >= 60 ? ` · faltan ~${Math.ceil(seg / 60)} min` : ` · faltan ~${seg} s`;
            }
            $('progress-pct').textContent = pct + '%';
            $('progress-info').textContent = t;
        };

        const errores = [];
        const listos = [];   // las que sí se generaron (para el historial)
        let hechas = 0;
        try {
            if (op.modo === 'unico') {
                // Un PDF por grupo (o uno solo con todas)
                for (const [g, lista] of grupos) {
                    if (this.cancelarGeneracion) break;
                    const doc = await this.construirPdfMultiple(lista, (n) => {
                        pintar(hechas + n, `${hechas + n} de ${total}` + (g ? ` · ${g}` : ''));
                        return !this.cancelarGeneracion;
                    });
                    if (!doc) break;
                    hechas += lista.length;
                    const nombre = g ? nombreArchivoSeguro(g).substring(0, 80) || 'Sin_grupo' : base;
                    try {
                        escribir(nombre + '.pdf', doc.output('arraybuffer'));
                        listos.push(...lista);
                    } catch (err) {
                        console.error('Error con el grupo', g, err);
                        errores.push(g || 'PDF');
                    }
                    await new Promise(r => setTimeout(r, 0));
                }
            } else if (op.modo === 'imagen') {
                // Una imagen PNG/JPG por persona
                await this.prepararFuentesImagen();
                const ext = op.formatoImagen === 'jpg' ? 'jpg' : 'png';
                for (const [g, lista] of grupos) {
                    for (const { alumno, indice } of lista) {
                        if (this.cancelarGeneracion) break;
                        pintar(hechas, `${hechas + 1} de ${total}: ${alumno.nombre}`);
                        try {
                            const bytes = await this.construirImagen(alumno, indice, ext, op.dpi || 150);
                            const nombre = this.nombreDeArchivo(alumno, indice) || `constancia_${indice + 1}`;
                            escribir(carpetaDe(g) + nombre + '.' + ext, bytes);
                            listos.push({ alumno, indice });
                        } catch (err) {
                            console.error('Error con', alumno.nombre, err);
                            errores.push(alumno.nombre);
                        }
                        hechas++;
                        await new Promise(r => setTimeout(r, 0));
                    }
                    if (this.cancelarGeneracion) break;
                }
            } else {
                // Un PDF por constancia
                for (const [g, lista] of grupos) {
                    for (const { alumno, indice } of lista) {
                        if (this.cancelarGeneracion) break;
                        pintar(hechas, `${hechas + 1} de ${total}: ${alumno.nombre}`);
                        try {
                            const doc = this.construirPdf(alumno, indice);
                            const nombre = this.nombreDeArchivo(alumno, indice) || `constancia_${indice + 1}`;
                            escribir(carpetaDe(g) + nombre + '.pdf', doc.output('arraybuffer'));
                            listos.push({ alumno, indice });
                        } catch (err) {
                            console.error('Error con', alumno.nombre, err);
                            errores.push(alumno.nombre);
                        }
                        hechas++;
                        // Ceder el control a la interfaz para que la barra se actualice
                        await new Promise(r => setTimeout(r, 0));
                    }
                    if (this.cancelarGeneracion) break;
                }
            }

            if (this.cancelarGeneracion) {
                UI.toast(enCarpeta
                    ? `Generación cancelada. ${escritos} archivo(s) quedaron en la carpeta.`
                    : 'Generación cancelada. No se descargó nada.', 'warning');
                return;
            }
            if (!escritos) throw new Error('No se generó ninguna constancia.');

            if (enCarpeta) {
                let msg = `¡Listo! ${listos.length} constancias guardadas en:\n${raiz}`;
                if (errores.length) msg += `\n${errores.length} con error: ${errores.slice(0, 3).join(', ')}${errores.length > 3 ? '…' : ''}`;
                UI.toast(msg, errores.length ? 'warning' : 'success', {
                    duracion: 12000,
                    accion: { texto: 'Abrir carpeta', fn: () => Escritorio.invocar('app:abrirCarpeta', raiz) }
                });
            } else {
                pintar(total, 'Preparando la descarga…');
                const archivos = Object.values(zip.files).filter(f => !f.dir);
                if (op.modo === 'unico' && archivos.length === 1) {
                    // Un solo PDF: se descarga directo, sin comprimir
                    const datos = await archivos[0].async('uint8array');
                    descargarBlob(new Blob([datos], { type: 'application/pdf' }), archivos[0].name);
                } else {
                    const contenido = await zip.generateAsync({
                        type: 'blob', compression: 'DEFLATE', compressionOptions: { level: 6 }
                    });
                    descargarBlob(contenido, `${base}.zip`);
                }
                // En ZIP, al historial solo cuando el archivo ya se descargó
                await Historial.registrar(listos, 'generada');
                let msg = `¡${listos.length} constancias generadas!`;
                if (errores.length) msg += `\n${errores.length} con error: ${errores.slice(0, 3).join(', ')}${errores.length > 3 ? '…' : ''}`;
                UI.toast(msg, errores.length ? 'warning' : 'success', { duracion: 8000 });
            }
        } catch (err) {
            console.error(err);
            UI.toast('Error al generar:\n' + err.message, 'error');
        } finally {
            // En carpeta, lo que quedó escrito va al historial aunque se cancele o falle a la mitad
            // (así esos folios no se vuelven a usar)
            if (enCarpeta && listos.length) await Historial.registrar(listos, 'generada');
            UI.cerrarModal('progress-modal');
        }
    },

    // ------------------------------------------------------------------
    // IMPRIMIR (una o varias constancias)
    // ------------------------------------------------------------------
    abrirImpresion() {
        if (!this.validarAntesDeGenerar()) return;
        const activos = this.indicesActivos();
        const a = this.alumnoActual();
        $('print-actual-label').textContent = a ? a.nombre : 'Ejemplo';
        $('print-todas-label').textContent = `${activos.length}`;
        document.querySelector('input[name="print-que"][value="todas"]').disabled = !activos.length;
        document.querySelector('input[name="print-que"][value="rango"]').disabled = !activos.length;
        $('print-desde').max = $('print-hasta').max = activos.length || 1;
        if (!$('print-desde').value) $('print-desde').value = 1;
        $('print-hasta').value = Math.min(parseInt($('print-hasta').value, 10) || activos.length, activos.length) || 1;
        document.querySelector('input[name="print-que"][value="actual"]').checked = true;
        UI.abrirModal('print-modal');
    },

    async imprimir() {
        const que = document.querySelector('input[name="print-que"]:checked').value;
        const activos = this.indicesActivos();
        let items;
        if (que === 'actual') {
            const a = this.alumnoActual();
            items = [{ alumno: a, indice: Math.max(0, a && !this.manual ? this.indiceVista : 0) }];
        } else if (que === 'todas') {
            items = activos.map(i => ({ alumno: Alumnos.lista[i], indice: i }));
        } else {
            let d = parseInt($('print-desde').value, 10) || 1, h = parseInt($('print-hasta').value, 10) || d;
            d = Math.max(1, Math.min(d, activos.length));
            h = Math.max(d, Math.min(h, activos.length));
            items = activos.slice(d - 1, h).map(i => ({ alumno: Alumnos.lista[i], indice: i }));
        }
        if (!items.length) return UI.toast('No hay constancias para imprimir.', 'warning');
        UI.cerrarModal('print-modal');

        const muchas = items.length > 1;
        if (muchas) {
            this.cancelarGeneracion = false;
            $('progress-title').textContent = 'Preparando impresión…';
            UI.abrirModal('progress-modal');
        }
        try {
            const doc = await this.construirPdfMultiple(items, (n) => {
                if (muchas) {
                    $('progress-fill').style.width = Math.round((n / items.length) * 100) + '%';
                    $('progress-pct').textContent = Math.round((n / items.length) * 100) + '%';
                    $('progress-info').textContent = `${n} de ${items.length}`;
                }
                return !this.cancelarGeneracion;
            });
            if (!doc) return;
            doc.autoPrint(); // el visor abre el cuadro de impresión
            const datos = new Uint8Array(doc.output('arraybuffer'));
            if (Escritorio.disponible) {
                const r = await Escritorio.invocar('imprimir:pdf', datos);
                if (r && r.externo) UI.toast('Se abrió el PDF en tu visor; imprime desde ahí.', 'info');
                Historial.registrar(items, 'impresa');
            } else {
                window.open(URL.createObjectURL(new Blob([datos], { type: 'application/pdf' })), '_blank');
            }
        } catch (err) {
            console.error(err);
            UI.toast('No se pudo preparar la impresión:\n' + err.message, 'error');
        } finally {
            if (muchas) {
                UI.cerrarModal('progress-modal');
                $('progress-title').textContent = 'Generando constancias…';
            }
        }
    }
};


// ==================================================================================
// 6b. ASISTENTE DE DISEÑO (crear la imagen de fondo sin otro programa)
// ==================================================================================
const PALETAS_ASISTENTE = [
    { n: 'Azul y oro', p: '#1d2a5b', a: '#b8913f', f: '#fffdf7' },
    { n: 'Guinda y oro', p: '#6d1f2f', a: '#c19a4b', f: '#fffaf3' },
    { n: 'Verde', p: '#1f5a3d', a: '#b8a04f', f: '#fbfdf8' },
    { n: 'Negro elegante', p: '#1f1f1f', a: '#a88a4a', f: '#ffffff' },
    { n: 'Azul cielo', p: '#1f6fa8', a: '#7cc4e8', f: '#ffffff' },
    { n: 'Morado', p: '#3d2a6b', a: '#c9a3e6', f: '#fdfbff' }
];
const MARCOS_ASISTENTE = [
    ['clasico', 'Clásico'], ['ornamental', 'Ornamental'], ['banda', 'Banda lateral'], ['ondas', 'Ondas'],
    ['moderno', 'Moderno'], ['geometrico', 'Geométrico'], ['minimal', 'Minimalista'], ['medalla', 'Con medalla']
];
const FUENTES_TITULO = [['Cinzel', 'Cinzel (clásica)'], ['PlayfairDisplay', 'Playfair Display'], ['CormorantGaramond', 'Cormorant Garamond'],
    ['EBGaramond', 'EB Garamond'], ['Montserrat', 'Montserrat (moderna)'], ['Oswald', 'Oswald'], ['GreatVibes', 'Great Vibes (caligráfica)']];

const Asistente = {
    o: null,        // opciones actuales
    logos: {},      // posicion → HTMLImageElement

    porDefecto() {
        return {
            hoja: 'carta', orientacion: 'horizontal', marco: 'clasico', paleta: 0,
            p: PALETAS_ASISTENTE[0].p, a: PALETAS_ASISTENTE[0].a, f: PALETAS_ASISTENTE[0].f,
            institucion: '', titulo: 'CONSTANCIA', subtitulo: 'se otorga la presente a', fuente: 'Cinzel',
            tamLogo: 26, firmas: [{ nombre: '', cargo: '' }, { nombre: '', cargo: '' }], nFirmas: 2,
            logos: {} // posicion → dataUrl
        };
    },

    medidas() {
        const [c, l] = TAMANOS_HOJA[this.o.hoja] || TAMANOS_HOJA.carta;
        return this.o.orientacion === 'vertical' ? { w: c, h: l } : { w: l, h: c };
    },

    /** Dibuja el fondo completo. s = píxeles por mm */
    dibujar(ctx, s, soloMarco = false) {
        const o = this.o;
        const { w, h } = this.medidas();
        const W = w * s, H = h * s;
        ctx.save();
        ctx.fillStyle = o.f;
        ctx.fillRect(0, 0, W, H);
        this.marco(ctx, o.marco, W, H, s, o);
        if (soloMarco) { ctx.restore(); return; }

        const fam = (id) => Fuentes.css(id);
        ctx.textAlign = 'center';
        ctx.textBaseline = 'middle';
        const x0 = o.marco === 'banda' ? 30 * s : 0;           // la banda ocupa la izquierda
        const cx = x0 + (W - x0) / 2;
        // Logos
        const lado = o.tamLogo * s;
        const yLogo = H * 0.115;
        const pos = { izq: x0 + (W - x0) * 0.16, centro: cx, der: x0 + (W - x0) * 0.84 };
        for (const [k, img] of Object.entries(this.logos)) {
            if (!img) continue;
            const prop = img.naturalHeight / img.naturalWidth;
            let lw = lado, lh = lado * prop;
            if (lh > lado) { lh = lado; lw = lado / prop; }
            ctx.drawImage(img, pos[k] - lw / 2, yLogo - lh / 2 + (k === 'centro' && o.institucion ? -2 * s : 0), lw, lh);
        }
        // Institución
        if (o.institucion) {
            ctx.fillStyle = o.p;
            ctx.font = `600 ${5.2 * s}px ${fam('Montserrat')}`;
            const yInst = this.logos.centro ? H * 0.115 + lado / 2 + 6 * s : H * 0.13;
            ctx.fillText(o.institucion.toUpperCase(), cx, yInst, (W - x0) * 0.8);
        }
        // Título
        ctx.fillStyle = o.p;
        const tituloPx = (o.fuente === 'GreatVibes' ? 22 : 15) * s;
        ctx.font = `700 ${tituloPx}px ${fam(o.fuente)}`;
        try { ctx.letterSpacing = (o.fuente === 'GreatVibes' ? 0 : 2.2 * s) + 'px'; } catch (e) { }
        ctx.fillText(o.titulo, cx, H * 0.28, (W - x0) * 0.85);
        try { ctx.letterSpacing = '0px'; } catch (e) { }
        // Adorno bajo el título
        ctx.strokeStyle = o.a;
        ctx.lineWidth = 0.6 * s;
        ctx.beginPath();
        ctx.moveTo(cx - 34 * s, H * 0.345); ctx.lineTo(cx - 4 * s, H * 0.345);
        ctx.moveTo(cx + 4 * s, H * 0.345); ctx.lineTo(cx + 34 * s, H * 0.345);
        ctx.stroke();
        ctx.fillStyle = o.a;
        ctx.beginPath();
        ctx.moveTo(cx, H * 0.345 - 2 * s); ctx.lineTo(cx + 2 * s, H * 0.345); ctx.lineTo(cx, H * 0.345 + 2 * s); ctx.lineTo(cx - 2 * s, H * 0.345);
        ctx.fill();
        // Subtítulo
        if (o.subtitulo) {
            ctx.fillStyle = '#444';
            ctx.font = `italic 400 ${5.6 * s}px ${fam('CormorantGaramond')}`;
            ctx.fillText(o.subtitulo, cx, H * 0.40, (W - x0) * 0.8);
        }
        // Línea para el nombre (el nombre lo pone el programa)
        ctx.strokeStyle = o.a;
        ctx.lineWidth = 0.4 * s;
        ctx.beginPath();
        ctx.moveTo(cx - (W - x0) * 0.3, H * 0.555); ctx.lineTo(cx + (W - x0) * 0.3, H * 0.555);
        ctx.stroke();
        // Firmas
        const n = Math.max(0, Math.min(3, o.nFirmas | 0));
        const xs = n === 1 ? [0.5] : n === 2 ? [0.3, 0.7] : [0.2, 0.5, 0.8];
        for (let i = 0; i < n; i++) {
            const fx = x0 + (W - x0) * xs[i];
            const fy = H * 0.845;
            const ancho = (W - x0) * (n === 3 ? 0.2 : 0.26);
            ctx.strokeStyle = '#555';
            ctx.lineWidth = 0.3 * s;
            ctx.beginPath(); ctx.moveTo(fx - ancho / 2, fy); ctx.lineTo(fx + ancho / 2, fy); ctx.stroke();
            const fir = o.firmas[i] || {};
            ctx.fillStyle = '#222';
            ctx.font = `600 ${3.9 * s}px ${fam('Lora')}`;
            if (fir.nombre) ctx.fillText(fir.nombre, fx, fy + 4.2 * s, ancho * 1.15);
            ctx.fillStyle = '#555';
            ctx.font = `400 ${3.3 * s}px ${fam('Lora')}`;
            if (fir.cargo) ctx.fillText(fir.cargo, fx, fy + (fir.nombre ? 8.6 : 4.2) * s, ancho * 1.15);
        }
        ctx.restore();
    },

    marco(ctx, tipo, W, H, s, o) {
        const m = 8 * s;
        ctx.lineJoin = 'miter';
        const rect = (x, y, w, h, color, lw) => { ctx.strokeStyle = color; ctx.lineWidth = lw; ctx.strokeRect(x, y, w, h); };
        switch (tipo) {
            case 'clasico': {
                rect(m, m, W - 2 * m, H - 2 * m, o.p, 2.2 * s);
                rect(m + 3.2 * s, m + 3.2 * s, W - 2 * m - 6.4 * s, H - 2 * m - 6.4 * s, o.a, 0.6 * s);
                ctx.fillStyle = o.a;
                for (const [x, y] of [[m, m], [W - m, m], [m, H - m], [W - m, H - m]]) ctx.fillRect(x - 2.6 * s, y - 2.6 * s, 5.2 * s, 5.2 * s);
                break;
            }
            case 'ornamental': {
                rect(m, m, W - 2 * m, H - 2 * m, o.a, 0.8 * s);
                rect(m + 2.4 * s, m + 2.4 * s, W - 2 * m - 4.8 * s, H - 2 * m - 4.8 * s, o.p, 0.4 * s);
                const esquina = (x, y, sx, sy) => {
                    ctx.save(); ctx.translate(x, y); ctx.scale(sx, sy);
                    ctx.strokeStyle = o.a; ctx.lineWidth = 0.7 * s;
                    for (let r = 6; r <= 18; r += 6) { ctx.beginPath(); ctx.arc(0, 0, r * s, 0, Math.PI / 2); ctx.stroke(); }
                    ctx.fillStyle = o.p;
                    ctx.beginPath(); ctx.moveTo(0, -0.5 * s); ctx.lineTo(4 * s, 4 * s); ctx.lineTo(-0.5 * s, 0); ctx.fill();
                    for (const [dx, dy] of [[22, 3], [3, 22]]) {
                        ctx.beginPath(); ctx.arc(dx * s, dy * s, 1.1 * s, 0, Math.PI * 2); ctx.fill();
                    }
                    ctx.restore();
                };
                esquina(m + 2.4 * s, m + 2.4 * s, 1, 1); esquina(W - m - 2.4 * s, m + 2.4 * s, -1, 1);
                esquina(m + 2.4 * s, H - m - 2.4 * s, 1, -1); esquina(W - m - 2.4 * s, H - m - 2.4 * s, -1, -1);
                break;
            }
            case 'banda': {
                ctx.fillStyle = o.p; ctx.fillRect(0, 0, 26 * s, H);
                ctx.fillStyle = o.a; ctx.fillRect(26 * s, 0, 2.2 * s, H);
                ctx.globalAlpha = 0.15; ctx.strokeStyle = '#fff'; ctx.lineWidth = 0.5 * s;
                for (let y = -26; y < H / s + 26; y += 7) { ctx.beginPath(); ctx.moveTo(0, y * s); ctx.lineTo(26 * s, (y + 26) * s); ctx.stroke(); }
                ctx.globalAlpha = 1;
                rect(34 * s, m, W - 34 * s - m, H - 2 * m, o.a, 0.5 * s);
                break;
            }
            case 'ondas': {
                const ola = (y0, alto, color, arriba) => {
                    ctx.fillStyle = color; ctx.beginPath();
                    ctx.moveTo(0, arriba ? 0 : H);
                    for (let x = 0; x <= W + 1; x += W / 60) ctx.lineTo(x, y0 + Math.sin(x / W * Math.PI * 2.2 + (arriba ? 0 : 1.7)) * alto);
                    ctx.lineTo(W, arriba ? 0 : H); ctx.closePath(); ctx.fill();
                };
                ola(14 * s, 4 * s, o.a, true); ola(10 * s, 4 * s, o.p, true);
                ola(H - 14 * s, 4 * s, o.a, false); ola(H - 10 * s, 4 * s, o.p, false);
                break;
            }
            case 'moderno': {
                ctx.fillStyle = o.p;
                ctx.beginPath(); ctx.moveTo(0, 0); ctx.lineTo(70 * s, 0); ctx.lineTo(0, 55 * s); ctx.fill();
                ctx.fillStyle = o.a;
                ctx.beginPath(); ctx.moveTo(0, 0); ctx.lineTo(40 * s, 0); ctx.lineTo(0, 32 * s); ctx.fill();
                ctx.fillStyle = o.p;
                ctx.beginPath(); ctx.moveTo(W, H); ctx.lineTo(W - 70 * s, H); ctx.lineTo(W, H - 55 * s); ctx.fill();
                ctx.fillStyle = o.a;
                ctx.beginPath(); ctx.moveTo(W, H); ctx.lineTo(W - 40 * s, H); ctx.lineTo(W, H - 32 * s); ctx.fill();
                rect(m, m, W - 2 * m, H - 2 * m, o.a, 0.4 * s);
                break;
            }
            case 'geometrico': {
                const fila = (y, color, alto, inv) => {
                    ctx.fillStyle = color;
                    const paso = 9 * s;
                    for (let x = 0; x < W + paso; x += paso) {
                        ctx.beginPath();
                        ctx.moveTo(x, y); ctx.lineTo(x + paso / 2, y + (inv ? -alto : alto)); ctx.lineTo(x + paso, y); ctx.fill();
                    }
                };
                ctx.fillStyle = o.p; ctx.fillRect(0, 0, W, 6 * s); ctx.fillRect(0, H - 6 * s, W, 6 * s);
                fila(6 * s, o.p, 5 * s, false); fila(H - 6 * s, o.p, 5 * s, true);
                ctx.fillStyle = o.a; ctx.fillRect(0, 12 * s, W, 0.8 * s); ctx.fillRect(0, H - 12.8 * s, W, 0.8 * s);
                break;
            }
            case 'minimal': {
                rect(m * 1.4, m * 1.4, W - 2.8 * m, H - 2.8 * m, o.a, 0.35 * s);
                ctx.fillStyle = o.p; ctx.fillRect(W / 2 - 20 * s, m * 1.4 - 0.9 * s, 40 * s, 1.8 * s);
                ctx.fillRect(W / 2 - 20 * s, H - m * 1.4 - 0.9 * s, 40 * s, 1.8 * s);
                break;
            }
            case 'medalla': {
                rect(m, m, W - 2 * m, H - 2 * m, o.p, 1.4 * s);
                rect(m + 2 * s, m + 2 * s, W - 2 * m - 4 * s, H - 2 * m - 4 * s, o.a, 0.4 * s);
                // Medalla con listones (abajo a la izquierda, sin tapar firmas ni nombre)
                const cx = W - 34 * s, cy = H * 0.66, r = 13 * s;
                ctx.fillStyle = o.p;
                for (const dx of [-5, 5]) {
                    ctx.beginPath(); ctx.moveTo(cx + dx * s - 3.5 * s, cy); ctx.lineTo(cx + dx * s + 3.5 * s, cy);
                    ctx.lineTo(cx + dx * s * 1.6 + 3.5 * s, cy + r + 10 * s); ctx.lineTo(cx + dx * s * 1.6, cy + r + 6 * s);
                    ctx.lineTo(cx + dx * s * 1.6 - 3.5 * s, cy + r + 10 * s); ctx.closePath(); ctx.fill();
                }
                ctx.fillStyle = o.a; ctx.beginPath();
                for (let i = 0; i < 48; i++) {
                    const ang = i / 48 * Math.PI * 2, rr = i % 2 ? r : r * 1.12;
                    ctx[i ? 'lineTo' : 'moveTo'](cx + Math.cos(ang) * rr, cy + Math.sin(ang) * rr);
                }
                ctx.closePath(); ctx.fill();
                ctx.strokeStyle = '#fff'; ctx.lineWidth = 0.6 * s;
                ctx.beginPath(); ctx.arc(cx, cy, r * 0.75, 0, Math.PI * 2); ctx.stroke();
                ctx.fillStyle = '#fff';
                ctx.beginPath();
                for (let i = 0; i < 10; i++) {
                    const ang = -Math.PI / 2 + i * Math.PI / 5, rr = i % 2 ? r * 0.22 : r * 0.5;
                    ctx[i ? 'lineTo' : 'moveTo'](cx + Math.cos(ang) * rr, cy + Math.sin(ang) * rr);
                }
                ctx.closePath(); ctx.fill();
                break;
            }
        }
    },

    // ---------------- Ventana ----------------
    crearModal() {
        if ($('wizard-modal')) return;
        const m = document.createElement('div');
        m.className = 'modal-backdrop';
        m.id = 'wizard-modal';
        const o = this.o;
        m.innerHTML = `
        <div class="modal" style="width:min(1240px,97vw);height:min(760px,94vh)">
            <div class="modal-header"><h3>🎨 Asistente de diseño</h3><button class="close-x" data-cerrar>✕</button></div>
            <div class="modal-body" style="flex-direction:row;gap:16px;overflow:hidden;flex:1;min-height:0">
                <div class="wz-panel">
                    <div class="wz-sec"><b>1 · Hoja</b>
                        <div class="row"><select id="wz-hoja" class="input"><option value="carta">Carta</option><option value="a4">A4</option><option value="oficio">Oficio</option></select>
                        <select id="wz-orient" class="input"><option value="horizontal">Horizontal</option><option value="vertical">Vertical</option></select></div></div>
                    <div class="wz-sec"><b>2 · Marco</b><div id="wz-marcos" class="wz-marcos"></div></div>
                    <div class="wz-sec"><b>3 · Colores</b><div id="wz-paletas" class="wz-paletas"></div>
                        <div class="row" style="align-items:center;font-size:12px">
                            <label>Principal <input type="color" id="wz-p"></label><label>Acento <input type="color" id="wz-a"></label><label>Fondo <input type="color" id="wz-f"></label></div></div>
                    <div class="wz-sec"><b>4 · Logos</b> <span class="hint">(PNG con fondo transparente se ve mejor)</span>
                        <div class="wz-logos">${['izq', 'centro', 'der'].map(k => `<div class="wz-logo" data-pos="${k}"><span>${{ izq: 'Izquierda', centro: 'Centro', der: 'Derecha' }[k]}</span>
                            <div class="wz-logo-img" id="wz-logo-${k}">＋</div><button class="icon-btn" data-quitar-logo="${k}" title="Quitar">✕</button></div>`).join('')}</div>
                        <label style="font-size:12px">Tamaño de logos <input type="range" id="wz-tamlogo" min="14" max="45"></label></div>
                    <div class="wz-sec"><b>5 · Textos</b>
                        <input id="wz-inst" class="input" placeholder="Institución (opcional), ej. CUCSur · Universidad de Guadalajara">
                        <div class="row"><input id="wz-titulo" class="input" placeholder="Título" style="flex:2"><select id="wz-fuente" class="input" style="flex:2">${FUENTES_TITULO.map(([id, n]) => `<option value="${id}">${n}</option>`).join('')}</select></div>
                        <input id="wz-sub" class="input" placeholder="Subtítulo, ej. se otorga la presente a"></div>
                    <div class="wz-sec"><b>6 · Firmas</b>
                        <select id="wz-nfirmas" class="input"><option value="0">Sin firmas</option><option value="1">1 firma</option><option value="2">2 firmas</option><option value="3">3 firmas</option></select>
                        <div id="wz-firmas"></div></div>
                </div>
                <div class="wz-vista"><canvas id="wz-canvas"></canvas>
                    <span class="hint">El nombre de cada persona y los demás datos se ponen después, sobre la línea, como siempre.</span></div>
            </div>
            <div class="modal-footer"><button class="btn" data-cerrar>Cancelar</button><button class="btn primary" id="wz-usar">✓ Usar este diseño</button></div>
        </div>`;
        document.body.appendChild(m);
        const st = document.createElement('style');
        st.textContent = `
            .wz-panel{width:400px;flex:0 0 400px;overflow:auto;display:flex;flex-direction:column;gap:12px;padding-right:6px}
            .wz-sec{display:flex;flex-direction:column;gap:6px;border-bottom:1px solid var(--borde,#ddd);padding-bottom:10px}
            .wz-sec .row{display:flex;gap:6px}
            .wz-marcos{display:grid;grid-template-columns:repeat(4,1fr);gap:6px}
            .wz-marcos button{border:2px solid transparent;border-radius:6px;padding:2px;background:#f6f6f8;cursor:pointer;font-size:10.5px;display:flex;flex-direction:column;align-items:center;gap:2px}
            .wz-marcos button.on{border-color:var(--oro,#b8913f);background:#fff}
            .wz-marcos canvas{width:100%;border-radius:3px;box-shadow:0 1px 3px rgba(0,0,0,.15)}
            .wz-paletas{display:flex;gap:6px;flex-wrap:wrap}
            .wz-paletas button{display:flex;border:2px solid transparent;border-radius:6px;overflow:hidden;padding:0;cursor:pointer;height:26px}
            .wz-paletas button.on{border-color:var(--oro,#b8913f)}
            .wz-paletas span{width:16px;height:22px;display:block}
            .wz-logos{display:flex;gap:8px}
            .wz-logo{flex:1;display:flex;flex-direction:column;align-items:center;gap:3px;font-size:11.5px;position:relative}
            .wz-logo-img{width:100%;height:58px;border:1px dashed #bbb;border-radius:6px;display:grid;place-items:center;cursor:pointer;color:#999;font-size:20px;background:#fafafa center/contain no-repeat}
            .wz-logo .icon-btn{position:absolute;top:16px;right:2px;font-size:11px}
            .wz-vista{flex:1;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:8px;background:#e9e9ef;border-radius:8px;padding:14px;min-width:0}
            .wz-vista canvas{max-width:100%;max-height:calc(100% - 30px);box-shadow:0 8px 24px rgba(0,0,0,.18);background:#fff}
            #wz-firmas .row{margin-top:4px}`;
        document.head.appendChild(st);

        m.querySelectorAll('[data-cerrar]').forEach(b => b.addEventListener('click', () => UI.cerrarModal('wizard-modal')));
        const enlazar = (id, prop, numero) => {
            const el = $(id);
            el.value = o[prop];
            el.addEventListener('input', () => { o[prop] = numero ? +el.value : el.value; this.pintar(); });
        };
        enlazar('wz-hoja', 'hoja'); enlazar('wz-orient', 'orientacion');
        enlazar('wz-inst', 'institucion'); enlazar('wz-titulo', 'titulo'); enlazar('wz-sub', 'subtitulo');
        enlazar('wz-fuente', 'fuente'); enlazar('wz-tamlogo', 'tamLogo', true);
        for (const k of ['p', 'a', 'f']) enlazar('wz-' + k, k);
        $('wz-hoja').addEventListener('change', () => this.pintarMarcos());
        $('wz-orient').addEventListener('change', () => this.pintarMarcos());
        $('wz-nfirmas').value = String(o.nFirmas);
        $('wz-nfirmas').addEventListener('change', e => { o.nFirmas = +e.target.value; this.pintarFirmas(); this.pintar(); });
        $('wz-marcos').addEventListener('click', e => {
            const b = e.target.closest('[data-marco]');
            if (!b) return;
            o.marco = b.dataset.marco;
            this.pintarMarcos();
            this.pintar();
        });
        $('wz-paletas').innerHTML = PALETAS_ASISTENTE.map((p, i) => `<button data-paleta="${i}" title="${p.n}"><span style="background:${p.p}"></span><span style="background:${p.a}"></span><span style="background:${p.f}"></span></button>`).join('');
        $('wz-paletas').addEventListener('click', e => {
            const b = e.target.closest('[data-paleta]');
            if (!b) return;
            const p = PALETAS_ASISTENTE[+b.dataset.paleta];
            Object.assign(o, { paleta: +b.dataset.paleta, p: p.p, a: p.a, f: p.f });
            for (const k of ['p', 'a', 'f']) $('wz-' + k).value = o[k];
            this.pintarMarcos();
            this.pintar();
        });
        m.querySelectorAll('.wz-logo-img').forEach(el => el.addEventListener('click', () => {
            const pos = el.parentElement.dataset.pos;
            const inp = document.createElement('input');
            inp.type = 'file';
            inp.accept = 'image/png,image/jpeg';
            inp.onchange = async () => {
                const file = inp.files[0];
                if (!file) return;
                try {
                    const { dataUrl } = await reducirImagen(file, 800);
                    o.logos[pos] = dataUrl;
                    await this.cargarLogos();
                    this.pintar();
                } catch (err) { UI.toast(err.message, 'error'); }
            };
            inp.click();
        }));
        m.addEventListener('click', e => {
            const q = e.target.closest('[data-quitar-logo]');
            if (!q) return;
            delete o.logos[q.dataset.quitarLogo];
            this.cargarLogos().then(() => this.pintar());
        });
        $('wz-usar').addEventListener('click', () => this.usar());
    },

    pintarFirmas() {
        const o = this.o;
        while (o.firmas.length < 3) o.firmas.push({ nombre: '', cargo: '' });
        $('wz-firmas').innerHTML = Array.from({ length: o.nFirmas }, (_, i) => `<div class="row">
            <input class="input" data-f="${i}" data-k="nombre" placeholder="Nombre de quien firma ${i + 1}" value="${escapeHtml(o.firmas[i].nombre)}">
            <input class="input" data-f="${i}" data-k="cargo" placeholder="Cargo" value="${escapeHtml(o.firmas[i].cargo)}"></div>`).join('');
        $('wz-firmas').oninput = (e) => {
            const t = e.target;
            if (t.dataset.f === undefined) return;
            o.firmas[+t.dataset.f][t.dataset.k] = t.value;
            this.pintar();
        };
    },

    pintarMarcos() {
        const o = this.o;
        const { w, h } = this.medidas();
        $('wz-marcos').innerHTML = MARCOS_ASISTENTE.map(([id, n]) => `<button data-marco="${id}" class="${o.marco === id ? 'on' : ''}"><canvas data-m="${id}"></canvas>${n}</button>`).join('');
        $('wz-marcos').querySelectorAll('canvas').forEach(c => {
            const s = 80 / w;
            c.width = 80; c.height = Math.round(h * s);
            const ctx = c.getContext('2d');
            ctx.fillStyle = o.f; ctx.fillRect(0, 0, c.width, c.height);
            this.marco(ctx, c.dataset.m, c.width, c.height, s, o);
        });
        $('wz-paletas').querySelectorAll('button').forEach(b => b.classList.toggle('on', +b.dataset.paleta === o.paleta));
    },

    async cargarLogos() {
        this.logos = {};
        for (const [k, url] of Object.entries(this.o.logos || {})) {
            if (!/^data:image\/(png|jpe?g);base64,/i.test(url)) continue;
            this.logos[k] = await new Promise(res => { const im = new Image(); im.onload = () => res(im); im.onerror = () => res(null); im.src = url; });
            const el = $('wz-logo-' + k);
            if (el) { el.style.backgroundImage = `url('${url}')`; el.textContent = ''; }
        }
        for (const k of ['izq', 'centro', 'der']) {
            if (!this.o.logos[k] && $('wz-logo-' + k)) { $('wz-logo-' + k).style.backgroundImage = ''; $('wz-logo-' + k).textContent = '＋'; }
        }
    },

    pintar() {
        clearTimeout(this._t);
        this._t = setTimeout(() => {
            const c = $('wz-canvas');
            const { w, h } = this.medidas();
            const s = 3; // 3 px por mm en la vista
            c.width = Math.round(w * s); c.height = Math.round(h * s);
            this.dibujar(c.getContext('2d'), s);
            this.guardar();
        }, 30);
    },

    guardar() {
        try { localStorage.setItem('gc_asistente', JSON.stringify(Object.assign({}, this.o, { logos: {} }))); } catch (e) { }
        BD.poner('ajustes', { id: 'asistente_logos', valor: this.o.logos || {} }).catch(() => { });
    },

    async abrir() {
        if (!this.o) {
            this.o = this.porDefecto();
            try { Object.assign(this.o, JSON.parse(localStorage.getItem('gc_asistente') || '{}')); } catch (e) { }
            this.o.logos = {};
            try { const r = await BD.leer('ajustes', 'asistente_logos'); if (r && r.valor) this.o.logos = r.valor; } catch (e) { }
        }
        this.crearModal();
        for (const id of FUENTES_TITULO.map(f => f[0]).concat(['Montserrat', 'CormorantGaramond', 'Lora'])) {
            for (const v of ['400', '700', 'italic 400']) document.fonts?.load(`${v} 20px ${Fuentes.css(id)}`).catch(() => { });
        }
        await document.fonts?.ready;
        this.pintarFirmas();
        await this.cargarLogos();
        this.pintarMarcos();
        UI.abrirModal('wizard-modal');
        this.pintar();
    },

    async usar() {
        const E = Editor;
        if (E.fondo) {
            const ok = await UI.confirmar('Usar este diseño', 'Se reemplazará la imagen de fondo actual por este diseño. Tus textos y tu lista se conservan.', 'Reemplazar');
            if (!ok) return;
        }
        const { w, h } = this.medidas();
        const s = 200 / 25.4; // 200 ppp: nítido al imprimir
        const c = document.createElement('canvas');
        c.width = Math.round(w * s); c.height = Math.round(h * s);
        this.dibujar(c.getContext('2d'), s);
        const dataUrl = c.toDataURL('image/jpeg', 0.93);
        E.pagina = { tamano: this.o.hoja, orientacion: this.o.orientacion, ajuste: 'estirar' };
        await E.ponerFondo(dataUrl);
        E.guardarFondo();
        // Acomodar el nombre sobre la línea y lo demás debajo
        const nom = E.campo('nombre');
        if (nom) { nom.x = this.o.marco === 'banda' ? 50 + (30 / w) * 50 : 50; nom.y = 50; nom.alineacion = 'center'; }
        const cod = E.campo('codigo');
        if (cod) { cod.x = nom ? nom.x : 50; cod.y = 60; }
        E.campos.filter(c => c.tipo === 'parrafo').forEach((c, i) => { c.x = nom ? nom.x : 50; c.y = 63 + i * 8; c.alineacion = 'center'; });
        E.actualizarPagina();
        E.refrescarTodo();
        UI.cerrarModal('wizard-modal');
        UI.toast('Diseño listo. Ahora carga tu lista o ajusta los textos.', 'success');
    }
};

// ==================================================================================
// 6c. ARREGLOS SUGERIDOS, COLORES DEL FONDO, PROYECTOS Y BUSCADOR DE ACCIONES (Ctrl+F)
// ==================================================================================

/** Propone cómo arreglar un correo mal escrito ('' si no hay sugerencia clara) */
function sugerirCorreo(c) {
    let s = String(c || '').trim().toLowerCase();
    if (!s) return '';
    s = s.replace(/\s+/g, '').replace(/[,;]+$/, '').replace(/^mailto:/, '');
    s = s.replace(/@@+/g, '@').replace(/\.\.+/g, '.').replace(/\.$/, '');
    // "alumno,udg.mx" → punto
    const partes = s.split('@');
    if (partes.length === 2) partes[1] = partes[1].replace(/,/g, '.');
    s = partes.join('@');
    const dominios = {
        'gmial.com': 'gmail.com', 'gmai.com': 'gmail.com', 'gmal.com': 'gmail.com', 'gamil.com': 'gmail.com', 'gmail.co': 'gmail.com',
        'gmail.con': 'gmail.com', 'gmail.cm': 'gmail.com', 'gmail.om': 'gmail.com', 'gmailcom': 'gmail.com', 'gnail.com': 'gmail.com',
        'hotmial.com': 'hotmail.com', 'hotmai.com': 'hotmail.com', 'hotmail.con': 'hotmail.com', 'hotmail.co': 'hotmail.com', 'hotmal.com': 'hotmail.com',
        'outlok.com': 'outlook.com', 'outlook.con': 'outlook.com', 'yaho.com': 'yahoo.com', 'yahoo.con': 'yahoo.com', 'yahoo.com.mx.': 'yahoo.com.mx',
        'alumnos.udg.mx.': 'alumnos.udg.mx', 'alumnos.udg.com': 'alumnos.udg.mx', 'alumnos.udg': 'alumnos.udg.mx', 'alumno.udg.mx': 'alumnos.udg.mx',
        'academicos.udg': 'academicos.udg.mx', 'udg.com.mx': 'udg.mx'
    };
    const p = s.split('@');
    if (p.length === 2 && dominios[p[1]]) s = p[0] + '@' + dominios[p[1]];
    if (p.length === 1 && /^(gmail|hotmail|outlook|yahoo)\.com$/.test(s.replace(/^[^.]*\./, ''))) return ''; // falta el usuario
    if (s === String(c || '').trim().toLowerCase()) return '';
    return esCorreoValido(s) ? s : '';
}

/** Colores principales de una imagen (para sugerirlos al elegir color de texto) */
function coloresDeImagen(img, cuantos = 6) {
    try {
        const c = document.createElement('canvas');
        const W = 80, H = Math.max(1, Math.round(80 * img.naturalHeight / img.naturalWidth));
        c.width = W; c.height = H;
        const ctx = c.getContext('2d');
        ctx.drawImage(img, 0, 0, W, H);
        const d = ctx.getImageData(0, 0, W, H).data;
        const cubos = new Map();
        for (let i = 0; i < d.length; i += 4) {
            if (d[i + 3] < 200) continue;
            const r = d[i], g = d[i + 1], b = d[i + 2];
            const max = Math.max(r, g, b), min = Math.min(r, g, b);
            if (max > 235 && min > 225) continue; // casi blanco (el papel)
            const clave = (r >> 4) + ',' + (g >> 4) + ',' + (b >> 4);
            const e = cubos.get(clave) || { n: 0, r: 0, g: 0, b: 0, sat: 0 };
            e.n++; e.r += r; e.g += g; e.b += b; e.sat += (max - min);
            cubos.set(clave, e);
        }
        const lista = [...cubos.values()].map(e => ({
            r: e.r / e.n, g: e.g / e.n, b: e.b / e.n, peso: e.n * (1 + e.sat / e.n / 60)
        })).sort((a, b) => b.peso - a.peso);
        const elegidos = [];
        for (const col of lista) {
            if (elegidos.every(o => Math.abs(o.r - col.r) + Math.abs(o.g - col.g) + Math.abs(o.b - col.b) > 70)) elegidos.push(col);
            if (elegidos.length >= cuantos) break;
        }
        const hex = (v) => Math.round(v).toString(16).padStart(2, '0');
        return elegidos.map(o => '#' + hex(o.r) + hex(o.g) + hex(o.b));
    } catch (e) {
        return [];
    }
}

/** Proyectos: cada evento guarda su diseño, su lista y su avance */
const Proyectos = {
    activo: '',
    _t: null,

    init() {
        try { this.activo = localStorage.getItem('gc_proyecto') || ''; } catch (e) { }
    },

    nombreActivo() {
        return this._nombre || '';
    },

    async lista() {
        try { return (await BD.todos('proyectos')).sort((a, b) => b.actualizado - a.actualizado); } catch (e) { return []; }
    },

    datos(nombre) {
        return {
            id: normalizar(nombre) || 'proyecto', nombre, actualizado: Date.now(),
            plantilla: Editor.datosPlantilla(Editor.nombrePlantilla || nombre),
            lista: Alumnos.lista, archivos: Alumnos.archivos, columnas: Alumnos.columnasExtra,
            filtro: Editor.filtroEvento, indice: Editor.indiceVista, miniatura: Editor.miniaturaPlantilla()
        };
    },

    async guardarComo() {
        const nombre = await UI.pedirTexto('Guardar como proyecto',
            'Un proyecto guarda el diseño, la lista de personas y su avance (envíos). Ponle el nombre del evento, por ejemplo «Semana de Teleinformática 2026».',
            this._nombre || Alumnos.eventos().find(e => e.nombre)?.nombre || '', 'Guardar');
        if (!nombre) return false;
        const d = this.datos(nombre.substring(0, 80));
        try {
            const ya = await BD.leer('proyectos', d.id);
            if (ya && d.id !== this.activo &&
                !await UI.confirmar('Ya existe ese proyecto', `Ya tienes un proyecto llamado «${ya.nombre}». ¿Reemplazarlo con lo que tienes ahora?`, 'Reemplazar', true)) return false;
            await BD.poner('proyectos', d);
        } catch (e) {
            UI.toast('No se pudo guardar el proyecto:\n' + e.message, 'error');
            return false;
        }
        clearTimeout(this._t);
        this.activo = d.id; this._nombre = d.nombre;
        this.listo = true;
        try { localStorage.setItem('gc_proyecto', d.id); } catch (e) { }
        UI.toast(`Proyecto «${nombre}» guardado. Los cambios se guardan solos mientras trabajas.`, 'success');
        document.dispatchEvent(new CustomEvent('gc-proyecto'));
        return true;
    },

    /** Guardado automático del proyecto activo */
    tocar() {
        if (!this.activo || !this.listo) return;
        clearTimeout(this._t);
        this._t = setTimeout(async () => {
            try {
                const id = this.activo;
                const previo = await BD.leer('proyectos', id);
                if (!previo) { this.cerrar(); return; }
                if (id !== this.activo || !this.listo) return; // cambió de proyecto mientras tanto
                await BD.poner('proyectos', Object.assign(this.datos(previo.nombre), { id: previo.id }));
                document.dispatchEvent(new CustomEvent('gc-guardado'));
            } catch (e) { console.warn('No se pudo guardar el proyecto', e); }
        }, 1500);
    },

    async abrir(id) {
        let p;
        try { p = await BD.leer('proyectos', id); } catch (e) { return UI.toast('No se pudo leer el proyecto:\n' + e.message, 'error'); }
        if (!p) return UI.toast('No se encontró el proyecto.', 'error');
        clearTimeout(this._t);
        this.listo = false; // que el guardado automático no escriba a medias
        try {
            if (p.plantilla) await Editor.aplicarPlantilla(p.plantilla, p.plantilla.nombre || p.nombre);
        } catch (e) {
            this.listo = true;
            return UI.toast('El diseño del proyecto está dañado:\n' + e.message, 'error');
        }
        Alumnos.lista = Array.isArray(p.lista) ? p.lista : [];
        Alumnos.archivos = Array.isArray(p.archivos) ? p.archivos : [];
        Alumnos.columnasExtra = Array.isArray(p.columnas) ? p.columnas : [];
        Editor.filtroEvento = typeof p.filtro === 'string' ? p.filtro : '__todos__';
        Editor.indiceVista = Math.min(Number.isInteger(p.indice) ? p.indice : 0, Alumnos.lista.length - 1);
        this.activo = p.id; this._nombre = p.nombre;
        this.listo = true;
        try { localStorage.setItem('gc_proyecto', p.id); } catch (e) { }
        Editor.trasCargar(false);
        UI.toast(`Proyecto «${p.nombre}» abierto: ${Alumnos.lista.length} personas.`, 'success');
        document.dispatchEvent(new CustomEvent('gc-proyecto'));
    },

    /**
     * Antes de algo que borra (vaciar la lista, plantilla nueva) con un proyecto abierto:
     * ¿es un evento nuevo (el proyecto queda como está) o se cambia este proyecto?
     * Devuelve false si se cancela.
     */
    async preguntarSiNuevoEvento(accion) {
        const r = await UI.elegir('Proyecto abierto',
            `Estás trabajando en el proyecto «${this._nombre || 'sin nombre'}», que se guarda solo.\n\nAntes de ${accion}: ¿empiezas otro evento o quieres cambiar este proyecto?`,
            [{ v: 'nuevo', t: 'Otro evento (el proyecto queda guardado)', clase: 'primary' },
             { v: 'este', t: 'Cambiar este proyecto' },
             { v: null, t: 'Cancelar' }]);
        if (!r) return false;
        if (r === 'nuevo') {
            clearTimeout(this._t);
            this.cerrar();
            UI.toast('El proyecto quedó guardado como estaba. Ahora trabajas sin proyecto.', 'info');
        }
        return true;
    },

    listo: false,
    /** Al abrir el programa: si había un proyecto abierto y la lista está vacía, se recupera solo */
    async cargarNombre() {
        try {
            if (this.activo) {
                const p = await BD.leer('proyectos', this.activo);
                if (!p) { this.activo = ''; this._nombre = ''; }
                else {
                    this._nombre = p.nombre;
                    if (!Alumnos.lista.length && (p.lista || []).length) await this.abrir(p.id);
                    else if (p.plantilla?.fondo) {
                        // Al cambiar de modo la lista sigue, pero el fondo podría no haberse recuperado: se toma del proyecto
                        await Editor._cargaFondo;
                        if (!Editor.fondo) await Editor.ponerFondo(p.plantilla.fondo).catch(() => { });
                    }
                }
            }
        } catch (e) { }
        this.listo = true;
        document.dispatchEvent(new CustomEvent('gc-proyecto'));
    },

    cerrar() {
        clearTimeout(this._t);
        this.activo = ''; this._nombre = '';
        try { localStorage.removeItem('gc_proyecto'); } catch (e) { }
        document.dispatchEvent(new CustomEvent('gc-proyecto'));
    },

    async borrar(id) {
        await BD.borrar('proyectos', id);
        if (id === this.activo) this.cerrar();
    },

    async elegir() {
        const lista = await this.lista();
        let m = $('projects-modal');
        if (!m) {
            m = document.createElement('div');
            m.className = 'modal-backdrop';
            m.id = 'projects-modal';
            m.innerHTML = `<div class="modal" style="width:min(820px,96vw)">
                <div class="modal-header"><h3>📁 Proyectos (eventos)</h3><button class="close-x" data-cerrar>✕</button></div>
                <div class="modal-body">
                    <div class="row" style="align-items:center"><span class="hint" style="flex:1">Cada proyecto guarda su diseño, su lista de personas y su avance. Ábrelo para seguir donde te quedaste.</span>
                    <button class="btn primary" id="pj-nuevo" style="flex:0;white-space:nowrap">💾 Guardar lo actual como proyecto</button></div>
                    <div id="pj-grid" style="display:grid;grid-template-columns:repeat(auto-fill,minmax(180px,1fr));gap:12px"></div>
                </div>
                <div class="modal-footer"><button class="btn" id="pj-cerrar-proy">Trabajar sin proyecto</button><button class="btn" data-cerrar>Cerrar</button></div></div>`;
            document.body.appendChild(m);
            m.querySelectorAll('[data-cerrar]').forEach(b => b.addEventListener('click', () => UI.cerrarModal('projects-modal')));
            $('pj-nuevo').addEventListener('click', async () => { if (await this.guardarComo()) UI.cerrarModal('projects-modal'); });
            $('pj-cerrar-proy').addEventListener('click', () => { this.cerrar(); UI.cerrarModal('projects-modal'); UI.toast('Trabajando sin proyecto (no se guarda automáticamente).', 'info'); });
            $('pj-grid').addEventListener('click', async e => {
                const del = e.target.closest('[data-borrar]');
                if (del) {
                    e.stopPropagation();
                    if (!await UI.confirmar('Borrar proyecto', `¿Borrar el proyecto «${del.dataset.nombre}»? Su historial de constancias NO se borra.`, 'Borrar', true)) return;
                    await this.borrar(del.dataset.borrar);
                    return this.elegir();
                }
                const card = e.target.closest('[data-proy]');
                if (!card) return;
                if (card.dataset.proy === this.activo) return UI.cerrarModal('projects-modal');
                if (!this.activo && (Alumnos.lista.length || Editor.fondo) &&
                    !await UI.confirmar('Abrir proyecto', 'Lo que tienes ahora (diseño y lista) no está guardado en un proyecto y se reemplazará. ¿Continuar?', 'Abrir proyecto', true)) return;
                UI.cerrarModal('projects-modal');
                await this.abrir(card.dataset.proy);
            });
        }
        const fmt = (t) => new Date(t).toLocaleDateString('es-MX', { day: 'numeric', month: 'short', year: 'numeric' });
        $('pj-grid').innerHTML = lista.length ? lista.map(p => `
            <div class="tpl-card" data-proy="${escapeHtml(p.id)}" style="${p.id === this.activo ? 'border-color:var(--oro,#b8913f);box-shadow:0 0 0 2px var(--oro-claro,#f4ead6)' : ''}">
                <div class="tpl-img" style="${p.miniatura ? `background-image:url('${p.miniatura}')` : ''}"></div>
                <div class="tpl-info"><b>${escapeHtml(p.nombre)}</b><small>${(p.lista || []).length} personas · ${fmt(p.actualizado)}${p.id === this.activo ? ' · abierto' : ''}</small></div>
                <button class="tpl-del" data-borrar="${escapeHtml(p.id)}" data-nombre="${escapeHtml(p.nombre)}" title="Borrar proyecto">✕</button>
            </div>`).join('') : '<div class="tpl-vacio">Aún no tienes proyectos. Guarda lo que tienes ahora con el botón de arriba.</div>';
        Editor.crearModalPlantillas(); // estilos de las tarjetas
        UI.abrirModal('projects-modal');
    }
};

/** Buscador de acciones y personas (Ctrl+F): escribe lo que quieres hacer */
const Comandos = {
    sel: 0,
    items: [],

    acciones() {
        const clic = (id) => () => $(id)?.click();
        const ir = (url, pref) => async () => {
            if (!await Editor.antesDeCambiarDeModo()) return;
            try { localStorage.setItem('gc_interfaz', pref); if (pref !== 'express') localStorage.setItem('gc_interfaz_avanzada', pref); } catch (e) { }
            location.href = url;
        };
        return [
            ['Cargar lista de Excel o CSV', 'archivo importar alumnos personas', clic('csv-upload')],
            ['Pegar tabla de Excel', 'pegar copiar google sheets', () => Editor.abrirPegado()],
            ['Ver / editar la lista', 'lista alumnos corregir tabla', clic('open-list-btn')],
            ['Corregir nombres (mayúsculas y acentos)', 'nombres mayusculas acentos', () => { $('open-list-btn')?.click(); Editor.abrirCorreccionNombres(); }],
            ['Hacer una sola constancia', 'manual una persona escribir', clic('manual-btn')],
            ['Imagen de fondo', 'fondo diseño imagen subir', clic('image-upload')],
            ['Asistente de diseño', 'crear fondo marco logos firmas', () => Asistente.abrir()],
            ['Nueva plantilla', 'plantilla nueva limpiar', () => Editor.nuevaPlantilla()],
            ['Abrir plantilla', 'plantilla abrir recientes', () => Editor.abrirPlantillaUI()],
            ['Guardar plantilla', 'plantilla guardar', () => Editor.guardarPlantillaUI()],
            ['Proyectos (eventos guardados)', 'proyecto evento abrir guardar', () => Proyectos.elegir()],
            ['Guardar como proyecto', 'proyecto evento guardar', () => Proyectos.guardarComo()],
            ['Generar todas las constancias', 'generar pdf zip carpeta todas', clic('generate-pdf-btn')],
            ['PDF de la constancia que estoy viendo', 'pdf una vista actual', clic('generate-current-btn')],
            ['Imprimir', 'imprimir impresora', clic('print-btn')],
            ['Enviar por correo', 'correo email enviar gmail', clic('open-email-btn')],
            ['Historial de constancias', 'historial folios emitidas reimprimir reenviar', () => Editor.abrirHistorial()],
            ['Verificar un código', 'verificar autentica qr codigo', () => { Editor.abrirHistorial(); setTimeout(() => $('hi-verificar')?.click(), 50); }],
            ['Agregar párrafo con datos', 'parrafo texto datos variables', () => Editor.agregarParrafo()],
            ['Agregar código QR', 'qr verificacion', clic('add-qr-field-btn')],
            ['Agregar imagen (firma o sello)', 'firma sello logo imagen', () => Editor.elegirImagenCampo(null)],
            ['Agregar texto libre', 'texto libre', clic('add-text-field-btn')],
            ['Todas las fuentes / importar', 'fuentes tipo de letra ttf', clic('manage-fonts-btn')],
            ['Probar con el nombre más largo', 'nombre largo probar', clic('longest-btn')],
            ['¿Cómo debe ser mi tabla?', 'formato columnas plantilla csv', () => Editor.descargarPlantilla()],
            ['Exportar la lista a Excel', 'exportar excel lista', clic('export-xlsx-btn')],
            ['Tutorial', 'ayuda tutorial', clic('help-btn')],
            ['Buscar versión nueva', 'actualizar version github', () => Actualizaciones.revisar(true)],
            ['Cambiar a Modo Express', 'interfaz express rapido', ir('index_express.html', 'express')],
            ['Cambiar a modo completo', 'interfaz completo avanzado cinta', ir('index.html', 'completo')]
        ].map(([t, k, fn]) => ({ tipo: 'accion', t, k: normalizar(t + ' ' + k), fn }));
    },

    crear() {
        if ($('cmd-modal')) return;
        const m = document.createElement('div');
        m.className = 'modal-backdrop';
        m.id = 'cmd-modal';
        m.style.alignItems = 'flex-start';
        m.style.paddingTop = '12vh';
        m.style.zIndex = '9300';
        m.innerHTML = `<div class="modal" style="width:min(620px,94vw)">
            <div style="padding:10px 12px;border-bottom:1px solid var(--borde,#ddd)"><input id="cmd-input" class="input" style="width:100%;font-size:15px;height:38px" placeholder="¿Qué quieres hacer? Escribe una acción o el nombre de una persona…" autocomplete="off"></div>
            <div id="cmd-list" style="max-height:52vh;overflow:auto;padding:6px"></div>
            <div class="hint" style="padding:6px 12px;border-top:1px solid var(--borde,#ddd)">↑ ↓ para elegir · Enter para hacerlo · Esc para cerrar</div></div>`;
        document.body.appendChild(m);
        const st = document.createElement('style');
        st.textContent = `#cmd-list .cmd{display:flex;align-items:center;gap:10px;padding:8px 10px;border-radius:7px;cursor:pointer;font-size:13.5px}
            #cmd-list .cmd.sel{background:var(--oro-claro,#f4ead6)} #cmd-list .cmd small{color:#888;margin-left:auto;font-size:11.5px}
            #cmd-list .cmd .ic{width:20px;text-align:center}`;
        document.head.appendChild(st);
        m.addEventListener('mousedown', e => { if (e.target === m) this.cerrar(); });
        $('cmd-input').addEventListener('input', () => { this.sel = 0; this.pintar(); });
        $('cmd-input').addEventListener('keydown', e => {
            if (e.key === 'ArrowDown') { e.preventDefault(); this.sel = Math.min(this.items.length - 1, this.sel + 1); this.pintar(false); }
            else if (e.key === 'ArrowUp') { e.preventDefault(); this.sel = Math.max(0, this.sel - 1); this.pintar(false); }
            else if (e.key === 'Enter') { e.preventDefault(); this.ejecutar(this.items[this.sel]); }
            else if (e.key === 'Escape') { e.preventDefault(); this.cerrar(); }
        });
        $('cmd-list').addEventListener('click', e => {
            const d = e.target.closest('[data-k]');
            if (d) this.ejecutar(this.items[+d.dataset.k]);
        });
    },

    abrir(texto = '') {
        this.crear();
        $('cmd-input').value = texto;
        this.sel = 0;
        this.pintar();
        UI.abrirModal('cmd-modal');
        $('cmd-input').focus();
    },

    cerrar() { UI.cerrarModal('cmd-modal'); },

    pintar(recalcular = true) {
        if (recalcular) {
            const q = normalizar($('cmd-input').value);
            const palabras = q.split(' ').filter(Boolean);
            const acc = this.acciones().filter(a => palabras.every(p => a.k.includes(p)));
            let personas = [];
            if (q.length >= 2) {
                Alumnos.lista.forEach((a, i) => {
                    if (personas.length >= 8) return;
                    const t = normalizar(`${a.nombre} ${a.codigo} ${a.correo} ${a.folio}`);
                    if (palabras.every(p => t.includes(p))) personas.push({ tipo: 'persona', t: a.nombre, sub: [a.codigo, a.evento].filter(Boolean).join(' · '), i });
                });
            }
            const campos = q.length >= 2 ? Editor.campos.filter(c => palabras.every(p => normalizar(c.etiqueta).includes(p)))
                .slice(0, 4).map(c => ({ tipo: 'campo', t: 'Editar texto: ' + c.etiqueta, id: c.id })) : [];
            this.items = [...personas, ...acc.slice(0, q ? 12 : 10), ...campos];
        }
        const ic = { accion: '⚡', persona: '👤', campo: '✏️' };
        $('cmd-list').innerHTML = this.items.map((it, k) => `<div class="cmd ${k === this.sel ? 'sel' : ''}" data-k="${k}">
            <span class="ic">${ic[it.tipo]}</span><span>${escapeHtml(it.t)}</span>${it.sub ? `<small>${escapeHtml(it.sub)}</small>` : it.tipo === 'persona' ? '<small>ver su constancia</small>' : ''}</div>`).join('')
            || '<div class="hint" style="padding:12px">Nada coincide. Prueba con otra palabra (por ejemplo «imprimir», «correo» o un nombre).</div>';
        $('cmd-list').querySelector('.sel')?.scrollIntoView({ block: 'nearest' });
    },

    ejecutar(it) {
        if (!it) return;
        this.cerrar();
        if (it.tipo === 'accion') it.fn();
        else if (it.tipo === 'persona') {
            Editor.manual = null;
            if (Editor.filtroEvento !== '__todos__' && !Editor.indicesActivos().includes(it.i)) {
                Editor.filtroEvento = '__todos__';
                Editor.actualizarResumen();
            }
            Editor.indiceVista = it.i;
            Editor.actualizarVista();
        } else if (it.tipo === 'campo') {
            Editor.seleccionar(it.id);
        }
    }
};

// ==================================================================================
// 7. ENVÍO DE CONSTANCIAS POR CORREO
// ==================================================================================
// El envío real lo hace el proceso principal de Electron (main.js + correo.js)
// con la librería nodemailer. Aquí solo se arma cada PDF y el mensaje.

const MENSAJE_POR_DEFECTO = {
    asunto: 'Constancia de participación - {nombre}',
    cuerpo: 'Hola {nombre}:\n\nTe compartimos tu constancia en formato PDF, adjunta a este correo.\n\n¡Gracias por tu participación!\n\nSaludos cordiales.'
};

const AYUDA_PROVEEDOR = {
    gmail: 'Gmail pide una <b>contraseña de aplicación</b> (no tu contraseña normal): activa la verificación en 2 pasos y créala en ' +
        '<a href="https://myaccount.google.com/apppasswords" target="_blank">myaccount.google.com/apppasswords</a>. ' +
        'Si tu correo institucional usa Google, también elige esta opción. Límite aproximado: 500 correos al día.',
    outlook: 'Usa tu correo de Outlook/Office 365. Algunas instituciones bloquean el envío por SMTP; si la prueba falla, ' +
        'pide al área de sistemas que habilite “SMTP autenticado” o usa una cuenta de Gmail.',
    personalizado: 'Escribe los datos del servidor SMTP que te dé tu proveedor de correo.'
};

// Una sola dirección: sin espacios, comas, punto y coma ni signos < > ( ) "
const esCorreoValido = (c) => /^[^\s@,;<>()"]+@[^\s@,;<>()"]+\.[^\s@,;<>()".]+(\.[^\s@,;<>()".]+)*$/.test(String(c || '').trim());

const Envio = {
    enviando: false,
    cancelar: false,

    init() {
        $('open-email-btn').addEventListener('click', () => this.abrir());
        $('smtp-provider').addEventListener('change', () => { this.actualizarProveedor(); this.guardarConfig(); });
        $('smtp-pass-toggle').addEventListener('click', () => {
            const p = $('smtp-pass');
            p.type = p.type === 'password' ? 'text' : 'password';
        });
        ['smtp-host', 'smtp-port', 'smtp-security', 'smtp-user', 'smtp-name', 'smtp-remember',
            'mail-subject', 'mail-body', 'mail-delay', 'mail-only-pending'].forEach(id =>
                $(id).addEventListener('change', () => this.guardarConfig()));
        ['mail-subject', 'mail-body'].forEach(id => $(id).addEventListener('input', () => this.vistaPrevia()));
        $('mail-only-pending').addEventListener('change', () => this.estadisticas());
        $('mail-group').addEventListener('change', () => { this.estadisticas(); this.vistaPrevia(); this.guardarConfig(); });

        // Insertar {etiquetas} donde está el cursor
        let ultimoCampo = $('mail-body');
        $('mail-subject').addEventListener('focus', () => ultimoCampo = $('mail-subject'));
        $('mail-body').addEventListener('focus', () => ultimoCampo = $('mail-body'));
        $('mail-chips').addEventListener('mousedown', e => e.preventDefault());
        $('mail-chips').addEventListener('click', e => {
            const chip = e.target.closest('.chip');
            if (!chip) return;
            const el = ultimoCampo;
            const ini = el.selectionStart ?? el.value.length, fin = el.selectionEnd ?? el.value.length;
            el.value = el.value.slice(0, ini) + chip.dataset.tag + el.value.slice(fin);
            el.focus();
            el.selectionStart = el.selectionEnd = ini + chip.dataset.tag.length;
            this.vistaPrevia();
            this.guardarConfig();
        });

        $('smtp-test-btn').addEventListener('click', () => this.probarConexion());
        $('mail-test-btn').addEventListener('click', () => this.enviarPrueba());
        $('mail-send-btn').addEventListener('click', () => this.enviarTodos());
        $('mail-report-btn').addEventListener('click', () => this.descargarReporte());

        this.cargarConfig();
    },

    // ---------------- Configuración ----------------
    leerFormulario() {
        return {
            proveedor: $('smtp-provider').value,
            host: $('smtp-host').value.trim(),
            puerto: parseInt($('smtp-port').value, 10) || 587,
            seguridad: $('smtp-security').value,
            usuario: $('smtp-user').value.trim(),
            nombre: $('smtp-name').value.trim(),
            // Las contraseñas de aplicación de Gmail se muestran con espacios ("abcd efgh …"): solo ahí se quitan
            contrasena: $('smtp-provider').value === 'gmail'
                ? $('smtp-pass').value.replace(/\s/g, '')
                : $('smtp-pass').value,
            recordar: $('smtp-remember').checked
        };
    },

    async cargarConfig() {
        let cfg = null;
        try {
            if (Escritorio.disponible) cfg = await Escritorio.invocar('correo:leerConfig');
        } catch (e) { console.warn(e); }
        if (!cfg) {
            try { cfg = JSON.parse(localStorage.getItem('gc_correo') || 'null'); } catch (e) { }
        }
        cfg = Object.assign({
            proveedor: 'gmail', host: '', puerto: 587, seguridad: 'starttls', usuario: '', nombre: '',
            contrasena: '', recordar: false, asunto: MENSAJE_POR_DEFECTO.asunto, cuerpo: MENSAJE_POR_DEFECTO.cuerpo,
            pausa: 2, soloPendientes: true
        }, cfg || {});
        $('smtp-provider').value = cfg.proveedor;
        $('smtp-host').value = cfg.host;
        $('smtp-port').value = cfg.puerto;
        $('smtp-security').value = cfg.seguridad;
        $('smtp-user').value = cfg.usuario;
        $('smtp-name').value = cfg.nombre;
        $('smtp-pass').value = cfg.contrasena || '';
        $('smtp-remember').checked = !!cfg.recordar;
        $('mail-subject').value = cfg.asunto;
        $('mail-body').value = cfg.cuerpo;
        $('mail-delay').value = cfg.pausa;
        $('mail-only-pending').checked = cfg.soloPendientes !== false;
        $('mail-group').checked = cfg.agrupar !== false;
        this.actualizarProveedor();
    },

    guardarConfig() {
        const cfg = Object.assign(this.leerFormulario(), {
            asunto: $('mail-subject').value,
            cuerpo: $('mail-body').value,
            pausa: parseFloat($('mail-delay').value) || 0,
            soloPendientes: $('mail-only-pending').checked,
            agrupar: $('mail-group').checked
        });
        if (Escritorio.disponible) {
            Escritorio.invocar('correo:guardarConfig', cfg).catch(e => console.warn(e));
        } else {
            delete cfg.contrasena;
            try { localStorage.setItem('gc_correo', JSON.stringify(cfg)); } catch (e) { }
        }
    },

    actualizarProveedor() {
        const p = $('smtp-provider').value;
        $('smtp-custom').style.display = p === 'personalizado' ? 'flex' : 'none';
        $('smtp-help').innerHTML = AYUDA_PROVEEDOR[p];
        $('smtp-pass').placeholder = p === 'gmail' ? 'xxxx xxxx xxxx xxxx' : 'Contraseña';
    },

    validarCuenta() {
        const c = this.leerFormulario();
        if (!esCorreoValido(c.usuario)) { UI.toast('Escribe tu correo (el que envía).', 'warning'); $('smtp-user').focus(); return null; }
        if (!c.contrasena) { UI.toast('Escribe la contraseña de aplicación.', 'warning'); $('smtp-pass').focus(); return null; }
        if (c.proveedor === 'personalizado' && !c.host) { UI.toast('Escribe el servidor SMTP.', 'warning'); $('smtp-host').focus(); return null; }
        return c;
    },

    /**
     * Envía UNA constancia a UN destinatario (desde la lista o desde "Una constancia").
     * Usa la cuenta configurada en la ventana de correo.
     */
    async enviarUno(alumno, indice, para) {
        if (this._enviandoUno || this.enviando) return UI.toast('Ya se está enviando un correo; espera a que termine.', 'info');
        para = String(para || '').trim().toLowerCase();
        if (!esCorreoValido(para)) {
            UI.toast('Ese alumno no tiene un correo válido. Escríbelo en la lista y vuelve a intentar.', 'warning');
            return;
        }
        if (!Escritorio.disponible) {
            UI.toast('El envío por correo solo funciona en la aplicación de escritorio (Abrir Generador.bat).', 'error');
            return;
        }
        if (!Editor.validarAntesDeGenerar()) return;
        // Si la cuenta aún no está configurada, se abre la ventana de correo
        const c = this.leerFormulario();
        if (!esCorreoValido(c.usuario) || !c.contrasena || (c.proveedor === 'personalizado' && !c.host)) {
            this.abrir();
            UI.toast('Primero configura la cuenta que envía (paso 1) y presiona “Probar conexión”. Luego vuelve a intentar.', 'info', { duracion: 9000 });
            return;
        }
        const h = Historial.de(alumno);
        const yaFue = alumno.envio?.estado === 'ok' || (h && h.enviadoEn && String(h.enviadoA || '').trim().toLowerCase() === para);
        const ok = await UI.confirmar('Enviar constancia',
            (yaFue ? `⚠️ Ya se le había enviado su constancia a ${para}${h?.enviadoEn ? ' el ' + new Date(h.enviadoEn).toLocaleDateString('es-MX') : ''}.\n\n` : '') +
            `Se enviará la constancia de ${alumno.nombre} a:
${para}

Desde: ${c.usuario}`, yaFue ? 'Enviar otra vez' : 'Enviar');
        if (!ok) return;
        this._enviandoUno = true;
        UI.toast(`Enviando a ${para}…`, 'info', { duracion: 3000 });
        try {
            const r = await Escritorio.invocar('correo:enviar', c, this.armarMensaje(para, [{ alumno, indice }]));
            if (Alumnos.lista.includes(alumno)) {
                alumno.envio = { estado: r.ok ? 'ok' : 'error', detalle: r.ok ? '' : r.error, fecha: new Date().toISOString() };
            }
            if (r.ok) {
                UI.toast(`Constancia enviada a ${para}.`, 'success', { duracion: 7000 });
                Historial.registrar([{ alumno, indice }], 'enviada', { para });
            }
            else UI.toast('No se pudo enviar:\n' + r.error, 'error', { duracion: 12000 });
        } catch (e) {
            UI.toast('No se pudo enviar:\n' + e.message, 'error');
        } finally {
            this._enviandoUno = false;
            Escritorio.invocar('correo:cerrar').catch(() => { });
            Editor.actualizarResumen();
        }
    },

    // ---------------- Ventana ----------------
    abrir() {
        const aviso = $('email-unavailable');
        if (!Escritorio.disponible) {
            aviso.style.display = 'block';
            aviso.textContent = 'El envío por correo solo funciona en la aplicación de escritorio (npm start o el programa instalado).';
        } else {
            aviso.style.display = 'none';
        }
        this.estadisticas();
        this.vistaPrevia();
        UI.abrirModal('email-modal');
    },

    /**
     * Grupos de envío: cada grupo es un correo con una o varias constancias.
     * Si "un solo correo por alumno" está activo, se juntan todas las
     * constancias (de distintos talleres) que van a la misma dirección.
     */
    grupos() {
        const solo = $('mail-only-pending').checked;
        const juntar = $('mail-group').checked;
        const mapa = new Map();
        const lista = [];
        for (const i of Editor.indicesActivos()) {
            const alumno = Alumnos.lista[i];
            if (!esCorreoValido(alumno.correo)) continue;
            if (solo && alumno.envio?.estado === 'ok') continue;
            // También cuenta lo enviado en días anteriores (historial), aunque se haya cerrado el programa
            if (solo && this.yaEnviado(alumno)) continue;
            const clave = juntar ? alumno.correo.trim().toLowerCase() : 'r' + i;
            if (!mapa.has(clave)) {
                const g = { para: alumno.correo.trim(), items: [] };
                mapa.set(clave, g);
                lista.push(g);
            }
            mapa.get(clave).items.push({ alumno, indice: i });
        }
        return lista;
    },

    /** ¿Ya se le envió esta constancia a este mismo correo (según el historial)? */
    yaEnviado(alumno) {
        const r = Historial.de(alumno);
        return !!(r && r.enviadoEn && String(r.enviadoA || '').trim().toLowerCase() === String(alumno.correo || '').trim().toLowerCase());
    },

    estadisticas() {
        const activos = Editor.indicesActivos().map(i => Alumnos.lista[i]);
        const total = activos.length;
        const conCorreo = activos.filter(a => esCorreoValido(a.correo)).length;
        const invalidos = activos.filter(a => a.correo && !esCorreoValido(a.correo)).length;
        const enviados = activos.filter(a => a.envio?.estado === 'ok').length;
        const fallidos = activos.filter(a => a.envio?.estado === 'error').length;
        const grupos = this.grupos();
        const constancias = grupos.reduce((n, g) => n + g.items.length, 0);
        let t = `${total} constancias${Editor.filtroEvento !== '__todos__' ? ' del evento elegido' : ''} · ${conCorreo} con correo válido`;
        if (total - conCorreo - invalidos > 0) t += ` · ${total - conCorreo - invalidos} sin correo`;
        if (invalidos) t += ` · ⚠️ ${invalidos} con correo mal escrito`;
        if (enviados || fallidos) t += ` · ✅ ${enviados} enviadas · ❌ ${fallidos} con error`;
        $('mail-stats').textContent = t;
        $('mail-send-btn').textContent = grupos.length
            ? `🚀 Enviar ${grupos.length} correo${grupos.length === 1 ? '' : 's'}` +
            (constancias !== grupos.length ? ` (${constancias} constancias)` : '')
            : '🚀 Enviar';
        $('mail-send-btn').disabled = grupos.length === 0 || this.enviando;
    },

    /** Reemplaza {nombre}, {codigo}, {evento}, {folio}, {fecha} y cualquier {Columna} del CSV */
    rellenar(plantilla, items) {
        const unicos = (vals) => [...new Set(vals.filter(Boolean))];
        const lista = (vals) => {
            const u = unicos(vals);
            return u.length <= 1 ? (u[0] || '') : u.slice(0, -1).join(', ') + ' y ' + u[u.length - 1];
        };
        const curso = Editor.campo('curso');
        const fecha = Editor.campo('fecha');
        const primero = items[0]?.alumno || null;
        const datos = {
            nombre: primero ? primero.nombre : 'Nombre del Alumno',
            codigo: primero ? (primero.codigo || '') : '219000000',
            correo: primero ? (primero.correo || '') : 'alumno@correo.com',
            folio: lista(items.map(it => Editor.folioDe(it.alumno, it.indice))) || Editor.folioDe(null, 0),
            evento: lista(items.map(it => it.alumno?.evento)) || (curso ? curso.valor : ''),
            fecha: lista(items.map(it => it.alumno?.fecha)) || (fecha ? fecha.valor : fechaLarga())
        };
        datos.curso = datos.evento;
        datos.taller = datos.evento;
        return String(plantilla || '').replace(/\{([^{}]+)\}/g, (todo, k) => {
            const clave = normalizarEncabezado(k);
            if (Object.prototype.hasOwnProperty.call(datos, clave)) return datos[clave];
            // Otras columnas del CSV
            const col = Alumnos.columnasExtra.find(c => normalizarEncabezado(c) === clave);
            if (col) return lista(items.map(it => (it.alumno?.extra || {})[col]));
            return todo;
        });
    },

    itemsVistaPrevia() {
        const a = Editor.alumnoActual() || Alumnos.lista[Editor.indicesActivos()[0]] || null;
        if (!a) return [{ alumno: null, indice: 0 }];
        const i = Alumnos.lista.indexOf(a);
        // Constancia escrita a mano (no está en la lista): solo ella
        if (Editor.manual || i < 0) return [{ alumno: a, indice: 0 }];
        if (!$('mail-group').checked || !a.correo) return [{ alumno: a, indice: i }];
        // Mismo correo → todas sus constancias del evento elegido
        return Editor.indicesActivos()
            .filter(j => (Alumnos.lista[j].correo || '').trim().toLowerCase() === a.correo.trim().toLowerCase())
            .map(j => ({ alumno: Alumnos.lista[j], indice: j }));
    },

    vistaPrevia() {
        const items = this.itemsVistaPrevia();
        let t = `Asunto: ${this.rellenar($('mail-subject').value, items)}\n\n${this.rellenar($('mail-body').value, items)}`;
        if (items.length > 1) t += `\n\n📎 ${items.length} constancias adjuntas`;
        $('mail-preview').textContent = t;
    },

    armarMensaje(para, items) {
        const texto = this.rellenar($('mail-body').value, items);
        const html = '<div style="font-family:Arial,Helvetica,sans-serif;font-size:14px;line-height:1.5">' +
            escapeHtml(texto).replace(/\n/g, '<br>') + '</div>';
        const usados = new Set();
        const adjuntos = items.map(({ alumno, indice }) => {
            const doc = Editor.construirPdf(alumno, Math.max(0, indice));
            let base = Editor.nombreDeArchivo(alumno, Math.max(0, indice));
            // Varias constancias del mismo alumno: se distinguen por evento
            if (items.length > 1 && alumno && alumno.evento && !base.includes(nombreArchivoSeguro(alumno.evento))) {
                base += '_' + nombreArchivoSeguro(alumno.evento).substring(0, 40);
            }
            let nombre = base, k = 2;
            while (usados.has(nombre.toLowerCase())) nombre = `${base}_${k++}`;
            usados.add(nombre.toLowerCase());
            return { nombre: nombre + '.pdf', datos: new Uint8Array(doc.output('arraybuffer')) };
        });
        return { para, asunto: this.rellenar($('mail-subject').value, items), texto, html, adjuntos };
    },

    // ---------------- Acciones ----------------
    async probarConexion() {
        const cfg = this.validarCuenta();
        if (!cfg) return;
        const btn = $('smtp-test-btn');
        btn.disabled = true;
        btn.textContent = '⏳ Conectando…';
        try {
            const r = await Escritorio.invocar('correo:probar', cfg);
            if (r.ok) {
                UI.toast('Conexión correcta. Ya puedes enviar.', 'success');
                this.guardarConfig();
            } else {
                UI.toast('No se pudo conectar:\n' + r.error, 'error', { duracion: 12000 });
            }
        } catch (e) {
            UI.toast(e.message, 'error');
        } finally {
            btn.disabled = false;
            btn.textContent = '🔌 Probar conexión';
        }
    },

    async enviarPrueba() {
        const cfg = this.validarCuenta();
        if (!cfg || !Editor.validarAntesDeGenerar()) return;
        const items = this.itemsVistaPrevia();
        const btn = $('mail-test-btn');
        btn.disabled = true;
        btn.textContent = '⏳ Enviando…';
        try {
            const msg = this.armarMensaje(cfg.usuario, items);
            msg.asunto = '[PRUEBA] ' + msg.asunto;
            const r = await Escritorio.invocar('correo:enviar', cfg, msg);
            if (r.ok) UI.toast(`Prueba enviada a ${cfg.usuario} con ${msg.adjuntos.length} constancia(s). Revisa tu bandeja (y la carpeta de spam).`, 'success', { duracion: 8000 });
            else UI.toast('No se pudo enviar la prueba:\n' + r.error, 'error', { duracion: 12000 });
        } catch (e) {
            UI.toast(e.message, 'error');
        } finally {
            btn.disabled = false;
            btn.textContent = '🧪 Enviarme una prueba';
        }
    },

    async enviarTodos() {
        const cfg = this.validarCuenta();
        if (!cfg || !Editor.validarAntesDeGenerar()) return;
        const grupos = this.grupos();
        if (!grupos.length) return UI.toast('No hay alumnos con correo pendientes de envío.', 'warning');

        const constancias = grupos.reduce((n, g) => n + g.items.length, 0);
        const sinCorreo = Editor.indicesActivos().filter(i => !esCorreoValido(Alumnos.lista[i].correo)).length;
        let aviso = `Se enviarán ${grupos.length} correo(s) con ${constancias} constancia(s) desde ${cfg.usuario}.`;
        if (sinCorreo) aviso += `\n\n${sinCorreo} registro(s) no tienen un correo válido y se omitirán.`;
        if (cfg.proveedor === 'gmail' && grupos.length > 450) aviso += '\n\n⚠️ Gmail permite unos 500 correos al día; podrías alcanzar el límite.';
        if (!await UI.confirmar('Enviar por correo', aviso, 'Enviar')) return;

        this.enviando = true;
        this.cancelar = false;
        const pausa = Math.max(0, parseFloat($('mail-delay').value) || 0) * 1000;
        let ok = 0, err = 0, seguidosSinRed = 0, detenido = '';

        $('progress-title').textContent = 'Enviando correos…';
        UI.abrirModal('progress-modal');
        const pintar = (n, texto) => {
            const pct = Math.round((n / grupos.length) * 100);
            $('progress-fill').style.width = pct + '%';
            $('progress-pct').textContent = pct + '%';
            $('progress-info').textContent = texto;
        };
        $('progress-cancel').onclick = () => { this.cancelar = true; $('progress-info').textContent = 'Cancelando…'; };

        try {
            for (let k = 0; k < grupos.length; k++) {
                if (this.cancelar) break;
                const g = grupos[k];
                const quien = g.items[0].alumno.nombre;
                pintar(k, `${k + 1} de ${grupos.length}: ${quien} <${g.para}>` + (g.items.length > 1 ? ` · ${g.items.length} constancias` : ''));
                let r;
                try {
                    r = await Escritorio.invocar('correo:enviar', cfg, this.armarMensaje(g.para, g.items));
                } catch (e) {
                    r = { ok: false, error: e.message };
                }
                const envio = { estado: r.ok ? 'ok' : 'error', detalle: r.ok ? '' : r.error, fecha: new Date().toISOString() };
                g.items.forEach(it => { it.alumno.envio = { ...envio }; });
                if (r.ok) { ok++; Historial.registrar(g.items, 'enviada', { para: g.para }); } else err++;
                // Si falla el inicio de sesión, no tiene caso seguir
                if (!r.ok && r.tipo === 'auth') {
                    UI.toast('Se detuvo el envío: ' + r.error, 'error', { duracion: 12000 });
                    break;
                }
                // Límite diario del correo o sin internet: se detiene y se puede retomar después
                seguidosSinRed = (!r.ok && r.tipo === 'conexion') ? seguidosSinRed + 1 : 0;
                if (!r.ok && r.tipo === 'limite') { detenido = 'limite'; break; }
                if (seguidosSinRed >= 3) { detenido = 'conexion'; break; }
                if (k < grupos.length - 1 && pausa) await new Promise(res => setTimeout(res, pausa));
            }
        } finally {
            this.enviando = false;
            $('progress-cancel').onclick = null;
            UI.cerrarModal('progress-modal');
            $('progress-title').textContent = 'Generando constancias…';
            Escritorio.invocar('correo:cerrar').catch(() => { });
        }

        const tipo = err ? (ok ? 'warning' : 'error') : 'success';
        let msg = `${ok} correo(s) enviados`;
        if (err) msg += `, ${err} con error (márcalos en la lista con ✖; puedes reintentar solo esos)`;
        if (this.cancelar) msg += '. Envío cancelado.';
        if (detenido) {
            const faltan = this.grupos().length;
            const porque = detenido === 'limite'
                ? 'El correo llegó a su límite de envíos (Gmail permite unos 500 al día).'
                : 'Se perdió la conexión a internet.';
            UI.confirmar('Envío en pausa',
                `${porque}\n\nSe enviaron ${ok}. Faltan ${faltan} correo(s).\n\nPara continuar ${detenido === 'limite' ? 'mañana' : 'cuando regrese el internet'}: ` +
                'abre el programa, carga la misma lista y presiona «Enviar» con la casilla «Solo pendientes» marcada. ' +
                'Seguirá con los que faltan y no repetirá a quien ya lo recibió (aunque hayas cerrado el programa).',
                'Entendido');
        }
        UI.toast(msg, tipo, {
            duracion: 10000,
            accion: err ? { texto: 'Reporte', fn: () => this.descargarReporte() } : undefined
        });
        this.estadisticas();
        Editor.actualizarResumen();
    },

    descargarReporte() {
        if (!Alumnos.lista.length) return UI.toast('La lista está vacía.', 'warning');
        const filas = Alumnos.lista.map(a => ({
            Codigo: a.codigo, Nombre: a.nombre, Correo: a.correo || '', Evento: a.evento || '',
            Folio: Editor.folioDe(a, Alumnos.lista.indexOf(a)),
            Estado: a.envio ? (a.envio.estado === 'ok' ? 'Enviado' : 'Error')
                : this.yaEnviado(a) ? 'Enviado (antes)' : (esCorreoValido(a.correo) ? 'Pendiente' : 'Sin correo'),
            Detalle: a.envio?.detalle || '',
            Fecha: a.envio?.fecha ? new Date(a.envio.fecha).toLocaleString('es-MX')
                : (this.yaEnviado(a) ? new Date(Historial.de(a).enviadoEn).toLocaleString('es-MX') : '')
        }));
        descargarBlob(new Blob(['\uFEFF' + Papa.unparse(filas.map(filaSegura))], { type: 'text/csv;charset=utf-8' }), 'reporte_envios.csv');
    }
};

// ==================================================================================
// 8. ARRANQUE
// ==================================================================================
// Soltar un archivo sobre la ventana NO debe abrirlo como página (seguridad)
['dragover', 'drop'].forEach(ev => window.addEventListener(ev, e => e.preventDefault()));

document.addEventListener('DOMContentLoaded', async () => {
    try {
        await Fuentes.init();
    } catch (e) {
        console.error('Error al preparar las fuentes', e);
    }
    Fuentes.llenarSelect($('prop-font'));
    await Historial.init();
    Proyectos.init();
    Editor.init();
    Actualizaciones.init();
    Proyectos.cargarNombre();
    // Buscador de acciones y personas: Ctrl+F (Ctrl+K ya es «cursiva», como en Word)
    document.addEventListener('keydown', e => {
        if ((e.ctrlKey || e.metaKey) && !e.altKey && e.key.toLowerCase() === 'f') {
            if (document.querySelector('.modal-backdrop.open:not(#cmd-modal)')) return;
            e.preventDefault();
            Comandos.abrir();
        }
    });
    // Cuando terminen de cargar las fuentes de la vista previa, volver a dibujar
    if (document.fonts) document.fonts.ready.then(() => Editor.renderCampos());
});
