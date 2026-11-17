// ================================================================
// MÓDULO PRINCIPAL DE ELECTRON - EL "CEREBRO" DE LA APLICACIÓN
// ================================================================
// Este archivo es el punto de entrada de la aplicación: crea la ventana
// y atiende las peticiones que la interfaz no puede hacer por sí sola
// (enviar correos, guardar la contraseña cifrada, ubicar la carpeta de datos).

const { app, BrowserWindow, ipcMain, safeStorage, shell, dialog } = require('electron');
const path = require('path');
const fs = require('fs');
const os = require('os');

// ================================================================
// VENTANA PRINCIPAL
// ================================================================
function createWindow() {
    const win = new BrowserWindow({
        width: 1280,
        height: 720,
        minWidth: 1280,
        minHeight: 720,
        autoHideMenuBar: true,
        webPreferences: {
            nodeIntegration: true,      // Permite usar Node.js en la interfaz
            contextIsolation: false     // Necesario para que funcionen las librerías
        },
        // El ícono solo se usa si existe (la carpeta assets no viene en el proyecto)
        icon: fs.existsSync(path.join(__dirname, 'assets/icon.ico')) ? path.join(__dirname, 'assets/icon.ico') : undefined
    });

    // Los enlaces externos (ej. ayuda de Gmail) se abren en el navegador
    win.webContents.setWindowOpenHandler(({ url }) => {
        if (/^https?:\/\//i.test(url)) shell.openExternal(url);
        return { action: 'deny' };
    });

    // SEGURIDAD: esta ventana tiene acceso a Node.js, así que solo puede mostrar
    // las páginas del propio programa. Cualquier otra navegación (un enlace, un
    // archivo soltado sobre la ventana, etc.) se bloquea; los enlaces web se abren
    // en el navegador del sistema.
    const paginasPermitidas = ['index.html', 'index_express.html']
        .map(f => require('url').pathToFileURL(path.join(__dirname, f)).href.toLowerCase());
    win.webContents.on('will-navigate', (e, url) => {
        const limpia = url.split('#')[0].split('?')[0].toLowerCase();
        if (!paginasPermitidas.includes(limpia)) {
            e.preventDefault();
            if (/^https?:\/\//i.test(url)) shell.openExternal(url);
        }
    });
    win.webContents.on('will-attach-webview', e => e.preventDefault());

    win.maximize();
    win.loadFile(path.join(__dirname, 'index.html'));
}

// ================================================================
// CORREO ELECTRÓNICO
// ================================================================
// El módulo se carga solo cuando se usa, así el programa abre aunque
// falte instalar "nodemailer".
let correo = null;
function moduloCorreo() {
    if (!correo) {
        try {
            correo = require('./correo');
        } catch (e) {
            throw new Error('Falta la librería "nodemailer". En la carpeta del proyecto ejecuta: npm install');
        }
    }
    return correo;
}

function respuestaError(err) {
    try {
        return { ok: false, ...moduloCorreo().traducirError(err) };
    } catch (e) {
        return { ok: false, tipo: 'otro', error: err.message };
    }
}

ipcMain.handle('correo:probar', async (_e, cfg) => {
    try {
        await moduloCorreo().probar(cfg);
        return { ok: true };
    } catch (err) {
        return respuestaError(err);
    }
});

ipcMain.handle('correo:enviar', async (_e, cfg, mensaje) => {
    try {
        const r = await moduloCorreo().enviar(cfg, mensaje);
        return { ok: true, ...r };
    } catch (err) {
        return respuestaError(err);
    }
});

ipcMain.handle('correo:cerrar', () => {
    if (correo) correo.cerrar();
    return true;
});

// ----------------------------------------------------------------
// Configuración del correo (la contraseña se guarda CIFRADA con
// safeStorage, que en Windows usa la protección de la cuenta de usuario)
// ----------------------------------------------------------------
const rutaConfigCorreo = () => path.join(app.getPath('userData'), 'correo.json');

ipcMain.handle('correo:leerConfig', () => {
    try {
        const cfg = JSON.parse(fs.readFileSync(rutaConfigCorreo(), 'utf8'));
        if (cfg.contrasenaCifrada && safeStorage.isEncryptionAvailable()) {
            try {
                cfg.contrasena = safeStorage.decryptString(Buffer.from(cfg.contrasenaCifrada, 'base64'));
            } catch (e) {
                cfg.contrasena = '';
            }
        }
        delete cfg.contrasenaCifrada;
        return cfg;
    } catch (e) {
        return null; // Aún no hay configuración
    }
});

ipcMain.handle('correo:guardarConfig', (_e, cfg) => {
    const datos = { ...cfg };
    delete datos.contrasena;
    if (cfg.recordar && cfg.contrasena && safeStorage.isEncryptionAvailable()) {
        datos.contrasenaCifrada = safeStorage.encryptString(cfg.contrasena).toString('base64');
    }
    fs.mkdirSync(app.getPath('userData'), { recursive: true });
    fs.writeFileSync(rutaConfigCorreo(), JSON.stringify(datos, null, 2));
    return true;
});

// Carpeta de datos del programa (fuentes importadas, configuración)
ipcMain.handle('app:rutaDatos', () => app.getPath('userData'));

// Versión instalada y repositorio de GitHub para avisar de versiones nuevas
ipcMain.handle('app:info', () => {
    let repo = '';
    try {
        const pkg = JSON.parse(fs.readFileSync(path.join(__dirname, 'package.json'), 'utf8'));
        repo = (pkg.actualizaciones && pkg.actualizaciones.github) || '';
    } catch (e) { /* sin package.json */ }
    return { version: app.getVersion(), repo };
});

// ================================================================
// GUARDAR EN CARPETA E IMPRIMIR
// ================================================================
// Ventana para elegir la carpeta donde se guardan las constancias
ipcMain.handle('dialogo:carpeta', async (e) => {
    const win = BrowserWindow.fromWebContents(e.sender);
    const r = await dialog.showOpenDialog(win, {
        title: 'Elige dónde guardar las constancias',
        properties: ['openDirectory', 'createDirectory']
    });
    return r.canceled ? null : r.filePaths[0];
});

// Abrir la carpeta en el Explorador de Windows
ipcMain.handle('app:abrirCarpeta', (_e, ruta) => {
    // Solo carpetas que existen (nunca ejecutar un archivo)
    try {
        if (typeof ruta !== 'string' || !fs.statSync(ruta).isDirectory()) return 'No es una carpeta';
    } catch (e) { return 'No existe'; }
    return shell.openPath(ruta);
});

// Imprimir: el PDF (ya trae la orden de imprimir) se abre en una ventana con el visor de PDF.
// Si el visor no está disponible, se abre con el programa de PDF de la computadora.
ipcMain.handle('imprimir:pdf', async (_e, datos) => {
    if (!(datos instanceof Uint8Array) && !Buffer.isBuffer(datos) && !(datos instanceof ArrayBuffer)) throw new Error('Datos de impresión no válidos');
    const ruta = path.join(os.tmpdir(), `constancias_imprimir_${Date.now()}.pdf`);
    fs.writeFileSync(ruta, Buffer.from(datos));
    const win = new BrowserWindow({
        width: 1000,
        height: 800,
        title: 'Imprimir constancias',
        autoHideMenuBar: true,
        webPreferences: { plugins: true, nodeIntegration: false, contextIsolation: true, sandbox: true }
    });
    win.webContents.on('will-navigate', e => e.preventDefault());
    win.webContents.setWindowOpenHandler(() => ({ action: 'deny' }));
    // El PDF temporal tiene datos de los alumnos: se borra al cerrar la ventana de impresión
    let externo = false;
    win.on('closed', () => { if (!externo) setTimeout(() => { try { fs.unlinkSync(ruta); } catch (e) { } }, 2000); });
    try {
        await win.loadFile(ruta);
        return { ok: true };
    } catch (err) {
        externo = true; // lo abre el visor de la computadora: el archivo debe seguir ahí
        win.close();
        await shell.openPath(ruta);
        return { ok: true, externo: true };
    }
});

// ================================================================
// CICLO DE VIDA
// ================================================================
// Una sola ventana del programa a la vez (dos ventanas pueden pelearse el historial y los proyectos)
if (!app.requestSingleInstanceLock()) {
    app.quit();
} else {
    app.on('second-instance', () => {
        const w = BrowserWindow.getAllWindows()[0];
        if (w) { if (w.isMinimized()) w.restore(); w.focus(); }
    });
    app.whenReady().then(createWindow);
}

app.on('window-all-closed', () => {
    if (correo) correo.cerrar();
    if (process.platform !== 'darwin') {
        app.quit();
    }
});

app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
        createWindow();
    }
});
