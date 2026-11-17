// ================================================================
// MÓDULO DE CORREO (se ejecuta en el proceso principal de Electron)
// ================================================================
// Usa la librería "nodemailer" para conectarse a un servidor SMTP
// (Gmail, Outlook u otro) y enviar cada constancia como adjunto.
// La interfaz (renderer.js) le pide los envíos a main.js por IPC.

const nodemailer = require('nodemailer');

// Datos de conexión de los proveedores más comunes
const PROVEEDORES = {
    gmail: { host: 'smtp.gmail.com', port: 465, secure: true },
    outlook: { host: 'smtp.office365.com', port: 587, secure: false, requireTLS: true }
};

// Se reutiliza la misma conexión mientras no cambien los datos de la cuenta
let actual = { clave: null, transporte: null };

function opcionesServidor(cfg) {
    if (cfg.proveedor && PROVEEDORES[cfg.proveedor]) return { ...PROVEEDORES[cfg.proveedor] };
    const seguridad = cfg.seguridad || 'starttls';
    return {
        host: cfg.host,
        port: parseInt(cfg.puerto, 10) || (seguridad === 'ssl' ? 465 : 587),
        secure: seguridad === 'ssl',
        requireTLS: seguridad === 'starttls',
        ignoreTLS: seguridad === 'ninguna'
    };
}

function transporte(cfg) {
    const servidor = opcionesServidor(cfg);
    const clave = JSON.stringify([servidor, cfg.usuario, cfg.contrasena]);
    if (actual.clave !== clave) {
        cerrar();
        actual = {
            clave,
            transporte: nodemailer.createTransport({
                ...servidor,
                auth: { user: cfg.usuario, pass: cfg.contrasena },
                pool: true,             // una sola conexión para muchos correos
                maxConnections: 1,
                maxMessages: 100,
                connectionTimeout: 20000,
                greetingTimeout: 20000,
                socketTimeout: 60000
            })
        };
    }
    return actual.transporte;
}

/** Comprueba usuario, contraseña y servidor */
async function probar(cfg) {
    await transporte(cfg).verify();
    return true;
}

/**
 * Envía un correo con una o varias constancias adjuntas.
 * @param cfg  { proveedor, host, puerto, seguridad, usuario, contrasena, nombre }
 * @param m    { para, asunto, texto, html, adjuntos: [{ nombre, datos(Uint8Array) }] }
 */
async function enviar(cfg, m) {
    const info = await transporte(cfg).sendMail({
        from: { name: cfg.nombre || cfg.usuario, address: cfg.usuario },
        to: m.para,
        subject: m.asunto,
        text: m.texto,
        html: m.html,
        // Una o varias constancias (si el alumno participó en varios talleres)
        attachments: (m.adjuntos || (m.adjunto ? [m.adjunto] : [])).map(a => ({
            filename: a.nombre,
            content: Buffer.from(a.datos),
            contentType: 'application/pdf'
        }))
    });
    if (info.rejected && info.rejected.length) {
        const e = new Error('El servidor rechazó la dirección ' + info.rejected.join(', '));
        e.code = 'EENVELOPE';
        throw e;
    }
    return { id: info.messageId };
}

function cerrar() {
    if (actual.transporte) {
        try { actual.transporte.close(); } catch (e) { /* nada */ }
    }
    actual = { clave: null, transporte: null };
}

/** Convierte los errores técnicos en mensajes entendibles */
function traducirError(err) {
    const msg = String((err && err.message) || err || 'Error desconocido');
    const code = err && err.code;
    if (code === 'EAUTH' || /auth|credential|password|username/i.test(msg) && /535|534|invalid|failed|not accepted/i.test(msg)) {
        return { tipo: 'auth', error: 'Correo o contraseña incorrectos. En Gmail necesitas una “contraseña de aplicación” (no tu contraseña normal). Detalle: ' + msg };
    }
    if (code === 'ETIMEDOUT' || code === 'ECONNECTION' || code === 'ESOCKET' || code === 'ECONNREFUSED' || code === 'ENOTFOUND' || code === 'EDNS') {
        return { tipo: 'conexion', error: 'No se pudo conectar al servidor de correo. Revisa tu internet, el servidor y el puerto. Detalle: ' + msg };
    }
    // Primero el límite: Gmail responde "550 5.4.5 Daily user sending limit exceeded" (también trae 550)
    if (/5\.4\.5|daily.*(limit|quota)|sending limit|quota|rate limit|too many|421|454|4\.7\.0/i.test(msg)) {
        return { tipo: 'limite', error: 'El servidor limitó los envíos (cuota diaria o demasiados seguidos). Espera un rato o aumenta la pausa. Detalle: ' + msg };
    }
    if (code === 'EENVELOPE' || /recipient|550|553|501/i.test(msg)) {
        return { tipo: 'destinatario', error: 'La dirección del destinatario fue rechazada. Detalle: ' + msg };
    }
    return { tipo: 'otro', error: msg };
}

module.exports = { probar, enviar, cerrar, traducirError, PROVEEDORES };
