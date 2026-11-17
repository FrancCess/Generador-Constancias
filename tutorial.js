/**
 * ==================================================================================
 * TUTORIAL DE PRIMER USO (recorrido guiado)
 * ==================================================================================
 * - Se muestra solo la primera vez que se abre el programa.
 * - Se puede volver a ver cuando quieras con el botón "Tutorial" (arriba) o la tecla F1.
 * - Resalta cada parte de la pantalla y explica con palabras sencillas qué hace.
 * - Funciona con las dos cintas (con pestañas y única): si el botón está en otra
 *   pestaña, la abre sola.
 * ==================================================================================
 */
(() => {
    // Cada pantalla puede traer su propio recorrido (el Modo Express define el suyo)
    const CLAVE_VISTO = window.TUTORIAL_CLAVE || 'gc_tutorial_visto';

    // ------------------------------------------------------------------
    // PASOS DEL RECORRIDO
    // objetivo: id del elemento a resaltar (o función que lo regresa)
    // pestana: pestaña que se abre antes (solo en la cinta con pestañas)
    // ------------------------------------------------------------------
    const padre = (id, sel) => () => {
        const el = document.getElementById(id);
        return el ? (sel ? el.closest(sel) : el.parentElement) : null;
    };

    const PASOS_AVANZADO = [
        {
            titulo: '¡Bienvenido al Generador de Constancias! 🎓',
            texto: 'Con este programa haces constancias en PDF para muchas personas al mismo tiempo, ' +
                'o solo una si lo necesitas.<br><br>Te mostramos en <b>2 minutos</b> dónde está cada cosa. ' +
                'Usa el botón <b>Siguiente</b> para avanzar.'
        },
        {
            objetivo: padre('image-upload', 'label'),
            pestana: 'diseno',
            titulo: 'Paso 1 · Elige el diseño',
            texto: 'Primero sube la <b>imagen de tu constancia</b> (el fondo, con logos y firmas). ' +
                'Puede ser JPG o PNG.<br><br>💡 Puedes diseñarla en Canva, Word o PowerPoint y guardarla como imagen.'
        },
        {
            objetivo: padre('page-size', '.group'),
            pestana: 'diseno',
            titulo: 'Tamaño de la hoja',
            texto: 'Aquí eliges el tamaño (Carta, Oficio o A4) y si la hoja va <b>horizontal</b> (acostada) ' +
                'o <b>vertical</b> (parada).'
        },
        {
            objetivo: padre('csv-upload', 'label'),
            pestana: 'inicio',
            titulo: 'Paso 2 · Carga tu lista',
            texto: 'Carga la lista de personas desde un archivo de <b>Excel</b> o <b>CSV</b>. ' +
                'Basta con una columna llamada <b>Nombre</b>; también puede traer Código, Correo, Taller, Folio, Fecha, Carrera o Grupo.<br><br>' +
                'Después de cargarla, el programa te muestra la lista para que <b>corrijas errores</b>.'
        },
        {
            objetivo: () => document.querySelector('.paste-btn'),
            pestana: 'inicio',
            titulo: 'O pega tu tabla',
            texto: '¿Tienes la lista abierta en Excel o Google Sheets? Cópiala (<b>Ctrl+C</b>), presiona <b>Pegar tabla</b> ' +
                'y pégala (<b>Ctrl+V</b>). No necesitas guardar ningún archivo.<br><br>' +
                '❓ Pasa el mouse sobre el <b>signo de interrogación</b> amarillo para ver <b>cómo debe ser tu tabla</b>: ' +
                'te muestra las columnas según los textos que lleva tu constancia.'
        },
        {
            objetivo: () => document.getElementById('manual-btn'),
            pestana: 'inicio',
            titulo: '¿Solo necesitas una?',
            texto: 'Si solo quieres <b>una constancia</b>, presiona aquí y escribe el nombre a mano. ' +
                'No necesitas ninguna lista.'
        },
        {
            objetivo: () => document.getElementById('certificate-canvas'),
            titulo: 'Paso 3 · Acomoda los textos',
            texto: 'Aquí ves cómo quedará la constancia.<br><br>🖱️ <b>Arrastra cada texto con el mouse</b> ' +
                'para ponerlo en su lugar (por ejemplo, el nombre sobre la línea).'
        },
        {
            objetivo: () => document.querySelector('.pane'),
            titulo: '¿Qué textos lleva?',
            texto: 'En esta lista <b>marca con una palomita ✔</b> los textos que quieres en la constancia: ' +
                'nombre, código, taller, fecha, folio…<br><br>Haz clic en uno para cambiar lo que dice.'
        },
        {
            objetivo: () => {
                const f = document.getElementById('prop-font');
                return f ? (f.closest('.group') || f) : null;
            },
            pestana: 'texto',
            titulo: 'Letra, tamaño y color',
            texto: 'Con el texto seleccionado, cambia aquí el <b>tipo de letra</b>, el <b>tamaño</b>, ' +
                'el <b>color</b> y pon <b>negritas</b>, <i>cursiva</i> o subrayado, igual que en Word.'
        },
        {
            objetivo: () => document.querySelector('.statusbar'),
            titulo: 'Revisa antes de generar',
            texto: 'Con las flechas <b>◀ ▶</b> de esta barra ves la constancia de <b>cada persona</b>. ' +
                'Revisa que los nombres se vean bien antes de generar.'
        },
        {
            objetivo: () => document.getElementById('generate-pdf-btn'),
            pestana: 'inicio',
            titulo: 'Paso 4 · Genera las constancias',
            texto: 'Cuando todo esté listo, presiona <b>Generar</b>. Podrás elegir:<br>' +
                '• un PDF por persona o <b>uno solo con todas</b>,<br>' +
                '• guardarlas en <b>una carpeta</b> o en un ZIP,<br>' +
                '• separarlas por <b>taller, grupo o carrera</b>.'
        },
        {
            objetivo: () => document.getElementById('print-btn'),
            pestana: 'inicio',
            titulo: 'Imprimir',
            texto: 'Para imprimir <b>una o varias</b> constancias directo desde el programa.'
        },
        {
            objetivo: () => document.getElementById('open-email-btn'),
            pestana: 'correo',
            titulo: 'Enviar por correo',
            texto: 'Si tu lista trae correos, aquí envías a cada persona <b>su constancia en PDF</b>. ' +
                'La primera vez te pedirá tu cuenta de Gmail y una “contraseña de aplicación” (te explica cómo crearla).'
        },
        {
            objetivo: () => document.getElementById('save-template-btn'),
            titulo: 'Guarda tu diseño',
            texto: 'Con este botón 💾 guardas el diseño (imagen, textos y posiciones) para usarlo otra vez ' +
                'en el próximo evento. Con 📂 lo abres.'
        },
        {
            objetivo: () => document.getElementById('mesa-btn'),
            titulo: 'Mesa de trabajo',
            texto: 'Si haces constancias seguido, prueba la <b>Mesa de trabajo</b>: la hoja, la lista y lo que falta, todo en una sola pantalla.<br><br>' +
                '🔎 En cualquier pantalla, <b>Ctrl+F</b> busca una acción («imprimir», «correo», «QR») o a una persona.'
        },
        {
            objetivo: () => document.getElementById('express-btn'),
            titulo: '¿Tienes prisa? Modo Express',
            texto: 'Si solo necesitas algo rápido y sencillo, presiona <b>Modo Express</b>: ' +
                'las mismas constancias en <b>4 pasos</b>, sin opciones de más.<br><br>' +
                'Este modo completo (avanzado) siempre seguirá aquí.'
        },
        {
            objetivo: () => document.getElementById('help-btn'),
            titulo: '¡Listo! 🎉',
            texto: 'Si se te olvida algo, presiona <b>Tutorial</b> (o la tecla <b>F1</b>) para ver esta guía ' +
                'otra vez, cuantas veces quieras.<br><br>Pasos básicos: <b>1</b> imagen → <b>2</b> lista → ' +
                '<b>3</b> acomoda textos → <b>4</b> Generar.'
        }
    ];

    const PASOS = window.TUTORIAL_PASOS || PASOS_AVANZADO;

    // ------------------------------------------------------------------
    // ESTILOS
    // ------------------------------------------------------------------
    const css = `
    .tuto-bloqueo { position: fixed; inset: 0; z-index: 12500; background: transparent; }
    .tuto-bloqueo.sin-objetivo { background: rgba(15, 14, 30, 0.62); }
    .tuto-foco {
        position: fixed; z-index: 12501; pointer-events: none; border-radius: 8px;
        box-shadow: 0 0 0 4px #b8913f, 0 0 0 9999px rgba(15, 14, 30, 0.62);
        transition: all 0.25s ease;
    }
    .tuto-globo {
        position: fixed; z-index: 12502; width: min(380px, calc(100vw - 24px));
        background: #fff; color: #252423; border-radius: 12px;
        box-shadow: 0 18px 50px rgba(0, 0, 0, 0.35); padding: 18px 20px 14px;
        font-family: 'Segoe UI', system-ui, sans-serif; font-size: 15px; line-height: 1.5;
        transition: top 0.25s ease, left 0.25s ease;
    }
    .tuto-globo.centro { width: min(460px, calc(100vw - 24px)); }
    .tuto-globo h3 { font-size: 18px; margin: 0 0 8px; color: #1d1b33; }
    .tuto-globo .tuto-texto { margin-bottom: 14px; }
    .tuto-pie { display: flex; align-items: center; gap: 8px; }
    .tuto-pie .tuto-paso { flex: 1; min-width: 0; font-size: 12.5px; color: #605e5c; }
    .tuto-puntos { display: flex; flex-wrap: wrap; gap: 4px; margin-top: 4px; }
    .tuto-puntos span { width: 7px; height: 7px; border-radius: 50%; background: #d8d8dd; }
    .tuto-puntos span.on { background: #b8913f; }
    .tuto-btn {
        height: 36px; padding: 0 16px; border-radius: 8px; border: 1px solid #c4c4cc;
        background: #fff; cursor: pointer; font-size: 14px; font-family: inherit;
        white-space: nowrap; flex: 0 0 auto;
    }
    .tuto-btn:hover { background: #f0f0f3; }
    .tuto-btn.primario { background: #1d1b33; color: #fff; border-color: #1d1b33; font-weight: 600; }
    .tuto-btn.primario:hover { background: #2a2747; }
    .tuto-cerrar {
        position: absolute; top: 8px; right: 10px; border: none; background: none;
        font-size: 18px; cursor: pointer; color: #8a8886; width: 30px; height: 30px; border-radius: 6px;
    }
    .tuto-cerrar:hover { background: #f0f0f3; }
    .tuto-flecha {
        position: absolute; width: 14px; height: 14px; background: #fff; transform: rotate(45deg);
    }
    `;

    // ------------------------------------------------------------------
    // LÓGICA
    // ------------------------------------------------------------------
    let paso = 0;
    let capas = null;

    const crearCapas = () => {
        if (capas) return capas;
        const estilo = document.createElement('style');
        estilo.textContent = css;
        document.head.appendChild(estilo);
        const bloqueo = document.createElement('div');
        bloqueo.className = 'tuto-bloqueo';
        const foco = document.createElement('div');
        foco.className = 'tuto-foco';
        const globo = document.createElement('div');
        globo.className = 'tuto-globo';
        globo.setAttribute('role', 'dialog');
        globo.setAttribute('aria-live', 'polite');
        document.body.append(bloqueo, foco, globo);
        capas = { bloqueo, foco, globo };
        return capas;
    };

    const quitarCapas = () => {
        if (!capas) return;
        capas.bloqueo.remove();
        capas.foco.remove();
        capas.globo.remove();
        capas = null;
    };

    const visible = (el) => {
        if (!el) return false;
        const r = el.getBoundingClientRect();
        return r.width > 0 && r.height > 0;
    };

    const mostrar = (n) => {
        paso = Math.max(0, Math.min(PASOS.length - 1, n));
        const p = PASOS[paso];
        // Cerrar ventanas abiertas para que no tapen
        document.querySelectorAll('.modal-backdrop.open').forEach(m => m.classList.remove('open'));
        if (p.pestana && window.gcAbrirPestana) window.gcAbrirPestana(p.pestana);

        const { bloqueo, foco, globo } = crearCapas();
        const objetivo = p.objetivo ? p.objetivo() : null;
        const hay = visible(objetivo);
        if (hay) objetivo.scrollIntoView({ block: 'nearest', inline: 'nearest' });

        const ultimo = paso === PASOS.length - 1;
        const puntos = PASOS.map((_, i) => `<span class="${i === paso ? 'on' : ''}"></span>`).join('');
        globo.className = 'tuto-globo' + (hay ? '' : ' centro');
        globo.innerHTML = `
            <button class="tuto-cerrar" title="Cerrar tutorial (Esc)">✕</button>
            <h3>${p.titulo}</h3>
            <div class="tuto-texto">${p.texto}</div>
            <div class="tuto-pie">
                <div class="tuto-paso">Paso ${paso + 1} de ${PASOS.length}<div class="tuto-puntos">${puntos}</div></div>
                ${paso > 0 ? '<button class="tuto-btn" data-t="ant">◀ Anterior</button>' : '<button class="tuto-btn" data-t="salir">Ahora no</button>'}
                <button class="tuto-btn primario" data-t="sig">${ultimo ? '¡Entendido!' : (paso === 0 ? 'Empezar ▶' : 'Siguiente ▶')}</button>
            </div>`;
        globo.querySelector('.tuto-cerrar').onclick = terminar;
        globo.querySelectorAll('[data-t]').forEach(b => b.onclick = () => {
            if (b.dataset.t === 'sig') (ultimo ? terminar() : mostrar(paso + 1));
            else if (b.dataset.t === 'ant') mostrar(paso - 1);
            else terminar();
        });

        bloqueo.classList.toggle('sin-objetivo', !hay);
        foco.style.display = hay ? 'block' : 'none';
        // Esperar un instante a que la pestaña se dibuje antes de medir
        requestAnimationFrame(() => colocar(objetivo, hay));
        setTimeout(() => globo.querySelector('[data-t="sig"]')?.focus(), 50);
    };

    const colocar = (objetivo, hay) => {
        if (!capas) return;
        const { foco, globo } = capas;
        const g = globo.getBoundingClientRect();
        const vw = window.innerWidth, vh = window.innerHeight;
        globo.querySelector('.tuto-flecha')?.remove();
        if (!hay) {
            globo.style.left = Math.round((vw - g.width) / 2) + 'px';
            globo.style.top = Math.round((vh - g.height) / 2) + 'px';
            return;
        }
        const r = objetivo.getBoundingClientRect();
        const m = 6; // margen del resaltado
        foco.style.left = (r.left - m) + 'px';
        foco.style.top = (r.top - m) + 'px';
        foco.style.width = (r.width + m * 2) + 'px';
        foco.style.height = (r.height + m * 2) + 'px';

        // El globo va abajo del elemento; si no cabe, arriba; si tampoco, a un lado
        const sep = 16;
        let top, left, flecha = 'arriba';
        if (r.bottom + sep + g.height < vh) {
            top = r.bottom + sep;
        } else if (r.top - sep - g.height > 0) {
            top = r.top - sep - g.height;
            flecha = 'abajo';
        } else {
            top = Math.max(12, Math.min(vh - g.height - 12, r.top + r.height / 2 - g.height / 2));
            flecha = r.right + sep + g.width < vw ? 'izquierda' : 'derecha';
        }
        if (flecha === 'izquierda') left = r.right + sep;
        else if (flecha === 'derecha') left = r.left - sep - g.width;
        else left = r.left + r.width / 2 - g.width / 2;
        left = Math.max(12, Math.min(vw - g.width - 12, left));
        top = Math.max(12, Math.min(vh - g.height - 12, top));
        globo.style.left = Math.round(left) + 'px';
        globo.style.top = Math.round(top) + 'px';

        // Flechita que apunta al elemento
        const f = document.createElement('div');
        f.className = 'tuto-flecha';
        if (flecha === 'arriba' || flecha === 'abajo') {
            const x = Math.max(16, Math.min(g.width - 30, r.left + r.width / 2 - left - 7));
            f.style.left = x + 'px';
            f.style[flecha === 'arriba' ? 'top' : 'bottom'] = '-7px';
        } else {
            const y = Math.max(16, Math.min(g.height - 30, r.top + r.height / 2 - top - 7));
            f.style.top = y + 'px';
            f.style[flecha === 'izquierda' ? 'left' : 'right'] = '-7px';
        }
        globo.appendChild(f);
    };

    const terminar = () => {
        quitarCapas();
        try { localStorage.setItem(CLAVE_VISTO, 'si'); } catch (e) { }
        if (window.gcAbrirPestana) window.gcAbrirPestana('inicio');
    };

    const iniciar = () => mostrar(0);

    // Teclado: → / Enter siguiente, ← anterior, Esc salir; F1 abre el tutorial
    document.addEventListener('keydown', (e) => {
        if (e.key === 'F1') {
            e.preventDefault();
            iniciar();
            return;
        }
        if (!capas) return;
        if (e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); terminar(); }
        else if (e.key === 'ArrowRight') { e.preventDefault(); e.stopPropagation(); paso < PASOS.length - 1 ? mostrar(paso + 1) : terminar(); }
        else if (e.key === 'ArrowLeft') { e.preventDefault(); e.stopPropagation(); if (paso > 0) mostrar(paso - 1); }
    }, true);

    window.addEventListener('resize', () => {
        if (!capas) return;
        const p = PASOS[paso];
        const objetivo = p.objetivo ? p.objetivo() : null;
        colocar(objetivo, visible(objetivo));
    });

    // Botón "Tutorial" y primera vez
    window.addEventListener('load', () => {
        const btn = document.getElementById('help-btn');
        if (btn) btn.addEventListener('click', iniciar);
        let visto = false;
        try { visto = localStorage.getItem(CLAVE_VISTO) === 'si'; } catch (e) { }
        if (!visto) setTimeout(iniciar, 700);
    });

    window.Tutorial = { iniciar, terminar };
})();
