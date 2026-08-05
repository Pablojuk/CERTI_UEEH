const { app, BrowserWindow, ipcMain, dialog, session, safeStorage, net } = require('electron');
const path = require('path');
const fs = require('fs');
const crypto = require('crypto');
const { spawn, spawnSync } = require('child_process');
const { pathToFileURL } = require('url');

let mainWindow;
let descargaFormatoEnCurso = false;
let autoUpdaterInstancia = null;
let actualizadorConfigurado = false;
let comprobacionActualizacionEnCurso = null;
let descargaActualizacionEnCurso = null;
let actualizacionDisponibleInfo = null;
let actualizacionDescargada = false;
let temporizadorActualizacionInicial = null;
let temporizadorActualizacionPeriodica = null;
const CAPACIDAD_MATERIAS_FORMATO_BGU3 = 16;
const MAX_EXCEL_BYTES = 25 * 1024 * 1024;
const MAX_HTML_BYTES = 20 * 1024 * 1024;
const MAX_PAYLOAD_BYTES = 20 * 1024 * 1024;
const MAX_LOGO_BYTES = 5 * 1024 * 1024;
const MAX_PYTHON_OUTPUT_BYTES = 25 * 1024 * 1024;
const PYTHON_TIMEOUT_MS = 120000;
const PDF_READY_TIMEOUT_MS = 15000;
const PDF_READY_POLL_MS = 100;
const RETRASO_ACTUALIZACION_INICIAL_MS = 8 * 1000;
const INTERVALO_ACTUALIZACIONES_MS = 4 * 60 * 60 * 1000;
const DURACION_PRUEBA_MS = 15 * 24 * 60 * 60 * 1000;
const TOLERANCIA_RELOJ_MS = 1000;
const HASH_LICENCIA_ESPERADO = Buffer.from(
    '09fa7b2c0e4a19c7dc686f28d1eac2db8ed1824ee8f97fb4861886b45f2baad9',
    'hex'
);
const CLAVE_REGISTRO_LICENCIA = 'HKCU\\Software\\CERTI_UEEH';
const VALOR_REGISTRO_LICENCIA = 'LicenseStateV1';
const CANALES_LICENCIA = new Set([
    'obtener-estado-licencia',
    'activar-licencia',
    'iniciar-prueba'
]);
const CANALES_ACTUALIZACIONES = new Set([
    'actualizaciones:buscar',
    'actualizaciones:descargar',
    'actualizaciones:instalar'
]);
let estadoActualizacion = {
    estado: 'deshabilitado',
    versionInstalada: null,
    nuevaVersion: null,
    notas: '',
    porcentaje: null,
    manual: false,
    mensaje: 'Las actualizaciones solo se comprueban en la aplicación instalada.'
};
let estadoLicenciaMemoria = null;
let colaOperacionesLicencia = Promise.resolve();
const EXTENSIONES_EXCEL_PERMITIDAS = new Set(['.xlsx', '.xls']);
const EXTENSIONES_IMAGEN_PERMITIDAS = Object.freeze({
    'image/png': 'png',
    'image/jpeg': 'jpg',
    'image/webp': 'webp'
});
const PLANTILLAS_CERTIFICADO_PERMITIDAS = Object.freeze([
    'FORMATO INICIAL 1.html',
    'FORMATO INICIAL 2.html',
    'PRIMERO DE EGB.html',
    'FORMALO DE ELEMENTAL.html',
    'FORMATO EGBM.html',
    'FORMATO EGBS.html',
    'FORMATO DE 1 Y 2 DE BGU.html',
    'FORMATO DE 3 DE BGU.html'
]);
const CANALES_IPC_PERMITIDOS = Object.freeze([
    ...CANALES_LICENCIA,
    ...CANALES_ACTUALIZACIONES,
    'seleccionar-archivo',
    'descargar-formato',
    'analizar-excel',
    'obtener-catalogo-asignaturas',
    'obtener-escala-cualitativa',
    'generar-boletines',
    'generar-certificados',
    'actualizar-supletorios',
    'leer-plantilla',
    'leer-certificado-generado',
    'imprimir-certificados',
    'abrir-vista-previa-certificado'
]);

function normalizarNotasVersion(releaseNotes) {
    if (typeof releaseNotes === 'string') return releaseNotes.trim();
    if (!Array.isArray(releaseNotes)) return '';
    return releaseNotes
        .map(entrada => {
            if (typeof entrada === 'string') return entrada.trim();
            if (!entrada || typeof entrada !== 'object') return '';
            const version = typeof entrada.version === 'string' ? entrada.version.trim() : '';
            const nota = typeof entrada.note === 'string' ? entrada.note.trim() : '';
            return [version && `Versión ${version}`, nota].filter(Boolean).join('\n');
        })
        .filter(Boolean)
        .join('\n\n');
}

function esVersionSuperior(nuevaVersion, versionInstalada) {
    const analizar = valor => {
        const coincidencia = String(valor || '').trim().match(
            /^v?(\d+)\.(\d+)\.(\d+)(?:-([0-9A-Za-z.-]+))?$/
        );
        if (!coincidencia) return null;
        return {
            numeros: coincidencia.slice(1, 4).map(Number),
            prerelease: coincidencia[4] || ''
        };
    };
    const nueva = analizar(nuevaVersion);
    const instalada = analizar(versionInstalada);
    if (!nueva || !instalada) return false;
    for (let indice = 0; indice < 3; indice += 1) {
        if (nueva.numeros[indice] !== instalada.numeros[indice]) {
            return nueva.numeros[indice] > instalada.numeros[indice];
        }
    }
    if (nueva.prerelease === instalada.prerelease) return false;
    if (!nueva.prerelease) return true;
    if (!instalada.prerelease) return false;
    return nueva.prerelease.localeCompare(instalada.prerelease, undefined, {
        numeric: true,
        sensitivity: 'base'
    }) > 0;
}

function esErrorSinConexion(error) {
    const codigo = String(error?.code || '').toUpperCase();
    const mensaje = String(error?.message || error || '').toLowerCase();
    return [
        'ENOTFOUND',
        'EAI_AGAIN',
        'ECONNREFUSED',
        'ECONNRESET',
        'ETIMEDOUT',
        'ERR_INTERNET_DISCONNECTED',
        'ERR_NAME_NOT_RESOLVED',
        'ERR_NETWORK_CHANGED'
    ].some(valor => codigo.includes(valor) || mensaje.includes(valor.toLowerCase()))
        || /internet|network|conexi[oó]n|socket|dns|offline/.test(mensaje);
}

function obtenerEstadoActualizacionPublico() {
    return {
        estado: estadoActualizacion.estado,
        versionInstalada: estadoActualizacion.versionInstalada,
        nuevaVersion: estadoActualizacion.nuevaVersion,
        notas: estadoActualizacion.notas,
        porcentaje: estadoActualizacion.porcentaje,
        manual: estadoActualizacion.manual,
        mensaje: estadoActualizacion.mensaje
    };
}

function publicarEstadoActualizacion(cambios) {
    estadoActualizacion = {
        ...estadoActualizacion,
        ...cambios,
        versionInstalada: app?.getVersion?.() || estadoActualizacion.versionInstalada
    };
    const publico = obtenerEstadoActualizacionPublico();
    if (mainWindow && !mainWindow.isDestroyed?.()) {
        mainWindow.webContents.send('actualizaciones:estado', publico);
    }
    return publico;
}

function sanitizarValorLogActualizacion(valor) {
    if (typeof valor === 'number' || typeof valor === 'boolean' || valor === null) return valor;
    return String(valor || '')
        .replace(/([?&](?:token|access_token|auth|key)=)[^&\s]+/gi, '$1[REDACTADO]')
        .replace(/(authorization\s*[:=]\s*)([^\s,;]+)/gi, '$1[REDACTADO]')
        .slice(0, 500);
}

function registrarEventoActualizacion(evento, detalles = {}, dependencias = {}) {
    const appActual = dependencias.app || app;
    const fsActual = dependencias.fs || fs;
    const pathActual = dependencias.path || path;
    try {
        const directorio = pathActual.join(appActual.getPath('userData'), 'logs');
        fsActual.mkdirSync(directorio, { recursive: true });
        const permitidos = [
            'versionInstalada',
            'proveedor',
            'url',
            'nuevaVersion',
            'porcentaje',
            'mensaje',
            'empaquetada'
        ];
        const datos = Object.fromEntries(
            permitidos
                .filter(clave => Object.prototype.hasOwnProperty.call(detalles, clave))
                .map(clave => [clave, sanitizarValorLogActualizacion(detalles[clave])])
        );
        const registro = {
            fecha: new Date().toISOString(),
            evento: sanitizarValorLogActualizacion(evento),
            ...datos
        };
        fsActual.appendFileSync(
            pathActual.join(directorio, 'actualizaciones.log'),
            `${JSON.stringify(registro)}\n`,
            'utf8'
        );
        return true;
    } catch (error) {
        if (!appActual?.isPackaged) {
            console.error('[actualizaciones/log] No se pudo escribir el registro.', error);
        }
        return false;
    }
}

function obtenerActualizador() {
    if (!app?.isPackaged) return null;
    if (!autoUpdaterInstancia) {
        const electronUpdater = require('electron-updater');
        autoUpdaterInstancia = electronUpdater.autoUpdater;
    }
    return autoUpdaterInstancia;
}

function configurarActualizador() {
    if (!app?.isPackaged || actualizadorConfigurado) return false;
    const updater = obtenerActualizador();
    if (!updater) return false;

    updater.autoDownload = false;
    updater.autoInstallOnAppQuit = false;
    updater.allowDowngrade = false;
    registrarEventoActualizacion('configurado', {
        versionInstalada: app.getVersion(),
        proveedor: 'github:Pablojuk/CERTI_UEEH',
        url: 'https://github.com/Pablojuk/CERTI_UEEH/releases',
        empaquetada: app.isPackaged
    });

    updater.on('checking-for-update', () => {
        registrarEventoActualizacion('buscando', {
            versionInstalada: app.getVersion(),
            proveedor: 'github:Pablojuk/CERTI_UEEH'
        });
        publicarEstadoActualizacion({
            estado: 'buscando',
            mensaje: 'Buscando actualización...',
            porcentaje: null
        });
    });
    updater.on('update-available', info => {
        if (!esVersionSuperior(info?.version, app.getVersion())) {
            actualizacionDisponibleInfo = null;
            actualizacionDescargada = false;
            publicarEstadoActualizacion({
                estado: 'no-disponible',
                nuevaVersion: null,
                notas: '',
                porcentaje: null,
                mensaje: 'La aplicación está actualizada.'
            });
            return;
        }
        actualizacionDisponibleInfo = info;
        actualizacionDescargada = false;
        registrarEventoActualizacion('disponible', {
            versionInstalada: app.getVersion(),
            nuevaVersion: info?.version || ''
        });
        publicarEstadoActualizacion({
            estado: 'disponible',
            nuevaVersion: info?.version || null,
            notas: normalizarNotasVersion(info?.releaseNotes),
            porcentaje: null,
            mensaje: 'Nueva versión disponible.'
        });
    });
    updater.on('update-not-available', () => {
        actualizacionDisponibleInfo = null;
        actualizacionDescargada = false;
        registrarEventoActualizacion('no-disponible', {
            versionInstalada: app.getVersion()
        });
        publicarEstadoActualizacion({
            estado: 'no-disponible',
            nuevaVersion: null,
            notas: '',
            porcentaje: null,
            mensaje: 'La aplicación está actualizada.'
        });
    });
    updater.on('download-progress', progreso => {
        const porcentaje = Math.max(0, Math.min(100, Number(progreso?.percent) || 0));
        registrarEventoActualizacion('progreso', { porcentaje: Number(porcentaje.toFixed(1)) });
        publicarEstadoActualizacion({
            estado: 'descargando',
            porcentaje,
            mensaje: `Descargando actualización: ${porcentaje.toFixed(1)} %.`
        });
    });
    updater.on('update-downloaded', info => {
        actualizacionDisponibleInfo = info || actualizacionDisponibleInfo;
        actualizacionDescargada = true;
        descargaActualizacionEnCurso = null;
        registrarEventoActualizacion('descargada', {
            nuevaVersion: actualizacionDisponibleInfo?.version || ''
        });
        publicarEstadoActualizacion({
            estado: 'descargada',
            nuevaVersion: actualizacionDisponibleInfo?.version || null,
            notas: normalizarNotasVersion(actualizacionDisponibleInfo?.releaseNotes),
            porcentaje: 100,
            mensaje: 'La actualización está lista para instalar.'
        });
    });
    updater.on('error', error => {
        comprobacionActualizacionEnCurso = null;
        descargaActualizacionEnCurso = null;
        const sinConexion = esErrorSinConexion(error);
        registrarErrorSeguro('actualizaciones', error);
        registrarEventoActualizacion('error', {
            mensaje: error?.message || String(error || 'Error desconocido')
        });
        publicarEstadoActualizacion({
            estado: sinConexion ? 'sin-conexion' : 'error',
            porcentaje: null,
            mensaje: sinConexion
                ? 'Sin conexión. La aplicación seguirá funcionando normalmente.'
                : 'Error al comprobar la actualización.'
        });
    });

    actualizadorConfigurado = true;
    publicarEstadoActualizacion({
        estado: 'listo',
        mensaje: 'Actualizaciones automáticas activadas.'
    });
    return true;
}

async function comprobarActualizaciones({ manual = false } = {}) {
    if (!app?.isPackaged) {
        return publicarEstadoActualizacion({
            estado: 'deshabilitado',
            manual,
            mensaje: 'Las actualizaciones solo se comprueban en la aplicación instalada.'
        });
    }
    configurarActualizador();
    if (comprobacionActualizacionEnCurso) {
        return obtenerEstadoActualizacionPublico();
    }
    if (net?.isOnline && !net.isOnline()) {
        return publicarEstadoActualizacion({
            estado: 'sin-conexion',
            manual,
            mensaje: 'Sin conexión. No se pudo buscar una actualización.'
        });
    }

    publicarEstadoActualizacion({
        estado: 'buscando',
        manual,
        mensaje: 'Buscando actualización...',
        porcentaje: null
    });
    const updater = obtenerActualizador();
    comprobacionActualizacionEnCurso = Promise.resolve()
        .then(() => updater.checkForUpdates())
        .catch(error => {
            const sinConexion = esErrorSinConexion(error);
            registrarErrorSeguro('actualizaciones/comprobar', error);
            return publicarEstadoActualizacion({
                estado: sinConexion ? 'sin-conexion' : 'error',
                manual,
                mensaje: sinConexion
                    ? 'Sin conexión. No se pudo buscar una actualización.'
                    : 'Error al comprobar la actualización.'
            });
        })
        .finally(() => {
            comprobacionActualizacionEnCurso = null;
        });
    await comprobacionActualizacionEnCurso;
    return obtenerEstadoActualizacionPublico();
}

async function descargarActualizacion() {
    if (!app?.isPackaged || !actualizacionDisponibleInfo) {
        return publicarEstadoActualizacion({
            estado: 'error',
            mensaje: 'No hay una actualización disponible para descargar.'
        });
    }
    if (actualizacionDescargada) return obtenerEstadoActualizacionPublico();
    if (descargaActualizacionEnCurso) return obtenerEstadoActualizacionPublico();

    publicarEstadoActualizacion({
        estado: 'descargando',
        porcentaje: 0,
        mensaje: 'Descargando actualización: 0 %.'
    });
    const updater = obtenerActualizador();
    descargaActualizacionEnCurso = Promise.resolve()
        .then(() => updater.downloadUpdate())
        .catch(error => {
            const sinConexion = esErrorSinConexion(error);
            registrarErrorSeguro('actualizaciones/descargar', error);
            return publicarEstadoActualizacion({
                estado: sinConexion ? 'sin-conexion' : 'error',
                porcentaje: null,
                mensaje: sinConexion
                    ? 'Sin conexión. La descarga no pudo continuar.'
                    : 'Error al descargar la actualización.'
            });
        })
        .finally(() => {
            descargaActualizacionEnCurso = null;
        });
    await descargaActualizacionEnCurso;
    return obtenerEstadoActualizacionPublico();
}

function instalarActualizacion() {
    if (!app?.isPackaged || !actualizacionDescargada) {
        return publicarEstadoActualizacion({
            estado: 'error',
            mensaje: 'La actualización debe descargarse correctamente antes de instalarla.'
        });
    }
    const updater = obtenerActualizador();
    registrarEventoActualizacion('instalacion-solicitada', {
        versionInstalada: app.getVersion(),
        nuevaVersion: actualizacionDisponibleInfo?.version || ''
    });
    const respuesta = publicarEstadoActualizacion({
        estado: 'instalando',
        mensaje: 'Reiniciando para instalar la actualización...'
    });
    setImmediate(() => updater.quitAndInstall(false, true));
    return respuesta;
}

function programarActualizaciones() {
    if (!app?.isPackaged) return false;
    if (!actualizadorConfigurado && !configurarActualizador()) return false;
    if (!temporizadorActualizacionInicial) {
        temporizadorActualizacionInicial = setTimeout(() => {
            temporizadorActualizacionInicial = null;
            comprobarActualizaciones().catch(error => {
                registrarErrorSeguro('actualizaciones/inicial', error);
            });
        }, RETRASO_ACTUALIZACION_INICIAL_MS);
        temporizadorActualizacionInicial.unref?.();
    }
    if (!temporizadorActualizacionPeriodica) {
        temporizadorActualizacionPeriodica = setInterval(() => {
            if (!net?.isOnline || net.isOnline()) {
                comprobarActualizaciones().catch(error => {
                    registrarErrorSeguro('actualizaciones/periodica', error);
                });
            }
        }, INTERVALO_ACTUALIZACIONES_MS);
        temporizadorActualizacionPeriodica.unref?.();
    }
    return true;
}

function calcularSha256(valor, cryptoActual = crypto) {
    return cryptoActual.createHash('sha256').update(String(valor), 'utf8').digest();
}

function validarLicenciaIngresada(valor, cryptoActual = crypto) {
    if (typeof valor !== 'string' || valor.length < 1 || valor.length > 256) return false;
    const hashIngresado = calcularSha256(valor, cryptoActual);
    return hashIngresado.length === HASH_LICENCIA_ESPERADO.length
        && cryptoActual.timingSafeEqual(hashIngresado, HASH_LICENCIA_ESPERADO);
}

function normalizarEstadoLicencia(valor) {
    if (!esObjetoPlano(valor)) return null;
    const machineHash = typeof valor.machineHash === 'string'
        ? valor.machineHash.toLowerCase()
        : '';
    const trialStart = Number(valor.trialStart);
    const trialEnd = Number(valor.trialEnd);
    const lastRun = Number(valor.lastRun);
    if (
        !/^[a-f0-9]{64}$/.test(machineHash)
        || !Number.isSafeInteger(trialStart)
        || !Number.isSafeInteger(trialEnd)
        || !Number.isSafeInteger(lastRun)
        || trialStart <= 0
        || trialEnd < trialStart
        || lastRun <= 0
    ) {
        return null;
    }
    return {
        machineHash,
        trialStart,
        trialEnd,
        lastRun,
        activated: valor.activated === true,
        expired: valor.expired === true
    };
}

function crearEstadoPrueba(machineHash, ahora = Date.now()) {
    return {
        machineHash,
        trialStart: ahora,
        trialEnd: ahora + DURACION_PRUEBA_MS,
        lastRun: ahora,
        activated: false,
        expired: false
    };
}

function fusionarEstadosLicencia(estados, machineHash) {
    const compatibles = (Array.isArray(estados) ? estados : [])
        .map(normalizarEstadoLicencia)
        .filter(estado => estado?.machineHash === machineHash);
    if (compatibles.length === 0) return null;

    const trialStart = Math.min(...compatibles.map(estado => estado.trialStart));
    const limiteOriginal = trialStart + DURACION_PRUEBA_MS;
    return {
        machineHash,
        trialStart,
        trialEnd: Math.min(
            limiteOriginal,
            ...compatibles.map(estado => estado.trialEnd)
        ),
        lastRun: Math.max(...compatibles.map(estado => estado.lastRun)),
        activated: compatibles.some(estado => estado.activated),
        expired: compatibles.some(estado => estado.expired)
    };
}

function evaluarEstadoLicencia(estado, ahora = Date.now()) {
    const normalizado = normalizarEstadoLicencia(estado);
    if (!normalizado) {
        return {
            tipo: 'no_iniciada',
            autorizado: false,
            activada: false,
            diasRestantes: 15
        };
    }
    if (normalizado.activated) {
        return {
            tipo: 'activada',
            autorizado: true,
            activada: true,
            diasRestantes: null
        };
    }
    if (normalizado.expired || ahora >= normalizado.trialEnd) {
        return {
            tipo: 'vencida',
            autorizado: false,
            activada: false,
            diasRestantes: 0
        };
    }
    if (ahora + TOLERANCIA_RELOJ_MS < normalizado.lastRun) {
        return {
            tipo: 'fecha_invalida',
            autorizado: false,
            activada: false,
            diasRestantes: Math.max(
                0,
                Math.ceil((normalizado.trialEnd - normalizado.lastRun) / (24 * 60 * 60 * 1000))
            )
        };
    }
    return {
        tipo: 'prueba',
        autorizado: true,
        activada: false,
        diasRestantes: Math.max(
            0,
            Math.ceil((normalizado.trialEnd - ahora) / (24 * 60 * 60 * 1000))
        )
    };
}

function registrarErrorSeguro(contexto, error) {
    if (!app?.isPackaged) {
        console.error(`[${contexto}]`, error);
        return;
    }
    console.error(`[${contexto}] La operación no pudo completarse.`);
}

function obtenerEjecutableRegistro() {
    const raizWindows = process.env.SystemRoot;
    if (typeof raizWindows === 'string' && path.isAbsolute(raizWindows)) {
        return path.join(raizWindows, 'System32', 'reg.exe');
    }
    return 'reg.exe';
}

function ejecutarRegistro(argumentos, spawnSyncActual = spawnSync) {
    try {
        return spawnSyncActual(obtenerEjecutableRegistro(), argumentos, {
            shell: false,
            windowsHide: true,
            encoding: 'utf8',
            timeout: 5000,
            maxBuffer: 256 * 1024
        });
    } catch {
        return null;
    }
}

function obtenerHashMaquina(spawnSyncActual = spawnSync, cryptoActual = crypto) {
    const resultado = ejecutarRegistro([
        'query',
        'HKLM\\SOFTWARE\\Microsoft\\Cryptography',
        '/v',
        'MachineGuid'
    ], spawnSyncActual);
    const coincidencia = resultado?.status === 0
        ? String(resultado.stdout || '').match(/MachineGuid\s+REG_\w+\s+([^\r\n]+)/i)
        : null;
    const identificador = coincidencia?.[1]?.trim();
    if (!identificador || identificador.length > 256) {
        throw new Error('No se pudo obtener el identificador estable del equipo.');
    }
    return calcularSha256(identificador.toUpperCase(), cryptoActual).toString('hex');
}

function obtenerUbicacionesLicencia(appActual = app) {
    const appData = appActual.getPath('appData');
    const localAppData = (
        typeof process.env.LOCALAPPDATA === 'string'
        && path.isAbsolute(process.env.LOCALAPPDATA)
    )
        ? process.env.LOCALAPPDATA
        : path.resolve(appData, '..', 'Local');
    const programData = (
        typeof process.env.PROGRAMDATA === 'string'
        && path.isAbsolute(process.env.PROGRAMDATA)
    )
        ? process.env.PROGRAMDATA
        : path.join(path.parse(appData).root, 'ProgramData');
    return {
        archivos: [
            path.join(localAppData, 'CERTI_UEEH', 'license_state.dat'),
            path.join(programData, 'CERTI_UEEH', 'license_state.dat')
        ],
        registro: {
            clave: CLAVE_REGISTRO_LICENCIA,
            valor: VALOR_REGISTRO_LICENCIA
        }
    };
}

function descifrarEstadoLicencia(textoCifrado, safeStorageActual = safeStorage) {
    if (
        typeof textoCifrado !== 'string'
        || textoCifrado.length === 0
        || textoCifrado.length > 256 * 1024
        || !/^[A-Za-z0-9+/=]+$/.test(textoCifrado)
    ) {
        return null;
    }
    try {
        const plano = safeStorageActual.decryptString(Buffer.from(textoCifrado, 'base64'));
        if (Buffer.byteLength(plano, 'utf8') > 64 * 1024) return null;
        return normalizarEstadoLicencia(JSON.parse(plano));
    } catch {
        return null;
    }
}

function cifrarEstadoLicencia(estado, safeStorageActual = safeStorage) {
    const normalizado = normalizarEstadoLicencia(estado);
    if (!normalizado) throw new Error('El estado de licencia no es válido.');
    return safeStorageActual.encryptString(JSON.stringify(normalizado)).toString('base64');
}

function leerEstadosLicencia(dependencias = {}) {
    const fsActual = dependencias.fs || fs;
    const safeStorageActual = dependencias.safeStorage || safeStorage;
    const ubicaciones = dependencias.ubicaciones || obtenerUbicacionesLicencia();
    const estados = [];

    for (const rutaArchivo of ubicaciones.archivos) {
        try {
            const stat = fsActual.lstatSync(rutaArchivo);
            if (!stat.isFile() || stat.isSymbolicLink() || stat.size > 256 * 1024) continue;
            const estado = descifrarEstadoLicencia(
                fsActual.readFileSync(rutaArchivo, 'utf8').trim(),
                safeStorageActual
            );
            if (estado) estados.push(estado);
        } catch {
            // Una copia ausente o dañada no invalida las otras copias redundantes.
        }
    }

    const consulta = ejecutarRegistro([
        'query',
        ubicaciones.registro.clave,
        '/v',
        ubicaciones.registro.valor
    ], dependencias.spawnSync || spawnSync);
    if (consulta?.status === 0) {
        const patron = new RegExp(
            `${ubicaciones.registro.valor}\\s+REG_\\w+\\s+([^\\r\\n]+)`,
            'i'
        );
        const estado = descifrarEstadoLicencia(
            String(consulta.stdout || '').match(patron)?.[1]?.trim() || '',
            safeStorageActual
        );
        if (estado) estados.push(estado);
    }
    return estados;
}

function escribirArchivoEstadoSeguro(rutaArchivo, contenido, fsActual = fs) {
    const directorio = path.dirname(rutaArchivo);
    const temporal = `${rutaArchivo}.tmp-${process.pid}`;
    fsActual.mkdirSync(directorio, { recursive: true });
    try {
        fsActual.writeFileSync(temporal, contenido, { encoding: 'utf8', flag: 'w' });
        fsActual.copyFileSync(temporal, rutaArchivo);
    } finally {
        try {
            fsActual.unlinkSync(temporal);
        } catch {
            // La copia definitiva ya fue escrita; la limpieza se reintentará al guardar.
        }
    }
}

function guardarEstadoLicencia(estado, dependencias = {}) {
    const fsActual = dependencias.fs || fs;
    const safeStorageActual = dependencias.safeStorage || safeStorage;
    if (!safeStorageActual?.isEncryptionAvailable?.()) {
        throw new Error('El cifrado seguro del sistema no está disponible.');
    }
    const ubicaciones = dependencias.ubicaciones || obtenerUbicacionesLicencia();
    const cifrado = cifrarEstadoLicencia(estado, safeStorageActual);
    let copiasGuardadas = 0;

    for (const rutaArchivo of ubicaciones.archivos) {
        try {
            escribirArchivoEstadoSeguro(rutaArchivo, cifrado, fsActual);
            copiasGuardadas += 1;
        } catch (error) {
            registrarErrorSeguro('licencia/guardar-archivo', error);
        }
    }

    const escrituraRegistro = ejecutarRegistro([
        'add',
        ubicaciones.registro.clave,
        '/v',
        ubicaciones.registro.valor,
        '/t',
        'REG_SZ',
        '/d',
        cifrado,
        '/f'
    ], dependencias.spawnSync || spawnSync);
    if (escrituraRegistro?.status === 0) copiasGuardadas += 1;
    if (copiasGuardadas === 0) {
        throw new Error('No fue posible conservar el estado de licencia.');
    }
    return copiasGuardadas;
}

function mensajeEstadoLicencia(tipo) {
    const mensajes = {
        no_iniciada: 'Puede iniciar la prueba gratuita de 15 días o ingresar su licencia.',
        prueba: 'La aplicación está funcionando en período de prueba.',
        activada: 'Licencia activada permanentemente en este equipo.',
        vencida: 'La prueba de 15 días terminó. Ingrese una licencia para continuar.',
        fecha_invalida: 'La fecha del sistema retrocedió. Corríjala o ingrese una licencia.',
        error: 'No fue posible verificar la licencia de forma segura.'
    };
    return mensajes[tipo] || mensajes.error;
}

function convertirEstadoLicenciaPublico(evaluacion) {
    return {
        success: evaluacion.tipo !== 'error',
        estado: evaluacion.tipo,
        autorizado: evaluacion.autorizado === true,
        activada: evaluacion.activada === true,
        diasRestantes: Number.isInteger(evaluacion.diasRestantes)
            ? evaluacion.diasRestantes
            : null,
        puedeIniciarPrueba: evaluacion.tipo === 'no_iniciada'
            || evaluacion.tipo === 'prueba',
        mensaje: mensajeEstadoLicencia(evaluacion.tipo)
    };
}

function encolarOperacionLicencia(operacion) {
    const ejecucion = colaOperacionesLicencia.then(operacion, operacion);
    colaOperacionesLicencia = ejecucion.catch(() => undefined);
    return ejecucion;
}

async function obtenerEstadoLicenciaInterno({ registrarEjecucion = true } = {}) {
    return encolarOperacionLicencia(async () => {
        try {
            if (!safeStorage?.isEncryptionAvailable?.()) {
                return { tipo: 'error', autorizado: false, activada: false, diasRestantes: null };
            }
            const machineHash = obtenerHashMaquina();
            if (!estadoLicenciaMemoria) {
                estadoLicenciaMemoria = fusionarEstadosLicencia(
                    leerEstadosLicencia(),
                    machineHash
                );
            }
            const ahora = Date.now();
            const evaluacion = evaluarEstadoLicencia(estadoLicenciaMemoria, ahora);
            if (estadoLicenciaMemoria && registrarEjecucion) {
                if (evaluacion.tipo === 'vencida') estadoLicenciaMemoria.expired = true;
                if (ahora >= estadoLicenciaMemoria.lastRun) estadoLicenciaMemoria.lastRun = ahora;
                guardarEstadoLicencia(estadoLicenciaMemoria);
            }
            return evaluacion;
        } catch (error) {
            registrarErrorSeguro('licencia/verificar', error);
            return { tipo: 'error', autorizado: false, activada: false, diasRestantes: null };
        }
    });
}

function esObjetoPlano(valor) {
    if (!valor || typeof valor !== 'object' || Array.isArray(valor)) return false;
    const prototipo = Object.getPrototypeOf(valor);
    return prototipo === Object.prototype || prototipo === null;
}

function tamanoJson(valor) {
    try {
        return Buffer.byteLength(JSON.stringify(valor), 'utf8');
    } catch {
        return Number.POSITIVE_INFINITY;
    }
}

function esSegmentoRutaSeguro(valor, maximo = 180) {
    if (typeof valor !== 'string') return false;
    const texto = valor.trim();
    return Boolean(
        texto
        && texto.length <= maximo
        && texto !== '.'
        && texto !== '..'
        && !/[\\/\0-\x1f\x7f]/.test(texto)
    );
}

function resolverRutaHija(directorioBase, ...segmentos) {
    const base = path.resolve(directorioBase);
    const destino = path.resolve(base, ...segmentos);
    const relativa = path.relative(base, destino);
    if (!relativa || (!relativa.startsWith('..') && !path.isAbsolute(relativa))) {
        return destino;
    }
    return null;
}

function validarArchivoExcel(ruta, fsActual = fs, pathActual = path) {
    if (
        typeof ruta !== 'string'
        || ruta.length === 0
        || ruta.length > 2048
        || ruta.includes('\0')
        || !pathActual.isAbsolute(ruta)
        || !EXTENSIONES_EXCEL_PERMITIDAS.has(pathActual.extname(ruta).toLowerCase())
    ) {
        return false;
    }
    try {
        const estado = fsActual.statSync(ruta);
        return estado.isFile() && estado.size > 0 && estado.size <= MAX_EXCEL_BYTES;
    } catch {
        return false;
    }
}

function validarHtmlRecibido(htmlContent) {
    return (
        typeof htmlContent === 'string'
        && htmlContent.length > 0
        && Buffer.byteLength(htmlContent, 'utf8') <= MAX_HTML_BYTES
    );
}

function normalizarSolicitudImpresionCertificados(solicitud) {
    if (typeof solicitud === 'string') {
        return { html: solicitud, diagnostico: null };
    }
    if (!esObjetoPlano(solicitud) || !esObjetoPlano(solicitud.diagnostico)) return null;
    const diagnostico = solicitud.diagnostico;
    if (
        !Number.isSafeInteger(diagnostico.notasEsperadas)
        || diagnostico.notasEsperadas < 0
        || diagnostico.notasEsperadas > 100000
        || !Number.isSafeInteger(diagnostico.estudiantesEsperados)
        || diagnostico.estudiantesEsperados < 1
        || diagnostico.estudiantesEsperados > 2000
        || !esSegmentoRutaSeguro(String(diagnostico.cursoId || ''), 100)
        || typeof diagnostico.plantilla !== 'string'
        || !PLANTILLAS_CERTIFICADO_PERMITIDAS.includes(diagnostico.plantilla)
    ) {
        return null;
    }
    return {
        html: solicitud.html,
        diagnostico: {
            notasEsperadas: diagnostico.notasEsperadas,
            estudiantesEsperados: diagnostico.estudiantesEsperados,
            cursoId: String(diagnostico.cursoId),
            plantilla: diagnostico.plantilla
        }
    };
}

function obtenerDestinoPdfPruebaEmpaquetada(dependencias = {}) {
    const appActual = dependencias.app || app;
    const fsActual = dependencias.fs || fs;
    const pathActual = dependencias.path || path;
    const entorno = dependencias.env || process.env;
    if (entorno.CERTI_MODO_PRUEBA_EMPAQUETADA !== '1') return null;
    const destino = entorno.CERTI_PRUEBA_PDF_DESTINO;
    const baseTemporal = pathActual.resolve(appActual.getPath('temp'));
    const destinoResuelto = typeof destino === 'string' && pathActual.isAbsolute(destino)
        ? pathActual.resolve(destino)
        : '';
    const relativa = destinoResuelto ? pathActual.relative(baseTemporal, destinoResuelto) : '..';
    if (
        !appActual.isPackaged
        || !destinoResuelto
        || pathActual.extname(destinoResuelto).toLowerCase() !== '.pdf'
        || relativa.startsWith('..')
        || pathActual.isAbsolute(relativa)
        || !fsActual.existsSync(pathActual.dirname(destinoResuelto))
    ) {
        const error = new Error('El destino de la prueba PDF empaquetada no es válido.');
        error.code = 'PDF_DESTINO_PRUEBA_INVALIDO';
        throw error;
    }
    return destinoResuelto;
}

async function esperarDocumentoListoParaPdf(webContents, diagnostico, dependencias = {}) {
    const ahora = dependencias.ahora || Date.now;
    const esperar = dependencias.esperar || (milisegundos => (
        new Promise(resolve => setTimeout(resolve, milisegundos))
    ));
    const timeoutMs = dependencias.timeoutMs || PDF_READY_TIMEOUT_MS;
    const intervaloMs = dependencias.intervaloMs || PDF_READY_POLL_MS;
    const inicio = ahora();

    while (ahora() - inicio <= timeoutMs) {
        const estado = await webContents.executeJavaScript(`(() => {
            const celdas = Array.from(document.querySelectorAll(
                '.cert-page tr[data-subject] td[data-academic-value="true"]'
            ));
            return {
                readyState: document.readyState,
                marca: window.__CERTI_PDF_READY__ === true,
                fuentes: !document.fonts || document.fonts.status === 'loaded',
                imagenes: Array.from(document.images).every(imagen => imagen.complete),
                paginas: document.querySelectorAll('.cert-page').length,
                notas: celdas.filter(celda => celda.textContent.trim() !== '').length
            };
        })()`, true);

        if (
            estado?.readyState === 'complete'
            && estado.marca
            && estado.fuentes
            && estado.imagenes
        ) {
            if (diagnostico) {
                if (estado.paginas !== diagnostico.estudiantesEsperados) {
                    const error = new Error(
                        `Se esperaban ${diagnostico.estudiantesEsperados} certificado(s), pero el documento contiene ${estado.paginas}.`
                    );
                    error.code = 'PDF_CERTIFICADOS_INCOMPLETOS';
                    throw error;
                }
                if (estado.notas !== diagnostico.notasEsperadas) {
                    const error = new Error(
                        `Se esperaban ${diagnostico.notasEsperadas} notas, pero el documento contiene ${estado.notas}.`
                    );
                    error.code = 'PDF_NOTAS_INCOMPLETAS';
                    throw error;
                }
            }
            return estado;
        }
        await esperar(intervaloMs);
    }
    const error = new Error('El certificado no terminó de cargar antes de generar el PDF.');
    error.code = 'PDF_DOCUMENTO_NO_LISTO';
    throw error;
}

function validarPayloadGeneracion(datos) {
    if (!esObjetoPlano(datos) || tamanoJson(datos) > MAX_PAYLOAD_BYTES) {
        return 'Los datos enviados no tienen una estructura o tamaño válido.';
    }
    if (!esSegmentoRutaSeguro(datos.cursoActivoId || 'default', 100)) {
        return 'El identificador del curso no es válido.';
    }
    const estudiantes = Array.isArray(datos.datos_consolidados)
        ? datos.datos_consolidados
        : [];
    if (estudiantes.length > 2000) {
        return 'La cantidad de estudiantes supera el límite permitido.';
    }
    const idInvalido = estudiantes.some(
        estudiante => !esObjetoPlano(estudiante)
            || !esSegmentoRutaSeguro(String(estudiante.id_real || ''), 180)
    );
    return idInvalido ? 'Uno de los identificadores de estudiante no es válido.' : null;
}

function esSolicitudIpcAutorizada(evento, ventana = mainWindow) {
    if (!evento?.sender || !ventana?.webContents || evento.sender !== ventana.webContents) {
        return false;
    }
    const urlEmisor = evento.senderFrame?.url || evento.sender.getURL?.() || '';
    return urlEmisor === pathToFileURL(path.join(__dirname, 'index.html')).href;
}

function registrarManejadorIpcSeguro(canal, manejador) {
    if (!CANALES_IPC_PERMITIDOS.includes(canal)) {
        throw new Error(`Canal IPC no autorizado: ${canal}`);
    }
    ipcMain.handle(canal, async (evento, ...argumentos) => {
        if (!esSolicitudIpcAutorizada(evento)) {
            return {
                success: false,
                code: 'IPC_NO_AUTORIZADO',
                error: 'La solicitud no proviene de una ventana autorizada.'
            };
        }
        if (!CANALES_LICENCIA.has(canal)) {
            const estadoLicencia = await obtenerEstadoLicenciaInterno();
            if (!estadoLicencia.autorizado) {
                return {
                    success: false,
                    code: 'LICENCIA_REQUERIDA',
                    error: mensajeEstadoLicencia(estadoLicencia.tipo)
                };
            }
        }
        return manejador(evento, ...argumentos);
    });
}

function endurecerVentana(ventana) {
    ventana.webContents.setWindowOpenHandler(() => ({ action: 'deny' }));
    ventana.webContents.on('will-navigate', evento => evento.preventDefault());
}

function normalizarOpcionesSeleccionArchivo(opciones) {
    const filtros = Array.isArray(opciones?.filters) ? opciones.filters : [];
    const permitidas = new Set(['xlsx', 'xls', 'png', 'jpg', 'jpeg', 'webp']);
    const filtrosSeguros = filtros.slice(0, 5).map(filtro => ({
        name: typeof filtro?.name === 'string'
            ? filtro.name.replace(/[\0-\x1f\x7f]/g, '').slice(0, 80)
            : 'Archivo permitido',
        extensions: Array.isArray(filtro?.extensions)
            ? filtro.extensions
                .map(extension => String(extension).toLowerCase().replace(/^\./, ''))
                .filter(extension => permitidas.has(extension))
                .slice(0, 6)
            : []
    })).filter(filtro => filtro.extensions.length > 0);
    return {
        properties: ['openFile'],
        filters: filtrosSeguros.length > 0
            ? filtrosSeguros
            : [{ name: 'Archivos permitidos', extensions: [...permitidas] }]
    };
}

function guardarImagenBase64(base64Str, prefix, cursoId, logosDir) {
    if (!base64Str) return null;
    if (typeof base64Str !== 'string' || !esSegmentoRutaSeguro(cursoId, 100)) {
        return null;
    }
    const coincidencia = base64Str.match(/^data:(image\/(?:png|jpeg|webp));base64,([A-Za-z0-9+/=\r\n]+)$/);
    if (!coincidencia || !EXTENSIONES_IMAGEN_PERMITIDAS[coincidencia[1]]) {
        return null;
    }
    const buffer = Buffer.from(coincidencia[2], 'base64');
    if (buffer.length === 0 || buffer.length > MAX_LOGO_BYTES) return null;
    const extension = EXTENSIONES_IMAGEN_PERMITIDAS[coincidencia[1]];
    const tempPath = resolverRutaHija(logosDir, `${prefix}_${cursoId}.${extension}`);
    if (!tempPath) return null;
    fs.writeFileSync(tempPath, buffer, { flag: 'w' });
    return tempPath;
}

const FORMATOS_NOTAS = Object.freeze({
    egb_2_7: Object.freeze({
        archivoInterno: 'formato_2_a_7_egb.xlsx',
        nombreDescarga: 'FORMATO PARA 2.º a 7.º EGB.xlsx',
        tituloDialogo: 'Guardar formato para 2.º a 7.º de EGB'
    }),
    egb_8_10: Object.freeze({
        archivoInterno: 'formato_8_a_10_egb.xlsx',
        nombreDescarga: 'FORMATO PARA 8.º a 10.º EGB.xlsx',
        tituloDialogo: 'Guardar formato para 8.º a 10.º de EGB'
    }),
    bgu_1_2: Object.freeze({
        archivoInterno: 'formato_1_y_2_bgu.xlsx',
        nombreDescarga: 'FORMATO PARA 1.º y 2.º BGU.xlsx',
        tituloDialogo: 'Guardar formato para 1.º y 2.º de BGU'
    }),
    bgu_3: Object.freeze({
        archivoInterno: 'formato_3_bgu.xlsx',
        nombreDescarga: 'FORMATO PARA 3.º BGU.xlsx',
        tituloDialogo: 'Guardar formato para 3.º de BGU'
    })
});

const MENSAJES_FORMATOS_NOTAS = Object.freeze({
    invalido: 'La opción seleccionada no corresponde a un formato autorizado.',
    noDisponible: 'No se encontró el formato seleccionado dentro de los recursos de la aplicación. Vuelva a instalar la aplicación o comuníquese con el administrador.',
    copiaFallida: 'No fue posible copiar el archivo en la ubicación seleccionada. Verifique los permisos de la carpeta e intente nuevamente.',
    sinOptativas: 'Seleccione al menos una asignatura optativa para generar el formato de 3.º de BGU.',
    optativasInvalidas: 'La selección contiene asignaturas que no están autorizadas como optativas de 3.º de BGU.',
    capacidadExcedida: 'La selección supera la capacidad de 16 asignaturas del formato de 3.º de BGU. Reduzca la cantidad de optativas.',
    generacionFallida: 'No fue posible generar el formato personalizado de 3.º de BGU. Verifique la selección e intente nuevamente.',
    destinoNoPermitido: 'La plantilla original es de solo lectura. Seleccione otra ubicación para guardar la copia generada.',
    enCurso: 'Ya existe una descarga de formato en curso.'
});

function obtenerCatalogoAsignaturasSeguro(directorioAplicacion = __dirname, fsActual = fs, pathActual = path) {
    const rutaCatalogo = pathActual.join(directorioAplicacion, 'catalogo_asignaturas.json');
    const datos = JSON.parse(fsActual.readFileSync(rutaCatalogo, 'utf-8'));
    return Array.isArray(datos) ? datos : [];
}

function validarSeleccionOptativasBgu3(optativas, catalogo) {
    const entradas = Array.isArray(catalogo) ? catalogo : [];
    const autorizadas = entradas
        .filter(entrada => entrada?.es_optativa_bgu3 === true)
        .sort((a, b) => Number(a.orden || 999) - Number(b.orden || 999));
    const nombresAutorizados = new Map(
        autorizadas.map(entrada => [String(entrada.nombre), entrada])
    );
    const seleccion = Array.isArray(optativas)
        ? optativas.map(nombre => String(nombre || '').trim()).filter(Boolean)
        : [];
    if (seleccion.length === 0) {
        return {
            valida: false,
            code: 'OPTATIVAS_REQUERIDAS',
            error: MENSAJES_FORMATOS_NOTAS.sinOptativas
        };
    }
    if (
        new Set(seleccion).size !== seleccion.length
        || seleccion.some(nombre => !nombresAutorizados.has(nombre))
    ) {
        return {
            valida: false,
            code: 'OPTATIVAS_INVALIDAS',
            error: MENSAJES_FORMATOS_NOTAS.optativasInvalidas
        };
    }

    const materiasFijas = entradas.filter(
        entrada => (
            Array.isArray(entrada?.grados)
            && entrada.grados.includes('BGU_3')
            && entrada.es_optativa_bgu3 !== true
        )
    ).length + 1; // Evaluación comportamental no forma parte del catálogo.
    const totalMaterias = materiasFijas + seleccion.length;
    if (totalMaterias > CAPACIDAD_MATERIAS_FORMATO_BGU3) {
        return {
            valida: false,
            code: 'CAPACIDAD_EXCEDIDA',
            error: MENSAJES_FORMATOS_NOTAS.capacidadExcedida,
            totalMaterias,
            capacidad: CAPACIDAD_MATERIAS_FORMATO_BGU3
        };
    }

    const seleccionOrdenada = autorizadas
        .filter(entrada => seleccion.includes(entrada.nombre))
        .map(entrada => entrada.nombre);
    return {
        valida: true,
        optativas: seleccionOrdenada,
        materiasFijas,
        totalMaterias,
        capacidad: CAPACIDAD_MATERIAS_FORMATO_BGU3
    };
}

function obtenerDirectorioFormatosNotas(
    appActual = app,
    resourcesPathActual = process.resourcesPath,
    directorioDesarrollo = __dirname
) {
    const base = appActual && appActual.isPackaged
        ? resourcesPathActual
        : directorioDesarrollo;
    return path.join(base, 'assets', 'formatos-notas');
}

function obtenerDirectorioProcesador(
    appActual = app,
    resourcesPathActual = process.resourcesPath,
    directorioDesarrollo = __dirname
) {
    return appActual?.isPackaged
        ? path.join(resourcesPathActual, 'app.asar.unpacked')
        : directorioDesarrollo;
}

async function generarFormatoBgu3ConPython(payload, dependencias = {}) {
    const ejecutarPython = dependencias.runPython || runPython;
    const resultado = await ejecutarPython(['--generar-formato-bgu3'], payload);
    const datos = JSON.parse(resultado.stdout || '{}');
    if (!datos.success) throw new Error(datos.error || 'Generación fallida');
    return datos;
}

async function descargarFormatoNotasSeguro(formatoId, opciones = {}, dependencias = {}) {
    // Conserva compatibilidad con llamadas y pruebas anteriores que pasaban
    // las dependencias como segundo argumento.
    if (
        opciones
        && typeof opciones === 'object'
        && (opciones.fs || opciones.path || opciones.dialog || opciones.app)
    ) {
        dependencias = opciones;
        opciones = {};
    }
    if (
        typeof formatoId !== 'string'
        || !Object.prototype.hasOwnProperty.call(FORMATOS_NOTAS, formatoId)
    ) {
        return {
            success: false,
            cancelled: false,
            code: 'FORMATO_INVALIDO',
            error: MENSAJES_FORMATOS_NOTAS.invalido
        };
    }

    const formato = FORMATOS_NOTAS[formatoId];
    const fsActual = dependencias.fs || fs;
    const pathActual = dependencias.path || path;
    const dialogActual = dependencias.dialog || dialog;
    const appActual = dependencias.app || app;
    const ventanaActual = dependencias.mainWindow || mainWindow;
    const resourcesPathActual = dependencias.resourcesPath || process.resourcesPath;
    const directorio = dependencias.directorioFormatos || obtenerDirectorioFormatosNotas(
        appActual,
        resourcesPathActual,
        __dirname
    );
    let seleccionBgu3 = null;
    if (formatoId === 'bgu_3') {
        let catalogo;
        try {
            catalogo = dependencias.catalogoAsignaturas || obtenerCatalogoAsignaturasSeguro(
                dependencias.directorioAplicacion || __dirname,
                fsActual,
                pathActual
            );
        } catch (error) {
            registrarErrorSeguro('descargar-formato/catalogo', error);
            return {
                success: false,
                cancelled: false,
                code: 'OPTATIVAS_INVALIDAS',
                error: MENSAJES_FORMATOS_NOTAS.optativasInvalidas
            };
        }
        seleccionBgu3 = validarSeleccionOptativasBgu3(opciones.optativas, catalogo);
        if (!seleccionBgu3.valida) {
            return {
                success: false,
                cancelled: false,
                ...seleccionBgu3
            };
        }
    }
    const directorioResuelto = pathActual.resolve(directorio);
    const sourcePath = pathActual.resolve(directorioResuelto, formato.archivoInterno);
    const rutaRelativa = pathActual.relative(directorioResuelto, sourcePath);

    if (rutaRelativa.startsWith('..') || pathActual.isAbsolute(rutaRelativa)) {
        registrarErrorSeguro('descargar-formato/ruta', new Error('Ruta interna no autorizada.'));
        return {
            success: false,
            cancelled: false,
            code: 'FORMATO_INVALIDO',
            error: MENSAJES_FORMATOS_NOTAS.invalido
        };
    }

    try {
        const estado = fsActual.statSync(sourcePath);
        if (!estado.isFile()) throw new Error('El recurso no es un archivo regular.');
        fsActual.accessSync(sourcePath, fsActual.constants.R_OK);
    } catch (error) {
        registrarErrorSeguro('descargar-formato/recurso', error);
        return {
            success: false,
            cancelled: false,
            code: 'FORMATO_NO_DISPONIBLE',
            error: MENSAJES_FORMATOS_NOTAS.noDisponible
        };
    }

    let result;
    try {
        result = await dialogActual.showSaveDialog(ventanaActual, {
            title: formato.tituloDialogo,
            defaultPath: pathActual.join(appActual.getPath('downloads'), formato.nombreDescarga),
            filters: [
                { name: 'Archivo de Excel', extensions: ['xlsx'] }
            ]
        });
    } catch (error) {
        registrarErrorSeguro('descargar-formato/dialogo', error);
        return {
            success: false,
            cancelled: false,
            code: 'COPIA_FALLIDA',
            error: MENSAJES_FORMATOS_NOTAS.copiaFallida
        };
    }

    if (result.canceled || !result.filePath) {
        return { success: false, cancelled: true };
    }

    let directorioTemporal = null;
    try {
        if (pathActual.resolve(result.filePath) === sourcePath) {
            return {
                success: false,
                cancelled: false,
                code: 'DESTINO_NO_PERMITIDO',
                error: MENSAJES_FORMATOS_NOTAS.destinoNoPermitido
            };
        }
        let archivoACopiar = sourcePath;
        if (formatoId === 'bgu_3') {
            const baseTemporal = dependencias.directorioTemporalBase || appActual.getPath('temp');
            directorioTemporal = fsActual.mkdtempSync(
                pathActual.join(baseTemporal, 'certi-bgu3-')
            );
            archivoACopiar = pathActual.join(directorioTemporal, 'formato_3_bgu_generado.xlsx');
            const generarFormato = dependencias.generarFormatoBgu3 || generarFormatoBgu3ConPython;
            await generarFormato(
                {
                    origen: sourcePath,
                    destino: archivoACopiar,
                    optativas: seleccionBgu3.optativas
                },
                dependencias
            );
            const generado = fsActual.statSync(archivoACopiar);
            if (!generado.isFile() || generado.size === 0) {
                throw new Error('El generador no produjo un archivo XLSX válido.');
            }
        }
        fsActual.copyFileSync(archivoACopiar, result.filePath);
        return {
            success: true,
            cancelled: false,
            path: result.filePath,
            nombre: formato.nombreDescarga,
            formatoId,
            optativas: seleccionBgu3?.optativas || []
        };
    } catch (error) {
        registrarErrorSeguro('descargar-formato/copia', error);
        return {
            success: false,
            cancelled: false,
            code: formatoId === 'bgu_3' ? 'GENERACION_FALLIDA' : 'COPIA_FALLIDA',
            error: formatoId === 'bgu_3'
                ? MENSAJES_FORMATOS_NOTAS.generacionFallida
                : MENSAJES_FORMATOS_NOTAS.copiaFallida
        };
    } finally {
        if (directorioTemporal) {
            try {
                fsActual.rmSync(directorioTemporal, { recursive: true, force: true });
            } catch (error) {
                registrarErrorSeguro('descargar-formato/limpieza', error);
            }
        }
    }
}

if (app && BrowserWindow && ipcMain && dialog) {

function createWindow() {
    mainWindow = new BrowserWindow({
        width: 1280,
        height: 800,
        title: "UEEH • Sistema Académico - Boletines",
        webPreferences: {
            preload: path.join(__dirname, 'preload.js'),
            contextIsolation: true,
            nodeIntegration: false,
            sandbox: true,
            webSecurity: true,
            allowRunningInsecureContent: false,
            experimentalFeatures: false,
            devTools: !app.isPackaged
        }
    });

    endurecerVentana(mainWindow);
    mainWindow.webContents.once('did-finish-load', () => {
        mainWindow.webContents.send(
            'actualizaciones:estado',
            obtenerEstadoActualizacionPublico()
        );
    });
    mainWindow.loadFile('index.html');
}

app.whenReady().then(() => {
    if (session?.defaultSession) {
        session.defaultSession.setPermissionRequestHandler(
            (_webContents, _permission, callback) => callback(false)
        );
        session.defaultSession.setPermissionCheckHandler(() => false);
    }
    createWindow();
    programarActualizaciones();

    app.on('activate', function () {
        if (BrowserWindow.getAllWindows().length === 0) createWindow();
    });
});

app.on('window-all-closed', function () {
    if (process.platform !== 'darwin') app.quit();
});

// IPC: Estado, activación e inicio de la prueba. La validación permanece en main.
registrarManejadorIpcSeguro('actualizaciones:buscar', async () => {
    return comprobarActualizaciones({ manual: true });
});

registrarManejadorIpcSeguro('actualizaciones:descargar', async () => {
    return descargarActualizacion();
});

registrarManejadorIpcSeguro('actualizaciones:instalar', async () => {
    return instalarActualizacion();
});

registrarManejadorIpcSeguro('obtener-estado-licencia', async () => {
    const evaluacion = await obtenerEstadoLicenciaInterno();
    return convertirEstadoLicenciaPublico(evaluacion);
});

registrarManejadorIpcSeguro('activar-licencia', async (event, licenciaIngresada) => {
    if (!validarLicenciaIngresada(licenciaIngresada)) {
        const evaluacionActual = await obtenerEstadoLicenciaInterno({
            registrarEjecucion: false
        });
        return {
            ...convertirEstadoLicenciaPublico(evaluacionActual),
            success: false,
            licenciaValida: false,
            mensaje: 'La licencia ingresada no es válida.'
        };
    }
    return encolarOperacionLicencia(async () => {
        try {
            if (!safeStorage?.isEncryptionAvailable?.()) {
                throw new Error('El cifrado seguro del sistema no está disponible.');
            }
            const ahora = Date.now();
            const machineHash = obtenerHashMaquina();
            if (!estadoLicenciaMemoria) {
                estadoLicenciaMemoria = fusionarEstadosLicencia(
                    leerEstadosLicencia(),
                    machineHash
                );
            }
            if (!estadoLicenciaMemoria) {
                estadoLicenciaMemoria = crearEstadoPrueba(machineHash, ahora);
            }
            estadoLicenciaMemoria.activated = true;
            estadoLicenciaMemoria.lastRun = Math.max(estadoLicenciaMemoria.lastRun, ahora);
            guardarEstadoLicencia(estadoLicenciaMemoria);
            return convertirEstadoLicenciaPublico(
                evaluarEstadoLicencia(estadoLicenciaMemoria, ahora)
            );
        } catch (error) {
            registrarErrorSeguro('licencia/activar', error);
            return convertirEstadoLicenciaPublico({
                tipo: 'error',
                autorizado: false,
                activada: false,
                diasRestantes: null
            });
        }
    });
});

registrarManejadorIpcSeguro('iniciar-prueba', async () => {
    return encolarOperacionLicencia(async () => {
        try {
            if (!safeStorage?.isEncryptionAvailable?.()) {
                throw new Error('El cifrado seguro del sistema no está disponible.');
            }
            const ahora = Date.now();
            const machineHash = obtenerHashMaquina();
            if (!estadoLicenciaMemoria) {
                estadoLicenciaMemoria = fusionarEstadosLicencia(
                    leerEstadosLicencia(),
                    machineHash
                );
            }
            let evaluacion = evaluarEstadoLicencia(estadoLicenciaMemoria, ahora);
            if (evaluacion.tipo === 'no_iniciada') {
                estadoLicenciaMemoria = crearEstadoPrueba(machineHash, ahora);
                guardarEstadoLicencia(estadoLicenciaMemoria);
                evaluacion = evaluarEstadoLicencia(estadoLicenciaMemoria, ahora);
            } else if (evaluacion.tipo === 'prueba' || evaluacion.tipo === 'activada') {
                estadoLicenciaMemoria.lastRun = Math.max(
                    estadoLicenciaMemoria.lastRun,
                    ahora
                );
                guardarEstadoLicencia(estadoLicenciaMemoria);
            }
            return convertirEstadoLicenciaPublico(evaluacion);
        } catch (error) {
            registrarErrorSeguro('licencia/iniciar-prueba', error);
            return convertirEstadoLicenciaPublico({
                tipo: 'error',
                autorizado: false,
                activada: false,
                diasRestantes: null
            });
        }
    });
});

// IPC: Diálogo para seleccionar archivos (Logo o Excel)
registrarManejadorIpcSeguro('seleccionar-archivo', async (event, opciones = {}) => {
    const opcionesSeguras = normalizarOpcionesSeleccionArchivo(opciones);
    const result = await dialog.showOpenDialog(mainWindow, {
        properties: opcionesSeguras.properties,
        filters: opcionesSeguras.filters
    });
    if (result.canceled || result.filePaths.length === 0) {
        return null;
    }
    return result.filePaths[0];
});

// IPC: Descargar un formato Excel autorizado
registrarManejadorIpcSeguro('descargar-formato', async (event, formatoId, opciones = {}) => {
    if (descargaFormatoEnCurso) {
        return {
            success: false,
            cancelled: false,
            code: 'DESCARGA_EN_CURSO',
            error: MENSAJES_FORMATOS_NOTAS.enCurso
        };
    }

    descargaFormatoEnCurso = true;
    try {
        return await descargarFormatoNotasSeguro(formatoId, opciones);
    } finally {
        descargaFormatoEnCurso = false;
    }
});

// Helper para ejecutar Python
function runPython(args, inputData = null) {
    return new Promise((resolve, reject) => {
        const accionesPermitidas = new Set([
            '--generar-formato-bgu3',
            '--analizar',
            '--generar',
            '--certificados',
            '--supletorios'
        ]);
        if (
            !Array.isArray(args)
            || args.length === 0
            || args.length > 12
            || !accionesPermitidas.has(args[0])
            || args.some(argumento => typeof argumento !== 'string' || argumento.length > 4096)
        ) {
            reject({ success: false, error: 'Parámetros de procesamiento no autorizados.' });
            return;
        }

        const ejecutableConfigurado = process.env.CERTI_PYTHON_EXECUTABLE;
        const procesadorEmpaquetado = process.platform === 'win32'
            ? path.join(process.resourcesPath || '', 'python', 'certi-python.exe')
            : path.join(process.resourcesPath || '', 'python', 'certi-python');
        const usaEjecutableConfigurado = (
            typeof ejecutableConfigurado === 'string'
            && path.isAbsolute(ejecutableConfigurado)
            && fs.existsSync(ejecutableConfigurado)
        );
        const usaProcesadorEmpaquetado = (
            app.isPackaged
            && !usaEjecutableConfigurado
            && fs.existsSync(procesadorEmpaquetado)
        );
        const pythonExecutable = usaEjecutableConfigurado
            ? ejecutableConfigurado
            : (usaProcesadorEmpaquetado
                ? procesadorEmpaquetado
                : (process.platform === 'win32' ? 'python.exe' : 'python3'));
        const directorioProcesador = obtenerDirectorioProcesador();
        const scriptPath = path.join(directorioProcesador, 'procesador_notas.py');
        if (!usaProcesadorEmpaquetado && !fs.existsSync(scriptPath)) {
            reject({ success: false, error: 'El procesador local no está disponible.' });
            return;
        }

        const processArgs = usaProcesadorEmpaquetado ? args : [scriptPath, ...args];
        const pyProcess = spawn(pythonExecutable, processArgs, {
            cwd: directorioProcesador,
            shell: false,
            windowsHide: true,
            env: {
                ...process.env,
                PYTHONIOENCODING: 'utf-8'
            },
            stdio: ['pipe', 'pipe', 'pipe']
        });

        let stdoutData = '';
        let stderrData = '';
        let finalizado = false;
        const finalizar = (callback, valor) => {
            if (finalizado) return;
            finalizado = true;
            clearTimeout(timeout);
            callback(valor);
        };
        const timeout = setTimeout(() => {
            pyProcess.kill();
            finalizar(reject, {
                success: false,
                error: 'El procesamiento excedió el tiempo permitido.'
            });
        }, PYTHON_TIMEOUT_MS);

        if (inputData) {
            const serializado = JSON.stringify(inputData);
            if (Buffer.byteLength(serializado, 'utf8') > MAX_PAYLOAD_BYTES) {
                pyProcess.kill();
                finalizar(reject, {
                    success: false,
                    error: 'Los datos exceden el tamaño permitido.'
                });
                return;
            }
            pyProcess.stdin.write(serializado);
            pyProcess.stdin.end();
        }

        pyProcess.stdout.on('data', (data) => {
            stdoutData += data.toString('utf8');
            if (Buffer.byteLength(stdoutData, 'utf8') > MAX_PYTHON_OUTPUT_BYTES) {
                pyProcess.kill();
                finalizar(reject, {
                    success: false,
                    error: 'La respuesta del procesador excede el tamaño permitido.'
                });
            }
        });

        pyProcess.stderr.on('data', (data) => {
            stderrData += data.toString('utf8');
            if (Buffer.byteLength(stderrData, 'utf8') > MAX_PYTHON_OUTPUT_BYTES) {
                pyProcess.kill();
                finalizar(reject, {
                    success: false,
                    error: 'El procesador produjo una respuesta inválida.'
                });
            }
        });

        pyProcess.on('close', (code) => {
            if (code === 0) {
                finalizar(resolve, { success: true, stdout: stdoutData.trim() });
            } else {
                finalizar(reject, {
                    success: false,
                    code,
                    error: 'El procesador local finalizó con un error.'
                });
            }
        });

        pyProcess.on('error', (err) => {
            registrarErrorSeguro('python/inicio', err);
            finalizar(reject, {
                success: false,
                error: 'No se pudo iniciar el procesador local.'
            });
        });
    });
}

// IPC: Analizar archivos Excel para listar estudiantes
registrarManejadorIpcSeguro('analizar-excel', async (event, solicitud = {}) => {
    try {
        if (!esObjetoPlano(solicitud) || tamanoJson(solicitud) > 16384) {
            return { error: "La solicitud de análisis no es válida." };
        }
        const grado = typeof solicitud.grado === 'string' ? solicitud.grado.trim() : '';
        if (!grado || grado.length > 120 || /[\0-\x1f\x7f]/.test(grado)) {
            return { error: "No se recibió un grado válido del curso seleccionado." };
        }

        const args = ['--analizar', '--grado', grado];
        const periodos = ['t1', 't2', 't3', 'su'];
        const rutasValidas = {};

        for (const periodo of periodos) {
            const ruta = typeof solicitud[periodo] === 'string' ? solicitud[periodo].trim() : '';
            if (!ruta) continue;
            if (!validarArchivoExcel(ruta)) {
                return {
                    error: `El archivo seleccionado para ${periodo.toUpperCase()} no es un Excel válido, accesible o está fuera del límite de 25 MB.`
                };
            }
            rutasValidas[periodo] = ruta;
            args.push(`--${periodo}`, ruta);
        }

        if (Object.keys(rutasValidas).length === 0) {
            return { error: "Ruta inválida: no se recibió una ruta real del archivo Excel seleccionado." };
        }
        
        const result = await runPython(args);
        const parsed = JSON.parse(result.stdout);
        return parsed;
    } catch (error) {
        registrarErrorSeguro('analizar-excel', error);
        return { error: "No se pudo procesar el archivo de notas. Verifique que use el formato oficial." };
    }
});

// Catálogo de solo lectura compartido con la migración segura de datos locales.
registrarManejadorIpcSeguro('obtener-catalogo-asignaturas', async () => {
    try {
        const catalogoPath = path.join(__dirname, 'catalogo_asignaturas.json');
        return JSON.parse(fs.readFileSync(catalogoPath, 'utf-8'));
    } catch (error) {
        registrarErrorSeguro('catalogo-asignaturas', error);
        return [];
    }
});

registrarManejadorIpcSeguro('obtener-escala-cualitativa', async () => {
    try {
        const escalaPath = path.join(__dirname, 'escala_cualitativa.json');
        return JSON.parse(fs.readFileSync(escalaPath, 'utf-8'));
    } catch (error) {
        registrarErrorSeguro('escala-cualitativa', error);
        return { minimo: 1, maximo: 10, rangos: [] };
    }
});

// IPC: Generar Boletines (PDF)
registrarManejadorIpcSeguro('generar-boletines', async (event, datos) => {
    try {
        const errorValidacion = validarPayloadGeneracion(datos);
        if (errorValidacion) return { success: false, error: errorValidacion };

        const gradoArchivo = String(datos.institucion?.grado || 'Curso')
            .normalize('NFD')
            .replace(/[\u0300-\u036f]/g, '')
            .replace(/[^A-Za-z0-9_-]+/g, '_')
            .slice(0, 60) || 'Curso';
        const paraleloArchivo = String(datos.institucion?.paralelo || 'P')
            .replace(/[^A-Za-z0-9_-]+/g, '_')
            .slice(0, 20) || 'P';
        const destino = await dialog.showSaveDialog(mainWindow, {
            title: 'Guardar boletines como PDF',
            defaultPath: path.join(
                app.getPath('documents'),
                `Boletines_Consolidados_${gradoArchivo}_${paraleloArchivo}.pdf`
            ),
            filters: [{ name: 'Archivo PDF', extensions: ['pdf'] }]
        });
        if (destino.canceled || !destino.filePath) {
            return { success: false, cancelled: true };
        }
        if (path.extname(destino.filePath).toLowerCase() !== '.pdf') {
            return { success: false, error: 'El destino debe ser un archivo PDF.' };
        }
        datos.outputPath = destino.filePath;

        const appDataPath = app.getPath('userData');
        const logosDir = path.join(appDataPath, 'logos');
        if (!fs.existsSync(logosDir)) {
            fs.mkdirSync(logosDir, { recursive: true });
        }

        if (datos.logos) {
            const cursoId = datos.cursoActivoId || 'default';
            datos.logos.logo1 = guardarImagenBase64(datos.logos.logo1, 'logo1', cursoId, logosDir);
            datos.logos.logo2 = guardarImagenBase64(datos.logos.logo2, 'logo2', cursoId, logosDir);
        }

        // Ejecutamos Python pasando el payload JSON por stdin
        const result = await runPython(['--generar'], datos);
        return JSON.parse(result.stdout);
    } catch (error) {
        registrarErrorSeguro('generar-boletines', error);
        return { success: false, error: error.error || "Error de ejecución del generador de PDF" };
    }
});

// IPC: Generar certificados HTML iniciales y detectar supletorios
registrarManejadorIpcSeguro('generar-certificados', async (event, datos) => {
    try {
        const errorValidacion = validarPayloadGeneracion(datos);
        if (errorValidacion) return { success: false, error: errorValidacion };
        if (!PLANTILLAS_CERTIFICADO_PERMITIDAS.includes(datos.plantillaName)) {
            return { success: false, error: 'La plantilla solicitada no está autorizada.' };
        }

        const appDataPath = app.getPath('userData');
        const logosDir = path.join(appDataPath, 'logos');
        if (!fs.existsSync(logosDir)) {
            fs.mkdirSync(logosDir, { recursive: true });
        }

        // Crear directorio de salida para certificados generados (fuera de assets/certificados)
        const cursoId = datos.cursoActivoId || 'default';
        const certOutputDir = resolverRutaHija(
            path.join(appDataPath, 'certificados_temporales'),
            cursoId
        );
        if (!certOutputDir) {
            return { success: false, error: 'El identificador del curso no es válido.' };
        }
        if (!fs.existsSync(certOutputDir)) {
            fs.mkdirSync(certOutputDir, { recursive: true });
        }
        datos.certOutputDir = certOutputDir;

        if (datos.logos) {
            datos.logos.logo1 = guardarImagenBase64(datos.logos.logo1, 'logo1', cursoId, logosDir);
            datos.logos.logo2 = guardarImagenBase64(datos.logos.logo2, 'logo2', cursoId, logosDir);
        }

        // Ejecutamos Python con --certificados y pasamos el payload
        const result = await runPython(['--certificados'], datos);
        return JSON.parse(result.stdout);
    } catch (error) {
        registrarErrorSeguro('generar-certificados', error);
        return { success: false, error: error.error || "Error de ejecución al procesar certificados" };
    }
});

// IPC: Actualizar certificados con notas manuales de supletorio
registrarManejadorIpcSeguro('actualizar-supletorios', async (event, datos) => {
    try {
        let updates = [];
        let cursoId = 'default';
        if (esObjetoPlano(datos) && Array.isArray(datos.updates) && datos.cursoId) {
            updates = datos.updates;
            cursoId = datos.cursoId;
        } else if (Array.isArray(datos)) {
            updates = datos || [];
        }
        if (
            !esSegmentoRutaSeguro(cursoId, 100)
            || updates.length > 500
            || tamanoJson(updates) > 1024 * 1024
            || updates.some(actualizacion => (
                !esObjetoPlano(actualizacion)
                || !esSegmentoRutaSeguro(String(actualizacion.id || ''), 180)
                || typeof actualizacion.asignatura !== 'string'
                || actualizacion.asignatura.length > 200
                || !Number.isFinite(Number(actualizacion.nota_supletorio))
                || Number(actualizacion.nota_supletorio) < 0
                || Number(actualizacion.nota_supletorio) > 10
            ))
        ) {
            return { success: false, error: 'Los datos de supletorio no son válidos.' };
        }

        const appDataPath = app.getPath('userData');
        const certOutputDir = resolverRutaHija(
            path.join(appDataPath, 'certificados_temporales'),
            cursoId
        );
        if (!certOutputDir) {
            return { success: false, error: 'El identificador del curso no es válido.' };
        }

        // Inyectar cert_output_dir en cada update para Python
        const updatesWithDir = updates.map(u => ({
            ...u,
            cert_output_dir: certOutputDir
        }));

        const result = await runPython(['--supletorios'], updatesWithDir);
        return JSON.parse(result.stdout);
    } catch (error) {
        registrarErrorSeguro('actualizar-supletorios', error);
        return { success: false, error: error.error || "Error de ejecución al actualizar supletorios" };
    }
});


// IPC: Leer plantilla de certificado desde assets/certificados/
registrarManejadorIpcSeguro('leer-plantilla', async (event, nombreArchivo) => {
    try {
        if (
            typeof nombreArchivo !== 'string'
            || !PLANTILLAS_CERTIFICADO_PERMITIDAS.includes(nombreArchivo)
        ) {
            return { error: 'La plantilla solicitada no está autorizada.' };
        }
        const directorioPlantillas = path.join(__dirname, 'assets', 'certificados');
        const templatePath = resolverRutaHija(directorioPlantillas, nombreArchivo);
        if (!templatePath || !fs.existsSync(templatePath)) {
            return { error: 'La plantilla solicitada no está disponible.' };
        }
        const content = fs.readFileSync(templatePath, 'utf-8');
        if (Buffer.byteLength(content, 'utf8') > MAX_HTML_BYTES) {
            return { error: 'La plantilla excede el tamaño permitido.' };
        }
        return { content };
    } catch (error) {
        registrarErrorSeguro('leer-plantilla', error);
        return { error: "Error al leer la plantilla de certificado." };
    }
});

// IPC: Leer certificado generado desde userData/certificados_temporales/
registrarManejadorIpcSeguro('leer-certificado-generado', async (event, cursoId, nombreArchivo) => {
    try {
        if (
            !esSegmentoRutaSeguro(cursoId, 100)
            || !esSegmentoRutaSeguro(nombreArchivo, 220)
            || !nombreArchivo.startsWith('certificado_')
            || path.extname(nombreArchivo).toLowerCase() !== '.html'
        ) {
            return { error: 'El certificado solicitado no es válido.' };
        }
        const appDataPath = app.getPath('userData');
        const directorioCertificados = path.join(appDataPath, 'certificados_temporales');
        const filePath = resolverRutaHija(directorioCertificados, cursoId, nombreArchivo);
        if (!filePath || !fs.existsSync(filePath)) {
            return { error: 'El certificado generado no está disponible.' };
        }
        const content = fs.readFileSync(filePath, 'utf-8');
        if (Buffer.byteLength(content, 'utf8') > MAX_HTML_BYTES) {
            return { error: 'El certificado excede el tamaño permitido.' };
        }
        return { content };
    } catch (error) {
        registrarErrorSeguro('leer-certificado-generado', error);
        return { error: "Error al leer el certificado generado." };
    }
});

// IPC: Generar PDF de certificados desde HTML ensamblado
registrarManejadorIpcSeguro('imprimir-certificados', async (event, solicitud) => {
    let directorioTemporal = null;
    try {
        const normalizada = normalizarSolicitudImpresionCertificados(solicitud);
        const htmlContent = normalizada?.html;
        if (!validarHtmlRecibido(htmlContent)) {
            return { success: false, error: 'El contenido del certificado no es válido.' };
        }
        const destinoPrueba = obtenerDestinoPdfPruebaEmpaquetada();
        // Diálogo para elegir dónde guardar el PDF, salvo en la prueba empaquetada aislada.
        const saveResult = destinoPrueba
            ? { canceled: false, filePath: destinoPrueba }
            : await dialog.showSaveDialog(mainWindow, {
                title: 'Guardar Certificados como PDF',
                defaultPath: path.join(app.getPath('documents'), 'Certificados_UEEH.pdf'),
                filters: [{ name: 'Archivo PDF', extensions: ['pdf'] }]
            });

        if (saveResult.canceled || !saveResult.filePath) {
            return { success: false, cancelled: true };
        }
        if (path.extname(saveResult.filePath).toLowerCase() !== '.pdf') {
            return { success: false, error: 'El destino debe ser un archivo PDF.' };
        }

        directorioTemporal = fs.mkdtempSync(path.join(app.getPath('temp'), 'certi-pdf-'));
        const tempHtmlPath = path.join(directorioTemporal, 'certificados.html');
        fs.writeFileSync(tempHtmlPath, htmlContent, 'utf-8');

        // Crear ventana oculta para renderizar el HTML
        const printWindow = new BrowserWindow({
            show: false,
            width: 1024,
            height: 768,
            webPreferences: {
                nodeIntegration: false,
                contextIsolation: true,
                sandbox: true,
                webSecurity: true,
                allowRunningInsecureContent: false,
                experimentalFeatures: false,
                devTools: false
            }
        });
        endurecerVentana(printWindow);

        try {
            await printWindow.loadFile(tempHtmlPath);

            await esperarDocumentoListoParaPdf(
                printWindow.webContents,
                normalizada.diagnostico
            );

            // Generar el PDF
            const pdfBuffer = await printWindow.webContents.printToPDF({
                printBackground: true,
                landscape: false,
                pageSize: 'A4',
                margins: {
                    marginType: 'custom',
                    top: 0.4,
                    bottom: 0.4,
                    left: 0.2,
                    right: 0.2
                }
            });

            // Guardar el PDF en la ruta elegida por el usuario
            fs.writeFileSync(saveResult.filePath, pdfBuffer);
            return { success: true, path: saveResult.filePath };
        } finally {
            printWindow.close();
        }
    } catch (error) {
        registrarErrorSeguro('imprimir-certificados', error);
        if (error?.code === 'PDF_NOTAS_INCOMPLETAS') {
            return {
                success: false,
                error: 'No se generó el PDF porque faltan notas respecto de las cargadas en la interfaz. Revise los archivos del curso e inténtelo nuevamente.'
            };
        }
        if (error?.code === 'PDF_CERTIFICADOS_INCOMPLETOS') {
            return {
                success: false,
                error: 'No se generó el PDF porque faltan certificados de estudiantes seleccionados.'
            };
        }
        if (error?.code === 'PDF_DOCUMENTO_NO_LISTO') {
            return { success: false, error: error.message };
        }
        return { success: false, error: "Error al generar el PDF de certificados." };
    } finally {
        if (directorioTemporal) {
            try {
                fs.rmSync(directorioTemporal, { recursive: true, force: true });
            } catch (error) {
                registrarErrorSeguro('imprimir-certificados/limpieza', error);
            }
        }
    }
});

// IPC: Abrir vista previa de certificado en ventana nueva
registrarManejadorIpcSeguro('abrir-vista-previa-certificado', async (event, htmlContent) => {
    let directorioTemporal = null;
    try {
        if (!validarHtmlRecibido(htmlContent)) {
            return { success: false, error: 'El contenido de la vista previa no es válido.' };
        }
        directorioTemporal = fs.mkdtempSync(path.join(app.getPath('temp'), 'certi-preview-'));
        const tempHtmlPath = path.join(directorioTemporal, 'certificado.html');
        fs.writeFileSync(tempHtmlPath, htmlContent, 'utf-8');

        const previewWindow = new BrowserWindow({
            width: 900,
            height: 700,
            title: 'Vista Previa del Certificado',
            autoHideMenuBar: true,
            webPreferences: {
                nodeIntegration: false,
                contextIsolation: true,
                sandbox: true,
                webSecurity: true,
                allowRunningInsecureContent: false,
                experimentalFeatures: false,
                devTools: false
            }
        });
        endurecerVentana(previewWindow);

        await previewWindow.loadFile(tempHtmlPath);

        // Limpiar archivo temporal cuando la ventana se cierre
        previewWindow.on('closed', () => {
            try {
                fs.rmSync(directorioTemporal, { recursive: true, force: true });
            } catch (error) {
                registrarErrorSeguro('vista-previa/limpieza', error);
            }
        });

        return { success: true };
    } catch (error) {
        registrarErrorSeguro('abrir-vista-previa-certificado', error);
        if (directorioTemporal) {
            try {
                fs.rmSync(directorioTemporal, { recursive: true, force: true });
            } catch (errorLimpieza) {
                registrarErrorSeguro('vista-previa/limpieza', errorLimpieza);
            }
        }
        return { success: false, error: 'No se pudo abrir la vista previa.' };
    }
});
}

module.exports = {
    FORMATOS_NOTAS,
    MENSAJES_FORMATOS_NOTAS,
    CAPACIDAD_MATERIAS_FORMATO_BGU3,
    CANALES_IPC_PERMITIDOS,
    DURACION_PRUEBA_MS,
    RETRASO_ACTUALIZACION_INICIAL_MS,
    INTERVALO_ACTUALIZACIONES_MS,
    PDF_READY_TIMEOUT_MS,
    PLANTILLAS_CERTIFICADO_PERMITIDAS,
    normalizarNotasVersion,
    esVersionSuperior,
    esErrorSinConexion,
    obtenerEstadoActualizacionPublico,
    sanitizarValorLogActualizacion,
    registrarEventoActualizacion,
    comprobarActualizaciones,
    descargarActualizacion,
    instalarActualizacion,
    programarActualizaciones,
    calcularSha256,
    validarLicenciaIngresada,
    normalizarEstadoLicencia,
    crearEstadoPrueba,
    fusionarEstadosLicencia,
    evaluarEstadoLicencia,
    convertirEstadoLicenciaPublico,
    obtenerHashMaquina,
    descifrarEstadoLicencia,
    cifrarEstadoLicencia,
    leerEstadosLicencia,
    guardarEstadoLicencia,
    obtenerDirectorioFormatosNotas,
    obtenerDirectorioProcesador,
    validarSeleccionOptativasBgu3,
    validarArchivoExcel,
    validarHtmlRecibido,
    normalizarSolicitudImpresionCertificados,
    obtenerDestinoPdfPruebaEmpaquetada,
    esperarDocumentoListoParaPdf,
    validarPayloadGeneracion,
    esSegmentoRutaSeguro,
    resolverRutaHija,
    esSolicitudIpcAutorizada,
    normalizarOpcionesSeleccionArchivo,
    generarFormatoBgu3ConPython,
    descargarFormatoNotasSeguro
};
