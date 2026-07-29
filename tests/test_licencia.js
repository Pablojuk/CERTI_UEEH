const test = require('node:test');
const assert = require('node:assert/strict');
const crypto = require('node:crypto');
const fs = require('node:fs');
const path = require('node:path');

const {
    DURACION_PRUEBA_MS,
    crearEstadoPrueba,
    fusionarEstadosLicencia,
    evaluarEstadoLicencia,
    validarLicenciaIngresada,
    convertirEstadoLicenciaPublico,
    cifrarEstadoLicencia,
    descifrarEstadoLicencia
} = require('../main.js');

const DIA = 24 * 60 * 60 * 1000;
const MAQUINA_A = 'a'.repeat(64);
const MAQUINA_B = 'b'.repeat(64);

test('la primera instalación crea exactamente 15 días de prueba', () => {
    const ahora = Date.UTC(2026, 0, 1);
    const estado = crearEstadoPrueba(MAQUINA_A, ahora);
    assert.equal(DURACION_PRUEBA_MS, 15 * DIA);
    assert.equal(estado.trialEnd - estado.trialStart, 15 * DIA);
    assert.equal(evaluarEstadoLicencia(estado, ahora).diasRestantes, 15);
});

test('el reinicio y la reinstalación conservan el inicio más antiguo', () => {
    const inicio = Date.UTC(2026, 0, 1);
    const original = crearEstadoPrueba(MAQUINA_A, inicio);
    const copiaPosterior = crearEstadoPrueba(MAQUINA_A, inicio + 4 * DIA);
    const fusionado = fusionarEstadosLicencia([copiaPosterior, original], MAQUINA_A);
    assert.equal(fusionado.trialStart, inicio);
    assert.equal(fusionado.trialEnd, inicio + 15 * DIA);
    assert.equal(evaluarEstadoLicencia(fusionado, inicio + 5 * DIA).diasRestantes, 10);
});

test('un registro vencido impide reiniciar la prueba', () => {
    const inicio = Date.UTC(2026, 0, 1);
    const vencido = { ...crearEstadoPrueba(MAQUINA_A, inicio), expired: true };
    const nuevo = crearEstadoPrueba(MAQUINA_A, inicio + DIA);
    const fusionado = fusionarEstadosLicencia([nuevo, vencido], MAQUINA_A);
    assert.equal(fusionado.expired, true);
    assert.equal(evaluarEstadoLicencia(fusionado, inicio + 2 * DIA).tipo, 'vencida');
});

test('el vencimiento ocurre al completar exactamente los 15 días', () => {
    const inicio = Date.UTC(2026, 0, 1);
    const estado = crearEstadoPrueba(MAQUINA_A, inicio);
    assert.equal(evaluarEstadoLicencia(estado, estado.trialEnd - 1).tipo, 'prueba');
    assert.equal(evaluarEstadoLicencia(estado, estado.trialEnd).tipo, 'vencida');
});

test('un cambio de reloj hacia atrás bloquea y no aumenta días', () => {
    const inicio = Date.UTC(2026, 0, 1);
    const estado = {
        ...crearEstadoPrueba(MAQUINA_A, inicio),
        lastRun: inicio + 8 * DIA
    };
    const evaluacion = evaluarEstadoLicencia(estado, inicio + 3 * DIA);
    assert.equal(evaluacion.tipo, 'fecha_invalida');
    assert.equal(evaluacion.autorizado, false);
    assert.equal(evaluacion.diasRestantes, 7);
});

test('una licencia incorrecta se rechaza', () => {
    assert.equal(validarLicenciaIngresada('licencia-invalida'), false);
    assert.equal(validarLicenciaIngresada(''), false);
});

test('la comparación válida usa timingSafeEqual sin guardar la clave en la prueba', () => {
    const hashEsperado = Buffer.from(
        '09fa7b2c0e4a19c7dc686f28d1eac2db8ed1824ee8f97fb4861886b45f2baad9',
        'hex'
    );
    let comparacionConstanteInvocada = false;
    const cryptoControlado = {
        createHash: () => ({
            update: () => ({
                digest: () => Buffer.from(hashEsperado)
            })
        }),
        timingSafeEqual: (a, b) => {
            comparacionConstanteInvocada = true;
            return crypto.timingSafeEqual(a, b);
        }
    };
    assert.equal(validarLicenciaIngresada('candidata-de-prueba', cryptoControlado), true);
    assert.equal(comparacionConstanteInvocada, true);
});

test('la activación permanece autorizada aunque la prueba anterior haya vencido', () => {
    const inicio = Date.UTC(2026, 0, 1);
    const estado = {
        ...crearEstadoPrueba(MAQUINA_A, inicio),
        activated: true,
        expired: true
    };
    const evaluacion = evaluarEstadoLicencia(estado, inicio + 100 * DIA);
    assert.equal(evaluacion.tipo, 'activada');
    assert.equal(evaluacion.autorizado, true);
});

test('los registros de otro equipo no se reutilizan', () => {
    const estadoAjeno = crearEstadoPrueba(MAQUINA_A, Date.UTC(2026, 0, 1));
    assert.equal(fusionarEstadosLicencia([estadoAjeno], MAQUINA_B), null);
});

test('el estado expuesto al renderer no revela equipo ni fechas internas', () => {
    const publico = convertirEstadoLicenciaPublico({
        tipo: 'prueba',
        autorizado: true,
        activada: false,
        diasRestantes: 12
    });
    assert.equal(publico.diasRestantes, 12);
    assert.equal('machineHash' in publico, false);
    assert.equal('trialStart' in publico, false);
    assert.equal('lastRun' in publico, false);
});

test('el estado se cifra y descifra exclusivamente mediante safeStorage', () => {
    const estado = crearEstadoPrueba(MAQUINA_A, Date.UTC(2026, 0, 1));
    const seguroControlado = {
        encryptString: texto => Buffer.from([...Buffer.from(texto)].reverse()),
        decryptString: buffer => Buffer.from([...buffer].reverse()).toString('utf8')
    };
    const cifrado = cifrarEstadoLicencia(estado, seguroControlado);
    assert.equal(cifrado.includes(MAQUINA_A), false);
    assert.deepEqual(descifrarEstadoLicencia(cifrado, seguroControlado), estado);
});

test('preload expone solo los tres IPC específicos de licencia y NSIS conserva datos', () => {
    const raiz = path.resolve(__dirname, '..');
    const preload = fs.readFileSync(path.join(raiz, 'preload.js'), 'utf8');
    const html = fs.readFileSync(path.join(raiz, 'index.html'), 'utf8');
    const paquete = JSON.parse(fs.readFileSync(path.join(raiz, 'package.json'), 'utf8'));
    assert.match(preload, /obtenerEstadoLicencia:\s*\(\)\s*=>/);
    assert.match(preload, /activarLicencia:\s*\(licencia\)\s*=>/);
    assert.match(preload, /iniciarPrueba:\s*\(\)\s*=>/);
    assert.doesNotMatch(preload, /verificarLicencia/);
    assert.match(html, /id="licencia-overlay"/);
    assert.equal(paquete.build.nsis.deleteAppDataOnUninstall, false);
});
