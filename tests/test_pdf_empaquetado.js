'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

const raiz = path.resolve(__dirname, '..');
const fuenteMain = fs.readFileSync(path.join(raiz, 'main.js'), 'utf8');
const fuenteIndex = fs.readFileSync(path.join(raiz, 'index.html'), 'utf8');
const fuentePreload = fs.readFileSync(path.join(raiz, 'preload.js'), 'utf8');
const {
    esperarDocumentoListoParaPdf,
    normalizarSolicitudImpresionCertificados,
    obtenerDestinoPdfPruebaEmpaquetada,
    registrarEventoActualizacion
} = require('../main.js');

test('la solicitud de impresión exige el conteo esperado y una plantilla autorizada', () => {
    const solicitud = normalizarSolicitudImpresionCertificados({
        html: '<html></html>',
        diagnostico: {
            notasEsperadas: 48,
            estudiantesEsperados: 1,
            cursoId: 'curso_bgu2_a',
            plantilla: 'FORMATO DE 1 Y 2 DE BGU.html'
        }
    });
    assert.equal(solicitud.diagnostico.notasEsperadas, 48);
    assert.equal(normalizarSolicitudImpresionCertificados({ html: '<html></html>' }), null);
    assert.equal(normalizarSolicitudImpresionCertificados({
        html: '<html></html>',
        diagnostico: {
            notasEsperadas: -1,
            estudiantesEsperados: 1,
            cursoId: 'curso',
            plantilla: 'FORMATO DE 1 Y 2 DE BGU.html'
        }
    }), null);
});

test('la impresión espera DOM, fuentes, imágenes y la marca explícita', async () => {
    let consultas = 0;
    let reloj = 0;
    const webContents = {
        executeJavaScript: async () => {
            consultas += 1;
            if (consultas === 1) {
                return {
                    readyState: 'complete',
                    marca: false,
                    fuentes: true,
                    imagenes: true,
                    paginas: 1,
                    notas: 48
                };
            }
            return {
                readyState: 'complete',
                marca: true,
                fuentes: true,
                imagenes: true,
                paginas: 1,
                notas: 48
            };
        }
    };
    const estado = await esperarDocumentoListoParaPdf(
        webContents,
        { notasEsperadas: 48, estudiantesEsperados: 1 },
        {
            ahora: () => reloj,
            esperar: async milisegundos => { reloj += milisegundos; },
            timeoutMs: 1000,
            intervaloMs: 10
        }
    );
    assert.equal(consultas, 2);
    assert.equal(estado.notas, 48);
    assert.doesNotMatch(fuenteMain, /setTimeout\(resolve\s*=>\s*resolve\([^)]*\),\s*2500\)/);
});

test('la impresión se detiene si el HTML tiene menos notas que la interfaz', async () => {
    const webContents = {
        executeJavaScript: async () => ({
            readyState: 'complete',
            marca: true,
            fuentes: true,
            imagenes: true,
            paginas: 1,
            notas: 11
        })
    };
    await assert.rejects(
        esperarDocumentoListoParaPdf(
            webContents,
            { notasEsperadas: 48, estudiantesEsperados: 1 }
        ),
        error => error.code === 'PDF_NOTAS_INCOMPLETAS'
    );
});

test('el destino E2E solo acepta PDF dentro de temp y una app empaquetada', () => {
    const dependencias = {
        app: { isPackaged: true, getPath: () => 'C:\\temp-certi' },
        fs: { existsSync: () => true },
        path: path.win32,
        env: {
            CERTI_MODO_PRUEBA_EMPAQUETADA: '1',
            CERTI_PRUEBA_PDF_DESTINO: 'C:\\temp-certi\\resultado.pdf'
        }
    };
    assert.equal(
        obtenerDestinoPdfPruebaEmpaquetada(dependencias),
        'C:\\temp-certi\\resultado.pdf'
    );
    assert.throws(
        () => obtenerDestinoPdfPruebaEmpaquetada({
            ...dependencias,
            env: {
                ...dependencias.env,
                CERTI_PRUEBA_PDF_DESTINO: 'C:\\datos\\fuera.pdf'
            }
        }),
        /no es válido/
    );
    assert.equal(obtenerDestinoPdfPruebaEmpaquetada({
        ...dependencias,
        env: {}
    }), null);
});

test('renderer y preload transmiten diagnóstico y una marca PDF controlada', () => {
    assert.match(fuenteIndex, /window\.__CERTI_PDF_READY__\s*=\s*true/);
    assert.match(fuenteIndex, /contarNotasEsperadasCertificados\(estudiantesData, tDoc\)/);
    assert.match(fuenteIndex, /imprimirCertificados\(\{[\s\S]*diagnostico:/);
    assert.match(
        fuentePreload,
        /imprimirCertificados:\s*\(solicitud\)\s*=>\s*ipcRenderer\.invoke\('imprimir-certificados', solicitud\)/
    );
});

test('el logger del actualizador persiste en userData y redacta secretos', () => {
    const operaciones = [];
    const fsFalso = {
        mkdirSync: (...argumentos) => operaciones.push(['mkdir', ...argumentos]),
        appendFileSync: (...argumentos) => operaciones.push(['append', ...argumentos])
    };
    assert.equal(registrarEventoActualizacion(
        'error',
        {
            versionInstalada: '1.1.5',
            proveedor: 'github:Pablojuk/CERTI_UEEH',
            mensaje: 'falló https://example.test/latest.yml?token=secreto'
        },
        {
            app: { getPath: () => 'C:\\datos-certi', isPackaged: true },
            fs: fsFalso,
            path: path.win32
        }
    ), true);
    const append = operaciones.find(operacion => operacion[0] === 'append');
    assert.equal(append[1], 'C:\\datos-certi\\logs\\actualizaciones.log');
    assert.doesNotMatch(append[2], /secreto/);
    assert.match(append[2], /\[REDACTADO\]/);
});
