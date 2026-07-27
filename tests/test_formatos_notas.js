'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const crypto = require('node:crypto');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const vm = require('node:vm');

const {
    FORMATOS_NOTAS,
    MENSAJES_FORMATOS_NOTAS,
    obtenerDirectorioFormatosNotas,
    descargarFormatoNotasSeguro
} = require('../main.js');

const raiz = path.resolve(__dirname, '..');
const recursos = path.join(raiz, 'assets', 'formatos-notas');
const indexSource = fs.readFileSync(path.join(raiz, 'index.html'), 'utf8');
const preloadSource = fs.readFileSync(path.join(raiz, 'preload.js'), 'utf8');
const packageJson = JSON.parse(fs.readFileSync(path.join(raiz, 'package.json'), 'utf8'));

const originalesPorId = {
    egb_2_7: 'FORMATO PARA 2.º a 7.º EGB.xlsx',
    egb_8_10: 'FORMATO PARA 8.º a 10.º EGB.xlsx',
    bgu_1_2: 'FORMATO PARA 1.º y 2.º BGU.xlsx',
    bgu_3: 'FORMATO PARA 3.º BGU.xlsx'
};

function sha256(ruta) {
    return crypto.createHash('sha256').update(fs.readFileSync(ruta)).digest('hex');
}

function extraerFuncion(nombre) {
    const inicio = indexSource.indexOf(`function ${nombre}(`);
    assert.ok(inicio >= 0, `No se encontró ${nombre}`);
    const inicioCuerpo = indexSource.indexOf('{', inicio);
    let profundidad = 0;
    for (let indice = inicioCuerpo; indice < indexSource.length; indice += 1) {
        if (indexSource[indice] === '{') profundidad += 1;
        if (indexSource[indice] === '}') profundidad -= 1;
        if (profundidad === 0) return indexSource.slice(inicio, indice + 1);
    }
    throw new Error(`No se pudo extraer ${nombre}`);
}

function crearDependencias(directorioFormatos, dialogo) {
    return {
        fs,
        path,
        directorioFormatos,
        app: {
            isPackaged: false,
            getPath: () => os.tmpdir()
        },
        mainWindow: {},
        dialog: dialogo
    };
}

test('los cuatro recursos conservan tamaño y hash del original', () => {
    assert.deepEqual(Object.keys(FORMATOS_NOTAS), ['egb_2_7', 'egb_8_10', 'bgu_1_2', 'bgu_3']);
    Object.entries(FORMATOS_NOTAS).forEach(([formatoId, formato]) => {
        const original = path.join(raiz, originalesPorId[formatoId]);
        const recurso = path.join(recursos, formato.archivoInterno);
        assert.ok(fs.statSync(original).size > 0);
        assert.equal(fs.statSync(recurso).size, fs.statSync(original).size);
        assert.equal(sha256(recurso), sha256(original));
    });
    assert.notEqual(
        sha256(path.join(recursos, FORMATOS_NOTAS.bgu_1_2.archivoInterno)),
        sha256(path.join(recursos, FORMATOS_NOTAS.bgu_3.archivoInterno))
    );
});

test('la lista blanca y sus entradas están congeladas', () => {
    assert.equal(Object.isFrozen(FORMATOS_NOTAS), true);
    Object.values(FORMATOS_NOTAS).forEach(formato => assert.equal(Object.isFrozen(formato), true));
    assert.equal(FORMATOS_NOTAS.bgu_1_2.archivoInterno, 'formato_1_y_2_bgu.xlsx');
    assert.equal(FORMATOS_NOTAS.bgu_3.archivoInterno, 'formato_3_bgu.xlsx');
});

test('las rutas de desarrollo y producción se resuelven desde la aplicación', () => {
    assert.equal(
        obtenerDirectorioFormatosNotas({ isPackaged: false }, 'R:\\recursos', 'D:\\app'),
        path.join('D:\\app', 'assets', 'formatos-notas')
    );
    assert.equal(
        obtenerDirectorioFormatosNotas({ isPackaged: true }, 'R:\\recursos', 'D:\\app'),
        path.join('R:\\recursos', 'assets', 'formatos-notas')
    );
});

test('cada ID descarga únicamente su recurso y sugiere el nombre visible correcto', async () => {
    const temporal = fs.mkdtempSync(path.join(os.tmpdir(), 'certi-formatos-'));
    try {
        for (const [formatoId, formato] of Object.entries(FORMATOS_NOTAS)) {
            const destino = path.join(temporal, formato.nombreDescarga);
            let opcionesDialogo;
            const dialogo = {
                showSaveDialog: async (_ventana, opciones) => {
                    opcionesDialogo = opciones;
                    return { canceled: false, filePath: destino };
                }
            };
            const resultado = await descargarFormatoNotasSeguro(
                formatoId,
                crearDependencias(recursos, dialogo)
            );
            assert.deepEqual(
                {
                    success: resultado.success,
                    cancelled: resultado.cancelled,
                    nombre: resultado.nombre,
                    formatoId: resultado.formatoId
                },
                {
                    success: true,
                    cancelled: false,
                    nombre: formato.nombreDescarga,
                    formatoId
                }
            );
            assert.equal(path.basename(opcionesDialogo.defaultPath), formato.nombreDescarga);
            assert.deepEqual(opcionesDialogo.filters, [
                { name: 'Archivo de Excel', extensions: ['xlsx'] }
            ]);
            assert.equal(sha256(destino), sha256(path.join(recursos, formato.archivoInterno)));
        }
    } finally {
        fs.rmSync(temporal, { recursive: true, force: true });
    }
});

test('todos los identificadores arbitrarios y traversal se rechazan antes del diálogo', async () => {
    const invalidos = [
        '../../archivo',
        '..\\..\\archivo',
        'C:\\Windows\\archivo.xlsx',
        'formato_inexistente',
        null,
        undefined,
        123,
        {},
        [],
        'formato_2_a_7_egb.xlsx'
    ];
    let dialogos = 0;
    const dialogo = {
        showSaveDialog: async () => {
            dialogos += 1;
            return { canceled: true };
        }
    };
    for (const valor of invalidos) {
        const resultado = await descargarFormatoNotasSeguro(
            valor,
            crearDependencias(recursos, dialogo)
        );
        assert.equal(resultado.code, 'FORMATO_INVALIDO');
        assert.equal(resultado.error, MENSAJES_FORMATOS_NOTAS.invalido);
    }
    assert.equal(dialogos, 0);
});

test('cancelar no copia ni produce un error', async () => {
    const resultado = await descargarFormatoNotasSeguro(
        'egb_2_7',
        crearDependencias(recursos, {
            showSaveDialog: async () => ({ canceled: true })
        })
    );
    assert.deepEqual(resultado, { success: false, cancelled: true });
});

test('un recurso faltante devuelve mensaje controlado y no abre diálogo', async () => {
    const temporal = fs.mkdtempSync(path.join(os.tmpdir(), 'certi-vacio-'));
    let dialogos = 0;
    try {
        const resultado = await descargarFormatoNotasSeguro(
            'bgu_3',
            crearDependencias(temporal, {
                showSaveDialog: async () => {
                    dialogos += 1;
                    return { canceled: true };
                }
            })
        );
        assert.equal(resultado.code, 'FORMATO_NO_DISPONIBLE');
        assert.equal(resultado.error, MENSAJES_FORMATOS_NOTAS.noDisponible);
        assert.equal(dialogos, 0);
    } finally {
        fs.rmSync(temporal, { recursive: true, force: true });
    }
});

test('un error de copia no expone ENOENT ni detalles técnicos', async () => {
    const destinoInvalido = path.join(os.tmpdir(), 'directorio-que-no-existe-certi', 'archivo.xlsx');
    const resultado = await descargarFormatoNotasSeguro(
        'egb_8_10',
        crearDependencias(recursos, {
            showSaveDialog: async () => ({ canceled: false, filePath: destinoInvalido })
        })
    );
    assert.equal(resultado.code, 'COPIA_FALLIDA');
    assert.equal(resultado.error, MENSAJES_FORMATOS_NOTAS.copiaFallida);
    assert.doesNotMatch(resultado.error, /ENOENT|EPERM|EACCES|stack/i);
});

test('el modal es único, accesible y contiene exactamente las cuatro tarjetas', () => {
    assert.equal((indexSource.match(/id="formatos-notas-modal"/g) || []).length, 1);
    assert.match(indexSource, /id="formatos-notas-modal"[\s\S]*?role="dialog"/);
    assert.match(indexSource, /aria-modal="true"/);
    assert.match(indexSource, /aria-labelledby="formatos-notas-modal-title"/);
    assert.match(indexSource, />\s*Descargar formatos\s*</);
    Object.keys(FORMATOS_NOTAS).forEach(formatoId => {
        assert.equal(
            (indexSource.match(new RegExp(`data-formato-card data-formato-id="${formatoId}"`, 'g')) || []).length,
            1
        );
    });
    assert.equal((indexSource.match(/<article data-formato-card/g) || []).length, 4);
    assert.match(indexSource, /event\.key !== 'Escape'/);
    assert.equal((indexSource.match(/document\.addEventListener\('keydown'/g) || []).length, 1);

    const ids = Array.from(indexSource.matchAll(/\sid="([^"]+)"/g), coincidencia => coincidencia[1]);
    const duplicados = ids.filter((id, indice) => ids.indexOf(id) !== indice);
    assert.deepEqual(duplicados, []);
});

test('todos los bloques JavaScript embebidos conservan sintaxis válida', () => {
    const bloques = Array.from(
        indexSource.matchAll(/<script(?:\s[^>]*)?>([\s\S]*?)<\/script>/gi),
        coincidencia => coincidencia[1]
    ).filter(codigo => codigo.trim());
    assert.ok(bloques.length > 0);
    bloques.forEach((codigo, indice) => {
        assert.doesNotThrow(
            () => new vm.Script(codigo, { filename: `index-inline-${indice}.js` })
        );
    });
});

test('la recomendación reutiliza la normalización real del curso activo', () => {
    const contexto = { cursoActual: null };
    contexto.obtenerCursoActivo = () => contexto.cursoActual;
    vm.createContext(contexto);
    vm.runInContext([
        extraerFuncion('normalizarTextoAsignatura'),
        extraerFuncion('codigoGradoCatalogo'),
        extraerFuncion('obtenerFormatoRecomendadoCursoActivo')
    ].join('\n'), contexto);

    const casos = {
        '5TO DE EGB': 'egb_2_7',
        '9NO DE EGB': 'egb_8_10',
        '2DO DE BACHILLERATO CIENCIAS': 'bgu_1_2',
        '2DO BGU': 'bgu_1_2',
        '3RO DE BACHILLERATO CIENCIAS': 'bgu_3',
        '3RO BGU': 'bgu_3',
        'INICIAL 1': null,
        'INICIAL 2': null,
        '1RO DE EGB': null
    };
    Object.entries(casos).forEach(([grado, esperado]) => {
        contexto.cursoActual = {
            nombreVisible: `${grado} A`,
            datosInstitucion: { gradoCurso: grado }
        };
        assert.equal(contexto.obtenerFormatoRecomendadoCursoActivo(), esperado, grado);
    });
});

test('el contrato preload solo transmite formatoId y conserva el aislamiento', () => {
    assert.match(
        preloadSource,
        /descargarFormato:\s*\(formatoId\)\s*=>\s*ipcRenderer\.invoke\('descargar-formato', formatoId\)/
    );
    assert.doesNotMatch(preloadSource, /\bfs\s*:/);
    assert.doesNotMatch(preloadSource, /\bpath\s*:/);
    assert.doesNotMatch(preloadSource, /ipcRenderer:\s*ipcRenderer/);
    const mainSource = fs.readFileSync(path.join(raiz, 'main.js'), 'utf8');
    assert.match(mainSource, /contextIsolation:\s*true/);
    assert.match(mainSource, /nodeIntegration:\s*false/);
});

test('electron-builder incluye los XLSX como extraResources fuera del ASAR', () => {
    assert.equal(packageJson.build.asar, true);
    assert.deepEqual(packageJson.build.extraResources, [{
        from: 'assets/formatos-notas',
        to: 'assets/formatos-notas',
        filter: ['*.xlsx']
    }]);
    assert.match(packageJson.scripts.build, /electron-builder --win nsis/);
    assert.match(packageJson.scripts['build:dir'], /electron-builder --dir/);
});
