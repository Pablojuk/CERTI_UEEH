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
    CAPACIDAD_MATERIAS_FORMATO_BGU3,
    obtenerDirectorioFormatosNotas,
    validarSeleccionOptativasBgu3,
    descargarFormatoNotasSeguro
} = require('../main.js');

const raiz = path.resolve(__dirname, '..');
const recursos = path.join(raiz, 'assets', 'formatos-notas');
const indexSource = fs.readFileSync(path.join(raiz, 'index.html'), 'utf8');
const preloadSource = fs.readFileSync(path.join(raiz, 'preload.js'), 'utf8');
const packageJson = JSON.parse(fs.readFileSync(path.join(raiz, 'package.json'), 'utf8'));
const catalogoAsignaturas = JSON.parse(
    fs.readFileSync(path.join(raiz, 'catalogo_asignaturas.json'), 'utf8')
);
const optativasBgu3 = catalogoAsignaturas
    .filter(entrada => entrada.es_optativa_bgu3 === true)
    .sort((a, b) => a.orden - b.orden)
    .map(entrada => entrada.nombre);

const hashesEsperadosPorId = {
    egb_2_7: 'be04d082453ad9fc4cf90113fa4909fba556380e09c1e3053ecda685739bb037',
    egb_8_10: '38e7148cf0925266d9a1900f06b816e15313f4cd5300564556f44a74ad884f67',
    bgu_1_2: 'b095f77f65a8373cb55c2c0168fa17f1bf6c8bf77fd3e05c27f38da0e1be42dc',
    bgu_3: 'e6a84c5b9acfb451eab0ff838f52c4303a225beb9d9ab3672bc2b49a0d8be753'
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

function crearDependencias(directorioFormatos, dialogo, adicionales = {}) {
    return {
        fs,
        path,
        directorioFormatos,
        directorioAplicacion: raiz,
        catalogoAsignaturas,
        app: {
            isPackaged: false,
            getPath: tipo => tipo === 'temp' ? os.tmpdir() : os.tmpdir()
        },
        mainWindow: {},
        dialog: dialogo,
        ...adicionales
    };
}

test('los cuatro recursos empaquetados conservan sus hashes conocidos', () => {
    assert.deepEqual(Object.keys(FORMATOS_NOTAS), ['egb_2_7', 'egb_8_10', 'bgu_1_2', 'bgu_3']);
    Object.entries(FORMATOS_NOTAS).forEach(([formatoId, formato]) => {
        const recurso = path.join(recursos, formato.archivoInterno);
        assert.ok(fs.statSync(recurso).size > 0);
        assert.equal(sha256(recurso), hashesEsperadosPorId[formatoId]);
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
                formatoId === 'bgu_3'
                    ? { cursoId: 'curso_bgu3', optativas: optativasBgu3.slice(0, 3) }
                    : {},
                crearDependencias(recursos, dialogo, {
                    generarFormatoBgu3: async payload => {
                        fs.copyFileSync(payload.origen, payload.destino);
                    }
                })
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
            {},
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
            { cursoId: 'curso_bgu3', optativas: optativasBgu3.slice(0, 3) },
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
    assert.equal((indexSource.match(/id="optativas-bgu3-modal"/g) || []).length, 1);
    assert.match(indexSource, /id="optativas-bgu3-modal"[\s\S]*?aria-modal="true"/);
    assert.match(indexSource, /id="btn-generar-formato-bgu3"[^>]*disabled/);
    assert.match(indexSource, /entrada\?\.es_optativa_bgu3 === true/);
    assert.match(indexSource, /Number\(a\.orden \|\| 999\) - Number\(b\.orden \|\| 999\)/);
    assert.match(indexSource, /curso\.optativasFormatoBgu3 = estado\.seleccion\.slice\(\)/);
    assert.match(indexSource, /await saveStateToDB\(\)/);
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

test('el contrato preload transmite opciones validadas sin exponer APIs peligrosas', () => {
    assert.match(
        preloadSource,
        /descargarFormato:\s*\(formatoId,\s*opciones = \{\}\)\s*=>\s*ipcRenderer\.invoke\('descargar-formato', formatoId, opciones\)/
    );
    assert.doesNotMatch(preloadSource, /\bfs\s*:/);
    assert.doesNotMatch(preloadSource, /\bpath\s*:/);
    assert.doesNotMatch(preloadSource, /ipcRenderer:\s*ipcRenderer/);
    const mainSource = fs.readFileSync(path.join(raiz, 'main.js'), 'utf8');
    assert.match(mainSource, /contextIsolation:\s*true/);
    assert.match(mainSource, /nodeIntegration:\s*false/);
});

test('3.º BGU exige optativas, las ordena por catálogo y respeta la capacidad de 16 materias', () => {
    assert.equal(CAPACIDAD_MATERIAS_FORMATO_BGU3, 16);
    const vacia = validarSeleccionOptativasBgu3([], catalogoAsignaturas);
    assert.equal(vacia.valida, false);
    assert.equal(vacia.code, 'OPTATIVAS_REQUERIDAS');

    const seleccionDesordenada = [
        optativasBgu3[7],
        optativasBgu3[0],
        optativasBgu3[3]
    ];
    const tres = validarSeleccionOptativasBgu3(seleccionDesordenada, catalogoAsignaturas);
    assert.equal(tres.valida, true);
    assert.deepEqual(tres.optativas, [
        optativasBgu3[0],
        optativasBgu3[3],
        optativasBgu3[7]
    ]);
    assert.equal(tres.totalMaterias, 14);

    const exceso = validarSeleccionOptativasBgu3(
        optativasBgu3.slice(0, 6),
        catalogoAsignaturas
    );
    assert.equal(exceso.valida, false);
    assert.equal(exceso.code, 'CAPACIDAD_EXCEDIDA');
});

test('la descarga personalizada usa una copia temporal, conserva el original y limpia al finalizar', async () => {
    const temporal = fs.mkdtempSync(path.join(os.tmpdir(), 'certi-bgu3-descarga-'));
    const destino = path.join(temporal, 'formato_personalizado.xlsx');
    const hashOriginal = sha256(path.join(recursos, FORMATOS_NOTAS.bgu_3.archivoInterno));
    let temporalGenerado = null;
    try {
        const resultado = await descargarFormatoNotasSeguro(
            'bgu_3',
            {
                cursoId: 'curso_bgu3',
                optativas: optativasBgu3.slice(0, 3)
            },
            crearDependencias(
                recursos,
                {
                    showSaveDialog: async () => ({
                        canceled: false,
                        filePath: destino
                    })
                },
                {
                    directorioTemporalBase: temporal,
                    generarFormatoBgu3: async payload => {
                        temporalGenerado = payload.destino;
                        fs.copyFileSync(payload.origen, payload.destino);
                    }
                }
            )
        );
        assert.equal(resultado.success, true);
        assert.deepEqual(resultado.optativas, optativasBgu3.slice(0, 3));
        assert.equal(sha256(destino), hashOriginal);
        assert.equal(
            sha256(path.join(recursos, FORMATOS_NOTAS.bgu_3.archivoInterno)),
            hashOriginal
        );
        assert.equal(fs.existsSync(temporalGenerado), false);
        assert.equal(fs.readdirSync(temporal).includes('formato_personalizado.xlsx'), true);
    } finally {
        fs.rmSync(temporal, { recursive: true, force: true });
    }
});

test('la selección guardada se recupera por curso y el resto de formatos sigue directo', () => {
    const contexto = {
        catalogoAsignaturas,
        cursoA: { id: 'curso_a', optativasFormatoBgu3: optativasBgu3.slice(0, 2) },
        cursoB: { id: 'curso_b', optativasFormatoBgu3: optativasBgu3.slice(4, 5) }
    };
    vm.createContext(contexto);
    vm.runInContext([
        extraerFuncion('obtenerOptativasBgu3Disponibles'),
        extraerFuncion('obtenerSeleccionOptativasBgu3Curso')
    ].join('\n'), contexto);
    assert.deepEqual(
        Array.from(contexto.obtenerSeleccionOptativasBgu3Curso(contexto.cursoA)),
        optativasBgu3.slice(0, 2)
    );
    assert.deepEqual(
        Array.from(contexto.obtenerSeleccionOptativasBgu3Curso(contexto.cursoB)),
        optativasBgu3.slice(4, 5)
    );
    assert.match(
        extraerFuncion('descargarFormatoNotas'),
        /formatoId === 'bgu_3'[\s\S]*abrirModalOptativasBgu3\(\)[\s\S]*return;[\s\S]*ejecutarDescargaFormatoNotas\(formatoId\)/
    );
});

test('el botón Generar y descargar permanece bloqueado sin selección', () => {
    const boton = { disabled: false };
    const estado = {
        textContent: '',
        classList: { toggle() {} }
    };
    const contexto = {
        catalogoAsignaturas,
        CAPACIDAD_MATERIAS_FORMATO_BGU3: 16,
        descargaFormatoNotasEnCurso: false,
        seleccion: [],
        document: {
            querySelectorAll: () => contexto.seleccion.map(value => ({ value })),
            getElementById: id => id === 'btn-generar-formato-bgu3' ? boton : estado
        }
    };
    vm.createContext(contexto);
    vm.runInContext([
        extraerFuncion('obtenerOptativasBgu3Disponibles'),
        extraerFuncion('cantidadMateriasFijasBgu3'),
        extraerFuncion('seleccionOptativasBgu3Actual'),
        extraerFuncion('actualizarEstadoOptativasBgu3')
    ].join('\n'), contexto);

    const vacia = contexto.actualizarEstadoOptativasBgu3();
    assert.equal(vacia.valida, false);
    assert.equal(boton.disabled, true);

    contexto.seleccion = optativasBgu3.slice(0, 3);
    const tres = contexto.actualizarEstadoOptativasBgu3();
    assert.equal(tres.valida, true);
    assert.equal(tres.total, 14);
    assert.equal(boton.disabled, false);
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
