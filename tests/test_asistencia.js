'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const asistencia = require('../asistencia.js');
const {
    toggleCollapsibleSection,
    enriquecerEstudiantesConAsistencia,
    obtenerValoresAsistenciaDocumento,
    inyectarAsistenciaDocumento,
    actualizarFaltaPorClic
} = require('../asistencia_ui.js');

const raiz = path.resolve(__dirname, '..');

function periodoConDiasLectivos(inicio, cantidad) {
    let fin = inicio;
    const trimestre = {
        fechaInicio: inicio,
        fechaFin: inicio,
        diasSinClases: [],
        estudiantes: {}
    };
    while (true) {
        trimestre.fechaFin = fin;
        if (asistencia.contarDiasLectivos(trimestre) === cantidad) return trimestre;
        fin = asistencia.siguienteFechaISO(fin);
    }
}

test('septiembre de 2026 cuenta solo lunes a viernes', () => {
    const trimestre = {
        fechaInicio: '2026-09-01',
        fechaFin: '2026-09-30',
        diasSinClases: [],
        estudiantes: {}
    };
    assert.equal(asistencia.contarDiasLectivos(trimestre), 22);
    assert.equal(asistencia.esDiaLectivo('2026-09-05', trimestre), false);
    assert.equal(asistencia.esDiaLectivo('2026-09-06', trimestre), false);
    assert.equal(asistencia.esDiaLectivo('2026-09-07', trimestre), true);
});

test('valida rangos invertidos y periodos superpuestos', () => {
    const datos = asistencia.crearAsistenciaVacia();
    datos.T1.fechaInicio = '2026-09-01';
    datos.T1.fechaFin = '2026-09-30';
    assert.equal(
        asistencia.validarPeriodo(datos, 'T2', '2026-10-30', '2026-10-01').valido,
        false
    );
    const solapado = asistencia.validarPeriodo(datos, 'T2', '2026-09-20', '2026-10-20');
    assert.equal(solapado.valido, false);
    assert.match(solapado.mensaje, /T1/);
    assert.equal(
        asistencia.validarPeriodo(datos, 'T2', '2026-10-01', '2026-10-30').valido,
        true
    );
});

test('un día sin clases se excluye y no admite faltas', () => {
    const trimestre = {
        fechaInicio: '2026-09-01',
        fechaFin: '2026-09-30',
        diasSinClases: ['2026-09-07'],
        estudiantes: {
            alumno: {
                faltas: {
                    '2026-09-07': { tipo: 'injustificada', observacion: '' },
                    '2026-09-08': { tipo: 'justificada', observacion: 'Cita médica' }
                }
            }
        }
    };
    assert.equal(asistencia.contarDiasLectivos(trimestre), 21);
    assert.equal(asistencia.esDiaLectivo('2026-09-07', trimestre), false);
    assert.deepEqual(Object.keys(asistencia.normalizarFaltas(trimestre.estudiantes.alumno.faltas, trimestre)), ['2026-09-08']);
});

test('cada fecha tiene un solo estado y el cambio no duplica la falta', () => {
    const trimestre = periodoConDiasLectivos('2026-09-01', 22);
    trimestre.estudiantes.alumno = {
        faltas: {
            '2026-09-01': { tipo: 'justificada', observacion: '' }
        }
    };
    assert.deepEqual(asistencia.resumenEstudiante(trimestre, 'alumno'), {
        configurado: true,
        justificadas: 1,
        injustificadas: 0,
        totalFaltas: 1,
        diasLectivos: 22,
        totalAsistencia: 21
    });
    trimestre.estudiantes.alumno.faltas['2026-09-01'] = {
        tipo: 'injustificada',
        observacion: 'Sin justificación'
    };
    const resumen = asistencia.resumenEstudiante(trimestre, 'alumno');
    assert.equal(resumen.justificadas, 0);
    assert.equal(resumen.injustificadas, 1);
    assert.equal(resumen.totalFaltas, 1);
    assert.equal(resumen.totalAsistencia, 21);
});

test('un segundo clic con la misma marca desmarca el día', () => {
    const faltas = {};
    const fecha = '2026-09-09';

    assert.equal(actualizarFaltaPorClic(faltas, fecha, 'justificada'), true);
    assert.deepEqual(faltas[fecha], {
        tipo: 'justificada',
        observacion: ''
    });

    assert.equal(actualizarFaltaPorClic(faltas, fecha, 'justificada'), true);
    assert.equal(faltas[fecha], undefined);
});

test('cambiar el tipo de falta conserva la observación y borrar un día vacío no genera cambios', () => {
    const fecha = '2026-09-09';
    const faltas = {
        [fecha]: {
            tipo: 'justificada',
            observacion: 'Cita médica'
        }
    };

    assert.equal(actualizarFaltaPorClic(faltas, fecha, 'injustificada'), true);
    assert.deepEqual(faltas[fecha], {
        tipo: 'injustificada',
        observacion: 'Cita médica'
    });
    assert.equal(actualizarFaltaPorClic(faltas, '2026-09-10', 'borrar'), false);
});

test('el acumulado anual suma T1, T2 y T3 exactamente', () => {
    const datos = asistencia.crearAsistenciaVacia();
    datos.T1 = periodoConDiasLectivos('2026-01-05', 60);
    datos.T2 = periodoConDiasLectivos('2026-04-06', 65);
    datos.T3 = periodoConDiasLectivos('2026-08-03', 65);
    datos.T1.estudiantes.alumno = {
        faltas: {
            '2026-01-05': { tipo: 'justificada', observacion: '' },
            '2026-01-06': { tipo: 'injustificada', observacion: '' },
            '2026-01-07': { tipo: 'injustificada', observacion: '' }
        }
    };
    datos.T2.estudiantes.alumno = {
        faltas: {
            '2026-04-06': { tipo: 'justificada', observacion: '' },
            '2026-04-07': { tipo: 'justificada', observacion: '' },
            '2026-04-08': { tipo: 'injustificada', observacion: '' }
        }
    };
    datos.T3.estudiantes.alumno = {
        faltas: {
            '2026-08-03': { tipo: 'injustificada', observacion: '' }
        }
    };
    assert.deepEqual(asistencia.resumenAnual(datos, 'alumno'), {
        configurada: true,
        justificadas: 3,
        injustificadas: 4,
        totalFaltas: 7,
        diasLectivos: 190,
        totalAsistencia: 183
    });
    assert.deepEqual(asistencia.resumenAnual(datos, 'sin_faltas'), {
        configurada: true,
        justificadas: 0,
        injustificadas: 0,
        totalFaltas: 0,
        diasLectivos: 190,
        totalAsistencia: 190
    });
});

test('solo T1 aporta al acumulado y un alumno sin faltas conserva los días lectivos', () => {
    const datos = asistencia.crearAsistenciaVacia();
    datos.T1 = periodoConDiasLectivos('2026-09-01', 22);
    assert.deepEqual(asistencia.resumenAnual(datos, 'sin_faltas'), {
        configurada: true,
        justificadas: 0,
        injustificadas: 0,
        totalFaltas: 0,
        diasLectivos: 22,
        totalAsistencia: 22
    });
});

test('sin configuración no inventa valores', () => {
    assert.deepEqual(asistencia.resumenAnual(asistencia.crearAsistenciaVacia(), 'alumno'), {
        configurada: false,
        justificadas: null,
        injustificadas: null,
        totalFaltas: null,
        diasLectivos: null,
        totalAsistencia: null
    });
});

test('los cursos y trimestres permanecen aislados', () => {
    const cursoA = asistencia.crearAsistenciaVacia();
    const cursoB = asistencia.crearAsistenciaVacia();
    cursoA.T1 = periodoConDiasLectivos('2026-09-01', 5);
    cursoB.T1 = periodoConDiasLectivos('2026-09-01', 5);
    cursoA.T1.estudiantes.alumno = {
        faltas: { '2026-09-01': { tipo: 'justificada', observacion: '' } }
    };
    assert.equal(asistencia.resumenAnual(cursoA, 'alumno').totalFaltas, 1);
    assert.equal(asistencia.resumenAnual(cursoB, 'alumno').totalFaltas, 0);
    assert.equal(asistencia.resumenEstudiante(cursoA.T2, 'alumno').configurado, false);
});

test('el resumen para documentos conserva curso e id_real aunque se repita el nombre', () => {
    const asistenciaCurso = asistencia.crearAsistenciaVacia();
    asistenciaCurso.T1 = periodoConDiasLectivos('2026-09-01', 68);
    asistenciaCurso.T1.estudiantes.estudiante_a = {
        faltas: Object.fromEntries(
            ['01', '02', '03', '04', '07', '08'].map(dia => [
                `2026-09-${dia}`,
                { tipo: 'injustificada', observacion: '' }
            ])
        )
    };
    const curso = { id: 'curso_a', asistencia: asistenciaCurso };
    const anteriorUtils = global.AsistenciaUtils;
    const anteriorMostrar = global.mostrarValorSeguro;
    global.AsistenciaUtils = asistencia;
    global.mostrarValorSeguro = valor => valor == null ? '' : String(valor);
    try {
        const enriquecidos = enriquecerEstudiantesConAsistencia(curso, [
            { id_real: 'estudiante_a', nombre: 'NOMBRE REPETIDO' },
            { id_real: 'estudiante_b', nombre: 'NOMBRE REPETIDO' }
        ]);
        assert.deepEqual(enriquecidos[0].asistencia, {
            configurada: true,
            justificadas: 0,
            injustificadas: 6,
            totalFaltas: 6,
            diasLectivos: 68,
            totalAsistencia: 62,
            T1: {
                configurado: true,
                justificadas: 0,
                injustificadas: 6,
                totalFaltas: 6,
                diasLectivos: 68,
                totalAsistencia: 62
            },
            T2: {
                configurado: false,
                justificadas: 0,
                injustificadas: 0,
                totalFaltas: 0,
                diasLectivos: 0,
                totalAsistencia: 0
            },
            T3: {
                configurado: false,
                justificadas: 0,
                injustificadas: 0,
                totalFaltas: 0,
                diasLectivos: 0,
                totalAsistencia: 0
            },
            anual: {
                configurada: true,
                justificadas: 0,
                injustificadas: 6,
                totalFaltas: 6,
                diasLectivos: 68,
                totalAsistencia: 62
            },
            cursoId: 'curso_a',
            estudianteId: 'estudiante_a'
        });
        assert.equal(enriquecidos[1].asistencia.totalFaltas, 0);
        assert.equal(enriquecidos[1].asistencia.estudianteId, 'estudiante_b');
    } finally {
        global.AsistenciaUtils = anteriorUtils;
        global.mostrarValorSeguro = anteriorMostrar;
    }
});

test('el certificado recibe T1, T2, T3 y el anual exactos, incluidos ceros y filas vacías', () => {
    const asistenciaCurso = asistencia.crearAsistenciaVacia();
    asistenciaCurso.T1 = periodoConDiasLectivos('2026-01-05', 22);
    asistenciaCurso.T2 = periodoConDiasLectivos('2026-04-06', 22);
    asistenciaCurso.T3 = periodoConDiasLectivos('2026-08-03', 22);
    asistenciaCurso.T1.estudiantes.alumno = {
        faltas: {
            '2026-01-05': { tipo: 'injustificada', observacion: '' },
            '2026-01-06': { tipo: 'injustificada', observacion: '' },
            '2026-01-07': { tipo: 'injustificada', observacion: '' }
        }
    };
    asistenciaCurso.T2.estudiantes.alumno = {
        faltas: {
            '2026-04-06': { tipo: 'justificada', observacion: '' },
            '2026-04-07': { tipo: 'injustificada', observacion: '' }
        }
    };
    asistenciaCurso.T3.estudiantes.alumno = {
        faltas: {
            '2026-08-03': { tipo: 'injustificada', observacion: '' }
        }
    };
    const anteriorUtils = global.AsistenciaUtils;
    const anteriorMostrar = global.mostrarValorSeguro;
    global.AsistenciaUtils = asistencia;
    global.mostrarValorSeguro = valor => valor == null ? '' : String(valor);
    try {
        const [alumno, sinFaltas] = enriquecerEstudiantesConAsistencia(
            { id: 'curso_a', asistencia: asistenciaCurso },
            [
                { id_real: 'alumno', nombre: 'ALUMNO' },
                { id_real: 'sin_faltas', nombre: 'SIN FALTAS' }
            ]
        );
        assert.deepEqual(
            obtenerValoresAsistenciaDocumento(alumno.asistencia, 'alumno', 'curso_a'),
            {
                T1: { registro: 3, justificacion: 0, injustificado: 3, total: 19 },
                T2: { registro: 2, justificacion: 1, injustificado: 1, total: 20 },
                T3: { registro: 1, justificacion: 0, injustificado: 1, total: 21 },
                ANUAL: { registro: 6, justificacion: 1, injustificado: 5, total: 60 }
            }
        );
        assert.deepEqual(
            obtenerValoresAsistenciaDocumento(sinFaltas.asistencia, 'sin_faltas', 'curso_a').T1,
            { registro: 0, justificacion: 0, injustificado: 0, total: 22 }
        );

        asistenciaCurso.T3 = asistencia.crearTrimestreVacio();
        const [parcial] = enriquecerEstudiantesConAsistencia(
            { id: 'curso_a', asistencia: asistenciaCurso },
            [{ id_real: 'alumno', nombre: 'ALUMNO' }]
        );
        assert.deepEqual(
            obtenerValoresAsistenciaDocumento(parcial.asistencia, 'alumno', 'curso_a').T3,
            { registro: '', justificacion: '', injustificado: '', total: '' }
        );
        assert.deepEqual(
            obtenerValoresAsistenciaDocumento(parcial.asistencia, 'otro', 'curso_a').ANUAL,
            { registro: '', justificacion: '', injustificado: '', total: '' }
        );
        assert.deepEqual(
            obtenerValoresAsistenciaDocumento(parcial.asistencia, 'alumno', 'otro_curso').T1,
            { registro: '', justificacion: '', injustificado: '', total: '' }
        );
    } finally {
        global.AsistenciaUtils = anteriorUtils;
        global.mostrarValorSeguro = anteriorMostrar;
    }
});

test('la vista previa llena las mismas dieciséis celdas trimestrales y anuales', () => {
    const resumen = {
        cursoId: 'curso_a',
        estudianteId: 'alumno',
        T1: { configurado: true, totalFaltas: 3, justificadas: 0, injustificadas: 3, totalAsistencia: 19 },
        T2: { configurado: true, totalFaltas: 2, justificadas: 1, injustificadas: 1, totalAsistencia: 20 },
        T3: { configurado: true, totalFaltas: 1, justificadas: 0, injustificadas: 1, totalAsistencia: 21 },
        anual: { configurada: true, totalFaltas: 6, justificadas: 1, injustificadas: 5, totalAsistencia: 60 }
    };
    const celdas = [];
    ['T1', 'T2', 'T3', 'ANUAL'].forEach(periodo => {
        ['registro', 'justificacion', 'injustificado', 'total'].forEach(campo => {
            celdas.push({
                dataset: {
                    asistenciaPeriodo: periodo,
                    asistenciaCampo: campo
                },
                textContent: ''
            });
        });
    });
    const doc = {
        querySelectorAll(selector) {
            return selector === '[data-asistencia-periodo][data-asistencia-campo]'
                ? celdas
                : [];
        }
    };
    const anteriorMostrar = global.mostrarValorSeguro;
    global.mostrarValorSeguro = valor => valor == null ? '' : String(valor);
    try {
        inyectarAsistenciaDocumento(
            doc,
            { id_real: 'alumno', asistencia: resumen },
            'curso_a'
        );
        assert.deepEqual(
            celdas.map(celda => celda.textContent),
            ['3', '0', '3', '19', '2', '1', '1', '20', '1', '0', '1', '21', '6', '1', '5', '60']
        );
    } finally {
        global.mostrarValorSeguro = anteriorMostrar;
    }
});

test('acepta el 29 de febrero únicamente en años bisiestos', () => {
    assert.equal(asistencia.esFechaISOValida('2024-02-29'), true);
    assert.equal(asistencia.esFechaISOValida('2025-02-29'), false);
    assert.equal(asistencia.formatoFecha('2024-02-29'), '29/02/2024');
});

test('detecta faltas que quedarían fuera al modificar un periodo', () => {
    const trimestre = {
        fechaInicio: '2026-09-01',
        fechaFin: '2026-09-30',
        diasSinClases: [],
        estudiantes: {
            a: { faltas: { '2026-09-02': { tipo: 'justificada', observacion: '' } } },
            b: { faltas: { '2026-09-25': { tipo: 'injustificada', observacion: '' } } }
        }
    };
    assert.deepEqual(
        asistencia.faltasFueraDePeriodo(trimestre, '2026-09-10', '2026-09-30'),
        ['2026-09-02']
    );
});

test('la opción de asistencia está debajo de Gestión de Estudiantes y tiene vista propia', () => {
    const html = fs.readFileSync(path.join(raiz, 'index.html'), 'utf8');
    const gestion = html.indexOf('Gestión de Estudiantes');
    const asistenciaModulo = html.indexOf('id="asistencia-modulo"');
    assert.ok(gestion >= 0);
    assert.ok(asistenciaModulo > gestion);
    assert.match(html, /Registrar asistencia/);
    assert.match(html, /Seleccione un curso para registrar la asistencia de sus estudiantes/);
    assert.match(html, /id="asistencia-trimestre"/);
    assert.match(html, /id="asistencia-calendario"/);
    assert.match(html, /id="asistencia-buscar"/);
});

test('las tres secciones comparten cabeceras plegables accesibles e identificadores únicos', () => {
    const html = fs.readFileSync(path.join(raiz, 'index.html'), 'utf8');
    const acordeones = [
        ['institucion-menu-btn', 'institucion-content', 'institucion-menu-icon'],
        ['gestion-estudiantes-menu-btn', 'gestion-estudiantes-content', 'gestion-estudiantes-menu-icon'],
        ['asistencia-menu-btn', 'asistencia-view', 'asistencia-menu-icon']
    ];

    acordeones.forEach(([botonId, contenidoId, iconoId]) => {
        assert.match(
            html,
            new RegExp(`<button id="${botonId}"[^>]*[\\s\\S]*?aria-expanded="false"[^>]*aria-controls="${contenidoId}"`)
        );
        assert.match(html, new RegExp(`<div id="${contenidoId}" class="hidden`));
        assert.match(html, new RegExp(`id="${iconoId}"[^>]*fa-chevron-down[^>]*transition-transform`));
        [botonId, contenidoId, iconoId].forEach(id => {
            assert.equal((html.match(new RegExp(`id="${id}"`, 'g')) || []).length, 1);
        });
    });
});

test('el acordeón conserva estado, flecha e independencia después de diez aperturas y cierres', () => {
    const crearClases = (...iniciales) => {
        const clases = new Set(iniciales);
        return {
            contains: clase => clases.has(clase),
            add: clase => clases.add(clase),
            remove: clase => clases.delete(clase),
            toggle(clase, forzar) {
                const activo = forzar === undefined ? !clases.has(clase) : forzar;
                if (activo) clases.add(clase);
                else clases.delete(clase);
                return activo;
            }
        };
    };
    const crearAcordeon = prefijo => {
        const atributos = new Map([['aria-expanded', 'false']]);
        return {
            contenido: {
                id: `${prefijo}-content`,
                classList: crearClases('hidden'),
                style: {},
                scrollHeight: 1200
            },
            icono: {
                id: `${prefijo}-icon`,
                classList: crearClases()
            },
            cabecera: {
                id: `${prefijo}-button`,
                classList: crearClases(),
                getAttribute: nombre => atributos.get(nombre) ?? null,
                setAttribute: (nombre, valor) => atributos.set(nombre, valor)
            }
        };
    };

    const institucion = crearAcordeon('institucion');
    const gestion = crearAcordeon('gestion');
    const elementos = new Map([
        [institucion.contenido.id, institucion.contenido],
        [institucion.icono.id, institucion.icono],
        [institucion.cabecera.id, institucion.cabecera],
        [gestion.contenido.id, gestion.contenido],
        [gestion.icono.id, gestion.icono],
        [gestion.cabecera.id, gestion.cabecera]
    ]);
    const documentoAnterior = global.document;
    global.document = { getElementById: id => elementos.get(id) || null };

    try {
        for (let clic = 1; clic <= 10; clic += 1) {
            const abierto = clic % 2 === 1;
            assert.equal(
                toggleCollapsibleSection(
                    institucion.contenido.id,
                    institucion.icono.id,
                    institucion.cabecera.id
                ),
                abierto
            );
            assert.equal(institucion.cabecera.getAttribute('aria-expanded'), String(abierto));
            assert.equal(institucion.contenido.classList.contains('hidden'), !abierto);
            assert.equal(institucion.icono.classList.contains('rotate-180'), abierto);
            assert.equal(gestion.cabecera.getAttribute('aria-expanded'), 'false');
        }
    } finally {
        global.document = documentoAnterior;
    }
});

test('la persistencia reutiliza el store actual y agrega asistencia dentro del curso', () => {
    const html = fs.readFileSync(path.join(raiz, 'index.html'), 'utf8');
    const ui = fs.readFileSync(path.join(raiz, 'asistencia_ui.js'), 'utf8');
    assert.match(html, /const DB_VERSION = 1/);
    assert.match(html, /const STORE_NAME = 'cursosState'/);
    assert.match(html, /asistencia: AsistenciaUtils\.crearAsistenciaVacia\(\)/);
    assert.match(ui, /curso\.asistencia = AsistenciaUtils\.normalizarAsistencia/);
});

test('el renderer propaga metadatos y excluye materias no elegibles de supletorios', () => {
    const html = fs.readFileSync(path.join(raiz, 'index.html'), 'utf8');
    assert.match(html, /metadatosAsignaturas:\s*estudianteExcel\.metadatos_asignaturas/);
    assert.match(html, /m\.tipo !== 'cualitativa' && m\.permite_supletorio !== false/);
    assert.match(html, /sub\.tipo === 'cualitativa' \|\| sub\.permite_supletorio === false/);
    assert.match(html, /!metadatosAsignaturaCatalogo\(subName\)\.permite_supletorio/);
    assert.match(html, /asignaturas generales y \$\{optativas\} asignaturas optativas/);
});

test('Electron mantiene aislamiento de contexto y no expone APIs peligrosas', () => {
    const main = fs.readFileSync(path.join(raiz, 'main.js'), 'utf8');
    const preload = fs.readFileSync(path.join(raiz, 'preload.js'), 'utf8');
    assert.match(main, /contextIsolation:\s*true/);
    assert.match(main, /nodeIntegration:\s*false/);
    assert.doesNotMatch(preload, /ipcRenderer:\s*ipcRenderer/);
    assert.doesNotMatch(preload, /\bfs\s*:/);
    assert.doesNotMatch(preload, /child_process/);
});
