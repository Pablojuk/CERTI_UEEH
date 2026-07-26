'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const asistencia = require('../asistencia.js');

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

test('la persistencia reutiliza el store actual y agrega asistencia dentro del curso', () => {
    const html = fs.readFileSync(path.join(raiz, 'index.html'), 'utf8');
    const ui = fs.readFileSync(path.join(raiz, 'asistencia_ui.js'), 'utf8');
    assert.match(html, /const DB_VERSION = 1/);
    assert.match(html, /const STORE_NAME = 'cursosState'/);
    assert.match(html, /asistencia: AsistenciaUtils\.crearAsistenciaVacia\(\)/);
    assert.match(ui, /curso\.asistencia = AsistenciaUtils\.normalizarAsistencia/);
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
