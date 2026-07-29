(function (root, factory) {
    const api = factory();
    if (typeof module === 'object' && module.exports) module.exports = api;
    root.AsistenciaUtils = api;
})(typeof globalThis !== 'undefined' ? globalThis : this, function () {
    'use strict';

    const TRIMESTRES = ['T1', 'T2', 'T3'];
    const TIPOS_FALTA = new Set(['justificada', 'injustificada']);

    function crearTrimestreVacio() {
        return { fechaInicio: '', fechaFin: '', diasSinClases: [], estudiantes: {} };
    }

    function crearAsistenciaVacia() {
        return {
            T1: crearTrimestreVacio(),
            T2: crearTrimestreVacio(),
            T3: crearTrimestreVacio()
        };
    }

    function normalizarAsistencia(asistencia) {
        const salida = asistencia && typeof asistencia === 'object' ? asistencia : {};
        TRIMESTRES.forEach(clave => {
            const actual = salida[clave] && typeof salida[clave] === 'object' ? salida[clave] : {};
            salida[clave] = {
                fechaInicio: esFechaISOValida(actual.fechaInicio) ? actual.fechaInicio : '',
                fechaFin: esFechaISOValida(actual.fechaFin) ? actual.fechaFin : '',
                diasSinClases: Array.from(new Set(
                    (Array.isArray(actual.diasSinClases) ? actual.diasSinClases : [])
                        .filter(esFechaISOValida)
                )).sort(),
                estudiantes: actual.estudiantes && typeof actual.estudiantes === 'object'
                    ? actual.estudiantes
                    : {}
            };
        });
        return salida;
    }

    function esFechaISOValida(valor) {
        if (!/^\d{4}-\d{2}-\d{2}$/.test(String(valor || ''))) return false;
        const [anio, mes, dia] = valor.split('-').map(Number);
        const fecha = new Date(anio, mes - 1, dia, 12, 0, 0);
        return fecha.getFullYear() === anio
            && fecha.getMonth() === mes - 1
            && fecha.getDate() === dia;
    }

    function fechaDesdeISO(valor) {
        if (!esFechaISOValida(valor)) return null;
        const [anio, mes, dia] = valor.split('-').map(Number);
        return new Date(anio, mes - 1, dia, 12, 0, 0);
    }

    function isoDesdePartes(anio, mes, dia) {
        const valor = `${String(anio).padStart(4, '0')}-${String(mes).padStart(2, '0')}-${String(dia).padStart(2, '0')}`;
        return esFechaISOValida(valor) ? valor : '';
    }

    function siguienteFechaISO(valor) {
        const fecha = fechaDesdeISO(valor);
        if (!fecha) return '';
        fecha.setDate(fecha.getDate() + 1);
        return isoDesdePartes(fecha.getFullYear(), fecha.getMonth() + 1, fecha.getDate());
    }

    function formatoFecha(valor) {
        if (!esFechaISOValida(valor)) return '';
        const [anio, mes, dia] = valor.split('-');
        return `${dia}/${mes}/${anio}`;
    }

    function nombreDia(valor) {
        const fecha = fechaDesdeISO(valor);
        if (!fecha) return '';
        return ['Domingo', 'Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado'][fecha.getDay()];
    }

    function esFinDeSemana(valor) {
        const fecha = fechaDesdeISO(valor);
        if (!fecha) return false;
        return fecha.getDay() === 0 || fecha.getDay() === 6;
    }

    function rangosSeSuperponen(inicioA, finA, inicioB, finB) {
        if (![inicioA, finA, inicioB, finB].every(esFechaISOValida)) return false;
        return inicioA <= finB && inicioB <= finA;
    }

    function validarPeriodo(asistencia, trimestre, inicio, fin) {
        if (!TRIMESTRES.includes(trimestre)) return { valido: false, mensaje: 'Seleccione primero un trimestre.' };
        if (!esFechaISOValida(inicio)) return { valido: false, mensaje: 'La fecha inicial es obligatoria y debe ser válida.' };
        if (!esFechaISOValida(fin)) return { valido: false, mensaje: 'La fecha final es obligatoria y debe ser válida.' };
        if (inicio > fin) return { valido: false, mensaje: 'La fecha inicial no puede ser posterior a la fecha final.' };

        const datos = normalizarAsistencia(asistencia);
        for (const otraClave of TRIMESTRES) {
            if (otraClave === trimestre) continue;
            const otro = datos[otraClave];
            if (otro.fechaInicio && otro.fechaFin && rangosSeSuperponen(inicio, fin, otro.fechaInicio, otro.fechaFin)) {
                return {
                    valido: false,
                    mensaje: `El periodo se superpone con ${otraClave}.`
                };
            }
        }
        return { valido: true, mensaje: '' };
    }

    function esDiaLectivo(valor, trimestre) {
        if (!esFechaISOValida(valor) || !trimestre) return false;
        if (!esFechaISOValida(trimestre.fechaInicio) || !esFechaISOValida(trimestre.fechaFin)) return false;
        if (valor < trimestre.fechaInicio || valor > trimestre.fechaFin) return false;
        if (esFinDeSemana(valor)) return false;
        return !(trimestre.diasSinClases || []).includes(valor);
    }

    function contarDiasLectivos(trimestre) {
        if (!trimestre || !esFechaISOValida(trimestre.fechaInicio) || !esFechaISOValida(trimestre.fechaFin)) return 0;
        if (trimestre.fechaInicio > trimestre.fechaFin) return 0;
        let total = 0;
        let cursor = trimestre.fechaInicio;
        while (cursor && cursor <= trimestre.fechaFin) {
            if (esDiaLectivo(cursor, trimestre)) total += 1;
            cursor = siguienteFechaISO(cursor);
        }
        return total;
    }

    function normalizarFaltas(faltas, trimestre) {
        const salida = {};
        Object.entries(faltas && typeof faltas === 'object' ? faltas : {}).forEach(([fecha, registro]) => {
            if (!esDiaLectivo(fecha, trimestre)) return;
            if (!registro || !TIPOS_FALTA.has(registro.tipo)) return;
            salida[fecha] = {
                tipo: registro.tipo,
                observacion: String(registro.observacion || '').trim()
            };
        });
        return salida;
    }

    function resumenEstudiante(trimestre, estudianteId) {
        const configurado = !!(
            trimestre
            && esFechaISOValida(trimestre.fechaInicio)
            && esFechaISOValida(trimestre.fechaFin)
            && trimestre.fechaInicio <= trimestre.fechaFin
        );
        if (!configurado) {
            return {
                configurado: false,
                justificadas: 0,
                injustificadas: 0,
                totalFaltas: 0,
                diasLectivos: 0,
                totalAsistencia: 0
            };
        }

        const faltas = normalizarFaltas(
            trimestre.estudiantes?.[estudianteId]?.faltas || {},
            trimestre
        );
        const valores = Object.values(faltas);
        const justificadas = valores.filter(item => item.tipo === 'justificada').length;
        const injustificadas = valores.filter(item => item.tipo === 'injustificada').length;
        const totalFaltas = justificadas + injustificadas;
        const diasLectivos = contarDiasLectivos(trimestre);
        return {
            configurado: true,
            justificadas,
            injustificadas,
            totalFaltas,
            diasLectivos,
            totalAsistencia: Math.max(diasLectivos - totalFaltas, 0)
        };
    }

    function resumenAnual(asistencia, estudianteId) {
        const datos = normalizarAsistencia(asistencia);
        const resumenes = TRIMESTRES.map(clave => resumenEstudiante(datos[clave], estudianteId));
        const configurados = resumenes.filter(item => item.configurado);
        if (configurados.length === 0) {
            return {
                configurada: false,
                justificadas: null,
                injustificadas: null,
                totalFaltas: null,
                diasLectivos: null,
                totalAsistencia: null
            };
        }
        const suma = campo => configurados.reduce((total, item) => total + item[campo], 0);
        const justificadas = suma('justificadas');
        const injustificadas = suma('injustificadas');
        const totalFaltas = justificadas + injustificadas;
        const diasLectivos = suma('diasLectivos');
        return {
            configurada: true,
            justificadas,
            injustificadas,
            totalFaltas,
            diasLectivos,
            totalAsistencia: Math.max(diasLectivos - totalFaltas, 0)
        };
    }

    function faltasFueraDePeriodo(trimestre, inicio, fin) {
        const fechas = new Set();
        Object.values(trimestre?.estudiantes || {}).forEach(estudiante => {
            Object.keys(estudiante?.faltas || {}).forEach(fecha => {
                if (fecha < inicio || fecha > fin || esFinDeSemana(fecha)) fechas.add(fecha);
            });
        });
        return Array.from(fechas).sort();
    }

    return {
        TRIMESTRES,
        crearTrimestreVacio,
        crearAsistenciaVacia,
        normalizarAsistencia,
        esFechaISOValida,
        fechaDesdeISO,
        isoDesdePartes,
        siguienteFechaISO,
        formatoFecha,
        nombreDia,
        esFinDeSemana,
        rangosSeSuperponen,
        validarPeriodo,
        esDiaLectivo,
        contarDiasLectivos,
        normalizarFaltas,
        resumenEstudiante,
        resumenAnual,
        faltasFueraDePeriodo
    };
});
