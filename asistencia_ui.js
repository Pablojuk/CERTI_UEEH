'use strict';

function asegurarAsistenciaCurso(curso) {
    if (!curso) return null;
    curso.asistencia = AsistenciaUtils.normalizarAsistencia(curso.asistencia);
    return curso.asistencia;
}

async function persistirAsistencia() {
    const guardado = await saveStateToDB();
    if (!guardado) {
        mostrarAlerta(
            'No se pudo guardar',
            'La información permanece visible, pero no pudo confirmarse su almacenamiento permanente. Intente guardar nuevamente.',
            'error'
        );
        return false;
    }
    return true;
}

function abrirModuloAsistencia() {
    const vista = document.getElementById('asistencia-view');
    const boton = document.getElementById('asistencia-menu-btn');
    const icono = document.getElementById('asistencia-menu-icon');
    const abrir = vista.classList.contains('hidden');
    vista.classList.toggle('hidden', !abrir);
    boton.classList.toggle('bg-indigo-50', abrir);
    boton.classList.toggle('border-indigo-200', abrir);
    icono.className = `fa-solid ${abrir ? 'fa-chevron-up' : 'fa-chevron-down'} text-slate-400`;
    if (abrir) renderModuloAsistencia();
}

function cambiarTrimestreAsistencia() {
    trimestreAsistenciaActivo = document.getElementById('asistencia-trimestre').value;
    renderModuloAsistencia();
}

function obtenerEstudianteIdAsistencia(estudiante) {
    return mostrarValorSeguro(estudiante?.id_real || estudiante?.estudianteId || estudiante?.cedula);
}

function resumenAsistenciaAnualCurso(curso, estudianteId) {
    return AsistenciaUtils.resumenAnual(
        curso ? asegurarAsistenciaCurso(curso) : null,
        estudianteId
    );
}

function enriquecerEstudiantesConAsistencia(curso, estudiantes) {
    return estudiantes.map(estudiante => ({
        ...estudiante,
        asistencia: resumenAsistenciaAnualCurso(curso, obtenerEstudianteIdAsistencia(estudiante))
    }));
}

function renderModuloAsistencia() {
    const vista = document.getElementById('asistencia-view');
    if (!vista || vista.classList.contains('hidden')) return;

    const curso = obtenerCursoActivo();
    document.getElementById('asistencia-sin-curso').classList.toggle('hidden', !!curso);
    document.getElementById('asistencia-con-curso').classList.toggle('hidden', !curso);
    if (!curso) return;

    const trimestre = asegurarAsistenciaCurso(curso)[trimestreAsistenciaActivo];
    document.getElementById('asistencia-trimestre').value = trimestreAsistenciaActivo;
    document.getElementById('asistencia-curso-nombre').textContent = curso.nombreVisible || 'Curso sin nombre';
    document.getElementById('asistencia-anio-lectivo').textContent =
        curso.datosInstitucion?.anioLectivo ? `Año lectivo: ${curso.datosInstitucion.anioLectivo}` : '';
    document.getElementById('asistencia-fecha-inicio').value = trimestre.fechaInicio || '';
    document.getElementById('asistencia-fecha-fin').value = trimestre.fechaFin || '';
    document.getElementById('asistencia-rango-guardado').textContent =
        trimestre.fechaInicio && trimestre.fechaFin
            ? `${AsistenciaUtils.formatoFecha(trimestre.fechaInicio)} – ${AsistenciaUtils.formatoFecha(trimestre.fechaFin)}`
            : 'Sin configurar';
    renderDiasSinClases();
    renderTablaAsistencia();
}

async function guardarPeriodoAsistencia() {
    const curso = obtenerCursoActivo();
    if (!curso) {
        mostrarAlerta('Sin curso', 'Seleccione un curso para registrar la asistencia de sus estudiantes.', 'warning');
        return;
    }
    const asistencia = asegurarAsistenciaCurso(curso);
    const inicio = document.getElementById('asistencia-fecha-inicio').value;
    const fin = document.getElementById('asistencia-fecha-fin').value;
    const validacion = AsistenciaUtils.validarPeriodo(asistencia, trimestreAsistenciaActivo, inicio, fin);
    if (!validacion.valido) {
        mostrarAlerta('Periodo inválido', validacion.mensaje, 'warning');
        return;
    }

    const trimestre = asistencia[trimestreAsistenciaActivo];
    let registrosFuera = 0;
    Object.values(trimestre.estudiantes || {}).forEach(estudiante => {
        Object.keys(estudiante?.faltas || {}).forEach(fecha => {
            if (fecha < inicio || fecha > fin || AsistenciaUtils.esFinDeSemana(fecha)) registrosFuera += 1;
        });
    });
    const diasSinClaseFuera = (trimestre.diasSinClases || []).filter(fecha => fecha < inicio || fecha > fin);
    if (registrosFuera > 0 || diasSinClaseFuera.length > 0) {
        const aceptar = confirm(
            `El nuevo periodo dejaría fuera ${registrosFuera} registro(s) de falta y ${diasSinClaseFuera.length} día(s) sin clases. ` +
            '¿Desea eliminar exclusivamente esos registros fuera del nuevo rango?'
        );
        if (!aceptar) return;
    }

    trimestre.fechaInicio = inicio;
    trimestre.fechaFin = fin;
    trimestre.diasSinClases = (trimestre.diasSinClases || []).filter(
        fecha => fecha >= inicio && fecha <= fin && !AsistenciaUtils.esFinDeSemana(fecha)
    );
    Object.values(trimestre.estudiantes || {}).forEach(estudiante => {
        estudiante.faltas = AsistenciaUtils.normalizarFaltas(estudiante?.faltas || {}, trimestre);
    });
    if (!await persistirAsistencia()) return;
    renderModuloAsistencia();
    mostrarAlerta('Periodo guardado', 'Periodo del trimestre guardado correctamente.', 'success');
}

async function agregarDiaSinClase() {
    const curso = obtenerCursoActivo();
    if (!curso) {
        mostrarAlerta('Sin curso', 'Seleccione primero un curso.', 'warning');
        return;
    }
    const trimestre = asegurarAsistenciaCurso(curso)[trimestreAsistenciaActivo];
    const fecha = document.getElementById('asistencia-dia-sin-clase').value;
    if (!trimestre.fechaInicio || !trimestre.fechaFin) {
        mostrarAlerta('Periodo pendiente', 'Configure las fechas del trimestre antes de agregar días sin clases.', 'warning');
        return;
    }
    if (!AsistenciaUtils.esFechaISOValida(fecha) || fecha < trimestre.fechaInicio || fecha > trimestre.fechaFin) {
        mostrarAlerta('Fecha inválida', 'El día sin clases debe estar dentro del periodo seleccionado.', 'warning');
        return;
    }
    if (AsistenciaUtils.esFinDeSemana(fecha)) {
        mostrarAlerta('Fecha no lectiva', 'Los sábados y domingos ya están excluidos de los días lectivos.', 'warning');
        return;
    }
    if ((trimestre.diasSinClases || []).includes(fecha)) {
        mostrarAlerta('Fecha existente', 'Este día ya fue configurado como día sin clases.', 'warning');
        return;
    }

    let faltasEnFecha = 0;
    Object.values(trimestre.estudiantes || {}).forEach(estudiante => {
        if (estudiante?.faltas?.[fecha]) faltasEnFecha += 1;
    });
    if (faltasEnFecha > 0) {
        const aceptar = confirm(
            `Existen ${faltasEnFecha} falta(s) registradas en esta fecha. ` +
            'Para convertirla en día sin clases deben eliminarse. ¿Desea continuar?'
        );
        if (!aceptar) return;
        Object.values(trimestre.estudiantes || {}).forEach(estudiante => {
            if (estudiante?.faltas) delete estudiante.faltas[fecha];
        });
    }

    trimestre.diasSinClases.push(fecha);
    trimestre.diasSinClases.sort();
    document.getElementById('asistencia-dia-sin-clase').value = '';
    if (!await persistirAsistencia()) return;
    renderModuloAsistencia();
    mostrarAlerta('Día sin clases guardado', 'La fecha fue excluida de los días lectivos.', 'success');
}

async function eliminarDiaSinClase(fecha) {
    const curso = obtenerCursoActivo();
    if (!curso) return;
    const trimestre = asegurarAsistenciaCurso(curso)[trimestreAsistenciaActivo];
    trimestre.diasSinClases = (trimestre.diasSinClases || []).filter(item => item !== fecha);
    if (!await persistirAsistencia()) return;
    renderModuloAsistencia();
}

function renderDiasSinClases() {
    const contenedor = document.getElementById('asistencia-dias-sin-clases-lista');
    contenedor.innerHTML = '';
    const curso = obtenerCursoActivo();
    if (!curso) return;
    const trimestre = asegurarAsistenciaCurso(curso)[trimestreAsistenciaActivo];
    if ((trimestre.diasSinClases || []).length === 0) {
        const vacio = document.createElement('span');
        vacio.className = 'text-xs text-slate-400';
        vacio.textContent = 'No hay días sin clases configurados.';
        contenedor.appendChild(vacio);
        return;
    }
    trimestre.diasSinClases.forEach(fecha => {
        const etiqueta = document.createElement('span');
        etiqueta.className = 'inline-flex items-center gap-2 bg-slate-200 text-slate-700 rounded-full px-3 py-1 text-xs';
        etiqueta.append(AsistenciaUtils.formatoFecha(fecha));
        const boton = document.createElement('button');
        boton.type = 'button';
        boton.className = 'text-slate-500 hover:text-rose-600';
        boton.setAttribute('aria-label', `Eliminar ${AsistenciaUtils.formatoFecha(fecha)}`);
        boton.innerHTML = '<i class="fa-solid fa-xmark"></i>';
        boton.onclick = () => eliminarDiaSinClase(fecha);
        etiqueta.appendChild(boton);
        contenedor.appendChild(etiqueta);
    });
}

function renderTablaAsistencia() {
    const cuerpo = document.getElementById('asistencia-estudiantes-body');
    if (!cuerpo) return;
    cuerpo.innerHTML = '';
    const curso = obtenerCursoActivo();
    if (!curso) return;
    const trimestre = asegurarAsistenciaCurso(curso)[trimestreAsistenciaActivo];
    const termino = normalizarNombre(document.getElementById('asistencia-buscar')?.value || '');
    const estudiantes = [...estudiantesCargados]
        .sort((a, b) => String(a.nombre || '').localeCompare(String(b.nombre || ''), 'es'))
        .filter(estudiante => {
            if (!termino) return true;
            const texto = normalizarNombre(
                `${estudiante.nombre || ''} ${estudiante.cedula || ''} ${estudiante.id_real || ''}`
            );
            return texto.includes(termino);
        });

    let totalJustificadas = 0;
    let totalInjustificadas = 0;
    estudiantesCargados.forEach(estudiante => {
        const resumen = AsistenciaUtils.resumenEstudiante(
            trimestre,
            obtenerEstudianteIdAsistencia(estudiante)
        );
        totalJustificadas += resumen.justificadas;
        totalInjustificadas += resumen.injustificadas;
    });
    document.getElementById('asistencia-resumen-estudiantes').textContent = estudiantesCargados.length;
    document.getElementById('asistencia-resumen-dias').textContent = AsistenciaUtils.contarDiasLectivos(trimestre);
    document.getElementById('asistencia-resumen-justificadas').textContent = totalJustificadas;
    document.getElementById('asistencia-resumen-injustificadas').textContent = totalInjustificadas;
    document.getElementById('asistencia-resumen-total').textContent = totalJustificadas + totalInjustificadas;

    if (estudiantes.length === 0) {
        const fila = cuerpo.insertRow();
        const celda = fila.insertCell();
        celda.colSpan = 8;
        celda.className = 'p-8 text-center text-slate-400';
        celda.textContent = estudiantesCargados.length
            ? 'No hay estudiantes que coincidan con la búsqueda.'
            : 'No se encontraron estudiantes en el curso seleccionado.';
        return;
    }

    estudiantes.forEach((estudiante, indice) => {
        const estudianteId = obtenerEstudianteIdAsistencia(estudiante);
        const resumen = AsistenciaUtils.resumenEstudiante(trimestre, estudianteId);
        const fila = cuerpo.insertRow();
        fila.className = 'hover:bg-slate-50';
        [
            String(indice + 1),
            estudiante.nombre || '',
            estudiante.cedula || estudianteId || 'S/C',
            String(resumen.justificadas),
            String(resumen.injustificadas),
            String(resumen.totalFaltas),
            resumen.configurado ? String(resumen.totalAsistencia) : ''
        ].forEach((texto, celdaIndice) => {
            const celda = fila.insertCell();
            celda.className = `p-3 ${celdaIndice >= 3 ? 'text-center' : ''}`;
            celda.textContent = texto;
        });
        const accion = fila.insertCell();
        accion.className = 'p-3 text-center';
        const boton = document.createElement('button');
        boton.type = 'button';
        boton.className = 'bg-indigo-50 text-indigo-700 border border-indigo-200 hover:bg-indigo-100 px-3 py-1.5 rounded text-xs font-semibold';
        boton.innerHTML = '<i class="fa-solid fa-calendar-days mr-1"></i> Gestionar faltas';
        boton.onclick = () => abrirModalAsistencia(estudianteId);
        accion.appendChild(boton);
    });
}

function abrirModalAsistencia(estudianteId) {
    const curso = obtenerCursoActivo();
    if (!curso) {
        mostrarAlerta('Sin curso', 'Seleccione primero un curso.', 'warning');
        return;
    }
    const estudiante = estudiantesCargados.find(item => obtenerEstudianteIdAsistencia(item) === estudianteId);
    if (!estudiante) {
        mostrarAlerta('Estudiante no encontrado', 'El estudiante no pertenece al curso activo.', 'warning');
        return;
    }
    const trimestre = asegurarAsistenciaCurso(curso)[trimestreAsistenciaActivo];
    if (!trimestre.fechaInicio || !trimestre.fechaFin) {
        mostrarAlerta('Periodo pendiente', 'Configure las fechas del trimestre antes de registrar faltas.', 'warning');
        return;
    }
    const faltas = AsistenciaUtils.normalizarFaltas(
        trimestre.estudiantes?.[estudianteId]?.faltas || {},
        trimestre
    );
    const fechaInicial = AsistenciaUtils.fechaDesdeISO(trimestre.fechaInicio);
    asistenciaModalEstado = {
        cursoId: curso.id,
        trimestre: trimestreAsistenciaActivo,
        estudianteId,
        faltas: JSON.parse(JSON.stringify(faltas)),
        modo: 'justificada',
        anio: fechaInicial.getFullYear(),
        mes: fechaInicial.getMonth() + 1,
        cambios: false
    };
    document.getElementById('asistencia-modal-estudiante').textContent = estudiante.nombre || '';
    document.getElementById('asistencia-modal-contexto').textContent =
        `${curso.nombreVisible} · ${trimestreAsistenciaActivo} · ` +
        `${AsistenciaUtils.formatoFecha(trimestre.fechaInicio)} – ${AsistenciaUtils.formatoFecha(trimestre.fechaFin)}`;
    prepararSelectoresCalendario(trimestre);
    seleccionarModoAsistencia('justificada');
    renderCalendarioAsistencia();
    renderListaFaltasAsistencia();
    document.getElementById('asistencia-modal').classList.remove('hidden');
}

function prepararSelectoresCalendario(trimestre) {
    const meses = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'];
    const selectorMes = document.getElementById('asistencia-cal-mes');
    selectorMes.innerHTML = '';
    meses.forEach((nombre, indice) => {
        const opcion = document.createElement('option');
        opcion.value = String(indice + 1);
        opcion.textContent = nombre;
        selectorMes.appendChild(opcion);
    });
    const selectorAnio = document.getElementById('asistencia-cal-anio');
    selectorAnio.innerHTML = '';
    const inicioAnio = Number(trimestre.fechaInicio.slice(0, 4));
    const finAnio = Number(trimestre.fechaFin.slice(0, 4));
    for (let anio = inicioAnio; anio <= finAnio; anio += 1) {
        const opcion = document.createElement('option');
        opcion.value = String(anio);
        opcion.textContent = String(anio);
        selectorAnio.appendChild(opcion);
    }
    selectorMes.value = String(asistenciaModalEstado.mes);
    selectorAnio.value = String(asistenciaModalEstado.anio);
}

function seleccionarModoAsistencia(modo) {
    if (!asistenciaModalEstado) return;
    asistenciaModalEstado.modo = modo;
    ['justificada', 'injustificada', 'borrar'].forEach(item => {
        const boton = document.getElementById(`asistencia-modo-${item}`);
        const activo = item === modo;
        boton.className = `px-3 py-2 rounded border text-sm font-semibold ${
            activo ? 'bg-indigo-600 text-white border-indigo-600' : 'bg-white text-slate-700 border-slate-300'
        }`;
    });
}

function cambiarMesCalendario() {
    if (!asistenciaModalEstado) return;
    asistenciaModalEstado.mes = Number(document.getElementById('asistencia-cal-mes').value);
    asistenciaModalEstado.anio = Number(document.getElementById('asistencia-cal-anio').value);
    renderCalendarioAsistencia();
}

function moverMesAsistencia(delta) {
    if (!asistenciaModalEstado) return;
    const fecha = new Date(asistenciaModalEstado.anio, asistenciaModalEstado.mes - 1 + delta, 1, 12);
    const selectorAnio = document.getElementById('asistencia-cal-anio');
    const anios = Array.from(selectorAnio.options).map(opcion => Number(opcion.value));
    if (!anios.includes(fecha.getFullYear())) return;
    asistenciaModalEstado.anio = fecha.getFullYear();
    asistenciaModalEstado.mes = fecha.getMonth() + 1;
    selectorAnio.value = String(asistenciaModalEstado.anio);
    document.getElementById('asistencia-cal-mes').value = String(asistenciaModalEstado.mes);
    renderCalendarioAsistencia();
}

function renderCalendarioAsistencia() {
    const contenedor = document.getElementById('asistencia-calendario');
    contenedor.innerHTML = '';
    if (!asistenciaModalEstado) return;
    const curso = obtenerCursoActivo();
    if (!curso || curso.id !== asistenciaModalEstado.cursoId) return;
    const trimestre = asegurarAsistenciaCurso(curso)[asistenciaModalEstado.trimestre];
    const primerDia = new Date(asistenciaModalEstado.anio, asistenciaModalEstado.mes - 1, 1, 12).getDay();
    const espacios = (primerDia + 6) % 7;
    const diasMes = new Date(asistenciaModalEstado.anio, asistenciaModalEstado.mes, 0, 12).getDate();
    for (let i = 0; i < espacios; i += 1) {
        const vacio = document.createElement('div');
        vacio.className = 'min-h-14';
        contenedor.appendChild(vacio);
    }
    for (let dia = 1; dia <= diasMes; dia += 1) {
        const fecha = AsistenciaUtils.isoDesdePartes(asistenciaModalEstado.anio, asistenciaModalEstado.mes, dia);
        const habilitado = AsistenciaUtils.esDiaLectivo(fecha, trimestre);
        const registro = asistenciaModalEstado.faltas[fecha];
        const boton = document.createElement('button');
        boton.type = 'button';
        boton.disabled = !habilitado;
        boton.className = 'min-h-14 rounded border p-1 text-sm flex flex-col items-center justify-center gap-0.5 transition-all ';
        if (!habilitado) {
            boton.className += 'bg-slate-100 border-slate-200 text-slate-400 cursor-not-allowed';
            boton.innerHTML = `<span>${dia}</span><i class="fa-solid fa-ban text-[10px]"></i>`;
        } else if (registro?.tipo === 'justificada') {
            boton.className += 'bg-emerald-50 border-emerald-400 text-emerald-800';
            boton.innerHTML = `<span>${dia}</span><span class="text-[9px] font-bold"><i class="fa-solid fa-circle-check"></i> J</span>`;
        } else if (registro?.tipo === 'injustificada') {
            boton.className += 'bg-rose-50 border-rose-400 text-rose-800';
            boton.innerHTML = `<span>${dia}</span><span class="text-[9px] font-bold"><i class="fa-solid fa-circle-xmark"></i> I</span>`;
        } else {
            boton.className += 'bg-white border-slate-300 text-slate-700 hover:border-indigo-400 hover:bg-indigo-50';
            boton.textContent = String(dia);
        }
        boton.title = habilitado ? AsistenciaUtils.formatoFecha(fecha) : 'Día no habilitado';
        if (habilitado) boton.onclick = () => aplicarModoAsistencia(fecha);
        contenedor.appendChild(boton);
    }
}

function aplicarModoAsistencia(fecha) {
    if (!asistenciaModalEstado) return;
    if (asistenciaModalEstado.modo === 'borrar') {
        delete asistenciaModalEstado.faltas[fecha];
    } else {
        asistenciaModalEstado.faltas[fecha] = {
            tipo: asistenciaModalEstado.modo,
            observacion: asistenciaModalEstado.faltas[fecha]?.observacion || ''
        };
    }
    asistenciaModalEstado.cambios = true;
    renderCalendarioAsistencia();
    renderListaFaltasAsistencia();
}

function renderListaFaltasAsistencia() {
    const cuerpo = document.getElementById('asistencia-faltas-lista');
    cuerpo.innerHTML = '';
    if (!asistenciaModalEstado) return;
    const fechas = Object.keys(asistenciaModalEstado.faltas).sort();
    if (fechas.length === 0) {
        const fila = cuerpo.insertRow();
        const celda = fila.insertCell();
        celda.colSpan = 3;
        celda.className = 'p-4 text-center text-slate-400';
        celda.textContent = 'Sin faltas registradas.';
        return;
    }
    fechas.forEach(fecha => {
        const registro = asistenciaModalEstado.faltas[fecha];
        const fila = cuerpo.insertRow();
        const fechaCelda = fila.insertCell();
        fechaCelda.className = 'p-2 align-top whitespace-nowrap';
        fechaCelda.textContent = `${AsistenciaUtils.formatoFecha(fecha)}\n${AsistenciaUtils.nombreDia(fecha)}`;
        fechaCelda.style.whiteSpace = 'pre-line';
        const detalle = fila.insertCell();
        detalle.className = 'p-2';
        const tipo = document.createElement('span');
        tipo.className = registro.tipo === 'justificada'
            ? 'text-emerald-700 font-semibold'
            : 'text-rose-700 font-semibold';
        tipo.textContent = registro.tipo === 'justificada' ? 'Justificada' : 'Injustificada';
        const observacion = document.createElement('input');
        observacion.type = 'text';
        observacion.value = registro.observacion || '';
        observacion.placeholder = 'Observación opcional';
        observacion.className = 'mt-1 w-full border border-slate-300 rounded px-2 py-1 text-xs';
        observacion.oninput = () => {
            asistenciaModalEstado.faltas[fecha].observacion = observacion.value;
            asistenciaModalEstado.cambios = true;
        };
        detalle.append(tipo, observacion);
        const accion = fila.insertCell();
        accion.className = 'p-2 text-center align-top';
        const eliminar = document.createElement('button');
        eliminar.type = 'button';
        eliminar.className = 'text-rose-600 hover:text-rose-800';
        eliminar.innerHTML = '<i class="fa-solid fa-trash"></i>';
        eliminar.setAttribute('aria-label', `Eliminar falta del ${AsistenciaUtils.formatoFecha(fecha)}`);
        eliminar.onclick = () => borrarFaltaAsistencia(fecha);
        accion.appendChild(eliminar);
    });
}

function borrarFaltaAsistencia(fecha) {
    if (!asistenciaModalEstado) return;
    delete asistenciaModalEstado.faltas[fecha];
    asistenciaModalEstado.cambios = true;
    renderCalendarioAsistencia();
    renderListaFaltasAsistencia();
}

async function guardarCambiosAsistencia() {
    if (!asistenciaModalEstado) return;
    const curso = obtenerCursoActivo();
    if (!curso || curso.id !== asistenciaModalEstado.cursoId) {
        mostrarAlerta('Curso diferente', 'La asistencia no puede guardarse en otro curso.', 'warning');
        return;
    }
    const pertenece = estudiantesCargados.some(
        item => obtenerEstudianteIdAsistencia(item) === asistenciaModalEstado.estudianteId
    );
    if (!pertenece) {
        mostrarAlerta('Estudiante no encontrado', 'El estudiante ya no pertenece al curso activo.', 'warning');
        return;
    }
    const trimestre = asegurarAsistenciaCurso(curso)[asistenciaModalEstado.trimestre];
    trimestre.estudiantes[asistenciaModalEstado.estudianteId] = {
        faltas: AsistenciaUtils.normalizarFaltas(asistenciaModalEstado.faltas, trimestre)
    };
    asistenciaModalEstado.cambios = false;
    if (!await persistirAsistencia()) return;
    document.getElementById('asistencia-modal').classList.add('hidden');
    asistenciaModalEstado = null;
    renderModuloAsistencia();
    mostrarAlerta('Asistencia guardada', 'Asistencia guardada correctamente.', 'success');
}

function cerrarModalAsistencia() {
    if (!asistenciaModalEstado) {
        document.getElementById('asistencia-modal').classList.add('hidden');
        return;
    }
    if (asistenciaModalEstado.cambios && !confirm('Existen cambios sin guardar. ¿Desea cerrar sin guardarlos?')) return;
    document.getElementById('asistencia-modal').classList.add('hidden');
    asistenciaModalEstado = null;
}

function inyectarAsistenciaDocumento(doc, estudiante) {
    const resumen = estudiante?.asistencia || {
        configurada: false,
        totalFaltas: null,
        justificadas: null,
        injustificadas: null,
        totalAsistencia: null
    };
    const valores = {
        registro: resumen.configurada ? resumen.totalFaltas : '',
        justificadas: resumen.configurada ? resumen.justificadas : '',
        injustificadas: resumen.configurada ? resumen.injustificadas : '',
        total: resumen.configurada ? resumen.totalAsistencia : ''
    };
    let celdas = Array.from(doc.querySelectorAll('[data-asistencia]'));
    if (celdas.length === 0) {
        const titulo = Array.from(doc.querySelectorAll('h3')).find(
            item => normalizarTextoAsignatura(item.textContent) === 'ASISTENCIA ANUAL'
        );
        const fila = titulo?.parentElement?.querySelector('tbody tr');
        celdas = fila ? Array.from(fila.cells) : [];
        ['registro', 'justificadas', 'injustificadas', 'total'].forEach((clave, indice) => {
            if (celdas[indice]) celdas[indice].dataset.asistencia = clave;
        });
    }
    celdas.forEach(celda => {
        const clave = celda.dataset.asistencia;
        if (Object.prototype.hasOwnProperty.call(valores, clave)) {
            celda.textContent = String(valores[clave]);
        }
    });
}
