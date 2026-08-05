const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
    obtenerEstadoLicencia: () => ipcRenderer.invoke('obtener-estado-licencia'),
    activarLicencia: (licencia) => ipcRenderer.invoke('activar-licencia', licencia),
    iniciarPrueba: () => ipcRenderer.invoke('iniciar-prueba'),
    buscarActualizaciones: () => ipcRenderer.invoke('actualizaciones:buscar'),
    descargarActualizacion: () => ipcRenderer.invoke('actualizaciones:descargar'),
    reiniciarEInstalarActualizacion: () => ipcRenderer.invoke('actualizaciones:instalar'),
    onEstadoActualizacion: (callback) => {
        if (typeof callback !== 'function') return () => {};
        const listener = (_event, estado) => callback(estado);
        ipcRenderer.on('actualizaciones:estado', listener);
        return () => ipcRenderer.removeListener('actualizaciones:estado', listener);
    },
    analizarExcel: (solicitud) => ipcRenderer.invoke('analizar-excel', solicitud),
    obtenerCatalogoAsignaturas: () => ipcRenderer.invoke('obtener-catalogo-asignaturas'),
    obtenerEscalaCualitativa: () => ipcRenderer.invoke('obtener-escala-cualitativa'),
    generarBoletines: (datos) => ipcRenderer.invoke('generar-boletines', datos),
    seleccionarArchivo: (opciones) => ipcRenderer.invoke('seleccionar-archivo', opciones),
    descargarFormato: (formatoId, opciones = {}) => ipcRenderer.invoke('descargar-formato', formatoId, opciones),
    leerPlantilla: (nombre) => ipcRenderer.invoke('leer-plantilla', nombre),
    imprimirCertificados: (solicitud) => ipcRenderer.invoke('imprimir-certificados', solicitud),
    abrirVistaPreviaCertificado: (html) => ipcRenderer.invoke('abrir-vista-previa-certificado', html),
    generarCertificados: (datos) => ipcRenderer.invoke('generar-certificados', datos),
    actualizarSupletorios: (datos) => ipcRenderer.invoke('actualizar-supletorios', datos),
    leerCertificadoGenerado: (cursoId, nombre) => ipcRenderer.invoke('leer-certificado-generado', cursoId, nombre)
});
