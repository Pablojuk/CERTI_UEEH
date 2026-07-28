const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
    verificarLicencia: () => ipcRenderer.invoke('verificar-licencia'),
    analizarExcel: (solicitud) => ipcRenderer.invoke('analizar-excel', solicitud),
    obtenerCatalogoAsignaturas: () => ipcRenderer.invoke('obtener-catalogo-asignaturas'),
    obtenerEscalaCualitativa: () => ipcRenderer.invoke('obtener-escala-cualitativa'),
    generarBoletines: (datos) => ipcRenderer.invoke('generar-boletines', datos),
    seleccionarArchivo: (opciones) => ipcRenderer.invoke('seleccionar-archivo', opciones),
    descargarFormato: (formatoId, opciones = {}) => ipcRenderer.invoke('descargar-formato', formatoId, opciones),
    leerPlantilla: (nombre) => ipcRenderer.invoke('leer-plantilla', nombre),
    imprimirCertificados: (html) => ipcRenderer.invoke('imprimir-certificados', html),
    abrirVistaPreviaCertificado: (html) => ipcRenderer.invoke('abrir-vista-previa-certificado', html),
    generarCertificados: (datos) => ipcRenderer.invoke('generar-certificados', datos),
    actualizarSupletorios: (datos) => ipcRenderer.invoke('actualizar-supletorios', datos),
    leerCertificadoGenerado: (cursoId, nombre) => ipcRenderer.invoke('leer-certificado-generado', cursoId, nombre)
});
