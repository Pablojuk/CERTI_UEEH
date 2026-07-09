const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
    verificarLicencia: () => ipcRenderer.invoke('verificar-licencia'),
    analizarExcel: (rutas) => ipcRenderer.invoke('analizar-excel', rutas),
    generarBoletines: (datos) => ipcRenderer.invoke('generar-boletines', datos),
    seleccionarArchivo: (opciones) => ipcRenderer.invoke('seleccionar-archivo', opciones),
    descargarFormato: () => ipcRenderer.invoke('descargar-formato'),
    leerPlantilla: (nombre) => ipcRenderer.invoke('leer-plantilla', nombre),
    imprimirCertificados: (html) => ipcRenderer.invoke('imprimir-certificados', html),
    abrirVistaPreviaCertificado: (html) => ipcRenderer.invoke('abrir-vista-previa-certificado', html),
    generarCertificados: (datos) => ipcRenderer.invoke('generar-certificados', datos),
    actualizarSupletorios: (datos) => ipcRenderer.invoke('actualizar-supletorios', datos)
});
