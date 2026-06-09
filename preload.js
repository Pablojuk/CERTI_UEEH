const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
    verificarLicencia: () => ipcRenderer.invoke('verificar-licencia'),
    analizarExcel: (rutas) => ipcRenderer.invoke('analizar-excel', rutas),
    generarBoletines: (datos) => ipcRenderer.invoke('generar-boletines', datos),
    seleccionarArchivo: (opciones) => ipcRenderer.invoke('seleccionar-archivo', opciones),
    descargarFormato: () => ipcRenderer.invoke('descargar-formato')
});
