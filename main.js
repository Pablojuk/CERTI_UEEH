const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const fs = require('fs');
const { spawn } = require('child_process');

let mainWindow;

function createWindow() {
    mainWindow = new BrowserWindow({
        width: 1280,
        height: 800,
        title: "UEEH • Sistema Académico - Boletines",
        webPreferences: {
            preload: path.join(__dirname, 'preload.js'),
            contextIsolation: true,
            nodeIntegration: false
        }
    });

    mainWindow.loadFile('index.html');
    // mainWindow.webContents.openDevTools(); // TEMP: habilitado para depuración
}

app.whenReady().then(() => {
    createWindow();

    app.on('activate', function () {
        if (BrowserWindow.getAllWindows().length === 0) createWindow();
    });
});

app.on('window-all-closed', function () {
    if (process.platform !== 'darwin') app.quit();
});

// IPC: Verificar Licencia
ipcMain.handle('verificar-licencia', async () => {
    const appDataPath = process.env.APPDATA || (process.platform === 'darwin' ? path.join(process.env.HOME, 'Library/Application Support') : path.join(process.env.HOME, '.config'));
    const licensePath = path.join(appDataPath, 'UEEH', 'license_info.dat');
    const exists = fs.existsSync(licensePath);
    if (exists) {
        const stats = fs.statSync(licensePath);
        return {
            valido: stats.size > 0,
            path: licensePath,
            mensaje: stats.size > 0 ? "Licencia activa de la Fase 1 detectada." : "Archivo de licencia vacío."
        };
    }
    return {
        valido: false,
        path: licensePath,
        mensaje: "No se detectó licencia en la ruta especificada."
    };
});

// IPC: Diálogo para seleccionar archivos (Logo o Excel)
ipcMain.handle('seleccionar-archivo', async (event, opciones = {}) => {
    const result = await dialog.showOpenDialog(mainWindow, {
        properties: opciones.properties || ['openFile'],
        filters: opciones.filters || []
    });
    if (result.canceled || result.filePaths.length === 0) {
        return null;
    }
    console.log("[seleccionar-archivo] Ruta seleccionada:", result.filePaths[0]);
    return result.filePaths[0];
});

// IPC: Descargar formato Excel de notas
ipcMain.handle('descargar-formato', async () => {
    try {
        const sourcePath = path.join(__dirname, 'FORMATO PARA CARGAR LAS NOTAS.xlsx');
        if (!fs.existsSync(sourcePath)) {
            return { success: false, error: "El archivo oficial de formato 'FORMATO PARA CARGAR LAS NOTAS.xlsx' no existe en el directorio de la aplicación." };
        }
        
        const result = await dialog.showSaveDialog(mainWindow, {
            title: "Guardar Formato de Carga de Notas",
            defaultPath: path.join(app.getPath('downloads'), "FORMATO PARA CARGAR LAS NOTAS.xlsx"),
            filters: [
                { name: "Archivos de Excel", extensions: ["xlsx"] }
            ]
        });

        if (result.canceled || !result.filePath) {
            return { success: false, cancelled: true };
        }

        fs.copyFileSync(sourcePath, result.filePath);
        return { success: true, path: result.filePath };
    } catch (error) {
        console.error("Error al descargar formato:", error);
        return { success: false, error: error.message };
    }
});

// Helper para ejecutar Python
function runPython(args, inputData = null) {
    return new Promise((resolve, reject) => {
        const pythonExecutable = 'python'; // Usa el python por defecto del sistema/entorno
        const scriptPath = path.join(__dirname, 'procesador_notas.py');
        
        const processArgs = [scriptPath, ...args];
        const pyProcess = spawn(pythonExecutable, processArgs, {
            env: {
                ...process.env,
                PYTHONIOENCODING: 'utf-8'
            }
        });
        
        let stdoutData = '';
        let stderrData = '';
        
        if (inputData) {
            pyProcess.stdin.write(JSON.stringify(inputData));
            pyProcess.stdin.end();
        }
        
        pyProcess.stdout.on('data', (data) => {
            stdoutData += data.toString('utf8');
        });
        
        pyProcess.stderr.on('data', (data) => {
            stderrData += data.toString('utf8');
        });
        
        pyProcess.on('close', (code) => {
            if (code === 0) {
                resolve({ success: true, stdout: stdoutData.trim() });
            } else {
                reject({ success: false, code: code, stderr: stderrData.trim() });
            }
        });
        
        pyProcess.on('error', (err) => {
            reject({ success: false, error: err.message });
        });
    });
}

// IPC: Analizar archivos Excel para listar estudiantes
ipcMain.handle('analizar-excel', async (event, rutas = {}) => {
    try {
        const args = ['--analizar'];
        const periodos = ['t1', 't2', 't3', 'su'];
        const rutasValidas = {};

        for (const periodo of periodos) {
            const ruta = typeof rutas[periodo] === 'string' ? rutas[periodo].trim() : '';
            if (!ruta) continue;
            if (!fs.existsSync(ruta)) {
                return { error: `Ruta inválida: el archivo seleccionado para ${periodo.toUpperCase()} no existe o no es accesible: ${ruta}` };
            }
            rutasValidas[periodo] = ruta;
            args.push(`--${periodo}`, ruta);
        }

        if (Object.keys(rutasValidas).length === 0) {
            return { error: "Ruta inválida: no se recibió una ruta real del archivo Excel seleccionado." };
        }
        
        console.log("[analizar-excel] Rutas validadas:", rutasValidas);
        console.log("[analizar-excel] Args enviados a Python:", args);
        const result = await runPython(args);
        console.log("[analizar-excel] stdout length:", result.stdout.length);
        console.log("[analizar-excel] stdout preview:", result.stdout.substring(0, 300));
        const parsed = JSON.parse(result.stdout);
        console.log("[analizar-excel] Parsed keys:", Object.keys(parsed));
        if (parsed.datosInstitucion) {
            console.log("[analizar-excel] datosInstitucion.nombreInstitucion:", parsed.datosInstitucion.nombreInstitucion);
        } else {
            console.log("[analizar-excel] WARNING: No datosInstitucion found in parsed result!");
        }
        return parsed;
    } catch (error) {
        console.error("Error en analizar-excel:", error);
        return { error: error.stderr || error.error || "Error al ejecutar el procesador de notas" };
    }
});

// IPC: Generar Boletines (PDF)
ipcMain.handle('generar-boletines', async (event, datos) => {
    try {
        const appDataPath = app.getPath('userData');
        const logosDir = path.join(appDataPath, 'logos');
        if (!fs.existsSync(logosDir)) {
            fs.mkdirSync(logosDir, { recursive: true });
        }

        // Helper to write base64 image to file
        const saveBase64Image = (base64Str, prefix) => {
            if (!base64Str || !base64Str.startsWith('data:image/')) {
                return base64Str; // Return as-is if it's already a filepath or empty
            }
            try {
                const matches = base64Str.match(/^data:([A-Za-z-+\/]+);base64,(.+)$/);
                if (matches && matches.length === 3) {
                    const ext = matches[1].split('/')[1] || 'png';
                    const buffer = Buffer.from(matches[2], 'base64');
                    const cursoId = datos.cursoActivoId || 'default';
                    const tempPath = path.join(logosDir, `${prefix}_${cursoId}.${ext}`);
                    fs.writeFileSync(tempPath, buffer);
                    return tempPath;
                }
            } catch (err) {
                console.error("Error saving base64 logo:", err);
            }
            return null;
        };

        if (datos.logos) {
            datos.logos.logo1 = saveBase64Image(datos.logos.logo1, 'logo1');
            datos.logos.logo2 = saveBase64Image(datos.logos.logo2, 'logo2');
        }

        // Ejecutamos Python pasando el payload JSON por stdin
        const result = await runPython(['--generar'], datos);
        return JSON.parse(result.stdout);
    } catch (error) {
        console.error("Error en generar-boletines:", error);
        return { success: false, error: error.stderr || error.error || "Error de ejecución del generador de PDF" };
    }
});

// IPC: Generar certificados HTML iniciales y detectar supletorios
ipcMain.handle('generar-certificados', async (event, datos) => {
    try {
        const appDataPath = app.getPath('userData');
        const logosDir = path.join(appDataPath, 'logos');
        if (!fs.existsSync(logosDir)) {
            fs.mkdirSync(logosDir, { recursive: true });
        }

        // Crear directorio de salida para certificados generados (fuera de assets/certificados)
        const cursoId = datos.cursoActivoId || 'default';
        const certOutputDir = path.join(appDataPath, 'certificados_temporales', cursoId);
        if (!fs.existsSync(certOutputDir)) {
            fs.mkdirSync(certOutputDir, { recursive: true });
        }
        datos.certOutputDir = certOutputDir;

        // Guardar logos base64 en archivos si es necesario
        const saveBase64Image = (base64Str, prefix) => {
            if (!base64Str || !base64Str.startsWith('data:image/')) {
                return base64Str;
            }
            try {
                const matches = base64Str.match(/^data:([A-Za-z-+\/]+);base64,(.+)$/);
                if (matches && matches.length === 3) {
                    const ext = matches[1].split('/')[1] || 'png';
                    const buffer = Buffer.from(matches[2], 'base64');
                    const tempPath = path.join(logosDir, `${prefix}_${cursoId}.${ext}`);
                    fs.writeFileSync(tempPath, buffer);
                    return tempPath;
                }
            } catch (err) {
                console.error("Error saving base64 logo:", err);
            }
            return null;
        };

        if (datos.logos) {
            datos.logos.logo1 = saveBase64Image(datos.logos.logo1, 'logo1');
            datos.logos.logo2 = saveBase64Image(datos.logos.logo2, 'logo2');
        }

        // Ejecutamos Python con --certificados y pasamos el payload
        const result = await runPython(['--certificados'], datos);
        return JSON.parse(result.stdout);
    } catch (error) {
        console.error("Error en generar-certificados:", error);
        return { success: false, error: error.stderr || error.error || "Error de ejecución al procesar certificados" };
    }
});

// IPC: Actualizar certificados con notas manuales de supletorio
ipcMain.handle('actualizar-supletorios', async (event, datos) => {
    try {
        let updates = [];
        let cursoId = 'default';
        if (datos && datos.updates && datos.cursoId) {
            updates = datos.updates;
            cursoId = datos.cursoId;
        } else {
            updates = datos || [];
        }

        const appDataPath = app.getPath('userData');
        const certOutputDir = path.join(appDataPath, 'certificados_temporales', cursoId);

        // Inyectar cert_output_dir en cada update para Python
        const updatesWithDir = updates.map(u => ({
            ...u,
            cert_output_dir: certOutputDir
        }));

        const payloadString = JSON.stringify(updatesWithDir);
        const result = await runPython(['--supletorios', payloadString]);
        return JSON.parse(result.stdout);
    } catch (error) {
        console.error("Error en actualizar-supletorios:", error);
        return { success: false, error: error.stderr || error.error || "Error de ejecución al actualizar supletorios" };
    }
});


// IPC: Leer plantilla de certificado desde assets/certificados/
ipcMain.handle('leer-plantilla', async (event, nombreArchivo) => {
    try {
        const templatePath = path.join(__dirname, 'assets', 'certificados', nombreArchivo);
        if (!fs.existsSync(templatePath)) {
            return { error: `Plantilla no encontrada: ${nombreArchivo}. Verifique que exista en assets/certificados/.` };
        }
        const content = fs.readFileSync(templatePath, 'utf-8');
        return { content };
    } catch (error) {
        console.error("Error en leer-plantilla:", error);
        return { error: error.message || "Error al leer la plantilla de certificado." };
    }
});

// IPC: Leer certificado generado desde userData/certificados_temporales/
ipcMain.handle('leer-certificado-generado', async (event, cursoId, nombreArchivo) => {
    try {
        const appDataPath = app.getPath('userData');
        const filePath = path.join(appDataPath, 'certificados_temporales', cursoId, nombreArchivo);
        if (!fs.existsSync(filePath)) {
            return { error: `Certificado generado no encontrado: ${nombreArchivo}` };
        }
        const content = fs.readFileSync(filePath, 'utf-8');
        return { content };
    } catch (error) {
        console.error("Error en leer-certificado-generado:", error);
        return { error: error.message || "Error al leer el certificado generado." };
    }
});

// IPC: Generar PDF de certificados desde HTML ensamblado
ipcMain.handle('imprimir-certificados', async (event, htmlContent) => {
    try {
        // Diálogo para elegir dónde guardar el PDF
        const saveResult = await dialog.showSaveDialog(mainWindow, {
            title: 'Guardar Certificados como PDF',
            defaultPath: path.join(app.getPath('documents'), 'Certificados_UEEH.pdf'),
            filters: [{ name: 'Archivo PDF', extensions: ['pdf'] }]
        });

        if (saveResult.canceled || !saveResult.filePath) {
            return { success: false, cancelled: true };
        }

        // Guardar HTML temporal en directorio de datos de la app
        const tempHtmlPath = path.join(app.getPath('userData'), 'temp_certificados.html');
        fs.writeFileSync(tempHtmlPath, htmlContent, 'utf-8');

        // Crear ventana oculta para renderizar el HTML
        const printWindow = new BrowserWindow({
            show: false,
            width: 1024,
            height: 768,
            webPreferences: {
                nodeIntegration: false,
                contextIsolation: true
            }
        });

        try {
            // Cargar el HTML temporal
            const fileUrl = 'file:///' + tempHtmlPath.replace(/\\/g, '/');
            await printWindow.loadURL(fileUrl);

            // Esperar a que Tailwind CSS (CDN) compile los estilos
            await new Promise(resolve => setTimeout(resolve, 2500));

            // Generar el PDF
            const pdfBuffer = await printWindow.webContents.printToPDF({
                printBackground: true,
                landscape: false,
                pageSize: 'A4',
                margins: {
                    marginType: 'custom',
                    top: 0.4,
                    bottom: 0.4,
                    left: 0.2,
                    right: 0.2
                }
            });

            // Guardar el PDF en la ruta elegida por el usuario
            fs.writeFileSync(saveResult.filePath, pdfBuffer);
            return { success: true, path: saveResult.filePath };
        } finally {
            printWindow.close();
            // Limpiar archivo temporal
            try { fs.unlinkSync(tempHtmlPath); } catch (e) { /* ignorar */ }
        }
    } catch (error) {
        console.error("Error en imprimir-certificados:", error);
        return { success: false, error: error.message || "Error al generar el PDF de certificados." };
    }
});

// IPC: Abrir vista previa de certificado en ventana nueva
ipcMain.handle('abrir-vista-previa-certificado', async (event, htmlContent) => {
    try {
        const tempHtmlPath = path.join(app.getPath('userData'), 'temp_preview_certificado.html');
        fs.writeFileSync(tempHtmlPath, htmlContent, 'utf-8');

        const previewWindow = new BrowserWindow({
            width: 900,
            height: 700,
            title: 'Vista Previa del Certificado',
            autoHideMenuBar: true,
            webPreferences: {
                nodeIntegration: false,
                contextIsolation: true
            }
        });

        const fileUrl = 'file:///' + tempHtmlPath.replace(/\\/g, '/');
        await previewWindow.loadURL(fileUrl);

        // Limpiar archivo temporal cuando la ventana se cierre
        previewWindow.on('closed', () => {
            try { fs.unlinkSync(tempHtmlPath); } catch (e) { /* ignorar */ }
        });

        return { success: true };
    } catch (error) {
        console.error("Error en abrir-vista-previa-certificado:", error);
        return { success: false, error: error.message };
    }
});
