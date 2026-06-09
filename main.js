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
        const pyProcess = spawn(pythonExecutable, processArgs);
        
        let stdoutData = '';
        let stderrData = '';
        
        if (inputData) {
            pyProcess.stdin.write(JSON.stringify(inputData));
            pyProcess.stdin.end();
        }
        
        pyProcess.stdout.on('data', (data) => {
            stdoutData += data.toString();
        });
        
        pyProcess.stderr.on('data', (data) => {
            stderrData += data.toString();
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
