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
    // mainWindow.webContents.openDevTools(); // Descomentar para depuración
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
ipcMain.handle('seleccionar-archivo', async (event, opciones) => {
    const result = await dialog.showOpenDialog(mainWindow, {
        properties: opciones.properties || ['openFile'],
        filters: opciones.filters || []
    });
    if (result.canceled || result.filePaths.length === 0) {
        return null;
    }
    return result.filePaths[0];
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
ipcMain.handle('analizar-excel', async (event, rutas) => {
    try {
        const args = ['--analizar'];
        if (rutas.t1) { args.push('--t1', rutas.t1); }
        if (rutas.t2) { args.push('--t2', rutas.t2); }
        if (rutas.t3) { args.push('--t3', rutas.t3); }
        if (rutas.su) { args.push('--su', rutas.su); }
        
        const result = await runPython(args);
        return JSON.parse(result.stdout);
    } catch (error) {
        console.error("Error en analizar-excel:", error);
        return { error: error.stderr || error.error || "Error al ejecutar el procesador de notas" };
    }
});

// IPC: Generar Boletines (PDF)
ipcMain.handle('generar-boletines', async (event, datos) => {
    try {
        // Ejecutamos Python pasando el payload JSON por stdin
        const result = await runPython(['--generar'], datos);
        return JSON.parse(result.stdout);
    } catch (error) {
        console.error("Error en generar-boletines:", error);
        return { success: false, error: error.stderr || error.error || "Error de ejecución del generador de PDF" };
    }
});
