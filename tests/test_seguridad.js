const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { pathToFileURL } = require('node:url');

const raiz = path.resolve(__dirname, '..');
const main = require(path.join(raiz, 'main.js'));

test('la lista blanca IPC es explícita y no expone canales genéricos', () => {
    assert.ok(Object.isFrozen(main.CANALES_IPC_PERMITIDOS));
    assert.deepEqual(
        new Set(main.CANALES_IPC_PERMITIDOS),
        new Set([
            'obtener-estado-licencia',
            'activar-licencia',
            'iniciar-prueba',
            'actualizaciones:buscar',
            'actualizaciones:descargar',
            'actualizaciones:instalar',
            'seleccionar-archivo',
            'descargar-formato',
            'analizar-excel',
            'obtener-catalogo-asignaturas',
            'obtener-escala-cualitativa',
            'generar-boletines',
            'generar-certificados',
            'actualizar-supletorios',
            'leer-plantilla',
            'leer-certificado-generado',
            'imprimir-certificados',
            'abrir-vista-previa-certificado'
        ])
    );
    assert.equal(main.CANALES_IPC_PERMITIDOS.includes('fs'), false);
    assert.equal(main.CANALES_IPC_PERMITIDOS.includes('shell'), false);
});

test('las rutas hijas y los identificadores bloquean traversal y separadores', () => {
    const base = path.join(raiz, 'assets', 'certificados');
    assert.equal(
        main.resolverRutaHija(base, 'FORMATO EGBM.html'),
        path.resolve(base, 'FORMATO EGBM.html')
    );
    assert.equal(main.resolverRutaHija(base, '..', 'main.js'), null);
    assert.equal(main.esSegmentoRutaSeguro('curso_123'), true);
    assert.equal(main.esSegmentoRutaSeguro('../curso'), false);
    assert.equal(main.esSegmentoRutaSeguro('curso\\otro'), false);
    assert.equal(main.esSegmentoRutaSeguro('curso\0otro'), false);
});

test('el análisis acepta solo Excel regular, absoluto, no vacío y de tamaño limitado', () => {
    const temporal = fs.mkdtempSync(path.join(os.tmpdir(), 'certi-seguridad-'));
    try {
        const excel = path.join(temporal, 'notas.xlsx');
        const texto = path.join(temporal, 'notas.txt');
        fs.writeFileSync(excel, Buffer.from('xlsx-de-prueba'));
        fs.writeFileSync(texto, Buffer.from('texto'));
        assert.equal(main.validarArchivoExcel(excel), true);
        assert.equal(main.validarArchivoExcel(texto), false);
        assert.equal(main.validarArchivoExcel('notas.xlsx'), false);
        assert.equal(main.validarArchivoExcel(path.join(temporal, 'ausente.xlsx')), false);
    } finally {
        fs.rmSync(temporal, { recursive: true, force: true });
    }
});

test('los payloads rechazan cursos o estudiantes utilizables como ruta', () => {
    assert.equal(main.validarPayloadGeneracion({
        cursoActivoId: 'curso_1',
        datos_consolidados: [{ id_real: 'cedula_0102030405' }]
    }), null);
    assert.match(main.validarPayloadGeneracion({
        cursoActivoId: '../curso',
        datos_consolidados: []
    }), /curso/i);
    assert.match(main.validarPayloadGeneracion({
        cursoActivoId: 'curso_1',
        datos_consolidados: [{ id_real: '../../estudiante' }]
    }), /estudiante/i);
    assert.equal(main.validarHtmlRecibido('<!doctype html><p>ok</p>'), true);
    assert.equal(main.validarHtmlRecibido(''), false);
});

test('IPC exige el webContents principal y la URL local exacta', () => {
    const webContents = {};
    const ventana = { webContents };
    const urlPrincipal = pathToFileURL(path.join(raiz, 'index.html')).href;
    assert.equal(main.esSolicitudIpcAutorizada({
        sender: webContents,
        senderFrame: { url: urlPrincipal }
    }, ventana), true);
    assert.equal(main.esSolicitudIpcAutorizada({
        sender: {},
        senderFrame: { url: urlPrincipal }
    }, ventana), false);
    assert.equal(main.esSolicitudIpcAutorizada({
        sender: webContents,
        senderFrame: { url: 'https://sitio-no-autorizado.example/' }
    }, ventana), false);
});

test('las opciones del selector descartan propiedades y extensiones arbitrarias', () => {
    const opciones = main.normalizarOpcionesSeleccionArchivo({
        properties: ['openDirectory', 'multiSelections'],
        filters: [{
            name: 'Prueba\r\n',
            extensions: ['xlsx', 'exe', '.PNG']
        }]
    });
    assert.deepEqual(opciones.properties, ['openFile']);
    assert.deepEqual(opciones.filters[0].extensions, ['xlsx', 'png']);
    assert.equal(opciones.filters[0].name.includes('\n'), false);
});

test('la descarga nunca sobrescribe la plantilla original', async () => {
    const temporal = fs.mkdtempSync(path.join(os.tmpdir(), 'certi-original-'));
    try {
        const original = path.join(temporal, 'formato_2_a_7_egb.xlsx');
        fs.writeFileSync(original, Buffer.from('plantilla-original'));
        const resultado = await main.descargarFormatoNotasSeguro(
            'egb_2_7',
            {},
            {
                directorioFormatos: temporal,
                fs,
                path,
                app: { getPath: () => temporal, isPackaged: false },
                dialog: {
                    showSaveDialog: async () => ({
                        canceled: false,
                        filePath: original
                    })
                }
            }
        );
        assert.equal(resultado.success, false);
        assert.equal(resultado.code, 'DESTINO_NO_PERMITIDO');
        assert.equal(fs.readFileSync(original, 'utf8'), 'plantilla-original');
    } finally {
        fs.rmSync(temporal, { recursive: true, force: true });
    }
});

test('Electron, CSP y preload conservan una superficie mínima', () => {
    const fuenteMain = fs.readFileSync(path.join(raiz, 'main.js'), 'utf8');
    const fuenteIndex = fs.readFileSync(path.join(raiz, 'index.html'), 'utf8');
    const fuentePreload = fs.readFileSync(path.join(raiz, 'preload.js'), 'utf8');

    for (const patron of [
        /contextIsolation:\s*true/,
        /nodeIntegration:\s*false/,
        /sandbox:\s*true/,
        /webSecurity:\s*true/,
        /allowRunningInsecureContent:\s*false/,
        /experimentalFeatures:\s*false/,
        /setWindowOpenHandler/,
        /setPermissionRequestHandler/
    ]) {
        assert.match(fuenteMain, patron);
    }
    assert.match(fuenteIndex, /Content-Security-Policy/);
    assert.doesNotMatch(fuenteIndex, /parentDiv\.innerHTML/);
    assert.doesNotMatch(fuenteIndex, /titleElem\.innerHTML|msgElem\.innerHTML/);
    assert.doesNotMatch(fuenteIndex, /console\.log/);
    assert.doesNotMatch(fuentePreload, /require\(['"](?:fs|path|child_process)['"]\)/);
    assert.doesNotMatch(fuentePreload, /ipcRenderer\s*:/);
});

test('todas las plantillas autorizadas incluyen CSP', () => {
    for (const nombre of main.PLANTILLAS_CERTIFICADO_PERMITIDAS) {
        const contenido = fs.readFileSync(
            path.join(raiz, 'assets', 'certificados', nombre),
            'utf8'
        );
        assert.match(contenido, /Content-Security-Policy/, nombre);
        assert.match(contenido, /object-src 'none'/, nombre);
    }
});

test('el procesador Python queda en una ruta ejecutable fuera de ASAR', () => {
    assert.equal(
        main.obtenerDirectorioProcesador(
            { isPackaged: false },
            'C:\\recursos',
            raiz
        ),
        raiz
    );
    assert.equal(
        main.obtenerDirectorioProcesador(
            { isPackaged: true },
            'C:\\recursos',
            raiz
        ),
        path.join('C:\\recursos', 'app.asar.unpacked')
    );
    const paquete = JSON.parse(fs.readFileSync(path.join(raiz, 'package.json'), 'utf8'));
    for (const recurso of [
        'procesador_notas.py',
        'catalogo_asignaturas.py',
        'catalogo_asignaturas.json',
        'escala_cualitativa.json',
        'assets/certificados/**/*'
    ]) {
        assert.ok(paquete.build.asarUnpack.includes(recurso), recurso);
    }
    assert.deepEqual(paquete.build.electronFuses, {
        runAsNode: false,
        enableCookieEncryption: true,
        enableNodeOptionsEnvironmentVariable: false,
        enableNodeCliInspectArguments: false,
        enableEmbeddedAsarIntegrityValidation: true,
        onlyLoadAppFromAsar: true,
        loadBrowserProcessSpecificV8Snapshot: false,
        grantFileProtocolExtraPrivileges: true
    });
});
