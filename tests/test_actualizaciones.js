'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

const raiz = path.resolve(__dirname, '..');
const main = require(path.join(raiz, 'main.js'));
const fuenteMain = fs.readFileSync(path.join(raiz, 'main.js'), 'utf8');
const fuenteIndex = fs.readFileSync(path.join(raiz, 'index.html'), 'utf8');
const fuentePreload = fs.readFileSync(path.join(raiz, 'preload.js'), 'utf8');
const paquete = JSON.parse(fs.readFileSync(path.join(raiz, 'package.json'), 'utf8'));
const workflow = fs.readFileSync(
    path.join(raiz, '.github', 'workflows', 'release.yml'),
    'utf8'
);

test('una aplicación empaquetada anterior reconoce solo una versión superior', () => {
    assert.equal(main.esVersionSuperior('1.1.0', '1.0.0'), true);
    assert.equal(main.esVersionSuperior('2.0.0', '1.99.99'), true);
    assert.equal(main.esVersionSuperior('1.1.0', '1.1.0'), false);
    assert.equal(main.esVersionSuperior('1.0.9', '1.1.0'), false);
    assert.equal(main.esVersionSuperior('1.1.0-beta.1', '1.1.0'), false);
    assert.equal(main.esVersionSuperior('versión-inválida', '1.0.0'), false);
    assert.match(fuenteMain, /allowDowngrade\s*=\s*false/);
});

test('npm start no activa el actualizador y la comprobación periódica evita duplicados', () => {
    assert.equal(paquete.scripts.start, 'electron .');
    assert.match(fuenteMain, /if\s*\(!app\?\.isPackaged\)\s*return null/);
    assert.match(fuenteMain, /if\s*\(comprobacionActualizacionEnCurso\)/);
    assert.match(fuenteMain, /setTimeout\(\(\)\s*=>/);
    assert.match(fuenteMain, /setInterval\(\(\)\s*=>/);
    assert.equal(main.RETRASO_ACTUALIZACION_INICIAL_MS >= 3000, true);
    assert.equal(main.INTERVALO_ACTUALIZACIONES_MS >= 60 * 1000, true);
});

test('preload ofrece IPC específico sin exponer ipcRenderer ni credenciales', () => {
    for (const patron of [
        /buscarActualizaciones:\s*\(\)\s*=>/,
        /descargarActualizacion:\s*\(\)\s*=>/,
        /reiniciarEInstalarActualizacion:\s*\(\)\s*=>/,
        /onEstadoActualizacion:\s*\(callback\)\s*=>/
    ]) {
        assert.match(fuentePreload, patron);
    }
    assert.doesNotMatch(fuentePreload, /ipcRenderer\s*:/);
    assert.doesNotMatch(fuentePreload, /GH_TOKEN|GITHUB_TOKEN|github_pat_|ghp_/i);
});

test('la interfaz permite posponer, buscar otra vez y muestra progreso y notas', () => {
    for (const texto of [
        'Buscar actualizaciones',
        'Actualizar ahora',
        'Más tarde',
        'Reiniciar e instalar',
        'Versión instalada',
        'Nueva versión',
        'Notas de la versión'
    ]) {
        assert.match(fuenteIndex, new RegExp(texto));
    }
    assert.match(fuenteIndex, /onclick="buscarActualizacionesManual\(\)"/);
    assert.match(fuenteIndex, /window\.electronAPI\.buscarActualizaciones\(\)/);
    const inicioCerrar = fuenteIndex.indexOf('function cerrarModalActualizacion()');
    const finCerrar = fuenteIndex.indexOf('function textoBreveActualizacion', inicioCerrar);
    const cuerpoCerrar = fuenteIndex.slice(inicioCerrar, finCerrar);
    assert.match(cuerpoCerrar, /classList\.add\('hidden'\)/);
    assert.doesNotMatch(cuerpoCerrar, /quit|install|reiniciar/i);
});

test('sin conexión no cierra la app y solo instala después de descargar', () => {
    assert.equal(main.esErrorSinConexion({ code: 'ENOTFOUND' }), true);
    assert.equal(main.esErrorSinConexion(new Error('internet disconnected')), true);
    assert.equal(main.esErrorSinConexion(new Error('firma inválida')), false);
    assert.match(fuenteMain, /Sin conexión\. La aplicación seguirá funcionando normalmente/);
    assert.match(
        fuenteMain,
        /if\s*\(!app\?\.isPackaged\s*\|\|\s*!actualizacionDescargada\)/
    );
    assert.match(fuenteMain, /quitAndInstall\(false,\s*true\)/);
    assert.match(fuenteMain, /autoDownload\s*=\s*false/);
    assert.match(fuenteMain, /autoInstallOnAppQuit\s*=\s*false/);
});

test('NSIS, GitHub Releases y los recursos preservan datos y archivos actuales', () => {
    assert.match(paquete.version, /^\d+\.\d+\.\d+$/);
    assert.equal(paquete.build.appId, 'ec.edu.ueeh.certi');
    assert.equal(paquete.build.productName, 'CERTI_UEEH');
    assert.equal(paquete.build.publish.provider, 'github');
    assert.equal(paquete.build.publish.owner, 'Pablojuk');
    assert.equal(paquete.build.publish.repo, 'CERTI_UEEH');
    assert.equal(paquete.build.publish.releaseType, 'draft');
    assert.equal(paquete.build.nsis.deleteAppDataOnUninstall, false);
    assert.equal(paquete.build.nsis.differentialPackage, true);
    assert.ok(paquete.build.extraResources.some(recurso => (
        recurso.from === 'build/python/certi-python'
        && recurso.to === 'python'
    )));
    assert.ok(paquete.build.extraResources.some(recurso => (
        recurso.from === 'assets/formatos-notas'
        && recurso.to === 'assets/formatos-notas'
    )));
    assert.ok(paquete.build.files.includes('assets/certificados/**/*'));
});

test('el workflow valida la etiqueta, prueba, empaqueta y publica sin tokens propios', () => {
    assert.match(workflow, /tags:\s*\r?\n\s*-\s*"v\*"/);
    assert.match(workflow, /contents:\s*write/);
    assert.match(workflow, /npm ci/);
    assert.match(workflow, /npm test/);
    assert.match(workflow, /unittest discover/);
    assert.match(workflow, /npm run release/);
    assert.match(workflow, /secrets\.GITHUB_TOKEN/);
    assert.match(workflow, /gh release edit/);
    assert.match(workflow, /--draft=false/);
    assert.match(workflow, /Falta el asset requerido/);
    assert.doesNotMatch(workflow, /PERSONAL_ACCESS_TOKEN|github_pat_|ghp_/i);
    assert.match(workflow, /latest\.yml/);
    assert.match(workflow, /\.exe\.blockmap/);
});

test('las notas de versión se normalizan sin HTML ejecutable', () => {
    assert.equal(
        main.normalizarNotasVersion([
            { version: '1.1.0', note: 'Mejoras de estabilidad' },
            { version: '1.0.1', note: 'Correcciones' }
        ]),
        'Versión 1.1.0\nMejoras de estabilidad\n\nVersión 1.0.1\nCorrecciones'
    );
    assert.match(fuenteIndex, /actualizacion-notas'\)\.textContent/);
});
