'use strict';

const packageJson = require('../package.json');

const tag = process.env.RELEASE_TAG || process.env.GITHUB_REF_NAME || '';
if (!/^v\d+\.\d+\.\d+$/.test(tag)) {
    throw new Error(`La etiqueta debe tener el formato vX.Y.Z; se recibió "${tag}".`);
}

const tagVersion = tag.slice(1);
if (packageJson.version !== tagVersion) {
    throw new Error(
        `La versión de package.json (${packageJson.version}) no coincide con la etiqueta (${tag}).`
    );
}

process.stdout.write(`Versión validada: ${tag}\n`);
