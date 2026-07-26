'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const raiz = path.resolve(__dirname, '..');
const indexSource = fs.readFileSync(path.join(raiz, 'index.html'), 'utf8');
const catalogo = JSON.parse(fs.readFileSync(path.join(raiz, 'catalogo_asignaturas.json'), 'utf8'));
const metadatosPorNombre = new Map(catalogo.map(entrada => [entrada.nombre, entrada]));

function extraerFuncion(nombre) {
    const inicio = indexSource.indexOf(`function ${nombre}(`);
    assert.ok(inicio >= 0, `No se encontró ${nombre} en index.html`);
    const inicioCuerpo = indexSource.indexOf('{', inicio);
    let profundidad = 0;
    for (let indice = inicioCuerpo; indice < indexSource.length; indice += 1) {
        if (indexSource[indice] === '{') profundidad += 1;
        if (indexSource[indice] === '}') profundidad -= 1;
        if (profundidad === 0) return indexSource.slice(inicio, indice + 1);
    }
    throw new Error(`La función ${nombre} no tiene cierre`);
}

const contexto = {
    mostrarValorSeguro: valor => (
        valor === null || valor === undefined || valor === '' || Number.isNaN(valor)
            ? ''
            : String(valor)
    ),
    tipoAsignaturaCatalogo: () => 'cuantitativa'
};
vm.createContext(contexto);
vm.runInContext(
    [
        extraerFuncion('_certConvertirNotaOptativa'),
        extraerFuncion('_certInjectOptativasBGU3'),
        extraerFuncion('_certInjectNotas')
    ].join('\n'),
    contexto
);

class ElementoFalso {
    constructor(tagName, id = '') {
        this.tagName = tagName.toUpperCase();
        this.id = id;
        this.children = [];
        this.dataset = {};
        this.style = {};
        this.className = '';
        this.textContent = '';
    }

    get cells() {
        return this.tagName === 'TR'
            ? this.children.filter(hijo => ['TD', 'TH'].includes(hijo.tagName))
            : undefined;
    }

    appendChild(hijo) {
        this.children.push(hijo);
        return hijo;
    }

    replaceChildren(...hijos) {
        this.children = [...hijos];
    }
}

class DocumentoFalso {
    constructor() {
        this.raiz = new ElementoFalso('div');
        this.optativas = new ElementoFalso('tbody', 'optativas-bgu3');
        this.estaticas = new ElementoFalso('tbody', 'materias-estaticas');
        this.raiz.appendChild(this.optativas);
        this.raiz.appendChild(this.estaticas);
    }

    createElement(tagName) {
        return new ElementoFalso(tagName);
    }

    querySelector(selector) {
        if (selector === '#optativas-bgu3') return this.optativas;
        return null;
    }

    querySelectorAll(selector) {
        const encontrados = [];
        const recorrer = elemento => {
            elemento.children.forEach(hijo => {
                if (selector === 'tr' && hijo.tagName === 'TR') encontrados.push(hijo);
                if (selector === 'th' && hijo.tagName === 'TH') encontrados.push(hijo);
                recorrer(hijo);
            });
        };
        recorrer(this.raiz);
        return encontrados;
    }

    agregarCivica() {
        const fila = this.createElement('tr');
        ['CÍVICA Y ACOMPAÑAMIENTO INTEGRAL EN EL AULA', '', '', ''].forEach(texto => {
            const celda = this.createElement('td');
            celda.textContent = texto;
            fila.appendChild(celda);
        });
        this.estaticas.appendChild(fila);
        return fila;
    }
}

function materiaOptativa(nombre, t1, t2 = null, t3 = null) {
    const metadatos = metadatosPorNombre.get(nombre);
    assert.ok(metadatos, `No se encontró ${nombre} en el catálogo`);
    return {
        ...metadatos,
        t1,
        t2,
        t3,
    };
}

function contenidoFilas(tbody) {
    return tbody.children.map(fila => fila.cells.map(celda => celda.textContent));
}

test('la escala JavaScript respeta todos los límites y vacíos de Python', () => {
    const casos = [
        [1.00, 'E-'], [1.49, 'E-'], [1.50, 'E+'], [2.49, 'E+'],
        [2.50, 'D-'], [3.49, 'D-'], [3.50, 'D+'], [4.49, 'D+'],
        [4.50, 'C-'], [5.49, 'C-'], [5.50, 'C+'], [6.49, 'C+'],
        [6.50, 'B-'], [7.49, 'B-'], [7.50, 'B+'], [8.49, 'B+'],
        [8.50, 'A-'], [9.49, 'A-'], [9.50, 'A+'], [10.00, 'A+'],
        [null, ''], [undefined, ''], ['', ''], ['   ', ''], ['None', ''],
        ['null', ''], ['undefined', ''], ['NaN', ''], [NaN, ''],
        ['texto', ''], [0, ''], [0.99, ''], [10.01, '']
    ];

    casos.forEach(([nota, esperado]) => {
        assert.equal(contexto._certConvertirNotaOptativa(nota), esperado, String(nota));
    });
});

test('la vista previa genera las cinco optativas reales, ordenadas y con T2/T3 vacíos', () => {
    const doc = new DocumentoFalso();
    const est = {
        materias: {
            'SOCIOLOGÍA': materiaOptativa('SOCIOLOGÍA', 9.41),
            'QUÍMICA SUPERIOR': materiaOptativa('QUÍMICA SUPERIOR', 8.40),
            'MATEMÁTICA SUPERIOR': materiaOptativa('MATEMÁTICA SUPERIOR', 9.55),
            'INVESTIGACIÓN EN CIENCIA Y TECNOLOGÍA': materiaOptativa('INVESTIGACIÓN EN CIENCIA Y TECNOLOGÍA', 8.18),
            'REDACCIÓN CREATIVA': materiaOptativa('REDACCIÓN CREATIVA', 7.76)
        }
    };

    contexto._certInjectOptativasBGU3(doc, est);

    assert.deepEqual(contenidoFilas(doc.optativas), [
        ['INVESTIGACIÓN EN CIENCIA Y TECNOLOGÍA', 'B+', '', ''],
        ['REDACCIÓN CREATIVA', 'B+', '', ''],
        ['SOCIOLOGÍA', 'A-', '', ''],
        ['QUÍMICA SUPERIOR', 'B+', '', ''],
        ['MATEMÁTICA SUPERIOR', 'A+', '', '']
    ]);
    assert.equal(doc.optativas.children[0].cells[0].className, 'text-left-cell font-semibold w-1/3');
    assert.equal(doc.optativas.children[3].cells[0].style.overflowWrap, 'anywhere');
});

test('otra combinación institucional muestra solo las optativas presentes', () => {
    const doc = new DocumentoFalso();
    contexto._certInjectOptativasBGU3(doc, {
        materias: {
            'TEATRO': materiaOptativa('TEATRO', 8.10, null, 9.60),
            'PSICOLOGÍA': materiaOptativa('PSICOLOGÍA', 6.80),
            'BIOLOGÍA SUPERIOR': materiaOptativa('BIOLOGÍA SUPERIOR', 4.80),
            'LENGUA EXTRANJERA: FRANCÉS': materiaOptativa('LENGUA EXTRANJERA: FRANCÉS', 9.60)
        }
    });

    assert.deepEqual(contenidoFilas(doc.optativas), [
        ['TEATRO', 'B+', '', 'A+'],
        ['PSICOLOGÍA', 'B-', '', ''],
        ['LENGUA EXTRANJERA: FRANCÉS', 'A+', '', ''],
        ['BIOLOGÍA SUPERIOR', 'C-', '', '']
    ]);
    assert.equal(JSON.stringify(contenidoFilas(doc.optativas)).includes('SOCIOLOGÍA'), false);
});

test('los metadatos bastan para renderizar una optativa futura sin nombres fijos', () => {
    const doc = new DocumentoFalso();
    contexto._certInjectOptativasBGU3(doc, {
        materias: {
            'NUEVA OPTATIVA OFICIAL CON UN NOMBRE EXTENSO': {
                t1: 9.8,
                t2: 8.8,
                t3: 7.8,
                tipo: 'cuantitativa',
                categoria: 'optativa',
                es_optativa_bgu3: true,
                presentacion_certificado: 'escala_cualitativa',
                permite_supletorio: false,
                orden: 400
            },
            'MATERIA GENERAL': {
                t1: 10,
                es_optativa_bgu3: false,
                presentacion_certificado: 'numerica',
                orden: 1
            }
        }
    });

    assert.deepEqual(contenidoFilas(doc.optativas), [
        ['NUEVA OPTATIVA OFICIAL CON UN NOMBRE EXTENSO', 'A+', 'A-', 'B+']
    ]);
});

test('cinco reaperturas no duplican filas ni alteran Cívica o las letras', () => {
    const doc = new DocumentoFalso();
    const civica = doc.agregarCivica();
    const est = {
        materias: {
            'MATEMÁTICA SUPERIOR': materiaOptativa('MATEMÁTICA SUPERIOR', 9.55),
            'SOCIOLOGÍA': materiaOptativa('SOCIOLOGÍA', null),
            'CÍVICA Y ACOMPAÑAMIENTO INTEGRAL EN EL AULA': {
                t1: 'B+',
                t2: null,
                t3: null,
                tipo: 'cualitativa',
                es_optativa_bgu3: false,
                presentacion_certificado: 'numerica'
            }
        }
    };

    for (let apertura = 0; apertura < 5; apertura += 1) {
        contexto._certInjectOptativasBGU3(doc, est);
        contexto._certInjectNotas(doc, est);
        assert.deepEqual(contenidoFilas(doc.optativas), [
            ['SOCIOLOGÍA', '', '', ''],
            ['MATEMÁTICA SUPERIOR', 'A+', '', '']
        ]);
        assert.equal(doc.estaticas.children.length, 1);
        assert.deepEqual(civica.cells.map(celda => celda.textContent), [
            'CÍVICA Y ACOMPAÑAMIENTO INTEGRAL EN EL AULA', 'B+', '', ''
        ]);
    }
});

test('vistaPrevia inyecta optativas antes de notas y no usa innerHTML para sus filas', () => {
    const vistaInicio = indexSource.indexOf('async function vistaPrevia(');
    const vistaFin = indexSource.indexOf('// Limpiar inputs de archivo', vistaInicio);
    const vista = indexSource.slice(vistaInicio, vistaFin);
    assert.ok(vista.indexOf('_certInjectOptativasBGU3(doc, est);') >= 0);
    assert.ok(
        vista.indexOf('_certInjectOptativasBGU3(doc, est);')
        < vista.indexOf('_certInjectNotas(doc, est);')
    );
    assert.doesNotMatch(extraerFuncion('_certInjectOptativasBGU3'), /innerHTML/);
});
