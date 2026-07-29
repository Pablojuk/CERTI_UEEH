# CERTI_UEEH

## Actualizaciones automáticas

La versión `1.1.5` incorpora actualizaciones automáticas para las instalaciones
NSIS de Windows. Las instalaciones anteriores, que todavía no incluyen el
actualizador, deben instalar manualmente esta versión una vez. A partir de ella,
las siguientes versiones se pueden descargar e instalar desde la aplicación.

El actualizador solo se activa en una aplicación empaquetada. `npm start` no
consulta GitHub Releases. Los datos académicos, la licencia y los archivos
guardados por el usuario se conservan al actualizar o desinstalar.

### Publicar una actualización

1. Modificar y probar la aplicación.
2. Aumentar la versión semántica de `package.json`.
3. Hacer commit y push a `main`.
4. Crear y subir la etiqueta `vX.Y.Z`; debe coincidir exactamente con la versión.
5. Esperar que GitHub Actions publique la Release.
6. Probar la actualización en otra PC con una versión anterior instalada.

Ejemplo:

```powershell
npm test
git add .
git commit -m "release: versión 1.1.1"
git push origin main
git tag v1.1.1
git push origin v1.1.1
```

El workflow `.github/workflows/release.yml` ejecuta las pruebas en Windows,
prepara el ejecutable Python y genera el instalador NSIS. Primero adjunta el
instalador, `latest.yml` y el `blockmap` a un borrador; después de validar los
tres assets publica la Release no borrador. Todo usa el `GITHUB_TOKEN` de GitHub
Actions y no se incluyen tokens en la aplicación.

### Construcción local

Instale las dependencias de compilación y genere el instalador:

```powershell
npm ci
python -m pip install -r requirements.txt
python -m pip install -r requirements-build.txt
npm run build
```

El instalador queda en `dist/`. El recurso Python autónomo se prepara en
`build/python/` y se incluye en `resources/python/` dentro de la aplicación.
