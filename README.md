# Ollama Office Add-in (Word + Excel)

Este proyecto contiene un add-in de Office (panel de tareas) que se conecta a un servidor local de Ollama para generar texto con modelos locales y luego insertarlo en Word o Excel.

## Archivos principales
- `manifest.xml`: manifiesto del add-in para Word y Excel.
- `taskpane.html`: UI del panel.
- `taskpane.js`: lógica para conectar con Ollama y escribir en Word/Excel.
- `styles.css`: estilos básicos.

## Requisitos
- Ollama en ejecución en `http://localhost:11434`.
- Node.js (para el servidor HTTPS local en desarrollo).

## Servir en https://localhost:3000
1) Instala dependencias:
```
npm install
```

2) Levanta el servidor HTTPS:
```
npm run start-dev
```

Esto instala certificados de desarrollo, sirve la carpeta en `https://localhost:3000` y hace proxy de `/ollama` hacia `http://localhost:11434` para evitar problemas de mixed content y CORS.

## Cargar el add-in en Word/Excel
1) Abre Word o Excel.
2) Home (o Insertar) → Complementos → Más complementos.
3) Pestaña “Mis complementos” → “Cargar mi complemento” → “Desde archivo”.
4) Selecciona `manifest.xml`.

> Si no aparece la opción de cargar XML, puede estar bloqueado por tu cuenta empresarial.

## Instalación en macOS (sideload por carpeta Wef)
Excel en Mac solo lee add-ins desde esta ruta:
`~/Library/Containers/com.microsoft.Excel/Data/Documents/wef`

Word en Mac solo lee add-ins desde esta ruta:
`~/Library/Containers/com.microsoft.Word/Data/Documents/wef`

1) Crea la carpeta (si no existe):
```
mkdir -p ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef
mkdir -p ~/Library/Containers/com.microsoft.Word/Data/Documents/wef
```

2) Copia SOLO el `manifest.xml`:
```
cp manifest.xml ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/
cp manifest.xml ~/Library/Containers/com.microsoft.Word/Data/Documents/wef/
```

3) Borra cache (recomendado tras cambios):
- `~/Library/Containers/com.microsoft.Word/Data/Library/Application Support/Microsoft/Office/16.0/Wef/`
- `~/Library/Containers/com.microsoft.Excel/Data/Library/Application Support/Microsoft/Office/16.0/Wef/`
- `~/Library/Containers/com.microsoft.Office365ServiceV2/Data/Library/Application Support/Microsoft/Office/16.0/Wef/`

4) Cierra y abre Word/Excel, luego carga el add‑in.

**Cambios recientes**
- Bump de versión en `manifest.xml` → `1.0.7.0`
- Cache‑buster en `taskpane.html` (`taskpane.js?v=20260204_7` y `styles.css?v=20260204_7`)

Verifica que:
`https://localhost:3000/taskpane.js` muestre:
```
const OLLAMA_BASE_URL = "/ollama";
```

## Instalación en Windows (Shared Folder Catalog)
Esto evita “modo depuración” y normalmente funciona aunque el Store esté bloqueado.

1) Crea una carpeta y compártela:
   - Crea, por ejemplo: `C:\OfficeAddinCatalog`
   - Click derecho → Propiedades → Compartir → Uso compartido avanzado
   - Marca “Compartir esta carpeta”
   - Nombre del recurso compartido: `OfficeAddinCatalog`
   - Permisos: Lectura para tu usuario (y si quieres “Todos” lectura)

2) Obtén la ruta UNC:
```
\\TU-PC\OfficeAddinCatalog
```
Para saber el nombre exacto de tu PC:
- Inicio → Configuración → Sistema → Acerca de → “Nombre del dispositivo”
- o en PowerShell: `hostname`

3) Copia `manifest.xml` a:
```
C:\OfficeAddinCatalog\
```

4) En Excel:
File → Options → Trust Center → Trust Center Settings → Trusted Add-in Catalogs

En “Catalog Url”, pon:
```
\\TU-PC\OfficeAddinCatalog
```

5) Cárgalo desde “Shared Folder”:
Insert → Get Add-ins (o Complementos) → pestaña “SHARED FOLDER”.

## Nota sobre CORS
Si ves errores de red, habilita CORS en el servidor de Ollama para permitir el origen `https://localhost:3000`.

## Troubleshooting
**No aparecen modelos**
- Verifica que Ollama responde en el navegador:
  - `http://localhost:11434/api/tags`
  - `https://localhost:3000/ollama/api/tags`

**Bloqueo por mixed content**
- Si ves errores “requested insecure content”, significa que el add‑in está llamando a `http://`.
- Usa el proxy `/ollama` y verifica que `taskpane.js` tenga:
```
const OLLAMA_BASE_URL = "/ollama";
```

**Cache vieja en macOS**
- Si aún ves `http://localhost:11434` en el inspector, borra cache Wef y vuelve a cargar el manifest.

**No aparece “Cargar mi complemento”**
- En cuentas empresariales, el sideload puede estar bloqueado.
- Usa “Shared Folder Catalog” en Windows o pide habilitar sideloading al admin.

**Certificado HTTPS no confiado**
- Ejecuta:
```
npx office-addin-dev-certs install
```
- Luego reinicia Word/Excel.

**Limpiar cache rápido en macOS**
- Ejecuta:
```
npm run clean-mac-cache
```

**Limpiar cache rápido en Windows**
- Ejecuta en PowerShell:
```
npm run clean-windows-cache
```
