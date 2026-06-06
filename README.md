# BankTrack 💳

App móvil PWA para registrar movimientos bancarios con IA, a partir de capturas de pantalla y comprobantes.

## ¿Qué hace?
- 📸 Analiza comprobantes bancarios con IA (Claude)
- 💰 Registra ingresos y egresos automáticamente
- 📊 Estadísticas por categoría
- 📥 Exporta a Excel (.xlsx)
- 📱 Instalable como app en cualquier celular (PWA)

---

## 🚀 Cómo subir a Vercel (gratis) en 5 pasos

### Paso 1 — Instala las dependencias
```bash
npm install
```

### Paso 2 — Prueba en local (opcional)
```bash
npm run dev
```
Abre http://localhost:5173 en tu navegador.

### Paso 3 — Sube el código a GitHub
1. Crea una cuenta en https://github.com (gratis)
2. Crea un repositorio nuevo (botón verde "New")
3. Sube esta carpeta:
```bash
git init
git add .
git commit -m "BankTrack PWA"
git branch -M main
git remote add origin https://github.com/TU_USUARIO/banktrack.git
git push -u origin main
```

### Paso 4 — Conecta con Vercel
1. Ve a https://vercel.com y crea cuenta gratis
2. Haz clic en "Add New Project"
3. Selecciona tu repositorio de GitHub
4. Vercel detecta automáticamente que es Vite — haz clic en **Deploy**
5. En ~1 minuto tendrás una URL como: `https://banktrack-xxx.vercel.app`

### Paso 5 — Instala en tu celular
1. Abre la URL en tu celular (Chrome en Android / Safari en iPhone)
2. **Android**: toca el menú ⋮ → "Añadir a pantalla de inicio"
3. **iPhone**: toca el botón compartir ⬆ → "Añadir a pantalla de inicio"
4. ¡Listo! Ya tienes el ícono en tu celular como app nativa 🎉

---

## 🔑 Nota sobre la API Key

La app usa la API de Anthropic para analizar imágenes.
El proxy de Claude.ai maneja la autenticación automáticamente cuando se usa desde claude.ai.

Si la despliegas en tu propio servidor y quieres que funcione de forma independiente,
necesitas agregar tu API key de Anthropic. Crea un archivo `.env`:

```
VITE_ANTHROPIC_API_KEY=sk-ant-...
```

Y modifica el fetch en `src/App.jsx` para incluir el header:
```js
headers: {
  'Content-Type': 'application/json',
  'x-api-key': import.meta.env.VITE_ANTHROPIC_API_KEY,
  'anthropic-version': '2023-06-01',
  'anthropic-dangerous-direct-browser-access': 'true',
}
```

---

## 📁 Estructura del proyecto

```
banktrack/
├── public/
│   └── favicon.svg
├── src/
│   ├── App.jsx       ← App completa
│   ├── index.css     ← Estilos globales
│   └── main.jsx      ← Entry point
├── index.html
├── package.json
├── vite.config.js    ← Config PWA
└── vercel.json       ← Config despliegue
```
