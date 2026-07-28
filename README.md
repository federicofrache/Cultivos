# Campo & Costo

App de seguimiento de costos e ingresos agrícolas: Campañas → Cultivos → Gastos e Ingresos (Ventas / Remitos).

## 1) Configurar Firebase

1. En https://console.firebase.google.com abrí tu proyecto (o creá uno nuevo).
2. **Agregar app web**: ⚙️ Configuración del proyecto → pestaña "General" → "Tus apps" → ícono `</>` → registrá la app → copiá el objeto `firebaseConfig`.
3. Pegalo en `src/firebase.js`, reemplazando los valores `TU_...`.
4. **Authentication**: en el menú lateral → Authentication → "Comenzar" → pestaña "Sign-in method" → habilitá **Correo electrónico/contraseña**.
5. **Firestore**: menú lateral → Firestore Database → "Crear base de datos" → elegí una región (ej. `southamerica-east1`) → modo producción.
6. **Storage** (para las facturas adjuntas): menú lateral → Storage → "Comenzar" → misma región.
7. Copiá las reglas de este proyecto a Firebase:
   - `firestore.rules` → Firestore Database → pestaña "Reglas" → pegar → Publicar.
   - `storage.rules` → Storage → pestaña "Reglas" → pegar → Publicar.
8. **Usuarios**: como el login es con usuario y contraseña, cada persona puede crear su cuenta desde la pantalla de login ("Crear una cuenta nueva"), o vos podés crearlas manualmente en Authentication → Users → "Add user".

## 2) Probar en tu computadora (opcional)

```bash
npm install
npm run dev
```

Abre `http://localhost:5173`.

## 3) Desplegar en Netlify

**Opción A — conectando un repositorio (recomendado):**
1. Subí esta carpeta a un repositorio de GitHub/GitLab.
2. En Netlify: "Add new site" → "Import an existing project" → elegí el repo.
3. Build command: `npm run build` — Publish directory: `dist` (ya viene configurado en `netlify.toml`).
4. Deploy. Netlify te da un link público (ej. `https://tu-app.netlify.app`), que podés compartir con tu equipo.

**Opción B — subida manual (rápida, sin GitHub):**
1. En tu computadora: `npm install && npm run build` (esto genera la carpeta `dist/`).
2. En Netlify: "Add new site" → "Deploy manually" → arrastrá la carpeta `dist/`.

## Notas importantes

- **Datos compartidos**: todos los usuarios logueados ven y editan la misma información (campañas, cultivos, gastos, ventas, remitos), en tiempo real.
- **Extracción con IA (voz / foto de factura)**: esas funciones llaman a la API de Anthropic directamente desde el navegador, lo cual expone tu clave si la agregás así. Para producción real te recomiendo mover esas llamadas a una Netlify Function (servidor) que guarde la API key como variable de entorno, en vez de hacer el `fetch` desde el cliente. Puedo armarte esa función cuando quieras — avisame.
- **Facturas adjuntas** se guardan en Firebase Storage y quedan enlazadas al gasto correspondiente.
