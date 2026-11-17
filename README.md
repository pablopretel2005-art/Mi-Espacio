# Mi espacio

App personal de notas, textos para copiar, listas, finanzas, hábitos, Universidad y accesos directos. Funciona en el móvil y en el portátil con los mismos datos, guardados en una hoja de Google y protegidos con contraseña.

| Archivo | Para qué sirve |
|---|---|
| `index.html` | La app entera. Es lo que publica GitHub Pages. |
| `Codigo.gs` | El servidor. Va pegado dentro de la hoja de Google (Apps Script). GitHub Pages no lo usa: aquí está como copia de referencia. |
| `sw.js` | Permite abrir la app sin conexión. Guarda solo la app, nunca tus datos. |
| `manifest.webmanifest`, `icon-180.png`, `icon-512.png` | Nombre e icono al añadirla a la pantalla de inicio. |

## Cómo está montado

- **GitHub Pages** enseña la app. El código es público, por eso no lleva ninguna contraseña.
- **Google Sheets** guarda los datos, en la pestaña «datos». La hoja no se comparte con nadie.
- **Apps Script** (`Codigo.gs`) hace de portero: comprueba la contraseña y es lo único que lee y escribe la hoja.

## Montarlo desde cero (una sola vez)

1. Crea una hoja de cálculo en blanco en Google Sheets, por ejemplo «Mi espacio (datos)». No la compartas.
2. En la hoja: **Extensiones → Apps Script**. Borra lo que haya, pega todo `Codigo.gs` y pulsa el disquete para guardar.
3. Vuelve a la hoja y recarga la página. Arriba aparece el menú **Mi espacio**.
4. **Mi espacio → 1. Preparar la hoja.** Google pedirá permiso: Continuar → elige tu cuenta → «Configuración avanzada» → «Ir a … (no seguro)» → Permitir. Es normal con un script propio.
5. **Mi espacio → 2. Poner o cambiar la contraseña.** Mínimo 8 caracteres.
6. **Mi espacio → 3. Conectar Google Calendar** (para el widget de hoy y los exámenes).
7. **Mi espacio → 4. Preparar la pestaña «añadir»** (opcional, para meter cosas desde la hoja).
8. En Apps Script: **Implementar → Nueva implementación** → engranaje → **Aplicación web**.
   - Ejecutar como: **Yo**
   - Quién tiene acceso: **Cualquier usuario**
   Pulsa Implementar y copia la **URL de la aplicación web** (acaba en `/exec`).
9. Pon esa dirección en `index.html`, en la línea `const SERVIDOR_URL = '…';`, y sube el archivo a GitHub.

«Nueva implementación» se usa **solo esta primera vez**. Después, nunca (mira abajo).

## Cambiar algo más adelante

- **La app** (`index.html` y demás): sube los archivos a GitHub. Antes, cambia `APP_VER` en `index.html` para saber qué versión tiene cada dispositivo.
- **El servidor** (`Codigo.gs`): pégalo en Apps Script, guarda y ve a **Implementar → Gestionar implementaciones → lápiz → Versión: Nueva versión → Implementar**.
  - **No uses «Nueva implementación»**: cambia la dirección `/exec` y deja de funcionar la app y los atajos de Siri hasta que pongas la dirección nueva en todos sitios.
- No edites la pestaña «datos» a mano. Para meter cosas desde la hoja, usa la pestaña «añadir».

## Usarlo

- En cada dispositivo abre la app y escribe la contraseña. Se queda conectado hasta que pulses **Configuración → Salir de este dispositivo** o pasen 60 días sin usarlo.
- En el iPhone: Safari → Compartir → **Añadir a pantalla de inicio**.
- Sin internet puedes seguir usándola: los cambios se envían solos cuando vuelve la conexión.
- Si dos dispositivos cambian la misma nota a la vez, se guardan las dos versiones (una con «(conflicto)» en el título).
- Atajos de Siri: **Configuración → Atajos de Siri**. Usan una clave propia, no la contraseña.

## Menú de la hoja

- 1. Preparar la hoja · 2. Contraseña · 3. Conectar Google Calendar · 4. Preparar la pestaña «añadir»
- Pasar ahora lo de «añadir» a la app
- Ver dispositivos conectados
- Cerrar todas las sesiones (por ejemplo, si pierdes el móvil)
- Limpiar borrados antiguos (quita filas de cosas borradas hace más de 30 días)
- Desactivar atajos de Siri

## Seguridad

- La contraseña no está en el código ni en GitHub. El script guarda solo una huella cifrada.
- Tras 5 intentos fallidos se bloquea la entrada 15 minutos.
- Tu cuenta de Google es la llave maestra de todo: actívale la verificación en dos pasos.
