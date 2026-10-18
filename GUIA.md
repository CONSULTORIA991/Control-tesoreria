# Control de Tesorería en GitHub: guía paso a paso

Esta carpeta tiene una **copia** de la app de claude.ai, preparada para funcionar solo con GitHub, sin ningún otro servicio. La versión de claude.ai sigue igual y funcionando; así pueden probar las dos y escoger con cuál trabajar.

- **Repositorio:** https://github.com/CONSULTORIA991/Control-tesoreria
- **Dirección de la app** (cuando actives GitHub Pages): https://consultoria991.github.io/Control-tesoreria/

## Qué hay en la carpeta

| Carpeta / archivo | Para qué sirve | ¿Se sube a GitHub? |
|---|---|---|
| `SUBIR-A-GITHUB/` | Lo único que va al repositorio: `index.html` (cifrado) y `README.md` | **Sí** |
| `NO-SUBIR/fuente/app.html` | La app sin cifrar | **No, nunca** |
| `NO-SUBIR/cifrador.html` | Herramienta para cifrar la app y manejar usuarios y contraseñas | No |
| `NO-SUBIR/version-claude-original.html` | Respaldo de la versión de claude.ai, sin cambios | No |

## Cómo funciona y cómo queda protegida

- **La app** está en la rama `main` del repositorio, como un solo `index.html` **cifrado** (AES-256). Quien lo abra sin clave solo ve la pantalla de acceso.
- **Los datos** (pagos, ingresos, movimientos del banco, soportes, actividad) se guardan en la rama `datos` del mismo repositorio. **Cada archivo está cifrado** y tiene un nombre ilegible, así que en el repositorio público solo se ven archivos sin sentido.
- **Cada persona** entra con su propio usuario y contraseña. Los creas tú en el cifrador.
- **La app guarda en GitHub con un token** que creas una sola vez. El token va dentro de la app cifrada; solo quien entra con usuario y contraseña puede usarlo.
- **Quién hizo qué:** cada registro guarda quién lo creó o modificó y cuándo, y la pestaña Actividad muestra los tiempos de cada persona.

**Límites que debes conocer**
- Los permisos por rol los aplica la pantalla, igual que en claude.ai. Como todos usan el mismo token, alguien con conocimientos técnicos y una contraseña válida podría saltárselos. Para un equipo de confianza es suficiente.
- Los cambios de otras personas aparecen en un minuto como máximo, o al volver a la pestaña.
- La lectura automática de comprobantes no está incluida: Jesús escribe el tercero y el valor a mano.
- **Usa contraseñas largas** (12 o más caracteres). Como el archivo cifrado es público, la seguridad depende de que sean difíciles de adivinar.

---

## Paso 1: Crear el token de GitHub (una vez, 5 minutos)

1. En GitHub, haz clic en tu foto (arriba a la derecha) → **Settings**.
2. Abajo a la izquierda: **Developer settings** → **Personal access tokens** → **Fine-grained tokens** → **Generate new token**.
3. Llena así:
   - **Token name:** `tesoreria-app`
   - **Expiration:** la más larga que te deje (por ejemplo 1 año). Anota la fecha: ese día habrá que crear otro.
   - **Repository access:** **Only select repositories** → elige `Control-tesoreria`.
   - **Permissions → Repository permissions → Contents:** **Read and write**. No actives nada más.
4. Presiona **Generate token** y **copia el token** (empieza por `github_pat_`). GitHub no lo vuelve a mostrar. No lo pegues en ningún otro lado: solo en el cifrador.

## Paso 2: Cifrar la app

1. Abre `NO-SUBIR/cifrador.html` con doble clic; se abre en Chrome o Edge.
2. Deja marcado **Instalación nueva**.
3. Pega el token y presiona **Probar conexión con GitHub**. Debe decir "Conexión correcta".
4. En **Archivo de la app**, elige `NO-SUBIR/fuente/app.html`.
5. En **Personas que pueden entrar**, escribe un usuario y una contraseña para cada persona: tú, Jesús, Cristian y Laura. Anótalas para entregárselas a cada uno en privado.
6. Presiona **Generar index.html cifrado**. Se descarga `index.html` en tu carpeta de Descargas.
7. Mueve ese `index.html` a la carpeta `SUBIR-A-GITHUB`.

## Paso 3: Subirlo y publicarlo

1. Abre el repositorio → **Add file → Upload files**.
2. Arrastra `SUBIR-A-GITHUB/index.html`. **No arrastres la carpeta NO-SUBIR.** Presiona **Commit changes**.
3. Ve a **Settings → Pages**. En **Source** elige **Deploy from a branch**, en **Branch** elige `main` y `/ (root)`, y presiona **Save**.
4. Espera uno o dos minutos y abre https://consultoria991.github.io/Control-tesoreria/

## Paso 4: Primer ingreso y roles

1. Entra tú primero: quedas como administradora. Al entrar por primera vez, la app crea sola la rama `datos`.
2. Pide a Jesús, Cristian y Laura que entren una vez. Cada uno verá un aviso de acceso pendiente.
3. En la pestaña **Usuarios**, escribe su nombre y asígnale su rol: Pagos, Ingresos o Movimientos de banco.

## Cuando cambie algo

Abre el cifrador y elige **Actualizar**. Elige el `index.html` publicado (el de `SUBIR-A-GITHUB`) y entra con tu usuario y contraseña de administradora. Desde ahí puedes:
- **Agregar una persona:** "Agregar persona" con su usuario y contraseña.
- **Cambiar una contraseña:** escribe la nueva. Si dejas la contraseña vacía, la persona conserva la actual.
- **Quitar a alguien:** presiona ✕ en su fila. Si se fue de la empresa en malos términos, crea también un token nuevo (paso 1), revoca el anterior en GitHub y pégalo en "Token nuevo".
- **Renovar el token vencido:** pega el nuevo en "Token nuevo".
- **Nueva versión de la app:** elige el nuevo `app.html`.

Después presiona **Generar**, mueve el nuevo `index.html` a `SUBIR-A-GITHUB` y súbelo al repositorio (**Add file → Upload files**), reemplazando el anterior. Los datos no se pierden.

## Diferencias con la versión de claude.ai

| | claude.ai | GitHub |
|---|---|---|
| Entrada | Cuenta de Claude de cada persona | Usuario y contraseña que creas en el cifrador |
| Datos | En claude.ai | Cifrados en la rama `datos` del repositorio |
| Lectura automática de comprobantes | Incluida | No incluida |
| Actualización de pantallas | Al instante | En un minuto como máximo |

Los datos de las dos versiones están separados: lo que se registra en una no aparece en la otra.
