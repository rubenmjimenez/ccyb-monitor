# Monitor mensual del CCyB (ESRB)

Avisa cada mes por email si algún país va a cambiar su colchón de capital
anticíclico (CCyB) dentro del próximo mes, usando el Excel oficial que
publica la ESRB.

## Recomendado: haz todo desde GitHub, sin depender de tu PC

Si tu Python local da problemas (SSL, permisos, antivirus corporativo,
etc. — algo habitual en equipos de empresa), no hace falta que instales
nada localmente. Sube el proyecto a GitHub (ver sección 4) y usa la
pestaña **Actions** para todo, incluida la calibración:

1. Ve a la pestaña **Actions** de tu repo → workflow **"CCyB monthly
   check"** → **Run workflow**.
2. En el desplegable **mode**, elige **diagnose** → **Run workflow**.
3. Espera a que termine (un círculo verde) y abre el log del job para ver
   las columnas detectadas, igual que si lo hubieras ejecutado en tu PC.
4. Si necesitas ajustar `COLUMN_HINTS` en `ccyb_monitor.py`, edítalo
   directamente en GitHub (botón del lápiz) y vuelve a lanzar el workflow
   en modo **diagnose** para comprobarlo.
5. Cuando esté bien, prueba una vez en modo **dry-run** (hace todo el
   análisis pero no envía email ni guarda nada) y luego ya puedes lanzarlo
   en modo **check** (el normal, el que se ejecutará solo cada mes).

Con esto tu PC no interviene en ningún momento: todo corre en la nube de
GitHub. Puedes saltarte por completo las secciones 1 y 2 de abajo.

---

## 1. (Alternativa local) Calibrar el lector del Excel

Solo si prefieres ejecutar el script en tu propio PC y tienes un Python
que funcione con normalidad (con SSL operativo):

```bash
pip install -r requirements.txt
python ccyb_monitor.py --diagnose
```

Esto imprime el nombre real de las columnas del Excel y qué columna ha
asignado el script a cada categoría (país, tasa actual, fecha actual, tasa
futura, fecha futura). Si alguna sale como "NO ENCONTRADA", abre
`ccyb_monitor.py`, busca el diccionario `COLUMN_HINTS` cerca del principio
del archivo, y añade ahí la palabra exacta que has visto impresa (en
minúsculas y sin tildes) en la categoría que corresponda.

### Si te da un error de SSL ("SSL module is not available")

No es un problema de certificados reales, es que tu instalación de Python
no tiene cargadas las DLLs de OpenSSL (`libssl-3-x64.dll` /
`libcrypto-3-x64.dll` para Python 3.12/3.13). Es muy típico en equipos
corporativos donde el antivirus o una política de IT las bloquea, o donde
Python se instaló de forma incompleta. La forma más simple de arreglarlo
es reinstalar Python desde el instalador oficial de
[python.org](https://www.python.org/downloads/) (no la Microsoft Store),
marcando la casilla "Add python.exe to PATH" — el instalador oficial
incluye esas DLLs. Si no tienes permisos de administrador en el equipo
para reinstalar, usa la opción de GitHub Actions de arriba, que evita el
problema por completo.

## 2. Probar sin enviar nada

```bash
python ccyb_monitor.py --dry-run
```

Descarga el Excel, hace todo el análisis y te imprime por pantalla el
email que enviaría, pero no lo manda ni actualiza el snapshot guardado.

## 3. Configurar el envío de email

El script necesita estas variables de entorno:

| Variable        | Qué es                                              |
|-----------------|------------------------------------------------------|
| `MAIL_TO`       | Tu correo corporativo (el destinatario)              |
| `MAIL_FROM`     | Dirección remitente (puede ser un "no-reply")        |
| `SMTP_HOST`     | Servidor SMTP                                        |
| `SMTP_PORT`     | Normalmente 587                                      |
| `SMTP_USER`     | Usuario SMTP                                         |
| `SMTP_PASSWORD` | Contraseña / contraseña de aplicación                |

### Opción A — Microsoft 365 / Outlook directamente

- `SMTP_HOST=smtp.office365.com`, `SMTP_PORT=587`
- `SMTP_USER` = tu cuenta de Microsoft 365 completa (ej. `tu.nombre@empresa.com`)
- `SMTP_PASSWORD` = una **contraseña de aplicación**, no tu contraseña normal

**Importante:** muchos tenants de Microsoft 365 tienen deshabilitado por
defecto el "SMTP AUTH" por seguridad (Microsoft lo desactivó por defecto
desde 2022 y muchas empresas no lo reactivan). Si al probar te da un error
de autenticación, es casi seguro que sea esto — tendrías que pedirle a tu
departamento de IT que active "SMTP AUTH client submission" para tu buzón.
No es algo que puedas activar tú mismo si no eres administrador del tenant.

### Opción B — Servicio de envío gratuito (recomendada si la A no es viable)

Si no quieres depender de que IT active nada, la forma más sencilla y
fiable es usar un servicio externo gratuito solo para el envío, con tu
correo corporativo como **destinatario** (`MAIL_TO`), sin tocar nada de la
configuración de tu empresa:

- **Brevo (antes Sendinblue)**: plan gratuito con 300 emails/día.
  `SMTP_HOST=smtp-relay.brevo.com`, `SMTP_PORT=587`.
- **Gmail con contraseña de aplicación**: si tienes o creas una cuenta de
  Gmail personal solo para esto.
  `SMTP_HOST=smtp.gmail.com`, `SMTP_PORT=587`.

En ambos casos el correo llega perfectamente a tu bandeja corporativa; lo
único que cambia es qué servidor lo envía.

## 4. Programarlo cada mes con GitHub Actions (recomendado)

1. Crea un repositorio nuevo en GitHub (puede ser privado) y sube todos
   estos archivos tal cual (mantén la carpeta `.github/workflows/` y
   `data/`).
2. Ve a **Settings → Secrets and variables → Actions → New repository
   secret** y crea estos 6 secrets con los valores de la sección 3:
   `MAIL_TO`, `MAIL_FROM`, `SMTP_HOST`, `SMTP_PORT`, `SMTP_USER`,
   `SMTP_PASSWORD`.
3. Ve a la pestaña **Actions**, elige el workflow "CCyB monthly check" y
   pulsa "Run workflow" una vez para probar que todo funciona.
4. A partir de ahí se ejecutará solo, automáticamente, el día 1 de cada
   mes a las 08:00 UTC (puedes cambiar la hora/día editando la línea
   `cron:` en `.github/workflows/monthly_check.yml`).

No necesitas dejar tu ordenador encendido ni ningún servidor propio: se
ejecuta en la nube de GitHub de forma gratuita.

### Alternativa: en tu propio PC/servidor

Si prefieres no usar GitHub, puedes programar la ejecución con el
**Programador de tareas de Windows** o con **cron** (Linux/Mac), llamando
mensualmente a:

```bash
python ccyb_monitor.py
```

con las variables de entorno de la sección 3 exportadas en el sistema (o
un fichero `.env` cargado antes de ejecutar el script).

## ¿Qué hace exactamente cada mes?

1. Descarga el Excel oficial de la ESRB.
2. Compara los datos con la última vez que se comprobó (`data/last_snapshot.json`)
   para saber si la ESRB ha publicado algo nuevo.
3. Busca cambios de tasa cuya fecha de entrada en vigor caiga dentro de los
   próximos 35 días.
4. Envía un único email con el resultado, por ejemplo:

   > Polonia subirá el CCyB del 1% al 2%, efectivo desde el 30 de
   > septiembre de 2026.

   o, si no hay nada:

   > No se han detectado cambios de CCyB en ningún país a un mes vista.
