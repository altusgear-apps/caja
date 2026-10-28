# Instalación — Caja Altus

Para Fernando. Una sola vez, unos 20 minutos. Hay tres partes: la hoja con su motor, la app en GitHub y el teléfono.

Todo se hace con la cuenta **altusgear@gmail.com**, no con la personal, para que la caja quede como activo de la empresa.

---

## Parte A · La hoja y su motor (10 min)

### A1. Crear la hoja
1. Entra a sheets.google.com con la cuenta de Altus.
2. Crea una hoja en blanco y nómbrala **Altus_Caja_2026**.
3. Muévela a `Altus 2026 / 05_Soporte / Administracion_y_Contabilidad/`.

Que viva ahí importa por dos razones. Junto a la hoja se crean dos carpetas: `fotos_caja/` (los comprobantes) y `buzon_claude/` (por donde Claude clasifica gastos y completa lo leído en fotos). Y tu cuenta personal debe poder escribir en esa carpeta, porque Claude trabaja con ella.

### A2. Pegar el motor
1. En la hoja: **Extensiones → Apps Script**.
2. Borra todo lo que trae `Código.gs` y pega el contenido de `apps-script/Codigo.gs`.
3. Guarda (ícono de disco).

### A3. Configurar
1. Regresa a la hoja y recárgala. Aparece el menú **Altus Caja**.
2. **Altus Caja → Configurar (primera vez)**.
3. Google pide autorización. Si dice "Google no verificó esta app": **Configuración avanzada → Ir a … (no seguro)**. Es tu propio código en tu propia cuenta.
4. Al terminar aparece la **clave de la app** (formato `XXXX-XXXX`). Anótala. Se puede volver a ver en **Altus Caja → Ver clave de la app**.

Configurar crea las hojas `MOVIMIENTOS`, `CONCEPTOS`, `CATALOGO`, `METAS`, `RESUMEN` y `BITACORA`, las dos carpetas, y una revisión automática del buzón cada hora.

### A4. Publicar el motor
1. En Apps Script: **Implementar → Nueva implementación**.
2. Tipo (engrane): **Aplicación web**.
3. Ejecutar como: **Yo (altusgear@gmail.com)**.
4. Quién tiene acceso: **Cualquier usuario**.
5. **Implementar** y copia la dirección. Termina en `/exec`.

"Cualquier usuario" es necesario porque la app del teléfono no inicia sesión en Google. Lo que protege los datos es la clave: sin ella, la hoja no responde.

---

## Parte B · La app en GitHub (5 min)

Si conectaste GitHub en Claude, esta parte la hago yo. Si no:

1. Entra al repositorio `caja`.
2. **Add file → Upload files** y arrastra todo el contenido de esta carpeta: `index.html`, `sw.js`, `manifest.webmanifest`, `README.md`, `GUIA_INSTALACION.md` y las carpetas `icons/` y `apps-script/`.
3. **Commit changes**.
4. **Settings → Pages**. En *Source* elige **Deploy from a branch**, rama **main**, carpeta **/ (root)**, y **Save**.
5. En uno o dos minutos la app queda en `https://altusgear.github.io/caja/` (cambia `altusgear` si el usuario quedó con otro nombre).

---

## Parte C · Cada teléfono (3 min por persona)

1. Abre la dirección de la app **en Safari** (en Android, en Chrome).
2. La primera pantalla es **Ajustes**. Pega la dirección `/exec` de la Parte A, escribe la clave y elige quién usa el teléfono.
3. **Guardar y probar conexión**. Debe decir "Conectado con la hoja".
4. Instálala: en Safari, **Compartir → Agregar a inicio**. Desde ahí en adelante, ábrela siempre desde ese ícono.

### Prueba de modo avión (hacerla en el teléfono de Alberto)
1. Abre la app desde el ícono, con señal, una vez.
2. Ciérrala. Activa **modo avión**.
3. Ábrela desde el ícono. Debe abrir y decir "Sin señal".
4. Registra un gasto de prueba de $1 con foto. Debe decir "Sin señal · 1 por subir".
5. Quita el modo avión. En unos segundos debe decir "Al día", y el gasto aparece en la hoja.
6. Elimina el gasto de prueba desde **Movimientos**.

Si el paso 3 falla, la app no quedó guardada en el teléfono. Avísale a Claude antes de usarla en el mercado.

---

## Cómo se usa la hoja

- **MOVIMIENTOS no se edita a mano.** Solo escriben la app y el buzón de Claude. Lo eliminado se marca como `eliminado`; no se borra, para que quede rastro.
- **CONCEPTOS:** la columna `clase` (fijo, semifijo, variable) sí se puede cambiar a mano. El cambio se aplica solo a todos los gastos con ese concepto.
- **METAS:** se pueden editar. El resumen de la app y la hoja `RESUMEN` las leen de ahí.
- **CATALOGO:** líneas, subcategorías, canales y formas de pago. Se edita con cuidado: cada valor es una lista en formato JSON. Mejor pídeselo a Claude.
- **RESUMEN:** P&L preliminar por mes con fórmulas. Si se rompe, **Altus Caja → Reconstruir resumen**.

## El buzón de Claude

Claude no puede editar la hoja directamente. En su lugar deja archivos en `buzon_claude/` con lo que hay que aplicar: clasificar un concepto de gasto o completar un movimiento que se guardó solo con foto. El motor los aplica cada hora y cada vez que la app se sincroniza, y los mueve a `buzon_claude/procesados/`. Cada aplicación queda registrada en `BITACORA`. Para aplicarlos al momento: **Altus Caja → Procesar buzón de Claude ahora**.

## Actualizaciones

- **App:** se suben los archivos nuevos al repositorio. Cada teléfono toma la versión nueva la siguiente vez que abre con señal (a veces hace falta abrirla dos veces).
- **Motor:** se pega el `Codigo.gs` nuevo y luego **Implementar → Gestionar implementaciones → ✏️ → Versión: Nueva versión → Implementar**. La dirección `/exec` no cambia y nadie tiene que reconfigurar su teléfono.
