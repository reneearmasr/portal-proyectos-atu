# Portal de Proyectos ATU

Página única que reúne los accesos a los proyectos de gestión presupuestal:

- Tablero de ejecución presupuestal 2026
- Programación y Modificación Presupuestal
- Previsiones Presupuestarias 2027–2034

## Logo
Sube el logo oficial de la ATU al repositorio con el nombre `logo-atu.png` (misma carpeta que `index.html`).
Si no está, se muestra el texto "ATU".

Es un acceso **adicional**: cada herramienta conserva su enlace e ingreso individual.

## Cómo publicarlo en GitHub Pages

1. En GitHub, crea un repositorio nuevo (por ejemplo `portal-proyectos-atu`).
2. Sube `index.html` y este `README.md` (botón **Add file → Upload files**).
3. Ve a **Settings → Pages**, en *Source* elige **Deploy from a branch**, rama `main`, carpeta `/ (root)` y guarda.
4. En 1–2 minutos el portal estará en `https://TU-USUARIO.github.io/portal-proyectos-atu/`.

## Cómo agregar o cambiar enlaces

Abre `index.html` (en GitHub puedes editar con el ícono del lápiz) y busca la lista `PROYECTOS`.
En cada proyecto pega el enlace en `url`:

```js
url: "https://tu-usuario.github.io/tablero-seguimiento/",
```

Si `url` queda vacío, la tarjeta aparece como **Enlace pendiente**.
Para agregar un proyecto nuevo, copia un bloque `{ ... }` y edita nombre, descripción y enlace.

> Nota: si el repositorio es público, la página del portal será visible para cualquiera con el enlace,
> pero cada herramienta seguirá pidiendo su propio ingreso.
