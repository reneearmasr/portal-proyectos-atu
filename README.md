# Portal de Proyectos ATU

Página única que reúne los accesos a los proyectos de gestión presupuestal de la ATU:
Tablero de ejecución 2026, Programación y Modificación Presupuestal, Previsiones 2027–2034,
Base de datos de locadores y POI Multianual.

Es un acceso **adicional**: cada herramienta conserva su enlace e ingreso individual.

## Archivos

- `index.html`: diseño y comportamiento de la página (no hace falta tocarlo para cambiar proyectos).
- `proyectos.json`: listado de proyectos, temas, formas de acceso y etapas del ciclo presupuestal.
- `logo-atu.png` (opcional): logo oficial. Si no está, se muestra el texto "ATU".

## Cómo agregar o cambiar un proyecto

En GitHub abre `proyectos.json`, pulsa el lápiz (Edit) y modifica la lista `proyectos`.
Cada proyecto tiene esta forma:

```json
{
  "id": "tablero-2026",
  "nombre": "Tablero de ejecución presupuestal 2026",
  "descripcion": "Texto corto que aparece en la tarjeta.",
  "detalle": ["Punto 1 de la ficha", "Punto 2 de la ficha"],
  "url": "https://reneearmasr.github.io/Tablero_ATU/",
  "acceso": "libre",
  "categoria": "ejecucion",
  "etiqueta": "2026",
  "icono": "chart",
  "estilo": "v1"
}
```

- `id`: nombre corto sin espacios; se usa en `ciclo` para ubicar el proyecto en una etapa.
- `acceso`: `libre`, `usuario` o `gmail` (definidos en `accesos`).
- `categoria`: `planeamiento`, `presupuesto`, `ejecucion` o `gestion` (definidos en `categorias`).
- `icono`: `chart`, `calendar`, `forecast`, `people` o `target`.
- `estilo` (color): `v1` azul, `v2` verde, `v3` dorado, `v4` morado, `v5` rojo.
- Si `url` queda vacío (`""`), la tarjeta aparece como **Enlace pendiente**.

Para que un proyecto aparezca al hacer clic en una etapa del ciclo, agrega su `id` en
`ciclo[...].proyectos`.

**Cuidado con el formato JSON**: comillas dobles, comas entre elementos y sin coma después del
último. Si el archivo queda con un error, la página muestra un aviso (y la última versión válida que
el navegador tenga guardada). Se puede validar pegando el contenido en https://jsonlint.com.

## Publicación (GitHub Pages)

Settings → Pages → *Deploy from a branch*, rama `main`, carpeta `/ (root)`.
Los cambios se ven en 1–2 minutos en https://reneearmasr.github.io/portal-proyectos-atu/
