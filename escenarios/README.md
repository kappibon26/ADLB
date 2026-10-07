# Escenarios · Biblioteca adlb.

Fondos vectoriales (SVG, 1080×1920, 9:16) para los reels con los maniquíes.

| Archivo | Escena |
|---|---|
| `01-parque.svg` | Parque con banca, farol y ciudad al fondo |
| `02-habitacion.svg` | Habitación con cama, ventana, skate y mochila |
| `03-oficina.svg` | Oficina BIM: escritorio, monitor con modelo, afiche |
| `04-obra.svg` | Obra: estructura, grúa, caseta, malla y cono |
| `05-paradero.svg` | Paradero con valla "La disciplina también es un proyecto" |
| `06-atardecer.svg` | Terraza al atardecer: termo, portátil, libreta |

## Cómo usarlos

- **Piso de los maniquíes:** y = 1362, el mismo del animatic Nº01. Todas las escenas tienen el suelo a esa altura.
- **Capas para parallax:** cada SVG trae grupos `cielo`, `fondo`, `medio` y `primer-plano`. En After Effects o Illustrator aparecen como capas separadas.
- **Zonas seguras de Instagram:** 220 px arriba, 420 px abajo y 130 px a la derecha.
- **Textos de afiches:** usan Archivo condensada. Si el programa no la tiene, el texto se ajusta a su ancho igual. Para entregar, conviértelos a contornos.

## Regenerar

```
python3 escenarios/generar.py
```

Colores, posiciones y piezas (árboles, nubes, ciudad, banca) están en ese script.
