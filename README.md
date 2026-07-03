# Combinaciones por Dirección — Energy Parts

App web para generar combinaciones de grupos G1–G7 por dirección de entrega,
a partir de un archivo Excel con puntos disponibles.

## Uso

1. Sube tu Excel (tabla dinámica: Etiquetas de fila · Suma de Puntos)
2. Selecciona las direcciones a procesar
3. Elige cuántas combinaciones por dirección (10/15/25/50)
4. Descarga el Excel con todas las combinaciones

## Deploy en Render

```bash
git init && git add . && git commit -m "Initial commit"
git remote add origin https://github.com/TU_USUARIO/combinaciones-direcciones.git
git push -u origin main
```

En Render → New → Web Service → conectar repo → Deploy.
**Sin variables de entorno requeridas.**

## Logo
Coloca `static/logo.png` con el logo de Energy Parts.
