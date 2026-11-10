# Verbarium

Aplicación web responsive para estudiar los 400 verbos del archivo `VERBS.pdf`.

## Ejecutar

La aplicación usa módulos ES y `fetch`, por lo que debe abrirse mediante un servidor HTTP local.

### Opción simple con Python

```bash
python -m http.server 4173
```

Después abre `http://localhost:4173`.

## Datos

- `public/data/verbs.json`: 400 verbos extraídos del PDF.
- `public/data/grammar.json`: reglas de uso de 12 tiempos verbales.
- `public/data/app-config.json`: configuración y alias para encabezados de Excel/CSV.

## Importación

La aplicación acepta `.xlsx` moderno (Office Open XML) y `.csv`. La primera fila define los encabezados. Columnas conocidas se mapean al modelo y columnas adicionales se conservan como campos dinámicos. Los datos importados se guardan en `localStorage` del navegador.

Encabezados recomendados:

`NUM, TYPE, SIMPLE FORM, THIRD PERSON, SIMPLE PAST, PAST PARTICIPLE, GERUND, MEANING`

Se admiten alias en español como `VERBO`, `TIPO`, `PASADO SIMPLE`, `PARTICIPIO`, `GERUNDIO`, `SIGNIFICADO`.

## Nota

Los datos base preservan el contenido del PDF proporcionado. La app no corrige automáticamente posibles erratas del documento fuente. Las equivalencias de tiempos en español son orientativas y explican la estructura gramatical.
