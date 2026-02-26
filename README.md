# Science Lab Notebook

Aplicación web académica para redactar informes de laboratorio con formato institucional y descarga directa en PDF.

## Funcionalidades principales

- Diseño institucional verde profesional (fondo verde oscuro + tarjetas blancas).
- Información obligatoria del estudiante:
  - `Title of Experiment`
  - `Student Name`
  - `Date`
- 12 secciones del informe (texto + tablas editables para `Raw Data` y `Processed Data`).
- En `Processed Data` se incluye `Sample Calculations`.
- Restricción anti copy-paste:
  - Bloquea `Ctrl+V / Cmd+V`
  - Bloquea evento `paste`
  - Bloquea clic derecho/context menu
  - Muestra alerta: `Paste is disabled. Please write your own work.`
- Autosave dual:
  - Cada 15 segundos
  - 3 segundos después de dejar de escribir
  - Guarda en base de datos (Supabase) y en `localStorage` como respaldo offline.
- Botones visibles: `Save Draft` y `Download Final Report`.
- Estados de documento:
  - `Draft`
  - `Submitted` (bloquea edición)
- Entrega final:
  - Genera PDF académico con márgenes de 1 inch.
  - Incluye solo secciones con contenido.
  - Numera secciones automáticamente.
  - Dibuja tablas con líneas visibles.
  - Descarga automática del PDF al estudiante.

## Stack

- Frontend: HTML + CSS + JavaScript moderno
- Backend: Node.js + Express
- Base de datos: Supabase (con fallback en memoria para desarrollo sin credenciales)
- PDF: PDFKit

## Instalación local

1. Entrar al proyecto:

```bash
cd libreta-de-laboratorio
```

2. Instalar dependencias:

```bash
npm install
```

3. Crear `.env` desde `.env.example`:

```bash
cp .env.example .env
```

4. Configurar Supabase en `.env`.

5. (Supabase) Crear tabla ejecutando:

Archivo: `sql/supabase_schema.sql`

6. Levantar servidor:

```bash
npm run dev
```

7. Abrir:

`http://localhost:3000`

## Variables de entorno

- `PORT`: puerto del servidor.
- `SUPABASE_URL`: URL del proyecto Supabase.
- `SUPABASE_SERVICE_ROLE_KEY`: key con permisos para insertar/actualizar.
- `SUPABASE_TABLE`: nombre de la tabla (default `lab_reports`).

## Subir a GitHub

```bash
git add .
git commit -m "feat: science lab notebook with autosave and local PDF download"
git push
```
