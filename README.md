# Conversor de Marcaciones (TXT → Excel)

Esta herramienta permite convertir archivos de texto (TXT/TSV) con marcaciones de reloj marcador en un archivo Excel (XLSX) profesional, con una pestaña por colaborador y cálculos automáticos de horas trabajadas mediante fórmulas de Excel.

## Características

- **Una pestaña por colaborador**: El sistema agrupa automáticamente las marcaciones por persona.
- **Cálculo automático**: Genera fórmulas de Excel para calcular el tiempo trabajado por la mañana, por la tarde y el total diario.
- **Fila de Resumen**: Incluye una fila de total general al final de cada hoja.
- **Formato Corporativo**: Fechas y horas formateadas según estándares (`dd/mm/yyyy`, `h:mm:ss`).
- **Procesamiento Seguro**: Los archivos se procesan localmente y no se almacenan en el servidor.

## Requisitos

- PHP 8.2 o superior.
- Composer.
- Extensiones de PHP: `xml`, `mbstring`, `zip`, `gd`, `curl`.

## Instalación

1. Clone el repositorio o descargue los archivos.
2. Instale las dependencias con Composer:
   ```bash
   composer install
   ```
3. Cree las carpetas necesarias y asigne permisos:
   ```bash
   mkdir -p var/ratelimit
   chmod -R 775 var
   ```
4. Inicie el servidor PHP apuntando a la carpeta `public`:
   ```bash
   php -S localhost:8000 -t public
   ```

## Cómo usar

1. Acceda a `http://localhost:8000` en su navegador.
2. Seleccione su archivo de marcaciones (TXT o TSV separado por TAB).
3. Haga clic en **"Generar Planilla Excel"**.
4. El navegador descargará automáticamente el archivo `marcaciones.xlsx`.

## Ejemplo de Archivo de Entrada (TXT/TSV)

El archivo debe estar separado por **TAB** y contener los siguientes encabezados (o variaciones comunes como "User ID", "Name", etc.):

```text
ID Colaborador	Nombre	Fecha	Hora
1015	Jose Leon	12/01/2026	09:47:08
1015	Jose Leon	12/01/2026	18:07:39
1016	Maximina Ara	06/01/2026	08:02:41
1016	Maximina Ara	06/01/2026	12:01:39
1016	Maximina Ara	06/01/2026	12:58:18
1016	Maximina Ara	06/01/2026	17:01:07
```

*Nota: Se aceptan hasta 4 marcaciones por día. Si hay más, se ignoran las excedentes.*

## Ejemplo de Salida (Excel)

Cada hoja del archivo Excel tendrá la siguiente estructura:

| Nombre | Fecha | 1 | 2 | 3 | 4 | Mañana | Tarde | Total |
| :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- |
| Jose Leon | 12/01/2026 | 09:47:08 | 18:07:39 | | | 08:20:31 | 00:00:00 | 08:20:31 |
| ... | ... | ... | ... | ... | ... | ... | ... | ... |
| | | | | | | | **Total** | **SUM(...)** |

## Soporte para Formatos Flexibles

El sistema es capaz de detectar automáticamente diferentes nombres de columnas:
- **ID**: "ID Colaborador", "User ID", "Enroll ID", "ID".
- **Nombre**: "Nome", "Name", "Nombre".
- **Fecha**: "Data", "Date", "Fecha".
- **Hora**: "Hora", "Time".
- **Marcaciones Directas**: Si el archivo ya tiene columnas numeradas "1", "2", "3", "4", el sistema las importará directamente.

---
© 2025 Conversor RH.
