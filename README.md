
# QIK PDF to XLSX

Herramienta para convertir estados de cuenta de **QIK Banco Digital (República Dominicana)** de PDF a XLSX.

> **Aviso:** Este es un proyecto independiente y no oficial. No tengo ninguna relación, afiliación ni asociación con QIK Banco Digital.

## 🤔 ¿Por qué?

QIK permite exportar los estados de cuenta en formato PDF, pero no ofrece una exportación directa a XLSX o CSV. Esto dificulta el procesamiento, análisis y organización de la información de los movimientos.

Este script convierte el PDF del estado de cuenta a un archivo **XLSX**, extrayendo la tabla de movimientos e ignorando el resto del contenido del documento.

## 📋 Requisitos

- Python 3.10+
- Camelot
- Pandas
- OpenPyXL

Instala las dependencias con:

```bash
pip install camelot-py pandas openpyxl
```

Dependiendo del sistema, Camelot puede requerir dependencias adicionales para trabajar con PDFs.

## 🚀 Uso

El script requiere un archivo PDF de entrada. El archivo de salida es opcional.

### Sin especificar output

```bash
python ./convert.py ./some/file.pdf
```

Si el PDF se llama:

```text
./some/file.pdf
```

el script creará automáticamente:

```text
./file.xlsx
```

El archivo se crea en el mismo directorio donde se encuentra el script.

### Especificando output

También puedes indicar dónde quieres guardar el archivo XLSX:

```bash
python ./convert.py ./some/file.pdf ./out/file.xlsx
```

Si la carpeta de destino no existe, el script la crea automáticamente.

### Ejemplo

```bash
python ./convert.py ./statements/julio-2026.pdf ./output/julio-2026.xlsx
```

El archivo generado contiene los movimientos del estado de cuenta en una hoja llamada `Movimientos`.

## ⚠️ Limitaciones

Este script está diseñado **específicamente para el formato de los estados de cuenta de QIK Banco Digital**.

Si QIK cambia el formato o diseño de sus estados de cuenta, el script podría requerir modificaciones.

## ℹ️ Disclaimer

Este proyecto no es oficial y no está afiliado, patrocinado, respaldado ni aprobado por QIK Banco Digital.

El nombre y las marcas de QIK pertenecen a sus respectivos propietarios.
