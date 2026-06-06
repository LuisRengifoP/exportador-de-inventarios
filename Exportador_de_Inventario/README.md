# Exportador Inteligente de Inventarios

Sistema de exportación y transformación de archivos Excel que permite cargar documentos de inventario con estructuras diferentes o desorganizadas y reorganizarlos automáticamente según la estructura estándar definida.

## Características Principales

- **Carga múltiple de archivos** Excel (.xlsx, .xls)
- **Análisis automático** de la estructura de cada archivo
- **Mapeo inteligente** de columnas con detección de similitud
- **Transformación automática** al formato estándar
- **Generación de Excel** estilizado y listo para importar
- **Validación y advertencias** de columnas faltantes
- **Interfaz moderna** con drag & drop
- **Diseño responsive** para móvil y escritorio

##  Estructura del Proyecto

```
exportador_inventarios/
├── app.py                  # Aplicación principal Flask
├── requirements.txt        # Dependencias Python
├── README.md              # Este archivo
├── static/                # Archivos estáticos (CSS, JS, imágenes)
├── templates/             # Plantillas HTML
│   └── index.html         # Interfaz principal
├── uploads/               # Archivos subidos temporalmente
├── outputs/               # Archivos procesados generados
└── modelo/                # Archivo de estructura de referencia
    └── Estructura de distriapp.xlsx
```

## Instalación y Configuración

### Paso 1: Preparar el entorno

**Opción A: Usar Python directamente**

```bash
# 1. Crear entorno virtual (recomendado)
python -m venv venv

# 2. Activar entorno virtual
# En Windows:
venv\Scripts\activate
# En macOS/Linux:
source venv/bin/activate

# 3. Instalar dependencias
pip install -r requirements.txt
```

**Opción B: Usar Conda**

```bash
# 1. Crear entorno Conda
conda create -n exportador python=3.11

# 2. Activar entorno
conda activate exportador

# 3. Instalar dependencias
pip install -r requirements.txt
```

### Paso 2: Copiar el archivo modelo

Copia el archivo **"Estructura de distriapp.xlsx"** (el archivo de referencia con la estructura objetivo) a la carpeta `modelo/` del proyecto:

```bash
mkdir -p modelo
cp "ruta/a/Estructura de distriapp.xlsx" modelo/
```

### Paso 3: Ejecutar la aplicación

```bash
# Asegúrate de estar en la carpeta del proyecto
cd exportador_inventarios

# Ejecutar la aplicación
python app.py
```

La aplicación estará disponible en: **http://localhost:5000**

## 📖 Guía de Uso

### 1. Cargar Archivo
- Arrastra y suelta archivos Excel en la zona de carga
- O haz clic en "Seleccionar Archivo" para elegir desde tu computadora
- Formatos soportados: `.xlsx` y `.xls`

### 2. Revisar Mapeo
- El sistema analiza automáticamente las columnas del archivo
- Muestra una tabla con el mapeo sugerido (columna origen → columna destino)
- Indicador de confianza: verde (alta), amarillo (media), rojo (baja)
- Puedes ajustar manualmente el mapeo si es necesario

### 3. Procesar
- Haz clic en "Procesar Archivo"
- El sistema transforma los datos al formato estándar
- Se muestra una vista previa de los primeros registros

### 4. Descargar
- Haz clic en "Descargar Resultado"
- El archivo Excel estandarizado se descarga automáticamente
- Listo para importar al sistema de inventarios

## 🔧 Estructura Objetivo

El sistema transforma los datos a las siguientes columnas estándar:

| Columna | Tipo | Descripción |
|---------|------|-------------|
| `sku` | Texto | Código único del producto |
| `locale` | Texto | Idioma/región (default: "es") |
| `attribute_family_code` | Texto | Familia de atributos (default: "default") |
| `type` | Texto | Tipo de producto (default: "simple") |
| `categories` | Texto | Categorías del producto |
| `name` | Texto | Nombre del producto |
| `description` | Texto | Descripción larga |
| `short_description` | Texto | Descripción corta |
| `status` | Entero | Estado (1=activo, 0=inactivo) |
| `visible_individually` | Entero | Visible individual (1=sí, 0=no) |
| `brand` | Texto | Marca del producto |
| `guest_checkout` | Entero | Permite compra invitado (1=sí, 0=no) |
| `price` | Entero | Precio del producto |
| `Iva` | Entero | Porcentaje de IVA (default: 19) |
| `url_key` | Texto | Clave URL amigable |

## Algoritmo de Mapeo Inteligente

El sistema utiliza múltiples estrategias para identificar columnas:

1. **Coincidencia exacta normalizada**: Compara nombres ignorando mayúsculas, acentos y espacios
2. **Similitud de secuencias**: Usa SequenceMatcher para encontrar coincidencias parciales
3. **Diccionario de aliases**: Conoce múltiples variantes de nombres comunes en español e inglés
4. **Inferencia por contenido**: Analiza los valores de muestra para determinar el tipo de dato
5. **Heurísticas específicas**: Reglas especiales para precios, códigos, nombres, etc.

## ⚙️ Configuración Avanzada

### Personalizar aliases de columnas

Edita el diccionario `COLUMN_ALIASES` en `app.py` para agregar más variantes de nombres de columnas:

```python
COLUMN_ALIASES = {
    'sku': ['sku', 'codigo', 'código', 'id', 'referencia', ...],
    # ...
}
```

### Modificar valores por defecto

Edita el diccionario `DEFAULT_VALUES` en `app.py`:

```python
DEFAULT_VALUES = {
    'locale': 'es',
    'Iva': 19,
    'status': 1,
    # ...
}
```

## Seguridad

- Los archivos subidos se almacenan temporalmente en `uploads/`
- Los archivos procesados se guardan en `outputs/`
- Límite de tamaño: 50MB por archivo
- Se recomienda limpiar periódicamente las carpetas `uploads/` y `outputs/`

## Solución de Problemas

### Error "No module named 'flask'"
```bash
pip install -r requirements.txt
```

### Error al leer archivo Excel
- Asegúrate de que el archivo no esté corrupto
- Verifica que sea formato `.xlsx` o `.xls`
- Intenta abrirlo primero en Excel/LibreOffice

### Columnas no detectadas correctamente
- Usa la tabla de mapeo para ajustar manualmente
- Verifica que la primera fila contenga los encabezados
- El sistema intenta detectar automáticamente la fila de encabezados

### Puerto 5000 en uso
```bash
# Ejecutar en otro puerto
python app.py --port 5001
# O editar app.py y cambiar port=5000 a port=5001
```

## Notas Técnicas

- **Backend**: Flask (Python 3.8+)
- **Procesamiento**: Pandas + OpenPyXL
- **Frontend**: HTML5, CSS3, JavaScript vanilla
- **Estilos**: Diseño moderno con CSS Grid y Flexbox
- **Responsive**: Compatible con móviles, tablets y escritorio

## Licencia

Proyecto desarrollado para estandarización de inventarios.

---

**Desarrollado para automatizar la gestión de inventarios**
