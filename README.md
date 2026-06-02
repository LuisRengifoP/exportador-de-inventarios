# exportador-de-inventarios
Flask + Pandas + OpenPyXL. Sistema de transformación de datos con mapeo inteligente de columnas basado en similitud de secuencias y diccionario de aliases. Soporta archivos .xls y .xlsx con detección automática de hojas y encabezados.
| Grupo              | Columnas                                                | Propósito                 |
| ------------------ | ------------------------------------------------------- | ------------------------- |
| **Identificación** | `sku`, `url_key`                                        | Código único del producto |
| **Configuración**  | `locale`, `type`, `attribute_family_code`, `categories` | Parámetros del sistema    |
| **Contenido**      | `name`, `description`, `short_description`              | Información del producto  |
| **Estado**         | `status`, `visible_individually`                        | Visibilidad y activación  |
| **Marca**          | `brand`                                                 | Fabricante/proveedor      |
| **Compra**         | `guest_checkout`                                        | Configuración de checkout |
| **Precios**        | `price`, `Iva`                                          | Valores económicos        |
