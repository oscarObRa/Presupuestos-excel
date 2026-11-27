# ADN_ICEL: Identidad Visual Corporativa

Este documento constituye la "fuente de verdad" para la generación de documentos (Excel/PDF) de **ICEL Contratistas S.A. de C.V.** 

Cualquier script de generación debe consultar este ADN para asegurar la consistencia de marca, evitando el uso de temas genéricos de Office.

---

## 1. Especificaciones de Color (Paleta Institucional)

Los colores deben aplicarse mediante sus valores HEX (sin el prefijo `#`) para asegurar la compatibilidad con la librería `openpyxl`.

| Elemento | Variable Sugerida | Valor HEX | Uso Principal |
| :--- | :--- | :--- | :--- |
| **Encabezado de Tabla** | `FILL_TABLE_HEADER` | `3C7D22` | Relleno de la fila de títulos de columna. Texto: Blanco. |
| **Acento de Sección** | `COLOR_SECTION_ACCENT` | `4EA72E` | Títulos de sección (I, II, III...) y líneas divisorias. |
| **Bordes de Celda** | `COLOR_BORDER_DARK` | `275317` | Bordes exteriores (medium) e interiores (thin). |
| **Fondo de Partida** | `FILL_PART_ROW` | `F2F5F0` | Resalte de filas que agrupan conceptos. |
| **Texto Principal** | `FONT_TEXT_BLACK` | `000000` | Cuerpo de texto y datos numéricos. |
| **Texto en Header** | `FONT_TEXT_WHITE` | `FFFFFF` | Texto sobre fondos oscuros. |

---

## 2. Maquetación y Retícula (Layout)

El diseño se basa en una cuadrícula de **17 columnas (A a Q)**. Esta estructura permite una maquetación tipo "revista" donde los elementos pueden alinearse de forma asimétrica y profesional.

### Anchos de Columna (`COLUMN_WIDTHS`)
| Col | Ancho | Col | Ancho | Col | Ancho | Col | Ancho |
| :-- | :--- | :-- | :--- | :-- | :--- | :-- | :--- |
| **A** | 3.8 | **E** | 4.2 | **I** | 12.2 | **M** | 11.8 |
| **B** | 13.0 | **F** | 1.5 | **J** | 4.5 | **N** | 18.0 |
| **C** | 4.6 | **G** | 4.0 | **K** | 72.8 | **O** | 7.5 |
| **D** | 6.4 | **H** | 3.5 | **L** | 11.5 | **P** | 7.2 |
| | | | | | | **Q** | 4.6 |

### Mapa de Coordenadas de la Carátula (`CELL_MAP`)
| Variable | Celda | Descripción |
| :--- | :--- | :--- |
| `lugar_fecha` | **A7** | Ciudad y fecha de la propuesta. |
| `folio` | **L7** | Código de identificación del documento (alineado derecha). |
| `etiqueta_propuesta` | **A10** | Texto fijo "Propuesta:". |
| `nombre_propuesta` | **D10** | Título del proyecto/propuesta. |
| `etiqueta_atencion` | **A13** | Texto fijo "Att'n:". |
| `destinatario` | **C13** | Nombre de la persona contacto. |
| `cargo` | **C14** | Cargo del destinatario. |
| `empresa` | **C15** | Empresa del cliente. |
| `introduccion` | **C17** | Párrafo descriptivo del alcance. |
| `seccion_economica` | **A20** | Título de la sección de costos. |

---

## 3. Tipografía y Estilo de Impresión

* **Fuente Base:** `Arial` (Estandarizada para evitar errores de conversión en PDF).
* **Jerarquías:**
    * **Empresa/Título:** 13 pt, Bold.
    * **Secciones:** 12 pt, Bold.
    * **Datos/Tablas:** 11 pt, Bold.
    * **Descripciones/Cuerpo:** 10 pt, Regular.
    * **Notas/Cláusulas:** 9 pt, Regular/Italic.
* **Configuración de Página:** 
    * Tamaño: Carta (Letter).
    * Orientación: Vertical (Portrait).
    * Ajuste: Escalar a 1 página de ancho (`fitToWidth = 1`).
    * Cuadrícula: Oculta (`showGridLines = False`).

---

## 4. Activos (Assets)

* **Logotipo:** `assets/Logo.png`
* **Pie de Página:** `assets/Pie de pagina.png`
