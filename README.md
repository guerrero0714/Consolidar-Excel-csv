# ⚡ HACK14 · Data Suite

Aplicación profesional de análisis y consolidación de datos con IA integrada.

---

## 🚀 Instalación rápida

```bash
# 1. Instalar dependencias
pip install -r requirements.txt

# 2. Ejecutar la aplicación
streamlit run hack14_app.py
```

La app abre automáticamente en `http://localhost:8501`

---

## 📦 Dependencias principales

| Paquete | Propósito |
|---------|-----------|
| `streamlit` | Interfaz web profesional |
| `pandas` | Procesamiento de datos |
| `openpyxl` | Lectura/escritura Excel |
| `python-docx` | Lectura de archivos Word |
| `pypdf` | Extracción de texto PDF |
| `anthropic` | IA (Claude) para análisis |
| `chardet` | Detección de encoding |

---

## 🗂️ Módulos de la aplicación

### 📊 Consolidador Excel/CSV
- Carga múltiples archivos `.xlsx`, `.xls`, `.xlsm`, `.csv`
- Valida estructura de columnas automáticamente
- Elimina duplicados y filas vacías
- Exporta el resultado en CSV y Excel
- Muestra estadísticas y diagnóstico de columnas

### 📄 Análisis de Documentos (requiere API key)
- Carga múltiples PDFs, Word (.docx) y TXT
- Extrae texto automáticamente
- Genera resúmenes con IA (Claude)
- Chat interactivo sobre el contenido de los documentos
- Sugerencias de preguntas rápidas

### ⚙️ Configuración IA
- Ingresa tu API key de Anthropic
- Verificación de la key en tiempo real
- La key solo se guarda en la sesión local

---

## 🔑 API Key de Anthropic

1. Ve a [console.anthropic.com](https://console.anthropic.com)
2. Crea una cuenta o inicia sesión
3. Genera una nueva API key
4. Pégala en ⚙️ Configuración dentro de la app

**Nota:** La app funciona sin API key para la funcionalidad de consolidación. 
El análisis con IA (resúmenes y chat) requiere la key.

---

## 💡 Uso sin IA

El consolidador de archivos Excel/CSV funciona completamente sin API key:
- Valida estructuras de datos
- Consolida múltiples archivos
- Reporte de errores detallado
- Exporta en CSV y Excel

---

*HACK14 Data Suite — Análisis de datos profesional*
