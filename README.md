# 📄 Generación Automática de Evidencias Word

Herramienta de escritorio desarrollada en Python que automatiza la creación de documentos Word de evidencias de pruebas a partir de una matriz de casos de prueba en formato Excel.

Diseñada para equipos de QA que buscan reducir el tiempo invertido en documentación manual y enfocarse en lo que realmente importa: **ejecutar pruebas**.

---

## ⚙️ Requisitos previos

- Python 3.x instalado en tu máquina
- Tener `pip` disponible en tu entorno

---

## 🚀 Instalación

**1. Clona el repositorio:**
```bash
git clone https://github.com/tu-usuario/tu-repositorio.git
cd tu-repositorio
```

**2. Crea y activa un entorno virtual:**
```bash
# 1. Crear entorno virtual
python -m venv env

# 2. Ir a la ruta del entorno virtual en Windows
cd env/Scripts/

# 3. Activar el entorno virtual
activate
```

**3. Instala las dependencias:**

```bash
# Regresar a la carpeta env
cd.. 

# Regresar a la carpeta qa-evidence-generator
cd..
```

```bash
# Estando en la ruta del principal del proyecto realizamos el siguiente comando para instalar las dependencias
pip install -r requirements.txt
```

---

## 📁 Estructura del proyecto

```
📂 tu-repositorio/
├── main.py                     # Backend: lógica de lectura del Excel y generación de Word
├── gui.py                      # Interfaz gráfica del programa
├── requirements.txt            # Librerías necesarias para ejecutar el proyecto
└── Matrix_TestCases_V1.xlsx    # Plantilla de referencia de la matriz de casos de prueba
```

---

## ▶️ Cómo ejecutar el programa

```bash
python gui.py
```

**Flujo de uso:**

1. Haz click en **Seleccionar Archivo** y elige tu matriz de casos de prueba `.xlsx`
2. Haz click en **Seleccionar Carpeta** y elige dónde quieres guardar los documentos Word
3. Haz click en **Generar** y espera a que la barra de progreso termine
4. ¡Listo! Tus evidencias Word estarán en la carpeta que seleccionaste

---

## 📊 Formato esperado del Excel

El archivo Excel debe contener una hoja llamada `TestCases` con los siguientes encabezados **en este orden exacto**:

| Columna | Descripción |
|---|---|
| Funcionalidad | Módulo o funcionalidad que se está probando |
| Id Caso de Prueba | Identificador único del caso (ej. MTC-FT-001) |
| Escenario de Prueba | Nombre descriptivo del escenario |
| Descripcion | Descripción del objetivo del caso de prueba |
| Pre-Requisitos | Condiciones necesarias antes de ejecutar |
| Datos de Prueba | Datos de entrada utilizados |
| No Paso | Número del paso |
| Descripcion del Paso | Detalle de la acción a ejecutar |
| Resultado Esperado | Resultado esperado por cada paso |
| Criticidad | Nivel de criticidad del caso |

> 📌 Usa el archivo `Matrix_TestCases_V1.xlsx` como plantilla de referencia.

---

## 📝 Resultado generado

Por cada `Id Caso de Prueba` encontrado en el Excel, el programa genera un archivo Word independiente con el siguiente formato:

```
📂 carpeta-destino/
├── MTC-FT-001.docx
├── MTC-FT-002.docx
└── MTC-FT-003.docx
```

Cada documento incluye todos los campos del formato de evidencia listos para ser completados por el Tester tras la ejecución.

---

## 👥 Equipo

Desarrollado para cualquier equipo de **QA**.

Made by: Claude Desktop + IvanBYA (Prompts and Validations)
