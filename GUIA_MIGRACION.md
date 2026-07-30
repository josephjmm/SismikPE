# Guía de Migración de SismikPE: de Python a C#

Basado en el análisis de las funcionalidades descritas en el proyecto (página de aterrizaje de SismikPE), el software actual automatiza procesos en ETABS, genera reportes en Excel, Word y LaTeX, y realiza cálculos complejos de ingeniería sísmica (RNE E.030). Al no contar con el código fuente en Python en este repositorio, el siguiente análisis infiere la arquitectura estándar utilizada en este tipo de soluciones y proporciona una guía detallada para migrar el sistema a **C# (.NET)**.

## 1. Análisis del Proyecto y Equivalencias de Librerías

Para lograr una funcionalidad idéntica en C#, aquí se detallan las librerías típicas de Python que sustentan un proyecto de esta naturaleza, frente a sus mejores alternativas en C#.

### Conexión y Automatización de ETABS (API)
- **Python**: Se suele utilizar `comtypes` o `pywin32` (win32com) para interactuar con la API COM de ETABS (CSi OAPI).
- **C#**:
  - **Uso directo de Interop (CSI OAPI)**: ETABS proporciona directamente la librería `CSiAPIv1.dll` (o `ETABSv1.dll`). Al estar desarrollado en C#, la integración es nativa y mucho más limpia, robusta y con tipado fuerte.

### Manejo de Datos y Cálculos Matriciales
- **Python**: `pandas` para estructuración de datos y `numpy` para cálculos (matrices de rigidez, análisis modal, escalamiento).
- **C#**:
  - **Math.NET Numerics**: La mejor alternativa para cálculos avanzados de álgebra lineal, matrices y estadística.
  - **Deedle** o **Microsoft.Data.Analysis** (DataFrame): Para manipular estructuras tabulares y limpieza de datos similares a `pandas`.
  - **LINQ**: Integrado en C#, es extremadamente potente para filtrar, agrupar y transformar colecciones de datos extraídos de ETABS.

### Exportación y Reportes en Excel
- **Python**: `pandas` (`to_excel`), `openpyxl` o `xlsxwriter`.
- **C#**:
  - **EPPlus** o **ClosedXML**: Ambas son excelentes para generar y manipular archivos `.xlsx` sin depender de tener Excel instalado, ofreciendo control total sobre el formato.

### Generación de Reportes en Word (.docx)
- **Python**: `python-docx`.
- **C#**:
  - **Open XML SDK** (de Microsoft): Muy robusto pero detallado.
  - **DocX (Xceed)**: Más fácil de usar, similar a python-docx, permite manipular texto, tablas e imágenes en Word con menos código.

### Generación de Reportes en LaTeX
- **Python**: `pylatex` o motores de plantillas como `Jinja2`.
- **C#**:
  - Generación basada en texto mediante plantillas con **Scriban** o **Fluid**, que permiten inyectar las variables del cálculo estructural dinámicamente sobre la estructura del código LaTeX.

### Gráficos y Visualización
- **Python**: `matplotlib`, `seaborn` o `plotly` (para exportar como imagen).
- **C#**:
  - **ScottPlot**: Excelente rendimiento para gráficos técnicos y científicos. Permite exportar los gráficos directamente a PNG para incluirlos en los reportes (Word/LaTeX).
  - **OxyPlot**: Otra buena alternativa para gráficos 2D.

---

## 2. Pasos para la Migración

1. **Estructurar la Solución en C# (Visual Studio / .NET 8+)**:
   - Crear un proyecto de tipo *WPF (Windows Presentation Foundation)* o *WinForms* si el software tiene interfaz de usuario de escritorio. Si es de consola/backend, un *Console App* o *Class Library*.
   - Agregar las referencias COM (CSI OAPI) al proyecto.
2. **Migrar Clases de Dominio**:
   - Traducir las clases de Python a clases fuertemente tipadas en C# (ej. `BuildingInfo`, `SeismicCase`, `DriftResult`).
3. **Migrar la Conexión a ETABS**:
   - Implementar el conector con `CSiAPIv1.cOAPI`. Aprovechar los bloques `try-catch` y la liberación de memoria con `Marshal.ReleaseComObject`.
4. **Implementar Lógica de Norma RNE E.030**:
   - Convertir las validaciones, penalidades de irregularidad y factores estáticos/dinámicos a métodos en C#.
5. **Reemplazar Scripts de Generación de Reportes**:
   - Utilizar ClosedXML para plantillas de Excel y DocX para inyectar datos en Word. Exportar los gráficos con ScottPlot y añadirlos al documento.

---

## 3. Protección y Encriptación del Código en C# (Anti-Piratería)

A diferencia de Python (que requiere empaquetadores como PyInstaller, fáciles de descompilar), C# compila a lenguaje intermedio (IL), el cual también puede ser leído por descompiladores como *dotPeek* o *dnSpy*. Para proteger el IP de SismikPE y evitar atacantes, se deben implementar las siguientes técnicas:

### A. Ofuscación de Código (Obfuscation)
Transforma el código (nombres de variables, métodos, clases) y altera el flujo de control para que el código descompilado sea incomprensible.
- **Dotfuscator**: (Tiene versión Community integrada en Visual Studio). Muy popular y robusto.
- **Obfuscar**: Opción Open Source. Esencial para renombrar métodos, clases y propiedades.
- **ConfuserEx (Forks actuales, ej. NeoConfuserEx)**: Open source avanzado que incluye empaquetado, encriptación de constantes y protección contra "dumping" de memoria.

### B. Compilación Ahead-of-Time (AOT) - *Recomendado*
Desde .NET 7 y .NET 8, puedes publicar tu aplicación usando **Native AOT**.
- Esto compila el código C# directamente a código máquina nativo (binarios de Windows .exe y .dll que no son IL).
- **Ventaja**: Hace que el uso de herramientas como dnSpy o dotPeek sea inútil. El proceso de ingeniería inversa es igual de difícil que en una aplicación en C++.
- *Nota*: Puede requerir ajustes si usas mucha reflexión (Reflection) o carga dinámica, pero es la forma más moderna de proteger C#.

### C. Encriptación de Licencias
Como SismikPE se vende como licencia única (según su página):
- **Sistema de Activación de Hardware (Hardware ID / HWID)**: Generar un identificador único en base al número de serie de la placa madre y procesador mediante `System.Management` (WMI).
- **Criptografía Asimétrica (RSA)**:
  - Generar un par de claves (Privada y Pública). Tu servidor o generador de licencias firma un archivo `.lic` con el HWID usando la **Llave Privada**.
  - El programa en C# incluye **solo la Llave Pública** incrustada y verifica si la firma del archivo `.lic` coincide con el HWID de la PC del cliente. De este modo, es imposible para el atacante generar licencias falsas válidas, ya que no tiene la llave privada.
- **Protección contra depuración (Anti-Debugging)**: En C# puedes detectar si alguien intenta usar un depurador conectándose a tu proceso usando `System.Diagnostics.Debugger.IsAttached` y cerrar la aplicación inmediatamente.

### D. Empaquetadores y Virtualizadores (Comerciales)
Para una seguridad extrema (nivel de protección comercial AAA):
- **VMProtect** o **Enigma Protector**: Toman el ejecutable final y lo envuelven (wrap). Transforman partes críticas del ejecutable en "Bytecode" de una máquina virtual única generada en ese momento. Muy difíciles de vulnerar, recomendados si el coste de la licencia es alto.

## Resumen de Acción
Para convertir de Python a C# protegiendo el código: **Diseña una arquitectura con tipado fuerte, conéctate nativamente al OAPI, usa `.NET 8 Native AOT` para compilar directamente a binario, e implementa un sistema de validación de licencias RSA asimétrico atado al Hardware ID del cliente.**
