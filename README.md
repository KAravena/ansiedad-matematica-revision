# Ansiedad Matemática en Educación Escolar: Revisión de la Literatura Empírica Reciente

[Presentación Quarto RevealJS](https://karavena.github.io/ansiedad-matematica-revision/Presentacion_Revealjs/presentacion_ansiedad_matematica.html#/)

Este repositorio contiene los materiales, scripts y la presentación finalcorrespondientes a la investigación: **"Ansiedad matemática en estudiantes de educación escolar: Definiciones, medición y dimensiones educativas asociadas en la literatura empírica reciente"**.

---

##  Resumen del Proyecto

La investigación tiene como propósito principal caracterizar cómo la literatura científica empírica reciente (2021-2026) define, mide y analiza la ansiedad matemática en el contexto de la educación escolar (primaria y secundaria), identificando además las dimensiones educativas con las que se asocia este fenómeno.

### Objetivos Específicos
1. **Definiciones y operacionalización:** Identificar las formas en que se define y operacionaliza (mide) la ansiedad matemática.
2. **Instrumentos y métodos:** Describir los diseños, metodologías e instrumentos empleados.
3. **Dimensiones educativas:** Sintetizar variables de aprendizaje, rendimiento escolar y contexto asociadas.

---

##  Metodología y Flujo de Trabajo

La revisión se construyó mediante un proceso riguroso utilizando bases de datos internacionales, soportado por automatización en Python para el procesamiento de los registros bibliográficos.

### 1. Bases de Datos Consultadas
- **ERIC:** Enfoque en investigación educativa.
- **Scopus:** Cobertura multidisciplinaria e internacional.
- **Web of Science (WoS):** Literatura indexada de alto impacto.

### 2. Ecuación de Búsqueda
Se utilizaron bloques booleanos combinando términos de **ansiedad matemática** (math anxiety, mathematics anxiety), **población escolar** (student, school, child), y **rendimiento/aprendizaje** (math achievement, performance); excluyendo literatura de educación superior o formación docente inicial.

### 3. Procesamiento y Selección (Scripts en Python)
Se desarrollaron scripts personalizados para asistir en la revisión sistemática:
1. prisma_pipeline_ansiedad_matematica.py

2. cribado_trazable_ansiedad_matematica.py
---

##  Principales Hallazgos (Síntesis)

Tras filtrar 1.369 registros iniciales, se caracterizó un corpus final de artículos empíricos de diversos contextos (Europa, Asia, África). Entre los resultados destacan:
- **Constructo Medible:** La ansiedad matemática es abordada predominantemente mediante modelos psicométricos estandarizados (SEM, CFA, modelos multinivel).
- **Mecanismo Intermedio:** La ansiedad matemática opera como una variable mediadora/moduladora entre el contexto (apoyo familiar, escolar) y la cognición (memoria de trabajo, atención, rendimiento). No actúa como un factor aislado, sino como parte del entramado de la experiencia educativa.

---

## 📁 Estructura del Repositorio

- `/Presentacion_Revealjs/`: Contiene el código fuente `.qmd` de la presentación académica, las hojas de estilo (`.scss`), imágenes, y bibliografía (`references.bib`).
- `/Procesamiento/`: Directorio para los scripts de Python y los datasets en sus fases raw y procesadas.

---

**Desarrollado para la presentación del Seminario de Investigación (2025).**
