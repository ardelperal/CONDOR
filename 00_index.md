# Índice de la Base de Conocimiento del Proyecto CONDOR

Este documento sirve como índice y guía para el agente de IA **CONDOR-Architect**. Describe el propósito de cada uno de los ficheros que componen su base de conocimiento.

---

### **1. Documentos de Especificación y Metodología (Las "Leyes")**

Estos son los documentos más importantes y deben ser tratados como la fuente de verdad absoluta.

*   **`Especificacion_Funcional.md`**
    *   **Propósito:** Describe **QUÉ** debe hacer la aplicación CONDOR. Contiene los requisitos del negocio, la definición de las tablas de la base de datos, los flujos de trabajo y los mapeos de campos.
    *   **Cuándo Usarlo:** Para responder cualquier pregunta sobre la funcionalidad, los datos que se deben manejar o las reglas de negocio. Es la "voz del cliente".

*   **`PLAN_DE_ACCION.md`**
    *   **Propósito:** Describe **CÓMO** se debe construir la aplicación. Es el documento de arquitectura y metodología principal.
    *   **Secciones Críticas:** Contiene los **"Principios de Arquitectura de Código"** (3 capas, Interfaces, Mocks, CamelCase) y el **"CICLO DE TRABAJO DE DESARROLLO (MODO TDD AUTÓNOMO)"**. Estas son las reglas inquebrantables que deben guiar toda la generación de código.
    *   **Cuándo Usarlo:** Para guiar cualquier tarea de desarrollo, refactorización o depuración. Define tu comportamiento y proceso.

*   **`README.md`**
    *   **Propósito:** Es la guía práctica para el desarrollador. Explica cómo usar las herramientas de automatización.
    *   **Contenido Clave:** Contiene la descripción y el uso de la herramienta `condor_cli.vbs`, el flujo de trabajo con Git, y los procedimientos para la verificación manual de pruebas.
    *   **Cuándo Usarlo:** Para recordar los comandos de la CLI o los procedimientos de trabajo.

### **2. Ejemplos de Código (Los "Artefactos")**

Estos ficheros sirven como ejemplos concretos de la aplicación de las reglas definidas en los documentos anteriores.

*   **`condor_cli.vbs.txt`**
    *   **Propósito:** Muestra un ejemplo completo de un script VBScript robusto para la automatización, con manejo de errores, acceso a ficheros y automatización COM de Access.

*   **`codigo_fuente_vba.zip`**
    *   **Propósito:** Contiene una instantánea completa del código fuente VBA del proyecto. Es el ejemplo práctico de nuestros **Principios de Arquitectura**.
    *   **Contenido:** Dentro encontrarás ejemplos de Módulos (`mod...`), Clases (`C...`), Interfaces (`I...`), Mocks (`CMock...`) y Módulos de Pruebas (`Test_...`).
    *   **Cuándo Usarlo:** Para ver cómo se aplican en la práctica las reglas de nombrado, el uso de `Implements`, la estructura de las pruebas, etc.