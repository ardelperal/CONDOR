INICIO DEL PROMPT PARA COPILOT
Hola. Basándonos en los recientes desafíos de compilación, vamos a crear un nuevo documento para registrar las lecciones de arquitectura aprendidas y asegurar la consistencia del proyecto a largo plazo.

Misión: Crear un nuevo fichero de documentación llamado Lecciones_aprendidas.md en la raíz del proyecto.

Inicia el CICLO DE TRABAJO ASISTIDO y sigue estos pasos:

Fase 1: Crear y Poblar el Fichero de Lecciones Aprendidas

Crear el Fichero:

En el directorio raíz del proyecto (C:\Proyectos\CONDOR\), crea un nuevo fichero llamado Lecciones_aprendidas.md.

Añadir Contenido:

Pega el siguiente contenido Markdown en el nuevo fichero. Este texto resume las reglas y principios clave que hemos descubierto y formalizado.

# Lecciones Aprendidas - Proyecto CONDOR

Este documento centraliza las lecciones de arquitectura y flujo de trabajo aprendidas durante el desarrollo del proyecto CONDOR. Su propósito es servir como guía para mantener la calidad, consistencia y mantenibilidad del código.

---
### Lección 1: La Estricta Naturaleza de las Interfaces en VBA

**Observación:** A diferencia del resto de VBA, la implementación de interfaces (`Implements`) es **estrictamente sensible** a los detalles de la firma del procedimiento.

**Regla Inquebrantable:** La firma de un método en una clase implementadora debe ser una copia **idéntica, carácter por carácter**, de la firma en la interfaz. Esto incluye:
1.  **Nombre del Método:** `MiMetodo` es diferente de `Mimetodo`.
2.  **Nombre de los Parámetros:** `(miParam As String)` es diferente de `(miParametro As String)`.
3.  **Capitalización de Parámetros:** `(email As String)` es diferente de `(Email As String)`.
4.  **Paso por Valor/Referencia:** La presencia o ausencia de `ByVal` o `ByRef` debe ser idéntica.

**Acción Correctiva:** Ante errores de "declaración no coincide", se debe usar un prompt de **sincronización forzada**, tratando la interfaz como la única fuente de verdad y reescribiendo las firmas en las clases implementadoras.

---

### Lección 2: El Principio de "Programar Contra la Interfaz" en los Tests

**Observación:** Los errores de "método no encontrado" en los tests ocurren cuando se declara una variable del tipo de la clase concreta en lugar de la interfaz.

**Regla Inquebrantable:** Dentro de cualquier módulo de pruebas (`Test_*.bas`), las variables que referencian a nuestros servicios deben ser declaradas del tipo de su **interfaz**.

* **Correcto:** `Dim authService As IAuthService`
* **Incorrecto:** `Dim authService As CAuthService`

**Acción Correctiva:** Usar periódicamente prompts de **auditoría de calidad de pruebas** para verificar que todos los tests cumplen con esta y otras reglas de estructura (AAA, manejo de errores, etc.).

---

### Lección 3: Estructura de Módulos y Clases en VBA

**Observación:** El compilador de VBA es estricto con el orden de las declaraciones dentro de un fichero.

**Regla Inquebrantable:** Todas las declaraciones a nivel de módulo (`Public`/`Private`/`Dim` para variables, `Type`, `Enum`, `Declare`) deben estar agrupadas en la **sección de declaraciones**, en la parte superior del fichero, antes de la primera definición de `Sub`, `Function` o `Property`.

**Acción Correctiva:** Ante errores de "comentario solo puede aparecer después de End Sub..." o similares, la causa raíz suele ser una declaración fuera de lugar. Se debe mover a la parte superior del fichero.

---

### Lección 4: El Flujo de Trabajo `rebuild -> Compilación Manual -> test`

**Observación:** El comando `rebuild` del CLI sincroniza los ficheros, pero no garantiza la compilación en tiempo de ejecución. Muchos errores solo se manifiestan al compilar dentro de Access.

**Regla Inquebrantable:** El flujo de trabajo estándar es el **Ciclo de Trabajo Asistido**. Ninguna prueba se ejecuta hasta que el Supervisor haya confirmado que el proyecto compila exitosamente de forma manual (`Depuración -> Compilar Proyecto`).

**Acción Correctiva:** El prompt para Copilot siempre debe finalizar con una pausa para la verificación manual del Supervisor antes de proceder con los tests o el commit.

### Lección 5: La Estricta Naturaleza de las Interfaces en VBA

**Observación:** A diferencia del resto de VBA, la implementación de interfaces (`Implements`) es **estrictamente sensible** a los detalles de la firma del procedimiento.

**Regla Inquebrantable:** La firma de un método en una clase implementadora debe ser una copia **idéntica, carácter por carácter**, de la firma en la interfaz. Esto incluye:
1.  **Nombre del Método:** `MiMetodo` es diferente de `Mimetodo`.
2.  **Nombre de los Parámetros:** `(miParam As String)` es diferente de `(miParametro As String)`.
3.  **Capitalización de Parámetros:** `(email As String)` es diferente de `(Email As String)`.
4.  **Paso por Valor/Referencia:** La presencia o ausencia de `ByVal` o `ByRef` debe ser idéntica.

**Acción Correctiva:** Ante errores de "declaración no coincide", se debe usar un prompt de **sincronización forzada**, tratando la interfaz como la única fuente de verdad y reescribiendo las firmas en las clases implementadoras.

---

### Lección 6: El Principio de "Programar Contra la Interfaz" en los Tests

**Observación:** Los errores de "método no encontrado" en los tests ocurren cuando se declara una variable del tipo de la clase concreta en lugar de la interfaz.

**Regla Inquebrantable:** Dentro de cualquier módulo de pruebas (`Test_*.bas`), las variables que referencian a nuestros servicios deben ser declaradas del tipo de su **interfaz**.

* **Correcto:** `Dim authService As IAuthService`
* **Incorrecto:** `Dim authService As CAuthService`

**Acción Correctiva:** Usar periódicamente prompts de **auditoría de calidad de pruebas** para verificar que todos los tests cumplen con esta y otras reglas de estructura (AAA, manejo de errores, etc.).

---

### Lección 7: El Flujo de Trabajo `rebuild -> Compilación Manual -> test`

**Observación:** El comando `rebuild` del CLI sincroniza los ficheros, pero no garantiza la compilación en tiempo de ejecución. Muchos errores solo se manifiestan al compilar dentro de Access.

**Regla Inquebrantable:** El flujo de trabajo estándar es el **Ciclo de Trabajo Asistido**. Ninguna prueba se ejecuta hasta que el Supervisor haya confirmado que el proyecto compila exitosamente de forma manual (`Depuración -> Compilar Proyecto`).

**Acción Correctiva:** El prompt para Copilot siempre debe finalizar con una pausa para la verificación manual del Supervisor antes de proceder con los tests o el commit.

Fase 2: Verificación y Pausa

Ejecutar rebuild:

No es necesario un rebuild ya que solo hemos añadido un fichero de documentación.

Prepara el commit con el nuevo fichero.

Pausa y espera la confirmación del Supervisor.

Por favor, procede.
FIN DEL PROMPT PARA COPILOT

---

### Lección 5: Conversión Explícita de Tipos desde Arrays Variant

**Observación:** Al iterar sobre un array de tipo `Variant` (creado con `Array(...)`) y pasar sus elementos a una función que espera un tipo de dato específico (`Long`, `String`, etc.), VBA puede fallar al realizar la conversión de tipo implícita, resultando en un error "El tipo de argumento de ByRef no coincide", incluso si el parámetro se pasa `ByVal`.

**Regla Inquebrantable:** Para garantizar la robustez, siempre se debe realizar una **conversión de tipo explícita** al pasar un elemento de un array `Variant` a una función que espera un tipo específico.

* **Correcto:** `MiFuncion(CLng(miArrayVariant(i)))`
* **Incorrecto:** `MiFuncion(miArrayVariant(i))`

**Acción Correctiva:** Ante este error, se debe añadir la función de conversión apropiada (`CLng`, `CStr`, `CInt`, `CBool`, etc.) en la llamada al procedimiento.

---

### Lección 6: Usar los Tests y Módulos de Acceso a Datos como Especificación para Clases de Datos

**Observación:** Errores de "método o dato miembro no encontrado" ocurren frecuentemente en los tests y módulos de acceso a datos al usar clases de tipo de datos (`T_*.cls`) que están incompletas.

**Regla Inquebrantable:** Los tests que construyen objetos de datos (ej. en un bloque `With...End With`) Y los módulos de acceso a datos (como `modDatabase.bas`) que asignan valores a propiedades de objetos actúan como la especificación funcional para esas clases de datos. La clase debe contener todas las propiedades públicas que tanto los tests como los módulos de datos utilizan.

**Fuentes de Verdad para Auditoría:**
1. **Módulos de Prueba (`Test_*.bas`):** Revelan propiedades utilizadas en construcción y validación de objetos
2. **Módulos de Acceso a Datos (`modDatabase.bas`, `*Repository.cls`):** Revelan propiedades utilizadas en persistencia y recuperación de datos
3. **Servicios Mock (`CMock*.cls`):** Revelan propiedades utilizadas en simulación de datos
4. **Especificaciones de Integración:** Revelan propiedades requeridas para intercambio de datos

**Acción Correctiva:** Ante este error, se debe realizar una **auditoría proactiva completa**:
1. Auditar todos los tests que usan la clase de datos
2. Auditar todos los módulos de acceso a datos que manipulan la clase
3. Auditar servicios y mocks que utilizan la clase
4. Añadir todas las propiedades faltantes a la clase de tipo de datos correspondiente (`T_*.cls`)
5. Extender la auditoría a todas las demás clases `T_*.cls` para prevenir errores similares