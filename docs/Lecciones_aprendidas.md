Lecciones Aprendidas - Proyecto CONDOR
Este documento centraliza las lecciones de arquitectura y flujo de trabajo aprendidas durante el desarrollo del proyecto CONDOR. Su propósito es servir como guía para mantener la calidad, consistencia y mantenibilidad del código.

Lección 1: La Estricta Naturaleza de las Interfaces en VBA
Observación: A diferencia del resto de VBA, la implementación de interfaces (Implements) es estrictamente sensible a los detalles de la firma del procedimiento.

Regla Inquebrantable: La firma de un método en una clase implementadora debe ser una copia idéntica, carácter por carácter, de la firma en la interfaz. Esto incluye:

Nombre del Método: MiMetodo es diferente de Mimetodo.

Nombre de los Parámetros: (miParam As String) es diferente de (miParametro As String).

Capitalización de Parámetros: (email As String) es diferente de (Email As String).

Paso por Valor/Referencia: La presencia o ausencia de ByVal o ByRef debe ser idéntica.

Acción Correctiva: Ante errores de "declaración no coincide", se debe usar un prompt de sincronización forzada, tratando la interfaz como la única fuente de verdad y reescribiendo las firmas en las clases implementadoras.

Lección 2: El Principio de "Programar Contra la Interfaz" en los Tests
Observación: Los errores de "método no encontrado" en los tests ocurren cuando se declara una variable del tipo de la clase concreta en lugar de la interfaz.

Regla Inquebrantable: Dentro de cualquier módulo de pruebas (Test_*.bas), las variables que referencian a nuestros servicios deben ser declaradas del tipo de su interfaz.

Correcto: Dim authService As IAuthService

Incorrecto: Dim authService As CAuthService

Acción Correctiva: Usar periódicamente prompts de auditoría de calidad de pruebas para verificar que todos los tests cumplen con esta y otras reglas de estructura (AAA, manejo de errores, etc.).

Lección 3: Estructura de Módulos y Clases en VBA
Observación: El compilador de VBA es estricto con el orden de las declaraciones dentro de un fichero.

Regla Inquebrantable: Todas las declaraciones a nivel de módulo (Public/Private/Dim para variables, Type, Enum, Declare) deben estar agrupadas en la sección de declaraciones, en la parte superior del fichero, antes de la primera definición de Sub, Function o Property.

Acción Correctiva: Ante errores de "comentario solo puede aparecer después de End Sub..." o similares, la causa raíz suele ser una declaración fuera de lugar. Se debe mover a la parte superior del fichero.

Lección 4: El Flujo de Trabajo rebuild -> Compilación Manual -> test
Observación: El comando rebuild del CLI sincroniza los ficheros, pero no garantiza la compilación en tiempo de ejecución. Muchos errores solo se manifiestan al compilar dentro de Access.

Regla Inquebrantable: El flujo de trabajo estándar es el Ciclo de Trabajo Asistido. Ninguna prueba se ejecuta hasta que el Supervisor haya confirmado que el proyecto compila exitosamente de forma manual (Depuración -> Compilar Proyecto).

Acción Correctiva: El prompt para Copilot siempre debe finalizar con una pausa para la verificación manual del Supervisor antes de proceder con los tests o el commit.

Lección 5: Conversión Explícita de Tipos desde Arrays Variant
Observación: Al iterar sobre un array de tipo Variant (creado con Array(...)) y pasar sus elementos a una función que espera un tipo de dato específico (Long, String, etc.), VBA puede fallar al realizar la conversión de tipo implícita, resultando en un error "El tipo de argumento de ByRef no coincide", incluso si el parámetro se pasa ByVal.

Regla Inquebrantable: Para garantizar la robustez, siempre se debe realizar una conversión de tipo explícita al pasar un elemento de un array Variant a una función que espera un tipo específico.

Correcto: MiFuncion(CLng(miArrayVariant(i)))

Incorrecto: MiFuncion(miArrayVariant(i))

Acción Correctiva: Ante este error, se debe añadir la función de conversión apropiada (CLng, CStr, CInt, CBool, etc.) en la llamada al procedimiento.

Lección 6: Usar los Tests y Módulos de Acceso a Datos como Especificación para Clases de Datos
Observación: Errores de "método o dato miembro no encontrado" ocurren frecuentemente en los tests y módulos de acceso a datos al usar clases de tipo de datos (T_*.cls) que están incompletas.

Regla Inquebrantable: Los tests que construyen objetos de datos (ej. en un bloque With...End With) Y los módulos de acceso a datos (como modDatabase.bas) que asignan valores a propiedades de objetos actúan como la especificación funcional para esas clases de datos. La clase debe contener todas las propiedades públicas que tanto los tests como los módulos de datos utilizan.

Fuentes de Verdad para Auditoría:

Módulos de Prueba (Test_*.bas): Revelan propiedades utilizadas en construcción y validación de objetos.

Módulos de Acceso a Datos (modDatabase.bas, *Repository.cls): Revelan propiedades utilizadas en persistencia y recuperación de datos.

Servicios Mock (CMock*.cls): Revelan propiedades utilizadas en simulación de datos.

Especificaciones de Integración: Revelan propiedades requeridas para intercambio de datos.

Acción Correctiva: Ante este error, se debe realizar una auditoría proactiva completa:

Auditar todos los tests que usan la clase de datos.

Auditar todos los módulos de acceso a datos que manipulan la clase.

Auditar servicios y mocks que utilizan la clase.

Añadir todas las propiedades faltantes a la clase de tipo de datos correspondiente (T_*.cls).

Extender la auditoría a todas las demás clases T_*.cls para prevenir errores similares.

Lección 7: La Batería de Pruebas Debe Ser Exhaustiva
Observación: Se han creado nuevos módulos de prueba (Test_*.bas) con sus respectivas suites de ejecución (Test_*_RunAll), pero se ha omitido añadirlos a la función principal de ejecución de pruebas en modTestRunner.bas. Esto provoca que los nuevos tests, aunque existan, nunca se ejecuten como parte de la batería completa, creando un falso sentido de seguridad.

Regla Inquebrantable: El módulo modTestRunner.bas es el único punto de verdad sobre la cobertura total de las pruebas. Cada vez que se cree un nuevo módulo de pruebas (Test_CSolicitudPC.bas, por ejemplo), su función de ejecución principal (Test_CSolicitudPC_RunAll) debe ser registrada inmediatamente dentro de la función RunAllTests en modTestRunner.bas.

Acción Correctiva: Todos los prompts que impliquen la creación de un nuevo módulo de pruebas deben incluir explícitamente un paso final para modificar modTestRunner.bas y añadir la llamada a la nueva suite de pruebas.
---
### Lección 8: La Centralización del Manejo de Errores es Obligatoria

**Observación:** Se ha detectado código que, si bien utiliza `On Error GoTo ErrorHandler`, no registra el error capturado en nuestro servicio central `modErrorHandler`. Esto crea "agujeros negros" en la traza de errores, haciendo la depuración casi imposible.

**Regla Inquebrantable:** Un bloque `ErrorHandler` sin una llamada a `modErrorHandler.LogError` (o `LogCriticalError`) se considera una implementación **incompleta y errónea**. El propósito de capturar un error es registrarlo de forma centralizada.

* **Incorrecto (Incompleto):**
    ```vba
    ErrorHandler:
        ' No hace nada o solo muestra un MsgBox
    ```
* **Correcto (Completo):**
    ```vba
    ErrorHandler:
        Call modErrorHandler.LogError(Err.Number, Err.Description, "NombreModulo.NombreFuncion")
    ```

**Acción Correctiva:** Todos los prompts que impliquen la creación o modificación de código deben incluir la directiva de implementar el manejo de errores centralizado. Además, se deben realizar auditorías proactivas para buscar bloques `ErrorHandler` que no cumplan con esta regla.

---
### Lección 9: La Auditoría de Operaciones es un Requisito, no una Opción

**Observación:** Mientras que el `modErrorHandler` nos dice *qué ha fallado*, el `IOperationLogger` nos dice *qué ha ocurrido*. La trazabilidad de las acciones de negocio (quién hizo qué y cuándo) es tan crítica como el registro de errores para la seguridad y el mantenimiento del sistema.

**Regla Inquebrantable:** Toda función o procedimiento que represente una acción de negocio significativa debe registrar dicha acción a través del servicio `IOperationLogger`. Esto es obligatorio para:
* Cualquier operación que cree, modifique o elimine datos (ej. `Save`, `Delete`).
* Cualquier cambio de estado en el workflow (ej. `ChangeState`).
* Acciones críticas como la autenticación de un usuario o la generación de un documento.

* **Incorrecto (Sin Trazabilidad):**
    ```vba
    Public Function Save(solicitud As ISolicitud) As Boolean
        ' ...código para guardar en la base de datos...
        Save = True
    End Function
    ```
* **Correcto (Con Trazabilidad):**
    ```vba
    Public Function Save(solicitud As ISolicitud) As Boolean
        ' ...código para guardar en la base de datos...
        If guardadoExitoso Then
            m_OperationLogger.LogOperation "Guardar Solicitud", solicitud.idSolicitud, "Datos guardados."
        End If
        Save = guardadoExitoso
    End Function
    ```

**Acción Correctiva:** Todos los prompts para la creación de nuevas funcionalidades deben incluir explícitamente el requisito de implementar el logging de auditoría en los puntos clave del negocio.

---
### Lección 10: El Aislamiento de las Pruebas Unitarias con Mocks no es Negociable

**Observación:** Se han detectado pruebas unitarias que, en lugar de aislar el componente bajo prueba, interactúan con dependencias reales (conexiones a bases de datos, acceso al sistema de ficheros). Esto hace que las pruebas sean lentas, frágiles y dependientes del entorno, convirtiéndolas en tests de integración en lugar de tests unitarios.

**Regla Inquebrantable:** Una prueba unitaria debe probar una única "unidad" de código de forma aislada. Todas las dependencias externas de esa unidad deben ser reemplazadas por **Mocks**. El objetivo es controlar el entorno de la prueba para validar la lógica interna del componente, no la de sus dependencias.

* **Incorrecto (Prueba de Integración):**
    ```vba
    ' La prueba depende de que CConfig se conecte a la BD real
    Dim authService As New CAuthService
    Dim config As New CConfig 
    authService.Initialize config, logger ' Se inyecta una dependencia real
    ```
* **Correcto (Prueba Unitaria Aislada):**
    ```vba
    ' La prueba controla el entorno con un Mock
    Dim authService As New CAuthService
    Dim mockConfig As New CMockConfig ' Se crea un Mock
    mockConfig.SetValue "CLAVE_NECESARIA", "VALOR_SIMULADO" ' Se configura el Mock
    authService.Initialize mockConfig, mockLogger ' Se inyecta el Mock
    ```

**Acción Correctiva:** Todos los prompts para la creación de pruebas unitarias deben incluir explícitamente el requisito de usar Mocks para todas las dependencias externas. Se deben realizar auditorías para identificar y refactorizar las pruebas que no cumplan con este principio de aislamiento.

---
### Lección 11: La Plantilla Estándar para Clases de Servicio

**Observación:** Para asegurar la consistencia y la aplicación de todos los principios aprendidos, toda nueva clase de servicio (ej. `CWorkflowService`, `CReportingService`, etc.) debe seguir una estructura estándar desde su creación.

**Regla Inquebrantable:** Toda nueva clase de servicio debe nacer con la siguiente estructura mínima:

1.  **Declaración de Interfaz:** `Implements INombreDelServicio`
2.  **Dependencias Privadas:** Variables privadas para cada dependencia que necesite (ej. `Private m_Config As IConfig`).
3.  **Método `Initialize` Público:** Un `Public Sub Initialize(...)` para la inyección de todas sus dependencias.
4.  **Implementación de la Interfaz:** Métodos `Private Sub/Function INombreDelServicio_...` que cumplan el contrato.
5.  **Manejo de Errores Centralizado:** Cada método debe tener un bloque `On Error GoTo` que llame a `modErrorHandler.LogError`.
6.  **Logging de Auditoría:** Los métodos que representen acciones de negocio deben llamar a `m_OperationLogger.LogOperation`.

**Plantilla de Inicio (Ejemplo para `CWorkflowService`):**
```vba
Option Compare Database
Option Explicit

Implements IWorkflowService

' 1. Dependencias
Private m_Config As IConfig
Private m_OperationLogger As IOperationLogger
Private m_SolicitudRepository As ISolicitudRepository

' 2. Método Initialize
Public Sub Initialize(config As IConfig, logger As IOperationLogger, repo As ISolicitudRepository)
    On Error GoTo ErrorHandler
    Set m_Config = config
    Set m_OperationLogger = logger
    Set m_SolicitudRepository = repo
    Exit Sub
ErrorHandler:
    Call modErrorHandler.LogError(Err.Number, Err.Description, "CWorkflowService.Initialize")
End Sub

' 3. Implementación de la Interfaz
Private Function IWorkflowService_ChangeState(...) As Boolean
    On Error GoTo ErrorHandler
    ' ... Lógica de negocio ...
    
    ' 4. Logging de Auditoría
    m_OperationLogger.LogOperation "ChangeState", idSolicitud, "Estado cambiado a " & newState
    
    Exit Function
ErrorHandler:
    ' 5. Manejo de Errores
    Call modErrorHandler.LogError(Err.Number, Err.Description, "CWorkflowService.IWorkflowService_ChangeState")
End Function