# Plan de Acción para Desarrollo con GitHub Copilot: Aplicación CONDOR

## 1. Introducción
Este documento sirve como guía estratégica para desarrollar la aplicación CONDOR utilizando GitHub Copilot. Se basa en la "Especificación Funcional para la Aplicación CONDOR (Versión Completa y Definitiva)". Los prompts están diseñados para maximizar la eficiencia, generar código idiomático y asegurar la coherencia con la arquitectura definida.

**Estrategia General:**
1.  **Contexto es Clave:** Antes de cada prompt complejo, proporciona a Copilot el contexto relevante (definiciones de tablas, interfaces, etc.) para mejorar la calidad de la sugerencia.
2.  **Iteración y Refinamiento:** Usa los prompts como punto de partida. Revisa y refina el código generado por Copilot para ajustarlo a las necesidades exactas y a las convenciones del proyecto.
3.  **De la Base a la Cima:** Sigue el orden de este plan, construyendo la base de datos, luego la lógica de negocio y finalmente la interfaz de usuario.

---

## 2. Fase 1: Creación de la Base de Datos (SQL DDL)
El objetivo es generar los scripts SQL para crear todas las tablas definidas en la especificación.

### Prompts para Copilot:
*   "Genera una sentencia SQL DDL para Microsoft Access para crear la tabla `Tb_Solicitudes` basándose en la sección 8.1 de la especificación."
*   "Genera una sentencia SQL DDL para Microsoft Access para crear la tabla `Tb_Datos_PC` basándose en la sección 8.2. Usa los tipos de datos apropiados para Texto, Memo, Sí/No y Fecha/Hora."
*   "Genera una sentencia SQL DDL para Microsoft Access para crear la tabla `Tb_Datos_CD_CA` basándose en la sección 8.3."
*   "Genera una sentencia SQL DDL para Microsoft Access para crear la tabla `Tb_Datos_CD_CA_SUB` basándose en la sección 8.4."
*   "Genera una sentencia SQL DDL para Microsoft Access para crear la tabla `Tb_Mapeo_Campos` basándose en la sección 8.5."
*   "Genera una sentencia SQL DDL para Microsoft Access para crear la tabla `Tb_Log_Cambios` basándose en la sección 8.6."
*   "Genera una sentencia SQL DDL para Microsoft Access para crear la tabla `Tb_Log_Errores` basándose en la sección 8.7."
*   "Genera una sentencia SQL DDL para Microsoft Access para crear la tabla `Tb_Adjuntos` basándose en la sección 8.8."

---

## 3. Fase 2: Arquitectura y Lógica de Negocio (VBA)
Aquí se crea el esqueleto de la aplicación: interfaces, clases de negocio y módulos de servicio.

### 3.1. Interfaces (Contratos)
*   **Prompt:** "Crea un Módulo de Clase VBA llamado `ISolicitud`. Esta clase debe definir la interfaz pública para gestionar solicitudes. Incluye los siguientes procedimientos: `Load(solicitudID As Long)`, `Save() As Boolean`, `Delete()`, `ChangeState(newState As String)`, `GenerateDocument()` y `Validate() As Boolean`."

### 3.2. Módulos de Servicio
*   **`modDatabase`:**
    *   **Prompt:** "Crea un módulo estándar de VBA `modDatabase`. Debe contener una función `GetRecordset(sql As String) As ADODB.Recordset` y un procedimiento `ExecuteSQL(sql As String)`. Implementa un manejo de errores robusto y gestión de la conexión para un backend de Access."
*   **`modFactory`:**
    *   **Prompt:** "Crea un módulo estándar de VBA `modFactory`. Escribe una función pública `CreateSolicitud(solicitudID As Long) As ISolicitud`. Esta función debe leer el `TipoSolicitud` de `Tb_Solicitudes` para el ID dado y devolver una instancia de la clase correspondiente (`CSolicitudPC`, `CSolicitudCDCA`, etc.)."
*   **`modWordManager`:**
    *   **Prompt:** "Crea un módulo estándar de VBA `modWordManager`. Escribe una función `FillWordTemplate(solicitudID As Long)`. Debería:
        1. Obtener el `TipoSolicitud` para determinar el nombre de la plantilla (ej. "PC").
        2. Consultar `Tb_Mapeo_Campos` para esa plantilla.
        3. Recorrer los resultados, obteniendo datos de la tabla de datos correspondiente (ej. `Tb_Datos_PC`).
        4. Abrir la plantilla de Word y rellenar los marcadores (`NombreCampoWord`) con los valores.
        5. Guardar el documento con un nombre versionado."
*   **`modLogging`:**
    *   **Prompt:** "Crea un módulo estándar de VBA `modLogging`. Escribe dos procedimientos públicos: `LogChange(solicitudID As Long, fieldName As String, oldValue As String, newValue As String, action As String)` y `LogError(errNumber As Long, errDescription As String, errSource As String)`. Estos procedimientos insertarán registros en `Tb_Log_Cambios` y `Tb_Log_Errores` respectivamente."

### 3.3. Clases de Negocio
Para cada clase, proporciona el contexto de la interfaz y la tabla correspondiente.

*   **`CSolicitudPC`:**
    *   **Prompt:** "Crea un Módulo de Clase VBA `CSolicitudPC` que implemente la interfaz `ISolicitud`.
        1.  Declara variables miembro privadas para cada campo de la tabla `Tb_Datos_PC`.
        2.  Crea `Property Get/Let` públicos para cada variable.
        3.  Implementa el método `Load` para cargar datos de `Tb_Solicitudes` y `Tb_Datos_PC` en las variables miembro.
        4.  Implementa el método `Save` para persistir las variables miembro en la base de datos."
*   **(Repetir para `CSolicitudCDCA` y `CSolicitudCDCASUB`)**
    *   **Prompt:** "Crea un Módulo de Clase VBA `CSolicitudCDCA` que implemente `ISolicitud` y se mapee a la tabla `Tb_Datos_CD_CA`..."
    *   **Prompt:** "Crea un Módulo de Clase VBA `CSolicitudCDCASUB` que implemente `ISolicitud` y se mapee a la tabla `Tb_Datos_CD_CA_SUB`..."

---

## 4. Fase 3: Implementación de Funcionalidades Clave
Ahora se implementa la lógica específica de la aplicación.

### 4.1. Máquina de Estados
*   **Prompt:** "En la clase base o en un módulo compartido, crea una función `ChangeState(solicitudID As Long, currentState As String, desiredState As String, userRole As String) As Boolean`. Esta función debe contener un `Select Case` para `currentState` para validar si la transición a `desiredState` está permitida según el flujo de trabajo definido en la sección 3. También debe verificar si el `userRole` tiene permiso para este cambio. Si es válido, actualiza `EstadoInterno` en `Tb_Solicitudes`."

### 4.2. Generación de `CodigoSolicitud`
*   **Prompt:** "Escribe una función de VBA `GenerateNewCodigoSolicitud(tipoSolicitud As String) As String`. Debería consultar `Tb_Solicitudes` para encontrar el último código usado para el tipo dado (ej. "PC-YYYY-NNN"), incrementar el número secuencial y devolver el nuevo código."

### 4.3. Notificaciones por Correo
*   **Prompt:** "Crea un módulo estándar de VBA `modMail`. Escribe un procedimiento `QueueEmail(recipient As String, subject As String, body As String)`. Este procedimiento insertará un nuevo registro en una tabla central de cola de correos (asumiendo que existe una según la arquitectura)." 

---

## 5. Fase 4: Interfaz de Usuario (Formularios)
Generar el código VBA para los formularios.

### 5.1. Formulario Principal (`Form0BDPrincipal`)
*   **Prompt:** "Para el evento `Form_Load` de `Form0BDPrincipal`, escribe código VBA para:
    1.  Verificar el rol del usuario.
    2.  Poblar un ListBox con una lista de solicitudes activas de `Tb_Solicitudes`.
    3.  Configurar la visibilidad y el estado (activado/desactivado) de los botones según el rol del usuario."
*   **Prompt:** "Para el evento `AfterUpdate` de los controles de filtro en `Form0BDPrincipal`, escribe código VBA para reconstruir el `RowSource` del ListBox principal basándose en los criterios de filtro seleccionados."

### 5.2. Formulario de Detalle (`FormDetalleSolicitud`)
*   **Prompt:** "Para el evento `Form_Load` de `FormDetalleSolicitud`, escribe código VBA que:
    1.  Recupere el `solicitudID` desde `OpenArgs`.
    2.  Use `modFactory.CreateSolicitud` para obtener el objeto de solicitud apropiado.
    3.  Llame al método `Load` del objeto.
    4.  Pueble los controles del formulario con las propiedades del objeto.
    5.  Establezca el estado de la UI del formulario (controles activados/desactivados) basándose en el `EstadoInterno` de la solicitud y el rol del usuario, según la matriz de la sección 7.1."
*   **Prompt:** "Para el evento `Click` del botón 'Guardar' en `FormDetalleSolicitud`, escribe código VBA para transferir datos desde los controles del formulario a las propiedades del objeto de negocio y luego llamar al método `Save` del objeto."

---

## 6. Fase 5: Pruebas y Calidad
Crear el framework de testing para asegurar la fiabilidad del código.

### 6.1. Framework de Pruebas
*   **Prompt:** "Crea un módulo estándar de VBA `modAssert`. Debe contener funciones de aserción como `Assert_AreEqual(expected, actual, message)`, `Assert_IsTrue(condition, message)` y `Assert_IsNotNull(obj, message)`. Estas funciones deben lanzar un error específico si la aserción falla."
*   **Prompt:** "Crea un módulo estándar de VBA `Test_CSolicitudPC`. Escribe una función de prueba `Test_Load_PopulatesAllFields()`. Esta prueba debería:
    1.  Crear un registro de prueba conocido en la base de datos.
    2.  Crear una instancia de `CSolicitudPC`.
    3.  Llamar al método `Load` con el ID del registro de prueba.
    4.  Usar funciones de `modAssert` para verificar que cada propiedad del objeto coincida con los datos del registro de prueba."

### 6.2. Manejo de Errores
*   **Prompt:** "Muéstrame cómo añadir manejo de errores estructurado a la función `FillWordTemplate`. Debería tener un bloque `Catch` que llame a `modLogging.LogError` y luego vuelva a lanzar el error para que sea manejado por el procedimiento que lo llamó en la capa de la interfaz de usuario."
