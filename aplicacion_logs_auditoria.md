# Aplicación de Logs de Auditoría en CONDOR

Este documento identifica los puntos clave dentro del código de la aplicación CONDOR donde sería beneficioso implementar el logging de auditoría (o logging de operaciones) utilizando el nuevo servicio `IOperationLogger`. El objetivo es registrar las acciones importantes realizadas por los usuarios y el sistema, proporcionando trazabilidad y soporte para auditorías.

## Criterios Generales para el Logging de Auditoría:

Una operación es candidata para ser loggeada si:
*   Modifica el estado de una entidad principal del negocio (creación, actualización, eliminación).
*   Representa un paso significativo en un flujo de trabajo o proceso de negocio.
*   Implica una una interacción con un sistema externo.
*   Es una acción crítica de seguridad o configuración.

## Posibles Puntos de Integración para el Logging de Auditoría:

A continuación, se listan los servicios y funciones donde se podría integrar el `IOperationLogger`, junto con ejemplos de `operationType` y `entityId`.

### 1. Gestión de Solicitudes (`CSolicitudService.cls`, `CSolicitudPC.cls`, `CSolicitudRepository.cls`)

Las solicitudes son la entidad central de la aplicación, por lo que cualquier cambio en ellas es de alta relevancia.

*   **Creación de Solicitud:**
    *   **Función:** `CSolicitudService.CreateSolicitud` (o `ISolicitudService.CreateNuevaSolicitud`)
    *   **`LogOperation`:** `m_OperationLogger.LogOperation "Crear Solicitud", CStr(newSolicitudID), "Nueva solicitud creada"`
*   **Actualización de Solicitud:**
    *   **Función:** `CSolicitudService.UpdateSolicitud` (o `ISolicitudService.SaveSolicitud` si maneja updates)
    *   **`LogOperation`:** `m_OperationLogger.LogOperation "Actualizar Solicitud", CStr(solicitud.idSolicitud), "Datos de solicitud actualizados"`
*   **Cambio de Estado de Solicitud:** (Muy crítico para auditoría)
    *   **Función:** `CSolicitudService.ChangeEstado` (o `ISolicitudService.UpdateEstadoSolicitud`)
    *   **`LogOperation`:** `m_OperationLogger.LogOperation "Cambio Estado Solicitud", CStr(solicitudID), "Estado cambiado de " & oldState & " a " & newState`
*   **Eliminación de Solicitud:**
    *   **Función:** `CSolicitudService.DeleteSolicitud`
    *   **`LogOperation`:** `m_OperationLogger.LogOperation "Eliminar Solicitud", CStr(solicitudID), "Solicitud eliminada"`
*   **Carga de Solicitud (para auditoría de acceso):**
    *   **Función:** `CSolicitudPC.ISolicitud_Load` (si se quiere registrar cada acceso a una solicitud)
    *   **`LogOperation`:** `m_OperationLogger.LogOperation "Cargar Solicitud", CStr(ID), "Solicitud cargada para visualización/edición"`

### 2. Gestión de Expedientes (`CExpedienteService.cls`)

Si la aplicación permite la creación o modificación directa de expedientes.

*   **Creación/Actualización de Expediente:**
    *   **Función:** `CExpedienteService.SaveExpediente` (o `IExpedienteService.SaveExpediente`)
    *   **`LogOperation`:** `m_OperationLogger.LogOperation "Guardar Expediente", CStr(expediente.idExpediente), "Expediente guardado/actualizado"`

### 3. Generación y Envío de Documentos/Notificaciones (`CDocumentService.cls`, `CNotificationService.cls`)

*   **Generación de Documento:**
    *   **Función:** `CDocumentService.GenerarDocumento` (o `IDocumentService_GenerarDocumento`)
    *   **`LogOperation`:** `m_OperationLogger.LogOperation "Generar Documento", CStr(idSolicitud), "Documento generado: " & rutaDocumento`
*   **Envío de Notificación/Correo:**
    *   **Función:** `CNotificationService.EnviarNotificacion` (o `INotificationService.EnviarNotificacion`)
    *   **`LogOperation`:** `m_OperationLogger.LogOperation "Enviar Notificacion", destinatarios, "Notificación enviada. Asunto: " & asunto`

### 4. Autenticación y Roles (`CAuthService.cls`)

*   **Inicio de Sesión Exitoso:**
    *   **Función:** `CAuthService.AuthenticateUser` (después de una autenticación exitosa)
    *   **`LogOperation`:** `m_OperationLogger.LogOperation "Login Exitoso", Email, "Usuario ha iniciado sesión"`
*   **Inicio de Sesión Fallido:**
    *   **Función:** `CAuthService.AuthenticateUser` (después de un intento fallido)
    *   **`LogOperation`:** `m_OperationLogger.LogOperation "Login Fallido", Email, "Intento de inicio de sesión fallido"`
*   **Cambio de Rol de Usuario:** (Si existe un servicio de gestión de usuarios)
    *   **Función:** `CUserService.ChangeUserRole` (ejemplo)
    *   **`LogOperation`:** `m_OperationLogger.LogOperation "Cambio Rol Usuario", CStr(userID), "Rol cambiado a: " & newRole`

### 5. Configuración del Sistema (`CConfig.cls`)

*   **Actualización de Configuración:**
    *   **Función:** `CConfig.SetValue` (si se permite la modificación de configuración en tiempo de ejecución)
    *   **`LogOperation`:** `m_OperationLogger.LogOperation "Actualizar Configuracion", key, "Valor cambiado a: " & value`

### 6. Interacciones con Sistemas Externos (`modDatabase.bas`, Repositorios)

*   **Ejecución de Consultas Externas:**
    *   **Función:** `modDatabase.ExecuteExternalQuery` (para registrar llamadas a sistemas externos como Lanzadera o Expedientes)
    *   **`LogOperation`:** `m_OperationLogger.LogOperation "Consulta Externa DB", dbPath, "SQL: " & sql`
*   **Operaciones de Repositorio (CRUD):**
    *   **Función:** `CSolicitudRepository.GuardarSolicitud`, `ObtenerSolicitudPorId`, `EliminarSolicitud`
    *   **`LogOperation`:** `m_OperationLogger.LogOperation "Repo Guardar Solicitud", CStr(solicitud.idSolicitud), "Solicitud persistida"`

## Ejemplo de Implementación en Código:

Para integrar el logger, primero asegúrate de que la clase que lo necesita tenga una referencia a `IOperationLogger` (inyectada a través de su método `Initialize` o un factory).

```vba
' Ejemplo en CSolicitudService.cls

Private m_OperationLogger As IOperationLogger
Private m_SolicitudRepository As ISolicitudRepository ' Asumiendo que ya existe

' Método de inicialización para inyectar dependencias
Public Sub Initialize(ByVal solicitudRepository As ISolicitudRepository, ByVal operationLogger As IOperationLogger)
    Set m_SolicitudRepository = solicitudRepository
    Set m_OperationLogger = operationLogger
End Sub

' Método de ejemplo donde se usaría el logger
Private Function ISolicitudService_SaveSolicitud(solicitud As ISolicitud) As Boolean
    On Error GoTo ErrorHandler
    
    Dim success As Boolean
    ' ... lógica para guardar la solicitud en el repositorio ...
    success = m_SolicitudRepository.Save(solicitud) ' Ejemplo de llamada al repositorio
    
    If success Then
        ' Registrar la operación después de que sea exitosa
        m_OperationLogger.LogOperation "Guardar Solicitud", CStr(solicitud.idSolicitud), "Solicitud guardada/actualizada exitosamente."
    End If
    
    ISolicitudService_SaveSolicitud = success
    Exit Function
    
ErrorHandler:
    Call modErrorHandler.LogError(Err.Number, Err.Description, "CSolicitudService.ISolicitudService_SaveSolicitud")
    ISolicitudService_SaveSolicitud = False
End Function
```

La implementación de estos logs de auditoría proporcionará una visibilidad sin precedentes sobre el funcionamiento interno de la aplicación y las acciones de los usuarios, lo que será invaluable para la depuración, la auditoría y la toma de decisiones.
