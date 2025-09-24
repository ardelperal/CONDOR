<<<<<<< HEAD
﻿Attribute VB_GlobalNameSpace = False
=======
﻿VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNotificationService"
Attribute VB_GlobalNameSpace = False
>>>>>>> main
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



' CNotificationService.cls
' Implementación segura del servicio de notificaciones asíncrono
' Reconstruido para cumplir con principios de inyección de dependencias y seguridad
' Basado en las Especificaciones de Integración - Sección 2

' Implementa la interfaz INotificationService
Implements INotificationService

' Variables privadas para dependencias inyectadas
Private m_Config As IConfig
Private m_OperationLogger As IOperationLogger
Private m_NotificationRepository As INotificationRepository
Private m_ErrorHandler As IErrorHandlerService

' Método público para inyectar dependencias
Public Sub Initialize(config As IConfig, logger As IOperationLogger, repository As INotificationRepository, ErrorHandler As IErrorHandlerService)
    ' Asignar ErrorHandler primero para poder usarlo en el manejo de errores
    Set m_ErrorHandler = ErrorHandler
    
    On Error GoTo ErrorHandler
    
    Set m_Config = config
    Set m_OperationLogger = logger
    Set m_NotificationRepository = repository
    
    Exit Sub
    
ErrorHandler:
    m_ErrorHandler.LogError Err.Number, Err.Description, "CNotificationService.Initialize"
End Sub

' Implementación del método SendNotification de la interfaz
Private Function INotificationService_SendNotification( _
    ByVal destinatarios As String, _
    ByVal asunto As String, _
    ByVal cuerpoHTML As String, _
    Optional ByVal conCopia As String = "", _
    Optional ByVal conCopiaOculta As String = "", _
    Optional ByVal urlAdjunto As String = "" _
) As Boolean
    On Error GoTo ErrorHandler
    
    ' Validar que las dependencias estén inyectadas
    If m_Config Is Nothing Or m_OperationLogger Is Nothing Or m_NotificationRepository Is Nothing Then
        INotificationService_SendNotification = False
        Exit Function
    End If
    
    ' Validar parámetros
    If Len(Trim(destinatarios)) = 0 Or Len(Trim(asunto)) = 0 Or Len(Trim(cuerpoHTML)) = 0 Then
        INotificationService_SendNotification = False
        Exit Function
    End If
    
    ' Obtener configuración para el repositorio
    Dim usuarioActual As String
    Dim correoAdministrador As String
    
    usuarioActual = m_Config.GetUsuarioActual()
    correoAdministrador = m_Config.GetCorreoAdministrador()
    
    ' Llamar al repositorio para encolar el correo (el repositorio calcula el ID)
    Dim resultado As Boolean
    resultado = m_NotificationRepository.EnqueueEmail(destinatarios, asunto, cuerpoHTML, conCopia, conCopiaOculta, urlAdjunto)
    
    If Not resultado Then
        INotificationService_SendNotification = False
        Exit Function
    End If
    
    ' Registrar operación exitosa
    Dim logEntry As EOperationLog
    Set logEntry = New EOperationLog
    With logEntry
        .tipoOperacion = "ENQUEUE_EMAIL"
        .idEntidadAfectada = 0 ' El ID lo gestiona el repositorio y no es devuelto al servicio
        .usuario = m_Config.GetUsuarioActual()
        .resultado = "Success"
        .descripcion = "Correo encolado para envío. Destinatarios: " & destinatarios
    End With
    m_OperationLogger.LogOperation logEntry
    
    INotificationService_SendNotification = True
    Exit Function
    
ErrorHandler:
    m_ErrorHandler.LogError Err.Number, Err.Description, "CNotificationService.SendNotification"
    
    ' No hay recursos específicos que limpiar ya que el repositorio maneja la BD
    
    INotificationService_SendNotification = False
End Function

'==============================================================================
' MÉTODOS PÚBLICOS DE CONVENIENCIA
'==============================================================================

' Método público de conveniencia que delega a la implementación de la interfaz
Public Function SendNotification( _
    ByVal destinatarios As String, _
    ByVal asunto As String, _
    ByVal cuerpoHTML As String, _
    Optional ByVal conCopia As String = "", _
    Optional ByVal conCopiaOculta As String = "", _
    Optional ByVal urlAdjunto As String = "" _
) As Boolean
    SendNotification = INotificationService_SendNotification(destinatarios, asunto, cuerpoHTML, conCopia, conCopiaOculta, urlAdjunto)
End Function

' Destructor
Private Sub Class_Terminate()
    Set m_Config = Nothing
    Set m_OperationLogger = Nothing
    Set m_NotificationRepository = Nothing
    Set m_ErrorHandler = Nothing
<<<<<<< HEAD
End Sub
=======
End Sub






















>>>>>>> main
