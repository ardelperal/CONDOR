VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNotificationService"
Attribute VB_GlobalNameSpace = False
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

' Método público para inyectar dependencias
Public Sub Initialize(config As IConfig, logger As IOperationLogger, repository As INotificationRepository)
    Dim ErrorHandler As IErrorHandlerService
    On Error GoTo ErrorHandler
    
    Set m_Config = config
    Set m_OperationLogger = logger
    Set m_NotificationRepository = repository
    
    Exit Sub
    
ErrorHandler:
    Set ErrorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    ErrorHandler.LogError Err.Number, Err.Description, "CNotificationService.Initialize"
End Sub

' Implementación del método EnviarNotificacion de la interfaz
Private Function INotificationService_EnviarNotificacion( _
    ByVal destinatarios As String, _
    ByVal asunto As String, _
    ByVal cuerpoHTML As String, _
    Optional ByVal urlAdjunto As String = "" _
) As Boolean
    Dim ErrorHandler As IErrorHandlerService
    On Error GoTo ErrorHandler
    
    ' Validar que las dependencias estén inyectadas
    If m_Config Is Nothing Or m_OperationLogger Is Nothing Or m_NotificationRepository Is Nothing Then
        INotificationService_EnviarNotificacion = False
        Exit Function
    End If
    
    ' Validar parámetros
    If Len(Trim(destinatarios)) = 0 Or Len(Trim(asunto)) = 0 Or Len(Trim(cuerpoHTML)) = 0 Then
        INotificationService_EnviarNotificacion = False
        Exit Function
    End If
    
    ' Obtener configuración para el repositorio
    Dim usuarioActual As String
    Dim correoAdministrador As String
    
    usuarioActual = m_Config.GetUsuarioActual()
    correoAdministrador = m_Config.GetCorreoAdministrador()
    
    ' Llamar al repositorio para encolar el correo (el repositorio calcula el ID)
    Dim resultado As Boolean
    resultado = m_NotificationRepository.EncolarCorreo(destinatarios, asunto, cuerpoHTML, usuarioActual, correoAdministrador, urlAdjunto)
    
    If Not resultado Then
        INotificationService_EnviarNotificacion = False
        Exit Function
    End If
    
    ' Registrar operación exitosa
    m_OperationLogger.LogOperation "NOTIFICATION", "SUCCESS", _
        "Correo encolado exitosamente. Destinatarios: " & destinatarios
    
    INotificationService_EnviarNotificacion = True
    Exit Function
    
ErrorHandler:
    Set ErrorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    ErrorHandler.LogError Err.Number, Err.Description, "CNotificationService.EnviarNotificacion"
    
    ' No hay recursos específicos que limpiar ya que el repositorio maneja la BD
    
    INotificationService_EnviarNotificacion = False
End Function

'==============================================================================
' MÉTODOS PÚBLICOS DE CONVENIENCIA
'==============================================================================

' Método público de conveniencia que delega a la implementación de la interfaz
Public Function EnviarNotificacion( _
    ByVal destinatarios As String, _
    ByVal asunto As String, _
    ByVal cuerpoHTML As String, _
    Optional ByVal urlAdjunto As String = "" _
) As Boolean
    EnviarNotificacion = INotificationService_EnviarNotificacion(destinatarios, asunto, cuerpoHTML, urlAdjunto)
End Function

' Destructor
Private Sub Class_Terminate()
    Set m_Config = Nothing
    Set m_OperationLogger = Nothing
    Set m_NotificationRepository = Nothing
End Sub
















