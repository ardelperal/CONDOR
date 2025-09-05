VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNotificationRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


' =====================================================
' Clase: CNotificationRepository
' Propósito: Implementación concreta del repositorio de notificaciones
' Autor: CONDOR-Expert
' Fecha: 2025-01-15
' =====================================================

Implements INotificationRepository

' Dependencias inyectadas
Private m_configService As IConfig
Private m_ErrorHandler As IErrorHandlerService
Private m_isInitialized As Boolean

' Inicializa el repositorio con sus dependencias
' @param configService: Servicio de configuración para obtener ruta de la BD de correos
' @param errorHandler: Servicio de manejo de errores
Public Sub Initialize(ByVal configService As IConfig, ByVal errorHandler As IErrorHandlerService)
    If configService Is Nothing Then
        Err.Raise 5, "CNotificationRepository.Initialize", "configService no puede ser Nothing"
    End If
    
    If errorHandler Is Nothing Then
        Err.Raise 5, "CNotificationRepository.Initialize", "errorHandler no puede ser Nothing"
    End If
    
    Set m_configService = configService
    Set m_ErrorHandler = errorHandler
    m_isInitialized = True
End Sub

' Implementación de INotificationRepository_EnqueueEmail
Private Function INotificationRepository_EnqueueEmail(ByVal destinatarios As String, ByVal asunto As String, ByVal cuerpoHTML As String, ByVal usuario As String, ByVal adminEmail As String, Optional ByVal adjuntoPath As String = "") As Boolean
    On Error GoTo errorHandler
    
    ' Validar inicialización
    If Not m_isInitialized Then
        m_ErrorHandler.LogError 5, "Repositorio no inicializado", "CNotificationRepository.EnqueueEmail"
        INotificationRepository_EnqueueEmail = False
        Exit Function
    End If
    
    ' Validar parámetros obligatorios
    If Len(Trim(destinatarios)) = 0 Or Len(Trim(asunto)) = 0 Or Len(Trim(cuerpoHTML)) = 0 Then
        m_ErrorHandler.LogError 5, "Parámetros obligatorios vacíos", "CNotificationRepository.EnqueueEmail"
        INotificationRepository_EnqueueEmail = False
        Exit Function
    End If
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim correosDbPath As String
    Dim dbPassword As String
    
    ' Obtener configuración de la base de datos de correos
    correosDbPath = m_configService.GetCorreosDBPath()
    dbPassword = m_configService.GetDatabasePassword()
    
    ' Validar que la configuración esté disponible
    If Len(Trim(correosDbPath)) = 0 Then
        m_ErrorHandler.LogError 5, "Ruta de base de datos de correos no configurada", "CNotificationRepository.EnqueueEmail"
        INotificationRepository_EnqueueEmail = False
        Exit Function
    End If
    
    ' Abrir conexión a la base de datos de correos
    If Len(Trim(dbPassword)) > 0 Then
        Set db = DBEngine.OpenDatabase(correosDbPath, False, False, ";PWD=" & dbPassword)
    Else
        Set db = DBEngine.OpenDatabase(correosDbPath, False, False)
    End If
    
    ' Abrir tabla de correos para insertar nuevo registro
    Set rs = db.OpenRecordset("TbCorreosEnviados", dbOpenDynaset)
    
    ' Insertar nuevo registro
    rs.AddNew
    rs("Destinatarios") = destinatarios
    rs("Asunto") = asunto
    rs("Cuerpo") = cuerpoHTML
    rs("DestinatariosConCopia") = usuario
    rs("DestinatariosConCopiaOculta") = adminEmail
    rs("URLAdjunto") = adjuntoPath
    rs("FechaGrabacion") = Now()
    rs.Update
    
    ' Cerrar recursos
    rs.Close
    db.Close
    
    Set rs = Nothing
    Set db = Nothing
    
    INotificationRepository_EnqueueEmail = True
    Exit Function
    
errorHandler:
    ' Limpiar recursos en caso de error
    If Not rs Is Nothing Then
        If rs.Updatable Then rs.CancelUpdate
        rs.Close
        Set rs = Nothing
    End If
    
    If Not db Is Nothing Then
        db.Close
        Set db = Nothing
    End If
    
    m_ErrorHandler.LogError Err.Number, Err.Description, "CNotificationRepository.EnqueueEmail"
    INotificationRepository_EnqueueEmail = False
End Function

' ============================================================================
' MÉTODOS PÚBLICOS DE CONVENIENCIA
' ============================================================================
' Estos métodos exponen públicamente la funcionalidad de la interfaz
' para facilitar las pruebas y el uso directo de la clase concreta

' Método público de conveniencia para EnqueueEmail
Public Function EnqueueEmail(ByVal destinatarios As String, ByVal asunto As String, ByVal cuerpoHTML As String, ByVal usuario As String, ByVal adminEmail As String, Optional ByVal adjuntoPath As String = "") As Boolean
    EnqueueEmail = INotificationRepository_EnqueueEmail(destinatarios, asunto, cuerpoHTML, usuario, adminEmail, adjuntoPath)
End Function

' Destructor
Private Sub Class_Terminate()
    Set m_configService = Nothing
    Set m_ErrorHandler = Nothing
End Sub

