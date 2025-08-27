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

Implements INotificationRepository

'******************************************************************************
' CLASE: CNotificationRepository
' Propósito: Repositorio concreto para operaciones de notificaciones en BD
' Patrón: Repository Pattern
' Versión: 1.0
'******************************************************************************

' Dependencia inyectada
Private m_Config As IConfig

' Método público para inyectar dependencias
Public Sub Initialize(config As IConfig)
    Dim ErrorHandler As IErrorHandlerService
    On Error GoTo ErrorHandler
    
    Set m_Config = config
    
    Exit Sub
    
ErrorHandler:
    Set ErrorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    ErrorHandler.LogError Err.Number, Err.Description, "CNotificationRepository.Initialize"
End Sub

' Implementación del método de la interfaz
Private Function INotificationRepository_EncolarCorreo(ByVal destinatarios As String, ByVal asunto As String, ByVal cuerpo As String, ByVal cc As String, ByVal bcc As String, ByVal adjunto As String) As Boolean
    Dim ErrorHandler As IErrorHandlerService
    On Error GoTo ErrorHandler
    
    ' Validar que la dependencia esté inyectada
    If m_Config Is Nothing Then
        INotificationRepository_EncolarCorreo = False
        Exit Function
    End If
    
    ' Obtener configuración desde IConfig inyectado
    Dim rutaCorreosDB As String
    Dim passwordDB As String
    
    rutaCorreosDB = m_Config.GetCorreosDBPath()
    passwordDB = m_Config.GetDatabasePassword()
    
    If Len(rutaCorreosDB) = 0 Then
        INotificationRepository_EncolarCorreo = False
        Exit Function
    End If
    
    ' Conectar a la base de datos de correos
    Dim db As DAO.Database
    Set db = DBEngine.OpenDatabase(rutaCorreosDB, dbFailOnError, False, ";PWD=" & passwordDB)
    
    ' Calcular nuevo IDCorreo usando consulta SQL robusta
    Dim nuevoId As Long
    Dim rsMaxId As DAO.Recordset
    Set rsMaxId = db.OpenRecordset("SELECT MAX(IDCorreo) AS MaxID FROM TbCorreosEnviados")
    nuevoId = Nz(rsMaxId!MaxID, 0) + 1
    rsMaxId.Close
    Set rsMaxId = Nothing
    
    ' Crear consulta parametrizada para eliminar vulnerabilidad SQL injection
    Dim qdf As DAO.QueryDef
    Set qdf = db.CreateQueryDef("", _
        "INSERT INTO TbCorreosEnviados (IDCorreo, Aplicacion, Asunto, Cuerpo, " & _
        "Destinatarios, DestinatariosConCopia, DestinatariosConCopiaOculta, " & _
        "URLAdjunto, FechaGrabacion, FechaEnvio) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)")
    
    ' Asignar parámetros de forma segura
    qdf.parameters(0) = nuevoId
    qdf.parameters(1) = "CONDOR"
    qdf.parameters(2) = asunto
    qdf.parameters(3) = cuerpo
    qdf.parameters(4) = destinatarios
    qdf.parameters(5) = cc
    qdf.parameters(6) = bcc
    qdf.parameters(7) = adjunto
    qdf.parameters(8) = Now()
    qdf.parameters(9) = Null
    
    ' Ejecutar consulta parametrizada
    qdf.Execute
    
    ' Limpiar recursos
    qdf.Close
    Set qdf = Nothing
    db.Close
    Set db = Nothing
    
    INotificationRepository_EncolarCorreo = True
    Exit Function
    
ErrorHandler:
    Set ErrorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    ErrorHandler.LogError Err.Number, Err.Description, "CNotificationRepository.EncolarCorreo"
    
    ' Limpiar recursos en caso de error
    If Not qdf Is Nothing Then
        qdf.Close
        Set qdf = Nothing
    End If
    If Not db Is Nothing Then
        db.Close
        Set db = Nothing
    End If
End Function

'==============================================================================
' MÉTODOS PÚBLICOS DE CONVENIENCIA
'==============================================================================

' Método público de conveniencia que delega a la implementación de la interfaz
Public Function EncolarCorreo(ByVal destinatarios As String, ByVal asunto As String, ByVal cuerpo As String, ByVal cc As String, ByVal bcc As String, ByVal adjunto As String) As Boolean
    EncolarCorreo = INotificationRepository_EncolarCorreo(destinatarios, asunto, cuerpo, cc, bcc, adjunto)
End Function

' Destructor
Private Sub Class_Terminate()
    Set m_Config = Nothing
End Sub