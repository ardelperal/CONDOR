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

Private m_configService As IConfig
Private m_ErrorHandler As IErrorHandlerService
Private m_isInitialized As Boolean

Public Sub Initialize(ByVal configSrv As IConfig, ByVal errorSrv As IErrorHandlerService)
    Set m_configService = configSrv
    Set m_ErrorHandler = errorSrv
    m_isInitialized = True
End Sub

Private Function INotificationRepository_EnqueueEmail(ByVal destinatarios As String, ByVal asunto As String, ByVal cuerpoHTML As String, Optional ByVal conCopia As String = "", Optional ByVal conCopiaOculta As String = "", Optional ByVal adjuntoPath As String = "") As Boolean
    On Error GoTo ErrorHandler
    
    INotificationRepository_EnqueueEmail = False
    If Not m_isInitialized Then Exit Function
    
    Dim db As DAO.Database, rs As DAO.Recordset
    Dim correosDbPath As String, dbPassword As String
    
    correosDbPath = m_configService.GetCorreosDBPath()
    dbPassword = m_configService.GetCorreosPassword()
    
    If Len(correosDbPath) = 0 Then Exit Function
    
    If Len(dbPassword) > 0 Then
        Set db = DBEngine.OpenDatabase(correosDbPath, False, False, ";PWD=" & dbPassword)
    Else
        Set db = DBEngine.OpenDatabase(correosDbPath, False, False)
    End If
    
    Set rs = db.OpenRecordset("TbCorreosEnviados", dbOpenDynaset)
    
    rs.AddNew
    rs("Destinatarios") = destinatarios
    rs("Asunto") = asunto
    rs("Cuerpo") = cuerpoHTML
    rs("DestinatariosConCopia") = conCopia
    rs("DestinatariosConCopiaOculta") = conCopiaOculta
    rs("URLAdjunto") = adjuntoPath
    rs("FechaGrabacion") = Now()
    rs.Update
    
    rs.Close
    db.Close
    
    INotificationRepository_EnqueueEmail = True
    Exit Function
    
ErrorHandler:
    If Not rs Is Nothing Then rs.Close
    If Not db Is Nothing Then db.Close
    m_ErrorHandler.LogError Err.Number, Err.Description, "CNotificationRepository.EnqueueEmail"
End Function