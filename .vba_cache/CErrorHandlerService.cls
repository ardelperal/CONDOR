VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CErrorHandlerService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Implements IErrorHandlerService

' Clase: CErrorHandlerService
' Descripción: Implementación concreta del servicio de manejo de errores del sistema CONDOR
' Arquitectura: Capa de Servicios - Implementación
' Implementa: IErrorHandlerService
' Método principal para registrar errores en el sistema

Private m_configService As IConfig
Private m_FileSystem As IFileSystem

Public Sub Initialize(config As IConfig, fileSystem As IFileSystem)
    On Error GoTo ErrorHandler
    Set m_configService = config
    Set m_FileSystem = fileSystem
    
    Exit Sub
ErrorHandler:
    Debug.Print "Error fatal inicializando CErrorHandlerService"
End Sub

Private Sub Class_Terminate()
    Set m_configService = Nothing
    Set m_FileSystem = Nothing
End Sub

' --- IMPLEMENTACION CORREGIDA ---
Private Sub IErrorHandlerService_LogError(ByVal errNumber As Long, ByVal errDescription As String, Optional ByVal moduleName As String, Optional ByVal isCritical As Boolean = False)
    On Error GoTo ErrorHandler
    
    ' Registrar en el log de operación si está disponible
    ' Llamar al helper de escritura con el nuevo parametro isCritical
    WriteToLog "ERROR", errDescription, moduleName, errNumber, isCritical
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error en CErrorHandlerService.IErrorHandlerService_LogError: " & Err.Description
End Sub

Private Sub IErrorHandlerService_LogInfo(ByVal message As String, Optional ByVal moduleName As String)
    On Error GoTo ErrorHandler
    
    ' Llamar al helper de escritura para INFO
    WriteToLog "INFO", message, moduleName, 0, False
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error en CErrorHandlerService.IErrorHandlerService_LogInfo: " & Err.Description
End Sub

Private Sub IErrorHandlerService_LogWarning(ByVal message As String, Optional ByVal moduleName As String)
    On Error GoTo ErrorHandler
    
    ' Llamar al helper de escritura para WARNING
    WriteToLog "WARNING", message, moduleName, 0, False
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error en CErrorHandlerService.IErrorHandlerService_LogWarning: " & Err.Description
End Sub

' ============================================================================
' MÉTODOS PÚBLICOS DE CONVENIENCIA (Lección 24)
' ============================================================================

Public Sub LogError(ByVal errNumber As Long, ByVal errDescription As String, ByVal moduleName As String, Optional ByVal isCritical As Boolean = False)
    IErrorHandlerService_LogError errNumber, errDescription, moduleName, isCritical
End Sub

Public Sub LogInfo(ByVal message As String, Optional ByVal moduleName As String)
    IErrorHandlerService_LogInfo message, moduleName
End Sub

Public Sub LogWarning(ByVal message As String, Optional ByVal moduleName As String)
    IErrorHandlerService_LogWarning message, moduleName
End Sub

' --- HELPER DE ESCRITURA CORREGIDO ---
Private Sub WriteToLog(level As String, message As String, moduleName As String, errNumber As Long, isCritical As Boolean)
    On Error GoTo ErrorHandler
    Dim logFile As Object
    Dim json As String
    Dim LogPath As String

    LogPath = m_configService.GetValue("LOG_FILE_PATH")

    Set logFile = m_FileSystem.OpenTextFile(LogPath, 8, True) ' 8 = ForAppending

    json = "{"
    json = json & """timestamp"" : """ & Format(Now, "yyyy-mm-dd hh:nn:ss") & """," 
    json = json & """level"" : """ & level & """," 
    json = json & """user"" : """ & Application.CurrentUser & """," 
    json = json & """module"" : """ & moduleName & """," 
    json = json & """message"" : """ & EscapeJSON(message) & """
    If errNumber <> 0 Then json = json & ",""errorNumber"" : " & errNumber
    ' --- CAMBIO: Usar isCritical en lugar de context ---
    If isCritical Then json = json & ",""isCritical"" : true"
    json = json & "}"

    logFile.WriteLine json
    logFile.Close
    Set logFile = Nothing
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error en CErrorHandlerService.WriteToLog: " & Err.Description
    If Not logFile Is Nothing Then
        logFile.Close
        Set logFile = Nothing
    End If
End Sub

Private Function EscapeJSON(ByVal str As String) As String
    On Error GoTo ErrorHandler
    
    Dim temp As String
    temp = str
    
    temp = Replace(temp, "\", "\\")
    temp = Replace(temp, Chr(34), "\" & Chr(34))
    temp = Replace(temp, vbTab, "\t")
    temp = Replace(temp, vbCrLf, "\n")
    temp = Replace(temp, vbCr, "\n")
    temp = Replace(temp, vbLf, "\n")
    temp = Replace(temp, Chr(8), "\b")
    temp = Replace(temp, Chr(12), "\f")
    temp = Replace(temp, "/", "\/")
    
    EscapeJSON = temp
    
    Exit Function
    
ErrorHandler:
    Debug.Print "Error en CErrorHandlerService.EscapeJSON: " & Err.Description
    EscapeJSON = str
End Function








