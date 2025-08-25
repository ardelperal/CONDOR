Option Compare Database
Option Explicit
Implements IErrorHandlerService

' Clase: CErrorHandlerService
' DescripciÃ³n: ImplementaciÃ³n concreta del servicio de manejo de errores del sistema CONDOR
' Arquitectura: Capa de Servicios - ImplementaciÃ³n
' Implementa: IErrorHandlerService

Private m_configService As IConfig
Private m_FileSystem As IFileSystem

Public Sub Initialize(config As IConfig, fileSystem As IFileSystem)
    On Error GoTo ErrorHandler
    Set m_configService = config
    Set m_FileSystem = fileSystem
    
    Exit Sub
ErrorHandler:
    ' No podemos usar el logger para loggear un error en el propio logger
    Debug.Print "Error fatal inicializando CErrorHandlerService"
End Sub

Private Sub Class_Terminate()
    Set m_configService = Nothing
    Set m_FileSystem = Nothing
End Sub

Private Sub IErrorHandlerService_LogError(ByVal errNumber As Long, ByVal errDescription As String, Optional ByVal moduleName As String, Optional ByVal context As String)
    On Error GoTo ErrorHandler
    
    WriteToLog "ERROR", errDescription, moduleName, errNumber, context
    
    Exit Sub
    
ErrorHandler:
    ' No podemos usar modErrorHandler aquí para evitar recursión infinita
    Debug.Print "Error en CErrorHandlerService.IErrorHandlerService_LogError: " & Err.Description
End Sub

Private Sub IErrorHandlerService_LogInfo(ByVal message As String, Optional ByVal moduleName As String)
    On Error GoTo ErrorHandler
    
    WriteToLog "INFO", message, moduleName
    
    Exit Sub
    
ErrorHandler:
    ' No podemos usar modErrorHandler aquí para evitar recursión infinita
    Debug.Print "Error en CErrorHandlerService.IErrorHandlerService_LogInfo: " & Err.Description
End Sub

Private Sub IErrorHandlerService_LogWarning(ByVal message As String, Optional ByVal moduleName As String)
    On Error GoTo ErrorHandler
    
    WriteToLog "WARNING", message, moduleName
    
    Exit Sub
    
ErrorHandler:
    ' No podemos usar modErrorHandler aquí para evitar recursión infinita
    Debug.Print "Error en CErrorHandlerService.IErrorHandlerService_LogWarning: " & Err.Description
End Sub

Public Sub LogError(ByVal errNumber As Long, ByVal errDescription As String, Optional ByVal moduleName As String, Optional ByVal context As String)
    IErrorHandlerService_LogError errNumber, errDescription, moduleName, context
End Sub

Public Sub LogInfo(ByVal message As String, Optional ByVal moduleName As String)
    IErrorHandlerService_LogInfo message, moduleName
End Sub

Public Sub LogWarning(ByVal message As String, Optional ByVal moduleName As String)
    IErrorHandlerService_LogWarning message, moduleName
End Sub

Private Sub WriteToLog(level As String, message As String, Optional moduleName As String, Optional errNumber As Long, Optional context As String)
    On Error Resume Next
    Dim logFile As Object
    Dim json As String
    Dim LogPath As String

    LogPath = m_configService.GetValue("LOG_FILE_PATH")

    ' Abrir el fichero en modo Append (8)
    Set logFile = m_FileSystem.OpenTextFile(LogPath, 8, True)

    ' Construir la lÃ­nea de log en formato JSON
    json = "{"
    json = json & """timestamp"": """ & Format(Now, "yyyy-mm-dd hh:nn:ss") & ""","
    json = json & """level"": """ & level & ""","
    json = json & """user"": """ & Application.CurrentUser & ""","
    json = json & """module"": """ & moduleName & ""","
    json = json & """message"": """ & EscapeJSON(message) & """"
    If errNumber <> 0 Then json = json & ",""errorNumber"": " & errNumber
    If context <> "" Then json = json & ",""context"": """ & EscapeJSON(context) & """"
    json = json & "}"

    logFile.WriteLine json
    logFile.Close
    Set logFile = Nothing
End Sub

Private Function EscapeJSON(ByVal str As String) As String
    On Error GoTo ErrorHandler
    
    ' Función auxiliar para escapar caracteres en JSON de forma robusta
    Dim temp As String
    temp = str
    
    ' Escapar caracteres especiales en el orden correcto
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
    ' No podemos usar modErrorHandler aquí para evitar recursión infinita
    Debug.Print "Error en CErrorHandlerService.EscapeJSON: " & Err.Description
    EscapeJSON = str ' Devolver string original si falla el escape
End Function






