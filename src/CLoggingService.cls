Option Compare Database
Option Explicit
Implements ILoggingService

' Clase: CLoggingService
' Descripción: Implementación concreta del servicio de logging del sistema CONDOR
' Arquitectura: Capa de Servicios - Implementación
' Implementa: ILoggingService

' Dependencias inyectadas
Private m_Config As IConfig
Private m_FileSystem As IFileSystem

' Método para inyectar dependencias
Public Sub Initialize(ByVal config As IConfig, ByVal fileSystem As IFileSystem)
    Set m_Config = config
    Set m_FileSystem = fileSystem
End Sub

Private Sub Class_Terminate()
    Set m_Config = Nothing
    Set m_FileSystem = Nothing
End Sub

Private Sub ILoggingService_LogError(ByVal errNumber As Long, ByVal errDescription As String, Optional ByVal moduleName As String, Optional ByVal context As String)
    On Error GoTo ErrorHandler
    
    WriteToLog "ERROR", errDescription, moduleName, errNumber, context
    
    Exit Sub
    
ErrorHandler:
    ' No podemos usar modErrorHandler aquí para evitar recursión infinita
    Debug.Print "Error en CLoggingService.ILoggingService_LogError: " & Err.Description
End Sub

Private Sub ILoggingService_LogInfo(ByVal message As String, Optional ByVal moduleName As String)
    On Error GoTo ErrorHandler
    
    WriteToLog "INFO", message, moduleName
    
    Exit Sub
    
ErrorHandler:
    ' No podemos usar modErrorHandler aquí para evitar recursión infinita
    Debug.Print "Error en CLoggingService.ILoggingService_LogInfo: " & Err.Description
End Sub

Private Sub ILoggingService_LogWarning(ByVal message As String, Optional ByVal moduleName As String)
    On Error GoTo ErrorHandler
    
    WriteToLog "WARNING", message, moduleName
    
    Exit Sub
    
ErrorHandler:
    ' No podemos usar modErrorHandler aquí para evitar recursión infinita
    Debug.Print "Error en CLoggingService.ILoggingService_LogWarning: " & Err.Description
End Sub

Public Sub LogWarning(ByVal message As String, Optional ByVal moduleName As String)
    ILoggingService_LogWarning message, moduleName
End Sub

Private Sub WriteToLog(level As String, message As String, moduleName As String, Optional errNumber As Long, Optional context As String)
    On Error Resume Next
    Dim logFile As Object
    Dim json As String
    Dim LogPath As String

    LogPath = m_Config.GetLogPath()

    ' Abrir el fichero en modo Append (8)
    Set logFile = m_FileSystem.OpenTextFile(LogPath, 8, True)

    ' Construir la línea de log en formato JSON
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
    
    ' Función auxiliar para escapar caracteres en JSON
    Dim temp As String
    temp = Replace(str, "\", "\\")
    temp = Replace(temp, """""", "\\""")
    temp = Replace(temp, vbCrLf, "\n")
    temp = Replace(temp, vbCr, "\n")
    temp = Replace(temp, vbLf, "\n")
    EscapeJSON = temp
    
    Exit Function
    
ErrorHandler:
    ' No podemos usar modErrorHandler aquí para evitar recursión infinita
    Debug.Print "Error en CLoggingService.EscapeJSON: " & Err.Description
    EscapeJSON = str ' Devolver string original si falla el escape
End Function






