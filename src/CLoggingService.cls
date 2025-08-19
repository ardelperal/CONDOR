VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CLoggingService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database


Option Explicit

Implements ILoggingService

' Clase: CLoggingService
' Descripción: Implementación concreta del servicio de logging del sistema CONDOR
' Arquitectura: Capa de Servicios - Implementación
' Implementa: ILoggingService

Private configSvc As IConfig
Private fso As Object ' Scripting.FileSystemObject

Private Sub Class_Initialize()
    Set configSvc = New CConfig ' TODO: Usar Factory cuando exista
    Set fso = CreateObject("Scripting.FileSystemObject")
End Sub

Private Sub Class_Terminate()
    Set configSvc = Nothing
    Set fso = Nothing
End Sub

Private Sub ILoggingService_LogError(ByVal errNumber As Long, ByVal errDescription As String, Optional ByVal moduleName As String, Optional ByVal context As String)
    WriteToLog "ERROR", errDescription, moduleName, errNumber, context
End Sub

Private Sub ILoggingService_LogInfo(ByVal message As String, Optional ByVal moduleName As String)
    WriteToLog "INFO", message, moduleName
End Sub

Private Sub WriteToLog(level As String, message As String, moduleName As String, Optional errNumber As Long, Optional context As String)
    On Error Resume Next
    Dim logFile As Object
    Dim json As String
    Dim LogPath As String

    LogPath = configSvc.GetLogPath()

    ' Abrir el fichero en modo Append (8)
    Set logFile = fso.OpenTextFile(LogPath, 8, True)

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
    ' Función auxiliar para escapar caracteres en JSON
    Dim temp As String
    temp = Replace(str, "\", "\\")
    temp = Replace(temp, """""", "\\""")
    temp = Replace(temp, vbCrLf, "\n")
    temp = Replace(temp, vbCr, "\n")
    temp = Replace(temp, vbLf, "\n")
    EscapeJSON = temp
End Function





