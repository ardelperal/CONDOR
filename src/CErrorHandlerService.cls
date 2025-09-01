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

Private m_config As IConfig
Private m_fileSystem As IFileSystem
Private m_logFilePath As String

' ============================================================================
' INICIALIZACIÓN CON DEPENDENCIAS (Lección 11)
' ============================================================================
Public Sub Initialize(ByVal config As IConfig, ByVal fileSystem As IFileSystem)
    Set m_config = config
    Set m_fileSystem = fileSystem
    m_logFilePath = m_config.GetValue("LOG_FILE_PATH")
End Sub

' ============================================================================
' IMPLEMENTACIÓN DE LA INTERFAZ IErrorHandlerService (Lección 1)
' ============================================================================
Private Sub IErrorHandlerService_Initialize(ByVal config As IConfig, ByVal fileSystem As IFileSystem)
    Call Initialize(config, fileSystem)
End Sub

Private Sub IErrorHandlerService_LogError(ByVal errNumber As Long, ByVal errDescription As String, Optional ByVal moduleName As String, Optional ByVal isCritical As Boolean = False)
    On Error Resume Next ' No podemos permitir que el manejador de errores falle
    
    Dim logPrefix As String
    If isCritical Then
        logPrefix = "CRITICAL ERROR"
    Else
        logPrefix = "ERROR"
    End If
    
    Dim logMessage As String
    logMessage = Now() & " | " & logPrefix & " | " & "Err: " & errNumber & " | " & errDescription & " | Source: " & moduleName
    
    If Len(m_logFilePath) > 0 Then
        Dim fso As Object
        Dim file As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        ' 8 = ForAppending, True = Create if not exists
        Set file = fso.OpenTextFile(m_logFilePath, 8, True)
        file.WriteLine logMessage
        file.Close
        Set file = Nothing
        Set fso = Nothing
    Else
        Debug.Print "LogFilePath no configurado. Error no guardado: " & logMessage
    End If
End Sub

Private Sub IErrorHandlerService_LogInfo(ByVal message As String, Optional ByVal moduleName As String)
    ' Implementación futura si es necesario
End Sub

Private Sub IErrorHandlerService_LogWarning(ByVal message As String, Optional ByVal moduleName As String)
    ' Implementación futura si es necesario
End Sub








