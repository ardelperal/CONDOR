VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMockErrorHandlerService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

#If DEV_MODE Then

' =====================================================
' CMockErrorHandlerService.cls
' Mock implementation of IErrorHandlerService
' Para uso en pruebas unitarias aisladas
' =====================================================

Implements IErrorHandlerService

' Variables para capturar las llamadas del mock
Private m_lastErrorNumber As Long
Private m_lastErrorDescription As String
Private m_lastErrorModuleName As String
Private m_lastErrorContext As String
Private m_lastInfoMessage As String
Private m_lastInfoModuleName As String
Private m_lastWarningMessage As String
Private m_lastWarningModuleName As String
Private m_logErrorCallCount As Long
Private m_logInfoCallCount As Long
Private m_logWarningCallCount As Long

' MÃ©todo Initialize (no forma parte de la interfaz pero es Ãºtil para el mock)
Public Sub Initialize(config As IConfig)
    ' Mock no necesita configuraciÃ³n real, pero mantiene la firma
    ' para compatibilidad con el patrÃ³n de inyecciÃ³n de dependencias
End Sub

' ImplementaciÃ³n mock de LogError
Private Sub IErrorHandlerService_LogError(ByVal errNumber As Long, ByVal errDescription As String, Optional ByVal moduleName As String, Optional ByVal context As String)
    m_lastErrorNumber = errNumber
    m_lastErrorDescription = errDescription
    m_lastErrorModuleName = moduleName
    m_lastErrorContext = context
    m_logErrorCallCount = m_logErrorCallCount + 1
End Sub

' ImplementaciÃ³n mock de LogInfo
Private Sub IErrorHandlerService_LogInfo(ByVal message As String, Optional ByVal moduleName As String)
    m_lastInfoMessage = message
    m_lastInfoModuleName = moduleName
    m_logInfoCallCount = m_logInfoCallCount + 1
End Sub

' ImplementaciÃ³n mock de LogWarning
Private Sub IErrorHandlerService_LogWarning(ByVal message As String, Optional ByVal moduleName As String)
    m_lastWarningMessage = message
    m_lastWarningModuleName = moduleName
    m_logWarningCallCount = m_logWarningCallCount + 1
End Sub

' =====================================================
' MÃ©todos pÃºblicos para verificaciÃ³n en pruebas
' =====================================================

' Verificar la Ãºltima llamada a LogError
Public Function GetLastErrorNumber() As Long
    GetLastErrorNumber = m_lastErrorNumber
End Function

Public Function GetLastErrorDescription() As String
    GetLastErrorDescription = m_lastErrorDescription
End Function

Public Function GetLastErrorModuleName() As String
    GetLastErrorModuleName = m_lastErrorModuleName
End Function

Public Function GetLastErrorContext() As String
    GetLastErrorContext = m_lastErrorContext
End Function

' Verificar la Ãºltima llamada a LogInfo
Public Function GetLastInfoMessage() As String
    GetLastInfoMessage = m_lastInfoMessage
End Function

Public Function GetLastInfoModuleName() As String
    GetLastInfoModuleName = m_lastInfoModuleName
End Function

' Verificar la Ãºltima llamada a LogWarning
Public Function GetLastWarningMessage() As String
    GetLastWarningMessage = m_lastWarningMessage
End Function

Public Function GetLastWarningModuleName() As String
    GetLastWarningModuleName = m_lastWarningModuleName
End Function

' Contadores de llamadas
Public Function GetLogErrorCallCount() As Long
    GetLogErrorCallCount = m_logErrorCallCount
End Function

Public Function GetLogInfoCallCount() As Long
    GetLogInfoCallCount = m_logInfoCallCount
End Function

Public Function GetLogWarningCallCount() As Long
    GetLogWarningCallCount = m_logWarningCallCount
End Function

' Resetear el estado del mock
Public Sub Reset()
    m_lastErrorNumber = 0
    m_lastErrorDescription = ""
    m_lastErrorModuleName = ""
    m_lastErrorContext = ""
    m_lastInfoMessage = ""
    m_lastInfoModuleName = ""
    m_lastWarningMessage = ""
    m_lastWarningModuleName = ""
    m_logErrorCallCount = 0
    m_logInfoCallCount = 0
    m_logWarningCallCount = 0
End Sub

' Verificar si se llamÃ³ a LogError con parÃ¡metros especÃ­ficos
Public Function WasLogErrorCalledWith(errorNumber As Long, errorDescription As String, Optional moduleName As String = "", Optional context As String = "") As Boolean
    WasLogErrorCalledWith = (m_lastErrorNumber = errorNumber) And _
                           (m_lastErrorDescription = errorDescription) And _
                           (m_lastErrorModuleName = moduleName) And _
                           (m_lastErrorContext = context)
End Function

' Verificar si se llamÃ³ a LogInfo con parÃ¡metros especÃ­ficos
Public Function WasLogInfoCalledWith(message As String, Optional moduleName As String = "") As Boolean
    WasLogInfoCalledWith = (m_lastInfoMessage = message) And _
                          (m_lastInfoModuleName = moduleName)
End Function

' Verificar si se llamÃ³ a LogWarning con parÃ¡metros especÃ­ficos
Public Function WasLogWarningCalledWith(message As String, Optional moduleName As String = "") As Boolean
    WasLogWarningCalledWith = (m_lastWarningMessage = message) And _
                             (m_lastWarningModuleName = moduleName)
End Function

#End If
