Option Compare Database
Option Explicit

Implements IOperationRepository

' Variables espía privadas (Lección 22)
Private m_SaveLog_WasCalled As Boolean
Private m_CallCount As Long
Private m_SaveLog_LastEntry As EOperationLog

' Propiedades públicas de solo lectura para verificación en tests
Public Property Get SaveLog_WasCalled() As Boolean
    SaveLog_WasCalled = m_SaveLog_WasCalled
End Property
Public Property Get CallCount() As Long
    CallCount = m_CallCount
End Property
Public Property Get SaveLog_LastEntry() As EOperationLog
    Set SaveLog_LastEntry = m_SaveLog_LastEntry
End Property

' Constructor
Private Sub Class_Initialize()
    Call Reset
End Sub

Private Sub IOperationRepository_SaveLog(ByVal logEntry As EOperationLog)
    ' Registrar la llamada para verificación en tests
    m_SaveLog_WasCalled = True
    m_CallCount = m_CallCount + 1
    Set m_SaveLog_LastEntry = logEntry
End Sub

' Resetea el estado del mock para asegurar el aislamiento entre tests
Public Sub Reset()
    m_SaveLog_WasCalled = False
    m_CallCount = 0
    Set m_SaveLog_LastEntry = Nothing
End Sub