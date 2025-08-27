VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMockErrorLogRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' ============================================================================
' Clase: CMockErrorLogRepository
' Descripción: Mock del repositorio de errores para pruebas unitarias.
' Implementa: IErrorLogRepository
' ============================================================================

Implements IErrorLogRepository

Public SaveErrorLog_WasCalled As Boolean
Public SaveErrorLog_CallCount As Long
Public SaveErrorLog_LastErrNumber As Long
Public SaveErrorLog_LastErrDescription As String
Public SaveErrorLog_LastModuleName As String
Public SaveErrorLog_LastIsCritical As Boolean

Private Sub Class_Initialize()
    Call Reset
End Sub

Public Sub Reset()
    SaveErrorLog_WasCalled = False
    SaveErrorLog_CallCount = 0
    SaveErrorLog_LastErrNumber = 0
    SaveErrorLog_LastErrDescription = ""
    SaveErrorLog_LastModuleName = ""
    SaveErrorLog_LastIsCritical = False
End Sub

Private Sub IErrorLogRepository_SaveErrorLog(ByVal errNumber As Long, ByVal errDescription As String, ByVal moduleName As String, ByVal isCritical As Boolean)
    SaveErrorLog_WasCalled = True
    SaveErrorLog_CallCount = SaveErrorLog_CallCount + 1
    SaveErrorLog_LastErrNumber = errNumber
    SaveErrorLog_LastErrDescription = errDescription
    SaveErrorLog_LastModuleName = moduleName
    SaveErrorLog_LastIsCritical = isCritical
End Sub
