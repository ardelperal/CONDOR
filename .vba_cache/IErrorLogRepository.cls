VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IErrorLogRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

' ============================================================================
' Interfaz: IErrorLogRepository
' Descripción: Define el contrato para los repositorios que gestionan la
'              persistencia de los registros de errores.
' ============================================================================

Public Sub SaveErrorLog(ByVal errNumber As Long, ByVal errDescription As String, ByVal moduleName As String, ByVal isCritical As Boolean)
End Sub
