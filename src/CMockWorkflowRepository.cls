VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMockWorkflowRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

#If DEV_MODE Then

'==============================================================================
' Clase: CMockWorkflowRepository
' PropÃ³sito: Mock del repositorio de workflow para pruebas unitarias
'            Implementa IWorkflowRepository permitiendo configurar reglas de transiciÃ³n
' Autor: CONDOR-Expert
' Fecha: 2025-01-21
'==============================================================================

Option Explicit

Implements IWorkflowRepository

'==============================================================================
' VARIABLES PRIVADAS
'==============================================================================

Private m_Rules As Collection

'==============================================================================
' EVENTOS DE CLASE
'==============================================================================

Private Sub Class_Initialize()
    Set m_Rules = New Collection
End Sub

Private Sub Class_Terminate()
    Set m_Rules = Nothing
End Sub

'==============================================================================
' MÃ‰TODOS PÃšBLICOS
'==============================================================================

' AÃ±ade una regla de transiciÃ³n vÃ¡lida al mock
' @param tipo: Tipo de solicitud
' @param origen: Estado origen
' @param destino: Estado destino
Public Sub AddRule(ByVal tipo As String, ByVal origen As String, ByVal destino As String)
    Dim ruleKey As String
    ruleKey = tipo & "|" & origen & "|" & destino
    
    On Error Resume Next
    m_Rules.Add True, ruleKey
    On Error GoTo 0
End Sub

' Limpia todas las reglas configuradas
Public Sub ClearRules()
    Set m_Rules = New Collection
End Sub

' Obtiene el nÃºmero de reglas configuradas (para verificaciÃ³n en pruebas)
Public Function GetRuleCount() As Long
    GetRuleCount = m_Rules.Count
End Function

'==============================================================================
' IMPLEMENTACIÃ“N DE IWorkflowRepository
'==============================================================================

Private Function IWorkflowRepository_IsValidTransition(ByVal tipoSolicitud As String, ByVal estadoOrigen As String, ByVal estadoDestino As String) As Boolean
    Dim ruleKey As String
    ruleKey = tipoSolicitud & "|" & estadoOrigen & "|" & estadoDestino
    
    On Error Resume Next
    Dim dummy As Boolean
    dummy = m_Rules(ruleKey)
    IWorkflowRepository_IsValidTransition = (Err.Number = 0)
    On Error GoTo 0
End Function

#End If
