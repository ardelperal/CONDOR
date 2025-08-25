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
Private m_IsValidTransition_WasCalled As Boolean

'==============================================================================
' EVENTOS DE CLASE
'==============================================================================

Private Sub Class_Initialize()
    Set m_Rules = New Collection
    m_IsValidTransition_WasCalled = False
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
' @param esValida: Indica si la transición es válida (True) o no (False)
Public Sub AddRule(ByVal tipo As String, ByVal origen As String, ByVal destino As String, Optional ByVal esValida As Boolean = True)
    Dim ruleKey As String
    ruleKey = tipo & "|" & origen & "|" & destino
    
    ' Verificar si la clave ya existe antes de añadir
    If Not RuleExists(ruleKey) Then
        m_Rules.Add esValida, ruleKey
    End If
End Sub

' Limpia todas las reglas configuradas
Public Sub ClearRules()
    Set m_Rules = New Collection
End Sub

' Obtiene el nÃºmero de reglas configuradas (para verificaciÃ³n en pruebas)
Public Function GetRuleCount() As Long
    GetRuleCount = m_Rules.Count
End Function

' Indica si el método IsValidTransition fue llamado
Public Property Get IsValidTransition_WasCalled() As Boolean
    IsValidTransition_WasCalled = m_IsValidTransition_WasCalled
End Property

'==============================================================================
' MÉTODOS PRIVADOS
'==============================================================================

' Verifica si una clave existe en la colección de reglas
' @param ruleKey: Clave a verificar
' @return: True si la clave existe, False en caso contrario
Private Function RuleExists(ByVal ruleKey As String) As Boolean
    Dim testValue As Variant
    Dim errorOccurred As Boolean
    
    errorOccurred = False
    
    ' Intentar acceder al elemento usando la clave
    On Error GoTo ErrorHandler
    testValue = m_Rules(ruleKey)
    RuleExists = True
    Exit Function
    
ErrorHandler:
    RuleExists = False
    Resume Next
End Function

'==============================================================================
' IMPLEMENTACIÃ“N DE IWorkflowRepository
'==============================================================================

Private Function IWorkflowRepository_IsValidTransition(ByVal tipoSolicitud As String, ByVal estadoOrigen As String, ByVal estadoDestino As String) As Boolean
    ' Marcar que el método fue llamado
    m_IsValidTransition_WasCalled = True
    
    Dim ruleKey As String
    ruleKey = tipoSolicitud & "|" & estadoOrigen & "|" & estadoDestino
    
    ' Verificar si la regla existe y es válida
    If RuleExists(ruleKey) Then
        IWorkflowRepository_IsValidTransition = m_Rules(ruleKey)
    Else
        IWorkflowRepository_IsValidTransition = False
    End If
End Function

#End If
