VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMockWorkflowRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


#If DEV_MODE Then

'==============================================================================
' Clase: CMockWorkflowRepository
' Propósito: Mock del repositorio de workflow para pruebas unitarias
'            Implementa IWorkflowRepository permitiendo configurar reglas de transición
' Autor: CONDOR-Expert
' Fecha: 2025-01-21
'==============================================================================


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
' MÉTODOS PÚBLICOS
'==============================================================================

' Añade una regla de transición válida al mock
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

' Obtiene el número de reglas configuradas (para verificación en pruebas)
Public Function GetRuleCount() As Long
    GetRuleCount = m_Rules.count
End Function

' Indica si el método IsValidTransition fue llamado
Public Property Get IsValidTransition_WasCalled() As Boolean
    IsValidTransition_WasCalled = m_IsValidTransition_WasCalled
End Property

' Reinicia el estado del mock a sus valores por defecto
Public Sub Reset()
    Set m_Rules = New Collection
    m_IsValidTransition_WasCalled = False
End Sub

'==============================================================================
' MÉTODOS PÚBLICOS DE CONVENIENCIA (DELEGACIÓN A INTERFAZ)
'==============================================================================

Public Function IsValidTransition(ByVal tipoSolicitud As String, ByVal estadoOrigen As String, ByVal estadoDestino As String) As Boolean
    IsValidTransition = IWorkflowRepository_IsValidTransition(tipoSolicitud, estadoOrigen, estadoDestino)
End Function

Public Function GetAvailableStates(ByVal tipoSolicitud As String) As Collection
    Set GetAvailableStates = IWorkflowRepository_GetAvailableStates(tipoSolicitud)
End Function

Public Function GetNextStates(ByVal estadoActual As String, ByVal tipoSolicitud As String, ByVal usuarioRol As String) As Collection
    Set GetNextStates = IWorkflowRepository_GetNextStates(estadoActual, tipoSolicitud, usuarioRol)
End Function

Public Function GetInitialState(ByVal tipoSolicitud As String) As String
    GetInitialState = IWorkflowRepository_GetInitialState(tipoSolicitud)
End Function

Public Function IsStateFinal(ByVal estadoCodigo As String, ByVal tipoSolicitud As String) As Boolean
    IsStateFinal = IWorkflowRepository_IsStateFinal(estadoCodigo, tipoSolicitud)
End Function

Public Function RecordStateChange(ByVal SolicitudID As Long, ByVal estadoAnterior As String, ByVal estadoNuevo As String, ByVal usuario As String, ByVal comentarios As String) As Boolean
    RecordStateChange = IWorkflowRepository_RecordStateChange(SolicitudID, estadoAnterior, estadoNuevo, usuario, comentarios)
End Function

Public Function GetStateHistory(ByVal SolicitudID As Long) As Collection
    Set GetStateHistory = IWorkflowRepository_GetStateHistory(SolicitudID)
End Function

Public Function HasTransitionPermission(ByVal usuarioRol As String, ByVal estadoOrigen As String, ByVal estadoDestino As String, ByVal tipoSolicitud As String) As Boolean
    HasTransitionPermission = IWorkflowRepository_HasTransitionPermission(usuarioRol, estadoOrigen, estadoDestino, tipoSolicitud)
End Function

Public Function RequiresApproval(ByVal estadoOrigen As String, ByVal estadoDestino As String, ByVal tipoSolicitud As String) As Boolean
    RequiresApproval = IWorkflowRepository_RequiresApproval(estadoOrigen, estadoDestino, tipoSolicitud)
End Function

Public Function GetTransitionRequiredRole(ByVal estadoOrigen As String, ByVal estadoDestino As String, ByVal tipoSolicitud As String) As String
    GetTransitionRequiredRole = IWorkflowRepository_GetTransitionRequiredRole(estadoOrigen, estadoDestino, tipoSolicitud)
End Function

'==============================================================================
' MÉTODOS PRIVADOS
'==============================================================================

' Verifica si una clave existe en la colección de reglas
' @param ruleKey: Clave a verificar
' @return: True si la clave existe, False en caso contrario
Private Function RuleExists(ByVal ruleKey As String) As Boolean
    Dim key As Variant
    
    ' Iterar sobre todas las claves de la colección
    For Each key In m_Rules
        If CStr(key) = ruleKey Then
            RuleExists = True
            Exit Function
        End If
    Next key
    
    ' Si llegamos aquí, la clave no fue encontrada
    RuleExists = False
End Function

'==============================================================================
' IMPLEMENTACIÓN DE IWorkflowRepository
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

Private Function IWorkflowRepository_GetAvailableStates(ByVal tipoSolicitud As String) As Collection
    ' Mock: devolver una colección con estados de prueba
    Dim estados As New Collection
    estados.Add "BORRADOR", "BORRADOR"
    estados.Add "EN_REVISION", "EN_REVISION"
    estados.Add "APROBADO", "APROBADO"
    Set IWorkflowRepository_GetAvailableStates = estados
End Function

Private Function IWorkflowRepository_GetNextStates(ByVal estadoActual As String, ByVal tipoSolicitud As String, ByVal usuarioRol As String) As Collection
    ' Mock: devolver estados siguientes basados en el estado actual
    Dim estadosSiguientes As New Collection
    
    Select Case estadoActual
        Case "BORRADOR"
            estadosSiguientes.Add "EN_REVISION", "EN_REVISION"
        Case "EN_REVISION"
            estadosSiguientes.Add "APROBADO", "APROBADO"
            estadosSiguientes.Add "RECHAZADO", "RECHAZADO"
        Case Else
            ' No hay estados siguientes
    End Select
    
    Set IWorkflowRepository_GetNextStates = estadosSiguientes
End Function

Private Function IWorkflowRepository_GetInitialState(ByVal tipoSolicitud As String) As String
    ' Mock: siempre devolver BORRADOR como estado inicial
    IWorkflowRepository_GetInitialState = "BORRADOR"
End Function

Private Function IWorkflowRepository_IsStateFinal(ByVal estadoCodigo As String, ByVal tipoSolicitud As String) As Boolean
    ' Mock: considerar APROBADO y RECHAZADO como estados finales
    IWorkflowRepository_IsStateFinal = (estadoCodigo = "APROBADO" Or estadoCodigo = "RECHAZADO")
End Function

Private Function IWorkflowRepository_RecordStateChange(ByVal SolicitudID As Long, ByVal estadoAnterior As String, ByVal estadoNuevo As String, ByVal usuario As String, ByVal comentarios As String) As Boolean
    ' Mock: siempre devolver True (éxito)
    IWorkflowRepository_RecordStateChange = True
End Function

Private Function IWorkflowRepository_GetStateHistory(ByVal SolicitudID As Long) As Collection
    ' Mock: devolver un historial de prueba
    Dim historial As New Collection
    historial.Add "2024-01-01 10:00:00 - TEST_USER: BORRADOR → EN_REVISION (Enviado para revisión)"
    historial.Add "2024-01-01 11:00:00 - REVIEWER: EN_REVISION → APROBADO (Aprobado)"
    Set IWorkflowRepository_GetStateHistory = historial
End Function

Private Function IWorkflowRepository_HasTransitionPermission(ByVal usuarioRol As String, ByVal estadoOrigen As String, ByVal estadoDestino As String, ByVal tipoSolicitud As String) As Boolean
    ' Mock: administradores siempre tienen permisos, otros roles dependen de la transición
    If usuarioRol = "ADMINISTRADOR" Then
        IWorkflowRepository_HasTransitionPermission = True
    Else
        ' Simular algunas reglas básicas
        Select Case estadoOrigen & "->" & estadoDestino
            Case "BORRADOR->EN_REVISION"
                IWorkflowRepository_HasTransitionPermission = True
            Case "EN_REVISION->APROBADO", "EN_REVISION->RECHAZADO"
                IWorkflowRepository_HasTransitionPermission = (usuarioRol = "REVISOR")
            Case Else
                IWorkflowRepository_HasTransitionPermission = False
        End Select
    End If
End Function

Private Function IWorkflowRepository_RequiresApproval(ByVal estadoOrigen As String, ByVal estadoDestino As String, ByVal tipoSolicitud As String) As Boolean
    ' Mock: transiciones a APROBADO requieren aprobación
    IWorkflowRepository_RequiresApproval = (estadoDestino = "APROBADO")
End Function

Private Function IWorkflowRepository_GetTransitionRequiredRole(ByVal estadoOrigen As String, ByVal estadoDestino As String, ByVal tipoSolicitud As String) As String
    ' Mock: devolver roles requeridos para ciertas transiciones
    Select Case estadoOrigen & "->" & estadoDestino
        Case "EN_REVISION->APROBADO", "EN_REVISION->RECHAZADO"
            IWorkflowRepository_GetTransitionRequiredRole = "REVISOR"
        Case Else
            IWorkflowRepository_GetTransitionRequiredRole = ""
    End Select
End Function

#End If


