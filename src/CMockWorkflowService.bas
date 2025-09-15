Option Compare Database
Option Explicit

Implements IWorkflowService

Private m_Validate_Result As Boolean
Private m_GetNext_Result As Object
Private m_Validate_WasCalled As Boolean

Private Sub Class_Initialize(): Call Reset: End Sub

Public Sub Reset()
    m_Validate_Result = False
    Set m_GetNext_Result = Nothing
    m_Validate_WasCalled = False
End Sub

Public Property Get ValidateTransition_WasCalled() As Boolean
    ValidateTransition_WasCalled = m_Validate_WasCalled
End Property

Public Sub ConfigureValidateTransition(ByVal value As Boolean)
    m_Validate_Result = value
End Sub

Public Sub ConfigureGetNextStates(ByVal value As Object)
    Set m_GetNext_Result = value
End Sub

Private Function IWorkflowService_ValidateTransition(ByVal SolicitudID As Long, ByVal estadoOrigen As String, ByVal estadoDestino As String, ByVal tipoSolicitud As String, ByVal usuarioRol As String) As Boolean
    m_Validate_WasCalled = True
    IWorkflowService_ValidateTransition = m_Validate_Result
End Function

Private Function IWorkflowService_GetNextStates(ByVal estadoActual As String, ByVal tipoSolicitud As String, ByVal usuarioRol As String) As Object
    Set IWorkflowService_GetNextStates = m_GetNext_Result
End Function

Private Function IWorkflowService_IsEstadoFinal(ByVal estadoActual As String) As Boolean
    ' Por defecto, un mock no considera ningún estado como final
    ' para no interferir con las pruebas, a menos que se configure explícitamente.
    IWorkflowService_IsEstadoFinal = False
End Function