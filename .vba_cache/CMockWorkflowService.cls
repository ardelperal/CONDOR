VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMockWorkflowService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Implements IWorkflowService

Private m_Validate_Result As Boolean
Private m_GetNext_Result As Scripting.Dictionary
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

Public Sub ConfigureGetNextStates(ByVal value As Scripting.Dictionary)
    Set m_GetNext_Result = value
End Sub

Private Function IWorkflowService_ValidateTransition(ByVal SolicitudID As Long, ByVal estadoOrigen As String, ByVal estadoDestino As String, ByVal tipoSolicitud As String, ByVal usuarioRol As String) As Boolean
    m_Validate_WasCalled = True
    IWorkflowService_ValidateTransition = m_Validate_Result
End Function

Private Function IWorkflowService_GetNextStates(ByVal estadoActual As String, ByVal tipoSolicitud As String, ByVal usuarioRol As String) As Scripting.Dictionary
    Set IWorkflowService_GetNextStates = m_GetNext_Result
End Function