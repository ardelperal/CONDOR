VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CWorkflowService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Implements IWorkflowService

Private m_Repo As IWorkflowRepository
Private m_Logger As IOperationLogger
Private m_ErrorHandler As IErrorHandlerService

Public Sub Initialize(ByVal repo As IWorkflowRepository, ByVal logger As IOperationLogger, ByVal errorHandler As IErrorHandlerService)
    Set m_Repo = repo
    Set m_Logger = logger
    Set m_ErrorHandler = errorHandler
End Sub

Private Function IWorkflowService_ValidateTransition(ByVal SolicitudID As Long, ByVal estadoOrigen As String, ByVal estadoDestino As String, ByVal tipoSolicitud As String, ByVal usuarioRol As String) As Boolean
    ' Lógica de negocio (placeholder)
    IWorkflowService_ValidateTransition = m_Repo.IsValidTransition(tipoSolicitud, estadoOrigen, estadoDestino)
End Function

Private Function IWorkflowService_GetNextStates(ByVal estadoActual As String, ByVal tipoSolicitud As String, ByVal usuarioRol As String) As Scripting.Dictionary
    ' Convertir estadoActual String a Long para el repositorio
    Dim idEstado As Long
    idEstado = CLng(estadoActual)
    Set IWorkflowService_GetNextStates = m_Repo.GetNextStates(idEstado, usuarioRol)
End Function


















