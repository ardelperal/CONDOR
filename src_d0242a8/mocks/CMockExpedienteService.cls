VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMockExpedienteService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

#If DEV_MODE Then

Implements IExpedienteService

' ============================================================================
' Clase: CMockExpedienteService
' Descripci?n: Implementaci?n Mock del servicio de expedientes para testing
'              NO se conecta a ninguna base de datos, devuelve datos hardcodeados
' Autor: CONDOR-Expert
' Fecha: Diciembre 2024
' ============================================================================

' Implementaci?n Mock de la interfaz IExpedienteService
' Devuelve datos predecibles y hardcodeados para pruebas unitarias
Private Function IExpedienteService_GetExpediente(ByVal IDExpediente As Long) As T_Expediente
    Dim resultado As T_Expediente
    
    ' Datos hardcodeados predecibles para testing
    Select Case IDExpediente
        Case 1001
            ' Expediente de prueba principal
            resultado.IDExpediente = 1001
            resultado.Nemotecnico = "TEST-001"
            resultado.Titulo = "Expediente de Prueba Principal"
            resultado.ResponsableCalidad = "Juan Perez"
            resultado.ResponsableTecnico = "Maria Garcia"
            resultado.Pecal = "SI"
            
        Case 1002
            ' Expediente de prueba secundario
            resultado.IDExpediente = 1002
            resultado.Nemotecnico = "TEST-002"
            resultado.Titulo = "Expediente de Prueba Secundario"
            resultado.ResponsableCalidad = "Ana Lopez"
            resultado.ResponsableTecnico = "Carlos Ruiz"
            resultado.Pecal = "NO"
            
        Case 1003
            ' Expediente con campos vac?os para probar manejo de nulos
            resultado.IDExpediente = 1003
            resultado.Nemotecnico = "TEST-003"
            resultado.Titulo = "Expediente con Campos Vacios"
            resultado.ResponsableCalidad = ""
            resultado.ResponsableTecnico = ""
            resultado.Pecal = ""
            
        Case Else
            ' Para cualquier otro ID, devolver expediente vac?o (simula "no encontrado")
            resultado.IDExpediente = 0
            resultado.Nemotecnico = ""
            resultado.Titulo = ""
            resultado.ResponsableCalidad = ""
            resultado.ResponsableTecnico = ""
            resultado.Pecal = ""
    End Select
    
    IExpedienteService_GetExpediente = resultado
End Function

Private Function IExpedienteService_GetExpedienteById(ByVal IDExpediente As Long) As T_Expediente
    ' Wrapper que llama al m├®todo principal GetExpediente
    IExpedienteService_GetExpedienteById = IExpedienteService_GetExpediente(IDExpediente)
End Function

Private Function IExpedienteService_CreateExpediente(ByVal numeroExpediente As String, ByVal descripcion As String, ByVal idUsuario As Long) As Long
    ' Mock implementation - simula crear expediente
    IExpedienteService_CreateExpediente = 999
End Function

Private Function IExpedienteService_UpdateExpediente(ByVal ID As Long, ByVal descripcion As String, ByVal estado As String) As Boolean
    ' Mock implementation - simula actualizar expediente
    IExpedienteService_UpdateExpediente = True
End Function

Private Function IExpedienteService_DeleteExpediente(ByVal ID As Long) As Boolean
    ' Mock implementation - simula eliminar expediente
    IExpedienteService_DeleteExpediente = True
End Function

Private Function IExpedienteService_SearchExpedientes(ByVal criterio As String) As Collection
    ' Mock implementation - devuelve colecci├│n ficticia
    Set IExpedienteService_SearchExpedientes = New Collection
End Function

Private Function IExpedienteService_GetExpedientesByUser(ByVal idUsuario As Long) As Collection
    ' Mock implementation - devuelve colecci├│n ficticia
    Set IExpedienteService_GetExpedientesByUser = New Collection
End Function

Private Function IExpedienteService_ValidateExpediente(expediente As T_Expediente) As Boolean
    ' Mock implementation - simula validaci├│n
    IExpedienteService_ValidateExpediente = True
End Function

#End If




