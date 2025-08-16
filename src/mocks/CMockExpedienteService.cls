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
Private Function IExpedienteService_GetExpedienteById(ByVal IDExpediente As Long) As T_Expediente
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
    
    IExpedienteService_GetExpedienteById = resultado
End Function

#End If




