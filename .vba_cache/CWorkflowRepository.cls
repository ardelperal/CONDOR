VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CWorkflowRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' Clase: CWorkflowRepository
' Propósito: Implementación real del repositorio de workflow que accede a la base de datos
' Autor: CONDOR-Expert
' Fecha: 2024
'==============================================================================

Option Compare Database
Option Explicit

Implements IWorkflowRepository

'==============================================================================
' VARIABLES PRIVADAS
'==============================================================================

' Dependencia para acceso a configuración
Private m_Config As IConfig

'==============================================================================
' MÉTODOS PÚBLICOS
'==============================================================================

'''
' Inicializa el repositorio con sus dependencias
' @param config: Servicio de configuración para obtener rutas de base de datos
'''
Public Sub Initialize(config As IConfig)
    Set m_Config = config
End Sub

'==============================================================================
' IMPLEMENTACIÓN DE INTERFAZ IWorkflowRepository
'==============================================================================

'''
' Valida si una transición de estado es permitida consultando la base de datos
' @param tipoSolicitud: Tipo de solicitud (PC, CDCA, CDCASUB)
' @param estadoOrigen: Estado actual de la solicitud
' @param estadoDestino: Estado al que se quiere transicionar
' @return Boolean: True si la transición es válida, False en caso contrario
'''
Private Function IWorkflowRepository_IsValidTransition(ByVal tipoSolicitud As String, ByVal estadoOrigen As String, ByVal estadoDestino As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim validCount As Long
    
    ' Construir la consulta SQL
    sql = "SELECT COUNT(*) AS Valid " & _
          "FROM (TbTransiciones " & _
          "INNER JOIN TbEstados AS Origen ON TbTransiciones.EstadoOrigenID = Origen.ID) " & _
          "INNER JOIN TbEstados AS Destino ON TbTransiciones.EstadoDestinoID = Destino.ID " & _
          "WHERE TbTransiciones.TipoSolicitud = '" & tipoSolicitud & "' " & _
          "AND Origen.CodigoEstado = '" & estadoOrigen & "' " & _
          "AND Destino.CodigoEstado = '" & estadoDestino & "' " & _
          "AND TbTransiciones.Activo = True;"
    
    ' Obtener ruta del backend desde configuración
    Dim backendPath As String
    backendPath = m_Config.GetValue("DATAPATH")
    
    ' Abrir conexión a la base de datos backend
    Set db = DBEngine.OpenDatabase(backendPath)
    
    ' Ejecutar la consulta
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    
    ' Verificar si hay resultados
    If Not rs.EOF Then
        validCount = rs.Fields("Valid").Value
        IWorkflowRepository_IsValidTransition = (validCount > 0)
    Else
        IWorkflowRepository_IsValidTransition = False
    End If
    
    ' Limpiar recursos
    rs.Close
    Set rs = Nothing
    db.Close
    Set db = Nothing
    
    Exit Function
    
ErrorHandler:
    ' En caso de error, asumir que la transición no es válida
    IWorkflowRepository_IsValidTransition = False
    
    ' Limpiar recursos si están abiertos
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
        Set rs = Nothing
    End If
    If Not db Is Nothing Then
        db.Close
        Set db = Nothing
    End If
    
    ' Registrar el error para debugging
    Debug.Print "Error en CWorkflowRepository.IsValidTransition: " & Err.Number & " - " & Err.Description
End Function