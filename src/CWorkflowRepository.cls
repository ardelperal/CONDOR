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
' PropÃ³sito: ImplementaciÃ³n real del repositorio de workflow que accede a la base de datos
' Autor: CONDOR-Expert
' Fecha: 2024
'==============================================================================

Option Compare Database
Option Explicit

Implements IWorkflowRepository

'==============================================================================
' VARIABLES PRIVADAS
'==============================================================================

' Dependencia para acceso a configuraciÃ³n
Private m_Config As IConfig

'==============================================================================
' MÃ‰TODOS PÃšBLICOS
'==============================================================================

'''
' Inicializa el repositorio con sus dependencias
' @param config: Servicio de configuraciÃ³n para obtener rutas de base de datos
'''
Public Sub Initialize(config As IConfig)
    Set m_Config = config
End Sub

'==============================================================================
' IMPLEMENTACIÃ“N DE INTERFAZ IWorkflowRepository
'==============================================================================

'''
' Valida si una transiciÃ³n de estado es permitida consultando la base de datos
' REFACTORIZADO: Usa consultas parametrizadas para prevenir inyecciÃ³n SQL
' @param tipoSolicitud: Tipo de solicitud (PC, CDCA, CDCASUB)
' @param estadoOrigen: Estado actual de la solicitud
' @param estadoDestino: Estado al que se quiere transicionar
' @return Boolean: True si la transiciÃ³n es vÃ¡lida, False en caso contrario
'''
Private Function IWorkflowRepository_IsValidTransition(ByVal tipoSolicitud As String, ByVal estadoOrigen As String, ByVal estadoDestino As String) As Boolean
    Dim errorHandler As IErrorHandlerService
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim rs As DAO.Recordset
    Dim validCount As Long
    
    ' Obtener ruta del backend desde configuraciÃ³n
    Dim backendPath As String
    backendPath = m_Config.GetDataPath()
    
    ' Abrir conexiÃ³n a la base de datos backend
    Set db = DBEngine.OpenDatabase(backendPath)
    
    ' Crear QueryDef con consulta parametrizada (SEGURA)
    Set qdf = db.CreateQueryDef("", _
        "SELECT COUNT(*) AS Valid " & _
        "FROM (TbTransiciones " & _
        "INNER JOIN TbEstados AS Origen ON TbTransiciones.EstadoOrigenID = Origen.ID) " & _
        "INNER JOIN TbEstados AS Destino ON TbTransiciones.EstadoDestinoID = Destino.ID " & _
        "WHERE TbTransiciones.TipoSolicitud = [pTipoSolicitud] " & _
        "AND Origen.CodigoEstado = [pEstadoOrigen] " & _
        "AND Destino.CodigoEstado = [pEstadoDestino] " & _
        "AND TbTransiciones.Activo = True;")
    
    ' Asignar valores a los parÃ¡metros de forma segura
    qdf.Parameters("pTipoSolicitud").Value = tipoSolicitud
    qdf.Parameters("pEstadoOrigen").Value = estadoOrigen
    qdf.Parameters("pEstadoDestino").Value = estadoDestino
    
    ' Ejecutar la consulta parametrizada
    Set rs = qdf.OpenRecordset(dbOpenSnapshot)
    
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
    Set qdf = Nothing
    db.Close
    Set db = Nothing
    
    Exit Function
    
ErrorHandler:
    ' En caso de error, asumir que la transiciÃ³n no es vÃ¡lida
    IWorkflowRepository_IsValidTransition = False
    
    ' Limpiar recursos si estÃ¡n abiertos
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    If Not qdf Is Nothing Then
        Set qdf = Nothing
    End If
    If Not db Is Nothing Then
        db.Close
        Set db = Nothing
    End If
    
    ' Registrar el error para debugging
    Debug.Print "Error en CWorkflowRepository.IsValidTransition: " & Err.Number & " - " & Err.Description
End Function