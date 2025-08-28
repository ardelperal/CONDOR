VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMapeoRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Implements IMapeoRepository

Private m_Config As IConfig
Private m_ErrorHandler As IErrorHandlerService

Public Sub Initialize(ByVal configService As IConfig, ByVal errorHandler As IErrorHandlerService)
    Set m_Config = configService
    Set m_ErrorHandler = errorHandler
End Sub

Public Function IMapeoRepository_GetMapeoPorTipo(ByVal params As Collection) As DAO.recordset
    On Error GoTo ErrorHandler
    
    ' Variables para la consulta parametrizada
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim rs As DAO.recordset
    Dim prm As QueryParameter
    
    Set db = DBEngine.OpenDatabase(m_Config.GetDataPath(), dbFailOnError, False, "MS Access;PWD=" & m_Config.GetDatabasePassword())
    
    ' Crear QueryDef con consulta desde modQueries
    Set qdf = db.CreateQueryDef("", GET_MAPEO_POR_TIPO)
    
    ' Asignar parámetros desde la colección
    For Each prm In params
        qdf.parameters(prm.ParameterName) = prm.ParameterValue
    Next prm
    
    ' Ejecutar consulta parametrizada
    Set rs = qdf.OpenRecordset()
    
    ' Limpiar recursos de consulta (mantener recordset abierto)
    Set qdf = Nothing
    
    ' Devolver el recordset
    Set IMapeoRepository_GetMapeoPorTipo = rs
    
    Exit Function
ErrorHandler:
    ' Limpiar recursos en caso de error
    If Not rs Is Nothing Then rs.Close
    If Not qdf Is Nothing Then Set qdf = Nothing
    If Not db Is Nothing Then
        db.Close
        Set db = Nothing
    End If
    m_ErrorHandler.LogError Err.Number, Err.Description, "CMapeoRepository.GetMapeoPorTipo"
    Set IMapeoRepository_GetMapeoPorTipo = Nothing
End Function

' ============================================================================
' MÉTODOS PÚBLICOS DE CONVENIENCIA
' ============================================================================
' Estos métodos exponen públicamente la funcionalidad de la interfaz
' para facilitar las pruebas y el uso directo de la clase concreta

' Método público de conveniencia para GetMapeoPorTipo
Public Function GetMapeoPorTipo(ByVal tipoSolicitud As String) As DAO.recordset
    Dim params As Collection
    Dim prm As QueryParameter
    
    ' Crear colección de parámetros
    Set params = New Collection
    
    Set prm = New QueryParameter
    prm.ParameterName = "pNombrePlantilla"
    prm.ParameterValue = tipoSolicitud
    params.Add prm
    
    Set GetMapeoPorTipo = IMapeoRepository_GetMapeoPorTipo(params)
End Function


