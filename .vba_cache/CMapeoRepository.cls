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

Public Sub Initialize(ByVal configService As IConfig)
    Set m_Config = configService
End Sub

Public Function IMapeoRepository_GetMapeoPorTipo(ByVal tipoSolicitud As String) As DAO.Recordset
    Dim errorHandler As IErrorHandlerService
    On Error GoTo ErrorHandler
    
    ' Variables para la consulta parametrizada
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim rs As DAO.Recordset
    Dim sql As String
    
    Set db = DBEngine.OpenDatabase(m_Config.GetDataPath(), False, False, "MS Access;PWD=" & m_Config.GetDatabasePassword())
    
    ' Construir la consulta SQL parametrizada
    sql = "SELECT nombreCampoTabla, nombreCampoWord, valorAsociado FROM tbMapeoCampos WHERE nombrePlantilla = ?"
    
    ' Crear QueryDef con parámetro
    Set qdf = db.CreateQueryDef("", sql)
    qdf.Parameters(0) = tipoSolicitud
    
    ' Ejecutar consulta parametrizada
    Set rs = qdf.OpenRecordset()
    
    ' Limpiar recursos de consulta (mantener recordset abierto)
    Set qdf = Nothing
    Set db = Nothing
    
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
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    errorHandler.LogError Err.Number, Err.Description, "CMapeoRepository.GetMapeoPorTipo"
    Set IMapeoRepository_GetMapeoPorTipo = Nothing
End Function
