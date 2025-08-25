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
    
    Dim db As DAO.Database
    Dim sql As String
    
    Set db = DBEngine.OpenDatabase(m_Config.GetDataPath(), False, False, "MS Access;PWD=" & m_Config.GetDatabasePassword())
    sql = "SELECT nombreCampoTabla, nombreCampoWord, valorAsociado FROM tbMapeoCampos WHERE nombrePlantilla = '" & tipoSolicitud & "'"
    Set IMapeoRepository_GetMapeoPorTipo = db.OpenRecordset(sql)
    
    Exit Function
ErrorHandler:
    If Not db Is Nothing Then
        db.Close
        Set db = Nothing
    End If
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    errorHandler.LogError Err.Number, Err.Description, "CMapeoRepository.GetMapeoPorTipo"
    Set IMapeoRepository_GetMapeoPorTipo = Nothing
End Function
