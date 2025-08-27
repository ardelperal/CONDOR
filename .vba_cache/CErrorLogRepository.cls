VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CErrorLogRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' ============================================================================
' Clase: CErrorLogRepository
' Descripción: Implementación del repositorio para la persistencia de errores.
' Implementa: IErrorLogRepository
' ============================================================================

Implements IErrorLogRepository

Private m_Config As IConfig

Public Sub Initialize(ByVal config As IConfig)
    Set m_Config = config
End Sub

Private Sub IErrorLogRepository_SaveErrorLog(ByVal errNumber As Long, ByVal errDescription As String, ByVal moduleName As String, ByVal isCritical As Boolean)
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    
    Set db = DBEngine.OpenDatabase(m_Config.GetDataPath, False, False, ";PWD=" & m_Config.GetDatabasePassword)
    Set rs = db.OpenRecordset("T_LogErrores", dbOpenDynaset)
    
    rs.AddNew
    rs!Fecha = Now()
    rs!NumeroError = errNumber
    rs!Descripcion = errDescription
    rs!Modulo = moduleName
    rs!EsCritico = isCritical
    rs.Update
    
    rs.Close
    db.Close
    Set rs = Nothing
    Set db = Nothing
    
    Exit Sub

ErrorHandler:
    ' If the repository fails, there is nowhere else to log it.
    ' The original error handler that called this will have to deal with it.
    If Not rs Is Nothing Then rs.Close
    If Not db Is Nothing Then db.Close
    Set rs = Nothing
    Set db = Nothing
    ' Re-raise the error to the calling service
    Err.Raise Err.Number, "CErrorLogRepository.SaveErrorLog", Err.Description
End Sub
