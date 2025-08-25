VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
' ImplementaciÃ³n del servicio de configuraciÃ³n. VersiÃ³n 2.0.
' AutÃ³noma y sin dependencias externas.

Implements IConfig

' Variables privadas
Private m_Settings As Collection
Private m_IsInitialized As Boolean

' Constructor de clase
Private Sub Class_Initialize()
    Set m_Settings = New Collection
    m_IsInitialized = False
End Sub

' Método público para cargar configuración desde una Collection (para pruebas unitarias)
Public Sub LoadFromCollection(ByVal settings As Collection)
    Set m_Settings = settings
    m_IsInitialized = True
End Sub

' ImplementaciÃ³n de IConfig.GetValue
Private Function IConfig_GetValue(ByVal key As String) As Variant
    Dim errorHandler As IErrorHandlerService
    On Error GoTo ErrorHandler
    
    IConfig_GetValue = m_Settings(UCase(key))
    Exit Function
    
ErrorHandler:
    IConfig_GetValue = Null
End Function

' ImplementaciÃ³n de IConfig.HasKey
Private Function IConfig_HasKey(ByVal key As String) As Boolean
    Dim errorHandler As IErrorHandlerService
    On Error GoTo ErrorHandler
    
    On Error Resume Next
    Dim testValue As Variant
    testValue = m_Settings(UCase(key))
    IConfig_HasKey = (Err.Number = 0)
    On Error GoTo ErrorHandler
    
    Exit Function
    
ErrorHandler:
    IConfig_HasKey = False
End Function



' Implementación de IConfig.GetDataPath
Private Function IConfig_GetDataPath() As String
    IConfig_GetDataPath = Me.IConfig_GetValue("BACKEND_DB_PATH")
End Function

' Implementación de IConfig.GetDatabasePassword
Private Function IConfig_GetDatabasePassword() As String
    IConfig_GetDatabasePassword = Me.IConfig_GetValue("DATABASE_PASSWORD")
End Function

' Método público para cargar configuración desde una base de datos específica
Public Sub Load(ByVal db As DAO.Database)
    Dim errorHandler As IErrorHandlerService
    On Error GoTo ErrorHandler
    
    Dim rs As DAO.Recordset
    
    ' Ejecutar consulta SELECT Clave, Valor FROM TbConfiguracion
    Set rs = db.OpenRecordset("SELECT Clave, Valor FROM TbConfiguracion", dbOpenSnapshot)
    
    ' Limpiar colección anterior
    Set m_Settings = New Collection
    
    ' Cargar todos los valores en la Collection privada
    Do While Not rs.EOF
        m_Settings.Add rs.Fields("Valor").Value, UCase(rs.Fields("Clave").Value)
        rs.MoveNext
    Loop
    
    ' Cerrar recursos
    rs.Close
    Set rs = Nothing
    
    ' Marcar como inicializado
    m_IsInitialized = True
    
    Exit Sub
    
ErrorHandler:
    ' Limpiar recursos en caso de error
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    m_IsInitialized = False
End Sub























