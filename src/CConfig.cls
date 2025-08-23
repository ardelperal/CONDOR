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
' Implementación del servicio de configuración. Versión 2.0.
' Autónoma y sin dependencias externas.

Implements IConfig

' Variables privadas
Private m_Settings As Collection
Private m_IsInitialized As Boolean

' Constructor de clase
Private Sub Class_Initialize()
    Set m_Settings = New Collection
    m_IsInitialized = False
    Call LoadConfigurationFromDatabase
End Sub

' Implementación de IConfig.GetValue
Private Function IConfig_GetValue(ByVal key As String) As Variant
    On Error GoTo ErrorHandler
    
    If Not m_IsInitialized Then
        Call LoadConfigurationFromDatabase
    End If
    
    IConfig_GetValue = m_Settings(UCase(key))
    Exit Function
    
ErrorHandler:
    IConfig_GetValue = Null
End Function

' Implementación de IConfig.HasKey
Private Function IConfig_HasKey(ByVal key As String) As Boolean
    On Error GoTo ErrorHandler
    
    If Not m_IsInitialized Then
        Call LoadConfigurationFromDatabase
    End If
    
    On Error Resume Next
    Dim testValue As Variant
    testValue = m_Settings(UCase(key))
    IConfig_HasKey = (Err.Number = 0)
    On Error GoTo ErrorHandler
    
    Exit Function
    
ErrorHandler:
    IConfig_HasKey = False
End Function

' Método privado para cargar configuración desde la base de datos
Private Sub LoadConfigurationFromDatabase()
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim backendPath As String
    
    ' Usar constante centralizada para la ruta del backend
    backendPath = Application.CurrentProject.Path & BACKEND_DB_PATH
    
    ' Conectar directamente a la base de datos del Backend
    Set db = DBEngine.OpenDatabase(backendPath, False, False, "MS Access;PWD=dpddpd")
    
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
    db.Close
    Set rs = Nothing
    Set db = Nothing
    
    ' Marcar como inicializado
    m_IsInitialized = True
    
    Exit Sub
    
ErrorHandler:
    ' Limpiar recursos en caso de error
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    If Not db Is Nothing Then
        db.Close
        Set db = Nothing
    End If
    m_IsInitialized = False
End Sub























