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

' Función auxiliar privada para obtener valores de la colección m_Settings
Private Function GetSettingValue(ByVal key As String) As Variant
    On Error GoTo ErrorHandler
    
    GetSettingValue = m_Settings(UCase(key))
    Exit Function
    
ErrorHandler:
    GetSettingValue = Null
End Function

' ImplementaciÃ³n de IConfig.GetValue
Private Function IConfig_GetValue(ByVal key As String) As Variant
    IConfig_GetValue = GetSettingValue(key)
End Function

' ImplementaciÃ³n de IConfig.HasKey
Private Function IConfig_HasKey(ByVal key As String) As Boolean
    Dim errorHandler As IErrorHandlerService
    On Error GoTo ErrorHandler
    
    ' Usar bucle For Each para una comprobación más limpia y explícita
    Dim settingKey As Variant
    For Each settingKey In m_Settings.Keys
        If UCase(settingKey) = UCase(key) Then
            IConfig_HasKey = True
            Exit Function
        End If
    Next settingKey
    
    IConfig_HasKey = False
    Exit Function
    
ErrorHandler:
    IConfig_HasKey = False
End Function



' Implementación de IConfig.GetDataPath
Private Function IConfig_GetDataPath() As String
    IConfig_GetDataPath = GetSettingValue("BACKEND_DB_PATH")
End Function

' Implementación de IConfig.GetDatabasePassword
Private Function IConfig_GetDatabasePassword() As String
    IConfig_GetDatabasePassword = GetSettingValue("DATABASE_PASSWORD")
End Function

' Implementación de IConfig.GetAttachmentsPath
Private Function IConfig_GetAttachmentsPath() As String
    IConfig_GetAttachmentsPath = GetSettingValue("ATTACHMENTS_PATH")
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























