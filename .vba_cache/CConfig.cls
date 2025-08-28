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
Private m_isInitialized As Boolean

' Constructor de clase
Private Sub Class_Initialize()
    Set m_Settings = New Collection
    m_isInitialized = False
End Sub

' Método público para cargar configuración desde una Collection (para pruebas unitarias)
Public Sub LoadFromCollection(ByVal settings As Collection)
    Set m_Settings = settings
    m_isInitialized = True
End Sub

' ============================================================================
' MÉTODOS PÚBLICOS DE CONVENIENCIA (Lección 24)
' ============================================================================

Public Function GetValue(ByVal key As String) As Variant
    GetValue = IConfig_GetValue(key)
End Function

Public Function HasKey(ByVal key As String) As Boolean
    HasKey = IConfig_HasKey(key)
End Function

Public Function GetDataPath() As String
    GetDataPath = IConfig_GetDataPath()
End Function

Public Function GetDatabasePassword() As String
    GetDatabasePassword = IConfig_GetDatabasePassword()
End Function

Public Function GetAttachmentsPath() As String
    GetAttachmentsPath = IConfig_GetAttachmentsPath()
End Function

Public Function GetCorreosDBPath() As String
    GetCorreosDBPath = IConfig_GetCorreosDBPath()
End Function

Public Function GetUsuarioActual() As String
    GetUsuarioActual = IConfig_GetUsuarioActual()
End Function

Public Function GetCorreoAdministrador() As String
    GetCorreoAdministrador = IConfig_GetCorreoAdministrador()
End Function

Public Function GetIDAplicacionCondor() As String
    GetIDAplicacionCondor = IConfig_GetIDAplicacionCondor()
End Function

Public Sub SetSetting(ByVal key As String, ByVal value As Variant)
    IConfig_SetSetting key, value
End Sub

' Función auxiliar privada para obtener valores de la colección m_Settings
Private Function GetSettingValue(ByVal key As String) As Variant
    On Error GoTo ErrorHandler
    
    GetSettingValue = m_Settings(UCase(key))
    Exit Function
    
ErrorHandler:
    GetSettingValue = Null
End Function

' Implementación de IConfig.GetValue
Private Function IConfig_GetValue(ByVal key As String) As Variant
    IConfig_GetValue = GetSettingValue(key)
End Function

' Implementación de IConfig.HasKey
Private Function IConfig_HasKey(ByVal key As String) As Boolean
    On Error GoTo ErrorHandler
    Dim temp As Variant
    temp = m_Settings(UCase(key))  ' Intenta acceder al elemento
    IConfig_HasKey = True
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
    On Error GoTo ErrorHandler
    
    Dim rs As DAO.recordset
    Dim qdf As DAO.QueryDef
    
    ' Crear QueryDef con la consulta centralizada
    Set qdf = db.CreateQueryDef("", GET_ALL_CONFIGURATION)
    
    ' Ejecutar consulta
    Set rs = qdf.OpenRecordset(dbOpenSnapshot)
    
    ' Limpiar colección anterior
    Set m_Settings = New Collection
    
    ' Cargar todos los valores en la Collection privada
    Do While Not rs.EOF
        m_Settings.Add rs.Fields("Valor").value, UCase(rs.Fields("Clave").value)
        rs.MoveNext
    Loop
    
    ' Cerrar recursos
    rs.Close
    Set rs = Nothing
    Set qdf = Nothing
    
    ' Marcar como inicializado
    m_isInitialized = True
    
    Exit Sub
    
ErrorHandler:
    ' Limpiar recursos en caso de error
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    If Not qdf Is Nothing Then
        Set qdf = Nothing
    End If
    m_isInitialized = False
End Sub

' Implementación de IConfig.GetCorreosDBPath
Private Function IConfig_GetCorreosDBPath() As String
    IConfig_GetCorreosDBPath = GetSettingValue("CORREOS_DB_PATH")
End Function

' Implementación de IConfig.GetUsuarioActual
Private Function IConfig_GetUsuarioActual() As String
    IConfig_GetUsuarioActual = GetSettingValue("USUARIO_ACTUAL")
End Function

' Implementación de IConfig.GetCorreoAdministrador
Private Function IConfig_GetCorreoAdministrador() As String
    IConfig_GetCorreoAdministrador = GetSettingValue("CORREO_ADMINISTRADOR")
End Function

' Implementación de IConfig.GetIDAplicacionCondor
Private Function IConfig_GetIDAplicacionCondor() As String
    IConfig_GetIDAplicacionCondor = GetSettingValue("IDAplicacion_CONDOR")
End Function

' Implementación de IConfig.SetSetting
Private Sub IConfig_SetSetting(ByVal key As String, ByVal value As Variant)
    On Error Resume Next
    m_Settings.Remove UCase(key)
    m_Settings.Add value, UCase(key)
    On Error GoTo 0
End Sub























