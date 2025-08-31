VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMockConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


' Mock para el servicio de configuración - Framework de Testing
' Implementa IConfig para permitir pruebas aisladas

Implements IConfig

' Variables privadas para almacenar configuraciones de prueba
Private m_mockSettings As Collection

' Propiedades de seguimiento de llamadas para auditoría de pruebas
Public GetValue_WasCalled As Boolean
Public GetValue_CallCount As Long
Public HasKey_WasCalled As Boolean
Public HasKey_CallCount As Long
Public GetDataPath_WasCalled As Boolean
Public GetDataPath_CallCount As Long
Public GetDatabasePassword_WasCalled As Boolean
Public GetDatabasePassword_CallCount As Long
Public GetAttachmentsPath_WasCalled As Boolean
Public GetAttachmentsPath_CallCount As Long
Public GetCorreosDBPath_WasCalled As Boolean
Public GetCorreosDBPath_CallCount As Long
Public GetUsuarioActual_WasCalled As Boolean
Public GetUsuarioActual_CallCount As Long
Public GetCorreoAdministrador_WasCalled As Boolean
Public GetCorreoAdministrador_CallCount As Long
Public GetIDAplicacionCondor_WasCalled As Boolean
Public GetIDAplicacionCondor_CallCount As Long
Public GetLanzaderaDataPath_WasCalled As Boolean
Public GetLanzaderaDataPath_CallCount As Long
Public GetLanzaderaPassword_WasCalled As Boolean
Public GetLanzaderaPassword_CallCount As Long

' Constructor de la clase
Private Sub Class_Initialize()
    Set m_mockSettings = New Collection
    Reset
End Sub

' Destructor de la clase
Private Sub Class_Terminate()
    Set m_mockSettings = Nothing
End Sub

' Método público para añadir configuraciones de prueba
Public Sub AddSetting(ByVal key As String, ByVal value As Variant)
    On Error Resume Next
    m_mockSettings.Remove key
    On Error GoTo 0
    m_mockSettings.Add value, key
End Sub

' Métodos de conveniencia para configurar valores específicos
Public Sub SetDataPath(ByVal Path As String)
    AddSetting "BACKEND_DB_PATH", Path
End Sub

Public Sub SetDatabasePassword(ByVal password As String)
    AddSetting "DATABASE_PASSWORD", password
End Sub

' Método público para limpiar todas las configuraciones y contadores
Public Sub Reset()
    Set m_mockSettings = New Collection
    GetValue_WasCalled = False
    GetValue_CallCount = 0
    HasKey_WasCalled = False
    HasKey_CallCount = 0
    GetDataPath_WasCalled = False
    GetDataPath_CallCount = 0
    GetDatabasePassword_WasCalled = False
    GetDatabasePassword_CallCount = 0
    GetAttachmentsPath_WasCalled = False
    GetAttachmentsPath_CallCount = 0
    GetCorreosDBPath_WasCalled = False
    GetCorreosDBPath_CallCount = 0
    GetUsuarioActual_WasCalled = False
    GetUsuarioActual_CallCount = 0
    GetCorreoAdministrador_WasCalled = False
    GetCorreoAdministrador_CallCount = 0
    GetIDAplicacionCondor_WasCalled = False
    GetIDAplicacionCondor_CallCount = 0
    GetLanzaderaDataPath_WasCalled = False
    GetLanzaderaDataPath_CallCount = 0
    GetLanzaderaPassword_WasCalled = False
    GetLanzaderaPassword_CallCount = 0
End Sub

'******************************************************************************
' IMPLEMENTACIÓN DE LA INTERFAZ IConfig
'******************************************************************************

' Función auxiliar privada para obtener valores de la colección m_mockSettings
Private Function GetMockSettingValue(ByVal key As String) As Variant
    On Error GoTo ErrorHandler
    GetMockSettingValue = m_mockSettings(key)
    Exit Function
    
ErrorHandler:
    ' Si la clave no existe, devolver Empty
    GetMockSettingValue = Empty
End Function

' Método privado para rastrear llamadas y eliminar código duplicado
Private Sub TrackCall(ByRef wasCalledFlag As Boolean, ByRef callCounter As Long)
    wasCalledFlag = True
    callCounter = callCounter + 1
End Sub

' Implementación de IConfig_GetValue
Private Function IConfig_GetValue(ByVal key As String) As Variant
    TrackCall GetValue_WasCalled, GetValue_CallCount
    
    IConfig_GetValue = GetMockSettingValue(key)
End Function

' Implementación de IConfig_HasKey
Private Function IConfig_HasKey(ByVal key As String) As Boolean
    TrackCall HasKey_WasCalled, HasKey_CallCount
    
    On Error GoTo ErrorHandler
    Dim temp As Variant
    temp = m_mockSettings(key)
    IConfig_HasKey = True
    Exit Function
    
ErrorHandler:
    IConfig_HasKey = False
End Function

' Implementación de IConfig_GetDataPath
Private Function IConfig_GetDataPath() As String
    TrackCall GetDataPath_WasCalled, GetDataPath_CallCount
    
    IConfig_GetDataPath = GetMockSettingValue("BACKEND_DB_PATH")
End Function

' Implementación de IConfig_GetDatabasePassword
Private Function IConfig_GetDatabasePassword() As String
    TrackCall GetDatabasePassword_WasCalled, GetDatabasePassword_CallCount
    
    IConfig_GetDatabasePassword = GetMockSettingValue("DATABASE_PASSWORD")
End Function

' Implementación de IConfig_GetAttachmentsPath
Private Function IConfig_GetAttachmentsPath() As String
    TrackCall GetAttachmentsPath_WasCalled, GetAttachmentsPath_CallCount
    
    IConfig_GetAttachmentsPath = GetMockSettingValue("ATTACHMENTS_PATH")
End Function

' Implementación de IConfig_GetCorreosDBPath
Private Function IConfig_GetCorreosDBPath() As String
    TrackCall GetCorreosDBPath_WasCalled, GetCorreosDBPath_CallCount
    
    IConfig_GetCorreosDBPath = GetMockSettingValue("CORREOS_DB_PATH")
End Function

' Implementación de IConfig_GetUsuarioActual
Private Function IConfig_GetUsuarioActual() As String
    TrackCall GetUsuarioActual_WasCalled, GetUsuarioActual_CallCount
    
    IConfig_GetUsuarioActual = GetMockSettingValue("USUARIO_ACTUAL")
End Function

' Implementación de IConfig_GetCorreoAdministrador
Private Function IConfig_GetCorreoAdministrador() As String
    TrackCall GetCorreoAdministrador_WasCalled, GetCorreoAdministrador_CallCount
    
    IConfig_GetCorreoAdministrador = GetMockSettingValue("CORREO_ADMINISTRADOR")
End Function

' Implementación de IConfig_GetIDAplicacionCondor
Private Function IConfig_GetIDAplicacionCondor() As String
    TrackCall GetIDAplicacionCondor_WasCalled, GetIDAplicacionCondor_CallCount
    
    IConfig_GetIDAplicacionCondor = GetMockSettingValue("IDAplicacion_CONDOR")
End Function

' Implementación de IConfig_GetLanzaderaDataPath
Private Function IConfig_GetLanzaderaDataPath() As String
    TrackCall GetLanzaderaDataPath_WasCalled, GetLanzaderaDataPath_CallCount
    
    IConfig_GetLanzaderaDataPath = GetMockSettingValue("LANZADERA_DB_PATH")
End Function

' Implementación de IConfig_GetLanzaderaPassword
Private Function IConfig_GetLanzaderaPassword() As String
    TrackCall GetLanzaderaPassword_WasCalled, GetLanzaderaPassword_CallCount
    
    IConfig_GetLanzaderaPassword = GetMockSettingValue("LANZADERA_DB_PASSWORD")
End Function

' Implementación de IConfig_SetSetting
Private Sub IConfig_SetSetting(ByVal key As String, ByVal value As Variant)
    AddSetting key, value
End Sub

'******************************************************************************
' MÉTODOS PÚBLICOS DE CONVENIENCIA - ESTÁNDAR DE ORO PARA MOCKS
'******************************************************************************

' Método público de conveniencia para GetValue
Public Function GetValue(ByVal key As String) As Variant
    GetValue = IConfig_GetValue(key)
End Function

' Método público de conveniencia para HasKey
Public Function HasKey(ByVal key As String) As Boolean
    HasKey = IConfig_HasKey(key)
End Function

' Método público de conveniencia para GetDataPath
Public Function GetDataPath() As String
    GetDataPath = IConfig_GetDataPath()
End Function

' Método público de conveniencia para GetDatabasePassword
Public Function GetDatabasePassword() As String
    GetDatabasePassword = IConfig_GetDatabasePassword()
End Function

' Método público de conveniencia para GetAttachmentsPath
Public Function GetAttachmentsPath() As String
    GetAttachmentsPath = IConfig_GetAttachmentsPath()
End Function

' Método público de conveniencia para GetCorreosDBPath
Public Function GetCorreosDBPath() As String
    GetCorreosDBPath = IConfig_GetCorreosDBPath()
End Function

' Método público de conveniencia para GetUsuarioActual
Public Function GetUsuarioActual() As String
    GetUsuarioActual = IConfig_GetUsuarioActual()
End Function

' Método público de conveniencia para GetCorreoAdministrador
Public Function GetCorreoAdministrador() As String
    GetCorreoAdministrador = IConfig_GetCorreoAdministrador()
End Function

' Método público de conveniencia para GetIDAplicacionCondor
Public Function GetIDAplicacionCondor() As String
    GetIDAplicacionCondor = IConfig_GetIDAplicacionCondor()
End Function

' Método público de conveniencia para GetLanzaderaDataPath
Public Function GetLanzaderaDataPath() As String
    GetLanzaderaDataPath = IConfig_GetLanzaderaDataPath()
End Function

' Método público de conveniencia para GetLanzaderaPassword
Public Function GetLanzaderaPassword() As String
    GetLanzaderaPassword = IConfig_GetLanzaderaPassword()
End Function

' Método público de conveniencia para SetSetting
Public Sub SetSetting(ByVal key As String, ByVal value As Variant)
    IConfig_SetSetting key, value
End Sub

