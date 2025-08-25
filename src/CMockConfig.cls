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

' Implementación de IConfig_GetValue
Private Function IConfig_GetValue(ByVal key As String) As Variant
    GetValue_WasCalled = True
    GetValue_CallCount = GetValue_CallCount + 1
    
    IConfig_GetValue = GetMockSettingValue(key)
End Function

' Implementación de IConfig_HasKey
Private Function IConfig_HasKey(ByVal key As String) As Boolean
    HasKey_WasCalled = True
    HasKey_CallCount = HasKey_CallCount + 1
    
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
    GetDataPath_WasCalled = True
    GetDataPath_CallCount = GetDataPath_CallCount + 1
    
    IConfig_GetDataPath = GetMockSettingValue("BACKEND_DB_PATH")
End Function

' Implementación de IConfig_GetDatabasePassword
Private Function IConfig_GetDatabasePassword() As String
    GetDatabasePassword_WasCalled = True
    GetDatabasePassword_CallCount = GetDatabasePassword_CallCount + 1
    
    IConfig_GetDatabasePassword = GetMockSettingValue("DATABASE_PASSWORD")
End Function

' Implementación de IConfig_GetAttachmentsPath
Private Function IConfig_GetAttachmentsPath() As String
    GetAttachmentsPath_WasCalled = True
    GetAttachmentsPath_CallCount = GetAttachmentsPath_CallCount + 1
    
    IConfig_GetAttachmentsPath = GetMockSettingValue("ATTACHMENTS_PATH")
End Function