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

Implements IConfig

Private m_Settings As Scripting.Dictionary

Private Sub Class_Initialize()
    Set m_Settings = New Scripting.Dictionary
    m_Settings.CompareMode = TextCompare ' Case-insensitive keys
End Sub

' Método para cargar la configuración desde la BD. Lanza errores si falla.
Public Sub LoadConfiguration()
    On Error GoTo errorHandler
    ' TODO: Lógica de lectura de BD
    m_Settings.Add "DATA_PATH", "C:\Path\To\Backend.accdb"
    m_Settings.Add "DATABASE_PASSWORD", "dpddpd"
    Exit Sub
errorHandler:
    Err.Raise vbObjectError + 1000, "CConfig.LoadConfiguration", "No se pudo cargar la configuración."
End Sub

' Método público para testabilidad que permite cargar configuración desde un diccionario.
Public Sub LoadFromDictionary(ByVal settings As Scripting.Dictionary)
    Set m_Settings = settings
    m_Settings.CompareMode = TextCompare
End Sub

Private Function IConfig_GetValue(ByVal key As String) As Variant
    If m_Settings.Exists(key) Then
        IConfig_GetValue = m_Settings(key)
    Else
        IConfig_GetValue = "" ' Devolver cadena vacía en lugar de Null
    End If
End Function

Private Sub IConfig_SetSetting(ByVal key As String, ByVal value As Variant)
    If m_Settings.Exists(key) Then
        m_Settings(key) = value
    Else
        m_Settings.Add key, value
    End If
End Sub

Private Function IConfig_HasKey(ByVal key As String) As Boolean
    IConfig_HasKey = m_Settings.Exists(key)
End Function

Private Function IConfig_GetDataPath() As String
    IConfig_GetDataPath = IConfig_GetValue("DATA_PATH")
End Function

Private Function IConfig_GetDatabasePassword() As String
    IConfig_GetDatabasePassword = IConfig_GetValue("DATABASE_PASSWORD")
End Function

Private Function IConfig_GetAttachmentsPath() As String
    IConfig_GetAttachmentsPath = IConfig_GetValue("ATTACHMENTS_PATH")
End Function

Private Function IConfig_GetCorreosDBPath() As String
    IConfig_GetCorreosDBPath = IConfig_GetValue("CORREOS_DB_PATH")
End Function

Private Function IConfig_GetUsuarioActual() As String
    IConfig_GetUsuarioActual = IConfig_GetValue("USUARIO_ACTUAL")
End Function

Private Function IConfig_GetCorreoAdministrador() As String
    IConfig_GetCorreoAdministrador = IConfig_GetValue("CORREO_ADMINISTRADOR")
End Function

Private Function IConfig_GetIDAplicacionCondor() As String
    IConfig_GetIDAplicacionCondor = IConfig_GetValue("ID_APLICACION_CONDOR")
End Function

Private Function IConfig_GetLanzaderaDataPath() As String
    IConfig_GetLanzaderaDataPath = IConfig_GetValue("LANZADERA_DATA_PATH")
End Function

Private Function IConfig_GetLanzaderaPassword() As String
    IConfig_GetLanzaderaPassword = IConfig_GetValue("LANZADERA_PASSWORD")
End Function

























