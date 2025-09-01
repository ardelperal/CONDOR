Option Compare Database
Option Explicit
Implements IConfig

Private m_Settings As Scripting.Dictionary

Private Sub Class_Initialize(): Call Reset: End Sub
Public Sub Reset(): Set m_Settings = New Scripting.Dictionary: m_Settings.CompareMode = TextCompare: End Sub

' --- Implementación COMPLETA de la Interfaz IConfig ---

Private Function IConfig_GetValue(ByVal key As String) As Variant
    If m_Settings.Exists(key) Then
        IConfig_GetValue = m_Settings(key)
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

' --- MÉTODOS PÚBLICOS DE CONVENIENCIA (Lección 24) ---

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

Public Function GetLanzaderaDataPath() As String
    GetLanzaderaDataPath = IConfig_GetLanzaderaDataPath()
End Function

Public Function GetLanzaderaPassword() As String
    GetLanzaderaPassword = IConfig_GetLanzaderaPassword()
End Function

Public Sub SetSetting(ByVal key As String, ByVal value As Variant)
    Call IConfig_SetSetting(key, value)
End Sub