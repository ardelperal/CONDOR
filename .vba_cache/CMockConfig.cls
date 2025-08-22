Option Compare Database
Option Explicit
' ============================================================================
' Clase: CMockConfig
' Descripción: Mock del servicio de configuración para pruebas unitarias.
' ============================================================================

Implements IConfig

Private m_plantillasPath As String
Private m_generatedDocsPath As String
Public IsTestMode As Boolean

' --- MÉTODOS PARA CONFIGURAR EL MOCK ---

Public Sub SetPlantillaPath(ByVal path As String)
    m_plantillasPath = path
End Sub

Public Sub SetGeneratedDocsPath(ByVal path As String)
    m_generatedDocsPath = path
End Sub

' --- IMPLEMENTACIÓN DE LA INTERFAZ IConfig ---

Private Function IConfig_GetPlantillasPath() As String
    IConfig_GetPlantillasPath = m_plantillasPath
End Function

Private Function IConfig_GetGeneratedDocsPath() As String
    IConfig_GetGeneratedDocsPath = m_generatedDocsPath
End Function

' --- MÉTODOS NO UTILIZADOS EN ESTAS PRUEBAS (IMPLEMENTACIONES VACÍAS) ---

Private Function IConfig_GetDatabasePassword() As String
End Function

Private Function IConfig_GetDatabasePath() As String
End Function

Private Function IConfig_GetDataPath() As String
End Function

Private Function IConfig_GetExpedientesDbPath() As String
End Function

Private Function IConfig_GetExpedientesPath() As String
End Function

Private Function IConfig_GetInstance() As IConfig
End Function

Private Function IConfig_GetLanzaderaDbPath() As String
End Function

Private Function IConfig_GetLogPath() As String
End Function

Private Function IConfig_GetValue(ByVal key As String) As Variant
End Function

Private Function IConfig_HasKey(ByVal key As String) As Boolean
End Function

Private Function IConfig_InitializeEnvironment() As Boolean
End Function

Private Sub IConfig_SetValue(ByVal key As String, ByVal value As Variant)
End Sub

Private Function IConfig_TestCConfig() As String
End Function

Private Function IConfig_ValidateConfiguration() As Boolean
End Function

Private Function IConfig_GetActiveEnvironment() As String
End Function

Private Function IConfig_IsTestMode() As Boolean
    IConfig_IsTestMode = IsTestMode
End Function
