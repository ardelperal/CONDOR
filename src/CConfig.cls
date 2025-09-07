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

Private m_Settings As Object

Private Sub Class_Initialize()
    Set m_Settings = New Scripting.Dictionary
    m_Settings.CompareMode = TextCompare ' Case-insensitive keys
End Sub

' Método para cargar la configuración desde la BD. Lanza errores si falla.
Public Sub LoadConfiguration()
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim rsConfig As DAO.Recordset
    Dim backendPath As String
    Dim backendPassword As String
    
    On Error GoTo ErrorHandler
    
    ' 1. Leer TbLocalConfig del Frontend para determinar el entorno
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT [Entorno] FROM [TbLocalConfig]")
    If rs.EOF Then
        Err.Raise vbObjectError + 1001, "CConfig.LoadConfiguration", "No se encontró ningún registro en TbLocalConfig."
    End If
    Dim entorno As String
    entorno = rs.Fields("Entorno").value
    rs.Close
    
    ' 2. Determinar la ruta del Backend según el entorno
    ' Esta lógica asume que las rutas base están definidas en la tabla de configuración del backend
    ' Por ahora, para que funcione, deducimos la ruta base de la ruta del frontend.
    Dim basePath As String
    basePath = Left(db.Name, InStrRev(db.Name, "\")) & "..\.."
    
    If entorno = "LOCAL" Then
        backendPath = basePath & "\back\CONDOR_datos.accdb"
        backendPassword = "dpddpd" ' Contraseña estándar para entornos locales
    Else ' Asumimos "OFICINA" o cualquier otro valor
        ' Aquí iría la lógica para obtener la ruta de red, por ahora usamos la local
        backendPath = basePath & "\back\CONDOR_datos.accdb"
        backendPassword = "dpddpd"
    End If

    ' Guardar la ruta del backend en la configuración para que otros servicios la usen
    m_Settings.Add "DATA_PATH", backendPath
    m_Settings.Add "DATABASE_PASSWORD", backendPassword
    
    ' 3. Conectar al Backend y leer tbConfiguracion
    Set db = DBEngine.OpenDatabase(backendPath, False, True, ";PWD=" & backendPassword)
    Set rsConfig = db.OpenRecordset("SELECT clave, valor FROM tbConfiguracion")
    
    Do While Not rsConfig.EOF
        If Not m_Settings.Exists(rsConfig!clave) Then
            m_Settings.Add rsConfig!clave, rsConfig!valor
        Else
            m_Settings(rsConfig!clave) = rsConfig!valor
        End If
        rsConfig.MoveNext
    Loop
    
    rsConfig.Close
    db.Close
    
    Exit Sub
    
ErrorHandler:
    If Not rs Is Nothing Then rs.Close
    If Not rsConfig Is Nothing Then rsConfig.Close
    If Not db Is Nothing Then db.Close
    Set rs = Nothing
    Set rsConfig = Nothing
    Set db = Nothing
    ' Propagar el error hacia arriba para que sea manejado por el ErrorHandler global
    Err.Raise Err.Number, "CConfig.LoadConfiguration", Err.Description
End Sub

' Método público para testabilidad que permite cargar configuración desde un diccionario.
Public Sub LoadFromDictionary(ByVal settings As Object)
    Set m_Settings = settings
    
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

Private Function IConfig_GetCondorDataPath() As String
    IConfig_GetCondorDataPath = IConfig_GetValue("CONDOR_DATA_PATH")
End Function

Private Function IConfig_GetExpedientesDataPath() As String
    IConfig_GetExpedientesDataPath = IConfig_GetValue("EXPEDIENTES_DATA_PATH")
End Function

Private Function IConfig_GetCondorPassword() As String
    IConfig_GetCondorPassword = IConfig_GetValue("CONDOR_PASSWORD")
End Function

Private Function IConfig_GetExpedientesPassword() As String
    IConfig_GetExpedientesPassword = IConfig_GetValue("EXPEDIENTES_PASSWORD")
End Function

Private Function IConfig_GetCorreosPassword() As String
    IConfig_GetCorreosPassword = IConfig_GetValue("CORREOS_PASSWORD")
End Function

'==============================================================================
' MÉTODOS PÚBLICOS DE CONVENIENCIA (Lección 24)
'==============================================================================

Public Function GetValue(ByVal key As String) As Variant
    GetValue = IConfig_GetValue(key)
End Function

Public Function HasKey(ByVal key As String) As Boolean
    HasKey = IConfig_HasKey(key)
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

Public Function GetCondorDataPath() As String
    GetCondorDataPath = IConfig_GetCondorDataPath()
End Function

Public Function GetExpedientesDataPath() As String
    GetExpedientesDataPath = IConfig_GetExpedientesDataPath()
End Function

Public Function GetCondorPassword() As String
    GetCondorPassword = IConfig_GetCondorPassword()
End Function

Public Function GetExpedientesPassword() As String
    GetExpedientesPassword = IConfig_GetExpedientesPassword()
End Function

Public Function GetCorreosPassword() As String
    GetCorreosPassword = IConfig_GetCorreosPassword()
End Function

Public Sub SetSetting(ByVal key As String, ByVal value As Variant)
    Call IConfig_SetSetting(key, value)
End Sub

























