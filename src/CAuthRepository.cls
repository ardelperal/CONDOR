VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CAuthRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



' =====================================================
' Clase: CAuthRepository
' Propósito: Implementación concreta del repositorio de autenticación
' Autor: CONDOR-Expert
' Fecha: 2025-01-15
' =====================================================

Implements IAuthRepository

' Dependencias inyectadas
Private m_configService As IConfig
Private m_ErrorHandler As IErrorHandlerService
Private m_isInitialized As Boolean

' Inicializa el repositorio con sus dependencias
' @param configService: Servicio de configuración para obtener ruta del Backend
' @param errorHandler: Servicio de manejo de errores
Public Sub Initialize(ByVal configService As IConfig, ByVal errorHandler As IErrorHandlerService)
    If configService Is Nothing Then
        Err.Raise 5, "CAuthRepository.Initialize", "configService no puede ser Nothing"
    End If
    
    If errorHandler Is Nothing Then
        Err.Raise 5, "CAuthRepository.Initialize", "errorHandler no puede ser Nothing"
    End If
    
    Set m_configService = configService
    Set m_ErrorHandler = errorHandler
    m_isInitialized = True
End Sub

' Implementación de IAuthRepository_GetUserAuthData
Private Function IAuthRepository_GetUserAuthData(ByVal userEmail As String) As EAuthData
    Dim authData As New EAuthData
    
    If Not m_isInitialized Then
        Err.Raise 5, "CAuthRepository.GetUserAuthData", "El repositorio no ha sido inicializado. Llame a Initialize() primero."
    End If
    
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim rs As DAO.Recordset
    
    On Error GoTo errorHandler
    
    ' Conectar a la base de datos Lanzadera
    Dim backendPath As String: backendPath = m_configService.GetLanzaderaDataPath()
    Dim backendPassword As String: backendPassword = m_configService.GetLanzaderaPassword()
    Set db = DBEngine.OpenDatabase(backendPath, dbFailOnError, False, ";PWD=" & backendPassword)
    
    ' Obtener el ID de la aplicación desde la configuración (CORRECCIÓN CLAVE)
    Dim idApp As String
    idApp = m_configService.GetIDAplicacionCondor()

    ' Eliminar la QueryDef temporal si existiera
    On Error Resume Next
    db.QueryDefs.Delete "tempAuthQuery"
    On Error GoTo errorHandler
    
    ' Crear QueryDef con la consulta desde modQueries
    Set qdf = db.CreateQueryDef("tempAuthQuery", GET_AUTH_DATA_BY_EMAIL)
    
    ' Asignar ambos parámetros a la consulta (CORRECCIÓN CLAVE)
    qdf.Parameters("pEmail") = userEmail
    qdf.Parameters("pIdAplicacion") = idApp
    
    Set rs = qdf.OpenRecordset(dbOpenSnapshot)
    
    ' Poblar el objeto AuthData basándose en los resultados
    If Not rs.EOF Then
        authData.UserExists = True
        If Not IsNull(rs("EsAdministrador").Value) And rs("EsAdministrador").Value = "Sí" Then
            authData.IsGlobalAdmin = True
        End If
        If Not IsNull(rs("EsUsuarioAdministrador").Value) And rs("EsUsuarioAdministrador").Value = True Then
            authData.IsAppAdmin = True
        End If
        ' ... (resto de los mapeos de roles)
    End If
    
Cleanup:
    On Error Resume Next
    Set IAuthRepository_GetUserAuthData = authData
    If Not rs Is Nothing Then rs.Close
    If Not db Is Nothing Then db.Close
    Set rs = Nothing: Set db = Nothing
    Exit Function
    
errorHandler:
    If Not m_ErrorHandler Is Nothing Then
        m_ErrorHandler.LogError Err.Number, Err.Description, "CAuthRepository.GetUserAuthData"
    End If
    GoTo Cleanup
End Function
' Método público de conveniencia que delega a la interfaz
Public Function GetUserAuthData(ByVal userEmail As String) As EAuthData
    Set GetUserAuthData = IAuthRepository_GetUserAuthData(userEmail)
End Function






