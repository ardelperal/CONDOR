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
    Dim rs As DAO.recordset
    Dim prm As EQueryParameter
    
    On Error GoTo ErrorHandler
    
    ' Conectar a la base de datos Lanzadera
    Dim backendPath As String
    backendPath = m_configService.GetLanzaderaDataPath()
    
    Dim backendPassword As String
    backendPassword = m_configService.GetLanzaderaPassword()
    
    Set db = DBEngine.OpenDatabase(backendPath, dbFailOnError, False, ";PWD=" & backendPassword)
    
    ' Crear QueryDef con consulta desde modQueries
    Set qdf = db.CreateQueryDef("", GET_AUTH_DATA_BY_EMAIL)
    
    ' Crear y asignar parámetros
    Set prm = New EQueryParameter
    prm.ParameterName = "pEmail"
    prm.ParameterValue = userEmail
    qdf.parameters(prm.ParameterName) = prm.ParameterValue
    
    Set rs = qdf.OpenRecordset(dbOpenSnapshot)
    
    ' Poblar el objeto AuthData basándose en los resultados
    If Not rs.EOF Then
        authData.UserExists = True
        
        ' Verificar si es administrador global
        If Not IsNull(rs("EsAdministrador")) And rs("EsAdministrador") = "Sí" Then
            authData.IsGlobalAdmin = True
        End If
        
        ' Verificar permisos específicos de aplicación
        If Not IsNull(rs("EsUsuarioAdministrador")) And rs("EsUsuarioAdministrador") = "Sí" Then
            authData.IsAppAdmin = True
        End If
        
        If Not IsNull(rs("EsUsuarioCalidad")) And rs("EsUsuarioCalidad") = "Sí" Then
            authData.IsCalidad = True
        End If
        
        If Not IsNull(rs("EsUsuarioTecnico")) And rs("EsUsuarioTecnico") = "Sí" Then
            authData.IsTecnico = True
        End If
    Else
        ' Usuario no encontrado - EAuthData ya inicializado con valores False
        authData.UserExists = False
    End If
    
    ' Cerrar recursos
    rs.Close
    Set rs = Nothing
    Set qdf = Nothing
    db.Close
    Set db = Nothing
    
    ' Logging removido - responsabilidad de la capa de servicio
    
    Set IAuthRepository_GetUserAuthData = authData
    Exit Function
    
ErrorHandler:
    If Not rs Is Nothing Then rs.Close
    If Not qdf Is Nothing Then Set qdf = Nothing
    If Not db Is Nothing Then db.Close
    
    If Not m_ErrorHandler Is Nothing Then
        m_ErrorHandler.LogError Err.Number, Err.Description, "CAuthRepository.GetUserAuthData"
    End If
    
    ' Devolver objeto con valores por defecto en caso de error
    Set IAuthRepository_GetUserAuthData = authData
End Function

' Método público de conveniencia que delega a la interfaz
Public Function GetUserAuthData(ByVal userEmail As String) As EAuthData
    Set GetUserAuthData = IAuthRepository_GetUserAuthData(userEmail)
End Function


