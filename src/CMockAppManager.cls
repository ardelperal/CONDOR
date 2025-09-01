VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMockAppManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Implements IAppManager

' =====================================================
' CLASE: CMockAppManager
' PROPOSITO: Mock del gestor de aplicación para pruebas unitarias
' AUTOR: CONDOR-Expert
' FECHA: 2025-01-14
' =====================================================

' Variables públicas para seguimiento de llamadas
Public Initialize_WasCalled As Boolean
Public Initialize_CallCount As Long
Public StartApplication_WasCalled As Boolean
Public StartApplication_CallCount As Long
Public StartApplication_ReturnValue As Boolean
Public StartApplication_LastUserEmail As String
Public GetCurrentUserRole_WasCalled As Boolean
Public GetCurrentUserRole_CallCount As Long
Public GetCurrentUserRole_ReturnValue As UserRole
Public IsApplicationInitialized_WasCalled As Boolean
Public IsApplicationInitialized_CallCount As Long
Public IsApplicationInitialized_ReturnValue As Boolean
Public Ping_WasCalled As Boolean
Public Ping_CallCount As Long
Public Ping_ReturnValue As String

' Variables para almacenar parámetros de Initialize
Public LastAuthService As IAuthService
Public LastConfig As IConfig
Public LastErrorHandler As IErrorHandlerService

' ============================================================================
' CONSTRUCTOR
' ============================================================================
Private Sub Class_Initialize()
    Reset
End Sub

' ============================================================================
' MÉTODO PÚBLICO RESET
' ============================================================================
Public Sub Reset()
    Initialize_WasCalled = False
    Initialize_CallCount = 0
    StartApplication_WasCalled = False
    StartApplication_CallCount = 0
    StartApplication_ReturnValue = True
    StartApplication_LastUserEmail = ""
    GetCurrentUserRole_WasCalled = False
    GetCurrentUserRole_CallCount = 0
    GetCurrentUserRole_ReturnValue = RolDesconocido
    IsApplicationInitialized_WasCalled = False
    IsApplicationInitialized_CallCount = 0
    IsApplicationInitialized_ReturnValue = True
    Ping_WasCalled = False
    Ping_CallCount = 0
    Ping_ReturnValue = "Pong"
    
    Set LastAuthService = Nothing
    Set LastConfig = Nothing
    Set LastErrorHandler = Nothing
End Sub

' ============================================================================
' IMPLEMENTACIÓN DE LA INTERFAZ IAppManager
' ============================================================================

Private Sub IAppManager_Initialize(ByVal authService As IAuthService, ByVal config As IConfig, ByVal errorHandler As IErrorHandlerService)
    Initialize_WasCalled = True
    Initialize_CallCount = Initialize_CallCount + 1
    Set LastAuthService = authService
    Set LastConfig = config
    Set LastErrorHandler = errorHandler
End Sub

Private Function IAppManager_StartApplication(Optional ByVal userEmail As String = "") As Boolean
    StartApplication_WasCalled = True
    StartApplication_CallCount = StartApplication_CallCount + 1
    StartApplication_LastUserEmail = userEmail
    IAppManager_StartApplication = StartApplication_ReturnValue
End Function

Private Function IAppManager_GetCurrentUserRole() As UserRole
    GetCurrentUserRole_WasCalled = True
    GetCurrentUserRole_CallCount = GetCurrentUserRole_CallCount + 1
    IAppManager_GetCurrentUserRole = GetCurrentUserRole_ReturnValue
End Function

Private Function IAppManager_IsApplicationInitialized() As Boolean
    IsApplicationInitialized_WasCalled = True
    IsApplicationInitialized_CallCount = IsApplicationInitialized_CallCount + 1
    IAppManager_IsApplicationInitialized = IsApplicationInitialized_ReturnValue
End Function

Private Function IAppManager_Ping() As String
    Ping_WasCalled = True
    Ping_CallCount = Ping_CallCount + 1
    IAppManager_Ping = Ping_ReturnValue
End Function

' ============================================================================
' MÉTODOS PÚBLICOS DE CONVENIENCIA - ESTÁNDAR DE ORO PARA MOCKS
' ============================================================================

Public Sub Initialize(ByVal authService As IAuthService, ByVal config As IConfig, ByVal errorHandler As IErrorHandlerService)
    IAppManager_Initialize authService, config, errorHandler
End Sub

Public Function StartApplication(Optional ByVal userEmail As String = "") As Boolean
    StartApplication = IAppManager_StartApplication(userEmail)
End Function

Public Function GetCurrentUserRole() As UserRole
    GetCurrentUserRole = IAppManager_GetCurrentUserRole()
End Function

Public Function IsApplicationInitialized() As Boolean
    IsApplicationInitialized = IAppManager_IsApplicationInitialized()
End Function

Public Function Ping() As String
    Ping = IAppManager_Ping()
End Function