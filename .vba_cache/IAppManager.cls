VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IAppManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


' =====================================================
' INTERFACE: IAppManager
' PROPOSITO: Interface para la gestión centralizada de la aplicación
' AUTOR: CONDOR-Expert
' FECHA: 2025-01-14
' =====================================================

' =====================================================
' METODO: Initialize
' PROPOSITO: Inicializar el gestor de aplicación con sus dependencias
' PARAMETROS:
'   - authService: Servicio de autenticación
'   - config: Servicio de configuración
'   - errorHandler: Manejador de errores
' =====================================================
Public Sub Initialize(ByVal authService As IAuthService, ByVal config As IConfig, ByVal errorHandler As IErrorHandlerService)
End Sub

' =====================================================
' METODO: StartApplication
' PROPOSITO: Iniciar la aplicación con autenticación y configuración
' PARAMETROS:
'   - userEmail: Email del usuario (opcional para testing)
' RETORNA: Boolean - True si la inicialización fue exitosa
' =====================================================
Public Function StartApplication(Optional ByVal userEmail As String = "") As Boolean
End Function

' =====================================================
' METODO: GetCurrentUserRole
' PROPOSITO: Obtener el rol del usuario actual autenticado
' RETORNA: UserRole - Rol del usuario actual
' =====================================================
Public Function GetCurrentUserRole() As UserRole
End Function

' =====================================================
' METODO: IsApplicationInitialized
' PROPOSITO: Verificar si la aplicación está correctamente inicializada
' RETORNA: Boolean - True si está inicializada
' =====================================================
Public Function IsApplicationInitialized() As Boolean
End Function

' =====================================================
' METODO: Ping
' PROPOSITO: Función de diagnóstico para smoke test
' RETORNA: String - "Pong"
' =====================================================
Public Function Ping() As String
End Function

