Option Compare Database
Option Explicit
' Definir constante de compilacion condicional para modo desarrollo
#Const DEV_MODE = 1

' =====================================================
' MODULO: modAppManager
' PROPOSITO: Gestion centralizada de la aplicacion y autenticacion
' AUTOR: CONDOR-Expert
' FECHA: 2025-01-14
' =====================================================

' Enum para definir los roles de usuario en el sistema CONDOR
Public Enum E_UserRole
    Rol_Desconocido = 0
    Rol_Tecnico = 1
    Rol_Calidad = 2
    Rol_Admin = 3
End Enum

' Variable global para almacenar el rol del usuario actual
Public g_CurrentUserRole As E_UserRole

' =====================================================
' FUNCION: GetCurrentUserEmail
' PROPOSITO: Obtener el email del usuario actual con capacidad de suplantacion en desarrollo
' RETORNA: String - Email del usuario
' =====================================================
Public Function GetCurrentUserEmail() As String
    #If DEV_MODE Then
        ' Descomentar la siguiente linea para suplantar a un usuario durante las pruebas
        ' GetCurrentUserEmail = "correo.de.prueba.calidad@example.com"
        If GetCurrentUserEmail = "" Then GetCurrentUserEmail = VBA.Command
    #Else
        ' En produccion, siempre se usa VBA.Command
        GetCurrentUserEmail = VBA.Command
    #End If
End Function

' =====================================================
' FLUJO DE ARRANQUE DE LA APLICACION
' =====================================================
Public Sub App_Start()
    On Error GoTo ErrorHandler

    ' 0. Obtener instancia del servicio de errores usando el factory
    Dim errorHandler As IErrorHandlerService
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService()

    ' 1. Obtener el email del usuario actual
    Dim UserEmail As String
    UserEmail = GetCurrentUserEmail()

    If UserEmail = "" Then
        Call errorHandler.LogError(vbObjectError + 514, "No se pudo obtener el email del usuario para la autenticaciÃ³n.", "modAppManager.App_Start", "CRITICAL_STARTUP_FAILURE")
        Exit Sub
    End If

    ' 2. Crear instancia del servicio de autenticacion usando el factory
    Dim authService As IAuthService
    Set authService = modAuthFactory.CreateAuthService()

    ' 3. Determinar el rol del usuario y guardarlo en la variable global
    g_CurrentUserRole = authService.GetUserRole(UserEmail)

    ' 4. Obtener instancia de configuraciÃ³n y verificar que se haya cargado correctamente
    Dim config As IConfig
    Set config = modConfig.CreateConfigService()
    
    If Not config.GetValue("IsInitialized") Then
        ' El error especÃ­fico ya ha sido logueado por CConfig.ValidateConfiguration
        ' Simplemente detenemos la ejecuciÃ³n para prevenir que la app se abra en un estado invÃ¡lido.
        Exit Sub
    End If

    ' 5. Continuar con la inicializacion de la aplicacion segun el rol
    Select Case g_CurrentUserRole
        Case Rol_Admin
            Debug.Print "Usuario autenticado como Administrador."
            ' TODO: Implementar inicializaciÃ³n especÃ­fica para Admin
        Case Rol_Calidad
            Debug.Print "Usuario autenticado como Calidad."
            ' TODO: Implementar inicializaciÃ³n especÃ­fica para Calidad
        Case Rol_Tecnico
            Debug.Print "Usuario autenticado como TÃ©cnico."
            ' TODO: Implementar inicializaciÃ³n especÃ­fica para TÃ©cnico
        Case Rol_Desconocido
            Call errorHandler.LogError(vbObjectError + 515, "El usuario '" & UserEmail & "' no tiene un rol definido en el sistema.", "modAppManager.App_Start", "CRITICAL_AUTH_FAILURE")
    End Select
    
    Exit Sub
    
ErrorHandler:
    Call errorHandler.LogError(Err.Number, Err.Description, "modAppManager.App_Start", "CRITICAL_STARTUP_ERROR")
End Sub

' =====================================================
' FUNCION: Ping
' PROPOSITO: Funcion de diagnostico para smoke test
' RETORNA: String - "Pong"
' =====================================================
Public Function Ping() As String
    Ping = "Pong"
End Function

' =====================================================
' SUBRUTINA: EJECUTAR_TODAS_LAS_PRUEBAS
' PROPOSITO: Punto de entrada manual para ejecutar todas las pruebas
' USO: Ejecutar desde la ventana de Macros (Alt+F8) y revisar resultados en Ventana Inmediato (Ctrl+G)
' =====================================================
Public Sub EJECUTAR_TODAS_LAS_PRUEBAS()
    Call modTestRunner.EjecutarTodasLasPruebas
End Sub











