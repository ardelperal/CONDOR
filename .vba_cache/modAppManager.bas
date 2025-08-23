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

    ' 1. Obtener el email del usuario actual
    Dim UserEmail As String
    UserEmail = GetCurrentUserEmail()

    If UserEmail = "" Then
        Call modErrorHandler.LogCriticalError(vbObjectError + 514, "No se pudo obtener el email del usuario para la autenticación.", "modAppManager.App_Start")
        Exit Sub
    End If

    ' 2. Crear instancia del servicio de autenticacion usando el factory
    Dim authService As IAuthService
    Set authService = modAuthFactory.CreateAuthService()

    ' 3. Determinar el rol del usuario y guardarlo en la variable global
    g_CurrentUserRole = authService.GetUserRole(UserEmail)

    ' 4. Obtener instancia de configuración y verificar que se haya cargado correctamente
    Dim config As IConfig
    Set config = modConfig.CreateConfigService()
    
    If Not config.GetValue("IsInitialized") Then
        ' El error específico ya ha sido logueado por CConfig.ValidateConfiguration
        ' Simplemente detenemos la ejecución para prevenir que la app se abra en un estado inválido.
        Exit Sub
    End If

    ' 5. Continuar con la inicializacion de la aplicacion segun el rol
    Select Case g_CurrentUserRole
        Case Rol_Admin
            Debug.Print "Usuario autenticado como Administrador."
            ' TODO: Implementar inicialización específica para Admin
        Case Rol_Calidad
            Debug.Print "Usuario autenticado como Calidad."
            ' TODO: Implementar inicialización específica para Calidad
        Case Rol_Tecnico
            Debug.Print "Usuario autenticado como Técnico."
            ' TODO: Implementar inicialización específica para Técnico
        Case Rol_Desconocido
            Call modErrorHandler.LogCriticalError(vbObjectError + 515, "El usuario '" & UserEmail & "' no tiene un rol definido en el sistema.", "modAppManager.App_Start")
    End Select
    
    Exit Sub
    
ErrorHandler:
    Call modErrorHandler.LogCriticalError(Err.Number, Err.Description, "modAppManager.App_Start")
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











