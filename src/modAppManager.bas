Attribute VB_Name = "modAppManager"
Option Compare Database
Option Explicit

' Definir constante de compilacion condicional para modo desarrollo
#Const DEV_MODE = True

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
' El flujo de arranque tipico en una subrutina principal App_Start() seria:
'
' Sub App_Start()
'     ' 1. Obtener el email del usuario actual
'     Dim userEmail As String
'     userEmail = GetCurrentUserEmail()
'
'     ' 2. Crear instancia del servicio de autenticacion
'     Dim authService As IAuthService
'     Set authService = New CAuthService
'
'     ' 3. Determinar el rol del usuario
'     Dim currentUserRole As E_UserRole
'     currentUserRole = authService.GetUserRole(userEmail)
'
'     ' 4. Guardar el rol en variable global para uso en toda la aplicacion
'     g_CurrentUserRole = currentUserRole
'
'     ' 5. Continuar con la inicializacion de la aplicacion segun el rol
'     Select Case g_CurrentUserRole
'         Case Rol_Admin
'             ' Inicializar funcionalidades de administrador
'         Case Rol_Calidad
'             ' Inicializar funcionalidades de calidad
'         Case Rol_Tecnico
' TODO: Implementar inicialización específica por rol cuando sea necesario

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
    Debug.Print modTestRunner.RunAllTests()
End Sub

' =====================================================
' FUNCION: OBTENER_RESULTADOS_PRUEBAS
' PROPOSITO: Obtener los resultados de todas las pruebas como string
' RETORNA: String - Resultados de las pruebas
' =====================================================
Public Function OBTENER_RESULTADOS_PRUEBAS() As String
    OBTENER_RESULTADOS_PRUEBAS = modTestRunner.RunAllTests()
End Function




