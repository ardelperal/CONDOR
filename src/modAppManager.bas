Attribute VB_Name = "modAppManager"
Option Compare Database
Option Explicit

' =====================================================
' MÓDULO: modAppManager
' PROPÓSITO: Gestión centralizada de la aplicación y autenticación
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
' FUNCIÓN: GetCurrentUserEmail
' PROPÓSITO: Obtener el email del usuario actual con capacidad de suplantación en desarrollo
' RETORNA: String - Email del usuario
' =====================================================
Public Function GetCurrentUserEmail() As String
    #If DEV_MODE Then
        ' Descomentar la siguiente línea para suplantar a un usuario durante las pruebas
        ' GetCurrentUserEmail = "correo.de.prueba.calidad@example.com"
        If GetCurrentUserEmail = "" Then GetCurrentUserEmail = VBA.Command
    #Else
        ' En producción, siempre se usa VBA.Command
        GetCurrentUserEmail = VBA.Command
    #End If
End Function

' =====================================================
' FLUJO DE ARRANQUE DE LA APLICACIÓN
' =====================================================
' El flujo de arranque típico en una subrutina principal App_Start() sería:
'
' Sub App_Start()
'     ' 1. Obtener el email del usuario actual
'     Dim userEmail As String
'     userEmail = GetCurrentUserEmail()
'
'     ' 2. Crear instancia del servicio de autenticación
'     Dim authService As IAuthService
'     Set authService = New CAuthService
'
'     ' 3. Determinar el rol del usuario
'     Dim currentUserRole As E_UserRole
'     currentUserRole = authService.GetUserRole(userEmail)
'
'     ' 4. Guardar el rol en variable global para uso en toda la aplicación
'     g_CurrentUserRole = currentUserRole
'
'     ' 5. Continuar con la inicialización de la aplicación según el rol
'     Select Case g_CurrentUserRole
'         Case Rol_Admin
'             ' Inicializar funcionalidades de administrador
'         Case Rol_Calidad
'             ' Inicializar funcionalidades de calidad
'         Case Rol_Tecnico
'             ' Inicializar funcionalidades técnicas
'         Case Rol_Desconocido
'             ' Mostrar error y cerrar aplicación
'             MsgBox "Usuario sin permisos para acceder a CONDOR", vbCritical
'             Application.Quit
'     End Select
' End Sub
