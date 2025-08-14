VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMockAuthService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' =====================================================
' CLASE: CMockAuthService
' PROPOSITO: Implementacion MOCK del servicio de autenticacion para pruebas
' AUTOR: CONDOR-Expert
' FECHA: 2025-01-14
' DESCRIPCION: Esta clase simula el comportamiento del servicio de autenticacion
'              sin conectarse a la base de datos real. Utiliza datos hardcodeados
'              para proporcionar respuestas predecibles en las pruebas.
' =====================================================

Implements IAuthService

' =====================================================
' IMPLEMENTACION: IAuthService_GetUserRole
' PROPOSITO: Obtener el rol de un usuario basado en su email (VERSION MOCK)
' PARAMETROS:
'   - userEmail: Email del usuario a consultar
' RETORNA: E_UserRole - Rol del usuario (hardcodeado para pruebas)
' NOTAS: Esta implementacion NO accede a ninguna base de datos.
'        Los roles se determinan mediante un Select Case con emails predefinidos.
' =====================================================
Private Function IAuthService_GetUserRole(ByVal userEmail As String) As E_UserRole
    ' Este es un servicio FALSO para pruebas. No toca la base de datos.
    ' Los datos estan hardcodeados para proporcionar resultados predecibles
    
    ' Normalizar el email (convertir a minusculas para comparacion)
    Dim normalizedEmail As String
    normalizedEmail = LCase(Trim(userEmail))
    
    ' Validar que el email no este vacio
    If Len(normalizedEmail) = 0 Then
        IAuthService_GetUserRole = Rol_Desconocido
        Exit Function
    End If
    
    ' Determinar el rol basado en emails predefinidos para pruebas
    Select Case normalizedEmail
        Case "admin.condor@example.com"
            IAuthService_GetUserRole = Rol_Admin
            
        Case "calidad.condor@example.com"
            IAuthService_GetUserRole = Rol_Calidad
            
        Case "tecnico.condor@example.com"
            IAuthService_GetUserRole = Rol_Tecnico
            
        Case "usuario.test@example.com"
            ' Email adicional para pruebas generales
            IAuthService_GetUserRole = Rol_Tecnico
            
        Case "supervisor.condor@example.com"
            ' Email adicional para pruebas de supervisor
            IAuthService_GetUserRole = Rol_Calidad
            
        Case Else
            ' Cualquier otro email retorna rol desconocido
            IAuthService_GetUserRole = Rol_Desconocido
    End Select
End Function

' =====================================================
' METODO PUBLICO: GetMockUserList
' PROPOSITO: Obtener la lista de usuarios mock disponibles para pruebas
' RETORNA: String - Lista formateada de usuarios mock
' NOTAS: Metodo auxiliar para documentacion y debugging de pruebas
' =====================================================
Public Function GetMockUserList() As String
    Dim userList As String
    
    userList = "USUARIOS MOCK DISPONIBLES PARA PRUEBAS:" & vbCrLf
    userList = userList & "========================================" & vbCrLf
    userList = userList & "admin.condor@example.com -> Rol_Admin" & vbCrLf
    userList = userList & "calidad.condor@example.com -> Rol_Calidad" & vbCrLf
    userList = userList & "tecnico.condor@example.com -> Rol_Tecnico" & vbCrLf
    userList = userList & "usuario.test@example.com -> Rol_Tecnico" & vbCrLf
    userList = userList & "supervisor.condor@example.com -> Rol_Calidad" & vbCrLf
    userList = userList & "[cualquier otro email] -> Rol_Desconocido" & vbCrLf
    userList = userList & "========================================" & vbCrLf
    
    GetMockUserList = userList
End Function

' =====================================================
' METODO PUBLICO: IsMockService
' PROPOSITO: Identificar que esta es una implementacion mock
' RETORNA: Boolean - Siempre True para esta clase
' NOTAS: Metodo auxiliar para identificacion en tiempo de ejecucion
' =====================================================
Public Function IsMockService() As Boolean
    IsMockService = True
End Function

