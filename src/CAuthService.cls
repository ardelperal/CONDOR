VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CAuthService"
Option Compare Database
Option Explicit

' =====================================================
' CLASE: CAuthService
' PROPÓSITO: Implementación concreta del servicio de autenticación
' IMPLEMENTA: IAuthService
' AUTOR: CONDOR-Expert
' FECHA: 2025-01-14
' =====================================================

Implements IAuthService

' =====================================================
' FUNCIÓN: IAuthService_GetUserRole
' PROPÓSITO: Determinar el rol de un usuario consultando la BD Lanzadera
' PARÁMETROS:
'   - userEmail: String - Email del usuario a verificar
' RETORNA: E_UserRole - Rol del usuario según los permisos en Lanzadera
' =====================================================
Private Function IAuthService_GetUserRole(ByVal userEmail As String) As E_UserRole
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim lanzaderaDbPath As String
    Dim userRole As E_UserRole
    
    ' Inicializar con rol desconocido
    userRole = Rol_Desconocido
    
    ' Validar que el email no esté vacío
    If Trim(userEmail) = "" Then
        IAuthService_GetUserRole = Rol_Desconocido
        Exit Function
    End If
    
    ' Obtener la ruta de la base de datos Lanzadera desde modConfig
    lanzaderaDbPath = GetLanzaderaDbPath()
    
    ' Verificar que la ruta no esté vacía
    If Trim(lanzaderaDbPath) = "" Then
        Debug.Print "Error: Ruta de base de datos Lanzadera no configurada"
        IAuthService_GetUserRole = Rol_Desconocido
        Exit Function
    End If
    
    ' Verificar que el archivo existe
    If Dir(lanzaderaDbPath) = "" Then
        Debug.Print "Error: No se encuentra la base de datos Lanzadera en: " & lanzaderaDbPath
        IAuthService_GetUserRole = Rol_Desconocido
        Exit Function
    End If
    
    ' Abrir conexión a la base de datos Lanzadera
    Set db = DBEngine.OpenDatabase(lanzaderaDbPath)
    
    ' Construir consulta SQL para obtener permisos del usuario
    ' Según la Sección 2.1 del documento, se consultan las tablas:
    ' - TbUsuariosAplicaciones: para verificar si el usuario tiene acceso
    ' - TbUsuariosAplicacionesPermisos: para obtener los permisos específicos
    sql = "SELECT uap.Permiso " & _
          "FROM TbUsuariosAplicaciones ua " & _
          "INNER JOIN TbUsuariosAplicacionesPermisos uap ON ua.IDUsuarioAplicacion = uap.IDUsuarioAplicacion " & _
          "WHERE ua.CorreoElectronico = '" & Replace(userEmail, "'", "''") & "' " & _
          "AND ua.IDAplicacion = " & IDAplicacion_CONDOR
    
    ' Ejecutar consulta
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    
    ' Procesar resultados para determinar el rol
    If Not rs.EOF Then
        ' El usuario tiene acceso a la aplicación
        ' Determinar el rol basado en los permisos
        Do While Not rs.EOF
            Select Case UCase(Trim(rs("Permiso")))
                Case "ADMIN", "ADMINISTRADOR"
                    userRole = Rol_Admin
                    Exit Do ' Admin es el rol más alto, no necesitamos seguir
                Case "CALIDAD"
                    If userRole < Rol_Calidad Then userRole = Rol_Calidad
                Case "TECNICO", "TÉCNICO"
                    If userRole < Rol_Tecnico Then userRole = Rol_Tecnico
            End Select
            rs.MoveNext
        Loop
    Else
        ' El usuario no tiene permisos o no existe en la base de datos
        Debug.Print "Usuario sin permisos: " & userEmail
        userRole = Rol_Desconocido
    End If
    
    ' Cerrar recursos
    rs.Close
    db.Close
    Set rs = Nothing
    Set db = Nothing
    
    ' Retornar el rol determinado
    IAuthService_GetUserRole = userRole
    
    Debug.Print "Usuario: " & userEmail & " - Rol determinado: " & GetRoleName(userRole)
    
    Exit Function
    
ErrorHandler:
    ' Manejo de errores
    Debug.Print "Error en IAuthService_GetUserRole: " & Err.Number & " - " & Err.Description
    
    ' Limpiar recursos en caso de error
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
    If Not db Is Nothing Then
        db.Close
        Set db = Nothing
    End If
    
    ' En caso de error, retornar rol desconocido
    IAuthService_GetUserRole = Rol_Desconocido
End Function

' =====================================================
' FUNCIÓN AUXILIAR: GetRoleName
' PROPÓSITO: Convertir el enum E_UserRole a texto para depuración
' PARÁMETROS:
'   - role: E_UserRole - Rol a convertir
' RETORNA: String - Nombre del rol
' =====================================================
Private Function GetRoleName(ByVal role As E_UserRole) As String
    Select Case role
        Case Rol_Admin
            GetRoleName = "Administrador"
        Case Rol_Calidad
            GetRoleName = "Calidad"
        Case Rol_Tecnico
            GetRoleName = "Técnico"
        Case Rol_Desconocido
            GetRoleName = "Desconocido"
        Case Else
            GetRoleName = "Indefinido"
    End Select
End Function