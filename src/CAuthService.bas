Attribute VB_Name = "CAuthService"
' =====================================================
' CLASE: CAuthService
' PROPOSITO: Implementacion concreta del servicio de autenticacion
' IMPLEMENTA: IAuthService
' AUTOR: CONDOR-Expert
' FECHA: 2025-01-14
' =====================================================

Implements IAuthService

' =====================================================
' FUNCION: IAuthService_GetUserRole
' PROPOSITO: Determinar el rol de un usuario consultando la BD Lanzadera
' PARAMETROS:
'   - userEmail: String - Email del usuario a verificar
' RETORNA: E_UserRole - Rol del usuario segun los permisos en Lanzadera
' =====================================================
Private Function IAuthService_GetUserRole(ByVal userEmail As String) As E_UserRole
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim LanzaderaDbPath As String
    Dim userRole As E_UserRole
    
    ' Inicializar con rol desconocido
    userRole = Rol_Desconocido
    
    ' Validar que el email no este vacio
    If Trim(userEmail) = "" Then
        IAuthService_GetUserRole = Rol_Desconocido
        Exit Function
    End If
    
    ' Obtener la ruta de la base de datos Lanzadera desde modConfig
    LanzaderaDbPath = GetLanzaderaDbPath()
    
    ' Verificar que la ruta no este vacia
    If Trim(LanzaderaDbPath) = "" Then
        Debug.Print "Error: Ruta de base de datos Lanzadera no configurada"
        IAuthService_GetUserRole = Rol_Desconocido
        Exit Function
    End If
    
    ' Verificar que el archivo existe
    If Dir(LanzaderaDbPath) = "" Then
        Debug.Print "Error: No se encuentra la base de datos Lanzadera en: " & LanzaderaDbPath
        IAuthService_GetUserRole = Rol_Desconocido
        Exit Function
    End If
    
    ' Abrir conexion a la base de datos Lanzadera
    Set db = DBEngine.OpenDatabase(LanzaderaDbPath)
    
    ' Construir consulta SQL para obtener permisos del usuario
    ' Segun la Seccion 2.1 del documento, se consultan las tablas:
    ' - TbUsuariosAplicaciones: para verificar si el usuario tiene acceso
    ' - TbUsuariosAplicacionesPermisos: para obtener los permisos especificos
    sql = "SELECT uap.Permiso " & _
          "FROM TbUsuariosAplicaciones ua " & _
          "INNER JOIN TbUsuariosAplicacionesPermisos uap ON ua.IDUsuarioAplicacion = uap.IDUsuarioAplicacion " & _
          "WHERE ua.CorreoElectronico = '" & Replace(userEmail, "'", "''") & "' " & _
          "AND ua.IDAplicacion = " & IDAplicacion_CONDOR
    
    ' Ejecutar consulta
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    
    ' Procesar resultados para determinar el rol
    If Not rs.EOF Then
        ' El usuario tiene acceso a la aplicacion
        ' Determinar el rol basado en los permisos
        Do While Not rs.EOF
            Select Case UCase(Trim(rs("Permiso")))
                Case "ADMIN", "ADMINISTRADOR"
                    userRole = Rol_Admin
                    Exit Do ' Admin es el rol mas alto, no necesitamos seguir
                Case "CALIDAD"
                    If userRole < Rol_Calidad Then userRole = Rol_Calidad
                Case "TECNICO", "TECNICO"
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
        On Error Resume Next
        rs.Close
        On Error GoTo 0
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
' FUNCION AUXILIAR: GetRoleName
' PROPOSITO: Convertir el enum E_UserRole a texto para depuracion
' PARAMETROS:
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
            GetRoleName = "Tecnico"
        Case Rol_Desconocido
            GetRoleName = "Desconocido"
        Case Else
            GetRoleName = "Indefinido"
    End Select
End Function





