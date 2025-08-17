Option Compare Database
Option Explicit
Implements ISolicitudRepository

Private Function ISolicitudRepository_LoadPC(ByVal solicitudId As Long) As CSolicitudPC
    ' TODO: Implementar lógica de carga desde Tb_Solicitudes y Tb_Datos_PC
    Debug.Print "CSolicitudRepository: LoadPC llamado con ID " & solicitudId
    Set ISolicitudRepository_LoadPC = New CSolicitudPC
End Function

Private Function ISolicitudRepository_SavePC(solicitudPC As CSolicitudPC) As Long
    ' TODO: Implementar lógica de guardado transaccional en Tb_Solicitudes y Tb_Datos_PC.
    Debug.Print "CSolicitudRepository: SavePC llamado."
    ISolicitudRepository_SavePC = 1 ' Simular ID devuelto
End Function

Private Function ISolicitudRepository_ObtenerPorId(ByVal idSolicitud As Long) As T_Solicitud
    Dim solicitud As T_Solicitud
    Dim rs As DAO.Recordset
    
    On Error GoTo ErrorHandler
    
    ' Consultar la tabla Tb_Solicitudes por ID
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM Tb_Solicitudes WHERE IdSolicitud = " & idSolicitud)
    
    If Not rs.EOF Then
        Set solicitud = New T_Solicitud
        With solicitud
            .IdSolicitud = rs("IdSolicitud")
            .TipoSolicitud = rs("TipoSolicitud")
            .EstadoInterno = rs("EstadoInterno")
            .FechaCreacion = rs("FechaCreacion")
            .UsuarioCreador = rs("UsuarioCreador")
            If Not IsNull(rs("FechaModificacion")) Then .FechaModificacion = rs("FechaModificacion")
            If Not IsNull(rs("UsuarioModificador")) Then .UsuarioModificador = rs("UsuarioModificador")
        End With
    End If
    
    rs.Close
    Set rs = Nothing
    Set ISolicitudRepository_ObtenerPorId = solicitud
    Exit Function
    
ErrorHandler:
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    Err.Raise Err.Number, "CSolicitudRepository.ObtenerPorId", "Error al obtener solicitud por ID " & idSolicitud & ": " & Err.Description
End Function