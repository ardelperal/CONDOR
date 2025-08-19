Attribute VB_Name = "modDatabase"
Option Compare Database


Option Explicit

' ============================================================================
' M?dulo: modDatabase
' Descripci?n: Servicio de acceso a datos para solicitudes
' Autor: CONDOR-Expert
' Fecha: Diciembre 2024
' ============================================================================

' ============================================================================
' FUNCIONES DE ACCESO A DATOS
' ============================================================================

' Funci?n que obtiene los datos de una solicitud con INNER JOIN
' Par?metros:
'   idSolicitud: ID de la solicitud a obtener
' Retorna: Recordset con los datos de Tb_Solicitudes y TbDatos_PC
Public Function GetSolicitudData(ByVal idSolicitud As Long) As DAO.Recordset
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim sql As String
    
    ' Conectar a la base de datos actual
    Set db = CurrentDb()
    
    ' Construir consulta SQL con INNER JOIN
    sql = "SELECT s.ID, s.NumeroExpediente, s.TipoSolicitud, s.EstadoInterno, s.EstadoRAC, " & _
          "s.FechaCreacion, s.FechaUltimaModificacion, s.Usuario, s.Observaciones, s.Activo, " & _
          "pc.ID as DatosID, pc.DescripcionCambio, pc.JustificacionCambio, " & _
          "pc.ImpactoSeguridad, pc.ImpactoCalidad, pc.Estado as EstadoDatos " & _
          "FROM Tb_Solicitudes s INNER JOIN TbDatos_PC pc ON s.ID = pc.SolicitudID " & _
          "WHERE s.ID = " & idSolicitud & " AND s.Activo = True AND pc.Activo = True"
    
    ' Ejecutar consulta
    Set GetSolicitudData = db.OpenRecordset(sql, dbOpenSnapshot)
    
    Exit Function
    
ErrorHandler:
    Call modErrorHandler.LogError(Err.Number, Err.Description, "modDatabase.GetSolicitudData")
    Set GetSolicitudData = Nothing
    If Not db Is Nothing Then db.Close
End Function

' Funci?n que guarda o actualiza una solicitud PC
' Par?metros:
'   solicitudData: Estructura T_Solicitud con los datos generales
'   pcData: Estructura T_Datos_PC con los datos espec?ficos
' Retorna: True si la operaci?n fue exitosa, False en caso contrario
Public Function SaveSolicitudPC(ByRef solicitudData As T_Solicitud, ByRef pcData As T_Datos_PC) As Boolean
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim rsSolicitud As DAO.Recordset
    Dim rsPC As DAO.Recordset
    Dim isNewRecord As Boolean
    
    SaveSolicitudPC = False
    Set db = CurrentDb
    
    ' --- INICIO DE LA TRANSACCIÓN ---
    ' Se gestiona a través del espacio de trabajo DBEngine
    DBEngine.BeginTrans
    
    ' Determinar si es un registro nuevo
    isNewRecord = (solicitudData.ID = 0)
    
    ' ============================================================================
    ' GUARDAR/ACTUALIZAR Tb_Solicitudes
    ' ============================================================================
    
    If isNewRecord Then
        ' INSERT en Tb_Solicitudes
        Set rsSolicitud = db.OpenRecordset("Tb_Solicitudes", dbOpenDynaset)
        rsSolicitud.AddNew
    Else
        ' UPDATE en Tb_Solicitudes
        Set rsSolicitud = db.OpenRecordset("SELECT * FROM Tb_Solicitudes WHERE ID = " & solicitudData.ID, dbOpenDynaset)
        If Not rsSolicitud.EOF Then
            rsSolicitud.Edit
        Else
            ' El registro no existe, tratarlo como nuevo
            rsSolicitud.Close
            Set rsSolicitud = db.OpenRecordset("Tb_Solicitudes", dbOpenDynaset)
            rsSolicitud.AddNew
            isNewRecord = True
        End If
    End If
    
    ' Asignar valores a Tb_Solicitudes
    With rsSolicitud
        !NumeroExpediente = solicitudData.NumeroExpediente
        !tipoSolicitud = solicitudData.tipoSolicitud
        !estadoInterno = solicitudData.estadoInterno
        !EstadoRAC = solicitudData.EstadoRAC
        !fechaCreacion = IIf(isNewRecord, Now(), solicitudData.fechaCreacion)
        !FechaUltimaModificacion = Now()
        !usuario = solicitudData.usuario
        !Observaciones = solicitudData.Observaciones
        !Activo = solicitudData.Activo
        .Update
        
        ' Obtener el ID generado si es nuevo registro
        If isNewRecord Then
            .Bookmark = .LastModified
            solicitudData.ID = !ID
            pcData.SolicitudID = !ID
        End If
    End With
    rsSolicitud.Close
    
    ' ============================================================================
    ' GUARDAR/ACTUALIZAR TbDatos_PC
    ' ============================================================================
    
    If pcData.ID = 0 Then
        ' INSERT en TbDatos_PC
        Set rsPC = db.OpenRecordset("TbDatos_PC", dbOpenDynaset)
        rsPC.AddNew
    Else
        ' UPDATE en TbDatos_PC
        Set rsPC = db.OpenRecordset("SELECT * FROM TbDatos_PC WHERE ID = " & pcData.ID, dbOpenDynaset)
        If Not rsPC.EOF Then
            rsPC.Edit
        Else
            ' El registro no existe, tratarlo como nuevo
            rsPC.Close
            Set rsPC = db.OpenRecordset("TbDatos_PC", dbOpenDynaset)
            rsPC.AddNew
        End If
    End If
    
    ' Asignar valores a TbDatos_PC
    With rsPC
        !SolicitudID = pcData.SolicitudID
        !NumeroExpediente = pcData.NumeroExpediente
        !tipoSolicitud = pcData.tipoSolicitud
        !descripcionCambio = pcData.descripcionCambio
        !JustificacionCambio = pcData.JustificacionCambio
        !ImpactoSeguridad = pcData.ImpactoSeguridad
        !impactoCalidad = pcData.impactoCalidad
        !fechaCreacion = IIf(pcData.ID = 0, Now(), pcData.fechaCreacion)
        !FechaUltimaModificacion = Now()
        !Estado = pcData.Estado
        !Activo = pcData.Activo
        .Update
        
        ' Obtener el ID generado si es nuevo registro
        If pcData.ID = 0 Then
            .Bookmark = .LastModified
            pcData.ID = !ID
        End If
    End With
    rsPC.Close
    
    ' --- FINALIZACIÓN DE LA TRANSACCIÓN ---
    DBEngine.CommitTrans
    SaveSolicitudPC = True
    
    Set db = Nothing
    Exit Function
    
ErrorHandler:
    SaveSolicitudPC = False
    ' Si hubo un error, deshacer todos los cambios de la transacción
    DBEngine.Rollback
    Call modErrorHandler.LogError(Err.Number, Err.Description, "modDatabase.SaveSolicitudPC")
    ' Limpiar objetos
    If Not rsSolicitud Is Nothing Then rsSolicitud.Close
    If Not rsPC Is Nothing Then rsPC.Close
    If Not db Is Nothing Then Set db = Nothing
End Function

' ============================================================================
' FUNCIONES AUXILIARES
' ============================================================================

' Funci?n auxiliar para verificar si una solicitud existe
Public Function SolicitudExists(ByVal idSolicitud As Long) As Boolean
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    
    Set db = CurrentDb()
    Set rs = db.OpenRecordset("SELECT COUNT(*) as Total FROM Tb_Solicitudes WHERE ID = " & idSolicitud & " AND Activo = True", dbOpenSnapshot)
    
    SolicitudExists = (rs!total > 0)
    
    rs.Close
    db.Close
    
    Exit Function
    
ErrorHandler:
    Call modErrorHandler.LogError(Err.Number, Err.Description, "modDatabase.SolicitudExists")
    SolicitudExists = False
    If Not rs Is Nothing Then rs.Close
    If Not db Is Nothing Then db.Close
End Function




















