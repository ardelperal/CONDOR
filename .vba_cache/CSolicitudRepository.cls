Option Compare Database
Option Explicit
' Implementación del repositorio de solicitudes
' Maneja la persistencia en las tablas tbSolicitud, tbDatosPC, tbDatosCDCA y tbDatosCDCASUB

Implements ISolicitudRepository

Private m_OperationLogger As IOperationLogger

Private Function ISolicitudRepository_GuardarSolicitud(ByVal solicitud As T_Solicitud) As Long
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim idSolicitudGenerado As Long
    
    ' Logging de auditoría: Inicio de guardado en repositorio
    If Not m_OperationLogger Is Nothing Then
        Dim operationType As String
        If solicitud.idSolicitud = 0 Then
            operationType = "INSERT"
        Else
            operationType = "UPDATE"
        End If
        m_OperationLogger.LogOperation "REPO_SAVE_START", operationType & " solicitud ID: " & solicitud.idSolicitud & ", Expediente: " & solicitud.idExpediente, "CSolicitudRepository.GuardarSolicitud"
    End If
    
    Set db = CurrentDb
    
    ' Iniciar transacción
    db.BeginTrans
    
    On Error GoTo ErrorHandler
    
    ' Guardar o actualizar solicitud principal
    If solicitud.idSolicitud = 0 Then
        ' INSERT - Nueva solicitud
        Set rs = db.OpenRecordset("tbSolicitud", dbOpenDynaset)
        rs.AddNew
        rs("idExpediente") = solicitud.idExpediente
        rs("tipoSolicitud") = solicitud.tipoSolicitud
        rs("subTipoSolicitud") = solicitud.subTipoSolicitud
        rs("codigoSolicitud") = solicitud.codigoSolicitud
        rs("estadoInterno") = solicitud.estadoInterno
        rs("fechaCreacion") = solicitud.fechaCreacion
        rs("usuarioCreacion") = solicitud.usuarioCreacion
        rs("fechaPaseTecnico") = solicitud.fechaPaseTecnico
        rs("fechaCompletadoTecnico") = solicitud.fechaCompletadoTecnico
        rs.Update
        
        ' Obtener el ID generado
        rs.MoveLast
        idSolicitudGenerado = rs("idSolicitud")
        rs.Close
    Else
        ' UPDATE - Solicitud existente
        idSolicitudGenerado = solicitud.idSolicitud
        Set rs = db.OpenRecordset("SELECT * FROM tbSolicitud WHERE idSolicitud = " & solicitud.idSolicitud, dbOpenDynaset)
        If Not rs.EOF Then
            rs.Edit
            rs("idExpediente") = solicitud.idExpediente
            rs("tipoSolicitud") = solicitud.tipoSolicitud
            rs("subTipoSolicitud") = solicitud.subTipoSolicitud
            rs("codigoSolicitud") = solicitud.codigoSolicitud
            rs("estadoInterno") = solicitud.estadoInterno
            rs("fechaCreacion") = solicitud.fechaCreacion
            rs("usuarioCreacion") = solicitud.usuarioCreacion
            rs("fechaPaseTecnico") = solicitud.fechaPaseTecnico
            rs("fechaCompletadoTecnico") = solicitud.fechaCompletadoTecnico
            rs.Update
        End If
        rs.Close
    End If
    
    ' Guardar datos específicos según el tipo de solicitud
    If solicitud.tipoSolicitud = "PC" And Not solicitud.datosPC Is Nothing Then
        Call GuardarDatosPC(db, idSolicitudGenerado, solicitud.datosPC)
    ElseIf solicitud.tipoSolicitud = "CD/CA" And Not solicitud.datosCDCA Is Nothing Then
        Call GuardarDatosCDCA(db, idSolicitudGenerado, solicitud.datosCDCA)
    ElseIf solicitud.tipoSolicitud = "CD/CA-SUB" And Not solicitud.datosCDCASUB Is Nothing Then
        Call GuardarDatosCDCASUB(db, idSolicitudGenerado, solicitud.datosCDCASUB)
    End If
    
    ' Confirmar transacción
    db.CommitTrans
    
    ' Logging de auditoría: Éxito en guardado
    If Not m_OperationLogger Is Nothing Then
        m_OperationLogger.LogOperation "REPO_SAVE_SUCCESS", "Solicitud guardada exitosamente ID: " & idSolicitudGenerado & ", Expediente: " & solicitud.idExpediente, "CSolicitudRepository.GuardarSolicitud"
    End If
    
    ISolicitudRepository_GuardarSolicitud = idSolicitudGenerado
    Exit Function
    
ErrorHandler:
    If Not m_OperationLogger Is Nothing Then
        m_OperationLogger.LogOperation "REPO_SAVE_ERROR", "Error en guardado de solicitud: " & Err.Description, "CSolicitudRepository.GuardarSolicitud"
    End If
    modErrorHandler.LogError Err.Number, Err.Description, "CSolicitudRepository.GuardarSolicitud"
    db.Rollback
    Err.Raise Err.Number, "CSolicitudRepository.GuardarSolicitud", Err.Description
End Function

Private Function ISolicitudRepository_ObtenerSolicitudPorId(ByVal idSolicitud As Long) As T_Solicitud
    ' Logging de auditoría: Consulta de solicitud por ID
    If Not m_OperationLogger Is Nothing Then
        m_OperationLogger.LogOperation "REPO_GET_SOLICITUD", "Consultando solicitud por ID: " & idSolicitud, "CSolicitudRepository.ObtenerSolicitudPorId"
    End If
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim solicitud As T_Solicitud
    
    On Error GoTo ErrorHandler
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT * FROM tbSolicitud WHERE idSolicitud = " & idSolicitud, dbOpenSnapshot)
    
    If Not rs.EOF Then
        Set solicitud = New T_Solicitud
        solicitud.idSolicitud = rs("idSolicitud")
        solicitud.idExpediente = rs("idExpediente") & ""
        solicitud.tipoSolicitud = rs("tipoSolicitud") & ""
        solicitud.subTipoSolicitud = rs("subTipoSolicitud") & ""
        solicitud.codigoSolicitud = rs("codigoSolicitud") & ""
        solicitud.estadoInterno = rs("estadoInterno") & ""
        If Not IsNull(rs("fechaCreacion")) Then solicitud.fechaCreacion = rs("fechaCreacion")
        solicitud.usuarioCreacion = rs("usuarioCreacion") & ""
        If Not IsNull(rs("fechaPaseTecnico")) Then solicitud.fechaPaseTecnico = rs("fechaPaseTecnico")
        If Not IsNull(rs("fechaCompletadoTecnico")) Then solicitud.fechaCompletadoTecnico = rs("fechaCompletadoTecnico")
        
        ' Cargar datos específicos según el tipo
        If solicitud.tipoSolicitud = "PC" Then
            Set solicitud.datosPC = CargarDatosPC(db, idSolicitud)
        ElseIf solicitud.tipoSolicitud = "CD/CA" Then
            Set solicitud.datosCDCA = CargarDatosCDCA(db, idSolicitud)
        ElseIf solicitud.tipoSolicitud = "CD/CA-SUB" Then
            Set solicitud.datosCDCASUB = CargarDatosCDCASUB(db, idSolicitud)
        End If
    End If
    
    rs.Close
    
    ' Logging de auditoría: Resultado de consulta
    If Not m_OperationLogger Is Nothing Then
        If Not solicitud Is Nothing Then
            m_OperationLogger.LogOperation "REPO_SOLICITUD_FOUND", "Solicitud encontrada ID: " & idSolicitud & ", Expediente: " & solicitud.idExpediente, "CSolicitudRepository.ObtenerSolicitudPorId"
        Else
            m_OperationLogger.LogOperation "REPO_SOLICITUD_NOT_FOUND", "Solicitud no encontrada ID: " & idSolicitud, "CSolicitudRepository.ObtenerSolicitudPorId"
        End If
    End If
    
    Set ISolicitudRepository_ObtenerSolicitudPorId = solicitud
    Exit Function
    
ErrorHandler:
    ' Logging de auditoría: Error en consulta
    If Not m_OperationLogger Is Nothing Then
        m_OperationLogger.LogOperation "REPO_GET_SOLICITUD_ERROR", "Error consultando solicitud ID: " & idSolicitud & ". Error: " & Err.Description, "CSolicitudRepository.ObtenerSolicitudPorId"
    End If
    
    modErrorHandler.LogError Err.Number, Err.Description, "CSolicitudRepository.ObtenerSolicitudPorId"
    Set ISolicitudRepository_ObtenerSolicitudPorId = Nothing
End Function

Private Function ISolicitudRepository_ObtenerSolicitudPorCodigo(ByVal codigoSolicitud As String) As T_Solicitud
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim solicitud As T_Solicitud
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT * FROM tbSolicitud WHERE codigoSolicitud = '" & codigoSolicitud & "'", dbOpenSnapshot)
    
    If Not rs.EOF Then
        Set solicitud = ISolicitudRepository_ObtenerSolicitudPorId(rs("idSolicitud"))
    End If
    
    rs.Close
    Set ISolicitudRepository_ObtenerSolicitudPorCodigo = solicitud
End Function

Private Function ISolicitudRepository_EliminarSolicitud(ByVal idSolicitud As Long) As Boolean
    ' Logging de auditoría: Eliminación de solicitud
    If Not m_OperationLogger Is Nothing Then
        m_OperationLogger.LogOperation "REPO_DELETE_SOLICITUD", "Eliminando solicitud ID: " & idSolicitud, "CSolicitudRepository.EliminarSolicitud"
    End If
    
    Dim db As DAO.Database
    
    Set db = CurrentDb
    db.BeginTrans
    
    On Error GoTo ErrorHandler
    
    ' Eliminar datos específicos primero (integridad referencial)
    db.Execute "DELETE FROM tbDatosPC WHERE idSolicitud = " & idSolicitud
    db.Execute "DELETE FROM tbDatosCDCA WHERE idSolicitud = " & idSolicitud
    db.Execute "DELETE FROM tbDatosCDCASUB WHERE idSolicitud = " & idSolicitud
    
    ' Eliminar solicitud principal
    db.Execute "DELETE FROM tbSolicitud WHERE idSolicitud = " & idSolicitud
    
    db.CommitTrans
    
    ' Logging de auditoría: Solicitud eliminada exitosamente
    If Not m_OperationLogger Is Nothing Then
        m_OperationLogger.LogOperation "REPO_SOLICITUD_DELETED", "Solicitud eliminada exitosamente ID: " & idSolicitud, "CSolicitudRepository.EliminarSolicitud"
    End If
    
    ISolicitudRepository_EliminarSolicitud = True
    Exit Function
    
ErrorHandler:
    ' Logging de auditoría: Error en eliminación
    If Not m_OperationLogger Is Nothing Then
        m_OperationLogger.LogOperation "REPO_DELETE_ERROR", "Error eliminando solicitud ID: " & idSolicitud & ". Error: " & Err.Description, "CSolicitudRepository.EliminarSolicitud"
    End If
    
    modErrorHandler.LogError Err.Number, Err.Description, "CSolicitudRepository.EliminarSolicitud"
    db.Rollback
    ISolicitudRepository_EliminarSolicitud = False
End Function

Private Function ISolicitudRepository_ObtenerSolicitudesPorExpediente(ByVal idExpediente As String) As Collection
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim solicitudes As Collection
    Dim solicitud As T_Solicitud
    
    Set solicitudes = New Collection
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT * FROM tbSolicitud WHERE idExpediente = '" & idExpediente & "' ORDER BY fechaCreacion", dbOpenSnapshot)
    
    While Not rs.EOF
        Set solicitud = ISolicitudRepository_ObtenerSolicitudPorId(rs("idSolicitud"))
        solicitudes.Add solicitud
        rs.MoveNext
    Wend
    
    rs.Close
    Set ISolicitudRepository_ObtenerSolicitudesPorExpediente = solicitudes
End Function

Private Function ISolicitudRepository_ExisteSolicitudConCodigo(ByVal codigoSolicitud As String) As Boolean
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT COUNT(*) AS Total FROM tbSolicitud WHERE codigoSolicitud = '" & codigoSolicitud & "'", dbOpenSnapshot)
    
    ISolicitudRepository_ExisteSolicitudConCodigo = (rs("Total") > 0)
    rs.Close
End Function

' Métodos privados para manejo de tablas secundarias
Private Sub GuardarDatosPC(ByVal db As DAO.Database, ByVal idSolicitud As Long, ByVal datosPC As T_Datos_PC)
    Dim rs As DAO.Recordset
    
    ' Verificar si ya existen datos PC para esta solicitud
    Set rs = db.OpenRecordset("SELECT * FROM tbDatosPC WHERE idSolicitud = " & idSolicitud, dbOpenDynaset)
    
    If rs.EOF Then
        ' INSERT
        rs.AddNew
        rs("idSolicitud") = idSolicitud
    Else
        ' UPDATE
        rs.Edit
    End If
    
    rs("refContratoInspeccionOficial") = datosPC.refContratoInspeccionOficial
    rs("refContratoInspeccionOficialNuevo") = datosPC.refContratoInspeccionOficialNuevo
    rs("refSuministrador") = datosPC.refSuministrador
    rs("refSuministradorNuevo") = datosPC.refSuministradorNuevo
    rs("refSubSuministrador") = datosPC.refSubSuministrador
    rs("refSubSuministradorNuevo") = datosPC.refSubSuministradorNuevo
    rs("importeOriginal") = datosPC.importeOriginal
    rs("importeNuevo") = datosPC.importeNuevo
    rs("justificacion") = datosPC.justificacion
    rs("descripcionCambio") = datosPC.descripcionCambio
    rs("impactoCalidad") = datosPC.impactoCalidad
    rs("impactoCoste") = datosPC.impactoCoste
    rs("impactoPlazos") = datosPC.impactoPlazos
    rs("requiereModificacionContrato") = datosPC.requiereModificacionContrato
    rs("efectoFechaEntrega") = datosPC.efectoFechaEntrega
    rs("observacionesAdicionales") = datosPC.observacionesAdicionales
    
    rs.Update
    rs.Close
End Sub

Private Sub GuardarDatosCDCA(ByVal db As DAO.Database, ByVal idSolicitud As Long, ByVal datosCDCA As T_Datos_CD_CA)
    Dim rs As DAO.Recordset
    
    Set rs = db.OpenRecordset("SELECT * FROM tbDatosCDCA WHERE idSolicitud = " & idSolicitud, dbOpenDynaset)
    
    If rs.EOF Then
        rs.AddNew
        rs("idSolicitud") = idSolicitud
    Else
        rs.Edit
    End If
    
    rs("refSuministrador") = datosCDCA.refSuministrador
    rs("descripcionImpactoNC") = datosCDCA.descripcionImpactoNC
    rs("refDesviacionesPrevias") = datosCDCA.refDesviacionesPrevias
    rs("causaNC") = datosCDCA.causaNC
    rs("impactoCoste") = datosCDCA.impactoCoste
    rs("clasificacionNC") = datosCDCA.clasificacionNC
    rs("requiereModificacionContrato") = datosCDCA.requiereModificacionContrato
    rs("efectoFechaEntrega") = datosCDCA.efectoFechaEntrega
    rs("identificacionAutoridadDiseno") = datosCDCA.identificacionAutoridadDiseno
    rs("esSuministradorAD") = datosCDCA.esSuministradorAD
    rs("racRef") = datosCDCA.racRef
    rs("racCodigo") = datosCDCA.racCodigo
    rs("observacionesRAC") = datosCDCA.observacionesRAC
    If datosCDCA.fechaFirmaRAC <> 0 Then rs("fechaFirmaRAC") = datosCDCA.fechaFirmaRAC
    rs("decisionFinal") = datosCDCA.decisionFinal
    rs("observacionesFinales") = datosCDCA.observacionesFinales
    If datosCDCA.fechaFirmaDecisionFinal <> 0 Then rs("fechaFirmaDecisionFinal") = datosCDCA.fechaFirmaDecisionFinal
    rs("cargoFirmanteFinal") = datosCDCA.cargoFirmanteFinal
    
    rs.Update
    rs.Close
End Sub

Private Sub GuardarDatosCDCASUB(ByVal db As DAO.Database, ByVal idSolicitud As Long, ByVal datosCDCASUB As T_Datos_CD_CA_SUB)
    Dim rs As DAO.Recordset
    
    Set rs = db.OpenRecordset("SELECT * FROM tbDatosCDCASUB WHERE idSolicitud = " & idSolicitud, dbOpenDynaset)
    
    If rs.EOF Then
        rs.AddNew
        rs("idSolicitud") = idSolicitud
    Else
        rs.Edit
    End If
    
    rs("refSuministrador") = datosCDCASUB.refSuministrador
    rs("refSubSuministrador") = datosCDCASUB.refSubSuministrador
    rs("suministradorPrincipalNombreDir") = datosCDCASUB.suministradorPrincipalNombreDir
    rs("subSuministradorNombreDir") = datosCDCASUB.subSuministradorNombreDir
    rs("identificacionMaterial") = datosCDCASUB.identificacionMaterial
    rs("numPlanoEspecificacion") = datosCDCASUB.numPlanoEspecificacion
    rs("cantidadPeriodo") = datosCDCASUB.cantidadPeriodo
    rs("numSerieLote") = datosCDCASUB.numSerieLote
    rs("afectaPrestaciones") = datosCDCASUB.afectaPrestaciones
    rs("afectaSeguridad") = datosCDCASUB.afectaSeguridad
    rs("afectaFiabilidad") = datosCDCASUB.afectaFiabilidad
    rs("afectaVidaUtil") = datosCDCASUB.afectaVidaUtil
    rs("afectaMedioambiente") = datosCDCASUB.afectaMedioambiente
    rs("afectaIntercambiabilidad") = datosCDCASUB.afectaIntercambiabilidad
    rs("afectaMantenibilidad") = datosCDCASUB.afectaMantenibilidad
    rs("afectaApariencia") = datosCDCASUB.afectaApariencia
    rs("afectaOtros") = datosCDCASUB.afectaOtros
    rs("esSubSuministradorAD") = datosCDCASUB.esSubSuministradorAD
    rs("nombreRepSubSuministrador") = datosCDCASUB.nombreRepSubSuministrador
    rs("decisionSuministradorPrincipal") = datosCDCASUB.decisionSuministradorPrincipal
    rs("obsSuministradorPrincipal") = datosCDCASUB.obsSuministradorPrincipal
    If datosCDCASUB.fechaFirmaSuministradorPrincipal <> 0 Then rs("fechaFirmaSuministradorPrincipal") = datosCDCASUB.fechaFirmaSuministradorPrincipal
    rs("firmaSuministradorPrincipalNombreCargo") = datosCDCASUB.firmaSuministradorPrincipalNombreCargo
    rs("obsRACDelegador") = datosCDCASUB.obsRACDelegador
    If datosCDCASUB.fechaFirmaRACDelegador <> 0 Then rs("fechaFirmaRACDelegador") = datosCDCASUB.fechaFirmaRACDelegador
    
    rs.Update
    rs.Close
End Sub

Private Function CargarDatosPC(ByVal db As DAO.Database, ByVal idSolicitud As Long) As T_Datos_PC
    Dim rs As DAO.Recordset
    Dim datosPC As T_Datos_PC
    
    Set rs = db.OpenRecordset("SELECT * FROM tbDatosPC WHERE idSolicitud = " & idSolicitud, dbOpenSnapshot)
    
    If Not rs.EOF Then
        Set datosPC = New T_Datos_PC
        datosPC.idDatosPC = rs("idDatosPC")
        datosPC.refContratoInspeccionOficial = rs("refContratoInspeccionOficial") & ""
        datosPC.refContratoInspeccionOficialNuevo = rs("refContratoInspeccionOficialNuevo") & ""
        datosPC.refSuministrador = rs("refSuministrador") & ""
        datosPC.refSuministradorNuevo = rs("refSuministradorNuevo") & ""
        datosPC.refSubSuministrador = rs("refSubSuministrador") & ""
        datosPC.refSubSuministradorNuevo = rs("refSubSuministradorNuevo") & ""
        If Not IsNull(rs("importeOriginal")) Then datosPC.importeOriginal = rs("importeOriginal")
        If Not IsNull(rs("importeNuevo")) Then datosPC.importeNuevo = rs("importeNuevo")
        datosPC.justificacion = rs("justificacion") & ""
        datosPC.descripcionCambio = rs("descripcionCambio") & ""
        datosPC.impactoCalidad = rs("impactoCalidad") & ""
        datosPC.impactoCoste = rs("impactoCoste") & ""
        datosPC.impactoPlazos = rs("impactoPlazos") & ""
        datosPC.requiereModificacionContrato = rs("requiereModificacionContrato")
        datosPC.efectoFechaEntrega = rs("efectoFechaEntrega") & ""
        datosPC.observacionesAdicionales = rs("observacionesAdicionales") & ""
    End If
    
    rs.Close
    Set CargarDatosPC = datosPC
End Function

Private Function CargarDatosCDCA(ByVal db As DAO.Database, ByVal idSolicitud As Long) As T_Datos_CD_CA
    Dim rs As DAO.Recordset
    Dim datosCDCA As T_Datos_CD_CA
    
    Set rs = db.OpenRecordset("SELECT * FROM tbDatosCDCA WHERE idSolicitud = " & idSolicitud, dbOpenSnapshot)
    
    If Not rs.EOF Then
        Set datosCDCA = New T_Datos_CD_CA
        datosCDCA.idDatosCDCA = rs("idDatosCDCA")
        datosCDCA.refSuministrador = rs("refSuministrador") & ""
        datosCDCA.descripcionImpactoNC = rs("descripcionImpactoNC") & ""
        datosCDCA.refDesviacionesPrevias = rs("refDesviacionesPrevias") & ""
        datosCDCA.causaNC = rs("causaNC") & ""
        datosCDCA.impactoCoste = rs("impactoCoste") & ""
        datosCDCA.clasificacionNC = rs("clasificacionNC") & ""
        datosCDCA.requiereModificacionContrato = rs("requiereModificacionContrato")
        datosCDCA.efectoFechaEntrega = rs("efectoFechaEntrega") & ""
        datosCDCA.identificacionAutoridadDiseno = rs("identificacionAutoridadDiseno") & ""
        datosCDCA.esSuministradorAD = rs("esSuministradorAD")
        datosCDCA.racRef = rs("racRef") & ""
        datosCDCA.racCodigo = rs("racCodigo") & ""
        datosCDCA.observacionesRAC = rs("observacionesRAC") & ""
        If Not IsNull(rs("fechaFirmaRAC")) Then datosCDCA.fechaFirmaRAC = rs("fechaFirmaRAC")
        datosCDCA.decisionFinal = rs("decisionFinal") & ""
        datosCDCA.observacionesFinales = rs("observacionesFinales") & ""
        If Not IsNull(rs("fechaFirmaDecisionFinal")) Then datosCDCA.fechaFirmaDecisionFinal = rs("fechaFirmaDecisionFinal")
        datosCDCA.cargoFirmanteFinal = rs("cargoFirmanteFinal") & ""
    End If
    
    rs.Close
    Set CargarDatosCDCA = datosCDCA
End Function

Private Function CargarDatosCDCASUB(ByVal db As DAO.Database, ByVal idSolicitud As Long) As T_Datos_CD_CA_SUB
    Dim rs As DAO.Recordset
    Dim datosCDCASUB As T_Datos_CD_CA_SUB
    
    Set rs = db.OpenRecordset("SELECT * FROM tbDatosCDCASUB WHERE idSolicitud = " & idSolicitud, dbOpenSnapshot)
    
    If Not rs.EOF Then
        Set datosCDCASUB = New T_Datos_CD_CA_SUB
        datosCDCASUB.idDatosCDCASUB = rs("idDatosCDCASUB")
        datosCDCASUB.refSuministrador = rs("refSuministrador") & ""
        datosCDCASUB.refSubSuministrador = rs("refSubSuministrador") & ""
        datosCDCASUB.suministradorPrincipalNombreDir = rs("suministradorPrincipalNombreDir") & ""
        datosCDCASUB.subSuministradorNombreDir = rs("subSuministradorNombreDir") & ""
        datosCDCASUB.identificacionMaterial = rs("identificacionMaterial") & ""
        datosCDCASUB.numPlanoEspecificacion = rs("numPlanoEspecificacion") & ""
        datosCDCASUB.cantidadPeriodo = rs("cantidadPeriodo")
        datosCDCASUB.numSerieLote = rs("numSerieLote") & ""
        datosCDCASUB.afectaPrestaciones = rs("afectaPrestaciones")
        datosCDCASUB.afectaSeguridad = rs("afectaSeguridad")
        datosCDCASUB.afectaFiabilidad = rs("afectaFiabilidad")
        datosCDCASUB.afectaVidaUtil = rs("afectaVidaUtil")
        datosCDCASUB.afectaMedioambiente = rs("afectaMedioambiente")
        datosCDCASUB.afectaIntercambiabilidad = rs("afectaIntercambiabilidad")
        datosCDCASUB.afectaMantenibilidad = rs("afectaMantenibilidad")
        datosCDCASUB.afectaApariencia = rs("afectaApariencia")
        datosCDCASUB.afectaOtros = rs("afectaOtros")
        datosCDCASUB.esSubSuministradorAD = rs("esSubSuministradorAD")
        datosCDCASUB.nombreRepSubSuministrador = rs("nombreRepSubSuministrador") & ""
        datosCDCASUB.decisionSuministradorPrincipal = rs("decisionSuministradorPrincipal") & ""
        datosCDCASUB.obsSuministradorPrincipal = rs("obsSuministradorPrincipal") & ""
        If Not IsNull(rs("fechaFirmaSuministradorPrincipal")) Then datosCDCASUB.fechaFirmaSuministradorPrincipal = rs("fechaFirmaSuministradorPrincipal")
        datosCDCASUB.firmaSuministradorPrincipalNombreCargo = rs("firmaSuministradorPrincipalNombreCargo") & ""
        datosCDCASUB.obsRACDelegador = rs("obsRACDelegador") & ""
        If Not IsNull(rs("fechaFirmaRACDelegador")) Then datosCDCASUB.fechaFirmaRACDelegador = rs("fechaFirmaRACDelegador")
    End If
    
    rs.Close
    Set CargarDatosCDCASUB = datosCDCASUB
End Function

' Implementación del método genérico Load (Principio de Abstracción de Interfaces)
Private Function ISolicitudRepository_Load(ByVal ID As Long) As ISolicitud
    On Error GoTo ErrorHandler
    
    ' Obtener la solicitud completa usando el método existente
    Dim solicitud As T_Solicitud
    Set solicitud = ISolicitudRepository_ObtenerSolicitudPorId(ID)
    
    If Not solicitud Is Nothing Then
        ' Crear la instancia apropiada según el tipo de solicitud
        If solicitud.tipoSolicitud = "PC" Then
            Dim solicitudPC As CSolicitudPC
            Set solicitudPC = New CSolicitudPC
            ' Copiar propiedades básicas
            solicitudPC.idSolicitud = solicitud.idSolicitud
            solicitudPC.idExpediente = solicitud.idExpediente
            solicitudPC.tipoSolicitud = solicitud.tipoSolicitud
            solicitudPC.codigoSolicitud = solicitud.codigoSolicitud
            solicitudPC.estadoInterno = solicitud.estadoInterno
            ' Asignar datos específicos
            If Not solicitud.datosPC Is Nothing Then
                Set solicitudPC.datosPC = solicitud.datosPC
            End If
            Set ISolicitudRepository_Load = solicitudPC
        ' Aquí se añadirían otros tipos cuando existan las clases CSolicitudCDCA, etc.
        Else
            ' Para tipos no implementados, devolver Nothing
            Set ISolicitudRepository_Load = Nothing
        End If
    Else
         Set ISolicitudRepository_Load = Nothing
     End If
     
     Exit Function
     
ErrorHandler:
    modErrorHandler.LogError Err.Number, Err.Description, "CSolicitudRepository.Load"
    Set ISolicitudRepository_Load = Nothing
 End Function

' Implementación del método genérico Save (Principio de Abstracción de Interfaces)
Private Function ISolicitudRepository_Save(ByVal solicitud As ISolicitud) As Long
    On Error GoTo ErrorHandler
    
    ' Determinar el tipo de solicitud y usar el método apropiado
    If TypeOf solicitud Is CSolicitudPC Then
        Dim solicitudPC As CSolicitudPC
        Set solicitudPC = solicitud
        ' Convertir CSolicitudPC a T_Solicitud y usar el método existente
        Dim solicitudData As T_Solicitud
        Set solicitudData = New T_Solicitud
        solicitudData.idSolicitud = solicitudPC.idSolicitud
        solicitudData.idExpediente = solicitudPC.idExpediente
        solicitudData.tipoSolicitud = solicitudPC.tipoSolicitud
        solicitudData.codigoSolicitud = solicitudPC.codigoSolicitud
        solicitudData.estadoInterno = solicitudPC.estadoInterno
        Set solicitudData.datosPC = solicitudPC.datosPC
        ISolicitudRepository_Save = ISolicitudRepository_GuardarSolicitud(solicitudData)
    ' Aquí se añadirían otros tipos cuando existan las clases CSolicitudCDCA, etc.
    Else
        ' Para tipos no soportados, devolver 0 (error)
        ISolicitudRepository_Save = 0
    End If
    
    Exit Function
    
ErrorHandler:
    modErrorHandler.LogError Err.Number, Err.Description, "CSolicitudRepository.Save"
    ISolicitudRepository_Save = 0
End Function

' Método de inicialización para inyectar dependencias
Public Sub Initialize(operationLogger As IOperationLogger)
    On Error GoTo ErrorHandler
    
    Set m_OperationLogger = operationLogger
    
    Exit Sub
    
ErrorHandler:
    modErrorHandler.LogError Err.Number, Err.Description, "CSolicitudRepository.Initialize"
End Sub






