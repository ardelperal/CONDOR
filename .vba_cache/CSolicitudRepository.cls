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
    DBEngine.BeginTrans
    
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
    DBEngine.CommitTrans
    
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
    DBEngine.Rollback
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
    DBEngine.BeginTrans
    
    On Error GoTo ErrorHandler
    
    ' Eliminar datos específicos primero (integridad referencial)
    db.Execute "DELETE FROM tbDatosPC WHERE idSolicitud = " & idSolicitud
    db.Execute "DELETE FROM tbDatosCDCA WHERE idSolicitud = " & idSolicitud
    db.Execute "DELETE FROM tbDatosCDCASUB WHERE idSolicitud = " & idSolicitud
    
    ' Eliminar solicitud principal
    db.Execute "DELETE FROM tbSolicitud WHERE idSolicitud = " & idSolicitud
    
    DBEngine.CommitTrans
    
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
    DBEngine.Rollback
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
    
    ' Campos básicos
    rs("refContratoInspeccionOficial") = datosPC.refContratoInspeccionOficial
    rs("refSuministrador") = datosPC.refSuministrador
    rs("suministradorNombreDir") = datosPC.suministradorNombreDir
    rs("objetoContrato") = datosPC.objetoContrato
    rs("descripcionMaterialAfectado") = datosPC.descripcionMaterialAfectado
    rs("numPlanoEspecificacion") = datosPC.numPlanoEspecificacion
    rs("descripcionPropuestaCambio") = datosPC.descripcionPropuestaCambio
    rs("descripcionPropuestaCambioCont") = datosPC.descripcionPropuestaCambioCont
    
    ' Motivos (campos booleanos)
    rs("motivoMejoraTecnica") = datosPC.motivoMejoraTecnica
    rs("motivoReduccionCoste") = datosPC.motivoReduccionCoste
    rs("motivoMejoraPlazo") = datosPC.motivoMejoraPlazo
    rs("motivoSolucionProblema") = datosPC.motivoSolucionProblema
    rs("motivoActualizacionNormativa") = datosPC.motivoActualizacionNormativa
    rs("motivoObsolescencia") = datosPC.motivoObsolescencia
    rs("motivoOtros") = datosPC.motivoOtros
    
    ' Incidencias
    rs("incidenciaCalidad") = datosPC.incidenciaCalidad
    rs("incidenciaSeguridad") = datosPC.incidenciaSeguridad
    rs("incidenciaFuncional") = datosPC.incidenciaFuncional
    rs("incidenciaOperacional") = datosPC.incidenciaOperacional
    rs("incidenciaMantenimiento") = datosPC.incidenciaMantenimiento
    
    ' Cambios que afectan material
    rs("cambioAfectaMaterialClaseSeguridad") = datosPC.cambioAfectaMaterialClaseSeguridad
    rs("cambioAfectaMaterialImportanteSeguridad") = datosPC.cambioAfectaMaterialImportanteSeguridad
    rs("cambioAfectaMaterialBasico") = datosPC.cambioAfectaMaterialBasico
    
    ' Campos de firma y autorización
    rs("nombreRepSuministrador") = datosPC.nombreRepSuministrador
    rs("cargoRepSuministrador") = datosPC.cargoRepSuministrador
    rs("fechaFirmaSuministrador") = datosPC.fechaFirmaSuministrador
    rs("identificacionAutoridadDiseno") = datosPC.identificacionAutoridadDiseno
    rs("esSuministradorAD") = datosPC.esSuministradorAD
    
    ' Campos RAC
    rs("racRef") = datosPC.racRef
    rs("racCodigo") = datosPC.racCodigo
    rs("observacionesRAC") = datosPC.observacionesRAC
    rs("fechaFirmaRAC") = datosPC.fechaFirmaRAC
    
    ' Campos de decisión final
    rs("decisionFinal") = datosPC.decisionFinal
    rs("observacionesFinales") = datosPC.observacionesFinales
    rs("fechaFirmaDecisionFinal") = datosPC.fechaFirmaDecisionFinal
    rs("cargoFirmanteFinal") = datosPC.cargoFirmanteFinal
    
    ' Campos de impacto y justificación
    rs("justificacion") = datosPC.justificacion
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
    
    ' Campos básicos
    rs("refSuministrador") = datosCDCA.refSuministrador
    rs("numContrato") = datosCDCA.numContrato
    rs("identificacionMaterial") = datosCDCA.identificacionMaterial
    rs("numPlanoEspecificacion") = datosCDCA.numPlanoEspecificacion
    rs("cantidadPeriodo") = datosCDCA.cantidadPeriodo
    rs("numSerieLote") = datosCDCA.numSerieLote
    
    ' Descripción del impacto
    rs("descripcionImpactoNC") = datosCDCA.descripcionImpactoNC
    rs("descripcionImpactoNCCont") = datosCDCA.descripcionImpactoNCCont
    
    ' Referencias y clasificación
    rs("refDesviacionesPrevias") = datosCDCA.refDesviacionesPrevias
    rs("causaNC") = datosCDCA.causaNC
    rs("impactoCoste") = datosCDCA.impactoCoste
    rs("clasificacionNC") = datosCDCA.clasificacionNC
    rs("requiereModificacionContrato") = datosCDCA.requiereModificacionContrato
    rs("efectoFechaEntrega") = datosCDCA.efectoFechaEntrega
    
    ' Autoridad de diseño
    rs("identificacionAutoridadDiseno") = datosCDCA.identificacionAutoridadDiseno
    rs("esSuministradorAD") = datosCDCA.esSuministradorAD
    
    ' Campos RAC
    rs("racRef") = datosCDCA.racRef
    rs("racCodigo") = datosCDCA.racCodigo
    rs("observacionesRAC") = datosCDCA.observacionesRAC
    If datosCDCA.fechaFirmaRAC <> 0 Then rs("fechaFirmaRAC") = datosCDCA.fechaFirmaRAC
    
    ' Decisión final
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
    
    ' Campos básicos
    rs("refSuministrador") = datosCDCASUB.refSuministrador
    rs("refSubSuministrador") = datosCDCASUB.refSubSuministrador
    rs("suministradorPrincipalNombreDir") = datosCDCASUB.suministradorPrincipalNombreDir
    rs("subSuministradorNombreDir") = datosCDCASUB.subSuministradorNombreDir
    rs("numContrato") = datosCDCASUB.numContrato
    rs("identificacionMaterial") = datosCDCASUB.identificacionMaterial
    rs("numPlanoEspecificacion") = datosCDCASUB.numPlanoEspecificacion
    rs("cantidadPeriodo") = datosCDCASUB.cantidadPeriodo
    rs("numSerieLote") = datosCDCASUB.numSerieLote
    
    ' Descripción del impacto
    rs("descripcionImpactoNC") = datosCDCASUB.descripcionImpactoNC
    rs("descripcionImpactoNCCont") = datosCDCASUB.descripcionImpactoNCCont
    
    ' Referencias y clasificación
    rs("refDesviacionesPrevias") = datosCDCASUB.refDesviacionesPrevias
    rs("causaNC") = datosCDCASUB.causaNC
    rs("impactoCoste") = datosCDCASUB.impactoCoste
    rs("clasificacionNC") = datosCDCASUB.clasificacionNC
    rs("requiereModificacionContrato") = datosCDCASUB.requiereModificacionContrato
    rs("efectoFechaEntrega") = datosCDCASUB.efectoFechaEntrega
    
    ' Autoridad de diseño
    rs("identificacionAutoridadDiseno") = datosCDCASUB.identificacionAutoridadDiseno
    rs("esSubSuministradorAD") = datosCDCASUB.esSubSuministradorAD
    rs("nombreRepSubSuministrador") = datosCDCASUB.nombreRepSubSuministrador
    
    ' Campos RAC
    rs("racRef") = datosCDCASUB.racRef
    rs("racCodigo") = datosCDCASUB.racCodigo
    rs("observacionesRAC") = datosCDCASUB.observacionesRAC
    If datosCDCASUB.fechaFirmaRAC <> 0 Then rs("fechaFirmaRAC") = datosCDCASUB.fechaFirmaRAC
    
    ' Campos de impacto en material
    rs("afectaPrestaciones") = datosCDCASUB.afectaPrestaciones
    rs("afectaSeguridad") = datosCDCASUB.afectaSeguridad
    rs("afectaFiabilidad") = datosCDCASUB.afectaFiabilidad
    rs("afectaVidaUtil") = datosCDCASUB.afectaVidaUtil
    rs("afectaMedioambiente") = datosCDCASUB.afectaMedioambiente
    rs("afectaIntercambiabilidad") = datosCDCASUB.afectaIntercambiabilidad
    rs("afectaMantenibilidad") = datosCDCASUB.afectaMantenibilidad
    rs("afectaApariencia") = datosCDCASUB.afectaApariencia
    rs("afectaOtros") = datosCDCASUB.afectaOtros
    
    ' Decisión del suministrador principal
    rs("decisionSuministradorPrincipal") = datosCDCASUB.decisionSuministradorPrincipal
    rs("obsSuministradorPrincipal") = datosCDCASUB.obsSuministradorPrincipal
    If datosCDCASUB.fechaFirmaSuministradorPrincipal <> 0 Then rs("fechaFirmaSuministradorPrincipal") = datosCDCASUB.fechaFirmaSuministradorPrincipal
    rs("firmaSuministradorPrincipalNombreCargo") = datosCDCASUB.firmaSuministradorPrincipalNombreCargo
    
    ' RAC Delegador
    rs("obsRACDelegador") = datosCDCASUB.obsRACDelegador
    If datosCDCASUB.fechaFirmaRACDelegador <> 0 Then rs("fechaFirmaRACDelegador") = datosCDCASUB.fechaFirmaRACDelegador
    
    ' Decisión final
    rs("decisionFinal") = datosCDCASUB.decisionFinal
    rs("observacionesFinales") = datosCDCASUB.observacionesFinales
    If datosCDCASUB.fechaFirmaDecisionFinal <> 0 Then rs("fechaFirmaDecisionFinal") = datosCDCASUB.fechaFirmaDecisionFinal
    rs("cargoFirmanteFinal") = datosCDCASUB.cargoFirmanteFinal
    
    rs.Update
    rs.Close
End Sub

Private Function CargarDatosPC(ByVal db As DAO.Database, ByVal idSolicitud As Long) As T_Datos_PC
    Dim rs As DAO.Recordset
    Dim datosPC As T_Datos_PC
    
    Set rs = db.OpenRecordset("SELECT * FROM tbDatosPC WHERE idSolicitud = " & idSolicitud, dbOpenSnapshot)
    
    If Not rs.EOF Then
        Set datosPC = New T_Datos_PC
        
        ' Campos básicos
        datosPC.idDatosPC = rs("idDatosPC")
        datosPC.idSolicitud = rs("idSolicitud")
        datosPC.refContratoInspeccionOficial = rs("refContratoInspeccionOficial") & ""
        datosPC.refSuministrador = rs("refSuministrador") & ""
        datosPC.suministradorNombreDir = rs("suministradorNombreDir") & ""
        datosPC.objetoContrato = rs("objetoContrato") & ""
        datosPC.descripcionMaterialAfectado = rs("descripcionMaterialAfectado") & ""
        datosPC.numPlanoEspecificacion = rs("numPlanoEspecificacion") & ""
        datosPC.descripcionPropuestaCambio = rs("descripcionPropuestaCambio") & ""
        datosPC.descripcionPropuestaCambioCont = rs("descripcionPropuestaCambioCont") & ""
        
        ' Motivos (campos booleanos)
        datosPC.motivoMejoraTecnica = rs("motivoMejoraTecnica")
        datosPC.motivoReduccionCoste = rs("motivoReduccionCoste")
        datosPC.motivoMejoraPlazo = rs("motivoMejoraPlazo")
        datosPC.motivoSolucionProblema = rs("motivoSolucionProblema")
        datosPC.motivoActualizacionNormativa = rs("motivoActualizacionNormativa")
        datosPC.motivoObsolescencia = rs("motivoObsolescencia")
        datosPC.motivoOtros = rs("motivoOtros")
        
        ' Incidencias
        datosPC.incidenciaCalidad = rs("incidenciaCalidad") & ""
        datosPC.incidenciaSeguridad = rs("incidenciaSeguridad") & ""
        datosPC.incidenciaFuncional = rs("incidenciaFuncional") & ""
        datosPC.incidenciaOperacional = rs("incidenciaOperacional") & ""
        datosPC.incidenciaMantenimiento = rs("incidenciaMantenimiento") & ""
        
        ' Cambios que afectan material
        datosPC.cambioAfectaMaterialClaseSeguridad = rs("cambioAfectaMaterialClaseSeguridad")
        datosPC.cambioAfectaMaterialImportanteSeguridad = rs("cambioAfectaMaterialImportanteSeguridad")
        datosPC.cambioAfectaMaterialBasico = rs("cambioAfectaMaterialBasico")
        
        ' Campos de firma y autorización
        datosPC.nombreRepSuministrador = rs("nombreRepSuministrador") & ""
        datosPC.cargoRepSuministrador = rs("cargoRepSuministrador") & ""
        If Not IsNull(rs("fechaFirmaSuministrador")) Then datosPC.fechaFirmaSuministrador = rs("fechaFirmaSuministrador")
        datosPC.identificacionAutoridadDiseno = rs("identificacionAutoridadDiseno") & ""
        datosPC.esSuministradorAD = rs("esSuministradorAD")
        
        ' Campos RAC
        datosPC.racRef = rs("racRef") & ""
        datosPC.racCodigo = rs("racCodigo") & ""
        datosPC.observacionesRAC = rs("observacionesRAC") & ""
        If Not IsNull(rs("fechaFirmaRAC")) Then datosPC.fechaFirmaRAC = rs("fechaFirmaRAC")
        
        ' Campos de decisión final
        datosPC.decisionFinal = rs("decisionFinal") & ""
        datosPC.observacionesFinales = rs("observacionesFinales") & ""
        If Not IsNull(rs("fechaFirmaDecisionFinal")) Then datosPC.fechaFirmaDecisionFinal = rs("fechaFirmaDecisionFinal")
        datosPC.cargoFirmanteFinal = rs("cargoFirmanteFinal") & ""
        
        ' Campos de impacto y justificación
        datosPC.justificacion = rs("justificacion") & ""
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
        
        ' Campos básicos
        datosCDCA.idDatosCDCA = rs("idDatosCDCA")
        datosCDCA.idSolicitud = rs("idSolicitud")
        datosCDCA.refSuministrador = rs("refSuministrador") & ""
        datosCDCA.numContrato = rs("numContrato") & ""
        datosCDCA.identificacionMaterial = rs("identificacionMaterial") & ""
        datosCDCA.numPlanoEspecificacion = rs("numPlanoEspecificacion") & ""
        datosCDCA.cantidadPeriodo = rs("cantidadPeriodo")
        datosCDCA.numSerieLote = rs("numSerieLote") & ""
        
        ' Descripción del impacto
        datosCDCA.descripcionImpactoNC = rs("descripcionImpactoNC") & ""
        datosCDCA.descripcionImpactoNCCont = rs("descripcionImpactoNCCont") & ""
        
        ' Referencias y clasificación
        datosCDCA.refDesviacionesPrevias = rs("refDesviacionesPrevias") & ""
        datosCDCA.causaNC = rs("causaNC") & ""
        datosCDCA.impactoCoste = rs("impactoCoste") & ""
        datosCDCA.clasificacionNC = rs("clasificacionNC") & ""
        datosCDCA.requiereModificacionContrato = rs("requiereModificacionContrato")
        datosCDCA.efectoFechaEntrega = rs("efectoFechaEntrega") & ""
        
        ' Autoridad de diseño
        datosCDCA.identificacionAutoridadDiseno = rs("identificacionAutoridadDiseno") & ""
        datosCDCA.esSuministradorAD = rs("esSuministradorAD")
        
        ' Campos RAC
        datosCDCA.racRef = rs("racRef") & ""
        datosCDCA.racCodigo = rs("racCodigo") & ""
        datosCDCA.observacionesRAC = rs("observacionesRAC") & ""
        If Not IsNull(rs("fechaFirmaRAC")) Then datosCDCA.fechaFirmaRAC = rs("fechaFirmaRAC")
        
        ' Decisión final
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
        
        ' Campos básicos
        datosCDCASUB.idDatosCDCASUB = rs("idDatosCDCASUB")
        datosCDCASUB.idSolicitud = rs("idSolicitud")
        datosCDCASUB.refSuministrador = rs("refSuministrador") & ""
        datosCDCASUB.refSubSuministrador = rs("refSubSuministrador") & ""
        datosCDCASUB.suministradorPrincipalNombreDir = rs("suministradorPrincipalNombreDir") & ""
        datosCDCASUB.subSuministradorNombreDir = rs("subSuministradorNombreDir") & ""
        datosCDCASUB.numContrato = rs("numContrato") & ""
        datosCDCASUB.identificacionMaterial = rs("identificacionMaterial") & ""
        datosCDCASUB.numPlanoEspecificacion = rs("numPlanoEspecificacion") & ""
        datosCDCASUB.cantidadPeriodo = rs("cantidadPeriodo")
        datosCDCASUB.numSerieLote = rs("numSerieLote") & ""
        
        ' Descripción del impacto
        datosCDCASUB.descripcionImpactoNC = rs("descripcionImpactoNC") & ""
        datosCDCASUB.descripcionImpactoNCCont = rs("descripcionImpactoNCCont") & ""
        
        ' Referencias y clasificación
        datosCDCASUB.refDesviacionesPrevias = rs("refDesviacionesPrevias") & ""
        datosCDCASUB.causaNC = rs("causaNC") & ""
        datosCDCASUB.impactoCoste = rs("impactoCoste") & ""
        datosCDCASUB.clasificacionNC = rs("clasificacionNC") & ""
        datosCDCASUB.requiereModificacionContrato = rs("requiereModificacionContrato")
        datosCDCASUB.efectoFechaEntrega = rs("efectoFechaEntrega") & ""
        
        ' Autoridad de diseño
        datosCDCASUB.identificacionAutoridadDiseno = rs("identificacionAutoridadDiseno") & ""
        datosCDCASUB.esSubSuministradorAD = rs("esSubSuministradorAD")
        datosCDCASUB.nombreRepSubSuministrador = rs("nombreRepSubSuministrador") & ""
        
        ' Campos RAC
        datosCDCASUB.racRef = rs("racRef") & ""
        datosCDCASUB.racCodigo = rs("racCodigo") & ""
        datosCDCASUB.observacionesRAC = rs("observacionesRAC") & ""
        If Not IsNull(rs("fechaFirmaRAC")) Then datosCDCASUB.fechaFirmaRAC = rs("fechaFirmaRAC")
        
        ' Campos de impacto en material
        datosCDCASUB.afectaPrestaciones = rs("afectaPrestaciones")
        datosCDCASUB.afectaSeguridad = rs("afectaSeguridad")
        datosCDCASUB.afectaFiabilidad = rs("afectaFiabilidad")
        datosCDCASUB.afectaVidaUtil = rs("afectaVidaUtil")
        datosCDCASUB.afectaMedioambiente = rs("afectaMedioambiente")
        datosCDCASUB.afectaIntercambiabilidad = rs("afectaIntercambiabilidad")
        datosCDCASUB.afectaMantenibilidad = rs("afectaMantenibilidad")
        datosCDCASUB.afectaApariencia = rs("afectaApariencia")
        datosCDCASUB.afectaOtros = rs("afectaOtros")
        
        ' Decisión del suministrador principal
        datosCDCASUB.decisionSuministradorPrincipal = rs("decisionSuministradorPrincipal") & ""
        datosCDCASUB.obsSuministradorPrincipal = rs("obsSuministradorPrincipal") & ""
        If Not IsNull(rs("fechaFirmaSuministradorPrincipal")) Then datosCDCASUB.fechaFirmaSuministradorPrincipal = rs("fechaFirmaSuministradorPrincipal")
        datosCDCASUB.firmaSuministradorPrincipalNombreCargo = rs("firmaSuministradorPrincipalNombreCargo") & ""
        
        ' RAC Delegador
        datosCDCASUB.obsRACDelegador = rs("obsRACDelegador") & ""
        If Not IsNull(rs("fechaFirmaRACDelegador")) Then datosCDCASUB.fechaFirmaRACDelegador = rs("fechaFirmaRACDelegador")
        
        ' Decisión final
        datosCDCASUB.decisionFinal = rs("decisionFinal") & ""
        datosCDCASUB.observacionesFinales = rs("observacionesFinales") & ""
        If Not IsNull(rs("fechaFirmaDecisionFinal")) Then datosCDCASUB.fechaFirmaDecisionFinal = rs("fechaFirmaDecisionFinal")
        datosCDCASUB.cargoFirmanteFinal = rs("cargoFirmanteFinal") & ""
    End If
    
    rs.Close
    Set CargarDatosCDCASUB = datosCDCASUB
End Function



' Método de inicialización para inyectar dependencias
Public Sub Initialize(operationLogger As IOperationLogger)
    On Error GoTo ErrorHandler
    
    Set m_OperationLogger = operationLogger
    
    Exit Sub
    
ErrorHandler:
    modErrorHandler.LogError Err.Number, Err.Description, "CSolicitudRepository.Initialize"
End Sub






