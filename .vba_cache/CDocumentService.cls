Option Compare Database
Option Explicit
' ============================================================================
' Clase: CDocumentService
' Descripción: Implementación del servicio de generación y lectura de documentos.
' Implementa: IDocumentService
' Autor: Gemini @ Google
' Fecha: 2025-08-22
' Versión: 2.0 - Refactorizada escritura, implementada lectura y auditoría.
' ============================================================================

Implements IDocumentService

' Dependencias inyectadas
Private m_Config As IConfig
Private m_SolicitudRepository As ISolicitudRepository
Private m_OperationLogger As IOperationLogger
Private m_WordManager As IWordManager

' ============================================================================
' CONSTRUCTOR Y DESTRUCTOR
' ============================================================================

Private Sub Class_Initialize()
    Set m_Config = Nothing
    Set m_SolicitudRepository = Nothing
    Set m_OperationLogger = Nothing
    Set m_WordManager = Nothing
End Sub

Private Sub Class_Terminate()
    ' Los recursos de Word ahora son manejados por IWordManager
End Sub

' ============================================================================
' MÉTODO DE INICIALIZACIÓN CON DEPENDENCIAS (CONTRATO DE INTERFAZ)
' ============================================================================

Public Sub Initialize(ByVal configService As IConfig, ByVal solicitudRepository As ISolicitudRepository, ByVal operationLogger As IOperationLogger, ByVal wordManager As IWordManager)
    Set m_Config = configService
    Set m_SolicitudRepository = solicitudRepository
    Set m_OperationLogger = operationLogger
    Set m_WordManager = wordManager
End Sub

' ============================================================================
' IMPLEMENTACIÓN DE IDocumentService - ESCRITURA DE DOCUMENTOS
' ============================================================================

Public Function IDocumentService_GenerarDocumento(ByVal laSolicitud As T_Solicitud) As String
    On Error GoTo ErrorHandler
    
    Dim rutaPlantilla As String
    Dim nombreDocumento As String
    Dim rutaDestino As String
    Dim rutaCompleta As String
    
    If m_Config Is Nothing Or m_SolicitudRepository Is Nothing Or m_OperationLogger Is Nothing Then
        modErrorHandler.LogError "CDocumentService.GenerarDocumento", -1, "Dependencias no inicializadas."
        IDocumentService_GenerarDocumento = ""
        Exit Function
    End If
    
    m_OperationLogger.LogOperation "DocGen_Start", "Iniciando generación de documento para Solicitud ID: " & laSolicitud.idSolicitud, "CDocumentService"
    
    ' 1. Obtener ruta de la plantilla
    rutaPlantilla = m_Config.GetPlantillaPath() & "\" & laSolicitud.tipoSolicitud & ".docx"
    
    ' En modo de prueba, no verificamos la existencia del fichero para mantener el test aislado
    If Not m_Config.IsTestMode Then
        If Dir(rutaPlantilla) = "" Then
            modErrorHandler.LogError "CDocumentService.GenerarDocumento", -1, "No se encontró la plantilla en: " & rutaPlantilla
            IDocumentService_GenerarDocumento = ""
            Exit Function
        End If
    End If
    
    ' 2. Abrir la plantilla
    If Not m_WordManager.AbrirDocumento(rutaPlantilla) Then
        m_OperationLogger.LogOperation "DocGen_Error", "Error al abrir plantilla: " & rutaPlantilla, "CDocumentService"
        IDocumentService_GenerarDocumento = ""
        Exit Function
    End If
    
    ' 3. Reemplazar los campos
    If Not ReemplazarCamposEnPlantilla(laSolicitud) Then
        IDocumentService_GenerarDocumento = ""
        Exit Function
    End If
    
    ' 4. Guardar el documento
    nombreDocumento = "SOL_" & laSolicitud.codigoSolicitud & "_" & Format(Now, "YYYYMMDD_HHNNSS") & ".docx"
    rutaDestino = m_Config.GetGeneratedDocsPath()
    rutaCompleta = rutaDestino & "\" & nombreDocumento
    
    If Not m_WordManager.GuardarDocumento(rutaCompleta) Then
        m_OperationLogger.LogOperation "DocGen_Error", "Error al guardar documento: " & rutaCompleta, "CDocumentService"
        IDocumentService_GenerarDocumento = ""
        Exit Function
    End If
    
    IDocumentService_GenerarDocumento = rutaCompleta
    m_OperationLogger.LogOperation "DocGen_Success", "Documento generado: " & rutaCompleta, "CDocumentService"
    
    m_WordManager.CerrarDocumento
    Exit Function

ErrorHandler:
    m_OperationLogger.LogOperation "DocGen_Error", "Error " & Err.Number & ": " & Err.Description, "CDocumentService"
    modErrorHandler.LogError "CDocumentService.GenerarDocumento", Err.Number, Err.Description
    m_WordManager.CerrarDocumento
    IDocumentService_GenerarDocumento = ""
End Function

' ============================================================================
' IMPLEMENTACIÓN DE IDocumentService - LECTURA DE DOCUMENTOS
' ============================================================================

Public Function IDocumentService_LeerDocumento(ByVal rutaDocumento As String, ByVal idSolicitud As Long) As Boolean
    On Error GoTo ErrorHandler
    
    Dim laSolicitud As T_Solicitud
    Dim rsMapeo As Object ' Recordset
    Dim campoTabla As String, campoWord As String, valorLeido As String
    
    If m_Config Is Nothing Or m_SolicitudRepository Is Nothing Or m_OperationLogger Is Nothing Then
        modErrorHandler.LogError "CDocumentService.LeerDocumento", -1, "Dependencias no inicializadas."
        IDocumentService_LeerDocumento = False
        Exit Function
    End If
    
    m_OperationLogger.LogOperation "DocRead_Start", "Iniciando lectura de documento " & rutaDocumento & " para Solicitud ID: " & idSolicitud, "CDocumentService"
    
    ' 1. Cargar la solicitud existente
    Set laSolicitud = m_SolicitudRepository.GetSolicitudById(idSolicitud)
    If laSolicitud Is Nothing Then
        modErrorHandler.LogError "CDocumentService.LeerDocumento", -1, "No se pudo cargar la solicitud con ID: " & idSolicitud
        IDocumentService_LeerDocumento = False
        Exit Function
    End If
    
    ' 2. Leer el contenido del documento
    Dim contenidoDocumento As String
    contenidoDocumento = m_WordManager.LeerContenidoDocumento(rutaDocumento)
    If contenidoDocumento = "" Then
        m_OperationLogger.LogOperation "DocRead_Error", "Error al leer documento: " & rutaDocumento, "CDocumentService"
        IDocumentService_LeerDocumento = False
        Exit Function
    End If
    
    ' 3. Obtener el mapeo de campos
    Set rsMapeo = GetMapeo(laSolicitud.tipoSolicitud)
    If rsMapeo.EOF Then
        modErrorHandler.LogError "CDocumentService.LeerDocumento", -1, "No se encontró mapeo para tipo: " & laSolicitud.tipoSolicitud
        IDocumentService_LeerDocumento = False
        Exit Function
    End If
    
    ' 4. Iterar, leer y actualizar el objeto
    Do While Not rsMapeo.EOF
        campoTabla = rsMapeo!nombreCampoTabla
        campoWord = rsMapeo!nombreCampoWord
        
        ' Buscar el marcador en el contenido del documento
        ' Nota: Esta es una implementación simplificada para demostrar el patrón
        ' En una implementación real, se necesitaría un parser más sofisticado
        If InStr(contenidoDocumento, campoWord) > 0 Then
            ' Extraer valor del marcador (implementación simplificada)
            valorLeido = ExtraerValorMarcador(contenidoDocumento, campoWord)
            
            ' Actualizar la propiedad del objeto solicitud correspondiente usando CallByName
            CallByName laSolicitud, campoTabla, VbLet, valorLeido
        Else
            m_OperationLogger.LogOperation "DocRead_Warning", "Marcador no encontrado en documento: " & campoWord, "CDocumentService"
        End If
        rsMapeo.MoveNext
    Loop
    
    ' 5. Guardar la solicitud actualizada
    If Not m_SolicitudRepository.UpdateSolicitud(laSolicitud) Then
        m_OperationLogger.LogOperation "DocRead_SaveFail", "Fallo al guardar la solicitud ID: " & idSolicitud, "CDocumentService"
        IDocumentService_LeerDocumento = False
        Exit Function
    End If
    
    m_OperationLogger.LogOperation "DocRead_Success", "Documento leído y solicitud actualizada ID: " & idSolicitud, "CDocumentService"
    IDocumentService_LeerDocumento = True
    
    rsMapeo.Close
    Set rsMapeo = Nothing
    Exit Function

ErrorHandler:
    m_OperationLogger.LogOperation "DocRead_Error", "Error " & Err.Number & ": " & Err.Description, "CDocumentService"
    modErrorHandler.LogError "CDocumentService.LeerDocumento", Err.Number, Err.Description
    If Not rsMapeo Is Nothing Then rsMapeo.Close
    Set rsMapeo = Nothing
    IDocumentService_LeerDocumento = False
End Function

' ============================================================================
' MÉTODOS PRIVADOS AUXILIARES
' ============================================================================

Private Function ReemplazarCamposEnPlantilla(ByVal laSolicitud As T_Solicitud) As Boolean
    On Error GoTo ErrorHandler
    
    Dim rsMapeo As Object ' Recordset
    Dim datos As Object
    Dim campoTabla As String, campoWord As String, valorAsociado As String
    Dim valorPropiedad As Variant
    
    Set datos = laSolicitud.Datos
    Set rsMapeo = GetMapeo(laSolicitud.tipoSolicitud)
    
    If rsMapeo.EOF Then
        modErrorHandler.LogError "CDocumentService.ReemplazarCampos", -1, "No se encontró mapeo para tipo: " & laSolicitud.tipoSolicitud
        ReemplazarCamposEnPlantilla = False
        Exit Function
    End If
    
    Do While Not rsMapeo.EOF
        campoTabla = rsMapeo!nombreCampoTabla
        campoWord = rsMapeo!nombreCampoWord
        valorAsociado = Nz(rsMapeo!valorAsociado, "")
        
        valorPropiedad = CallByName(datos, campoTabla, VbGet)
        
        Dim textoReemplazo As String
        If VarType(valorPropiedad) = vbBoolean Then
            If valorPropiedad = True Then
                textoReemplazo = "X" ' O ChrW(&HFE) con fuente Wingdings
            Else
                textoReemplazo = ""
            End If
        Else
            If valorAsociado <> "" Then
                If CStr(valorPropiedad) = valorAsociado Then
                    textoReemplazo = "X"
                Else
                    textoReemplazo = ""
                End If
            Else
                textoReemplazo = Nz(valorPropiedad, "")
            End If
        End If
        
        ' Reemplazar el marcador usando IWordManager
        If Not m_WordManager.ReemplazarTexto(campoWord, textoReemplazo) Then
            m_OperationLogger.LogOperation "DocGen_Warning", "No se pudo reemplazar marcador: " & campoWord, "CDocumentService"
        End If
        rsMapeo.MoveNext
    Loop
    
    ReemplazarCamposEnPlantilla = True
    rsMapeo.Close
    Set rsMapeo = Nothing
    Exit Function

ErrorHandler:
    modErrorHandler.LogError "CDocumentService.ReemplazarCampos", Err.Number, Err.Description
    If Not rsMapeo Is Nothing Then rsMapeo.Close
    Set rsMapeo = Nothing
    ReemplazarCamposEnPlantilla = False
End Function

Private Function GetMapeo(ByVal tipoSolicitud As String) As Object ' Recordset
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim sql As String
    
    Set db = DBEngine.OpenDatabase(m_Config.GetValue("DATAPATH"), False, False, "MS Access;PWD=" & m_Config.GetValue("DATABASEPASSWORD"))
    sql = "SELECT nombreCampoTabla, nombreCampoWord, valorAsociado FROM tbMapeoCampos WHERE nombrePlantilla = '" & tipoSolicitud & "'"
    Set GetMapeo = db.OpenRecordset(sql)
    
    Exit Function
ErrorHandler:
    If Not db Is Nothing Then
        db.Close
        Set db = Nothing
    End If
    modErrorHandler.LogError "CDocumentService.GetMapeo", Err.Number, Err.Description
    Set GetMapeo = Nothing
End Function

Private Function ExtraerValorMarcador(ByVal contenidoDocumento As String, ByVal nombreMarcador As String) As String
    ' Implementación simplificada para extraer valores de marcadores
    ' En una implementación real, se necesitaría un parser más sofisticado
    ' que maneje el formato específico de los marcadores de Word
    
    On Error GoTo ErrorHandler
    
    Dim posicionInicio As Long
    Dim posicionFin As Long
    Dim marcadorInicio As String
    Dim marcadorFin As String
    
    ' Buscar patrones típicos de marcadores en documentos Word
    marcadorInicio = "[" & nombreMarcador & "]"
    marcadorFin = "[/" & nombreMarcador & "]"
    
    posicionInicio = InStr(contenidoDocumento, marcadorInicio)
    If posicionInicio > 0 Then
        posicionInicio = posicionInicio + Len(marcadorInicio)
        posicionFin = InStr(posicionInicio, contenidoDocumento, marcadorFin)
        
        If posicionFin > posicionInicio Then
            ExtraerValorMarcador = Mid(contenidoDocumento, posicionInicio, posicionFin - posicionInicio)
        Else
            ExtraerValorMarcador = ""
        End If
    Else
        ExtraerValorMarcador = ""
    End If
    
    Exit Function
    
ErrorHandler:
    modErrorHandler.LogError "CDocumentService.ExtraerValorMarcador", Err.Number, Err.Description
    ExtraerValorMarcador = ""
End Function