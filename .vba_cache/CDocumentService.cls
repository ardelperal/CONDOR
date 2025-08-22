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

' Variables privadas para manejo de Word
Private m_WordApp As Object
Private m_WordDoc As Object
Private m_IsWordAppCreated As Boolean

' ============================================================================
' CONSTRUCTOR Y DESTRUCTOR
' ============================================================================

Private Sub Class_Initialize()
    m_IsWordAppCreated = False
    Set m_WordApp = Nothing
    Set m_WordDoc = Nothing
    Set m_Config = Nothing
    Set m_SolicitudRepository = Nothing
    Set m_OperationLogger = Nothing
End Sub

Private Sub Class_Terminate()
    Call LimpiarRecursos
End Sub

' ============================================================================
' MÉTODO DE INICIALIZACIÓN CON DEPENDENCIAS (CONTRATO DE INTERFAZ)
' ============================================================================

Public Sub Initialize(ByVal configService As IConfig, ByVal solicitudRepository As ISolicitudRepository, ByVal operationLogger As IOperationLogger)
    Set m_Config = configService
    Set m_SolicitudRepository = solicitudRepository
    Set m_OperationLogger = operationLogger
End Sub

' ============================================================================
' IMPLEMENTACIÓN DE IDocumentService - ESCRITURA DE DOCUMENTOS
' ============================================================================

Public Function IDocumentService_GenerarDocumento(ByVal laSolicitud As ISolicitud) As String
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
    If Not AbrirWordDocumento(rutaPlantilla) Then
        ' El error ya fue manejado dentro de AbrirWordDocumento
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
    
    If Not GuardarWordDocumento(rutaCompleta) Then
        IDocumentService_GenerarDocumento = ""
        Exit Function
    End If
    
    IDocumentService_GenerarDocumento = rutaCompleta
    m_OperationLogger.LogOperation "DocGen_Success", "Documento generado: " & rutaCompleta, "CDocumentService"
    
    Call LimpiarRecursos
    Exit Function

ErrorHandler:
    m_OperationLogger.LogOperation "DocGen_Error", "Error " & Err.Number & ": " & Err.Description, "CDocumentService"
    modErrorHandler.LogError "CDocumentService.GenerarDocumento", Err.Number, Err.Description
    Call LimpiarRecursos
    IDocumentService_GenerarDocumento = ""
End Function

' ============================================================================
' IMPLEMENTACIÓN DE IDocumentService - LECTURA DE DOCUMENTOS
' ============================================================================

Public Function IDocumentService_LeerDocumento(ByVal rutaDocumento As String, ByVal idSolicitud As Long) As Boolean
    On Error GoTo ErrorHandler
    
    Dim laSolicitud As ISolicitud
    Dim rsMapeo As Object ' Recordset
    Dim campoTabla As String, campoWord As String, valorLeido As String
    
    If m_Config Is Nothing Or m_SolicitudRepository Is Nothing Or m_OperationLogger Is Nothing Then
        modErrorHandler.LogError "CDocumentService.LeerDocumento", -1, "Dependencias no inicializadas."
        IDocumentService_LeerDocumento = False
        Exit Function
    End If
    
    m_OperationLogger.LogOperation "DocRead_Start", "Iniciando lectura de documento " & rutaDocumento & " para Solicitud ID: " & idSolicitud, "CDocumentService"
    
    ' 1. Cargar la solicitud existente
    Set laSolicitud = m_SolicitudRepository.Load(idSolicitud)
    If laSolicitud Is Nothing Then
        modErrorHandler.LogError "CDocumentService.LeerDocumento", -1, "No se pudo cargar la solicitud con ID: " & idSolicitud
        IDocumentService_LeerDocumento = False
        Exit Function
    End If
    
    ' 2. Abrir el documento para lectura
    If Not AbrirWordDocumento(rutaDocumento) Then
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
        
        If Not m_WordDoc.Bookmarks.Exists(campoWord) Then
            ' Opcional: registrar si un marcador no se encuentra, pero no detener el proceso
            m_OperationLogger.LogOperation "DocRead_Warning", "Marcador no encontrado en documento: " & campoWord, "CDocumentService"
        Else
            valorLeido = Trim(m_WordDoc.Bookmarks(campoWord).Range.Text)
            
            ' Actualizar la propiedad del objeto de datos correspondiente usando CallByName
            CallByName laSolicitud.Datos, campoTabla, VbLet, valorLeido
        End If
        rsMapeo.MoveNext
    Loop
    
    ' 5. Guardar la solicitud actualizada
    If Not m_SolicitudRepository.Save(laSolicitud) Then
        m_OperationLogger.LogOperation "DocRead_SaveFail", "Fallo al guardar la solicitud ID: " & idSolicitud, "CDocumentService"
        IDocumentService_LeerDocumento = False
        Exit Function
    End If
    
    m_OperationLogger.LogOperation "DocRead_Success", "Documento leído y solicitud actualizada ID: " & idSolicitud, "CDocumentService"
    IDocumentService_LeerDocumento = True
    
    Call LimpiarRecursos
    rsMapeo.Close
    Set rsMapeo = Nothing
    Exit Function

ErrorHandler:
    m_OperationLogger.LogOperation "DocRead_Error", "Error " & Err.Number & ": " & Err.Description, "CDocumentService"
    modErrorHandler.LogError "CDocumentService.LeerDocumento", Err.Number, Err.Description
    Call LimpiarRecursos
    If Not rsMapeo Is Nothing Then rsMapeo.Close
    Set rsMapeo = Nothing
    IDocumentService_LeerDocumento = False
End Function

' ============================================================================
' MÉTODOS PRIVADOS AUXILIARES
' ============================================================================

Private Function ReemplazarCamposEnPlantilla(ByVal laSolicitud As ISolicitud) As Boolean
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
        
        If m_WordDoc.Bookmarks.Exists(campoWord) Then
            valorPropiedad = CallByName(datos, campoTabla, VbGet)
            
            If VarType(valorPropiedad) = vbBoolean Then
                If valorPropiedad = True Then
                    m_WordDoc.Bookmarks(campoWord).Range.Text = "X" ' O ChrW(&HFE) con fuente Wingdings
                Else
                    m_WordDoc.Bookmarks(campoWord).Range.Text = ""
                End If
            Else
                If valorAsociado <> "" Then
                    If CStr(valorPropiedad) = valorAsociado Then
                        m_WordDoc.Bookmarks(campoWord).Range.Text = "X"
                    Else
                        m_WordDoc.Bookmarks(campoWord).Range.Text = ""
                    End If
                Else
                    m_WordDoc.Bookmarks(campoWord).Range.Text = Nz(valorPropiedad, "")
                End If
            End If
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
    
    Dim sql As String
    sql = "SELECT nombreCampoTabla, nombreCampoWord, valorAsociado FROM tbMapeoCampos WHERE nombrePlantilla = '" & tipoSolicitud & "'"
    Set GetMapeo = CurrentDb.OpenRecordset(sql)
    
    Exit Function
ErrorHandler:
    modErrorHandler.LogError "CDocumentService.GetMapeo", Err.Number, Err.Description
    Set GetMapeo = Nothing
End Function

Private Function AbrirWordDocumento(ByVal ruta As String) As Boolean
    On Error GoTo ErrorHandler
    
    If m_WordApp Is Nothing Then
        Set m_WordApp = CreateObject("Word.Application")
        m_IsWordAppCreated = True
        m_WordApp.Visible = False
    End If
    
    Set m_WordDoc = m_WordApp.Documents.Open(ruta)
    AbrirWordDocumento = True
    Exit Function

ErrorHandler:
    modErrorHandler.LogError "CDocumentService.AbrirWordDocumento", Err.Number, Err.Description
    Call LimpiarRecursos
    AbrirWordDocumento = False
End Function

Private Function GuardarWordDocumento(ByVal rutaCompleta As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Crear directorio si no existe
    Dim rutaDestino As String
    rutaDestino = Left(rutaCompleta, InStrRev(rutaCompleta, "\"))
    If Dir(rutaDestino, vbDirectory) = "" Then
        MkDir rutaDestino
    End If
    
    m_WordDoc.SaveAs2 rutaCompleta
    GuardarWordDocumento = True
    Exit Function

ErrorHandler:
    modErrorHandler.LogError "CDocumentService.GuardarWordDocumento", Err.Number, Err.Description
    GuardarWordDocumento = False
End Function

Private Sub LimpiarRecursos()
    On Error Resume Next
    If Not m_WordDoc Is Nothing Then
        m_WordDoc.Close False ' Cerrar sin guardar
        Set m_WordDoc = Nothing
    End If
    If m_IsWordAppCreated And Not m_WordApp Is Nothing Then
        m_WordApp.Quit
        Set m_WordApp = Nothing
    End If
    m_IsWordAppCreated = False
End Sub