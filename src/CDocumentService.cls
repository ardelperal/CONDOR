Option Compare Database
Option Explicit
' ============================================================================
' Clase: CDocumentService
' Descripción: Implementación del servicio de generación de documentos
' Implementa: IDocumentService
' Autor: CONDOR-Expert
' Fecha: Enero 2025
' Versión: 1.0 - Implementación inicial
' ============================================================================

Implements IDocumentService

' Variables privadas para manejo de Word y recursos
Private m_WordApp As Object
Private m_WordDoc As Object
Private m_IsWordAppCreated As Boolean
Private m_LastError As String

' Dependencias inyectadas
Private m_ValidationService As IValidationService
Private m_SolicitudRepository As ISolicitudRepository

' ============================================================================
' CONSTRUCTOR Y DESTRUCTOR
' ============================================================================

Private Sub Class_Initialize()
    m_IsWordAppCreated = False
    m_LastError = ""
    Set m_WordApp = Nothing
    Set m_WordDoc = Nothing
    ' Las dependencias se inyectan mediante Initialize
    Set m_ValidationService = Nothing
    Set m_SolicitudRepository = Nothing
End Sub

Private Sub Class_Terminate()
    Call LimpiarRecursos
End Sub

' ============================================================================
' MÉTODO PRIVADO DE LIMPIEZA DE RECURSOS
' ============================================================================

Private Sub LimpiarRecursos()
    On Error Resume Next
    
    ' Cerrar documento de Word si está abierto, sin guardar cambios
    If Not m_WordDoc Is Nothing Then
        m_WordDoc.Close False
        Set m_WordDoc = Nothing
    End If
    
    ' Cerrar aplicación de Word si está abierta
    If Not m_WordApp Is Nothing Then
        If m_IsWordAppCreated Then
            m_WordApp.Quit
        End If
        Set m_WordApp = Nothing
        m_IsWordAppCreated = False
    End If
    
    ' Limpiar error
    m_LastError = ""
End Sub

' ============================================================================
' MÉTODO DE INICIALIZACIÓN CON DEPENDENCIAS
' ============================================================================

Public Sub Initialize(ByVal validationService As IValidationService, ByVal solicitudRepository As ISolicitudRepository)
    Set m_ValidationService = validationService
    Set m_SolicitudRepository = solicitudRepository
End Sub

' ============================================================================
' IMPLEMENTACIÓN DE IDocumentService - MÉTODOS PRINCIPALES
' ============================================================================

Private Function IDocumentService_GenerarDocumento(ByVal idSolicitud As Long) As String
    On Error GoTo ErrorHandler
    
    Dim tipoSolicitud As String
    Dim nombrePlantilla As String
    Dim RutaPlantilla As String
    Dim MapeoCampos As String
    Dim nombreDocumento As String
    Dim rutaDestino As String
    
    ' Validar que las dependencias estén inyectadas
    If m_ValidationService Is Nothing Or m_SolicitudRepository Is Nothing Then
        m_LastError = "Dependencias no inicializadas. Llamar Initialize() primero."
        IDocumentService_GenerarDocumento = ""
        Exit Function
    End If
    
    ' Obtener datos de la solicitud usando el repositorio
    Dim solicitudPC As CSolicitudPC
    Set solicitudPC = m_SolicitudRepository.LoadPC(idSolicitud)
    
    If solicitudPC Is Nothing Then
        m_LastError = "No se pudo cargar la solicitud con ID: " & idSolicitud
        IDocumentService_GenerarDocumento = ""
        Exit Function
    End If
    
    ' Crear objeto T_Solicitud para validación
    Dim solicitudData As T_Solicitud
    Set solicitudData = New T_Solicitud
    solicitudData.idSolicitud = solicitudPC.idSolicitud
    solicitudData.tipoSolicitud = solicitudPC.tipoSolicitud
    solicitudData.estadoInterno = solicitudPC.estadoInterno
    
    ' Validar solicitud usando IValidationService
    Dim MensajeError As String
    If Not m_ValidationService.ValidarSolicitud(solicitudData, MensajeError) Then
        m_LastError = "Solicitud inválida para generar documento: " & MensajeError
        IDocumentService_GenerarDocumento = ""
        Exit Function
    End If
    
    ' Obtener tipo de solicitud del objeto cargado
    tipoSolicitud = solicitudPC.tipoSolicitud
    If tipoSolicitud = "" Then
        m_LastError = "Tipo de solicitud no definido para ID: " & idSolicitud
        IDocumentService_GenerarDocumento = ""
        Exit Function
    End If
    
    ' Obtener plantilla y mapeo
    nombrePlantilla = IDocumentService_ObtenerNombrePlantilla(tipoSolicitud)
    MapeoCampos = IDocumentService_ObtenerMapeoCampos(tipoSolicitud)
    
    ' Construir ruta completa de plantilla
    RutaPlantilla = CurrentProject.path & "\back\recursos\Plantillas\" & nombrePlantilla
    
    ' Validar que la plantilla existe
    If Not IDocumentService_ValidarPlantilla(RutaPlantilla) Then
        m_LastError = "Plantilla no encontrada: " & RutaPlantilla
        IDocumentService_GenerarDocumento = ""
        Exit Function
    End If
    
    ' Abrir plantilla
    If Not IDocumentService_AbrirPlantilla(RutaPlantilla) Then
        m_LastError = "Error al abrir plantilla: " & RutaPlantilla
        IDocumentService_GenerarDocumento = ""
        Exit Function
    End If
    
    ' Reemplazar campos
    If Not IDocumentService_ReemplazarCamposEnPlantilla(idSolicitud, MapeoCampos) Then
        m_LastError = "Error al reemplazar campos en plantilla"
        IDocumentService_GenerarDocumento = ""
        Exit Function
    End If
    
    ' Generar nombre y ruta de destino usando los datos ya cargados
    nombreDocumento = GenerarNombreDocumentoConDatos(solicitudPC)
    rutaDestino = CurrentProject.path & "\back\recursos\Documentos_Generados\"
    
    ' Crear directorio si no existe
    If Dir(rutaDestino, vbDirectory) = "" Then
        MkDir rutaDestino
    End If
    
    ' Guardar documento
    If IDocumentService_GuardarDocumento(rutaDestino, nombreDocumento) Then
        IDocumentService_GenerarDocumento = rutaDestino & nombreDocumento
    Else
        IDocumentService_GenerarDocumento = ""
    End If
    
    Exit Function
    
ErrorHandler:
    m_LastError = "Error en GenerarDocumento: " & Err.Description
    IDocumentService_GenerarDocumento = ""
End Function

Private Function IDocumentService_ObtenerNombrePlantilla(ByVal tipoSolicitud As String) As String
    On Error GoTo ErrorHandler
    
    Dim sql As String
    Dim rs As Object
    
    sql = "SELECT NombrePlantilla FROM Tb_Mapeo_Campos WHERE TipoSolicitud = '" & tipoSolicitud & "'"
    
    Set rs = CurrentDb.OpenRecordset(sql)
    
    If Not rs.EOF Then
        IDocumentService_ObtenerNombrePlantilla = rs("NombrePlantilla").value
    Else
        ' Plantilla por defecto basada en tipo
        IDocumentService_ObtenerNombrePlantilla = tipoSolicitud & "_template.docx"
    End If
    
    rs.Close
    Set rs = Nothing
    
    Exit Function
    
ErrorHandler:
    m_LastError = "Error al obtener nombre de plantilla: " & Err.Description
    IDocumentService_ObtenerNombrePlantilla = tipoSolicitud & "_template.docx"
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
End Function

' Función eliminada: ValidarSolicitudParaDocumento - Ahora se usa IValidationService.ValidarSolicitud

Private Function IDocumentService_ObtenerMapeoCampos(ByVal tipoSolicitud As String) As String
    On Error GoTo ErrorHandler
    
    Dim sql As String
    Dim rs As Object
    
    sql = "SELECT MapeoCampos FROM Tb_Mapeo_Campos WHERE TipoSolicitud = '" & tipoSolicitud & "'"
    
    Set rs = CurrentDb.OpenRecordset(sql)
    
    If Not rs.EOF Then
        IDocumentService_ObtenerMapeoCampos = rs("MapeoCampos").value
    Else
        ' Mapeo por defecto para tipos conocidos
        Select Case tipoSolicitud
            Case "PC"
                IDocumentService_ObtenerMapeoCampos = "NumeroExpediente,FechaSolicitud,TipoSolicitud"
            Case "RAC"
                IDocumentService_ObtenerMapeoCampos = "NumeroExpediente,FechaSolicitud,TipoSolicitud,FechaLimite"
            Case Else
                IDocumentService_ObtenerMapeoCampos = "NumeroExpediente,FechaSolicitud,TipoSolicitud"
        End Select
    End If
    
    rs.Close
    Set rs = Nothing
    
    Exit Function
    
ErrorHandler:
    m_LastError = "Error al obtener mapeo de campos: " & Err.Description
    IDocumentService_ObtenerMapeoCampos = "NumeroExpediente,FechaSolicitud,TipoSolicitud"
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
End Function

' ============================================================================
' IMPLEMENTACIÓN DE IDocumentService - PROCESAMIENTO DE PLANTILLAS
' ============================================================================

Private Function IDocumentService_AbrirPlantilla(ByVal RutaPlantilla As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Crear aplicación Word si no existe
    If m_WordApp Is Nothing Then
        Set m_WordApp = CreateObject("Word.Application")
        m_IsWordAppCreated = True
        m_WordApp.Visible = False
    End If
    
    ' Abrir plantilla
    Set m_WordDoc = m_WordApp.Documents.Open(RutaPlantilla)
    
    IDocumentService_AbrirPlantilla = True
    Exit Function
    
ErrorHandler:
    m_LastError = "Error al abrir plantilla: " & Err.Description
    IDocumentService_AbrirPlantilla = False
End Function

Private Function IDocumentService_ReemplazarCamposEnPlantilla(ByVal idSolicitud As Long, ByVal MapeoCampos As String) As Boolean
    On Error GoTo ErrorHandler
    
    If m_WordDoc Is Nothing Then
        m_LastError = "No hay documento abierto para reemplazar campos"
        IDocumentService_ReemplazarCamposEnPlantilla = False
        Exit Function
    End If
    
    ' Obtener datos de la solicitud
    Dim rs As Object
    Set rs = IDocumentService_ObtenerDatosSolicitud(idSolicitud)
    
    If rs Is Nothing Or rs.EOF Then
        m_LastError = "No se encontraron datos para la solicitud: " & idSolicitud
        IDocumentService_ReemplazarCamposEnPlantilla = False
        Exit Function
    End If
    
    ' Procesar cada campo del mapeo
    Dim campos() As String
    Dim i As Integer
    Dim NombreCampo As String
    Dim valorCampo As String
    Dim marcador As String
    
    campos = Split(MapeoCampos, ",")
    
    For i = 0 To UBound(campos)
        NombreCampo = Trim(campos(i))
        marcador = "{{" & NombreCampo & "}}"
        
        ' Obtener valor del campo
        valorCampo = ObtenerValorCampo(rs, NombreCampo)
        
        ' Reemplazar en el documento
        With m_WordDoc.Range.Find
            .Text = marcador
            .Replacement.Text = valorCampo
            .Execute Replace:=2 ' wdReplaceAll
        End With
    Next i
    
    rs.Close
    Set rs = Nothing
    
    IDocumentService_ReemplazarCamposEnPlantilla = True
    Exit Function
    
ErrorHandler:
    m_LastError = "Error al reemplazar campos: " & Err.Description
    IDocumentService_ReemplazarCamposEnPlantilla = False
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
End Function

Private Function IDocumentService_GuardarDocumento(ByVal rutaDestino As String, ByVal NombreArchivo As String) As Boolean
    On Error GoTo ErrorHandler
    
    If m_WordDoc Is Nothing Then
        m_LastError = "No hay documento abierto para guardar"
        IDocumentService_GuardarDocumento = False
        Exit Function
    End If
    
    Dim rutaCompleta As String
    rutaCompleta = rutaDestino & NombreArchivo
    
    ' Guardar como documento Word
    m_WordDoc.SaveAs2 rutaCompleta
    
    IDocumentService_GuardarDocumento = True
    Exit Function
    
ErrorHandler:
    m_LastError = "Error al guardar documento: " & Err.Description
    IDocumentService_GuardarDocumento = False
End Function

' ============================================================================
' IMPLEMENTACIÓN DE IDocumentService - CONSULTA DE DATOS
' ============================================================================

Private Function IDocumentService_ObtenerDatosSolicitud(ByVal idSolicitud As Long) As Object
    On Error GoTo ErrorHandler
    
    Dim sql As String
    
    ' Consulta principal que une datos de solicitud con datos específicos
    sql = "SELECT s.IdSolicitud, s.NumeroExpediente, s.FechaSolicitud, s.TipoSolicitud, s.EstadoInterno, " & _
          "pc.NombreSolicitante, pc.ApellidosSolicitante, pc.DNI " & _
          "FROM TbSolicitudes s " & _
          "LEFT JOIN TbDatos_PC pc ON s.IdSolicitud = pc.IdSolicitud " & _
          "WHERE s.IdSolicitud = " & idSolicitud
    
    Set IDocumentService_ObtenerDatosSolicitud = CurrentDb.OpenRecordset(sql)
    
    Exit Function
    
ErrorHandler:
    m_LastError = "Error al obtener datos de solicitud: " & Err.Description
    Set IDocumentService_ObtenerDatosSolicitud = Nothing
End Function

Private Function IDocumentService_ObtenerConfiguracionMapeo(ByVal tipoSolicitud As String) As Object
    On Error GoTo ErrorHandler
    
    Dim sql As String
    
    sql = "SELECT * FROM Tb_Mapeo_Campos WHERE TipoSolicitud = '" & tipoSolicitud & "'"
    
    Set IDocumentService_ObtenerConfiguracionMapeo = CurrentDb.OpenRecordset(sql)
    
    Exit Function
    
ErrorHandler:
    m_LastError = "Error al obtener configuración de mapeo: " & Err.Description
    Set IDocumentService_ObtenerConfiguracionMapeo = Nothing
End Function

' ============================================================================
' IMPLEMENTACIÓN DE IDocumentService - UTILIDADES
' ============================================================================

Private Function IDocumentService_ValidarPlantilla(ByVal RutaPlantilla As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Verificar que el archivo existe
    IDocumentService_ValidarPlantilla = (Dir(RutaPlantilla) <> "")
    
    Exit Function
    
ErrorHandler:
    m_LastError = "Error al validar plantilla: " & Err.Description
    IDocumentService_ValidarPlantilla = False
End Function

Private Function IDocumentService_GenerarNombreDocumento(ByVal idSolicitud As Long, ByVal tipoSolicitud As String) As String
    On Error GoTo ErrorHandler
    
    Dim fechaActual As String
    Dim NumeroExpediente As String
    Dim solicitud As T_Solicitud
    
    fechaActual = Format(Date, "yyyy-mm-dd")
    
    ' Usar el repositorio inyectado para obtener la solicitud y su número de expediente
    Set solicitud = m_SolicitudRepository.ObtenerPorId(idSolicitud)
    If Not solicitud Is Nothing Then
        NumeroExpediente = solicitud.NumeroExpediente
    Else
        NumeroExpediente = ""
    End If
    
    If NumeroExpediente <> "" Then
        IDocumentService_GenerarNombreDocumento = "DOC-" & tipoSolicitud & "-" & NumeroExpediente & "-" & fechaActual & ".docx"
    Else
        IDocumentService_GenerarNombreDocumento = "DOC-" & tipoSolicitud & "-" & idSolicitud & "-" & fechaActual & ".docx"
    End If
    
    Exit Function
    
ErrorHandler:
    m_LastError = "Error al generar nombre de documento: " & Err.Description
    IDocumentService_GenerarNombreDocumento = "DOC-" & tipoSolicitud & "-" & idSolicitud & ".docx"
End Function

Private Sub IDocumentService_LimpiarRecursos()
    On Error Resume Next
    
    ' Cerrar documento si está abierto
    If Not m_WordDoc Is Nothing Then
        m_WordDoc.Close False ' No guardar cambios
        Set m_WordDoc = Nothing
    End If
    
    ' Cerrar aplicación Word si la creamos nosotros
    If m_IsWordAppCreated And Not m_WordApp Is Nothing Then
        m_WordApp.Quit
        Set m_WordApp = Nothing
        m_IsWordAppCreated = False
    End If
    
    m_LastError = ""
End Sub

' ============================================================================
' MÉTODOS AUXILIARES PRIVADOS
' ============================================================================

' Función eliminada: ObtenerTipoSolicitud - Ahora se usa ISolicitudRepository.LoadPC

' Función auxiliar para generar nombre usando datos ya cargados
Private Function GenerarNombreDocumentoConDatos(ByVal solicitudPC As CSolicitudPC) As String
    On Error GoTo ErrorHandler
    
    Dim fechaActual As String
    Dim NumeroExpediente As String
    
    fechaActual = Format(Date, "yyyy-mm-dd")
    NumeroExpediente = solicitudPC.idExpediente
    
    If NumeroExpediente <> "" Then
        GenerarNombreDocumentoConDatos = "DOC-" & solicitudPC.tipoSolicitud & "-" & NumeroExpediente & "-" & fechaActual & ".docx"
    Else
        GenerarNombreDocumentoConDatos = "DOC-" & solicitudPC.tipoSolicitud & "-" & solicitudPC.idSolicitud & "-" & fechaActual & ".docx"
    End If
    
    Exit Function
    
ErrorHandler:
    m_LastError = "Error al generar nombre de documento: " & Err.Description
    GenerarNombreDocumentoConDatos = "DOC-" & solicitudPC.tipoSolicitud & "-" & solicitudPC.idSolicitud & ".docx"
End Function

Private Function ObtenerValorCampo(ByRef rs As Object, ByVal NombreCampo As String) As String
    On Error GoTo ErrorHandler
    
    Dim Valor As Variant
    
    ' Intentar obtener el valor del campo
    Valor = rs(NombreCampo).value
    
    If IsNull(Valor) Then
        ObtenerValorCampo = ""
    Else
        ' Formatear según el tipo de campo
        Select Case NombreCampo
            Case "FechaSolicitud", "FechaLimite"
                ObtenerValorCampo = Format(Valor, "dd/mm/yyyy")
            Case Else
                ObtenerValorCampo = CStr(Valor)
        End Select
    End If
    
    Exit Function
    
ErrorHandler:
    ObtenerValorCampo = "[Error: " & NombreCampo & "]"
End Function

' ============================================================================
' MÉTODOS PÚBLICOS ADICIONALES
' ============================================================================

Public Function GetLastError() As String
    GetLastError = m_LastError
End Function

Public Sub ClearLastError()
    m_LastError = ""
End Sub













