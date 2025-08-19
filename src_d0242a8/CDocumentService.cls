VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDocumentService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' ============================================================================
' Clase: CDocumentService
' Descripci├│n: Implementaci├│n del servicio de generaci├│n de documentos
' Implementa: IDocumentService
' Autor: CONDOR-Expert
' Fecha: Enero 2025
' Versi├│n: 1.0 - Implementaci├│n inicial
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
' M├ëTODO PRIVADO DE LIMPIEZA DE RECURSOS
' ============================================================================

Private Sub LimpiarRecursos()
    On Error Resume Next
    
    ' Cerrar documento de Word si est├í abierto, sin guardar cambios
    If Not m_WordDoc Is Nothing Then
        m_WordDoc.Close False
        Set m_WordDoc = Nothing
    End If
    
    ' Cerrar aplicaci├│n de Word si est├í abierta
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
' M├ëTODO DE INICIALIZACI├ôN CON DEPENDENCIAS
' ============================================================================

Public Sub Initialize(ByVal validationService As IValidationService, ByVal solicitudRepository As ISolicitudRepository)
    Set m_ValidationService = validationService
    Set m_SolicitudRepository = solicitudRepository
End Sub

' ============================================================================
' IMPLEMENTACI├ôN DE IDocumentService - M├ëTODOS PRINCIPALES
' ============================================================================

Private Function IDocumentService_GenerarDocumento(ByVal IdSolicitud As Long) As String
    On Error GoTo ErrorHandler
    
    Dim tipoSolicitud As String
    Dim nombrePlantilla As String
    Dim rutaPlantilla As String
    Dim mapeoCampos As String
    Dim nombreDocumento As String
    Dim rutaDestino As String
    
    ' Validar que las dependencias est├®n inyectadas
    If m_ValidationService Is Nothing Or m_SolicitudRepository Is Nothing Then
        m_LastError = "Dependencias no inicializadas. Llamar Initialize() primero."
        IDocumentService_GenerarDocumento = ""
        Exit Function
    End If
    
    ' Obtener datos de la solicitud usando el repositorio
    Dim solicitudPC As CSolicitudPC
    Set solicitudPC = m_SolicitudRepository.LoadPC(IdSolicitud)
    
    If solicitudPC Is Nothing Then
        m_LastError = "No se pudo cargar la solicitud con ID: " & IdSolicitud
        IDocumentService_GenerarDocumento = ""
        Exit Function
    End If
    
    ' Crear objeto T_Solicitud para validaci├│n
    Dim solicitudData As T_Solicitud
    Set solicitudData = New T_Solicitud
    solicitudData.idSolicitud = solicitudPC.idSolicitud
    solicitudData.TipoSolicitud = solicitudPC.TipoSolicitud
    solicitudData.EstadoInterno = solicitudPC.EstadoInterno
    
    ' Validar solicitud usando IValidationService
    Dim mensajeError As String
    If Not m_ValidationService.ValidarSolicitud(solicitudData, mensajeError) Then
        m_LastError = "Solicitud inv├ílida para generar documento: " & mensajeError
        IDocumentService_GenerarDocumento = ""
        Exit Function
    End If
    
    ' Obtener tipo de solicitud del objeto cargado
    tipoSolicitud = solicitudPC.TipoSolicitud
    If tipoSolicitud = "" Then
        m_LastError = "Tipo de solicitud no definido para ID: " & IdSolicitud
        IDocumentService_GenerarDocumento = ""
        Exit Function
    End If
    
    ' Obtener plantilla y mapeo
    nombrePlantilla = IDocumentService_ObtenerNombrePlantilla(tipoSolicitud)
    mapeoCampos = IDocumentService_ObtenerMapeoCampos(tipoSolicitud)
    
    ' Construir ruta completa de plantilla
    rutaPlantilla = CurrentProject.Path & "\back\recursos\Plantillas\" & nombrePlantilla
    
    ' Validar que la plantilla existe
    If Not IDocumentService_ValidarPlantilla(rutaPlantilla) Then
        m_LastError = "Plantilla no encontrada: " & rutaPlantilla
        IDocumentService_GenerarDocumento = ""
        Exit Function
    End If
    
    ' Abrir plantilla
    If Not IDocumentService_AbrirPlantilla(rutaPlantilla) Then
        m_LastError = "Error al abrir plantilla: " & rutaPlantilla
        IDocumentService_GenerarDocumento = ""
        Exit Function
    End If
    
    ' Reemplazar campos
    If Not IDocumentService_ReemplazarCamposEnPlantilla(IdSolicitud, mapeoCampos) Then
        m_LastError = "Error al reemplazar campos en plantilla"
        IDocumentService_GenerarDocumento = ""
        Exit Function
    End If
    
    ' Generar nombre y ruta de destino usando los datos ya cargados
    nombreDocumento = GenerarNombreDocumentoConDatos(solicitudPC)
    rutaDestino = CurrentProject.Path & "\back\recursos\Documentos_Generados\"
    
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

Private Function IDocumentService_ObtenerNombrePlantilla(ByVal TipoSolicitud As String) As String
    On Error GoTo ErrorHandler
    
    Dim sql As String
    Dim rs As Object
    
    sql = "SELECT NombrePlantilla FROM Tb_Mapeo_Campos WHERE TipoSolicitud = '" & TipoSolicitud & "'"
    
    Set rs = CurrentDb.OpenRecordset(sql)
    
    If Not rs.EOF Then
        IDocumentService_ObtenerNombrePlantilla = rs("NombrePlantilla").Value
    Else
        ' Plantilla por defecto basada en tipo
        IDocumentService_ObtenerNombrePlantilla = TipoSolicitud & "_template.docx"
    End If
    
    rs.Close
    Set rs = Nothing
    
    Exit Function
    
ErrorHandler:
    m_LastError = "Error al obtener nombre de plantilla: " & Err.Description
    IDocumentService_ObtenerNombrePlantilla = TipoSolicitud & "_template.docx"
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
End Function

' Funci├│n eliminada: ValidarSolicitudParaDocumento - Ahora se usa IValidationService.ValidarSolicitud

Private Function IDocumentService_ObtenerMapeoCampos(ByVal TipoSolicitud As String) As String
    On Error GoTo ErrorHandler
    
    Dim sql As String
    Dim rs As Object
    
    sql = "SELECT MapeoCampos FROM Tb_Mapeo_Campos WHERE TipoSolicitud = '" & TipoSolicitud & "'"
    
    Set rs = CurrentDb.OpenRecordset(sql)
    
    If Not rs.EOF Then
        IDocumentService_ObtenerMapeoCampos = rs("MapeoCampos").Value
    Else
        ' Mapeo por defecto para tipos conocidos
        Select Case TipoSolicitud
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
' IMPLEMENTACI├ôN DE IDocumentService - PROCESAMIENTO DE PLANTILLAS
' ============================================================================

Private Function IDocumentService_AbrirPlantilla(ByVal RutaPlantilla As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Crear aplicaci├│n Word si no existe
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

Private Function IDocumentService_ReemplazarCamposEnPlantilla(ByVal IdSolicitud As Long, ByVal MapeoCampos As String) As Boolean
    On Error GoTo ErrorHandler
    
    If m_WordDoc Is Nothing Then
        m_LastError = "No hay documento abierto para reemplazar campos"
        IDocumentService_ReemplazarCamposEnPlantilla = False
        Exit Function
    End If
    
    ' Obtener datos de la solicitud
    Dim rs As Object
    Set rs = IDocumentService_ObtenerDatosSolicitud(IdSolicitud)
    
    If rs Is Nothing Or rs.EOF Then
        m_LastError = "No se encontraron datos para la solicitud: " & IdSolicitud
        IDocumentService_ReemplazarCamposEnPlantilla = False
        Exit Function
    End If
    
    ' Procesar cada campo del mapeo
    Dim campos() As String
    Dim i As Integer
    Dim nombreCampo As String
    Dim valorCampo As String
    Dim marcador As String
    
    campos = Split(MapeoCampos, ",")
    
    For i = 0 To UBound(campos)
        nombreCampo = Trim(campos(i))
        marcador = "{{" & nombreCampo & "}}"
        
        ' Obtener valor del campo
        valorCampo = ObtenerValorCampo(rs, nombreCampo)
        
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

Private Function IDocumentService_GuardarDocumento(ByVal RutaDestino As String, ByVal NombreArchivo As String) As Boolean
    On Error GoTo ErrorHandler
    
    If m_WordDoc Is Nothing Then
        m_LastError = "No hay documento abierto para guardar"
        IDocumentService_GuardarDocumento = False
        Exit Function
    End If
    
    Dim rutaCompleta As String
    rutaCompleta = RutaDestino & NombreArchivo
    
    ' Guardar como documento Word
    m_WordDoc.SaveAs2 rutaCompleta
    
    IDocumentService_GuardarDocumento = True
    Exit Function
    
ErrorHandler:
    m_LastError = "Error al guardar documento: " & Err.Description
    IDocumentService_GuardarDocumento = False
End Function

' ============================================================================
' IMPLEMENTACI├ôN DE IDocumentService - CONSULTA DE DATOS
' ============================================================================

Private Function IDocumentService_ObtenerDatosSolicitud(ByVal IdSolicitud As Long) As Object
    On Error GoTo ErrorHandler
    
    Dim sql As String
    
    ' Consulta principal que une datos de solicitud con datos espec├¡ficos
    sql = "SELECT s.IdSolicitud, s.NumeroExpediente, s.FechaSolicitud, s.TipoSolicitud, s.EstadoInterno, " & _
          "pc.NombreSolicitante, pc.ApellidosSolicitante, pc.DNI " & _
          "FROM TbSolicitudes s " & _
          "LEFT JOIN TbDatos_PC pc ON s.IdSolicitud = pc.IdSolicitud " & _
          "WHERE s.IdSolicitud = " & IdSolicitud
    
    Set IDocumentService_ObtenerDatosSolicitud = CurrentDb.OpenRecordset(sql)
    
    Exit Function
    
ErrorHandler:
    m_LastError = "Error al obtener datos de solicitud: " & Err.Description
    Set IDocumentService_ObtenerDatosSolicitud = Nothing
End Function

Private Function IDocumentService_ObtenerConfiguracionMapeo(ByVal TipoSolicitud As String) As Object
    On Error GoTo ErrorHandler
    
    Dim sql As String
    
    sql = "SELECT * FROM Tb_Mapeo_Campos WHERE TipoSolicitud = '" & TipoSolicitud & "'"
    
    Set IDocumentService_ObtenerConfiguracionMapeo = CurrentDb.OpenRecordset(sql)
    
    Exit Function
    
ErrorHandler:
    m_LastError = "Error al obtener configuraci├│n de mapeo: " & Err.Description
    Set IDocumentService_ObtenerConfiguracionMapeo = Nothing
End Function

' ============================================================================
' IMPLEMENTACI├ôN DE IDocumentService - UTILIDADES
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

Private Function IDocumentService_GenerarNombreDocumento(ByVal IdSolicitud As Long, ByVal TipoSolicitud As String) As String
    On Error GoTo ErrorHandler
    
    Dim fechaActual As String
    Dim numeroExpediente As String
    Dim solicitud As T_Solicitud
    
    fechaActual = Format(Date, "yyyy-mm-dd")
    
    ' Usar el repositorio inyectado para obtener la solicitud y su n├║mero de expediente
    Set solicitud = m_SolicitudRepository.ObtenerPorId(IdSolicitud)
    If Not solicitud Is Nothing Then
        numeroExpediente = solicitud.NumeroExpediente
    Else
        numeroExpediente = ""
    End If
    
    If numeroExpediente <> "" Then
        IDocumentService_GenerarNombreDocumento = "DOC-" & TipoSolicitud & "-" & numeroExpediente & "-" & fechaActual & ".docx"
    Else
        IDocumentService_GenerarNombreDocumento = "DOC-" & TipoSolicitud & "-" & IdSolicitud & "-" & fechaActual & ".docx"
    End If
    
    Exit Function
    
ErrorHandler:
    m_LastError = "Error al generar nombre de documento: " & Err.Description
    IDocumentService_GenerarNombreDocumento = "DOC-" & TipoSolicitud & "-" & IdSolicitud & ".docx"
End Function

Private Sub IDocumentService_LimpiarRecursos()
    On Error Resume Next
    
    ' Cerrar documento si est├í abierto
    If Not m_WordDoc Is Nothing Then
        m_WordDoc.Close False ' No guardar cambios
        Set m_WordDoc = Nothing
    End If
    
    ' Cerrar aplicaci├│n Word si la creamos nosotros
    If m_IsWordAppCreated And Not m_WordApp Is Nothing Then
        m_WordApp.Quit
        Set m_WordApp = Nothing
        m_IsWordAppCreated = False
    End If
    
    m_LastError = ""
End Sub

' ============================================================================
' M├ëTODOS AUXILIARES PRIVADOS
' ============================================================================

' Funci├│n eliminada: ObtenerTipoSolicitud - Ahora se usa ISolicitudRepository.LoadPC

' Funci├│n auxiliar para generar nombre usando datos ya cargados
Private Function GenerarNombreDocumentoConDatos(ByVal solicitudPC As CSolicitudPC) As String
    On Error GoTo ErrorHandler
    
    Dim fechaActual As String
    Dim numeroExpediente As String
    
    fechaActual = Format(Date, "yyyy-mm-dd")
    numeroExpediente = solicitudPC.IDExpediente
    
    If numeroExpediente <> "" Then
        GenerarNombreDocumentoConDatos = "DOC-" & solicitudPC.TipoSolicitud & "-" & numeroExpediente & "-" & fechaActual & ".docx"
    Else
        GenerarNombreDocumentoConDatos = "DOC-" & solicitudPC.TipoSolicitud & "-" & solicitudPC.idSolicitud & "-" & fechaActual & ".docx"
    End If
    
    Exit Function
    
ErrorHandler:
    m_LastError = "Error al generar nombre de documento: " & Err.Description
    GenerarNombreDocumentoConDatos = "DOC-" & solicitudPC.TipoSolicitud & "-" & solicitudPC.idSolicitud & ".docx"
End Function

Private Function ObtenerValorCampo(ByRef rs As Object, ByVal NombreCampo As String) As String
    On Error GoTo ErrorHandler
    
    Dim valor As Variant
    
    ' Intentar obtener el valor del campo
    valor = rs(NombreCampo).Value
    
    If IsNull(valor) Then
        ObtenerValorCampo = ""
    Else
        ' Formatear seg├║n el tipo de campo
        Select Case NombreCampo
            Case "FechaSolicitud", "FechaLimite"
                ObtenerValorCampo = Format(valor, "dd/mm/yyyy")
            Case Else
                ObtenerValorCampo = CStr(valor)
        End Select
    End If
    
    Exit Function
    
ErrorHandler:
    ObtenerValorCampo = "[Error: " & NombreCampo & "]"
End Function

' ============================================================================
' M├ëTODOS P├ÜBLICOS ADICIONALES
' ============================================================================

Public Function GetLastError() As String
    GetLastError = m_LastError
End Function

Public Sub ClearLastError()
    m_LastError = ""
End Sub
