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
' Descripción: Implementación del servicio de generación y lectura de documentos.
' Implementa: IDocumentService
' Autor: Gemini @ Google
' Fecha: 2025-08-22
' Versión: 2.0 - Refactorizada escritura, implementada lectura y auditoría.
' ============================================================================

Implements IDocumentService

' Dependencias inyectadas
Private m_Config As IConfig
Private m_solicitudRepository As ISolicitudRepository
Private m_OperationLogger As IOperationLogger
Private m_WordManager As IWordManager
Private m_MapeoRepository As IMapeoRepository

' ============================================================================
' CONSTRUCTOR Y DESTRUCTOR
' ============================================================================

Private Sub Class_Initialize()
    Set m_Config = Nothing
    Set m_solicitudRepository = Nothing
    Set m_OperationLogger = Nothing
    Set m_WordManager = Nothing
    Set m_MapeoRepository = Nothing
End Sub

Private Sub Class_Terminate()
    ' Los recursos de Word ahora son manejados por IWordManager
End Sub

' ============================================================================
' MÉTODO DE INICIALIZACIÓN CON DEPENDENCIAS (CONTRATO DE INTERFAZ)
' ============================================================================

Public Sub Initialize(ByVal configService As IConfig, ByVal solicitudRepository As ISolicitudRepository, ByVal operationLogger As IOperationLogger, ByVal wordManager As IWordManager, ByVal mapeoRepository As IMapeoRepository)
    Set m_Config = configService
    Set m_solicitudRepository = solicitudRepository
    Set m_OperationLogger = operationLogger
    Set m_WordManager = wordManager
    Set m_MapeoRepository = mapeoRepository
End Sub

' ============================================================================
' MÉTODOS PÚBLICOS DE CONVENIENCIA (Lección 24)
' ============================================================================

Public Function GenerarDocumento(ByVal solicitud As T_Solicitud) As String
    GenerarDocumento = IDocumentService_GenerarDocumento(solicitud)
End Function

Public Function LeerDocumento(ByVal rutaDocumento As String, ByVal idSolicitud As Long) As Boolean
    LeerDocumento = IDocumentService_LeerDocumento(rutaDocumento, idSolicitud)
End Function

' ============================================================================
' IMPLEMENTACIÓN DE IDocumentService - ESCRITURA DE DOCUMENTOS
' ============================================================================

Private Function IDocumentService_GenerarDocumento(ByVal solicitud As T_Solicitud) As String
    Dim ErrorHandler As IErrorHandlerService
    On Error GoTo ErrorHandler
    
    Dim rutaPlantilla As String
    Dim nombreDocumento As String
    Dim rutaDestino As String
    Dim rutaCompleta As String
    
    If m_Config Is Nothing Or m_solicitudRepository Is Nothing Or m_OperationLogger Is Nothing Then
        Set ErrorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
        ErrorHandler.LogError -1, "Dependencias no inicializadas.", "CDocumentService.GenerarDocumento"
        IDocumentService_GenerarDocumento = ""
        Exit Function
    End If
    
    m_OperationLogger.LogOperation "DocGen_Start", "Iniciando generación de documento para Solicitud ID: " & solicitud.idSolicitud, "CDocumentService"
    
    ' 1. Obtener ruta de la plantilla
    rutaPlantilla = m_Config.GetValue("PLANTILLA_PATH") & "\" & solicitud.tipoSolicitud & ".docx"
    
    ' En modo de prueba, no verificamos la existencia del fichero para mantener el test aislado
    If Not CBool(m_Config.GetValue("IS_TEST_MODE")) Then
        If Dir(rutaPlantilla) = "" Then
            Set ErrorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
            ErrorHandler.LogError -1, "No se encontró la plantilla en: " & rutaPlantilla, "CDocumentService.GenerarDocumento"
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
    If Not ReemplazarCamposEnPlantilla(solicitud) Then
        IDocumentService_GenerarDocumento = ""
        Exit Function
    End If
    
    ' 4. Guardar el documento
    nombreDocumento = "SOL_" & solicitud.codigoSolicitud & "_" & Format(Now, "YYYYMMDD_HHNNSS") & ".docx"
    rutaDestino = m_Config.GetValue("GENERATED_DOCS_PATH")
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
    Set ErrorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    ErrorHandler.LogError Err.Number, Err.Description, "CDocumentService.GenerarDocumento"
    m_WordManager.CerrarDocumento
    IDocumentService_GenerarDocumento = ""
End Function

' ============================================================================
' IMPLEMENTACIÓN DE IDocumentService - LECTURA DE DOCUMENTOS
' ============================================================================

Private Function IDocumentService_LeerDocumento(ByVal rutaDocumento As String, ByVal idSolicitud As Long) As Boolean
    Dim ErrorHandler As IErrorHandlerService
    On Error GoTo ErrorHandler
    
    Dim solicitudLeida As T_Solicitud
    Dim rsMapeo As Object ' Recordset
    Dim campoTabla As String, campoWord As String, valorLeido As String
    
    If m_Config Is Nothing Or m_solicitudRepository Is Nothing Or m_OperationLogger Is Nothing Then
        Set ErrorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
        ErrorHandler.LogError -1, "Dependencias no inicializadas.", "CDocumentService.LeerDocumento"
        IDocumentService_LeerDocumento = False
        Exit Function
    End If
    
    m_OperationLogger.LogOperation "DocRead_Start", "Iniciando lectura de documento " & rutaDocumento & " para Solicitud ID: " & idSolicitud, "CDocumentService"
    
    ' 1. Cargar la solicitud existente
    Set solicitudLeida = m_solicitudRepository.GetSolicitudById(idSolicitud)
    If solicitudLeida Is Nothing Then
        Set ErrorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
        ErrorHandler.LogError -1, "No se pudo cargar la solicitud con ID: " & idSolicitud, "CDocumentService.LeerDocumento"
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
    Set rsMapeo = m_MapeoRepository.GetMapeoPorTipo(solicitudLeida.tipoSolicitud)
    If rsMapeo.EOF Then
        Set ErrorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
        ErrorHandler.LogError -1, "No se encontró mapeo para tipo: " & solicitudLeida.tipoSolicitud, "CDocumentService.LeerDocumento"
        IDocumentService_LeerDocumento = False
        Exit Function
    End If
    
    ' 4. Iterar, leer y actualizar el objeto
    Do While Not rsMapeo.EOF
        campoTabla = rsMapeo!NombreCampoTabla
        campoWord = rsMapeo!NombreCampoWord
        
        ' Buscar el marcador en el contenido del documento
        ' Nota: Esta es una implementación simplificada para demostrar el patrón
        ' En una implementación real, se necesitaría un parser más sofisticado
        If InStr(contenidoDocumento, campoWord) > 0 Then
            ' Extraer valor del marcador (implementación simplificada)
            valorLeido = ExtraerValorMarcador(contenidoDocumento, campoWord)
            
            ' Actualizar la propiedad del objeto solicitud correspondiente usando CallByName
            CallByName solicitudLeida, campoTabla, VbLet, valorLeido
        Else
            m_OperationLogger.LogOperation "DocRead_Warning", "Marcador no encontrado en documento: " & campoWord, "CDocumentService"
        End If
        rsMapeo.MoveNext
    Loop
    
    ' 5. Guardar la solicitud actualizada
    If m_solicitudRepository.SaveSolicitud(solicitudLeida) = 0 Then
        m_OperationLogger.LogOperation "DocRead_SaveFail", "Fallo al guardar la solicitud ID: " & idSolicitud & " usando método ISolicitudRepository.SaveSolicitud", "CDocumentService"
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
    Set ErrorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    ErrorHandler.LogError Err.Number, Err.Description, "CDocumentService.LeerDocumento"
    If Not rsMapeo Is Nothing Then rsMapeo.Close
    Set rsMapeo = Nothing
    IDocumentService_LeerDocumento = False
End Function

' ============================================================================
' MÉTODOS PRIVADOS AUXILIARES
' ============================================================================

Private Function ReemplazarCamposEnPlantilla(ByVal solicitud As T_Solicitud) As Boolean
    Dim ErrorHandler As IErrorHandlerService
    On Error GoTo ErrorHandler
    
    Set ErrorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    
    Dim rsMapeo As Object ' Recordset
    Dim Datos As Object
    Dim campoTabla As String, campoWord As String, ValorAsociado As String
    Dim valorPropiedad As Variant
    
    Set Datos = solicitud.Datos
    Set rsMapeo = m_MapeoRepository.GetMapeoPorTipo(solicitud.tipoSolicitud)
    
    If rsMapeo.EOF Then
        ErrorHandler.LogError -1, "No se encontró mapeo para tipo: " & solicitud.tipoSolicitud, "CDocumentService.ReemplazarCampos"
        ReemplazarCamposEnPlantilla = False
        Exit Function
    End If
    
    Do While Not rsMapeo.EOF
        campoTabla = rsMapeo!NombreCampoTabla
        campoWord = rsMapeo!NombreCampoWord
        ValorAsociado = Nz(rsMapeo!ValorAsociado, "")
        
        valorPropiedad = CallByName(Datos, campoTabla, VbGet)
        
        Dim textoReemplazo As String
        If VarType(valorPropiedad) = vbBoolean Then
            If valorPropiedad = True Then
                textoReemplazo = "X" ' O ChrW(&HFE) con fuente Wingdings
            Else
                textoReemplazo = ""
            End If
        Else
            If ValorAsociado <> "" Then
                If CStr(valorPropiedad) = ValorAsociado Then
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
            m_OperationLogger.LogOperation "DocGen_Warning", "No se pudo reemplazar marcador: " & campoWord & " para Solicitud ID: " & solicitud.idSolicitud, "CDocumentService"
        End If
        rsMapeo.MoveNext
    Loop
    
    ReemplazarCamposEnPlantilla = True
    rsMapeo.Close
    Set rsMapeo = Nothing
    Exit Function

ErrorHandler:
    ErrorHandler.LogError Err.Number, Err.Description, "CDocumentService.ReemplazarCampos"
    If Not rsMapeo Is Nothing Then rsMapeo.Close
    Set rsMapeo = Nothing
    ReemplazarCamposEnPlantilla = False
End Function



Private Function ExtraerValorMarcador(ByVal contenidoDocumento As String, ByVal nombreMarcador As String) As String
    Dim ErrorHandler As IErrorHandlerService
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
    Set ErrorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    ErrorHandler.LogError Err.Number, Err.Description, "CDocumentService.ExtraerValorMarcador"
    ExtraerValorMarcador = ""
End Function

