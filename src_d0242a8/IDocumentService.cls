VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDocumentService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' ============================================================================
' Interfaz: IDocumentService
' Descripci├│n: Define el contrato para el servicio de generaci├│n de documentos
' Autor: CONDOR-Expert
' Fecha: Enero 2025
' Versi├│n: 1.0 - Implementaci├│n inicial
' ============================================================================

' ============================================================================
' M├ëTODOS PRINCIPALES DE GENERACI├ôN DE DOCUMENTOS
' ============================================================================

' Genera un documento oficial basado en una solicitud
' Par├ímetros:
'   - IdSolicitud: ID de la solicitud para generar el documento
' Retorna: Ruta completa del documento generado
Public Function GenerarDocumento(ByVal IdSolicitud As Long) As String
End Function

' Obtiene el nombre de la plantilla correspondiente al tipo de solicitud
' Par├ímetros:
'   - TipoSolicitud: Tipo de solicitud (PC, RAC, etc.)
' Retorna: Nombre del archivo de plantilla
Public Function ObtenerNombrePlantilla(ByVal TipoSolicitud As String) As String
End Function

' M├®todo eliminado: ValidarSolicitudParaDocumento
' Raz├│n: Se usa IValidationService.ValidarSolicitud en su lugar

' Obtiene el mapeo de campos para un tipo de solicitud espec├¡fico
' Par├ímetros:
'   - TipoSolicitud: Tipo de solicitud (PC, RAC, etc.)
' Retorna: String con los campos mapeados separados por comas
Public Function ObtenerMapeoCampos(ByVal TipoSolicitud As String) As String
End Function

' ============================================================================
' M├ëTODOS DE PROCESAMIENTO DE PLANTILLAS
' ============================================================================

' Abre una plantilla Word y la prepara para el procesamiento
' Par├ímetros:
'   - RutaPlantilla: Ruta completa a la plantilla Word
' Retorna: True si la plantilla se abri├│ correctamente
Public Function AbrirPlantilla(ByVal RutaPlantilla As String) As Boolean
End Function

' Reemplaza los marcadores en la plantilla con los datos de la solicitud
' Par├ímetros:
'   - IdSolicitud: ID de la solicitud con los datos
'   - MapeoCampos: String con el mapeo de campos
' Retorna: True si el reemplazo fue exitoso
Public Function ReemplazarCamposEnPlantilla(ByVal IdSolicitud As Long, ByVal MapeoCampos As String) As Boolean
End Function

' Guarda el documento procesado en la ubicaci├│n especificada
' Par├ímetros:
'   - RutaDestino: Ruta donde guardar el documento final
'   - NombreArchivo: Nombre del archivo a generar
' Retorna: True si el guardado fue exitoso
Public Function GuardarDocumento(ByVal RutaDestino As String, ByVal NombreArchivo As String) As Boolean
End Function

' ============================================================================
' M├ëTODOS DE CONSULTA DE DATOS
' ============================================================================

' Obtiene los datos de una solicitud desde la base de datos
' Par├ímetros:
'   - IdSolicitud: ID de la solicitud
' Retorna: Recordset con los datos de la solicitud
Public Function ObtenerDatosSolicitud(ByVal IdSolicitud As Long) As Object
End Function

' Obtiene la configuraci├│n de mapeo desde Tb_Mapeo_Campos
' Par├ímetros:
'   - TipoSolicitud: Tipo de solicitud para obtener el mapeo
' Retorna: Recordset con la configuraci├│n de mapeo
Public Function ObtenerConfiguracionMapeo(ByVal TipoSolicitud As String) As Object
End Function

' ============================================================================
' M├ëTODOS DE UTILIDAD Y VALIDACI├ôN
' ============================================================================

' Valida que la plantilla existe y es accesible
' Par├ímetros:
'   - RutaPlantilla: Ruta a la plantilla a validar
' Retorna: True si la plantilla es v├ílida
Public Function ValidarPlantilla(ByVal RutaPlantilla As String) As Boolean
End Function

' Genera un nombre ├║nico para el documento basado en la solicitud
' Par├ímetros:
'   - IdSolicitud: ID de la solicitud
'   - TipoSolicitud: Tipo de solicitud
' Retorna: Nombre ├║nico para el documento
Public Function GenerarNombreDocumento(ByVal IdSolicitud As Long, ByVal TipoSolicitud As String) As String
End Function

' Limpia recursos y cierra conexiones abiertas
Public Sub LimpiarRecursos()
End Sub
