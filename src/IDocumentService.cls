Option Compare Database
Option Explicit
' ============================================================================
' Interfaz: IDocumentService
' Descripción: Define el contrato para el servicio de generación de documentos
' Autor: CONDOR-Expert
' Fecha: Enero 2025
' Versión: 1.0 - Implementación inicial
' ============================================================================

' ============================================================================
' MÉTODOS PRINCIPALES DE GENERACIÓN DE DOCUMENTOS
' ============================================================================

' Genera un documento oficial basado en una solicitud
' Parámetros:
'   - IdSolicitud: ID de la solicitud para generar el documento
' Retorna: Ruta completa del documento generado
Public Function GenerarDocumento(ByVal idSolicitud As Long) As String
End Function

' Obtiene el nombre de la plantilla correspondiente al tipo de solicitud
' Parámetros:
'   - TipoSolicitud: Tipo de solicitud (PC, RAC, etc.)
' Retorna: Nombre del archivo de plantilla
Public Function ObtenerNombrePlantilla(ByVal tipoSolicitud As String) As String
End Function

' Método eliminado: ValidarSolicitudParaDocumento
' Razón: Se usa IValidationService.ValidarSolicitud en su lugar

' Obtiene el mapeo de campos para un tipo de solicitud específico
' Parámetros:
'   - TipoSolicitud: Tipo de solicitud (PC, RAC, etc.)
' Retorna: String con los campos mapeados separados por comas
Public Function ObtenerMapeoCampos(ByVal tipoSolicitud As String) As String
End Function

' ============================================================================
' MÉTODOS DE PROCESAMIENTO DE PLANTILLAS
' ============================================================================

' Abre una plantilla Word y la prepara para el procesamiento
' Parámetros:
'   - RutaPlantilla: Ruta completa a la plantilla Word
' Retorna: True si la plantilla se abrió correctamente
Public Function AbrirPlantilla(ByVal RutaPlantilla As String) As Boolean
End Function

' Reemplaza los marcadores en la plantilla con los datos de la solicitud
' Parámetros:
'   - IdSolicitud: ID de la solicitud con los datos
'   - MapeoCampos: String con el mapeo de campos
' Retorna: True si el reemplazo fue exitoso
Public Function ReemplazarCamposEnPlantilla(ByVal idSolicitud As Long, ByVal MapeoCampos As String) As Boolean
End Function

' Guarda el documento procesado en la ubicación especificada
' Parámetros:
'   - RutaDestino: Ruta donde guardar el documento final
'   - NombreArchivo: Nombre del archivo a generar
' Retorna: True si el guardado fue exitoso
Public Function GuardarDocumento(ByVal rutaDestino As String, ByVal NombreArchivo As String) As Boolean
End Function

' ============================================================================
' MÉTODOS DE CONSULTA DE DATOS
' ============================================================================

' Obtiene los datos de una solicitud desde la base de datos
' Parámetros:
'   - IdSolicitud: ID de la solicitud
' Retorna: Recordset con los datos de la solicitud
Public Function ObtenerDatosSolicitud(ByVal idSolicitud As Long) As Object
End Function

' Obtiene la configuración de mapeo desde Tb_Mapeo_Campos
' Parámetros:
'   - TipoSolicitud: Tipo de solicitud para obtener el mapeo
' Retorna: Recordset con la configuración de mapeo
Public Function ObtenerConfiguracionMapeo(ByVal tipoSolicitud As String) As Object
End Function

' ============================================================================
' MÉTODOS DE UTILIDAD Y VALIDACIÓN
' ============================================================================

' Valida que la plantilla existe y es accesible
' Parámetros:
'   - RutaPlantilla: Ruta a la plantilla a validar
' Retorna: True si la plantilla es válida
Public Function ValidarPlantilla(ByVal RutaPlantilla As String) As Boolean
End Function

' Genera un nombre único para el documento basado en la solicitud
' Parámetros:
'   - IdSolicitud: ID de la solicitud
'   - TipoSolicitud: Tipo de solicitud
' Retorna: Nombre único para el documento
Public Function GenerarNombreDocumento(ByVal idSolicitud As Long, ByVal tipoSolicitud As String) As String
End Function

' Limpia recursos y cierra conexiones abiertas
Public Sub LimpiarRecursos()
End Sub






