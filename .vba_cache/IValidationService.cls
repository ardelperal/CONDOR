Option Compare Database
Option Explicit
' ===================================================================
' Interfaz: IValidationService
' Descripción: Contrato para servicios de validación de reglas de negocio
' Autor: Sistema CONDOR
' Fecha: 2024
' ===================================================================


' ===================================================================
' MÉTODOS DE VALIDACIÓN PRINCIPAL
' ===================================================================

' Valida una solicitud completa según las reglas de negocio
' Parámetros:
'   solicitud: Objeto T_Solicitud a validar
'   mensajeError: Variable de salida con el mensaje de error (si aplica)
' Retorna: True si la validación es exitosa, False en caso contrario
Public Function ValidarSolicitud(ByVal solicitud As T_Solicitud, ByRef MensajeError As String) As Boolean
End Function

' ===================================================================
' MÉTODOS DE VALIDACIÓN ESPECÍFICA
' ===================================================================

' Valida que el número de expediente sea válido
Public Function ValidarNumeroExpediente(ByVal NumeroExpediente As String, ByRef MensajeError As String) As Boolean
End Function

' Valida que el tipo de solicitud sea válido
Public Function ValidarTipoSolicitud(ByVal tipoSolicitud As String, ByRef MensajeError As String) As Boolean
End Function

' Valida que los campos obligatorios estén completos
Public Function ValidarCamposObligatorios(ByVal solicitud As T_Solicitud, ByRef MensajeError As String) As Boolean
End Function

' Valida que los importes sean válidos (para solicitudes PC)
Public Function ValidarImportes(ByVal importeOriginal As Currency, ByVal importeNuevo As Currency, ByRef MensajeError As String) As Boolean
End Function

' Valida que la transición de estado sea permitida
Public Function ValidarTransicionEstado(ByVal estadoActual As String, ByVal estadoNuevo As String, ByRef MensajeError As String) As Boolean
End Function






