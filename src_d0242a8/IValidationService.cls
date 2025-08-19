VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IValidationService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ===================================================================
' Interfaz: IValidationService
' Descripci├│n: Contrato para servicios de validaci├│n de reglas de negocio
' Autor: Sistema CONDOR
' Fecha: 2024
' ===================================================================

Option Explicit

' ===================================================================
' M├ëTODOS DE VALIDACI├ôN PRINCIPAL
' ===================================================================

' Valida una solicitud completa seg├║n las reglas de negocio
' Par├ímetros:
'   solicitud: Objeto T_Solicitud a validar
'   mensajeError: Variable de salida con el mensaje de error (si aplica)
' Retorna: True si la validaci├│n es exitosa, False en caso contrario
Public Function ValidarSolicitud(ByVal solicitud As T_Solicitud, ByRef mensajeError As String) As Boolean
End Function

' ===================================================================
' M├ëTODOS DE VALIDACI├ôN ESPEC├ìFICA
' ===================================================================

' Valida que el n├║mero de expediente sea v├ílido
Public Function ValidarNumeroExpediente(ByVal numeroExpediente As String, ByRef mensajeError As String) As Boolean
End Function

' Valida que el tipo de solicitud sea v├ílido
Public Function ValidarTipoSolicitud(ByVal tipoSolicitud As String, ByRef mensajeError As String) As Boolean
End Function

' Valida que los campos obligatorios est├®n completos
Public Function ValidarCamposObligatorios(ByVal solicitud As T_Solicitud, ByRef mensajeError As String) As Boolean
End Function

' Valida que los importes sean v├ílidos (para solicitudes PC)
Public Function ValidarImportes(ByVal importeOriginal As Currency, ByVal importeNuevo As Currency, ByRef mensajeError As String) As Boolean
End Function

' Valida que la transici├│n de estado sea permitida
Public Function ValidarTransicionEstado(ByVal estadoActual As String, ByVal estadoNuevo As String, ByRef mensajeError As String) As Boolean
End Function
