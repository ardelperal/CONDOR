VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CValidationService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ===================================================================
' Clase: CValidationService
' Descripción: Implementación concreta del servicio de validaciones de negocio
' Autor: Sistema CONDOR
' Fecha: 2024
' ===================================================================

Option Explicit

' Implementa la interfaz IValidationService
Implements IValidationService

' ===================================================================
' IMPLEMENTACIÓN DE MÉTODOS DE VALIDACIÓN PRINCIPAL
' ===================================================================

Private Function IValidationService_ValidarSolicitud(ByVal solicitud As T_Solicitud, ByRef mensajeError As String) As Boolean
    On Error GoTo ErrorHandler
    
    mensajeError = ""
    
    ' Validar campos obligatorios
    If Not IValidationService_ValidarCamposObligatorios(solicitud, mensajeError) Then
        IValidationService_ValidarSolicitud = False
        Exit Function
    End If
    
    ' Validar número de expediente
    If Not IValidationService_ValidarNumeroExpediente(solicitud.NumeroExpediente, mensajeError) Then
        IValidationService_ValidarSolicitud = False
        Exit Function
    End If
    
    ' Validar tipo de solicitud
    If Not IValidationService_ValidarTipoSolicitud(solicitud.TipoSolicitud, mensajeError) Then
        IValidationService_ValidarSolicitud = False
        Exit Function
    End If
    
    ' Validar importes (solo para solicitudes PC)
    If solicitud.TipoSolicitud = "PC" Then
        If Not IValidationService_ValidarImportes(solicitud.ImporteOriginal, solicitud.ImporteNuevo, mensajeError) Then
            IValidationService_ValidarSolicitud = False
            Exit Function
        End If
    End If
    
    ' Si llegamos aquí, todas las validaciones pasaron
    IValidationService_ValidarSolicitud = True
    Exit Function
    
ErrorHandler:
    mensajeError = "Error interno en validación: " & Err.Description
    IValidationService_ValidarSolicitud = False
End Function

' ===================================================================
' IMPLEMENTACIÓN DE MÉTODOS DE VALIDACIÓN ESPECÍFICA
' ===================================================================

Private Function IValidationService_ValidarNumeroExpediente(ByVal numeroExpediente As String, ByRef mensajeError As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Validar que no esté vacío
    If Trim(numeroExpediente) = "" Then
        mensajeError = "El número de expediente es obligatorio"
        IValidationService_ValidarNumeroExpediente = False
        Exit Function
    End If
    
    ' Validar formato básico (debe contener al menos 3 caracteres)
    If Len(Trim(numeroExpediente)) < 3 Then
        mensajeError = "El número de expediente debe tener al menos 3 caracteres"
        IValidationService_ValidarNumeroExpediente = False
        Exit Function
    End If
    
    ' Validación exitosa
    IValidationService_ValidarNumeroExpediente = True
    Exit Function
    
ErrorHandler:
    mensajeError = "Error validando número de expediente: " & Err.Description
    IValidationService_ValidarNumeroExpediente = False
End Function

Private Function IValidationService_ValidarTipoSolicitud(ByVal tipoSolicitud As String, ByRef mensajeError As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Validar que no esté vacío
    If Trim(tipoSolicitud) = "" Then
        mensajeError = "El tipo de solicitud es obligatorio"
        IValidationService_ValidarTipoSolicitud = False
        Exit Function
    End If
    
    ' Validar que sea un tipo válido
    Select Case UCase(Trim(tipoSolicitud))
        Case "PC", "CD_CA", "CD_CA_SUB"
            IValidationService_ValidarTipoSolicitud = True
        Case Else
            mensajeError = "Tipo de solicitud no válido. Debe ser PC, CD_CA o CD_CA_SUB"
            IValidationService_ValidarTipoSolicitud = False
    End Select
    
    Exit Function
    
ErrorHandler:
    mensajeError = "Error validando tipo de solicitud: " & Err.Description
    IValidationService_ValidarTipoSolicitud = False
End Function

Private Function IValidationService_ValidarCamposObligatorios(ByVal solicitud As T_Solicitud, ByRef mensajeError As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Validar descripción
    If Trim(solicitud.Descripcion) = "" Then
        mensajeError = "La descripción es obligatoria"
        IValidationService_ValidarCamposObligatorios = False
        Exit Function
    End If
    
    ' Validar justificación
    If Trim(solicitud.Justificacion) = "" Then
        mensajeError = "La justificación es obligatoria"
        IValidationService_ValidarCamposObligatorios = False
        Exit Function
    End If
    
    ' Validar usuario creador
    If Trim(solicitud.UsuarioCreador) = "" Then
        mensajeError = "El usuario creador es obligatorio"
        IValidationService_ValidarCamposObligatorios = False
        Exit Function
    End If
    
    ' Validar estado
    If Trim(solicitud.Estado) = "" Then
        mensajeError = "El estado es obligatorio"
        IValidationService_ValidarCamposObligatorios = False
        Exit Function
    End If
    
    ' Todos los campos obligatorios están presentes
    IValidationService_ValidarCamposObligatorios = True
    Exit Function
    
ErrorHandler:
    mensajeError = "Error validando campos obligatorios: " & Err.Description
    IValidationService_ValidarCamposObligatorios = False
End Function

Private Function IValidationService_ValidarImportes(ByVal importeOriginal As Currency, ByVal importeNuevo As Currency, ByRef mensajeError As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Validar que los importes sean positivos
    If importeOriginal <= 0 Then
        mensajeError = "El importe original debe ser mayor que cero"
        IValidationService_ValidarImportes = False
        Exit Function
    End If
    
    If importeNuevo <= 0 Then
        mensajeError = "El importe nuevo debe ser mayor que cero"
        IValidationService_ValidarImportes = False
        Exit Function
    End If
    
    ' Validar que haya una diferencia significativa (al menos 1%)
    Dim porcentajeCambio As Double
    porcentajeCambio = Abs((importeNuevo - importeOriginal) / importeOriginal) * 100
    
    If porcentajeCambio < 1 Then
        mensajeError = "La diferencia entre importes debe ser al menos del 1%"
        IValidationService_ValidarImportes = False
        Exit Function
    End If
    
    ' Validación exitosa
    IValidationService_ValidarImportes = True
    Exit Function
    
ErrorHandler:
    mensajeError = "Error validando importes: " & Err.Description
    IValidationService_ValidarImportes = False
End Function

Private Function IValidationService_ValidarTransicionEstado(ByVal estadoActual As String, ByVal estadoNuevo As String, ByRef mensajeError As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Validar que los estados no estén vacíos
    If Trim(estadoActual) = "" Or Trim(estadoNuevo) = "" Then
        mensajeError = "Los estados actual y nuevo son obligatorios"
        IValidationService_ValidarTransicionEstado = False
        Exit Function
    End If
    
    ' Validar transiciones permitidas según las reglas de negocio
    Select Case UCase(Trim(estadoActual))
        Case "BORRADOR"
            ' Desde Borrador se puede ir a Enviado o Cancelado
            If UCase(Trim(estadoNuevo)) = "ENVIADO" Or UCase(Trim(estadoNuevo)) = "CANCELADO" Then
                IValidationService_ValidarTransicionEstado = True
            Else
                mensajeError = "Desde Borrador solo se puede pasar a Enviado o Cancelado"
                IValidationService_ValidarTransicionEstado = False
            End If
            
        Case "ENVIADO"
            ' Desde Enviado se puede ir a Aprobado, Rechazado o Cancelado
            If UCase(Trim(estadoNuevo)) = "APROBADO" Or UCase(Trim(estadoNuevo)) = "RECHAZADO" Or UCase(Trim(estadoNuevo)) = "CANCELADO" Then
                IValidationService_ValidarTransicionEstado = True
            Else
                mensajeError = "Desde Enviado solo se puede pasar a Aprobado, Rechazado o Cancelado"
                IValidationService_ValidarTransicionEstado = False
            End If
            
        Case "APROBADO"
            ' Desde Aprobado se puede ir a Finalizado
            If UCase(Trim(estadoNuevo)) = "FINALIZADO" Then
                IValidationService_ValidarTransicionEstado = True
            Else
                mensajeError = "Desde Aprobado solo se puede pasar a Finalizado"
                IValidationService_ValidarTransicionEstado = False
            End If
            
        Case "RECHAZADO", "CANCELADO", "FINALIZADO"
            ' Estados finales - no se puede cambiar
            mensajeError = "No se puede cambiar el estado desde " & estadoActual
            IValidationService_ValidarTransicionEstado = False
            
        Case Else
            mensajeError = "Estado actual no reconocido: " & estadoActual
            IValidationService_ValidarTransicionEstado = False
    End Select
    
    Exit Function
    
ErrorHandler:
    mensajeError = "Error validando transición de estado: " & Err.Description
    IValidationService_ValidarTransicionEstado = False
End Function