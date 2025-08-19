Attribute VB_Name = "Test_WorkflowService"
Option Compare Database


Option Explicit

'******************************************************************************
' MÓDULO: Test_WorkflowService
' DESCRIPCIÓN: Pruebas unitarias para el servicio de workflow y estados
' AUTOR: Sistema CONDOR
' FECHA: 2024
' NOTAS: Implementación TDD siguiendo las lecciones aprendidas
'        - Lección 2: Declarar variables del tipo de interfaz
'        - Lección 4: Estructura AAA (Arrange-Act-Assert)
'******************************************************************************


#If DEV_MODE Then

'******************************************************************************
' VARIABLES PRIVADAS DEL MÓDULO
'******************************************************************************

' Variable del servicio - LECCIÓN 2: Declarar del tipo de interfaz
Private workflowService As IWorkflowService

'******************************************************************************
' FUNCIÓN PRINCIPAL DE EJECUCIÓN DE PRUEBAS
'******************************************************************************

Public Function RunAll() As String
    Dim resultado As String
    Dim totalTests As Integer
    Dim passedTests As Integer
    
    resultado = "=== EJECUTANDO PRUEBAS DE WORKFLOW SERVICE ===" & vbCrLf
    totalTests = 0
    passedTests = 0
    
    ' ========================================
    ' EJECUTAR PRUEBAS DE TRANSICIONES VÁLIDAS
    ' ========================================
    
    totalTests = totalTests + 1
    If Test_ValidTransition_Borrador_To_EnProceso_ReturnsTrue() Then
        passedTests = passedTests + 1
        resultado = resultado & "? Test_ValidTransition_Borrador_To_EnProceso_ReturnsTrue" & vbCrLf
    Else
        resultado = resultado & "? Test_ValidTransition_Borrador_To_EnProceso_ReturnsTrue" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_ValidTransition_EnProceso_To_Aprobado_ReturnsTrue() Then
        passedTests = passedTests + 1
        resultado = resultado & "? Test_ValidTransition_EnProceso_To_Aprobado_ReturnsTrue" & vbCrLf
    Else
        resultado = resultado & "? Test_ValidTransition_EnProceso_To_Aprobado_ReturnsTrue" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_ValidTransition_EnProceso_To_Rechazado_ReturnsTrue() Then
        passedTests = passedTests + 1
        resultado = resultado & "? Test_ValidTransition_EnProceso_To_Rechazado_ReturnsTrue" & vbCrLf
    Else
        resultado = resultado & "? Test_ValidTransition_EnProceso_To_Rechazado_ReturnsTrue" & vbCrLf
    End If
    
    ' ========================================
    ' EJECUTAR PRUEBAS DE TRANSICIONES INVÁLIDAS
    ' ========================================
    
    totalTests = totalTests + 1
    If Test_InvalidTransition_Borrador_To_Aprobado_ReturnsFalse() Then
        passedTests = passedTests + 1
        resultado = resultado & "? Test_InvalidTransition_Borrador_To_Aprobado_ReturnsFalse" & vbCrLf
    Else
        resultado = resultado & "? Test_InvalidTransition_Borrador_To_Aprobado_ReturnsFalse" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_InvalidTransition_Aprobado_To_Borrador_ReturnsFalse() Then
        passedTests = passedTests + 1
        resultado = resultado & "? Test_InvalidTransition_Aprobado_To_Borrador_ReturnsFalse" & vbCrLf
    Else
        resultado = resultado & "? Test_InvalidTransition_Aprobado_To_Borrador_ReturnsFalse" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_InvalidTransition_NonExistentState_ReturnsFalse() Then
        passedTests = passedTests + 1
        resultado = resultado & "? Test_InvalidTransition_NonExistentState_ReturnsFalse" & vbCrLf
    Else
        resultado = resultado & "? Test_InvalidTransition_NonExistentState_ReturnsFalse" & vbCrLf
    End If
    
    ' ========================================
    ' EJECUTAR PRUEBAS DE VALIDACIÓN DE PERMISOS
    ' ========================================
    
    totalTests = totalTests + 1
    If Test_ValidatePermission_AdminRole_AllowsTransition_ReturnsTrue() Then
        passedTests = passedTests + 1
        resultado = resultado & "? Test_ValidatePermission_AdminRole_AllowsTransition_ReturnsTrue" & vbCrLf
    Else
        resultado = resultado & "? Test_ValidatePermission_AdminRole_AllowsTransition_ReturnsTrue" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_ValidatePermission_UserRole_DeniesTransition_ReturnsFalse() Then
        passedTests = passedTests + 1
        resultado = resultado & "? Test_ValidatePermission_UserRole_DeniesTransition_ReturnsFalse" & vbCrLf
    Else
        resultado = resultado & "? Test_ValidatePermission_UserRole_DeniesTransition_ReturnsFalse" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_ValidatePermission_EmptyRole_ReturnsFalse() Then
        passedTests = passedTests + 1
        resultado = resultado & "? Test_ValidatePermission_EmptyRole_ReturnsFalse" & vbCrLf
    Else
        resultado = resultado & "? Test_ValidatePermission_EmptyRole_ReturnsFalse" & vbCrLf
    End If
    
    ' ========================================
    ' EJECUTAR PRUEBAS DE OBTENCIÓN DE ESTADOS
    ' ========================================
    
    totalTests = totalTests + 1
    If Test_GetAvailableStates_ForTipoPC_ReturnsCollection() Then
        passedTests = passedTests + 1
        resultado = resultado & "? Test_GetAvailableStates_ForTipoPC_ReturnsCollection" & vbCrLf
    Else
        resultado = resultado & "? Test_GetAvailableStates_ForTipoPC_ReturnsCollection" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_GetNextStates_FromBorrador_ReturnsValidStates() Then
        passedTests = passedTests + 1
        resultado = resultado & "? Test_GetNextStates_FromBorrador_ReturnsValidStates" & vbCrLf
    Else
        resultado = resultado & "? Test_GetNextStates_FromBorrador_ReturnsValidStates" & vbCrLf
    End If
    
    ' ========================================
    ' EJECUTAR PRUEBAS DE HISTORIAL DE ESTADOS
    ' ========================================
    
    totalTests = totalTests + 1
    If Test_RecordStateChange_ValidTransition_CreatesHistoryRecord() Then
        passedTests = passedTests + 1
        resultado = resultado & "? Test_RecordStateChange_ValidTransition_CreatesHistoryRecord" & vbCrLf
    Else
        resultado = resultado & "? Test_RecordStateChange_ValidTransition_CreatesHistoryRecord" & vbCrLf
    End If
    
    ' ========================================
    ' RESUMEN DE RESULTADOS
    ' ========================================
    
    resultado = resultado & "=== RESUMEN DE PRUEBAS DE WORKFLOW SERVICE ===" & vbCrLf
    resultado = resultado & "Total de pruebas: " & totalTests & vbCrLf
    resultado = resultado & "Pruebas exitosas: " & passedTests & vbCrLf
    resultado = resultado & "Pruebas fallidas: " & (totalTests - passedTests) & vbCrLf
    
    If passedTests = totalTests Then
        resultado = resultado & "? TODAS LAS PRUEBAS DE WORKFLOW SERVICE PASARON" & vbCrLf
    Else
        resultado = resultado & "? ALGUNAS PRUEBAS DE WORKFLOW SERVICE FALLARON" & vbCrLf
    End If
    
    RunAll = resultado
End Function

'******************************************************************************
' PRUEBAS DE TRANSICIONES VÁLIDAS
'******************************************************************************

Public Function Test_ValidTransition_Borrador_To_EnProceso_ReturnsTrue() As Boolean
    ' Arrange
    Dim SolicitudID As Long
    Dim estadoOrigen As String
    Dim estadoDestino As String
    Dim tipoSolicitud As String
    Dim usuarioRol As String
    Dim resultado As Boolean
    
    Set workflowService = New CWorkflowService
    SolicitudID = 1
    estadoOrigen = "Borrador"
    estadoDestino = "EnProceso"
    tipoSolicitud = "PC"
    usuarioRol = "Usuario"
    
    ' Act
    resultado = workflowService.ValidateTransition(SolicitudID, estadoOrigen, estadoDestino, tipoSolicitud, usuarioRol)
    
    ' Assert
    Test_ValidTransition_Borrador_To_EnProceso_ReturnsTrue = resultado
End Function

Public Function Test_ValidTransition_EnProceso_To_Aprobado_ReturnsTrue() As Boolean
    ' Arrange
    Dim SolicitudID As Long
    Dim estadoOrigen As String
    Dim estadoDestino As String
    Dim tipoSolicitud As String
    Dim usuarioRol As String
    Dim resultado As Boolean
    
    Set workflowService = New CWorkflowService
    SolicitudID = 1
    estadoOrigen = "EnProceso"
    estadoDestino = "Aprobado"
    tipoSolicitud = "PC"
    usuarioRol = "Aprobador"
    
    ' Act
    resultado = workflowService.ValidateTransition(SolicitudID, estadoOrigen, estadoDestino, tipoSolicitud, usuarioRol)
    
    ' Assert
    Test_ValidTransition_EnProceso_To_Aprobado_ReturnsTrue = resultado
End Function

Public Function Test_ValidTransition_EnProceso_To_Rechazado_ReturnsTrue() As Boolean
    ' Arrange
    Dim SolicitudID As Long
    Dim estadoOrigen As String
    Dim estadoDestino As String
    Dim tipoSolicitud As String
    Dim usuarioRol As String
    Dim resultado As Boolean
    
    Set workflowService = New CWorkflowService
    SolicitudID = 1
    estadoOrigen = "EnProceso"
    estadoDestino = "Rechazado"
    tipoSolicitud = "PC"
    usuarioRol = "Aprobador"
    
    ' Act
    resultado = workflowService.ValidateTransition(SolicitudID, estadoOrigen, estadoDestino, tipoSolicitud, usuarioRol)
    
    ' Assert
    Test_ValidTransition_EnProceso_To_Rechazado_ReturnsTrue = resultado
End Function

'******************************************************************************
' PRUEBAS DE TRANSICIONES INVÁLIDAS
'******************************************************************************

Public Function Test_InvalidTransition_Borrador_To_Aprobado_ReturnsFalse() As Boolean
    ' Arrange
    Dim SolicitudID As Long
    Dim estadoOrigen As String
    Dim estadoDestino As String
    Dim tipoSolicitud As String
    Dim usuarioRol As String
    Dim resultado As Boolean
    
    Set workflowService = New CWorkflowService
    SolicitudID = 1
    estadoOrigen = "Borrador"
    estadoDestino = "Aprobado"
    tipoSolicitud = "PC"
    usuarioRol = "Usuario"
    
    ' Act
    resultado = workflowService.ValidateTransition(SolicitudID, estadoOrigen, estadoDestino, tipoSolicitud, usuarioRol)
    
    ' Assert
    Test_InvalidTransition_Borrador_To_Aprobado_ReturnsFalse = Not resultado
End Function

Public Function Test_InvalidTransition_Aprobado_To_Borrador_ReturnsFalse() As Boolean
    ' Arrange
    Dim SolicitudID As Long
    Dim estadoOrigen As String
    Dim estadoDestino As String
    Dim tipoSolicitud As String
    Dim usuarioRol As String
    Dim resultado As Boolean
    
    Set workflowService = New CWorkflowService
    SolicitudID = 1
    estadoOrigen = "Aprobado"
    estadoDestino = "Borrador"
    tipoSolicitud = "PC"
    usuarioRol = "Usuario"
    
    ' Act
    resultado = workflowService.ValidateTransition(SolicitudID, estadoOrigen, estadoDestino, tipoSolicitud, usuarioRol)
    
    ' Assert
    Test_InvalidTransition_Aprobado_To_Borrador_ReturnsFalse = Not resultado
End Function

Public Function Test_InvalidTransition_NonExistentState_ReturnsFalse() As Boolean
    ' Arrange
    Dim SolicitudID As Long
    Dim estadoOrigen As String
    Dim estadoDestino As String
    Dim tipoSolicitud As String
    Dim usuarioRol As String
    Dim resultado As Boolean
    
    Set workflowService = New CWorkflowService
    SolicitudID = 1
    estadoOrigen = "EstadoInexistente"
    estadoDestino = "OtroEstadoInexistente"
    tipoSolicitud = "PC"
    usuarioRol = "Usuario"
    
    ' Act
    resultado = workflowService.ValidateTransition(SolicitudID, estadoOrigen, estadoDestino, tipoSolicitud, usuarioRol)
    
    ' Assert
    Test_InvalidTransition_NonExistentState_ReturnsFalse = Not resultado
End Function

'******************************************************************************
' PRUEBAS DE VALIDACIÓN DE PERMISOS
'******************************************************************************

Public Function Test_ValidatePermission_AdminRole_AllowsTransition_ReturnsTrue() As Boolean
    ' Arrange
    Dim SolicitudID As Long
    Dim estadoOrigen As String
    Dim estadoDestino As String
    Dim tipoSolicitud As String
    Dim usuarioRol As String
    Dim resultado As Boolean
    
    Set workflowService = New CWorkflowService
    SolicitudID = 1
    estadoOrigen = "EnProceso"
    estadoDestino = "Aprobado"
    tipoSolicitud = "PC"
    usuarioRol = "Administrador"
    
    ' Act
    resultado = workflowService.ValidateTransition(SolicitudID, estadoOrigen, estadoDestino, tipoSolicitud, usuarioRol)
    
    ' Assert
    Test_ValidatePermission_AdminRole_AllowsTransition_ReturnsTrue = resultado
End Function

Public Function Test_ValidatePermission_UserRole_DeniesTransition_ReturnsFalse() As Boolean
    ' Arrange
    Dim SolicitudID As Long
    Dim estadoOrigen As String
    Dim estadoDestino As String
    Dim tipoSolicitud As String
    Dim usuarioRol As String
    Dim resultado As Boolean
    
    Set workflowService = New CWorkflowService
    SolicitudID = 1
    estadoOrigen = "EnProceso"
    estadoDestino = "Aprobado"
    tipoSolicitud = "PC"
    usuarioRol = "Usuario"
    
    ' Act
    resultado = workflowService.ValidateTransition(SolicitudID, estadoOrigen, estadoDestino, tipoSolicitud, usuarioRol)
    
    ' Assert
    Test_ValidatePermission_UserRole_DeniesTransition_ReturnsFalse = Not resultado
End Function

Public Function Test_ValidatePermission_EmptyRole_ReturnsFalse() As Boolean
    ' Arrange
    Dim SolicitudID As Long
    Dim estadoOrigen As String
    Dim estadoDestino As String
    Dim tipoSolicitud As String
    Dim usuarioRol As String
    Dim resultado As Boolean
    
    Set workflowService = New CWorkflowService
    SolicitudID = 1
    estadoOrigen = "Borrador"
    estadoDestino = "EnProceso"
    tipoSolicitud = "PC"
    usuarioRol = ""
    
    ' Act
    resultado = workflowService.ValidateTransition(SolicitudID, estadoOrigen, estadoDestino, tipoSolicitud, usuarioRol)
    
    ' Assert
    Test_ValidatePermission_EmptyRole_ReturnsFalse = Not resultado
End Function

'******************************************************************************
' PRUEBAS DE OBTENCIÓN DE ESTADOS
'******************************************************************************

Public Function Test_GetAvailableStates_ForTipoPC_ReturnsCollection() As Boolean
    ' Arrange
    Dim tipoSolicitud As String
    Dim estados As Collection
    Dim resultado As Boolean
    
    Set workflowService = New CWorkflowService
    tipoSolicitud = "PC"
    
    ' Act
    Set estados = workflowService.GetAvailableStates(tipoSolicitud)
    
    ' Assert
    resultado = Not (estados Is Nothing) And estados.Count > 0
    Test_GetAvailableStates_ForTipoPC_ReturnsCollection = resultado
End Function

Public Function Test_GetNextStates_FromBorrador_ReturnsValidStates() As Boolean
    ' Arrange
    Dim estadoActual As String
    Dim tipoSolicitud As String
    Dim usuarioRol As String
    Dim estadosSiguientes As Collection
    Dim resultado As Boolean
    
    Set workflowService = New CWorkflowService
    estadoActual = "Borrador"
    tipoSolicitud = "PC"
    usuarioRol = "Usuario"
    
    ' Act
    Set estadosSiguientes = workflowService.GetNextStates(estadoActual, tipoSolicitud, usuarioRol)
    
    ' Assert
    resultado = Not (estadosSiguientes Is Nothing) And estadosSiguientes.Count > 0
    Test_GetNextStates_FromBorrador_ReturnsValidStates = resultado
End Function

'******************************************************************************
' PRUEBAS DE HISTORIAL DE ESTADOS
'******************************************************************************

Public Function Test_RecordStateChange_ValidTransition_CreatesHistoryRecord() As Boolean
    ' Arrange
    Dim SolicitudID As Long
    Dim estadoAnterior As String
    Dim estadoNuevo As String
    Dim usuario As String
    Dim comentarios As String
    Dim resultado As Boolean
    
    Set workflowService = New CWorkflowService
    SolicitudID = 1
    estadoAnterior = "Borrador"
    estadoNuevo = "EnProceso"
    usuario = "TestUser"
    comentarios = "Transición de prueba"
    
    ' Act
    resultado = workflowService.RecordStateChange(SolicitudID, estadoAnterior, estadoNuevo, usuario, comentarios)
    
    ' Assert
    Test_RecordStateChange_ValidTransition_CreatesHistoryRecord = resultado
End Function

#End If





