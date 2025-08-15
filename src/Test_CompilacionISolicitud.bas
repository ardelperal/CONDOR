Attribute VB_Name = "Test_CompilacionISolicitud"
Option Compare Database
Option Explicit

' ============================================================================
' M?dulo: Test_CompilacionISolicitud
' Descripci?n: Prueba de compilaci?n para verificar la implementaci?n de ISolicitud
' Autor: CONDOR-Expert
' Fecha: Diciembre 2024
' ============================================================================

' Funci?n de prueba para verificar que la implementaci?n de ISolicitud funciona
Public Function Test_ImplementacionISolicitud() As Boolean
    On Error GoTo ErrorHandler
    
    Dim solicitud As ISolicitud
    Dim solicitudPC As CSolicitudPC
    
    ' Crear instancia de CSolicitudPC
    Set solicitudPC = New CSolicitudPC
    
    ' Asignar a la interfaz
    Set solicitud = solicitudPC
    
    ' Probar propiedades de la interfaz
    solicitud.idSolicitud = 123
    solicitud.IDExpediente = "EXP-001"
    solicitud.TipoSolicitud = "PC"
    solicitud.CodigoSolicitud = "PC-0123"
    solicitud.EstadoInterno = "BORRADOR"
    
    ' Verificar que los valores se asignaron correctamente
    If solicitud.idSolicitud <> 123 Then GoTo ErrorHandler
    If solicitud.IDExpediente <> "EXP-001" Then GoTo ErrorHandler
    If solicitud.TipoSolicitud <> "PC" Then GoTo ErrorHandler
    If solicitud.CodigoSolicitud <> "PC-0123" Then GoTo ErrorHandler
    If solicitud.EstadoInterno <> "BORRADOR" Then GoTo ErrorHandler
    
    ' Probar m?todos de la interfaz
    ' Nota: Estos m?todos pueden fallar por falta de datos, pero no deben dar error de compilaci?n
    Dim loadResult As Boolean
    Dim saveResult As Boolean
    Dim changeStateResult As Boolean
    
    loadResult = solicitud.Load(1)
    saveResult = solicitud.Save()
    changeStateResult = solicitud.ChangeState("ENVIADO")
    
    Test_ImplementacionISolicitud = True
    
    ' Limpiar objetos
    Set solicitud = Nothing
    Set solicitudPC = Nothing
    
    Exit Function
    
ErrorHandler:
    Test_ImplementacionISolicitud = False
    Set solicitud = Nothing
    Set solicitudPC = Nothing
End Function

' Funci?n para ejecutar la prueba y mostrar resultado
Public Sub Ejecutar_Test_Compilacion()
    If Test_ImplementacionISolicitud() Then
        Debug.Print "? Test de compilaci?n ISolicitud: EXITOSO"
        MsgBox "Test de compilaci?n ISolicitud: EXITOSO", vbInformation
    Else
        Debug.Print "? Test de compilaci?n ISolicitud: FALL?"
        MsgBox "Test de compilaci?n ISolicitud: FALL?", vbCritical
    End If
End Sub



