Attribute VB_Name = "Test_CompilacionISolicitud"
Option Compare Database
Option Explicit




' ============================================================================
' Módulo: Test_CompilacionISolicitud
' Descripción: Prueba de compilación para verificar la implementación de ISolicitud
' Autor: CONDOR-Expert
' Fecha: Diciembre 2024
' ============================================================================

' Función de prueba para verificar que la implementación de ISolicitud funciona
Public Function Test_ImplementacionISolicitud() As Boolean
    On Error GoTo ErrorHandler
    
    Dim solicitud As ISolicitud
    Dim solicitudPC As CSolicitudPC
    
    ' Crear instancia de CSolicitudPC
    Set solicitudPC = New CSolicitudPC
    
    ' Asignar a la interfaz
    Set solicitud = solicitudPC
    
    ' Probar propiedades de la interfaz
    solicitud.ID_Solicitud = 123
    solicitud.ID_Expediente = "EXP-001"
    solicitud.TipoSolicitud = "PC"
    solicitud.CodigoSolicitud = "PC-0123"
    solicitud.EstadoInterno = "BORRADOR"
    
    ' Verificar que los valores se asignaron correctamente
    If solicitud.ID_Solicitud <> 123 Then GoTo ErrorHandler
    If solicitud.ID_Expediente <> "EXP-001" Then GoTo ErrorHandler
    If solicitud.TipoSolicitud <> "PC" Then GoTo ErrorHandler
    If solicitud.CodigoSolicitud <> "PC-0123" Then GoTo ErrorHandler
    If solicitud.EstadoInterno <> "BORRADOR" Then GoTo ErrorHandler
    
    ' Probar métodos de la interfaz
    ' Nota: Estos métodos pueden fallar por falta de datos, pero no deben dar error de compilación
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

' Función para ejecutar la prueba y mostrar resultado
Public Sub Ejecutar_Test_Compilacion()
    If Test_ImplementacionISolicitud() Then
        Debug.Print "? Test de compilación ISolicitud: EXITOSO"
        MsgBox "Test de compilación ISolicitud: EXITOSO", vbInformation
    Else
        Debug.Print "? Test de compilación ISolicitud: FALLÓ"
        MsgBox "Test de compilación ISolicitud: FALLÓ", vbCritical
    End If
End Sub
