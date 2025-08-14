Attribute VB_Name = "Test_Solicitudes"
Option Compare Database
Option Explicit


' M�dulo de Pruebas para el Servicio de Solicitudes
' Utiliza CMockSolicitudService para pruebas sin dependencias de base de datos
' Version: 1.0
' Fecha: 2025-01-14


' Funci�n principal que ejecuta todas las pruebas del servicio de solicitudes
Public Function Test_Solicitudes_RunAll() As String
    Dim resultado As String
    Dim totalPruebas As Integer
    Dim pruebasExitosas As Integer
    
    resultado = "=== PRUEBAS DEL SERVICIO DE SOLICITUDES ===" & vbCrLf
    totalPruebas = 0
    pruebasExitosas = 0
    
    ' Ejecutar todas las pruebas
    If Test_CreateNuevaSolicitud() Then pruebasExitosas = pruebasExitosas + 1
    totalPruebas = totalPruebas + 1
    
    If Test_GetSolicitudPorID() Then pruebasExitosas = pruebasExitosas + 1
    totalPruebas = totalPruebas + 1
    
    If Test_SaveSolicitud() Then pruebasExitosas = pruebasExitosas + 1
    totalPruebas = totalPruebas + 1
    
    If Test_GetAllSolicitudes() Then pruebasExitosas = pruebasExitosas + 1
    totalPruebas = totalPruebas + 1
    
    If Test_DeleteSolicitud() Then pruebasExitosas = pruebasExitosas + 1
    totalPruebas = totalPruebas + 1
    
    If Test_UpdateEstadoSolicitud() Then pruebasExitosas = pruebasExitosas + 1
    totalPruebas = totalPruebas + 1
    
    ' Resumen de resultados
    resultado = resultado & vbCrLf & "RESUMEN:" & vbCrLf
    resultado = resultado & "Total de pruebas: " & totalPruebas & vbCrLf
    resultado = resultado & "Pruebas exitosas: " & pruebasExitosas & vbCrLf
    resultado = resultado & "Pruebas fallidas: " & (totalPruebas - pruebasExitosas) & vbCrLf
    
    If pruebasExitosas = totalPruebas Then
        resultado = resultado & "RESULTADO: TODAS LAS PRUEBAS PASARON" & vbCrLf
    Else
        resultado = resultado & "RESULTADO: ALGUNAS PRUEBAS FALLARON" & vbCrLf
    End If
    
    Test_Solicitudes_RunAll = resultado
End Function

' Prueba: Crear nueva solicitud
Private Function Test_CreateNuevaSolicitud() As Boolean
    On Error GoTo ErrorHandler
    
    Dim solicitudService As New CMockSolicitudService
    Dim solicitud As ISolicitud
    
    ' Probar creaci�n de solicitud tipo PC
    Set solicitud = solicitudService.CreateNuevaSolicitud("PC")
    
    ' Verificar que se cre� la solicitud
    If Not solicitud Is Nothing Then
        Debug.Print "? Test_CreateNuevaSolicitud: PAS�"
        Test_CreateNuevaSolicitud = True
    Else
        Debug.Print "? Test_CreateNuevaSolicitud: FALL� - No se cre� la solicitud"
        Test_CreateNuevaSolicitud = False
    End If
    
    Exit Function
    
ErrorHandler:
    Debug.Print "? Test_CreateNuevaSolicitud: ERROR - " & Err.Description
    Test_CreateNuevaSolicitud = False
End Function

' Prueba: Obtener solicitud por ID
Private Function Test_GetSolicitudPorID() As Boolean
    On Error GoTo ErrorHandler
    
    Dim solicitudService As New CMockSolicitudService
    Dim solicitud As ISolicitud
    
    ' Probar obtener solicitud con ID v�lido
    Set solicitud = solicitudService.GetSolicitudPorID(1)
    
    If Not solicitud Is Nothing Then
        Debug.Print "? Test_GetSolicitudPorID: PAS�"
        Test_GetSolicitudPorID = True
    Else
        Debug.Print "? Test_GetSolicitudPorID: FALL� - No se encontr� la solicitud"
        Test_GetSolicitudPorID = False
    End If
    
    Exit Function
    
ErrorHandler:
    Debug.Print "? Test_GetSolicitudPorID: ERROR - " & Err.Description
    Test_GetSolicitudPorID = False
End Function

' Prueba: Guardar solicitud
Private Function Test_SaveSolicitud() As Boolean
    On Error GoTo ErrorHandler
    
    Dim solicitudService As New CMockSolicitudService
    Dim solicitud As New CSolicitudPC
    Dim resultado As Boolean
    
    ' Probar guardar solicitud
    resultado = solicitudService.SaveSolicitud(solicitud)
    
    If resultado Then
        Debug.Print "? Test_SaveSolicitud: PAS�"
        Test_SaveSolicitud = True
    Else
        Debug.Print "? Test_SaveSolicitud: FALL� - No se pudo guardar la solicitud"
        Test_SaveSolicitud = False
    End If
    
    Exit Function
    
ErrorHandler:
    Debug.Print "? Test_SaveSolicitud: ERROR - " & Err.Description
    Test_SaveSolicitud = False
End Function

' Prueba: Obtener todas las solicitudes
Private Function Test_GetAllSolicitudes() As Boolean
    On Error GoTo ErrorHandler
    
    Dim solicitudService As New CMockSolicitudService
    Dim solicitudes As Collection
    
    ' Probar obtener todas las solicitudes
    Set solicitudes = solicitudService.GetAllSolicitudes()
    
    If Not solicitudes Is Nothing And solicitudes.Count > 0 Then
        Debug.Print "? Test_GetAllSolicitudes: PAS� - " & solicitudes.Count & " solicitudes encontradas"
        Test_GetAllSolicitudes = True
    Else
        Debug.Print "? Test_GetAllSolicitudes: FALL� - No se encontraron solicitudes"
        Test_GetAllSolicitudes = False
    End If
    
    Exit Function
    
ErrorHandler:
    Debug.Print "? Test_GetAllSolicitudes: ERROR - " & Err.Description
    Test_GetAllSolicitudes = False
End Function

' Prueba: Eliminar solicitud
Private Function Test_DeleteSolicitud() As Boolean
    On Error GoTo ErrorHandler
    
    Dim solicitudService As New CMockSolicitudService
    Dim resultado As Boolean
    
    ' Probar eliminar solicitud con ID v�lido
    resultado = solicitudService.DeleteSolicitud(1)
    
    If resultado Then
        Debug.Print "? Test_DeleteSolicitud: PAS�"
        Test_DeleteSolicitud = True
    Else
        Debug.Print "? Test_DeleteSolicitud: FALL� - No se pudo eliminar la solicitud"
        Test_DeleteSolicitud = False
    End If
    
    Exit Function
    
ErrorHandler:
    Debug.Print "? Test_DeleteSolicitud: ERROR - " & Err.Description
    Test_DeleteSolicitud = False
End Function

' Prueba: Actualizar estado de solicitud
Private Function Test_UpdateEstadoSolicitud() As Boolean
    On Error GoTo ErrorHandler
    
    Dim solicitudService As New CMockSolicitudService
    Dim resultado As Boolean
    
    ' Probar actualizar estado con par�metros v�lidos
    resultado = solicitudService.UpdateEstadoSolicitud(1, "APROBADA")
    
    If resultado Then
        Debug.Print "? Test_UpdateEstadoSolicitud: PAS�"
        Test_UpdateEstadoSolicitud = True
    Else
        Debug.Print "? Test_UpdateEstadoSolicitud: FALL� - No se pudo actualizar el estado"
        Test_UpdateEstadoSolicitud = False
    End If
    
    Exit Function
    
ErrorHandler:
    Debug.Print "? Test_UpdateEstadoSolicitud: ERROR - " & Err.Description
    Test_UpdateEstadoSolicitud = False
End Function
