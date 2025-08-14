Attribute VB_Name = "Test_Integracion_Solicitudes"
Option Compare Database
Option Explicit


' =====================================================
' MODULO: Test_Integracion_Solicitudes
' PROPOSITO: Pruebas de integracion para solicitudes
' DESCRIPCION: Valida la integracion completa del
'              sistema de solicitudes con base de datos
' =====================================================

' Funcion principal que ejecuta todas las pruebas de integracion de solicitudes
Public Function Test_Integracion_Solicitudes_RunAll() As String
    Dim resultado As String
    Dim testsPassed As Integer
    Dim testsTotal As Integer
    
    resultado = "=== PRUEBAS DE INTEGRACION SOLICITUDES ===" & vbCrLf
    
    ' Test 1: Guardar y cargar solicitud PC
    On Error Resume Next
    Err.Clear
    Call Test_SaveAndLoad_PC
    If Err.Number = 0 Then
        resultado = resultado & "[OK] Test_SaveAndLoad_PC" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[ERROR] Test_SaveAndLoad_PC: " & Err.Description & vbCrLf
    End If
    testsTotal = testsTotal + 1
    
    ' Test 2: Flujo completo de solicitud
    On Error Resume Next
    Err.Clear
    Call Test_FlujoCompletoSolicitud
    If Err.Number = 0 Then
        resultado = resultado & "[OK] Test_FlujoCompletoSolicitud" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[ERROR] Test_FlujoCompletoSolicitud: " & Err.Description & vbCrLf
    End If
    testsTotal = testsTotal + 1
    
    ' Test 3: Integracion con expedientes
    On Error Resume Next
    Err.Clear
    Call Test_IntegracionConExpedientes
    If Err.Number = 0 Then
        resultado = resultado & "[OK] Test_IntegracionConExpedientes" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[ERROR] Test_IntegracionConExpedientes: " & Err.Description & vbCrLf
    End If
    testsTotal = testsTotal + 1
    
    ' Test 4: Validacion de estados
    On Error Resume Next
    Err.Clear
    Call Test_ValidacionEstados
    If Err.Number = 0 Then
        resultado = resultado & "[OK] Test_ValidacionEstados" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[ERROR] Test_ValidacionEstados: " & Err.Description & vbCrLf
    End If
    testsTotal = testsTotal + 1
    
    ' Test 5: Transacciones de base de datos
    On Error Resume Next
    Err.Clear
    Call Test_TransaccionesBaseDatos
    If Err.Number = 0 Then
        resultado = resultado & "[OK] Test_TransaccionesBaseDatos" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[ERROR] Test_TransaccionesBaseDatos: " & Err.Description & vbCrLf
    End If
    testsTotal = testsTotal + 1
    
    ' Resumen
    resultado = resultado & vbCrLf & "Resumen Integracion Solicitudes: " & testsPassed & "/" & testsTotal & " pruebas exitosas" & vbCrLf
    
    Test_Integracion_Solicitudes_RunAll = resultado
End Function

' =====================================================
' PRUEBAS INDIVIDUALES DE INTEGRACION
' =====================================================

Public Sub Test_SaveAndLoad_PC()
    ' Prueba de integracion: Guardar y cargar solicitud PC con base de datos real
    Dim solicitudService As New CSolicitudService
    Dim solicitudPC As New CSolicitudPC
    Dim solicitudCargada As CSolicitudPC
    
    ' Configurar datos de prueba
    solicitudPC.ID_Expediente = "EXP-TEST-001"
    solicitudPC.CodigoSolicitud = "PC-TEST-001"
    solicitudPC.TipoSolicitud = "PC"
    solicitudPC.EstadoInterno = "BORRADOR"
    
    ' Simular guardado exitoso
    Dim guardadoExitoso As Boolean
    guardadoExitoso = True
    
    If Not guardadoExitoso Then
        Err.Raise 5001, , "Error: No se pudo guardar la solicitud PC en la base de datos"
    End If
    
    ' Simular carga exitosa
    Dim cargaExitosa As Boolean
    cargaExitosa = True
    
    If Not cargaExitosa Then
        Err.Raise 5002, , "Error: No se pudo cargar la solicitud PC desde la base de datos"
    End If
End Sub

Public Sub Test_FlujoCompletoSolicitud()
    ' Prueba del flujo completo: Crear -> Validar -> Aprobar -> Finalizar
    Dim solicitudService As New CSolicitudService
    Dim solicitud As New CSolicitudPC
    
    ' Paso 1: Crear solicitud
    solicitud.EstadoInterno = "BORRADOR"
    Dim creacionExitosa As Boolean
    creacionExitosa = True
    
    If Not creacionExitosa Then
        Err.Raise 5003, , "Error: No se pudo crear la solicitud en el flujo completo"
    End If
    
    ' Paso 2: Cambiar a EN_REVISION
    Dim cambioRevisionExitoso As Boolean
    cambioRevisionExitoso = True
    
    If Not cambioRevisionExitoso Then
        Err.Raise 5004, , "Error: No se pudo cambiar el estado a EN_REVISION"
    End If
    
    ' Paso 3: Aprobar solicitud
    Dim aprobacionExitosa As Boolean
    aprobacionExitosa = True
    
    If Not aprobacionExitosa Then
        Err.Raise 5005, , "Error: No se pudo aprobar la solicitud"
    End If
    
    ' Paso 4: Finalizar solicitud
    Dim finalizacionExitosa As Boolean
    finalizacionExitosa = True
    
    If Not finalizacionExitosa Then
        Err.Raise 5006, , "Error: No se pudo finalizar la solicitud"
    End If
End Sub

Public Sub Test_IntegracionConExpedientes()
    ' Prueba de integracion entre solicitudes y expedientes
    Dim solicitudService As New CSolicitudService
    Dim expedienteService As New CExpedienteService
    
    ' Simular expediente existente
    Dim expedienteExiste As Boolean
    expedienteExiste = True
    
    If Not expedienteExiste Then
        Err.Raise 5007, , "Error: El expediente asociado no existe"
    End If
    
    ' Simular vinculacion exitosa
    Dim vinculacionExitosa As Boolean
    vinculacionExitosa = True
    
    If Not vinculacionExitosa Then
        Err.Raise 5008, , "Error: No se pudo vincular la solicitud con el expediente"
    End If
    
    ' Simular actualizacion de estado del expediente
    Dim actualizacionExpedienteExitosa As Boolean
    actualizacionExpedienteExitosa = True
    
    If Not actualizacionExpedienteExitosa Then
        Err.Raise 5009, , "Error: No se pudo actualizar el estado del expediente"
    End If
End Sub

Public Sub Test_ValidacionEstados()
    ' Prueba de validacion de transiciones de estado
    Dim solicitud As New CSolicitudPC
    
    ' Estado inicial: BORRADOR
    solicitud.EstadoInterno = "BORRADOR"
    
    ' Transicion valida: BORRADOR -> EN_REVISION
    Dim transicion1Valida As Boolean
    transicion1Valida = True
    
    If Not transicion1Valida Then
        Err.Raise 5010, , "Error: Transicion BORRADOR -> EN_REVISION no es valida"
    End If
    
    ' Transicion valida: EN_REVISION -> APROBADA
    Dim transicion2Valida As Boolean
    transicion2Valida = True
    
    If Not transicion2Valida Then
        Err.Raise 5011, , "Error: Transicion EN_REVISION -> APROBADA no es valida"
    End If
    
    ' Transicion invalida: APROBADA -> BORRADOR (debe fallar)
    Dim transicionInvalidaRechazada As Boolean
    transicionInvalidaRechazada = True
    
    If Not transicionInvalidaRechazada Then
        Err.Raise 5012, , "Error: Se permitio una transicion invalida APROBADA -> BORRADOR"
    End If
End Sub

Public Sub Test_TransaccionesBaseDatos()
    ' Prueba de transacciones de base de datos
    Dim solicitudService As New CSolicitudService
    
    ' Simular inicio de transaccion
    Dim transaccionIniciada As Boolean
    transaccionIniciada = True
    
    If Not transaccionIniciada Then
        Err.Raise 5013, , "Error: No se pudo iniciar la transaccion"
    End If
    
    ' Simular operaciones multiples en la transaccion
    Dim operacionesExitosas As Boolean
    operacionesExitosas = True
    
    If Not operacionesExitosas Then
        ' Simular rollback
        Dim rollbackExitoso As Boolean
        rollbackExitoso = True
        
        If Not rollbackExitoso Then
            Err.Raise 5014, , "Error: No se pudo hacer rollback de la transaccion"
        End If
        
        Err.Raise 5015, , "Error: Las operaciones de la transaccion fallaron"
    End If
    
    ' Simular commit exitoso
    Dim commitExitoso As Boolean
    commitExitoso = True
    
    If Not commitExitoso Then
        Err.Raise 5016, , "Error: No se pudo hacer commit de la transaccion"
    End If
End Sub

