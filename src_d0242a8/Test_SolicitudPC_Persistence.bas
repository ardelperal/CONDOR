Attribute VB_Name = "Test_SolicitudPC_Persistence"
'******************************************************************************
' MODULO: Test_SolicitudPC_Persistence
' DESCRIPCION: Pruebas unitarias para la persistencia de CSolicitudPC
' AUTOR: Sistema CONDOR
' FECHA: 2024
'******************************************************************************

#If DEV_MODE Then

Option Explicit

'******************************************************************************
' PRUEBAS DE PERSISTENCIA
'******************************************************************************

Public Function Test_Save_PC_ShouldDelegateToRepository() As Boolean
    ' Arrange
    Dim solicitud As CSolicitudPC
    Dim mockRepo As ISolicitudRepository
    Dim datosPC As T_Datos_PC
    Dim resultado As Long
    
    Set solicitud = New CSolicitudPC
    Set mockRepo = modRepositoryFactory.CreateSolicitudRepository()
    
    ' Configurar datos de prueba
    solicitud.IDExpediente = 1001
    solicitud.TipoSolicitud = "PC"
    solicitud.CodigoSolicitud = "PC-TEST-001"
    solicitud.EstadoInterno = "Pendiente"
    
    datosPC.Procesador = "Intel i7-12700K"
    datosPC.RAM = "32GB DDR4"
    datosPC.Almacenamiento = "1TB NVMe SSD"
    datosPC.SistemaOperativo = "Windows 11 Pro"
    solicitud.DatosPC = datosPC
    
    ' Act
    resultado = solicitud.Save
    
    ' Assert
    If resultado > 0 Then
        Debug.Print "Test_Save_PC_ShouldDelegateToRepository: PASO"
        Test_Save_PC_ShouldDelegateToRepository = True
    Else
        Debug.Print "Test_Save_PC_ShouldDelegateToRepository: FALLO - No se guardo correctamente"
        Test_Save_PC_ShouldDelegateToRepository = False
    End If
End Function

Public Function Test_Load_PC_ShouldPopulateObjectFromRepository() As Boolean
    ' Arrange
    Dim solicitud As CSolicitudPC
    Dim mockRepo As ISolicitudRepository
    Dim resultado As Boolean
    
    Set solicitud = New CSolicitudPC
    Set mockRepo = modRepositoryFactory.CreateSolicitudRepository()
    
    ' Act - Cargar solicitud de prueba predefinida (ID 999)
    resultado = solicitud.Load(999)
    
    ' Assert
    If resultado And solicitud.idSolicitud = 999 And solicitud.CodigoSolicitud = "PC-TEST-001" Then
        Debug.Print "Test_Load_PC_ShouldPopulateObjectFromRepository: PASO"
        Test_Load_PC_ShouldPopulateObjectFromRepository = True
    Else
        Debug.Print "Test_Load_PC_ShouldPopulateObjectFromRepository: FALLO - No se cargo correctamente"
        Debug.Print "  ID: " & solicitud.idSolicitud & ", Codigo: " & solicitud.CodigoSolicitud
        Test_Load_PC_ShouldPopulateObjectFromRepository = False
    End If
End Function

Public Function Test_SaveAndLoad_PC_ShouldMaintainDataIntegrity() As Boolean
    ' Arrange
    Dim solicitudOriginal As CSolicitudPC
    Dim solicitudCargada As CSolicitudPC
    Dim mockRepo As ISolicitudRepository
    Dim datosPC As T_Datos_PC
    Dim savedId As Long
    Dim loadResult As Boolean
    
    Set solicitudOriginal = New CSolicitudPC
    Set solicitudCargada = New CSolicitudPC
    Set mockRepo = modRepositoryFactory.CreateSolicitudRepository()
    
    ' Configurar datos de prueba
    solicitudOriginal.IDExpediente = 2002
    solicitudOriginal.TipoSolicitud = "PC"
    solicitudOriginal.CodigoSolicitud = "PC-INTEGRATION-001"
    solicitudOriginal.EstadoInterno = "En Proceso"
    
    datosPC.Procesador = "AMD Ryzen 9 5900X"
    datosPC.RAM = "64GB DDR4"
    datosPC.Almacenamiento = "2TB NVMe SSD"
    datosPC.SistemaOperativo = "Windows 11 Enterprise"
    solicitudOriginal.DatosPC = datosPC
    
    ' La inyeccion es automatica a traves del Factory en Class_Initialize
    
    ' Act
    savedId = solicitudOriginal.Save
    loadResult = solicitudCargada.Load(savedId)
    
    ' Assert
    If loadResult And _
       solicitudCargada.IDExpediente = solicitudOriginal.IDExpediente And _
       solicitudCargada.CodigoSolicitud = solicitudOriginal.CodigoSolicitud And _
       solicitudCargada.DatosPC.Procesador = solicitudOriginal.DatosPC.Procesador Then
        Debug.Print "Test_SaveAndLoad_PC_ShouldMaintainDataIntegrity: PASO"
        Test_SaveAndLoad_PC_ShouldMaintainDataIntegrity = True
    Else
        Debug.Print "Test_SaveAndLoad_PC_ShouldMaintainDataIntegrity: FALLO - Integridad de datos comprometida"
        Debug.Print "  Original ID Expediente: " & solicitudOriginal.IDExpediente & ", Cargado: " & solicitudCargada.IDExpediente
        Debug.Print "  Original Codigo: " & solicitudOriginal.CodigoSolicitud & ", Cargado: " & solicitudCargada.CodigoSolicitud
        Test_SaveAndLoad_PC_ShouldMaintainDataIntegrity = False
    End If
End Function

'******************************************************************************
' SUITE DE PRUEBAS
'******************************************************************************

' ============================================================================
' FUNCI├ôN PRINCIPAL PARA EJECUTAR TODAS LAS PRUEBAS
' ============================================================================

Public Function Test_SolicitudPC_Persistence_RunAll() As String
    Dim resultado As String
    Dim testsPassed As Long, testsTotal As Long
    
    resultado = "=== PRUEBAS DE SOLICITUDPC PERSISTENCE ===" & vbCrLf
    testsPassed = 0
    testsTotal = 0
    
    ' Ejecutar todas las pruebas
    testsTotal = testsTotal + 1
    If Test_Save_PC_ShouldDelegateToRepository() Then
        resultado = resultado & "[OK] Test_Save_PC_ShouldDelegateToRepository" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_Save_PC_ShouldDelegateToRepository" & vbCrLf
    End If
    
    testsTotal = testsTotal + 1
    If Test_Load_PC_ShouldPopulateObjectFromRepository() Then
        resultado = resultado & "[OK] Test_Load_PC_ShouldPopulateObjectFromRepository" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_Load_PC_ShouldPopulateObjectFromRepository" & vbCrLf
    End If
    
    testsTotal = testsTotal + 1
    If Test_SaveAndLoad_PC_ShouldMaintainDataIntegrity() Then
        resultado = resultado & "[OK] Test_SaveAndLoad_PC_ShouldMaintainDataIntegrity" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_SaveAndLoad_PC_ShouldMaintainDataIntegrity" & vbCrLf
    End If
    
    ' Agregar resumen
    resultado = resultado & vbCrLf & "RESUMEN: " & testsPassed & "/" & testsTotal & " pruebas pasadas" & vbCrLf
    
    Test_SolicitudPC_Persistence_RunAll = resultado
End Function

Public Sub EJECUTAR_PRUEBAS_PERSISTENCIA_PC()
    Debug.Print "=== INICIANDO PRUEBAS DE PERSISTENCIA PC ==="
    
    Test_Save_PC_ShouldDelegateToRepository
    Test_Load_PC_ShouldPopulateObjectFromRepository
    Test_SaveAndLoad_PC_ShouldMaintainDataIntegrity
    
    Debug.Print "=== PRUEBAS DE PERSISTENCIA PC COMPLETADAS ==="
End Sub

#End If
