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

Public Sub Test_Save_PC_ShouldDelegateToRepository()
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
    Else
        Debug.Print "Test_Save_PC_ShouldDelegateToRepository: FALLO - No se guardo correctamente"
    End If
End Sub

Public Sub Test_Load_PC_ShouldPopulateObjectFromRepository()
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
    Else
        Debug.Print "Test_Load_PC_ShouldPopulateObjectFromRepository: FALLO - No se cargo correctamente"
        Debug.Print "  ID: " & solicitud.idSolicitud & ", Codigo: " & solicitud.CodigoSolicitud
    End If
End Sub

Public Sub Test_SaveAndLoad_PC_ShouldMaintainDataIntegrity()
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
    Else
        Debug.Print "Test_SaveAndLoad_PC_ShouldMaintainDataIntegrity: FALLO - Integridad de datos comprometida"
        Debug.Print "  Original ID Expediente: " & solicitudOriginal.IDExpediente & ", Cargado: " & solicitudCargada.IDExpediente
        Debug.Print "  Original Codigo: " & solicitudOriginal.CodigoSolicitud & ", Cargado: " & solicitudCargada.CodigoSolicitud
    End If
End Sub

'******************************************************************************
' SUITE DE PRUEBAS
'******************************************************************************

Public Sub EJECUTAR_PRUEBAS_PERSISTENCIA_PC()
    Debug.Print "=== INICIANDO PRUEBAS DE PERSISTENCIA PC ==="
    
    Test_Save_PC_ShouldDelegateToRepository
    Test_Load_PC_ShouldPopulateObjectFromRepository
    Test_SaveAndLoad_PC_ShouldMaintainDataIntegrity
    
    Debug.Print "=== PRUEBAS DE PERSISTENCIA PC COMPLETADAS ==="
End Sub

#End If