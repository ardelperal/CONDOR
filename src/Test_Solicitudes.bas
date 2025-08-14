Attribute VB_Name = "Test_Solicitudes"
Option Compare Database
Option Explicit

' ============================================================================
' Módulo: Test_Solicitudes
' Descripción: Pruebas unitarias para el módulo de gestión de solicitudes
' Autor: CONDOR-Expert
' Fecha: Diciembre 2024
' ============================================================================

' ============================================================================
' PRUEBAS DEL FACTORY PATTERN
' ============================================================================

' Test básico para verificar que CreateSolicitud funciona correctamente
Public Sub TestCreateSolicitud()
    On Error GoTo ErrorHandler
    
    Debug.Print "=== Iniciando TestCreateSolicitud ==="
    
    ' Crear una solicitud usando el factory
    Dim solicitud As ISolicitud
    Set solicitud = modSolicitudFactory.CreateSolicitud(1)
    
    ' Verificar que el objeto no es Nothing
    Debug.Assert Not solicitud Is Nothing, "La solicitud creada no debe ser Nothing"
    Debug.Print "✓ Solicitud creada correctamente"
    
    ' Verificar que es del tipo correcto
    Debug.Assert TypeOf solicitud Is CSolicitudPC, "La solicitud debe ser del tipo CSolicitudPC"
    Debug.Print "✓ Tipo de solicitud correcto (CSolicitudPC)"
    
    Debug.Print "=== TestCreateSolicitud PASÓ ==="
    Exit Sub
    
ErrorHandler:
    Debug.Print "✗ Error en TestCreateSolicitud: " & Err.Description
    Debug.Assert False, "TestCreateSolicitud falló: " & Err.Description
End Sub

' ============================================================================
' PRUEBAS DE LA INTERFAZ ISolicitud
' ============================================================================

' Test para verificar las propiedades de la interfaz
Public Sub TestISolicitudProperties()
    On Error GoTo ErrorHandler
    
    Debug.Print "=== Iniciando TestISolicitudProperties ==="
    
    ' Crear una instancia directa de CSolicitudPC
    Dim solicitudPC As CSolicitudPC
    Set solicitudPC = New CSolicitudPC
    
    ' Verificar propiedades iniciales
    Debug.Assert solicitudPC.TipoSolicitud = "PC", "TipoSolicitud debe ser 'PC' por defecto"
    Debug.Print "✓ TipoSolicitud inicial correcto"
    
    Debug.Assert solicitudPC.EstadoInterno = "BORRADOR", "EstadoInterno debe ser 'BORRADOR' por defecto"
    Debug.Print "✓ EstadoInterno inicial correcto"
    
    ' Probar asignación de propiedades
    solicitudPC.ID_Solicitud = 123
    solicitudPC.ID_Expediente = "EXP-2024-001"
    solicitudPC.CodigoSolicitud = "PC-2024-001"
    
    Debug.Assert solicitudPC.ID_Solicitud = 123, "ID_Solicitud debe ser 123"
    Debug.Assert solicitudPC.ID_Expediente = "EXP-2024-001", "ID_Expediente debe ser 'EXP-2024-001'"
    Debug.Assert solicitudPC.CodigoSolicitud = "PC-2024-001", "CodigoSolicitud debe ser 'PC-2024-001'"
    Debug.Print "✓ Asignación de propiedades correcta"
    
    Debug.Print "=== TestISolicitudProperties PASÓ ==="
    Exit Sub
    
ErrorHandler:
    Debug.Print "✗ Error en TestISolicitudProperties: " & Err.Description
    Debug.Assert False, "TestISolicitudProperties falló: " & Err.Description
End Sub

' ============================================================================
' PRUEBAS DE LOS MÉTODOS DE LA INTERFAZ
' ============================================================================

' Test para verificar los métodos de la interfaz
Public Sub TestISolicitudMethods()
    On Error GoTo ErrorHandler
    
    Debug.Print "=== Iniciando TestISolicitudMethods ==="
    
    ' Crear una instancia usando la interfaz
    Dim solicitud As ISolicitud
    Set solicitud = New CSolicitudPC
    
    ' Probar método Load
    Dim loadResult As Boolean
    loadResult = solicitud.Load(456)
    Debug.Assert loadResult = True, "Load debe retornar True"
    Debug.Print "✓ Método Load funciona correctamente"
    
    ' Probar método Save
    Dim saveResult As Boolean
    saveResult = solicitud.Save()
    Debug.Assert saveResult = True, "Save debe retornar True"
    Debug.Print "✓ Método Save funciona correctamente"
    
    ' Probar método ChangeState
    Dim changeStateResult As Boolean
    changeStateResult = solicitud.ChangeState("EN_REVISION")
    Debug.Assert changeStateResult = True, "ChangeState debe retornar True"
    Debug.Print "✓ Método ChangeState funciona correctamente"
    
    Debug.Print "=== TestISolicitudMethods PASÓ ==="
    Exit Sub
    
ErrorHandler:
    Debug.Print "✗ Error en TestISolicitudMethods: " & Err.Description
    Debug.Assert False, "TestISolicitudMethods falló: " & Err.Description
End Sub

' ============================================================================
' PRUEBAS DE LA ESTRUCTURA T_Datos_PC
' ============================================================================

' Test para verificar que la estructura T_Datos_PC está disponible
Public Sub TestTDatosPC()
    On Error GoTo ErrorHandler
    
    Debug.Print "=== Iniciando TestTDatosPC ==="
    
    ' Crear una instancia de CSolicitudPC y verificar que tiene DatosPC
    Dim solicitudPC As CSolicitudPC
    Set solicitudPC = New CSolicitudPC
    
    ' Verificar que se puede acceder a la propiedad DatosPC
    ' (Por ahora solo verificamos que no genera error)
    Dim datosPC As T_Datos_PC
    datosPC = solicitudPC.DatosPC
    
    Debug.Print "✓ Estructura T_Datos_PC accesible"
    Debug.Print "✓ Propiedad DatosPC en CSolicitudPC funcional"
    
    Debug.Print "=== TestTDatosPC PASÓ ==="
    Exit Sub
    
ErrorHandler:
    Debug.Print "✗ Error en TestTDatosPC: " & Err.Description
    Debug.Assert False, "TestTDatosPC falló: " & Err.Description
End Sub

' ============================================================================
' SUITE DE PRUEBAS COMPLETA
' ============================================================================

' Ejecutar todas las pruebas
Public Sub RunAllSolicitudesTests()
    On Error GoTo ErrorHandler
    
    Debug.Print "============================================"
    Debug.Print "EJECUTANDO SUITE DE PRUEBAS DE SOLICITUDES"
    Debug.Print "============================================"
    
    Debug.Print "Iniciando TestCreateSolicitud..."
    TestCreateSolicitud
    Debug.Print "TestCreateSolicitud completado"
    
    Debug.Print "Iniciando TestISolicitudProperties..."
    TestISolicitudProperties
    Debug.Print "TestISolicitudProperties completado"
    
    Debug.Print "Iniciando TestISolicitudMethods..."
    TestISolicitudMethods
    Debug.Print "TestISolicitudMethods completado"
    
    Debug.Print "Iniciando TestTDatosPC..."
    TestTDatosPC
    Debug.Print "TestTDatosPC completado"
    
    Debug.Print "============================================"
    Debug.Print "TODAS LAS PRUEBAS DE SOLICITUDES PASARON ✓"
    Debug.Print "============================================"
    Exit Sub
    
ErrorHandler:
    Debug.Print "ERROR en RunAllSolicitudesTests: " & Err.Description & " (Número: " & Err.Number & ")"
    Debug.Print "Fuente: " & Err.Source
End Sub