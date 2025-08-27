Option Compare Database
Option Explicit

#If DEV_MODE Then

' ============================================================================
' PRUEBAS DE INTEGRACIÓN PARA CSolicitudRepository
' ============================================================================
' Este módulo contiene pruebas de integración que validan el comportamiento
' del repositorio CSolicitudRepository contra una base de datos real.
' A diferencia de las pruebas unitarias, estas pruebas verifican la interacción
' completa con la capa de datos.

' Constantes para las rutas de base de datos de prueba
Private Const CONDOR_TEMPLATE_PATH As String = "back\test_db\templates\CONDOR_test_template.accdb"
Private Const CONDOR_ACTIVE_PATH As String = "back\test_db\active\CONDOR_integration_test.accdb"

' Función principal que ejecuta todas las pruebas del módulo
Public Function IntegrationTest_SolicitudRepository_RunAll() As CTestSuiteResult
    Dim suiteResult As New CTestSuiteResult
    suiteResult.Initialize "IntegrationTest_SolicitudRepository"
    
    ' Ejecutar todas las pruebas de integración
    suiteResult.AddTestResult Test_GetSolicitudById_Success()
    suiteResult.AddTestResult Test_GetSolicitudById_NotFound()
    suiteResult.AddTestResult Test_SaveSolicitud_New()
    suiteResult.AddTestResult Test_SaveSolicitud_Update()
    suiteResult.AddTestResult Test_ExecuteQuery()
    suiteResult.AddTestResult Test_CargarDatosEspecificos_PC()
    suiteResult.AddTestResult Test_CargarDatosEspecificos_CDCA()
    suiteResult.AddTestResult Test_CargarDatosEspecificos_CDCASUB()
    
    Set IntegrationTest_SolicitudRepository_RunAll = suiteResult
End Function

' ============================================================================
' PROCEDIMIENTOS DE CONFIGURACIÓN Y LIMPIEZA
' ============================================================================

' Procedimiento: Setup
' Propósito: Prepara el entorno de pruebas con una base de datos limpia
Private Sub Setup()
    On Error GoTo ErrorHandler

    ' Aprovisionar la base de datos de prueba antes de cada ejecución
    Dim fullTemplatePath As String
    Dim fullTestPath As String

    fullTemplatePath = modTestUtils.GetProjectPath() & CONDOR_TEMPLATE_PATH
    fullTestPath = modTestUtils.GetProjectPath() & CONDOR_ACTIVE_PATH

    modTestUtils.PrepareTestDatabase fullTemplatePath, fullTestPath
    
    ' Abrir la nueva base de datos y ejecutar sentencias INSERT para datos de prueba
    Dim db As DAO.Database
    Set db = DBEngine.OpenDatabase(fullTestPath)
    
    ' Insertar registro de prueba conocido en T_Solicitudes
    db.Execute "INSERT INTO T_Solicitudes (idSolicitud, idExpediente, fechaCreacion, estado) " & _
               "VALUES (1, 'EXP-TEST-001', #" & Format(Date, "mm/dd/yyyy") & "#, 'Pendiente')"
    
    ' Insertar solicitudes adicionales para pruebas de datos específicos
    db.Execute "INSERT INTO T_Solicitudes (idSolicitud, idExpediente, fechaCreacion, estado) " & _
               "VALUES (2, 'EXP-PC-001', #" & Format(Date, "mm/dd/yyyy") & "#, 'Pendiente')"
    
    db.Execute "INSERT INTO T_Solicitudes (idSolicitud, idExpediente, fechaCreacion, estado) " & _
               "VALUES (3, 'EXP-CDCA-001', #" & Format(Date, "mm/dd/yyyy") & "#, 'Pendiente')"
    
    db.Execute "INSERT INTO T_Solicitudes (idSolicitud, idExpediente, fechaCreacion, estado) " & _
               "VALUES (4, 'EXP-CDCASUB-001', #" & Format(Date, "mm/dd/yyyy") & "#, 'Pendiente')"
    
    ' Insertar datos específicos para PC
    db.Execute "INSERT INTO TbDatos_PC (idSolicitud, campo1, campo2) " & _
               "VALUES (2, 'Valor PC 1', 'Valor PC 2')"
    
    ' Insertar datos específicos para CD_CA
    db.Execute "INSERT INTO TbDatos_CD_CA (idSolicitud, campoA, campoB) " & _
               "VALUES (3, 'Valor CDCA A', 'Valor CDCA B')"
    
    ' Insertar datos específicos para CD_CA_SUB
    db.Execute "INSERT INTO TbDatos_CD_CA_SUB (idSolicitud, campoX, campoY) " & _
               "VALUES (4, 'Valor CDCASUB X', 'Valor CDCASUB Y')"
    
    db.Close
    Set db = Nothing
    
    Exit Sub

ErrorHandler:
    If Not db Is Nothing Then
        db.Close
        Set db = Nothing
    End If
    Debug.Print "Error en Setup (" & Err.Number & "): " & Err.Description
    Err.Raise Err.Number, "IntegrationTest_SolicitudRepository.Setup", Err.Description
End Sub

' Procedimiento: Teardown
' Propósito: Limpia el entorno de pruebas después de la ejecución
Private Sub Teardown()
    On Error GoTo ErrorHandler
    
    ' Eliminar la base de datos de prueba si existe usando IFileSystem
    Dim fs As IFileSystem
    Set fs = modFileSystemFactory.CreateFileSystem()
    
    Dim fullTestPath As String
    fullTestPath = modTestUtils.GetProjectPath() & CONDOR_ACTIVE_PATH
    
    If fs.FileExists(fullTestPath) Then
        fs.DeleteFile fullTestPath
    End If
    
    Set fs = Nothing
    Exit Sub
    
ErrorHandler:
    ' Ignorar errores de limpieza
    Resume Next
End Sub

' ============================================================================
' PRUEBAS DE GetSolicitudById
' ============================================================================

Private Function Test_GetSolicitudById_Success() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "GetSolicitudById debe obtener una solicitud correctamente"
    
    On Error GoTo ErrorHandler
    
    ' Setup - Preparar base de datos de prueba
    Setup
    
    ' Arrange - Configurar repositorio con base de datos de prueba
    Dim testConfig As New CConfig
    testConfig.SetSetting "DATABASE_PATH", modTestUtils.GetProjectPath() & CONDOR_ACTIVE_PATH
    testConfig.SetSetting "DB_PASSWORD", ""
    
    ' Crear instancia del repositorio con dependencias
    Dim repository As New CSolicitudRepository
    repository.Initialize testConfig
    
    ' Act - Ejecutar el método a probar
    Dim result As T_Solicitud
    Set result = repository.GetSolicitudById(1)
    
    ' Assert - Verificar resultados usando modAssert
    modAssert.AssertNotNull result, "El resultado no debería ser nulo."
    modAssert.AssertEquals 1, result.idSolicitud, "El ID de la solicitud no coincide."
    modAssert.AssertEquals "EXP-TEST-001", result.idExpediente, "El ID del expediente no coincide."
    
    testResult.Pass
    
Cleanup:
    Teardown
    Set result = Nothing
    Set repository = Nothing
    Set testConfig = Nothing
    Exit Function
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Description
    Resume Cleanup
End Function

Private Function Test_GetSolicitudById_NotFound() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "GetSolicitudById debe devolver Nothing si la solicitud no existe"
    
    On Error GoTo ErrorHandler
    
    ' Setup - Preparar base de datos de prueba
    Setup
    
    ' Arrange - Configurar repositorio con base de datos de prueba
    Dim testConfig As New CConfig
    testConfig.SetSetting "DATABASE_PATH", modTestUtils.GetProjectPath() & CONDOR_ACTIVE_PATH
    testConfig.SetSetting "DB_PASSWORD", ""
    
    ' Crear instancia del repositorio con dependencias
    Dim repository As New CSolicitudRepository
    repository.Initialize testConfig
    
    ' Act - Ejecutar el método a probar con ID inexistente
    Dim result As T_Solicitud
    Set result = repository.GetSolicitudById(999) ' ID que no existe
    
    ' Assert - Verificar que devuelve Nothing
    modAssert.AssertIsNull result, "El resultado debería ser nulo para un ID inexistente."
    
    testResult.Pass
    
Cleanup:
    Teardown
    Set result = Nothing
    Set repository = Nothing
    Set testConfig = Nothing
    Exit Function
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Description
    Resume Cleanup
End Function

' ============================================================================
' PRUEBAS DE SaveSolicitud
' ============================================================================

Private Function Test_SaveSolicitud_New() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "SaveSolicitud debe insertar una nueva solicitud correctamente"
    
    On Error GoTo ErrorHandler
    
    ' Setup - Preparar base de datos de prueba
    Setup
    
    ' Arrange - Configurar repositorio con base de datos de prueba
     Dim testConfig As New CConfig
     testConfig.SetSetting "DATABASE_PATH", modTestUtils.GetProjectPath() & CONDOR_ACTIVE_PATH
     testConfig.SetSetting "DB_PASSWORD", ""
    
    ' Crear instancia del repositorio con dependencias
     Dim repository As New CSolicitudRepository
     repository.Initialize testConfig
    
    ' Crear nueva solicitud para insertar
    Dim nuevaSolicitud As New T_Solicitud
    nuevaSolicitud.idSolicitud = 0 ' ID 0 indica nueva solicitud
    nuevaSolicitud.idExpediente = "EXP-NEW-001"
    nuevaSolicitud.fechaCreacion = Date
    nuevaSolicitud.estado = "Pendiente"
    
    ' Act - Ejecutar el método a probar
    Dim newId As Long
    newId = repository.SaveSolicitud(nuevaSolicitud)
    
    ' Assert - Verificar que se asignó un ID válido
    modAssert.AssertTrue newId > 1, "El nuevo ID debe ser mayor que 1."
    
    ' Verificar que el registro existe en la base de datos
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Set db = DBEngine.OpenDatabase(modTestUtils.GetProjectPath() & CONDOR_ACTIVE_PATH)
    Set rs = db.OpenRecordset("SELECT * FROM T_Solicitudes WHERE idSolicitud = " & newId)
    
    modAssert.AssertFalse rs.EOF, "El nuevo registro debe existir en la base de datos."
    modAssert.AssertEquals "EXP-NEW-001", rs("idExpediente").Value, "El ID del expediente debe coincidir."
    modAssert.AssertEquals "Pendiente", rs("estado").Value, "El estado debe coincidir."
    
    rs.Close
    db.Close
    
    testResult.Pass
    
Cleanup:
    If Not rs Is Nothing Then
        If Not rs.EOF And Not rs.BOF Then rs.Close
        Set rs = Nothing
    End If
    If Not db Is Nothing Then
        db.Close
        Set db = Nothing
    End If
    Teardown
    Set nuevaSolicitud = Nothing
    Set repository = Nothing
    Set testConfig = Nothing
    Exit Function
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Description
    Resume Cleanup
End Function

Private Function Test_SaveSolicitud_Update() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "SaveSolicitud debe actualizar una solicitud existente correctamente"
    
    On Error GoTo ErrorHandler
    
    ' Setup - Preparar base de datos de prueba
    Setup
    
    ' Arrange - Configurar repositorio con base de datos de prueba
    Dim testConfig As New CConfig
    testConfig.SetSetting "DATABASE_PATH", modTestUtils.GetProjectPath() & CONDOR_ACTIVE_PATH
    testConfig.SetSetting "DB_PASSWORD", ""
    
    ' Crear instancia del repositorio con dependencias
    Dim repository As New CSolicitudRepository
    repository.Initialize testConfig
    
    ' Obtener el objeto T_Solicitud con id = 1 usando GetSolicitudById(1)
    Dim solicitud As T_Solicitud
    Set solicitud = repository.GetSolicitudById(1)
    
    ' Modificar una de sus propiedades
    solicitud.idExpediente = "EXP-TEST-UPDATED"
    
    ' Act - Ejecutar el método a probar
    Dim updatedId As Long
    updatedId = repository.SaveSolicitud(solicitud)
    
    ' Assert - Verificar que devuelve el mismo ID
    modAssert.AssertEquals 1, updatedId, "El ID devuelto debe ser el mismo para actualización."
    
    ' Abrir una conexión DAO directa a la base de datos de prueba
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Set db = DBEngine.OpenDatabase(modTestUtils.GetProjectPath() & CONDOR_ACTIVE_PATH)
    Set rs = db.OpenRecordset("SELECT * FROM T_Solicitudes WHERE idSolicitud = 1")
    
    ' Verificar que el campo idExpediente en la base de datos es ahora "EXP-TEST-UPDATED"
    modAssert.AssertFalse rs.EOF, "El registro debe existir en la base de datos."
    modAssert.AssertEquals "EXP-TEST-UPDATED", rs("idExpediente").Value, "El ID del expediente debe estar actualizado."
    
    rs.Close
    db.Close
    
    testResult.Pass
    
Cleanup:
    If Not rs Is Nothing Then
        If Not rs.EOF And Not rs.BOF Then rs.Close
        Set rs = Nothing
    End If
    If Not db Is Nothing Then
        db.Close
        Set db = Nothing
    End If
    Teardown
    Set solicitud = Nothing
    Set repository = Nothing
    Set testConfig = Nothing
    Exit Function
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Description
    Resume Cleanup
End Function

' ============================================================================
' PRUEBAS DE ExecuteQuery
' ============================================================================

Private Function Test_ExecuteQuery() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "ExecuteQuery debe ejecutar una consulta parametrizada y devolver un recordset"
    
    On Error GoTo ErrorHandler
    
    ' Setup - Preparar base de datos de prueba
    Setup
    
    ' Arrange - Configurar repositorio con base de datos de prueba
    Dim testConfig As New CConfig
    testConfig.SetSetting "DATABASE_PATH", modTestUtils.GetProjectPath() & CONDOR_ACTIVE_PATH
    testConfig.SetSetting "DB_PASSWORD", ""
    
    ' Crear instancia del repositorio con dependencias
    Dim repository As New CSolicitudRepository
    repository.Initialize testConfig
    
    ' Define una consulta SQL parametrizada
    Dim sql As String
    sql = "SELECT idSolicitud FROM T_Solicitudes WHERE idExpediente = ?"
    
    ' Crea una colección de QueryParameter y añade el parámetro
    Dim params As New Collection
    Dim param1 As New QueryParameter
    param1.Initialize "idExpediente", "EXP-TEST-001"
    params.Add param1
    
    ' Act - Llama al método ExecuteQuery
    Dim rs As Object
    Set rs = repository.ExecuteQuery(sql, params)
    
    ' Assert - Utiliza modAssert para validar el recordset devuelto
    modAssert.AssertNotNull rs, "El recordset no debe ser nulo."
    modAssert.AssertFalse rs.EOF, "El recordset no debe estar vacío."
    modAssert.AssertEquals 1, rs!idSolicitud.Value, "El valor del campo no es el esperado."
    
    testResult.Pass
    
Cleanup:
    ' Asegúrate de que el recordset se cierra correctamente
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    Teardown
    Set repository = Nothing
    Set testConfig = Nothing
    Set params = Nothing
    Set param1 = Nothing
    Exit Function
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Description
    Resume Cleanup
End Function

' ============================================================================
' PRUEBAS DE CargarDatosEspecificos
' ============================================================================

Private Function Test_CargarDatosEspecificos_PC() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "CargarDatosEspecificos debe mapear correctamente datos PC"
    
    On Error GoTo ErrorHandler
    
    ' Setup - Preparar base de datos de prueba
    Setup
    
    ' Arrange - Configurar repositorio con base de datos de prueba
    Dim testConfig As New CConfig
    testConfig.SetSetting "DATABASE_PATH", modTestUtils.GetProjectPath() & CONDOR_ACTIVE_PATH
    testConfig.SetSetting "DB_PASSWORD", ""
    
    ' Crear instancia del repositorio con dependencias
    Dim repository As New CSolicitudRepository
    repository.Initialize testConfig
    
    ' Act - Llamar a GetSolicitudById con el ID correspondiente al tipo PC (ID 2)
    Dim solicitud As T_Solicitud
    Set solicitud = repository.GetSolicitudById(2)
    
    ' Assert - Verificar que la propiedad de datos específicos no sea nula
    modAssert.AssertNotNull solicitud, "La solicitud no debería ser nula."
    modAssert.AssertNotNull solicitud.datosPC, "Los datos PC no deberían ser nulos."
    modAssert.AssertEquals "Valor PC 1", solicitud.datosPC.campo1, "El campo1 de datos PC debe coincidir."
    modAssert.AssertEquals "Valor PC 2", solicitud.datosPC.campo2, "El campo2 de datos PC debe coincidir."
    
    testResult.Pass
    
Cleanup:
    Teardown
    Set solicitud = Nothing
    Set repository = Nothing
    Set testConfig = Nothing
    Exit Function
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Description
    Resume Cleanup
End Function

Private Function Test_CargarDatosEspecificos_CDCA() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "CargarDatosEspecificos debe mapear correctamente datos CD_CA"
    
    On Error GoTo ErrorHandler
    
    ' Setup - Preparar base de datos de prueba
    Setup
    
    ' Arrange - Configurar repositorio con base de datos de prueba
    Dim testConfig As New CConfig
    testConfig.SetSetting "DATABASE_PATH", modTestUtils.GetProjectPath() & CONDOR_ACTIVE_PATH
    testConfig.SetSetting "DB_PASSWORD", ""
    
    ' Crear instancia del repositorio con dependencias
    Dim repository As New CSolicitudRepository
    repository.Initialize testConfig
    
    ' Act - Llamar a GetSolicitudById con el ID correspondiente al tipo CDCA (ID 3)
    Dim solicitud As T_Solicitud
    Set solicitud = repository.GetSolicitudById(3)
    
    ' Assert - Verificar que la propiedad de datos específicos no sea nula
    modAssert.AssertNotNull solicitud, "La solicitud no debería ser nula."
    modAssert.AssertNotNull solicitud.datosCDCA, "Los datos CD_CA no deberían ser nulos."
    modAssert.AssertEquals "Valor CDCA A", solicitud.datosCDCA.campoA, "El campoA de datos CD_CA debe coincidir."
    modAssert.AssertEquals "Valor CDCA B", solicitud.datosCDCA.campoB, "El campoB de datos CD_CA debe coincidir."
    
    testResult.Pass
    
Cleanup:
    Teardown
    Set solicitud = Nothing
    Set repository = Nothing
    Set testConfig = Nothing
    Exit Function
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Description
    Resume Cleanup
End Function

Private Function Test_CargarDatosEspecificos_CDCASUB() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "CargarDatosEspecificos debe mapear correctamente datos CD_CA_SUB"
    
    On Error GoTo ErrorHandler
    
    ' Setup - Preparar base de datos de prueba
    Setup
    
    ' Arrange - Configurar repositorio con base de datos de prueba
    Dim testConfig As New CConfig
    testConfig.SetSetting "DATABASE_PATH", modTestUtils.GetProjectPath() & CONDOR_ACTIVE_PATH
    testConfig.SetSetting "DB_PASSWORD", ""
    
    ' Crear instancia del repositorio con dependencias
    Dim repository As New CSolicitudRepository
    repository.Initialize testConfig
    
    ' Act - Llamar a GetSolicitudById con el ID correspondiente al tipo CDCASUB (ID 4)
    Dim solicitud As T_Solicitud
    Set solicitud = repository.GetSolicitudById(4)
    
    ' Assert - Verificar que la propiedad de datos específicos no sea nula
    modAssert.AssertNotNull solicitud, "La solicitud no debería ser nula."
    modAssert.AssertNotNull solicitud.datosCDCASUB, "Los datos CD_CA_SUB no deberían ser nulos."
    modAssert.AssertEquals "Valor CDCASUB X", solicitud.datosCDCASUB.campoX, "El campoX de datos CD_CA_SUB debe coincidir."
    modAssert.AssertEquals "Valor CDCASUB Y", solicitud.datosCDCASUB.campoY, "El campoY de datos CD_CA_SUB debe coincidir."
    
    testResult.Pass
    
Cleanup:
    Teardown
    Set solicitud = Nothing
    Set repository = Nothing
    Set testConfig = Nothing
    Exit Function
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Description
    Resume Cleanup
End Function

#End If



