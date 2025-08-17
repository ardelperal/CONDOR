Attribute VB_Name = "modMockFramework"
Option Compare Database
Option Explicit

' ============================================================================
' Módulo: modMockFramework
' Descripción: Framework centralizado de mocks para pruebas unitarias CONDOR
' Autor: CONDOR-Expert
' Fecha: Diciembre 2024
' ============================================================================

' ============================================================================
' ESTRUCTURAS MOCK CENTRALIZADAS
' ============================================================================

' Mock para base de datos Lanzadera (usuarios y autenticación)
Type T_MockLanzaderaDB
    IsConnected As Boolean
    ShouldFail As Boolean
    UserExists As Boolean
    UserRole As String
    UserEmail As String
    ErrorNumber As Long
    ErrorDescription As String
    QueryExecuted As String
    RecordCount As Long
End Type

' Mock para base de datos Expedientes
Type T_MockExpedientesDB
    IsConnected As Boolean
    ShouldFail As Boolean
    ExpedienteExists As Boolean
    ExpedienteData As T_Expediente
    ErrorNumber As Long
    ErrorDescription As String
    QueryExecuted As String
    RecordCount As Long
End Type

' Mock para base de datos de Solicitudes
Type T_MockSolicitudesDB
    IsConnected As Boolean
    ShouldFail As Boolean
    SolicitudExists As Boolean
    SolicitudData As T_Solicitud
    PCData As T_Datos_PC
    LastInsertedID As Long
    TransactionActive As Boolean
    ErrorNumber As Long
    ErrorDescription As String
    QueryExecuted As String
    RecordCount As Long
    RecordsAffected As Long
End Type

' Mock para sistema de archivos
Type T_MockFileSystem
    CanReadFile As Boolean
    CanWriteFile As Boolean
    FileExists As Boolean
    DirectoryExists As Boolean
    LastReadContent As String
    LastWrittenContent As String
    FilePath As String
    ErrorOnAccess As Boolean
    AccessAttempts As Long
End Type

' Mock para configuración del sistema
Type T_MockConfiguration
    ConfigLoaded As Boolean
    DatabasePath As String
    LogLevel As String
    MaxRetries As Integer
    TimeoutSeconds As Integer
    EnableNotifications As Boolean
    AdminEmail As String
    ShouldFailLoad As Boolean
    ErrorMessage As String
End Type

' Mock para sistema de notificaciones
Type T_MockNotificationSystem
    IsEnabled As Boolean
    NotificationsSent As Long
    LastRecipient As String
    LastSubject As String
    LastMessage As String
    ShouldFailSend As Boolean
    ErrorMessage As String
    QueueSize As Long
End Type

' Mock para Recordset de DAO
Type T_MockRecordset
    IsOpen As Boolean
    IsEOF As Boolean
    IsBOF As Boolean
    RecordCount As Long
    CurrentRecord As Long
    FieldCount As Integer
    FieldNames As Variant
    FieldValues As Variant
    CanEdit As Boolean
    ShouldFailOperation As Boolean
End Type

' Mock para transacciones de base de datos
Type T_MockTransaction
    IsActive As Boolean
    CanCommit As Boolean
    CanRollback As Boolean
    OperationsCount As Long
    ShouldFailCommit As Boolean
    ShouldFailRollback As Boolean
    ErrorOnOperation As Boolean
End Type

' ============================================================================
' VARIABLES GLOBALES DE MOCKS
' ============================================================================

Private g_MockLanzadera As T_MockLanzaderaDB
Private g_MockExpedientes As T_MockExpedientesDB
Private g_MockSolicitudes As T_MockSolicitudesDB
Private g_MockFileSystem As T_MockFileSystem
Private g_MockConfig As T_MockConfiguration
Private g_MockNotifications As T_MockNotificationSystem
Private g_MockRecordset As T_MockRecordset
Private g_MockTransaction As T_MockTransaction

' ============================================================================
' FUNCIONES DE INICIALIZACIÓN DE MOCKS
' ============================================================================

Public Sub InitializeAllMocks()
    ' Inicializar todos los mocks con valores por defecto
    Call InitializeLanzaderaMock
    Call InitializeExpedientesMock
    Call InitializeSolicitudesMock
    Call InitializeFileSystemMock
    Call InitializeConfigurationMock
    Call InitializeNotificationMock
    Call InitializeRecordsetMock
    Call InitializeTransactionMock
End Sub

Public Sub InitializeLanzaderaMock()
    ' Inicializar mock de base de datos Lanzadera
    With g_MockLanzadera
        .IsConnected = True
        .ShouldFail = False
        .UserExists = True
        .UserRole = "Usuario"
        .UserEmail = "usuario.prueba@empresa.com"
        .ErrorNumber = 0
        .ErrorDescription = ""
        .QueryExecuted = ""
        .RecordCount = 1
    End With
End Sub

Public Sub InitializeExpedientesMock()
    ' Inicializar mock de base de datos Expedientes
    With g_MockExpedientes
        .IsConnected = True
        .ShouldFail = False
        .ExpedienteExists = True
        .ErrorNumber = 0
        .ErrorDescription = ""
        .QueryExecuted = ""
        .RecordCount = 1
        
        ' Inicializar datos de expediente por defecto
        With .ExpedienteData
            .ID = 123
            .IDExpediente = 123
            .Nemotecnico = "EXP-2024-001"
            .Titulo = "Expediente de prueba"
            .ResponsableCalidad = "usuario.prueba@empresa.com"
            .ResponsableTecnico = "jefe.proyecto@empresa.com"
            .Pecal = "PECAL-001"
        End With
    End With
End Sub

Public Sub InitializeSolicitudesMock()
    ' Inicializar mock de base de datos Solicitudes
    With g_MockSolicitudes
        .IsConnected = True
        .ShouldFail = False
        .SolicitudExists = True
        .LastInsertedID = 456
        .TransactionActive = False
        .ErrorNumber = 0
        .ErrorDescription = ""
        .QueryExecuted = ""
        .RecordCount = 1
        .RecordsAffected = 1
        
        ' Inicializar datos de solicitud por defecto
        With .SolicitudData
            .ID = 456
            .NumeroExpediente = "EXP-2024-001"
            .TipoSolicitud = "PC"
            .EstadoInterno = "Borrador"
            .EstadoRAC = "Pendiente"
            .Usuario = "usuario.prueba@empresa.com"
            .FechaCreacion = Date
            .Observaciones = "Solicitud de prueba"
            .Activo = True
        End With
        
        ' Inicializar datos PC por defecto
        With .PCData
            .ID = 789
            .SolicitudID = 456
            .NumeroExpediente = "EXP-2024-001"
            .TipoSolicitud = "PC"
            .DescripcionCambio = "Descripción de prueba"
            .JustificacionCambio = "Justificación de prueba"
            .ImpactoSeguridad = "Bajo"
            .ImpactoCalidad = "Medio"
            .Estado = "Activo"
            .FechaCreacion = Date
            .Activo = True
        End With
    End With
End Sub

Public Sub InitializeFileSystemMock()
    ' Inicializar mock del sistema de archivos
    With g_MockFileSystem
        .CanReadFile = True
        .CanWriteFile = True
        .FileExists = True
        .DirectoryExists = True
        .LastReadContent = ""
        .LastWrittenContent = ""
        .FilePath = ""
        .ErrorOnAccess = False
        .AccessAttempts = 0
    End With
End Sub

Public Sub InitializeConfigurationMock()
    ' Inicializar mock de configuración
    With g_MockConfig
        .ConfigLoaded = True
        .DatabasePath = "C:\Proyectos\CONDOR\CONDOR_datos.accdb"
        .LogLevel = "INFO"
        .MaxRetries = 3
        .TimeoutSeconds = 30
        .EnableNotifications = True
        .AdminEmail = "admin@condor.local"
        .ShouldFailLoad = False
        .ErrorMessage = ""
    End With
End Sub

Public Sub InitializeNotificationMock()
    ' Inicializar mock de notificaciones
    With g_MockNotifications
        .IsEnabled = True
        .NotificationsSent = 0
        .LastRecipient = ""
        .LastSubject = ""
        .LastMessage = ""
        .ShouldFailSend = False
        .ErrorMessage = ""
        .QueueSize = 0
    End With
End Sub

Public Sub InitializeRecordsetMock()
    ' Inicializar mock de Recordset
    With g_MockRecordset
        .IsOpen = True
        .IsEOF = False
        .IsBOF = False
        .RecordCount = 1
        .CurrentRecord = 1
        .FieldCount = 5
        .FieldNames = Array("ID", "Nombre", "Email", "Fecha", "Activo")
        .FieldValues = Array(123, "Usuario Prueba", "usuario@test.com", Date, True)
        .CanEdit = True
        .ShouldFailOperation = False
    End With
End Sub

Public Sub InitializeTransactionMock()
    ' Inicializar mock de transacciones
    With g_MockTransaction
        .IsActive = False
        .CanCommit = True
        .CanRollback = True
        .OperationsCount = 0
        .ShouldFailCommit = False
        .ShouldFailRollback = False
        .ErrorOnOperation = False
    End With
End Sub

' ============================================================================
' FUNCIONES DE CONFIGURACIÓN DE FALLOS
' ============================================================================

Public Sub ConfigureLanzaderaToFail(errorNum As Long, errorDesc As String)
    ' Configurar mock de Lanzadera para fallar
    With g_MockLanzadera
        .ShouldFail = True
        .IsConnected = False
        .ErrorNumber = errorNum
        .ErrorDescription = errorDesc
        .UserExists = False
        .RecordCount = 0
    End With
End Sub

Public Sub ConfigureExpedientesToFail(errorNum As Long, errorDesc As String)
    ' Configurar mock de Expedientes para fallar
    With g_MockExpedientes
        .ShouldFail = True
        .IsConnected = False
        .ErrorNumber = errorNum
        .ErrorDescription = errorDesc
        .ExpedienteExists = False
        .RecordCount = 0
    End With
End Sub

Public Sub ConfigureSolicitudesToFail(errorNum As Long, errorDesc As String)
    ' Configurar mock de Solicitudes para fallar
    With g_MockSolicitudes
        .ShouldFail = True
        .IsConnected = False
        .ErrorNumber = errorNum
        .ErrorDescription = errorDesc
        .SolicitudExists = False
        .RecordCount = 0
        .RecordsAffected = 0
    End With
End Sub

Public Sub ConfigureFileSystemToFail()
    ' Configurar mock del sistema de archivos para fallar
    With g_MockFileSystem
        .CanReadFile = False
        .CanWriteFile = False
        .FileExists = False
        .DirectoryExists = False
        .ErrorOnAccess = True
    End With
End Sub

Public Sub ConfigureConfigurationToFail(errorMsg As String)
    ' Configurar mock de configuración para fallar
    With g_MockConfig
        .ConfigLoaded = False
        .ShouldFailLoad = True
        .ErrorMessage = errorMsg
    End With
End Sub

Public Sub ConfigureNotificationToFail(errorMsg As String)
    ' Configurar mock de notificaciones para fallar
    With g_MockNotifications
        .IsEnabled = False
        .ShouldFailSend = True
        .ErrorMessage = errorMsg
    End With
End Sub

Public Sub ConfigureRecordsetToFail()
    ' Configurar mock de Recordset para fallar
    With g_MockRecordset
        .IsOpen = False
        .ShouldFailOperation = True
        .RecordCount = 0
        .CanEdit = False
    End With
End Sub

Public Sub ConfigureTransactionToFail(failCommit As Boolean, failRollback As Boolean)
    ' Configurar mock de transacciones para fallar
    With g_MockTransaction
        .ShouldFailCommit = failCommit
        .ShouldFailRollback = failRollback
        .ErrorOnOperation = True
        .CanCommit = Not failCommit
        .CanRollback = Not failRollback
    End With
End Sub

' ============================================================================
' FUNCIONES DE ACCESO A MOCKS (GETTERS)
' ============================================================================

Public Function GetLanzaderaMock() As T_MockLanzaderaDB
    GetLanzaderaMock = g_MockLanzadera
End Function

Public Function GetExpedientesMock() As T_MockExpedientesDB
    GetExpedientesMock = g_MockExpedientes
End Function

Public Function GetSolicitudesMock() As T_MockSolicitudesDB
    GetSolicitudesMock = g_MockSolicitudes
End Function

Public Function GetFileSystemMock() As T_MockFileSystem
    GetFileSystemMock = g_MockFileSystem
End Function

Public Function GetConfigurationMock() As T_MockConfiguration
    GetConfigurationMock = g_MockConfig
End Function

Public Function GetNotificationMock() As T_MockNotificationSystem
    GetNotificationMock = g_MockNotifications
End Function

Public Function GetRecordsetMock() As T_MockRecordset
    GetRecordsetMock = g_MockRecordset
End Function

Public Function GetTransactionMock() As T_MockTransaction
    GetTransactionMock = g_MockTransaction
End Function

' ============================================================================
' FUNCIONES DE CONFIGURACIÓN ESPECÍFICA
' ============================================================================

Public Sub SetLanzaderaUser(email As String, role As String, exists As Boolean)
    ' Configurar usuario específico en mock de Lanzadera
    With g_MockLanzadera
        .UserEmail = email
        .UserRole = role
        .UserExists = exists
        .RecordCount = IIf(exists, 1, 0)
    End With
End Sub

Public Sub SetExpedienteData(id As Long, nemotecnico As String, titulo As String, responsableCalidad As String)
    ' Configurar datos específicos de expediente
    With g_MockExpedientes.ExpedienteData
        .ID = id
        .IDExpediente = id
        .Nemotecnico = nemotecnico
        .Titulo = titulo
        .ResponsableCalidad = responsableCalidad
        .ResponsableTecnico = "jefe.proyecto@empresa.com"
        .Pecal = "PECAL-001"
    End With
    g_MockExpedientes.ExpedienteExists = True
    g_MockExpedientes.RecordCount = 1
End Sub

Public Sub SetSolicitudData(id As Long, expediente As String, tipo As String, estado As String)
    ' Configurar datos específicos de solicitud
    With g_MockSolicitudes.SolicitudData
        .ID = id
        .NumeroExpediente = expediente
        .TipoSolicitud = tipo
        .EstadoInterno = estado
        .FechaCreacion = Date
        .Activo = True
    End With
    g_MockSolicitudes.SolicitudExists = True
    g_MockSolicitudes.RecordCount = 1
End Sub

Public Sub SetPCData(id As Long, solicitudId As Long, descripcion As String, justificacion As String)
    ' Configurar datos específicos de PC
    With g_MockSolicitudes.PCData
        .ID = id
        .SolicitudID = solicitudId
        .DescripcionCambio = descripcion
        .JustificacionCambio = justificacion
        .FechaCreacion = Date
        .Activo = True
    End With
End Sub

Public Sub SetFileSystemPath(filePath As String, exists As Boolean, canRead As Boolean, canWrite As Boolean)
    ' Configurar ruta específica en sistema de archivos
    With g_MockFileSystem
        .FilePath = filePath
        .FileExists = exists
        .CanReadFile = canRead
        .CanWriteFile = canWrite
        .ErrorOnAccess = False
    End With
End Sub

Public Sub SetRecordsetData(fieldNames As Variant, fieldValues As Variant, recordCount As Long)
    ' Configurar datos específicos de Recordset
    With g_MockRecordset
        .FieldNames = fieldNames
        .FieldValues = fieldValues
        .RecordCount = recordCount
        .FieldCount = UBound(fieldNames) + 1
        .IsEOF = (recordCount = 0)
        .IsOpen = True
    End With
End Sub

' ============================================================================
' FUNCIONES DE SIMULACIÓN DE OPERACIONES
' ============================================================================

Public Sub SimulateQuery(mockType As String, query As String)
    ' Simular ejecución de consulta en el mock especificado
    Select Case UCase(mockType)
        Case "LANZADERA"
            g_MockLanzadera.QueryExecuted = query
        Case "EXPEDIENTES"
            g_MockExpedientes.QueryExecuted = query
        Case "SOLICITUDES"
            g_MockSolicitudes.QueryExecuted = query
    End Select
End Sub

Public Sub SimulateFileAccess(filePath As String, operation As String, content As String)
    ' Simular acceso a archivo
    With g_MockFileSystem
        .FilePath = filePath
        .AccessAttempts = .AccessAttempts + 1
        
        Select Case UCase(operation)
            Case "READ"
                .LastReadContent = content
            Case "WRITE"
                .LastWrittenContent = content
        End Select
    End With
End Sub

Public Sub SimulateNotification(recipient As String, subject As String, message As String)
    ' Simular envío de notificación
    With g_MockNotifications
        If Not .ShouldFailSend Then
            .NotificationsSent = .NotificationsSent + 1
            .LastRecipient = recipient
            .LastSubject = subject
            .LastMessage = message
            .QueueSize = .QueueSize + 1
        End If
    End With
End Sub

Public Sub SimulateTransaction(operation As String)
    ' Simular operación de transacción
    With g_MockTransaction
        Select Case UCase(operation)
            Case "BEGIN"
                .IsActive = True
                .OperationsCount = 0
            Case "COMMIT"
                If Not .ShouldFailCommit Then
                    .IsActive = False
                End If
            Case "ROLLBACK"
                If Not .ShouldFailRollback Then
                    .IsActive = False
                End If
            Case "OPERATION"
                .OperationsCount = .OperationsCount + 1
        End Select
    End With
End Sub

' ============================================================================
' FUNCIONES DE VALIDACIÓN Y VERIFICACIÓN
' ============================================================================

Public Function VerifyQueryExecuted(ByVal mockType As String, ByVal expectedQuery As String) As Boolean
    ' Verificar que se ejecutó la consulta esperada
    Dim actualQuery As String
    
    Select Case UCase(mockType)
        Case "LANZADERA"
            actualQuery = g_MockLanzadera.QueryExecuted
        Case "EXPEDIENTES"
            actualQuery = g_MockExpedientes.QueryExecuted
        Case "SOLICITUDES"
            actualQuery = g_MockSolicitudes.QueryExecuted
        Case Else
            VerifyQueryExecuted = False
            Exit Function
    End Select
    
    VerifyQueryExecuted = (InStr(UCase(actualQuery), UCase(expectedQuery)) > 0)
End Function

Public Function VerifyFileAccessed(ByVal expectedPath As String, ByVal expectedOperation As String) As Boolean
    ' Verificar que se accedió al archivo esperado
    With g_MockFileSystem
        VerifyFileAccessed = (.FilePath = expectedPath) And (.AccessAttempts > 0)
    End With
End Function

Public Function VerifyNotificationSent(ByVal expectedRecipient As String, ByVal expectedSubject As String) As Boolean
    ' Verificar que se envió la notificación esperada
    With g_MockNotifications
        VerifyNotificationSent = (.LastRecipient = expectedRecipient) And _
                                (InStr(.LastSubject, expectedSubject) > 0) And _
                                (.NotificationsSent > 0)
    End With
End Function

Public Function VerifyTransactionState(ByVal expectedState As String) As Boolean
    ' Verificar el estado de la transacción
    With g_MockTransaction
        Select Case UCase(expectedState)
            Case "ACTIVE"
                VerifyTransactionState = .IsActive
            Case "INACTIVE"
                VerifyTransactionState = Not .IsActive
            Case "CAN_COMMIT"
                VerifyTransactionState = .CanCommit
            Case "CAN_ROLLBACK"
                VerifyTransactionState = .CanRollback
            Case Else
                VerifyTransactionState = False
        End Select
    End With
End Function

' ============================================================================
' FUNCIONES DE LIMPIEZA Y RESET
' ============================================================================

Public Sub ResetAllMocks()
    ' Reinicializar todos los mocks a sus valores por defecto
    Call InitializeAllMocks
End Sub

Public Sub ResetMockCounters()
    ' Reinicializar solo los contadores de los mocks
    g_MockNotifications.NotificationsSent = 0
    g_MockFileSystem.AccessAttempts = 0
    g_MockTransaction.OperationsCount = 0
    g_MockSolicitudes.RecordsAffected = 0
End Sub

Public Sub ClearMockData()
    ' Limpiar datos específicos de los mocks
    g_MockLanzadera.QueryExecuted = ""
    g_MockExpedientes.QueryExecuted = ""
    g_MockSolicitudes.QueryExecuted = ""
    g_MockFileSystem.LastReadContent = ""
    g_MockFileSystem.LastWrittenContent = ""
    g_MockNotifications.LastRecipient = ""
    g_MockNotifications.LastSubject = ""
    g_MockNotifications.LastMessage = ""
End Sub

' ============================================================================
' FUNCIONES DE UTILIDAD PARA PRUEBAS
' ============================================================================

Public Function GetMockSummary() As String
    ' Obtener resumen del estado actual de todos los mocks
    Dim summary As String
    
    summary = "=== RESUMEN DE MOCKS ===" & vbCrLf
    summary = summary & "Lanzadera: " & IIf(g_MockLanzadera.IsConnected, "Conectado", "Desconectado")
    summary = summary & " | Fallos: " & g_MockLanzadera.ShouldFail & vbCrLf
    summary = summary & "Expedientes: " & IIf(g_MockExpedientes.IsConnected, "Conectado", "Desconectado")
    summary = summary & " | Fallos: " & g_MockExpedientes.ShouldFail & vbCrLf
    summary = summary & "Solicitudes: " & IIf(g_MockSolicitudes.IsConnected, "Conectado", "Desconectado")
    summary = summary & " | Fallos: " & g_MockSolicitudes.ShouldFail & vbCrLf
    summary = summary & "FileSystem: " & IIf(g_MockFileSystem.CanReadFile, "Disponible", "No disponible")
    summary = summary & " | Accesos: " & g_MockFileSystem.AccessAttempts & vbCrLf
    summary = summary & "Notificaciones: " & IIf(g_MockNotifications.IsEnabled, "Habilitado", "Deshabilitado")
    summary = summary & " | Enviadas: " & g_MockNotifications.NotificationsSent & vbCrLf
    summary = summary & "Transacciones: " & IIf(g_MockTransaction.IsActive, "Activa", "Inactiva")
    summary = summary & " | Operaciones: " & g_MockTransaction.OperationsCount & vbCrLf
    
    GetMockSummary = summary
End Function

Public Sub LogMockActivity(activity As String)
    ' Registrar actividad de mock para debugging
    Debug.Print Format(Now(), "hh:nn:ss") & " - MOCK: " & activity
End Sub