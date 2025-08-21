# Mejoras en el Manejo de Errores (CONDOR Project)

Este documento lista las funciones y subrutinas que, según el análisis, deberían implementar o mejorar su control de errores centralizado a través de `modErrorHandler.bas`.

## Criterio de Inclusión:
Se incluyen funciones/subs que realizan operaciones propensas a fallar (ej. acceso a base de datos, operaciones de archivo, creación de objetos, llamadas a servicios externos, lógica compleja) y que actualmente no registran sus errores en `modErrorHandler` o no lo hacen de forma consistente.

## Listado de Mejoras Sugeridas:

### `CAuthService.cls`
- `Public Sub Initialize(config As IConfig)`: Realiza asignación de objetos que podría fallar si la configuración es inválida o si `m_config` se utiliza de forma que genere errores.

### `CConfig.cls`
- `Private Sub CreateDirectoriesIfNeeded()`: Realiza operaciones de sistema de archivos (`MkDir`, `CreateObject`, `fso.CreateFolder`) que pueden fallar. Actualmente usa `On Error Resume Next` pero no registra errores en `modErrorHandler`.

### `CNotificationService.cls`
- `Private Sub Class_Initialize()`: Realiza la creación de objetos (`New CConfig`) que puede fallar.
- `Private Function INotificationService_EnviarNotificacion(...)`: Tiene `On Error GoTo ErrorHandler` pero **NO llama a `modErrorHandler.LogError`** en su bloque `ErrorHandler`. Realiza operaciones de base de datos.

### `CSolicitudPC.cls`
- `Private Sub Class_Initialize()`: Llama a `modRepositoryFactory.CreateSolicitudRepository` que puede fallar.
- `Private Function ISolicitud_Load(ByVal ID As Long) As Boolean`: Llama a `mRepository.Load(ID)` que realiza operaciones de base de datos y puede fallar.
- `Private Function ISolicitud_Save() As Boolean`: Llama a `mRepository.Save(Me)` que realiza operaciones de base de datos y puede fallar.
- `Private Function ISolicitud_ChangeState(ByVal newState As String) As Boolean`: Actualmente es una implementación `TODO`. Necesitará manejo de errores una vez implementada.

### `CSolicitudRepository.cls`
- `Private Function ISolicitudRepository_GuardarSolicitud(...)`: Tiene `On Error GoTo ErrorHandler` pero **NO llama a `modErrorHandler.LogError`** en su bloque `ErrorHandler`. Realiza operaciones de base de datos.
- `Private Function ISolicitudRepository_ObtenerSolicitudPorId(...)`: Realiza operaciones de base de datos (`CurrentDb`, `OpenRecordset`).
- `Private Function ISolicitudRepository_ObtenerSolicitudPorCodigo(...)`: Realiza operaciones de base de datos y llama a `ISolicitudRepository_ObtenerSolicitudPorId`.
- `Private Function ISolicitudRepository_EliminarSolicitud(...)`: Tiene `On Error GoTo ErrorHandler` pero **NO llama a `modErrorHandler.LogError`** en su bloque `ErrorHandler`. Realiza operaciones de base de datos.
- `Private Function ISolicitudRepository_ObtenerSolicitudesPorExpediente(...)`: Realiza operaciones de base de datos y llama a `ISolicitudRepository_ObtenerSolicitudPorId`.
- `Private Function ISolicitudRepository_ExisteSolicitudConCodigo(...)`: Realiza operaciones de base de datos.
- `Private Sub GuardarDatosPC(...)`: Realiza operaciones de base de datos.
- `Private Sub GuardarDatosCDCA(...)`: Realiza operaciones de base de datos.
- `Private Sub GuardarDatosCDCASUB(...)`: Realiza operaciones de base de datos.
- `Private Function CargarDatosPC(...)`: Realiza operaciones de base de datos.
- `Private Function CargarDatosCDCA(...)`: Realiza operaciones de base de datos.
- `Private Function CargarDatosCDCASUB(...)`: Realiza operaciones de base de datos.
- `Private Function ISolicitudRepository_Load(...)`: Llama a `ISolicitudRepository_ObtenerSolicitudPorId`.
- `Private Function ISolicitudRepository_Save(...)`: Llama a `ISolicitudRepository_GuardarSolicitud`.

### `CSolicitudService.cls`
- **Todas las implementaciones `TODO`**: (`CreateNuevaSolicitud`, `GetSolicitudPorID`, `SaveSolicitud`, `GetAllSolicitudes`, `DeleteSolicitud`, `UpdateEstadoSolicitud`, `GetSolicitud`, `CreateSolicitud`, `UpdateSolicitud`, `ChangeEstado`, `GetSolicitudesByExpediente`, `GetSolicitudesByTipo`, `GetSolicitudesByEstado`, `SearchSolicitudes`, `ValidateSolicitud`). Todas estas funciones necesitarán `On Error GoTo ErrorHandler` y llamadas a `modErrorHandler.LogError` una vez que su lógica real sea implementada.

### `CWorkflowService.cls`
- `Private Function IWorkflowService_ValidateTransition(...)`: Realiza lógica compleja y llama a otras funciones.
- `Private Function IWorkflowService_GetAvailableStates(...)`: Realiza operaciones de base de datos.
- `Private Function IWorkflowService_GetNextStates(...)`: Realiza operaciones de base de datos.
- `Private Function IWorkflowService_GetInitialState(...)`: Realiza operaciones de base de datos.
- `Private Function IWorkflowService_IsStateFinal(...)`: Realiza operaciones de base de datos.
- `Private Function IWorkflowService_RecordStateChange(...)`: Tiene `On Error GoTo ErrorHandler` pero **NO llama a `modErrorHandler.LogError`** en su bloque `ErrorHandler`. Realiza operaciones de base de datos.
- `Private Function IWorkflowService_GetStateHistory(...)`: Realiza operaciones de base de datos.
- `Private Function IWorkflowService_HasTransitionPermission(...)`: Realiza operaciones de base de datos.
- `Private Function IWorkflowService_RequiresApproval(...)`: Realiza operaciones de base de datos.
- `Private Function TransitionExists(...)`: Realiza operaciones de base de datos.

### `modAppManager.bas`
- `Public Function GetCurrentUserEmail() As String`: Interactúa con `VBA.Command`.

### `modAuthFactory.bas`
- `Public Function CreateAuthService() As IAuthService`: Llama a `New CAuthService`.

### `modRebuildMacro.bas`
- `Public Sub RebuildProject()`: Tiene `On Error GoTo ErrorHandler` pero **NO llama a `modErrorHandler.LogError`** en su bloque `ErrorHandler`. Realiza operaciones críticas del VBE y del sistema de archivos.

### `modRepositoryFactory.bas`
- `Public Function CreateSolicitudRepository() As ISolicitudRepository`: Llama a `New CMockSolicitudRepository`.

### `modSolicitudFactory.bas`
- `Public Function CreateSolicitud(ByVal idSolicitud As Long) As ISolicitud`: Tiene `On Error GoTo ErrorHandler` pero **NO llama a `modErrorHandler.LogError`** en su bloque `ErrorHandler`. Llama a `New CSolicitudPC` y `solicitud.Load`.
- `Private Function GetTipoSolicitud(ByVal idSolicitud As Long) As String`: Actualmente es una implementación `TODO`. Necesitará manejo de errores una vez implementada.
- `Private Function CreateSolicitudPC(ByVal idSolicitud As Long) As ISolicitud`: Llama a `New CSolicitudPC` y `solicitud.Load`.

### `modTest.bas`
- `Public Sub TestInterface()`: Tiene `On Error GoTo ErrorHandler` pero **NO llama a `modErrorHandler.LogError`** en su bloque `ErrorHandler`.

### `modTestRunner.bas`
- `Public Function RunAllTests() As String`: Tiene `On Error GoTo ErrorHandler` pero **NO llama a `modErrorHandler.LogError`** en su bloque `ErrorHandler`.
- `Public Sub ExecuteAllTests(strLogPath As String)`: Tiene `On Error GoTo TestRunnerErrorHandler` pero **NO llama a `modErrorHandler.LogError`** en su bloque `ErrorHandler`. Realiza operaciones de E/S de archivos.
- `Private Sub AnalyzeSuiteResult(...)`: Realiza manipulaciones y conversiones de cadenas.
- `Private Function ExtractTestName(line As String) As String`: Realiza manipulaciones de cadenas.
