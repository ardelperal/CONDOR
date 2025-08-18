# Guías de Codificación VBA - Sistema CONDOR

## Introducción

Este documento establece las guías de codificación específicas para el desarrollo en VBA del Sistema CONDOR, basadas en la arquitectura de 3 capas definida en la especificación funcional del proyecto.

---

## 1. Arquitectura de 3 Capas

### 1.1 Estructura General

```
┌─────────────────────────────────────────────────────────────┐
│                    CAPA DE PRESENTACIÓN                    │
│  ┌─────────────────┐  ┌─────────────────┐  ┌─────────────┐ │
│  │   Formularios   │  │    Informes     │  │   Módulos   │ │
│  │   (Forms)       │  │   (Reports)     │  │   de UI     │ │
│  └─────────────────┘  └─────────────────┘  └─────────────┘ │
└─────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────┐
│                   CAPA DE LÓGICA DE NEGOCIO                │
│  ┌─────────────────┐  ┌─────────────────┐  ┌─────────────┐ │
│  │    Clases de    │  │   Interfaces    │  │   Módulos   │ │
│  │    Negocio      │  │   (Contracts)   │  │ de Negocio  │ │
│  └─────────────────┘  └─────────────────┘  └─────────────┘ │
└─────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────┐
│                     CAPA DE SERVICIOS                      │
│  ┌─────────────────┐  ┌─────────────────┐  ┌─────────────┐ │
│  │   Servicios de  │  │   Repositorios  │  │   Módulos   │ │
│  │     Datos       │  │   (Data Access) │  │ de Servicio │ │
│  └─────────────────┘  └─────────────────┘  └─────────────┘ │
└─────────────────────────────────────────────────────────────┘
```

### 1.2 Principios de Separación

1. **Capa de Presentación**: Solo maneja la interfaz de usuario y delegación a la capa de negocio
2. **Capa de Lógica de Negocio**: Contiene todas las reglas de negocio y validaciones
3. **Capa de Servicios**: Maneja acceso a datos y servicios externos

---

## 2. Convenciones de Nomenclatura

### 2.1 Archivos y Módulos

#### Formularios (Capa de Presentación)
```vba
' Prefijo: frm
' Formato: frm[Entidad][Accion]
' Ejemplos:
frmSolicitudCrear.frm
frmSolicitudEditar.frm
frmSolicitudBuscar.frm
frmExpedienteConsultar.frm
```

#### Clases de Negocio (Capa de Lógica de Negocio)
```vba
' Prefijo: C (para clases concretas)
' Formato: C[Entidad][Responsabilidad]
' Ejemplos:
CSolicitudManager.cls
CExpedienteValidator.cls
CDocumentoGenerator.cls
CNotificacionService.cls
```

#### Interfaces (Capa de Lógica de Negocio)
```vba
' Prefijo: I
' Formato: I[Responsabilidad]
' Ejemplos:
ISolicitudService.cls
IExpedienteService.cls
IDocumentoService.cls
INotificacionService.cls
```

#### Tipos de Datos
```vba
' Prefijo: T_
' Formato: T_[Entidad]
' Ejemplos:
T_Solicitud.cls
T_Expediente.cls
T_Usuario.cls
T_Configuracion.cls
```

#### Módulos de Servicio (Capa de Servicios)
```vba
' Prefijo: Mod
' Formato: Mod[Responsabilidad]
' Ejemplos:
ModDatabaseService.bas
ModExpedienteService.bas
ModConfiguracionService.bas
ModLogService.bas
```

### 2.2 Variables y Constantes

#### Variables Locales
```vba
' Usar camelCase
' Prefijos por tipo:
Dim numeroSolicitud As String        ' str para String
Dim fechaCreacion As Date           ' dt para Date
Dim importeTotal As Currency        ' cur para Currency
Dim esValido As Boolean            ' b para Boolean
Dim contador As Integer            ' i para Integer
Dim indice As Long                 ' l para Long
```

#### Variables de Módulo
```vba
' Prefijo: m_
Private m_solicitudActual As T_Solicitud
Private m_usuarioLogueado As String
Private m_configuracionCargada As Boolean
```

#### Constantes
```vba
' Usar UPPER_CASE con prefijo según alcance
' Constantes públicas:
Public Const CONDOR_VERSION As String = "1.0.0"
Public Const MAX_INTENTOS_CONEXION As Integer = 3

' Constantes privadas:
Private Const TIMEOUT_DEFAULT As Integer = 30
Private Const RUTA_PLANTILLAS As String = "\\servidor\condor\plantillas\"
```

---

## 3. Estructura de Clases por Capa

### 3.1 Capa de Presentación - Formularios

```vba
' Archivo: frmSolicitudCrear.frm
' Responsabilidad: Interfaz para crear nuevas solicitudes

Option Compare Database
Option Explicit

' === VARIABLES DE MÓDULO ===
Private m_solicitudService As ISolicitudService
Private m_expedienteService As IExpedienteService
Private m_solicitudActual As T_Solicitud

' === EVENTOS DEL FORMULARIO ===
Private Sub Form_Load()
    ' Inicializar servicios usando inyección de dependencias
    Set m_solicitudService = CreateSolicitudService()
    Set m_expedienteService = CreateExpedienteService()
    
    ' Configurar interfaz inicial
    ConfigurarInterfazInicial
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' Limpiar referencias
    Set m_solicitudService = Nothing
    Set m_expedienteService = Nothing
    Set m_solicitudActual = Nothing
End Sub

' === EVENTOS DE CONTROLES ===
Private Sub cmdGuardar_Click()
    If ValidarFormulario() Then
        GuardarSolicitud
    End If
End Sub

Private Sub cmbExpediente_AfterUpdate()
    CargarDatosExpediente Me.cmbExpediente.Value
End Sub

' === MÉTODOS PRIVADOS ===
Private Sub ConfigurarInterfazInicial()
    ' Solo lógica de UI - NO lógica de negocio
    Me.txtFechaCreacion.Value = Date
    Me.cmbTipoSolicitud.RowSource = "PC;CD_CA;CD_CA_SUB"
    Me.txtUsuarioCreador.Value = CurrentUser()
End Sub

Private Function ValidarFormulario() As Boolean
    ' Validaciones básicas de UI
    If IsNull(Me.cmbExpediente.Value) Then
        MsgBox "Debe seleccionar un expediente", vbExclamation
        Me.cmbExpediente.SetFocus
        ValidarFormulario = False
        Exit Function
    End If
    
    ValidarFormulario = True
End Function

Private Sub GuardarSolicitud()
    On Error GoTo ErrorHandler
    
    ' Crear objeto de solicitud desde formulario
    Set m_solicitudActual = CrearSolicitudDesdeFormulario()
    
    ' Delegar a la capa de negocio
    Dim resultado As Boolean
    resultado = m_solicitudService.CrearSolicitud(m_solicitudActual)
    
    If resultado Then
        MsgBox "Solicitud creada exitosamente", vbInformation
        DoCmd.Close acForm, Me.Name
    Else
        MsgBox "Error al crear la solicitud", vbCritical
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error inesperado: " & Err.Description, vbCritical
End Sub

Private Function CrearSolicitudDesdeFormulario() As T_Solicitud
    ' Mapear datos del formulario a objeto de negocio
    Dim solicitud As New T_Solicitud
    
    solicitud.NumeroExpediente = Me.cmbExpediente.Value
    solicitud.TipoSolicitud = Me.cmbTipoSolicitud.Value
    solicitud.DescripcionSolicitud = Me.txtDescripcion.Value
    solicitud.UsuarioCreador = Me.txtUsuarioCreador.Value
    solicitud.FechaCreacion = Me.txtFechaCreacion.Value
    
    Set CrearSolicitudDesdeFormulario = solicitud
End Function
```

### 3.2 Capa de Lógica de Negocio - Interfaces

```vba
' Archivo: ISolicitudService.cls
' Responsabilidad: Contrato para servicios de solicitud

Option Compare Database
Option Explicit

' === MÉTODOS DE NEGOCIO ===
Public Function CrearSolicitud(solicitud As T_Solicitud) As Boolean
    ' Implementado por clase concreta
End Function

Public Function ActualizarSolicitud(solicitud As T_Solicitud) As Boolean
    ' Implementado por clase concreta
End Function

Public Function ObtenerSolicitud(solicitudId As String) As T_Solicitud
    ' Implementado por clase concreta
End Function

Public Function CambiarEstadoSolicitud(solicitudId As String, nuevoEstado As String) As Boolean
    ' Implementado por clase concreta
End Function

Public Function ValidarSolicitud(solicitud As T_Solicitud) As Collection
    ' Retorna colección de errores de validación
    ' Implementado por clase concreta
End Function

Public Function BuscarSolicitudes(criterios As T_CriteriosBusqueda) As Collection
    ' Implementado por clase concreta
End Function
```

### 3.3 Capa de Lógica de Negocio - Implementaciones

```vba
' Archivo: CSolicitudManager.cls
' Responsabilidad: Lógica de negocio para solicitudes

Option Compare Database
Option Explicit
Implements ISolicitudService

' === DEPENDENCIAS ===
Private m_solicitudRepository As ISolicitudRepository
Private m_expedienteService As IExpedienteService
Private m_notificacionService As INotificacionService
Private m_logService As ILogService

' === CONSTRUCTOR ===
Public Sub Initialize(solicitudRepo As ISolicitudRepository, _
                     expedienteServ As IExpedienteService, _
                     notificacionServ As INotificacionService, _
                     logServ As ILogService)
    Set m_solicitudRepository = solicitudRepo
    Set m_expedienteService = expedienteServ
    Set m_notificacionService = notificacionServ
    Set m_logService = logServ
End Sub

' === IMPLEMENTACIÓN DE INTERFAZ ===
Private Function ISolicitudService_CrearSolicitud(solicitud As T_Solicitud) As Boolean
    On Error GoTo ErrorHandler
    
    ' 1. Validar reglas de negocio
    Dim errores As Collection
    Set errores = ValidarReglasNegocio(solicitud)
    
    If errores.Count > 0 Then
        ' Log errores y retornar false
        m_logService.LogError "Validación fallida al crear solicitud", errores
        ISolicitudService_CrearSolicitud = False
        Exit Function
    End If
    
    ' 2. Verificar expediente existe
    If Not m_expedienteService.ExisteExpediente(solicitud.NumeroExpediente) Then
        m_logService.LogError "Expediente no encontrado: " & solicitud.NumeroExpediente
        ISolicitudService_CrearSolicitud = False
        Exit Function
    End If
    
    ' 3. Asignar valores por defecto
    AsignarValoresPorDefecto solicitud
    
    ' 4. Guardar en repositorio
    Dim solicitudId As String
    solicitudId = m_solicitudRepository.Insertar(solicitud)
    
    If Len(solicitudId) > 0 Then
        ' 5. Enviar notificaciones
        EnviarNotificacionCreacion solicitud
        
        ' 6. Log éxito
        m_logService.LogInfo "Solicitud creada exitosamente: " & solicitudId
        
        ISolicitudService_CrearSolicitud = True
    Else
        ISolicitudService_CrearSolicitud = False
    End If
    
    Exit Function
    
ErrorHandler:
    m_logService.LogError "Error al crear solicitud: " & Err.Description
    ISolicitudService_CrearSolicitud = False
End Function

' === MÉTODOS PRIVADOS DE NEGOCIO ===
Private Function ValidarReglasNegocio(solicitud As T_Solicitud) As Collection
    Dim errores As New Collection
    
    ' Regla: Descripción no puede estar vacía
    If Len(Trim(solicitud.DescripcionSolicitud)) = 0 Then
        errores.Add "La descripción de la solicitud es obligatoria"
    End If
    
    ' Regla: Usuario debe tener permisos
    If Not TienePermisosCreacion(solicitud.UsuarioCreador, solicitud.TipoSolicitud) Then
        errores.Add "El usuario no tiene permisos para crear este tipo de solicitud"
    End If
    
    ' Regla: No puede haber solicitudes duplicadas
    If ExisteSolicitudDuplicada(solicitud) Then
        errores.Add "Ya existe una solicitud similar para este expediente"
    End If
    
    Set ValidarReglasNegocio = errores
End Function

Private Sub AsignarValoresPorDefecto(solicitud As T_Solicitud)
    ' Asignar ID único
    solicitud.SolicitudId = GenerarIdSolicitud()
    
    ' Estado inicial
    solicitud.EstadoInterno = "Borrador"
    solicitud.EstadoRAC = "Pendiente"
    
    ' Fechas
    solicitud.FechaCreacion = Now()
    solicitud.FechaUltimaModificacion = Now()
End Sub

Private Sub EnviarNotificacionCreacion(solicitud As T_Solicitud)
    ' Notificar a Ingeniería si requiere revisión técnica
    If RequiereRevisionTecnica(solicitud.TipoSolicitud) Then
        m_notificacionService.NotificarRevisionTecnica solicitud
    End If
    
    ' Notificar al jefe de proyecto
    m_notificacionService.NotificarJefeProyecto solicitud
End Sub
```

### 3.4 Capa de Servicios - Repositorios

```vba
' Archivo: CSolicitudRepository.cls
' Responsabilidad: Acceso a datos de solicitudes

Option Compare Database
Option Explicit
Implements ISolicitudRepository

' === IMPLEMENTACIÓN DE INTERFAZ ===
Private Function ISolicitudRepository_Insertar(solicitud As T_Solicitud) As String
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    
    Set db = CurrentDb()
    Set rs = db.OpenRecordset("Tb_Solicitudes", dbOpenDynaset)
    
    rs.AddNew
    
    ' Mapear objeto a campos de base de datos
    rs!SolicitudId = solicitud.SolicitudId
    rs!NumeroExpediente = solicitud.NumeroExpediente
    rs!TipoSolicitud = solicitud.TipoSolicitud
    rs!DescripcionSolicitud = solicitud.DescripcionSolicitud
    rs!EstadoInterno = solicitud.EstadoInterno
    rs!EstadoRAC = solicitud.EstadoRAC
    rs!UsuarioCreador = solicitud.UsuarioCreador
    rs!FechaCreacion = solicitud.FechaCreacion
    rs!FechaUltimaModificacion = solicitud.FechaUltimaModificacion
    
    rs.Update
    
    ISolicitudRepository_Insertar = solicitud.SolicitudId
    
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    
    Exit Function
    
ErrorHandler:
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    
    ISolicitudRepository_Insertar = ""
End Function

Private Function ISolicitudRepository_ObtenerPorId(solicitudId As String) As T_Solicitud
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
    
    Set db = CurrentDb()
    
    sql = "SELECT * FROM Tb_Solicitudes WHERE SolicitudId = '" & solicitudId & "'"
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    
    If Not rs.EOF Then
        Dim solicitud As New T_Solicitud
        
        ' Mapear campos a objeto
        solicitud.SolicitudId = rs!SolicitudId & ""
        solicitud.NumeroExpediente = rs!NumeroExpediente & ""
        solicitud.TipoSolicitud = rs!TipoSolicitud & ""
        solicitud.DescripcionSolicitud = rs!DescripcionSolicitud & ""
        solicitud.EstadoInterno = rs!EstadoInterno & ""
        solicitud.EstadoRAC = rs!EstadoRAC & ""
        solicitud.UsuarioCreador = rs!UsuarioCreador & ""
        
        If Not IsNull(rs!FechaCreacion) Then
            solicitud.FechaCreacion = rs!FechaCreacion
        End If
        
        If Not IsNull(rs!FechaUltimaModificacion) Then
            solicitud.FechaUltimaModificacion = rs!FechaUltimaModificacion
        End If
        
        Set ISolicitudRepository_ObtenerPorId = solicitud
    Else
        Set ISolicitudRepository_ObtenerPorId = Nothing
    End If
    
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    
    Exit Function
    
ErrorHandler:
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Set ISolicitudRepository_ObtenerPorId = Nothing
End Function
```

---

## 4. Patrones de Diseño

### 4.1 Inyección de Dependencias

```vba
' Archivo: ModDependencyInjection.bas
' Responsabilidad: Configurar e inyectar dependencias

Option Compare Database
Option Explicit

' === FACTORY METHODS ===
Public Function CreateSolicitudService() As ISolicitudService
    ' Crear dependencias
    Dim solicitudRepo As ISolicitudRepository
    Dim expedienteServ As IExpedienteService
    Dim notificacionServ As INotificacionService
    Dim logServ As ILogService
    
    Set solicitudRepo = New CSolicitudRepository
    Set expedienteServ = CreateExpedienteService()
    Set notificacionServ = New CNotificacionService
    Set logServ = New CLogService
    
    ' Crear e inicializar servicio principal
    Dim solicitudManager As New CSolicitudManager
    solicitudManager.Initialize solicitudRepo, expedienteServ, notificacionServ, logServ
    
    Set CreateSolicitudService = solicitudManager
End Function

Public Function CreateExpedienteService() As IExpedienteService
    ' Determinar implementación según configuración
    If GetConfigValue("DEV_MODE") = "true" Then
        Set CreateExpedienteService = New CExpedienteServiceMock
    Else
        Set CreateExpedienteService = New CExpedienteService
    End If
End Function
```

### 4.2 Repository Pattern

```vba
' Archivo: ISolicitudRepository.cls
' Responsabilidad: Contrato para acceso a datos de solicitudes

Option Compare Database
Option Explicit

Public Function Insertar(solicitud As T_Solicitud) As String
End Function

Public Function Actualizar(solicitud As T_Solicitud) As Boolean
End Function

Public Function ObtenerPorId(solicitudId As String) As T_Solicitud
End Function

Public Function Eliminar(solicitudId As String) As Boolean
End Function

Public Function BuscarPorCriterios(criterios As T_CriteriosBusqueda) As Collection
End Function

Public Function ObtenerPorExpediente(numeroExpediente As String) As Collection
End Function
```

### 4.3 Command Pattern para Operaciones

```vba
' Archivo: ICommand.cls
' Responsabilidad: Interfaz para comandos

Option Compare Database
Option Explicit

Public Function Execute() As Boolean
End Function

Public Function Undo() As Boolean
End Function

Public Function GetDescription() As String
End Function
```

```vba
' Archivo: CCrearSolicitudCommand.cls
' Responsabilidad: Comando para crear solicitud

Option Compare Database
Option Explicit
Implements ICommand

Private m_solicitud As T_Solicitud
Private m_solicitudService As ISolicitudService
Private m_solicitudCreada As String

Public Sub Initialize(solicitud As T_Solicitud, solicitudService As ISolicitudService)
    Set m_solicitud = solicitud
    Set m_solicitudService = solicitudService
End Sub

Private Function ICommand_Execute() As Boolean
    ICommand_Execute = m_solicitudService.CrearSolicitud(m_solicitud)
    If ICommand_Execute Then
        m_solicitudCreada = m_solicitud.SolicitudId
    End If
End Function

Private Function ICommand_Undo() As Boolean
    If Len(m_solicitudCreada) > 0 Then
        ICommand_Undo = m_solicitudService.EliminarSolicitud(m_solicitudCreada)
    End If
End Function

Private Function ICommand_GetDescription() As String
    ICommand_GetDescription = "Crear solicitud: " & m_solicitud.DescripcionSolicitud
End Function
```

---

## 5. Manejo de Errores

### 5.1 Estrategia por Capas

#### Capa de Presentación
```vba
' Manejo básico - mostrar mensajes al usuario
Private Sub cmdGuardar_Click()
    On Error GoTo ErrorHandler
    
    ' Lógica del botón
    
    Exit Sub
    
ErrorHandler:
    Select Case Err.Number
        Case 3021 ' No current record
            MsgBox "No hay datos para guardar", vbExclamation
        Case 3314 ' Field required
            MsgBox "Faltan campos obligatorios", vbExclamation
        Case Else
            MsgBox "Error inesperado: " & Err.Description, vbCritical
    End Select
End Sub
```

#### Capa de Lógica de Negocio
```vba
' Manejo avanzado - log y propagación controlada
Private Function ISolicitudService_CrearSolicitud(solicitud As T_Solicitud) As Boolean
    On Error GoTo ErrorHandler
    
    ' Lógica de negocio
    
    Exit Function
    
ErrorHandler:
    ' Log del error
    m_logService.LogError "CSolicitudManager.CrearSolicitud", Err.Number, Err.Description
    
    ' Decidir si propagar o manejar
    Select Case Err.Number
        Case 3021, 3314 ' Errores de datos - no propagar
            ISolicitudService_CrearSolicitud = False
        Case Else ' Errores inesperados - propagar
            Err.Raise Err.Number, "CSolicitudManager.CrearSolicitud", Err.Description
    End Select
End Function
```

#### Capa de Servicios
```vba
' Manejo específico de datos - log detallado
Private Function ISolicitudRepository_Insertar(solicitud As T_Solicitud) As String
    On Error GoTo ErrorHandler
    
    ' Lógica de acceso a datos
    
    Exit Function
    
ErrorHandler:
    ' Log con contexto específico
    Dim contexto As String
    contexto = "Insertando solicitud ID: " & solicitud.SolicitudId & ", Expediente: " & solicitud.NumeroExpediente
    
    LogError "CSolicitudRepository.Insertar", Err.Number, Err.Description, contexto
    
    ' Limpiar recursos
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    
    ' Retornar valor de error
    ISolicitudRepository_Insertar = ""
End Function
```

### 5.2 Logging Centralizado

```vba
' Archivo: CLogService.cls
' Responsabilidad: Servicio centralizado de logging

Option Compare Database
Option Explicit
Implements ILogService

Private Function ILogService_LogError(modulo As String, errorNumber As Long, errorDescription As String, Optional contexto As String = "") As Boolean
    On Error Resume Next
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    
    Set db = CurrentDb()
    Set rs = db.OpenRecordset("TbLog_Errores", dbOpenDynaset)
    
    rs.AddNew
    rs!FechaHora = Now()
    rs!Modulo = modulo
    rs!NumeroError = errorNumber
    rs!DescripcionError = errorDescription
    rs!Contexto = contexto
    rs!Usuario = CurrentUser()
    rs!Severidad = "ERROR"
    rs.Update
    
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    
    ILogService_LogError = (Err.Number = 0)
End Function
```

---

## 6. Testing y Calidad

### 6.1 Estructura de Pruebas

```vba
' Archivo: TestSolicitudService.bas
' Responsabilidad: Pruebas unitarias para SolicitudService

Option Compare Database
Option Explicit

' === SETUP Y TEARDOWN ===
Public Sub SetupTest()
    ' Configurar mocks y datos de prueba
    ConfigurarMocks
    CrearDatosPrueba
End Sub

Public Sub TeardownTest()
    ' Limpiar datos de prueba
    LimpiarDatosPrueba
End Sub

' === PRUEBAS UNITARIAS ===
Public Sub Test_CrearSolicitud_ConDatosValidos_DebeRetornarTrue()
    ' Arrange
    Dim solicitudService As ISolicitudService
    Set solicitudService = CreateSolicitudServiceMock()
    
    Dim solicitud As New T_Solicitud
    solicitud.NumeroExpediente = "EXP-TEST-001"
    solicitud.TipoSolicitud = "PC"
    solicitud.DescripcionSolicitud = "Solicitud de prueba"
    solicitud.UsuarioCreador = "test@test.com"
    
    ' Act
    Dim resultado As Boolean
    resultado = solicitudService.CrearSolicitud(solicitud)
    
    ' Assert
    Assert.IsTrue resultado, "Crear solicitud con datos válidos debe retornar True"
End Sub

Public Sub Test_CrearSolicitud_SinDescripcion_DebeRetornarFalse()
    ' Arrange
    Dim solicitudService As ISolicitudService
    Set solicitudService = CreateSolicitudServiceMock()
    
    Dim solicitud As New T_Solicitud
    solicitud.NumeroExpediente = "EXP-TEST-001"
    solicitud.TipoSolicitud = "PC"
    solicitud.DescripcionSolicitud = "" ' Descripción vacía
    solicitud.UsuarioCreador = "test@test.com"
    
    ' Act
    Dim resultado As Boolean
    resultado = solicitudService.CrearSolicitud(solicitud)
    
    ' Assert
    Assert.IsFalse resultado, "Crear solicitud sin descripción debe retornar False"
End Sub
```

### 6.2 Mocks y Stubs

```vba
' Archivo: CSolicitudServiceMock.cls
' Responsabilidad: Mock para pruebas de SolicitudService

Option Compare Database
Option Explicit
Implements ISolicitudService

' === CONFIGURACIÓN DEL MOCK ===
Private m_shouldReturnTrue As Boolean
Private m_shouldThrowError As Boolean
Private m_errorToThrow As Long

Public Sub ConfigurarMock(returnTrue As Boolean, Optional throwError As Boolean = False, Optional errorNumber As Long = 0)
    m_shouldReturnTrue = returnTrue
    m_shouldThrowError = throwError
    m_errorToThrow = errorNumber
End Sub

' === IMPLEMENTACIÓN MOCK ===
Private Function ISolicitudService_CrearSolicitud(solicitud As T_Solicitud) As Boolean
    If m_shouldThrowError Then
        Err.Raise m_errorToThrow, "CSolicitudServiceMock", "Error simulado para pruebas"
    End If
    
    ISolicitudService_CrearSolicitud = m_shouldReturnTrue
End Function
```

---

## 7. Documentación de Código

### 7.1 Comentarios de Clase

```vba
'******************************************************************************
' Clase: CSolicitudManager
' Propósito: Implementa la lógica de negocio para la gestión de solicitudes
' Autor: Equipo CONDOR
' Fecha Creación: 2024-12-20
' Última Modificación: 2024-12-20
' 
' Responsabilidades:
' - Validar reglas de negocio para solicitudes
' - Coordinar operaciones entre repositorios y servicios externos
' - Gestionar el ciclo de vida de las solicitudes
' - Enviar notificaciones según el estado de las solicitudes
' 
' Dependencias:
' - ISolicitudRepository: Para acceso a datos de solicitudes
' - IExpedienteService: Para validar expedientes
' - INotificacionService: Para envío de notificaciones
' - ILogService: Para registro de eventos y errores
' 
' Notas:
' - Implementa el patrón Repository para separar lógica de negocio de acceso a datos
' - Utiliza inyección de dependencias para facilitar testing
' - Todas las operaciones son transaccionales
'******************************************************************************
```

### 7.2 Comentarios de Método

```vba
'******************************************************************************
' Método: CrearSolicitud
' Propósito: Crea una nueva solicitud aplicando todas las reglas de negocio
' 
' Parámetros:
'   solicitud (T_Solicitud): Objeto con los datos de la solicitud a crear
' 
' Retorna:
'   Boolean: True si la solicitud se creó exitosamente, False en caso contrario
' 
' Proceso:
'   1. Valida reglas de negocio específicas
'   2. Verifica existencia del expediente asociado
'   3. Asigna valores por defecto (ID, fechas, estados)
'   4. Persiste la solicitud en el repositorio
'   5. Envía notificaciones correspondientes
'   6. Registra la operación en el log
' 
' Excepciones:
'   - Propaga errores críticos del repositorio
'   - Maneja errores de validación sin propagar
' 
' Ejemplo de uso:
'   Dim solicitud As New T_Solicitud
'   solicitud.NumeroExpediente = "EXP-2024-001"
'   solicitud.TipoSolicitud = "PC"
'   
'   Dim resultado As Boolean
'   resultado = solicitudService.CrearSolicitud(solicitud)
'******************************************************************************
```

---

## 8. Performance y Optimización

### 8.1 Gestión de Memoria

```vba
' Siempre limpiar referencias a objetos
Private Sub LimpiarReferencias()
    Set m_solicitudService = Nothing
    Set m_expedienteService = Nothing
    Set m_solicitudActual = Nothing
End Sub

' Usar With para múltiples asignaciones
Private Sub ConfigurarSolicitud(solicitud As T_Solicitud)
    With solicitud
        .SolicitudId = GenerarId()
        .FechaCreacion = Now()
        .EstadoInterno = "Borrador"
        .EstadoRAC = "Pendiente"
        .UsuarioCreador = CurrentUser()
    End With
End Sub
```

### 8.2 Optimización de Consultas

```vba
' Usar parámetros en lugar de concatenación
Private Function BuscarSolicitudesPorExpediente(numeroExpediente As String) As DAO.Recordset
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    
    Set db = CurrentDb()
    Set qdf = db.CreateQueryDef("", "SELECT * FROM Tb_Solicitudes WHERE NumeroExpediente = [pExpediente]")
    qdf.Parameters("pExpediente") = numeroExpediente
    
    Set BuscarSolicitudesPorExpediente = qdf.OpenRecordset(dbOpenSnapshot)
End Function

' Cerrar recordsets inmediatamente después de usar
Private Function ContarSolicitudes(numeroExpediente As String) As Long
    Dim rs As DAO.Recordset
    Set rs = BuscarSolicitudesPorExpediente(numeroExpediente)
    
    ContarSolicitudes = rs.RecordCount
    
    rs.Close ' Cerrar inmediatamente
    Set rs = Nothing
End Function
```

---

## 9. Checklist de Calidad

### 9.1 Antes de Commit

- [ ] Código sigue convenciones de nomenclatura
- [ ] Todas las variables están declaradas (Option Explicit)
- [ ] Manejo de errores implementado apropiadamente
- [ ] Referencias a objetos se limpian correctamente
- [ ] Comentarios de documentación están completos
- [ ] Pruebas unitarias pasan exitosamente
- [ ] No hay código comentado sin justificación
- [ ] Separación de capas respetada

### 9.2 Code Review

- [ ] Lógica de negocio está en la capa correcta
- [ ] No hay dependencias circulares
- [ ] Interfaces están bien definidas
- [ ] Inyección de dependencias implementada
- [ ] Logging apropiado para debugging
- [ ] Performance considerado (consultas, memoria)
- [ ] Seguridad validada (SQL injection, etc.)
- [ ] Compatibilidad con versiones de Access

---

*Documento generado según la Especificación Funcional y Arquitectura CONDOR*  
*Versión: 1.0*  
*Fecha: Diciembre 2024*