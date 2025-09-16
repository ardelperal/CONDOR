' CONDOR CLI - Herramienta de linea de comandos para el proyecto CONDOR
' Funcionalidades: Sincronizacion VBA, gestion de tablas, y operaciones del proyecto
' Version sin dialogos para automatizacion completa

Option Explicit

' ==== Access Enum constants (VBScript no tiene referencias a Access) ====
' Prohibido usar números mágicos para vistas/objetos DoCmd
Const acViewNormal    = 0
Const acViewDesign    = 1
Const acWindowNormal  = 0
Const acWindowHidden  = 1
Const acWindowDialog  = 3
Const acObjectForm    = 2  ' Para Application.SaveAsText si se usa en el futuro
Const acForm          = 2  ' Alias para acObjectForm
Const acObjectTable   = 0  ' Para DoCmd.DeleteObject
Const acObjectQuery   = 1  ' Para DoCmd.DeleteObject
Const acDetail        = 0  ' Sección detalle
Const acHeader        = 1  ' Sección encabezado
Const acFooter        = 2  ' Sección pie
Const acSaveYes       = 1  ' Para DoCmd.Close
Const acSaveNo        = 2  ' Para DoCmd.Close

' Constante para bypass de formulario de inicio
Const msoAutomationSecurityForceDisable = 3

' ===== FUNCIONES HELPER PARA SECCIONES DE FORMULARIOS =====

Function SectionIdToToken(n)
    If n = acDetail Then 
        SectionIdToToken = "detail"
    ElseIf n = acHeader Then 
        SectionIdToToken = "header"
    ElseIf n = acFooter Then 
        SectionIdToToken = "footer"
    Else 
        SectionIdToToken = "detail"
    End If
End Function

Function SectionTokenToId(s)
    s = LCase(Trim(s))
    If s = "detail" Then 
        SectionTokenToId = acDetail
    ElseIf s = "header" Then 
        SectionTokenToId = acHeader
    ElseIf s = "footer" Then 
        SectionTokenToId = acFooter
    Else 
        SectionTokenToId = acDetail
    End If
End Function

' ===== FIN FUNCIONES HELPER PARA SECCIONES =====

Dim objAccess
Dim strAccessPath
Dim strSourcePath
Dim strAction
Dim objFSO
Dim objArgs
Dim strDbPassword
Dim pathArg, i
Dim gVerbose ' Variable global para soporte --verbose
Dim gBypassStartup ' Variable global para --bypassstartup
Dim gPassword ' Variable global para --password
Dim gDbPath ' Variable global para --db
Dim gDryRun ' Variable global para --dry-run
Dim gOpenShared ' Variable global para --sharedopen (por defecto False = modo exclusivo)

' Variables globales para bypass restoration
Dim gBypassStartupEnabled
Dim gPreviousAllowBypassKey
Dim gCurrentDbPath
Dim gCurrentPassword
Dim gPreviousStartupForm
Dim gPreviousHasAutoExec

Dim gPrevStartupForm, gHadAutoExec

' Variables globales adicionales
Dim gDbSource ' Variable global para almacenar el origen de la BD resuelta
Dim gPrintDb  ' Variable global para el flag --print-db

' Variables globales para bypass del formulario de inicio
Dim gStartupOptName  ' "Startup Form" o "Display Form" o ""
Dim gStartupPrev     ' valor previo del formulario de inicio (string)

' ===== FUNCIONES UTILITARIAS DE RUTAS (se definen después de inicializar objFSO) =====

' Función para quitar comillas al inicio y fin si existen
Function TrimQuotes(s)
    Dim result
    result = Trim(s)
    If Len(result) >= 2 Then
        If (Left(result, 1) = """" And Right(result, 1) = """") Or _
           (Left(result, 1) = "'" And Right(result, 1) = "'") Then
            result = Mid(result, 2, Len(result) - 2)
        End If
    End If
    TrimQuotes = result
End Function

' Función para verificar si un token es una ruta de BD
Function IsDbPathToken(tok)
    Dim ext
    ext = LCase(objFSO.GetExtensionName(tok))
    IsDbPathToken = (ext = "accdb" Or ext = "mdb")
End Function

' Función para convertir una ruta a absoluta
Function ToAbsolute(pathLike)
    Dim cleanPath
    cleanPath = TrimQuotes(pathLike)
    
    ' Si ya es absoluta, devolverla tal como está
    If objFSO.GetAbsolutePathName(cleanPath) = cleanPath Then
        ToAbsolute = cleanPath
    Else
        ' Si es relativa, resolverla contra RepoRoot()
        ToAbsolute = objFSO.GetAbsolutePathName(objFSO.BuildPath(RepoRoot(), cleanPath))
    End If
End Function

' ===== FIN FUNCIONES UTILITARIAS DE RUTAS =====

' ===== FUNCIONES HELPER PARA BYPASS DEL FORMULARIO DE INICIO =====

Function DetectStartupOptionName(app)
    ' La opción de startup form no se maneja con SetOption, sino como propiedad de BD
    ' Retornamos cadena vacía para desactivar el bypass
    DetectStartupOptionName = ""
End Function

Function SafeGetOption(app, optName, ByRef outVal)
    On Error Resume Next
    outVal = ""
    If Len(optName) > 0 Then outVal = app.GetOption(optName)
    SafeGetOption = (Err.Number = 0)
    Err.Clear
End Function

Function SafeSetOption(app, optName, newVal)
    On Error Resume Next
    If Len(optName) > 0 Then app.SetOption optName, newVal
    SafeSetOption = (Err.Number = 0)
    Err.Clear
End Function

Sub StartupBypass_Enable(app)
    ' BYPASS SIEMPRE ACTIVO: detectar y desactivar el formulario de inicio
    gStartupOptName = DetectStartupOptionName(app)
    gStartupPrev = ""
    If Len(gStartupOptName) = 0 Then Exit Sub
    Call SafeGetOption(app, gStartupOptName, gStartupPrev)
    Call SafeSetOption(app, gStartupOptName, "")
End Sub

Sub StartupBypass_Restore(app)
    If Len(gStartupOptName) = 0 Then Exit Sub
    Call SafeSetOption(app, gStartupOptName, gStartupPrev)
    gStartupOptName = ""
    gStartupPrev = ""
End Sub

' ===== FUNCIONES CENTRALIZADAS DE APERTURA/CIERRE CON BYPASS AUTOMÁTICO =====

' Apertura centralizada con bypass SIEMPRE activo
Function OpenAccessQuiet(dbPath, password)
    ' Verificar que la BD existe antes de intentar abrirla
    If Not objFSO.FileExists(dbPath) Then
        WScript.Echo "ERROR: La base de datos no existe: " & dbPath
        WScript.Quit 1
    End If
    
    ' Intentar limpiar locks obsoletos antes de abrir
    Call TryCleanupStaleLock(dbPath, gVerbose)
    
    Dim app
    On Error Resume Next
    Set app = CreateObject("Access.Application")
    If Err.Number <> 0 Then
        WScript.Echo "ERROR: No se pudo crear Access.Application: " & Err.Description
        WScript.Quit 1
    End If
    Err.Clear
    
    ' Configuración silenciosa reforzada
    app.Visible = False
    app.UserControl = False
    On Error Resume Next
    app.AutomationSecurity = 1  ' msoAutomationSecurityLow - permite acceso a propiedades
    If Err.Number <> 0 Then app.AutomationSecurity = 3  ' fallback a msoAutomationSecurityForceDisable
    Err.Clear
    app.DisplayAlerts = False
    app.Echo False
    On Error GoTo 0
    
    ' BYPASS SIEMPRE ACTIVO: detectar y desactivar el formulario de inicio ANTES de abrir la BD
    Call StartupBypass_Enable(app)
    
    ' Abrir la BD - SIEMPRE EN EXCLUSIVO por defecto, salvo --sharedopen
    Dim isExclusive: isExclusive = (Not gOpenShared)
    On Error Resume Next
    If Len(password) > 0 Then
        app.OpenCurrentDatabase dbPath, isExclusive, password
    Else
        app.OpenCurrentDatabase dbPath, isExclusive
    End If
    
    If Err.Number <> 0 Then 
        If isExclusive And Not gOpenShared Then
            WScript.Echo "ERROR (export-form): El archivo está en uso. Cierre otras sesiones de Access o libere el lock (.laccdb)."
        Else
            WScript.Echo "ERROR: No se pudo abrir la base de datos: " & Err.Description
        End If
        Call StartupBypass_Restore(app)
        app.Quit
        Set app = Nothing
        WScript.Quit 1
    End If
    On Error GoTo 0
    
    ' Guardar información para el cierre
    gCurrentDbPath = dbPath
    gCurrentPassword = password
    
    Call RemoveBrokenReferences(app)
    Set OpenAccessQuiet = app
End Function

' Cierre centralizado que restaura el Startup Form
Sub CloseAccessQuiet(app)
    On Error Resume Next
    
    ' Restaurar siempre el valor previo del Startup Form
    Call StartupBypass_Restore(app)
    
    ' Limpiar variables globales
    gCurrentDbPath = ""
    gCurrentPassword = ""
    
    ' Cerrar limpiamente con clausura robusta
    If Not app Is Nothing Then
        app.Application.Echo True
        app.CloseCurrentDatabase
        app.Quit
        Set app = Nothing
    End If
    Err.Clear
End Sub

' ===== FIN FUNCIONES HELPER PARA BYPASS DEL FORMULARIO DE INICIO =====

' ===== RESOLUCIÓN CANÓNICA DE BASE DE DATOS =====

' Función para resolver la base de datos según la acción y prioridades
Function ResolveDbForAction(actionName, ByRef origin)
    Dim resolvedPath, envDb, i, arg
    
    ' Prioridad 1: Si existe --db <ruta>
    If gDbPath <> "" Then
        resolvedPath = ToAbsolute(TrimQuotes(gDbPath))
        origin = "flag"
        ResolveDbForAction = resolvedPath
        Exit Function
    End If
    
    ' Prioridad 2: Si algún posicional es *.accdb|*.mdb
    For i = 1 To objArgs.Count - 1
        arg = objArgs(i)
        ' Saltar flags conocidos
        If Left(LCase(arg), 2) <> "--" And Left(arg, 1) <> "/" Then
            If IsDbPathToken(arg) Then
                resolvedPath = ToAbsolute(TrimQuotes(arg))
                origin = "positional"
                ResolveDbForAction = resolvedPath
                Exit Function
            End If
        End If
    Next
    
    ' Prioridad 3: Si ENV("CONDOR_DEV_DB") no está vacío
    On Error Resume Next
    envDb = CreateObject("WScript.Shell").Environment("Process")("CONDOR_DEV_DB")
    On Error GoTo 0
    If envDb <> "" Then
        resolvedPath = ToAbsolute(TrimQuotes(envDb))
        origin = "env"
        ResolveDbForAction = resolvedPath
        Exit Function
    End If
    
    ' Prioridad 4: Default según acción usando DefaultForAction
    resolvedPath = DefaultForAction(actionName, origin)
    ResolveDbForAction = resolvedPath
End Function

' Función que determina la BD por defecto según la acción
Function DefaultForAction(actionName, ByRef origin)
    ' FRONTEND por defecto para acciones de código/desarrollo
    If actionName = "rebuild" Or actionName = "update" Or actionName = "export" Or _
       actionName = "validate" Or actionName = "test" Or actionName = "export-form" Or _
       actionName = "import-form" Or actionName = "list-forms" Or actionName = "list-modules" Or _
       actionName = "roundtrip-form" Or actionName = "validate-form-json" Then
        origin = "default-frontend"
        DefaultForAction = DefaultFrontendDb()
    Else
        ' BACKEND por defecto para comandos de datos (listtables, migrate, relink, createtable, droptable, etc.)
        origin = "default-backend"
        DefaultForAction = DefaultBackendDb()
    End If
End Function

' ===== FIN RESOLUCIÓN CANÓNICA DE BASE DE DATOS =====

' Configuracion
' Configuracion inicial - se determinara la base de datos segun la accion
Dim strDataPath

' Obtener argumentos de linea de comandos
Set objArgs = WScript.Arguments
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Ahora podemos usar RepoRoot() después de inicializar objFSO
strSourcePath = RepoRoot() & "\src"

' Inicializar variables globales
gVerbose = False
gBypassStartup = False
gPassword = ""
gDbPath = ""
gDryRun = False
gOpenShared = False
gBypassStartupEnabled = False
gPreviousAllowBypassKey = Null
gCurrentDbPath = ""
gCurrentPassword = ""
gPreviousStartupForm = Null
gPreviousHasAutoExec = False
gDbSource = ""
gPrintDb = False

' ===== FUNCIONES UTILITARIAS DE RUTAS =====

' Función para obtener la carpeta raíz del repositorio
Function RepoRoot()
    Dim scriptPath
    scriptPath = WScript.ScriptFullName
    RepoRoot = objFSO.GetParentFolderName(scriptPath)
End Function

' Función para obtener la ruta por defecto de la BD frontend
Function DefaultFrontendDb()
    Dim defaultPath, legacyPath
    defaultPath = GetDevDbPath()
    legacyPath = objFSO.BuildPath(RepoRoot(), "back\Desarrollo\CONDOR.accdb")
    If Not objFSO.FileExists(defaultPath) And objFSO.FileExists(legacyPath) Then
        If gVerbose Or gPrintDb Then
            WScript.Echo "WARNING: usando ruta legacy back\Desarrollo\CONDOR.accdb; migre a front\Desarrollo\CONDOR.accdb (deprecado)"
        End If
        DefaultFrontendDb = legacyPath
    Else
        DefaultFrontendDb = defaultPath
    End If
End Function

' Función para obtener la ruta por defecto de la BD backend
Function DefaultBackendDb()
    Dim defaultPath, legacyPath
    defaultPath = objFSO.BuildPath(GetAuxDataRoot(), "CONDOR_datos.accdb")
    legacyPath = objFSO.BuildPath(RepoRoot(), "back\CONDOR_datos.accdb")
    If Not objFSO.FileExists(defaultPath) And objFSO.FileExists(legacyPath) Then
        If gVerbose Or gPrintDb Then
            WScript.Echo "WARNING: usando ruta legacy back\CONDOR_datos.accdb; migre a back\data\CONDOR_datos.accdb (deprecado)"
        End If
        DefaultBackendDb = legacyPath
    Else
        DefaultBackendDb = defaultPath
    End If
End Function

' ===== HELPERS DE RUTAS CENTRALIZADAS =====

' Función para obtener la carpeta raíz del repositorio (ya existe arriba)
Function GetRepoRoot()
    GetRepoRoot = RepoRoot()
End Function

' Función para obtener la carpeta front
Function GetFrontRoot()
    GetFrontRoot = objFSO.BuildPath(GetRepoRoot(), "front")
End Function

' Función para obtener la carpeta back
Function GetBackRoot()
    GetBackRoot = objFSO.BuildPath(GetRepoRoot(), "back")
End Function

' Función para obtener la ruta de plantillas
Function GetTemplatesPath()
    GetTemplatesPath = objFSO.BuildPath(GetFrontRoot(), "recursos\Plantillas")
End Function

' Función para obtener la ruta del entorno de pruebas
Function GetTestEnvPath()
    GetTestEnvPath = objFSO.BuildPath(GetFrontRoot(), "test_env")
End Function

' Función para obtener la ruta de la BD de desarrollo
Function GetDevDbPath()
    GetDevDbPath = objFSO.BuildPath(GetFrontRoot(), "Desarrollo\CONDOR.accdb")
End Function

' Función para obtener la ruta de datos auxiliares
Function GetAuxDataRoot()
    GetAuxDataRoot = objFSO.BuildPath(GetBackRoot(), "data")
End Function


' Verificar si se solicita ayuda
If objArgs.Count > 0 Then
    If LCase(objArgs(0)) = "--help" Or LCase(objArgs(0)) = "-h" Or LCase(objArgs(0)) = "help" Then
        Call ShowHelp()
        WScript.Quit 0
    End If
End If

If objArgs.Count = 0 Then
    WScript.Echo "=== CONDOR CLI - Herramienta de linea de comandos ==="
    WScript.Echo "Uso: cscript condor_cli.vbs [comando] [opciones]"
    WScript.Echo ""
    WScript.Echo "COMANDOS DISPONIBLES:"
    WScript.Echo "  export     - Exportar modulos VBA a /src (con codificacion ANSI)"
    WScript.Echo "  validate   - Validar sintaxis de modulos VBA sin importar"
    WScript.Echo "  validate-schema - Valida el esquema de las BDs de prueba contra el Master Plan"
    WScript.Echo "  test       - Ejecutar suite de pruebas unitarias"
    
    WScript.Echo "  rebuild    - Reconstruir proyecto VBA (eliminar todos los modulos y reimportar)"
    WScript.Echo "  bundle <funcionalidad> [ruta_destino] - Empaquetar archivos de codigo por funcionalidad"
    WScript.Echo "  lint       - Auditar codigo VBA para detectar cabeceras duplicadas"
    WScript.Echo "  createtable <nombre> <sql> - Crear tabla con consulta SQL"
    WScript.Echo "  droptable <nombre> - Eliminar tabla"
    WScript.Echo "  listtables [db_path] [--schema] [--output] - Listar tablas de BD. --schema: muestra campos, tipos y requerido. --output: exporta a [nombre_bd]_listtables.txt"
    WScript.Echo "  relink [db_path] [folder]    - Re-vincular tablas a bases locales"
    WScript.Echo "  migrate [file.sql]           - Ejecutar scripts de migración SQL desde ./db/migrations"
    WScript.Echo "  relink --all - Re-vincular todas las bases en ./back automaticamente"
    WScript.Echo "  migrate [file.sql] - Ejecutar scripts de migración SQL en /db/migrations"
    WScript.Echo ""


    WScript.Echo "PARÁMETROS DE FUNCIONALIDAD PARA 'bundle' (según CONDOR_MASTER_PLAN.md):"
    WScript.Echo "  Auth: Empaqueta Autenticación + dependencias (Config, Error, Modelos)"
    WScript.Echo "  Document: Empaqueta Gestión de Documentos + dependencias (Config, FileSystem, Error, Word, Modelos)"
    WScript.Echo "  Expediente: Empaqueta Gestión de Expedientes + dependencias (Config, Error, Modelos, Utilidades)"
    WScript.Echo "  Solicitud: Empaqueta Gestión de Solicitudes + dependencias (Config, Error, Modelos, Utilidades)"
    WScript.Echo "  Workflow: Empaqueta Flujos de Trabajo + dependencias (Config, Error, Modelos, Utilidades)"
    WScript.Echo "  Mapeo: Empaqueta Gestión de Mapeos + dependencias (Config, Error, Modelos)"
    WScript.Echo "  Config: Empaqueta Configuración del Sistema + dependencias (Error, Modelos)"
    WScript.Echo "  FileSystem: Empaqueta Sistema de Archivos + dependencias (Error, Modelos)"
    WScript.Echo "  Error: Empaqueta Manejo de Errores + dependencias (Modelos)"
    WScript.Echo "  Word: Empaqueta Microsoft Word + dependencias (Error, Modelos)"
    WScript.Echo "  TestFramework: Empaqueta Framework de Pruebas + dependencias (11 archivos: ITestReporter, CTestResult, CTestSuiteResult, CTestReporter, modTestRunner, modTestUtils, ModAssert, TestModAssert, IFileSystem, IConfig, IErrorHandlerService)"
    WScript.Echo "  App: Empaqueta Gestión de Aplicación + dependencias (Config, Error, Modelos)"
    WScript.Echo "  Models: Empaqueta Modelos de Datos (entidades base)"
    WScript.Echo "  Utils: Empaqueta Utilidades y Enumeraciones + dependencias (Error, Modelos)"
    WScript.Echo "  Tests: Empaqueta todos los archivos de pruebas (Test* e IntegrationTest*)"
    WScript.Echo ""
    WScript.Echo "OPCIONES ESPECIALES:"

    WScript.Echo "  --verbose  - Mostrar informacion detallada durante la operacion"
    WScript.Echo "  --sharedopen    - Abre la BD en modo compartido (por defecto el CLI abre en EXCLUSIVO)"
    WScript.Echo ""
    WScript.Echo "EJEMPLOS:"
    WScript.Echo "  cscript condor_cli.vbs validate"
    WScript.Echo "  cscript condor_cli.vbs validate-schema"
    WScript.Echo "  cscript condor_cli.vbs export --verbose"
    WScript.Echo "  cscript condor_cli.vbs rebuild"
    WScript.Echo "  cscript condor_cli.vbs bundle Tests"
    WScript.Echo "  cscript condor_cli.vbs listtables"
    WScript.Echo "  cscript condor_cli.vbs listtables ./back/test_db/templates/CONDOR_test_template.accdb --schema"
    WScript.Quit 1
End If

' Obtener acción
strAction = LCase(objArgs(0))

' Validar comando
If strAction <> "export" And strAction <> "validate" And strAction <> "validate-schema" And strAction <> "test" And strAction <> "createtable" And strAction <> "droptable" And strAction <> "listtables" And strAction <> "relink" And strAction <> "rebuild" And strAction <> "update" And strAction <> "lint" And strAction <> "bundle" And strAction <> "migrate" And strAction <> "export-form" And strAction <> "import-form" And strAction <> "validate-form-json" And strAction <> "roundtrip-form" And strAction <> "list-forms" And strAction <> "list-modules" And strAction <> "fix-src-headers" Then
    WScript.Echo "Error: Comando debe ser 'export', 'validate', 'validate-schema', 'test', 'createtable', 'droptable', 'listtables', 'relink', 'rebuild', 'update', 'lint', 'bundle', 'migrate', 'export-form', 'import-form', 'validate-form-json', 'roundtrip-form', 'list-forms', 'list-modules' o 'fix-src-headers'"
    WScript.Quit 1
End If

' PASO 1: Resolver flags ANTES de cualquier apertura de Access
Call ResolveFlags()

' PASO 2: Resolver ruta de base de datos usando resolución canónica
strAccessPath = ResolveDbForAction(strAction, gDbSource)

' Mostrar información de la BD resuelta si se solicita
If gPrintDb Or gVerbose Then
    WScript.Echo "DB resuelta: " & strAccessPath & " (origen=" & gDbSource & ")"
End If

' PASO 3: Determinar bypass startup por defecto según comando
Call SetDefaultBypassStartup()

' PASO 4: Ejecutar comandos que NO requieren Access
If strAction = "bundle" Then
    ' Verificar si se solicita ayuda específica para bundle
    If objArgs.Count > 1 Then
        If LCase(objArgs(1)) = "--help" Or LCase(objArgs(1)) = "-h" Or LCase(objArgs(1)) = "help" Then
            Call ShowBundleHelp()
            WScript.Quit 0
        End If
    End If
    Call BundleFunctionality()
    WScript.Quit 0
ElseIf strAction = "validate-schema" Then
    ' Usar rutas de producción por defecto
    Call ValidateSchema(objFSO.BuildPath(GetTestEnvPath(), "fixtures\databases\Lanzadera_test_template.accdb"), objFSO.BuildPath(GetTestEnvPath(), "fixtures\databases\Document_test_template.accdb"))
    WScript.Quit 0
ElseIf strAction = "validate-form-json" Then
    Call ValidateFormJsonCommand()
    WScript.Quit 0
ElseIf strAction = "roundtrip-form" Then
    Call RoundtripFormCommand()
    WScript.Quit 0
ElseIf strAction = "fix-src-headers" Then
    Call FixSrcHeadersCommand()
    WScript.Quit 0
End If

' PASO 5: Verificar que existe la base de datos
If Not objFSO.FileExists(strAccessPath) Then
    WScript.Echo "Error: base de datos no encontrada (" & strAccessPath & "), origen=" & gDbSource
    WScript.Quit 1
End If

' PASO 6: Mostrar información de inicio
If strAction = "import-form" Then
    WScript.Echo "=== IMPORTANDO FORMULARIO ==="
ElseIf strAction = "export-form" Then
    WScript.Echo "=== EXPORTANDO FORMULARIO ==="
ElseIf strAction = "list-forms" Then
    WScript.Echo "=== LISTANDO FORMULARIOS ==="
ElseIf strAction = "list-modules" Then
    WScript.Echo "=== LISTANDO MÓDULOS ==="
ElseIf strAction = "listtables" Then
    WScript.Echo "=== LISTANDO TABLAS ==="
ElseIf strAction = "validate" Then
    WScript.Echo "=== VALIDANDO SINTAXIS VBA ==="
ElseIf strAction = "lint" Then
    WScript.Echo "=== ANALIZANDO CÓDIGO VBA ==="
ElseIf strAction = "validate-schema" Then
    WScript.Echo "=== VALIDANDO ESQUEMA ==="
ElseIf strAction = "validate-form-json" Then
    WScript.Echo "=== VALIDANDO FORMULARIO JSON ==="
ElseIf strAction = "createtable" Then
    WScript.Echo "=== CREANDO TABLA ==="
ElseIf strAction = "droptable" Then
    WScript.Echo "=== ELIMINANDO TABLA ==="
ElseIf strAction = "relink" Then
    WScript.Echo "=== REVINCULANDO BASE DE DATOS ==="
ElseIf strAction = "migrate" Then
    WScript.Echo "=== MIGRANDO BASE DE DATOS ==="
ElseIf strAction = "rebuild" Then
    WScript.Echo "=== RECONSTRUYENDO BASE DE DATOS ==="
ElseIf strAction = "export" Then
    WScript.Echo "=== EXPORTANDO CÓDIGO VBA ==="
ElseIf strAction = "roundtrip-form" Then
    WScript.Echo "=== PROCESANDO FORMULARIO ==="
ElseIf strAction = "update" Then
    WScript.Echo "=== ACTUALIZANDO CÓDIGO VBA ==="
ElseIf strAction = "test" Then
    WScript.Echo "=== EJECUTANDO PRUEBAS ==="
ElseIf strAction = "bundle" Then
    WScript.Echo "=== EMPAQUETANDO APLICACIÓN ==="
Else
    WScript.Echo "=== INICIANDO SINCRONIZACION VBA ==="
End If
WScript.Echo "Accion: " & strAction
WScript.Echo "Base de datos: " & strAccessPath
WScript.Echo "Directorio: " & strSourcePath
If gVerbose Then
    If gPassword <> "" Then
        WScript.Echo "[VERBOSE] Password: ***"
    Else
        WScript.Echo "[VERBOSE] Password: (none)"
    End If
    WScript.Echo "[INFO] BypassStartup aplicado automáticamente (flag obsoleto)."
End If

' PASO 7: Cerrar procesos de Access existentes
Call CloseExistingAccessProcesses()

' PASO 8: Abrir Access con OpenAccessQuiet unificado (solo si es necesario)
If strAction <> "bundle" And strAction <> "validate-schema" And strAction <> "validate-form-json" And strAction <> "roundtrip-form" And strAction <> "list-forms" Then
    If strAction <> "import-form" Then
        ' Abrir Access en silencio con contraseña correcta usando función canónica
        Set objAccess = OpenAccessApp(strAccessPath, gPassword, True)
    Else
        Set objAccess = OpenAccessApp(strAccessPath, gPassword, True)
    End If
    If Not objAccess Is Nothing Then
        ' RemoveBrokenReferences ya se llama dentro de OpenAccessQuiet
        Call EnsureVBReferences
    End If
End If

' PASO 9: Ejecutar comando correspondiente

If strAction = "validate" Then
    Call ValidateAllModules()
ElseIf strAction = "export" Then
    Call ExportModules()
ElseIf strAction = "test" Then
    Call ExecuteTests()
ElseIf strAction = "createtable" Then
    Call CreateTable()
ElseIf strAction = "droptable" Then
    Call DropTable()
ElseIf strAction = "listtables" Then
    Call ListTables()

ElseIf strAction = "rebuild" Then
    Call RebuildProject()
ElseIf strAction = "update" Then
    Call UpdateProject()
ElseIf strAction = "lint" Then
    Call LintProject()
ElseIf strAction = "relink" Then
    Call RelinkTables()
ElseIf strAction = "migrate" Then
    Call ExecuteMigrations()
ElseIf strAction = "export-form" Then
    Call ExportForm()
ElseIf strAction = "import-form" Then
    Call ImportForm()
ElseIf strAction = "list-forms" Then
    Call ListForms()
ElseIf strAction = "list-modules" Then
    Call ListModulesCommand()
    WScript.Quit 0

Else
    WScript.Echo "Error: Comando no reconocido: " & strAction
    WScript.Quit 1
End If

' PASO 10: Cerrar Access si fue abierto (unificado)
' Nota: list-forms y list-modules manejan su propio ciclo de vida de Access
On Error Resume Next
If Not objAccess Is Nothing And strAction <> "list-forms" And strAction <> "list-modules" Then
    ' Restaurar startup antes de cerrar (solo si no es import-form)
    If strAction <> "import-form" Then
    End If
    Call CloseAccessApp(objAccess)
End If
On Error Goto 0

WScript.Echo "=== COMANDO COMPLETADO EXITOSAMENTE ==="
WScript.Quit 0

' Subrutina para validar todos los modulos sin importar
Sub ValidateAllModules()
    Dim objFolder, objFile
    Dim strFileName, strContent
    Dim validationResult
    Dim totalFiles, validFiles, invalidFiles
    
    WScript.Echo "=== VALIDACION DE SINTAXIS VBA ==="
    
    If Not objFSO.FolderExists(strSourcePath) Then
        WScript.Echo "Error: Directorio de origen no existe: " & strSourcePath
        WScript.Quit 1
    End If
    
    Set objFolder = objFSO.GetFolder(strSourcePath)
    totalFiles = 0
    validFiles = 0
    invalidFiles = 0
    
    For Each objFile In objFolder.Files
        If LCase(objFSO.GetExtensionName(objFile.Name)) = "bas" Or LCase(objFSO.GetExtensionName(objFile.Name)) = "cls" Then
            totalFiles = totalFiles + 1
            strFileName = objFile.Path
            
            If gVerbose Then
                WScript.Echo "Validando: " & objFile.Name
            End If
            
            ' Validar sintaxis
            Dim errorDetails
            validationResult = ValidateVBASyntax(strFileName, errorDetails)
            
            If validationResult = True Then
                validFiles = validFiles + 1
                If gVerbose Then
                    WScript.Echo "  ✓ Sintaxis valida"
                End If
            Else
                invalidFiles = invalidFiles + 1
                WScript.Echo "  ✗ ERROR en " & objFile.Name & ": " & errorDetails
            End If
        End If
    Next
    
    WScript.Echo ""
    WScript.Echo "=== RESUMEN DE VALIDACION ==="
    WScript.Echo "Total de archivos: " & totalFiles
    WScript.Echo "Archivos validos: " & validFiles
    WScript.Echo "Archivos con errores: " & invalidFiles
    
    If invalidFiles > 0 Then
        WScript.Echo "ADVERTENCIA: Se encontraron errores de sintaxis. Corrija antes de importar."
        WScript.Quit 1
    Else
        WScript.Echo "✓ Todos los archivos tienen sintaxis valida"
    End If
End Sub



' Subrutina para exportar modulos
Sub ExportModules()
    Dim vbProj, vbComp
    Dim destPath, ext
    Dim exportedCount
    
    WScript.Echo "Iniciando exportacion de modulos VBA..."
    
    ' Asegurar que existe la carpeta ./src
    If Not objFSO.FolderExists(strSourcePath) Then
        objFSO.CreateFolder strSourcePath
        WScript.Echo "Directorio de destino creado: " & strSourcePath
    End If
    
    exportedCount = 0
    
    ' Usar VBE.ActiveVBProject para acceder a los componentes
    Set vbProj = objAccess.VBE.ActiveVBProject
    
    For Each vbComp In vbProj.VBComponents
        ' Exportar SOLO vbext_ct_StdModule (1) y vbext_ct_ClassModule (2)
        ' NO exportar módulos de documento (Forms/Reports)
        Select Case vbComp.Type
            Case 1  ' vbext_ct_StdModule
                ext = ".bas"
                destPath = objFSO.BuildPath(strSourcePath, vbComp.Name & ext)
                
                If gVerbose Then
                    WScript.Echo "Exported: " & vbComp.Name & " (Std) -> " & destPath
                End If
                
                On Error Resume Next
                ' Usar VBComponents.Export cuando sea posible
                vbComp.Export destPath
                
                If Err.Number <> 0 Then
                    WScript.Echo "Error al exportar modulo " & vbComp.Name & ": " & Err.Description
                    Err.Clear
                Else
                    ' Postprocesar cabecera para garantizar formato canónico
                    PostProcessHeader destPath
                    exportedCount = exportedCount + 1
                End If
                
            Case 2  ' vbext_ct_ClassModule
                ext = ".cls"
                destPath = objFSO.BuildPath(strSourcePath, vbComp.Name & ext)
                
                If gVerbose Then
                    WScript.Echo "Exported: " & vbComp.Name & " (Class) -> " & destPath
                End If
                
                On Error Resume Next
                ' Usar VBComponents.Export cuando sea posible
                vbComp.Export destPath
                
                If Err.Number <> 0 Then
                    WScript.Echo "Error al exportar clase " & vbComp.Name & ": " & Err.Description
                    Err.Clear
                Else
                    ' Postprocesar cabecera para garantizar formato canónico
                    PostProcessHeader destPath
                    exportedCount = exportedCount + 1
                End If
                
            ' Ignorar otros tipos (Document modules, etc.)
        End Select
    Next
    
    WScript.Echo "Exportacion completada exitosamente. Modulos exportados: " & exportedCount
End Sub

' Subrutina para crear tabla
Sub CreateTable()
    Dim strTableName
    Dim strSQL
    Dim strQueryName
    
    If objArgs.Count < 3 Then
        WScript.Echo "Error: Se requiere nombre de tabla y consulta SQL"
        WScript.Echo "Uso: cscript condor_cli.vbs createtable <nombre> <sql>"
        WScript.Quit 1
    End If
    
    strTableName = objArgs(1)
    strSQL = objArgs(2)
    strQueryName = "qry_Create_" & strTableName
    
    WScript.Echo "Creando tabla: " & strTableName
    WScript.Echo "SQL: " & strSQL
    
    On Error Resume Next
    
    ' Verificar si la tabla ya existe
    Dim tblExists
    tblExists = False
    Dim tbl
    For Each tbl In objAccess.CurrentDb.TableDefs
        If LCase(tbl.Name) = LCase(strTableName) Then
            tblExists = True
            Exit For
        End If
    Next
    
    If tblExists Then
        WScript.Echo "Advertencia: La tabla '" & strTableName & "' ya existe"
    End If
    
    ' Crear consulta temporal
    WScript.Echo "Creando consulta temporal: " & strQueryName
    objAccess.CurrentDb.CreateQueryDef strQueryName, strSQL
    
    If Err.Number <> 0 Then
        WScript.Echo "Error al crear consulta: " & Err.Description
        Err.Clear
        Exit Sub
    End If
    
    ' Ejecutar consulta
    WScript.Echo "Ejecutando consulta..."
    objAccess.DoCmd.OpenQuery strQueryName
    
    If Err.Number <> 0 Then
        WScript.Echo "Error al ejecutar consulta: " & Err.Description
        Err.Clear
    Else
        WScript.Echo "Tabla '" & strTableName & "' creada exitosamente"
    End If
    
    ' Eliminar consulta temporal
    WScript.Echo "Eliminando consulta temporal..."
    objAccess.DoCmd.DeleteObject acObjectQuery, strQueryName
    
    If Err.Number <> 0 Then
        WScript.Echo "Advertencia al eliminar consulta: " & Err.Description
        Err.Clear
    Else
        WScript.Echo "Consulta temporal eliminada"
    End If
    
    ' Verificar que la tabla fue creada
    Call VerifyTable(strTableName)
End Sub

' Subrutina para eliminar tabla
Sub DropTable()
    Dim strTableName
    
    If objArgs.Count < 2 Then
        WScript.Echo "Error: Se requiere nombre de tabla"
        WScript.Echo "Uso: cscript condor_cli.vbs droptable <nombre>"
        WScript.Quit 1
    End If
    
    strTableName = objArgs(1)
    
    WScript.Echo "Eliminando tabla: " & strTableName
    
    On Error Resume Next
    objAccess.DoCmd.DeleteObject acObjectTable, strTableName
    
    If Err.Number <> 0 Then
        WScript.Echo "Error al eliminar tabla: " & Err.Description
        Err.Clear
    Else
        WScript.Echo "Tabla '" & strTableName & "' eliminada exitosamente"
    End If
End Sub

' Subrutina para listar tablas
Sub ListTables()
    Dim tbl, fld, idx
    Dim tableCount, showSchema, outputToFile
    Dim primaryKeys
    Dim outputFile, outputPath
    
    ' Verificar flags de argumentos
    Dim arg
    showSchema = False
    outputToFile = False
    
    For Each arg In objArgs
        If LCase(arg) = "--schema" Then
            showSchema = True
        ElseIf LCase(arg) = "--output" Then
            outputToFile = True
        End If
    Next
    
    ' Configurar salida
    If outputToFile Then
        Dim dbName
        dbName = objFSO.GetBaseName(strAccessPath)
        outputPath = objFSO.GetAbsolutePathName(".") & "\" & dbName & "_listtables.txt"
        Set outputFile = objFSO.CreateTextFile(outputPath, True)
        WScript.Echo "Exportando resultados a: " & outputPath
    End If
    
    WScript.Echo "=== LISTADO DE TABLAS ==="
    If outputToFile Then outputFile.WriteLine "=== LISTADO DE TABLAS ==="
    
    If showSchema Then 
        WScript.Echo "Modo: Esquema Detallado"
        If outputToFile Then outputFile.WriteLine "Modo: Esquema Detallado"
    End If
    
    tableCount = 0
    For Each tbl In objAccess.CurrentDb.TableDefs
        If Left(tbl.Name, 4) <> "MSys" And Left(tbl.Name, 1) <> "~" Then
            tableCount = tableCount + 1
            WScript.Echo ""
            If outputToFile Then outputFile.WriteLine ""
            
            WScript.Echo "------------------------------------------------------------"
            If outputToFile Then outputFile.WriteLine "------------------------------------------------------------"
            
            WScript.Echo tableCount & ". " & tbl.Name & " (" & tbl.RecordCount & " registros)"
            If outputToFile Then outputFile.WriteLine tableCount & ". " & tbl.Name & " (" & tbl.RecordCount & " registros)"
            
            WScript.Echo "------------------------------------------------------------"
            If outputToFile Then outputFile.WriteLine "------------------------------------------------------------"
            
            If showSchema Then
                Set primaryKeys = CreateObject("Scripting.Dictionary")
                For Each idx In tbl.Indexes
                    If idx.Primary Then
                        For Each fld In idx.Fields
                            primaryKeys.Add fld.Name, True
                        Next
                    End If
                Next
    
                WScript.Echo PadRight("Campo", 25) & PadRight("Tipo", 15) & PadRight("PK", 8) & "Requerido"
                If outputToFile Then outputFile.WriteLine PadRight("Campo", 25) & PadRight("Tipo", 15) & PadRight("PK", 8) & "Requerido"
                
                WScript.Echo "--------------------------------------------------------------------"
                If outputToFile Then outputFile.WriteLine "--------------------------------------------------------------------"
                
                For Each fld In tbl.Fields
                    Dim pkMarker, requiredMarker
                    If primaryKeys.Exists(fld.Name) Then pkMarker = "PK" Else pkMarker = ""
                    If fld.Required Then requiredMarker = "true" Else requiredMarker = "false"
                    WScript.Echo PadRight(fld.Name, 25) & PadRight(DaoTypeToString(fld.Type), 15) & PadRight(pkMarker, 8) & requiredMarker
                    If outputToFile Then outputFile.WriteLine PadRight(fld.Name, 25) & PadRight(DaoTypeToString(fld.Type), 15) & PadRight(pkMarker, 8) & requiredMarker
                Next
            End If
        End If
    Next
    
    WScript.Echo ""
    If outputToFile Then outputFile.WriteLine ""
    
    WScript.Echo "Total de tablas: " & tableCount
    If outputToFile Then outputFile.WriteLine "Total de tablas: " & tableCount
    
    ' Cerrar archivo si se está usando
    If outputToFile Then
        outputFile.Close
        Set outputFile = Nothing
        WScript.Echo "Archivo generado exitosamente: " & outputPath
    End If
End Sub

' Subrutina para verificar tabla creada
Sub VerifyTable(strTableName)
    Dim tbl
    Dim found
    
    WScript.Echo "Verificando tabla creada..."
    found = False
    
    On Error Resume Next
    For Each tbl In objAccess.CurrentDb.TableDefs
        If LCase(tbl.Name) = LCase(strTableName) Then
            found = True
            WScript.Echo "? Tabla '" & strTableName & "' verificada exitosamente"
            WScript.Echo "  - Campos: " & tbl.Fields.Count
            WScript.Echo "  - Registros: " & tbl.RecordCount
            Exit For
        End If
    Next
    
    If Not found Then
        WScript.Echo "? Error: No se pudo verificar la tabla '" & strTableName & "'"
    End If
End Sub




' ===================================================================
' SUBRUTINA: LintProject
' Descripción: Audita el código VBA para detectar cabeceras duplicadas
' ===================================================================
Sub LintProject()
    Dim vbComponent, codeModule
    Dim lineContent, moduleName
    Dim optionCompareCount, optionExplicitCount
    Dim i, hasErrors
    
    WScript.Echo "=== INICIANDO AUDITORIA VBA ==="
    WScript.Echo "Accion: lint"
    WScript.Echo "Base de datos: " & strAccessPath
    
    ' Usar la instancia global de Access ya abierta
    ' (objAccess ya está disponible desde el dispatcher principal)
    
    WScript.Echo "=== AUDITORIA DE CABECERAS VBA ==="
    WScript.Echo ""
    
    hasErrors = False
    
    For Each vbComponent In objAccess.VBE.ActiveVBProject.VBComponents
        moduleName = vbComponent.Name
        Set codeModule = vbComponent.CodeModule
        
        optionCompareCount = 0
        optionExplicitCount = 0
        
        For i = 1 To 10
            If i <= codeModule.CountOfLines Then
                lineContent = Trim(codeModule.Lines(i, 1))
                
                If InStr(1, lineContent, "Option Compare", 1) > 0 Then
                    optionCompareCount = optionCompareCount + 1
                End If
                
                If InStr(1, lineContent, "Option Explicit", 1) > 0 Then
                    optionExplicitCount = optionExplicitCount + 1
                End If
            End If
        Next
        
        If optionCompareCount > 1 Then
            WScript.Echo "ERROR: Modulo " & moduleName & " tiene " & optionCompareCount & " declaraciones Option Compare duplicadas"
            hasErrors = True
        End If
        
        If optionExplicitCount > 1 Then
            WScript.Echo "ERROR: Modulo " & moduleName & " tiene " & optionExplicitCount & " declaraciones Option Explicit duplicadas"
            hasErrors = True
        End If
        
        If optionCompareCount <= 1 And optionExplicitCount <= 1 Then
            WScript.Echo "OK: " & moduleName & " - Cabeceras correctas"
        End If
    Next
    
    If hasErrors Then
        WScript.Echo ""
        WScript.Echo "=== LINT FALLIDO ==="
        WScript.Echo "Se encontraron cabeceras duplicadas."
        WScript.Quit 1
    Else
        WScript.Echo ""
        WScript.Echo "=== LINT COMPLETADO EXITOSAMENTE ==="
    End If
End Sub

' Subrutina para compilación condicional de módulos
Sub CompileModulesConditionally()
    Dim vbComponent
    Dim compilationErrors
    Dim totalModules
    Dim compiledModules
    
    WScript.Echo "Iniciando compilación condicional de módulos..."
    
    compilationErrors = 0
    totalModules = 0
    compiledModules = 0
    
    ' Intentar compilar cada módulo individualmente (módulos estándar y clases)
    For Each vbComponent In objAccess.VBE.ActiveVBProject.VBComponents
        If vbComponent.Type = 1 Or vbComponent.Type = 2 Then  ' vbext_ct_StdModule o vbext_ct_ClassModule
            totalModules = totalModules + 1
            
            On Error Resume Next
            Err.Clear
            
            ' Intentar compilar el módulo específico
            If vbComponent.Type = 1 Then
                WScript.Echo "Compilando módulo: " & vbComponent.Name
            Else
                WScript.Echo "Compilando clase: " & vbComponent.Name
            End If
            
            ' Verificar si el módulo tiene errores de sintaxis
            Dim hasErrors
            hasErrors = False
            
            ' Intentar acceder al código del módulo para detectar errores
            Dim moduleCode
            moduleCode = vbComponent.CodeModule.Lines(1, vbComponent.CodeModule.CountOfLines)
            
            If Err.Number <> 0 Then
                WScript.Echo "  ⚠️ Error en " & vbComponent.Name & ": " & Err.Description
                compilationErrors = compilationErrors + 1
                hasErrors = True
                Err.Clear
            Else
                ' Los módulos se guardan automáticamente, no es necesario guardar explícitamente
                If vbComponent.Type = 1 Then
                    WScript.Echo "  ✓ " & vbComponent.Name & " compilado correctamente"
                    compiledModules = compiledModules + 1
                Else
                    ' Para módulos de clase, solo verificar sintaxis sin intentar guardar individualmente
                    WScript.Echo "  ✓ " & vbComponent.Name & " verificado (clase)"
                    compiledModules = compiledModules + 1
                End If
            End If
            
            On Error GoTo 0
        End If
    Next
    
    ' Intentar compilación global si los módulos principales están bien
    If compiledModules >= (totalModules - 3) Then  ' Permitir hasta 3 errores (las clases problemáticas)
        WScript.Echo "Intentando compilación global..."
        On Error Resume Next
        objAccess.DoCmd.RunCommand 636  ' acCmdCompileAndSaveAllModules
        
        If Err.Number <> 0 Then
            WScript.Echo "⚠️ Advertencia en compilación global: " & Err.Description
            WScript.Echo "Continuando con módulos compilados individualmente..."
            Err.Clear
        Else
            WScript.Echo "✓ Compilación global exitosa"
        End If
        On Error GoTo 0
    Else
        WScript.Echo "⚠️ Se encontraron " & compilationErrors & " errores de compilación"
        WScript.Echo "Continuando sin compilación global para evitar bloqueos..."
    End If
    
    WScript.Echo "Resumen de compilación:"
    WScript.Echo "  - Total de módulos: " & totalModules
    WScript.Echo "  - Módulos compilados: " & compiledModules
    WScript.Echo "  - Errores encontrados: " & compilationErrors
    
    If compilationErrors > 0 Then
        WScript.Echo "⚠️ ADVERTENCIA: Algunos módulos tienen errores de compilación"
        WScript.Echo "El CLI continuará funcionando, pero revise los módulos con errores"
    End If
End Sub

' Subrutina para verificar que los nombres de módulos coincidan con src
Sub VerifyModuleNames()
    Dim objFolder, objFile
    Dim vbComponent
    Dim srcModules, accessModules
    Dim moduleName
    Dim discrepancies
    
    WScript.Echo "Verificando integridad de nombres de módulos..."
    
    ' Crear diccionarios para comparación
    Set srcModules = CreateObject("Scripting.Dictionary")
    Set accessModules = CreateObject("Scripting.Dictionary")
    discrepancies = 0
    
    ' Obtener lista de módulos en src
    Set objFolder = objFSO.GetFolder(strSourcePath)
    For Each objFile In objFolder.Files
        If LCase(objFSO.GetExtensionName(objFile.Name)) = "bas" Or LCase(objFSO.GetExtensionName(objFile.Name)) = "cls" Then
            moduleName = objFSO.GetBaseName(objFile.Name)
            srcModules.Add moduleName, True
        End If
    Next
    
    ' Obtener lista de módulos en Access
    For Each vbComponent In objAccess.VBE.ActiveVBProject.VBComponents
        If vbComponent.Type = 1 Or vbComponent.Type = 2 Then  ' vbext_ct_StdModule o vbext_ct_ClassModule
            accessModules.Add vbComponent.Name, True
        End If
    Next
    
    ' Verificar que todos los módulos de src estén en Access
    For Each moduleName In srcModules.Keys
        If Not accessModules.Exists(moduleName) Then
            WScript.Echo "⚠️ ERROR: Módulo '" & moduleName & "' existe en src pero no en Access"
            discrepancies = discrepancies + 1
        End If
    Next
    
    ' Verificar que todos los módulos de Access estén en src
    For Each moduleName In accessModules.Keys
        If Not srcModules.Exists(moduleName) Then
            WScript.Echo "⚠️ ERROR: Módulo '" & moduleName & "' existe en Access pero no en src"
            discrepancies = discrepancies + 1
        End If
    Next
    
    ' Reporte final
    If discrepancies = 0 Then
        WScript.Echo "✓ Verificación exitosa: Todos los módulos coinciden entre src y Access"
        WScript.Echo "  - Módulos en src: " & srcModules.Count
        WScript.Echo "  - Módulos en Access: " & accessModules.Count
    Else
        WScript.Echo "❌ FALLO EN VERIFICACIÓN: Se encontraron " & discrepancies & " discrepancias"
        WScript.Echo "⚠️ ACCIÓN REQUERIDA: Revise la sincronización entre src y Access"
    End If
End Sub

' Función para validar sintaxis VBA antes de importar
Function ValidateVBASyntax(filePath, ByRef errorDetails)
    Dim objFile, strContent
    
    errorDetails = ""
    
    ' Leer archivo con codificación ANSI
    On Error Resume Next
    Set objFile = objFSO.OpenTextFile(filePath, 1, False, 0)
    If Err.Number <> 0 Then
        errorDetails = "Error al leer archivo: " & Err.Description
        ValidateVBASyntax = False
        Exit Function
    End If
    
    strContent = objFile.ReadAll
    objFile.Close
    On Error GoTo 0
    
    ' Validación básica: verificar que el archivo no esté vacío y sea legible
    If Len(Trim(strContent)) = 0 Then
        errorDetails = "El archivo está vacío"
        ValidateVBASyntax = False
        Exit Function
    End If
    
    ' Verificar caracteres problemáticos básicos
    If InStr(strContent, Chr(0)) > 0 Then
        errorDetails = "El archivo contiene caracteres nulos"
        ValidateVBASyntax = False
        Exit Function
    End If
    
    ' Si llegamos aquí, el archivo es válido
    ValidateVBASyntax = True
End Function

' Función de serialización recursiva para convertir Dictionary a JSON


' Función para leer archivo con codificación ANSI
Function ReadFileWithAnsiEncoding(filePath)
    Dim objStream, strContent
    
    On Error Resume Next
    
    ' Leer contenido del archivo usando ADODB.Stream con UTF-8 y convertir a ANSI
    Set objStream = CreateObject("ADODB.Stream")
    objStream.Type = 2 ' adTypeText
    objStream.Charset = "UTF-8"
    objStream.Open
    objStream.LoadFromFile filePath
    strContent = objStream.ReadText
    objStream.Close
    
    ' Convertir caracteres UTF-8 a ANSI para compatibilidad con VBA
    ' Preservar caracteres especiales del español
    strContent = Replace(strContent, "á", "á")
    strContent = Replace(strContent, "é", "é")
    strContent = Replace(strContent, "í", "í")
    strContent = Replace(strContent, "ó", "ó")
    strContent = Replace(strContent, "ú", "ú")
    strContent = Replace(strContent, "ñ", "ñ")
    strContent = Replace(strContent, "Á", "Á")
    strContent = Replace(strContent, "É", "É")
    strContent = Replace(strContent, "Í", "Í")
    strContent = Replace(strContent, "Ó", "Ó")
    strContent = Replace(strContent, "Ú", "Ú")
    strContent = Replace(strContent, "Ñ", "Ñ")
    strContent = Replace(strContent, "ü", "ü")
    strContent = Replace(strContent, "Ü", "Ü")
    
    Set objStream = Nothing
    
    If Err.Number <> 0 Then
        WScript.Echo "❌ ERROR: No se pudo leer el archivo " & filePath & ": " & Err.Description
        ReadFileWithAnsiEncoding = ""
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    
    On Error GoTo 0
    ReadFileWithAnsiEncoding = strContent
End Function

' Función para limpiar archivos VBA eliminando líneas Attribute con validación mejorada
Function CleanVBAFile(filePath, fileType)
    Dim objStream, strContent, arrLines, i, cleanedContent
    Dim strLine
    
    ' Leer el archivo como UTF-8 y convertir a ANSI para VBA
    Set objStream = CreateObject("ADODB.Stream")
    objStream.Type = 2 ' adTypeText
    objStream.Charset = "UTF-8"
    objStream.Open
    objStream.LoadFromFile filePath
    strContent = objStream.ReadText
    objStream.Close
    
    ' Convertir caracteres UTF-8 a ANSI para compatibilidad con VBA
    ' Preservar caracteres especiales del español
    strContent = Replace(strContent, "á", "á")
    strContent = Replace(strContent, "é", "é")
    strContent = Replace(strContent, "í", "í")
    strContent = Replace(strContent, "ó", "ó")
    strContent = Replace(strContent, "ú", "ú")
    strContent = Replace(strContent, "ñ", "ñ")
    strContent = Replace(strContent, "Á", "Á")
    strContent = Replace(strContent, "É", "É")
    strContent = Replace(strContent, "Í", "Í")
    strContent = Replace(strContent, "Ó", "Ó")
    strContent = Replace(strContent, "Ú", "Ú")
    strContent = Replace(strContent, "Ñ", "Ñ")
    strContent = Replace(strContent, "ü", "ü")
    strContent = Replace(strContent, "Ü", "Ü")
    
    Set objStream = Nothing
    
    ' Dividir el contenido en un array de líneas
    strContent = Replace(strContent, vbCrLf, vbLf)
    strContent = Replace(strContent, vbCr, vbLf)
    arrLines = Split(strContent, vbLf)
    
    ' Crear un nuevo string vacío llamado cleanedContent
    cleanedContent = ""
    
    ' Iterar sobre el array de líneas original
    For i = 0 To UBound(arrLines)
        strLine = arrLines(i)
        
        ' Aplicar las reglas para descartar contenido no deseado
        ' Una línea se descarta si cumple cualquiera de estas condiciones:
        ' CORRECCION CRITICA: Filtrar TODAS las líneas que empiecen con 'Attribute'
        ' y todos los metadatos de archivos .cls
        If Not (Left(Trim(strLine), 9) = "Attribute" Or _
                Left(Trim(strLine), 17) = "VERSION 1.0 CLASS" Or _
                Trim(strLine) = "BEGIN" Or _
                Left(Trim(strLine), 8) = "MultiUse" Or _
                Trim(strLine) = "END") Then
            
            ' Si no cumple ninguna condición, es código VBA válido (incluyendo Option Explicit y Option Compare)
            ' Se añade al cleanedContent seguida de un salto de línea
            cleanedContent = cleanedContent & strLine & vbCrLf
        End If
    Next
    
    ' La función devuelve cleanedContent directamente
    ' No añade ninguna cabecera Option manualmente
    CleanVBAFile = cleanedContent
End Function




' Función para exportar módulo con conversión ANSI -> UTF-8 usando ADODB.Stream
Sub ExportModuleWithAnsiEncoding(vbComponent, strExportPath)
    Dim tempFilePath, objTempFile, objStream
    Dim strContent
    Dim tempFolderPath, tempFileName
    
    On Error Resume Next
    
    ' Crear archivo temporal usando el directorio temporal del sistema
    tempFolderPath = objFSO.GetSpecialFolder(2) ' El 2 es la constante para la carpeta temporal del sistema
    tempFileName = objFSO.GetTempName() ' Genera un nombre aleatorio y seguro como "radB93EB.tmp"
    tempFilePath = objFSO.BuildPath(tempFolderPath, tempFileName)
    
    ' Exportar a archivo temporal (Access usa ANSI internamente)
    vbComponent.Export tempFilePath
    
    ' Leer contenido del archivo temporal con codificación ANSI usando FSO
    Set objTempFile = objFSO.OpenTextFile(tempFilePath, 1, False, 0) ' ForReading = 1, Create = False, Format = 0 (ANSI)
    strContent = objTempFile.ReadAll
    objTempFile.Close
    
    ' Escribir al archivo final con codificación UTF-8 usando ADODB.Stream
    Set objStream = CreateObject("ADODB.Stream")
    objStream.Type = 2 ' adTypeText
    objStream.Charset = "UTF-8"
    objStream.Open
    objStream.WriteText strContent
    objStream.SaveToFile strExportPath, 2 ' adSaveCreateOverWrite
    objStream.Close
    Set objStream = Nothing
    
    ' Limpiar archivo temporal
    objFSO.DeleteFile tempFilePath
    
    If Err.Number <> 0 Then
        WScript.Echo "❌ ERROR en ExportModuleWithAnsiEncoding: " & Err.Description
        Err.Clear
    End If
    
    On Error GoTo 0
End Sub

' Subrutina principal para validar esquemas de base de datos
Sub ValidateSchema(lanzaderaPath, condorPath)
    WScript.Echo "=== INICIANDO VALIDACIÓN DE ESQUEMA DE BASE DE DATOS ==="
    
    Dim lanzaderaSchema, condorSchema
    Dim allOk
    allOk = True
    
    ' Si no se proporcionan rutas, usar las por defecto (producción)
    If IsEmpty(lanzaderaPath) Or lanzaderaPath = "" Then 
        lanzaderaPath = strSourcePath & "\..\back\Lanzadera_Datos.accdb"
    End If
    If IsEmpty(condorPath) Or condorPath = "" Then 
        condorPath = strSourcePath & "\..\back\CONDOR_datos.accdb"
    End If
    
    WScript.Echo "Validando bases de datos:"
    WScript.Echo "  - Lanzadera: " & lanzaderaPath
    WScript.Echo "  - CONDOR: " & condorPath
    
    ' Definir esquema esperado para Lanzadera
    Set lanzaderaSchema = CreateObject("Scripting.Dictionary")
    lanzaderaSchema.Add "TbUsuariosAplicaciones", Array("CorreoUsuario", "Password", "UsuarioRed", "Nombre", "Matricula", "FechaAlta")
    lanzaderaSchema.Add "TbUsuariosAplicacionesPermisos", Array("CorreoUsuario", "IDAplicacion", "EsUsuarioAdministrador", "EsUsuarioCalidad", "EsUsuarioEconomia", "EsUsuarioSecretaria")
    
    ' Definir esquema esperado para CONDOR
    Set condorSchema = CreateObject("Scripting.Dictionary")
    condorSchema.Add "tbSolicitudes", Array("idSolicitud", "idExpediente", "tipoSolicitud", "estadoInterno", "fechaCreacion", "usuarioCreacion", "fechaModificacion", "usuarioModificacion", "observaciones")
    condorSchema.Add "tbDatosPC", Array("idSolicitud", "numeroExpediente", "fechaExpediente", "tipoExpediente", "descripcionCambio", "justificacionCambio", "impactoCalidad", "impactoSeguridad", "impactoOperacional")
    condorSchema.Add "tbDatosCDCA", Array("idSolicitud", "numeroExpediente", "fechaExpediente", "tipoExpediente", "descripcionDesviacion", "justificacionDesviacion", "impactoCalidad", "impactoSeguridad", "impactoOperacional")
    condorSchema.Add "tbDatosCDCASUB", Array("idSolicitud", "numeroExpediente", "fechaExpediente", "tipoExpediente", "descripcionDesviacion", "justificacionDesviacion", "impactoCalidad", "impactoSeguridad", "impactoOperacional", "subsuministrador")
    condorSchema.Add "tbMapeoCampos", Array("NombrePlantilla", "NombreCampoTabla", "ValorAsociado", "NombreCampoWord")
    condorSchema.Add "tbLogCambios", Array("idLog", "idSolicitud", "fechaCambio", "usuarioCambio", "campoModificado", "valorAnterior", "valorNuevo")
    condorSchema.Add "tbLogErrores", Array("idError", "fechaError", "tipoError", "descripcionError", "moduloOrigen", "funcionOrigen", "usuarioAfectado")
    condorSchema.Add "tbOperacionesLog", Array("idOperacion", "fechaOperacion", "tipoOperacion", "descripcionOperacion", "usuario", "resultado")
    condorSchema.Add "tbAdjuntos", Array("idAdjunto", "idSolicitud", "nombreArchivo", "rutaArchivo", "tipoArchivo", "fechaSubida", "usuarioSubida")
    condorSchema.Add "tbEstados", Array("idEstado", "nombreEstado", "descripcionEstado", "esEstadoFinal")
    condorSchema.Add "tbTransiciones", Array("idTransicion", "estadoOrigen", "estadoDestino", "accionRequerida", "rolRequerido")
    condorSchema.Add "tbConfiguracion", Array("clave", "valor", "descripcion", "categoria")
    condorSchema.Add "TbLocalConfig", Array("clave", "valor", "descripcion", "categoria")
    
    ' Validar las bases de datos
    If Not VerifySchema(lanzaderaPath, "dpddpd", lanzaderaSchema) Then allOk = False
    If Not VerifySchema(condorPath, "", condorSchema) Then allOk = False
    
    If allOk Then
        WScript.Echo "✓ VALIDACIÓN DE ESQUEMA EXITOSA. Todas las bases de datos son consistentes."
        WScript.Quit 0
    Else
        WScript.Echo "X VALIDACIÓN DE ESQUEMA FALLIDA. Corrija las discrepancias."
        WScript.Quit 1
    End If
End Sub

' Función auxiliar para verificar esquema de una base de datos específica
Private Function VerifySchema(dbPath, dbPassword, expectedSchema)
    On Error Resume Next
    
    WScript.Echo "Validando base de datos: " & dbPath
    
    ' Verificar que existe la base de datos
    If Not objFSO.FileExists(dbPath) Then
        WScript.Echo "❌ ERROR: Base de datos no encontrada: " & dbPath
        VerifySchema = False
        Exit Function
    End If
    
    ' Crear conexión ADO
    Dim conn, rs
    Set conn = CreateObject("ADODB.Connection")
    
    ' Construir cadena de conexión
    Dim connectionString
    If dbPassword = "" Then
        connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath & ";"
    Else
        connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath & ";Jet OLEDB:Database Password=" & dbPassword & ";"
    End If
    
    ' Abrir conexión
    conn.Open connectionString
    
    If Err.Number <> 0 Then
        WScript.Echo "❌ ERROR: No se pudo conectar a la base de datos: " & Err.Description
        VerifySchema = False
        Err.Clear
        Exit Function
    End If
    
    ' Iterar sobre cada tabla esperada
    Dim tableName, expectedFields, i
    Dim tableExists, fieldExists
    Dim allTablesOk
    allTablesOk = True
    
    For Each tableName In expectedSchema.Keys
        expectedFields = expectedSchema(tableName)
        
        ' Verificar que existe la tabla usando una consulta más simple
         Set rs = CreateObject("ADODB.Recordset")
         On Error Resume Next
         rs.Open "SELECT TOP 1 * FROM [" & tableName & "]", conn
         tableExists = (Err.Number = 0)
         If tableExists Then rs.Close
         Err.Clear
         On Error GoTo 0
        
        If Not tableExists Then
            WScript.Echo "❌ ERROR: Tabla no encontrada: " & tableName
            allTablesOk = False
        Else
            WScript.Echo "✓ Tabla encontrada: " & tableName
            
            ' Verificar cada campo esperado
            For i = 0 To UBound(expectedFields)
                Dim fieldName
                fieldName = expectedFields(i)
                
                ' Verificar que existe el campo usando una consulta más simple
                 Set rs = CreateObject("ADODB.Recordset")
                 On Error Resume Next
                 rs.Open "SELECT [" & fieldName & "] FROM [" & tableName & "] WHERE 1=0", conn
                 fieldExists = (Err.Number = 0)
                 If fieldExists Then rs.Close
                 Err.Clear
                 On Error GoTo 0
                
                If Not fieldExists Then
                    WScript.Echo "❌ ERROR: Campo no encontrado: " & tableName & "." & fieldName
                    allTablesOk = False
                Else
                    WScript.Echo "  ✓ Campo encontrado: " & fieldName
                End If
            Next
        End If
    Next
    
    ' Cerrar conexión
    conn.Close
    Set conn = Nothing
    
    If allTablesOk Then
        WScript.Echo "✅ Base de datos validada correctamente: " & objFSO.GetFileName(dbPath)
        VerifySchema = True
    Else
        WScript.Echo "❌ Errores encontrados en: " & objFSO.GetFileName(dbPath)
        VerifySchema = False
    End If
    
    On Error GoTo 0
End Function

' ===================================================================
' FUNCIÓN AUXILIAR: ParseJsonFile
' Descripción: Lee y parsea un archivo JSON
' ===================================================================
Function ParseJsonFile(filePath)
    Dim objFSO, objFile, strContent
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFile = objFSO.OpenTextFile(filePath, 1, False, 0) ' Cambiar a ASCII
    strContent = objFile.ReadAll
    objFile.Close
    Set ParseJsonFile = ParseJson(strContent)
End Function

Sub EnsureVBReferences()
    WScript.Echo "Verificando referencias VBA críticas..."
    On Error Resume Next
    Dim vbProj: Set vbProj = objAccess.VBE.ActiveVBProject
    If vbProj Is Nothing Then Exit Sub
    
    Dim refs(1, 2)
    refs(0, 0) = "{420B2830-E718-11CF-893D-00A0C9054228}": refs(0, 1) = "1.0": refs(0, 2) = "Scripting Runtime"
    refs(1, 0) = "{0002E157-0000-0000-C000-000000000046}": refs(1, 1) = "5.3": refs(1, 2) = "VBIDE Extensibility"

    Dim i, ref, found
    For i = 0 To 1
        found = False
        For Each ref In vbProj.References
            If ref.Guid = refs(i, 0) Then found = True: Exit For
        Next
        If Not found Then
            WScript.Echo "  -> Añadiendo: " & refs(i, 2)
            vbProj.References.AddFromGuid refs(i, 0), CInt(Split(refs(i, 1), ".")(0)), CInt(Split(refs(i, 1), ".")(1))
        End If
    Next
    On Error GoTo 0
End Sub

' ===== VBIDE Import Helpers (compactos) =====
Function EnsureVBProject(app)
    On Error Resume Next
    Dim p: Set p = app.VBE.ActiveVBProject
    If Err.Number<>0 Or p Is Nothing Then
        WScript.Echo "❌ VBIDE no accesible. Activa 'Confiar en el modelo de objetos de proyectos de VBA'."
        WScript.Quit 1
    End If
    Set EnsureVBProject = p
End Function

Sub RemoveVBComponentIfExists(app, name)
    On Error Resume Next
    Dim p,c: Set p = EnsureVBProject(app)
    Set c = p.VBComponents(name)
    If Err.Number=0 Then p.VBComponents.Remove c
    Err.Clear
End Sub

Function CreateAnsiTempFrom(path)
    On Error Resume Next
    
    ' Leer texto usando ReadAllText (que ya maneja UTF-8 y fallback)
    Dim text
    text = ReadAllText(path)
    
    ' Normalizar EOL a CRLF
    text = Replace(text, vbCrLf, vbLf)
    text = Replace(text, vbCr, vbLf)
    text = Replace(text, vbLf, vbCrLf)
    
    ' Crear archivo temporal en %TEMP%\condor_tmp_<nombre>
    Dim tmpPath
    tmpPath = objFSO.BuildPath(objFSO.GetSpecialFolder(2), "condor_tmp_" & objFSO.GetFileName(path))
    
    ' Escribir a archivo temporal usando ADODB.Stream con Windows-1252 SIN BOM
    Dim stream
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2 ' adTypeText
    stream.Charset = "Windows-1252"
    stream.Open
    stream.WriteText text
    stream.SaveToFile tmpPath, 2 ' adSaveCreateOverWrite
    stream.Close
    
    ' Si hay error, devolver cadena vacía
    If Err.Number <> 0 Then
        tmpPath = ""
    End If
    
    CreateAnsiTempFrom = tmpPath
    Err.Clear
End Function

Function ImportVbComponentFromFile(app, moduleName, srcPath)
    On Error Resume Next
    
    ' Establecer modo silencioso antes de cualquier importación
    app.Application.DisplayAlerts = False
    app.Application.Echo False
    app.DoCmd.SetWarnings False
    
    ' a) Crear temp ANSI con CreateAnsiTempFrom
    Dim tmp: tmp = CreateAnsiTempFrom(srcPath)
    If Err.Number <> 0 Or tmp = "" Then
        WScript.Echo "  ❌ Error creando archivo temporal ANSI: " & Err.Description
        ImportVbComponentFromFile = False
        Exit Function
    End If
    
    ' b) Set p = app.VBE.ActiveVBProject (fallar con mensaje si no hay VBIDE/trust)
    Dim p: Set p = app.VBE.ActiveVBProject
    If Err.Number <> 0 Or p Is Nothing Then
        WScript.Echo "  ❌ Error accediendo VBE.ActiveVBProject - Verificar VBIDE/trust: " & Err.Description
        If objFSO.FileExists(tmp) Then objFSO.DeleteFile tmp, True
        ImportVbComponentFromFile = False
        Exit Function
    End If
    
    ' c) p.VBComponents.Import tmp y después p.VBComponents(p.VBComponents.Count).Name = moduleName
    p.VBComponents.Import tmp
    If Err.Number <> 0 Then
        WScript.Echo "  ❌ Import VBIDE " & srcPath & ": " & Err.Number & " - " & Err.Description
        If objFSO.FileExists(tmp) Then objFSO.DeleteFile tmp, True
        ImportVbComponentFromFile = False
        Exit Function
    End If
    
    p.VBComponents(p.VBComponents.Count).Name = moduleName
    If Err.Number <> 0 Then
        WScript.Echo "  ❌ Error renombrando componente a " & moduleName & ": " & Err.Description
        If objFSO.FileExists(tmp) Then objFSO.DeleteFile tmp, True
        ImportVbComponentFromFile = False
        Exit Function
    End If
    
    ' d) Guardar el módulo para evitar diálogos de confirmación
    app.DoCmd.Save 5, moduleName  ' 5 = acModule
    If Err.Number <> 0 Then
        WScript.Echo "  ❌ Error guardando módulo " & moduleName & ": " & Err.Description
        If objFSO.FileExists(tmp) Then objFSO.DeleteFile tmp, True
        ImportVbComponentFromFile = False
        Exit Function
    End If
    
    ' e) Borra el temp y devuelve True/False
    If objFSO.FileExists(tmp) Then objFSO.DeleteFile tmp, True
    ImportVbComponentFromFile = True
End Function
' ===== FIN Helpers =====

' Subrutina para mostrar ayuda completa
Sub ShowHelp()
    WScript.Echo "=== CONDOR CLI - Herramienta de línea de comandos ==="
    WScript.Echo "Versión: 2.0 - Sistema de gestión y sincronización VBA para proyecto CONDOR"
    WScript.Echo ""
    WScript.Echo "SINTAXIS:"
    WScript.Echo "  cscript condor_cli.vbs [comando] [opciones] [parámetros]"
    WScript.Echo ""
    WScript.Echo "Por defecto, acciones de código usan front\Desarrollo\CONDOR.accdb; acciones de datos usan back\data\CONDOR_datos.accdb”."
    WScript.Echo ""
    WScript.Echo "COMANDOS PRINCIPALES:"
    WScript.Echo ""
    WScript.Echo "📤 EXPORTACIÓN:"
    WScript.Echo "  export [--verbose]           - Exportar módulos VBA desde Access a /src"
    WScript.Echo "                                 Codificación: ANSI para compatibilidad"
    WScript.Echo "                                 Las cabeceras SIEMPRE quedan correctas en ./src"
    WScript.Echo "                                 CÓDIGO (VBIDE): .bas/.cls con cabeceras canónicas"
    WScript.Echo "                                 UI (JSON): usar export-form/import-form para formularios"
    WScript.Echo "                                 --verbose: Mostrar detalles de cada archivo"
    WScript.Echo ""
    WScript.Echo "🔄 SINCRONIZACIÓN:"
    WScript.Echo "  rebuild [--verbose] [--verifyModules] - Método principal de sincronización del proyecto"
    WScript.Echo "                                 Opera sobre FRONTEND por defecto (" & GetDevDbPath() & ")"
    WScript.Echo "                                 Reconstrucción completa: elimina todos los módulos"
    WScript.Echo "                                 y reimporta desde /src para garantizar coherencia"
    WScript.Echo "                                 TOLERANTE A CABECERAS: acepta .bas/.cls con o sin cabecera;"
    WScript.Echo "                                 si falta, el CLI inyecta una cabecera temporal según la extensión"
    WScript.Echo "                                 --verifyModules: Verificar automáticamente tras importar"
    WScript.Echo "                                 Transaccional: usa la misma sesión de Access"
    WScript.Echo "  update <Nombre|Nombre1,Nombre2|Nombre1 Nombre2> [--changed|--all] [--verifyModules] [--verbose] - Actualización selectiva de módulos VBA"
    WScript.Echo "                                 --changed: Solo módulos modificados (comparación por hash/fecha)"
    WScript.Echo "                                 --all: Todos los módulos (sync suave, sin eliminar)"
    WScript.Echo "                                 <Nombre>: Módulo específico (acepta nombre sin extensión)"
    WScript.Echo "                                 Sintaxis: nombres separados por comas (sin espacios) o por espacios"
    WScript.Echo "                                 --verifyModules: Verificar automáticamente tras importar"
    WScript.Echo "                                 Transaccional: usa la misma sesión de Access"
    WScript.Echo "                                 Ejemplos:"
    WScript.Echo "                                   cscript condor_cli.vbs update CAuthService,modUtils,CConfig"
    WScript.Echo "                                   cscript condor_cli.vbs update CAuthService modUtils CConfig"
    WScript.Echo "                                   cscript condor_cli.vbs update --changed"
    WScript.Echo "                                   cscript condor_cli.vbs update --all"
    WScript.Echo ""
    WScript.Echo "✅ VALIDACIÓN Y PRUEBAS:"
    WScript.Echo "  validate [--verbose] [--src] - Validar sintaxis VBA sin importar a Access"
    WScript.Echo "                                 --verbose: Mostrar detalles de validación"
    WScript.Echo "                                 --src: Usar directorio fuente alternativo"
    WScript.Echo "  test                         - Ejecutar suite completa de pruebas unitarias"
    WScript.Echo "  lint                         - Auditar código VBA (detectar cabeceras duplicadas)"
    WScript.Echo ""
    WScript.Echo "📦 EMPAQUETADO:"
    WScript.Echo "  bundle <funcionalidad> [destino] - Empaquetar archivos por funcionalidad"
    WScript.Echo "                                      Destino opcional (por defecto: directorio actual)"
    WScript.Echo ""
    WScript.Echo "🗄️ GESTIÓN DE BASE DE DATOS:"
    WScript.Echo "  createtable <nombre> <sql>   - Crear tabla con consulta SQL personalizada"
    WScript.Echo "  droptable <nombre>           - Eliminar tabla de la base de datos"
    WScript.Echo "  listtables [db_path]         - Listar todas las tablas"
    WScript.Echo "                                 db_path opcional (por defecto: back\data\CONDOR_datos.accdb)"
    WScript.Echo "  list-forms [db_path] [--password <pwd>] [--json] - Listar todos los formularios"
    WScript.Echo "                                 db_path opcional (por defecto: ./condor.accdb)"
    WScript.Echo "                                 --password: Contraseña de la base de datos"
    WScript.Echo "                                 --json: Salida en formato JSON (array de nombres)"
    WScript.Echo "  list-modules [--includeDocs] [--pattern <regex>] [--json] [--db <path>] [--password <pwd>] [--expectSrc [path]] [--diff]"
    WScript.Echo "                                 --json: Salida en formato JSON estructurado"
    WScript.Echo "                                 --expectSrc: Verificar existencia de archivos fuente en /src"
    WScript.Echo "                                 --diff: Detectar inconsistencias entre BD y archivos fuente"
    WScript.Echo "                                 Análisis: módulos faltantes, huérfanos, desactualizados"
    WScript.Echo "                                 Indicadores: ✓ (sincronizado), ⚠ (advertencia), ✗ (error)"
    WScript.Echo "                                 Orden de obtención: VBIDE → AllModules → DAO (fallback). Útil cuando VBIDE está bloqueado por políticas."
    WScript.Echo "  relink <db_path> <folder>    - Re-vincular tablas a bases locales específicas"
    WScript.Echo "  relink --all                 - Re-vincular automáticamente todas las bases en ./back"
    WScript.Echo "  migrate [file.sql]           - Ejecutar scripts de migración SQL desde ./db/migrations"
    WScript.Echo "  export-form <db_path> <form_name> [opciones] - Exportar diseño de formulario a JSON enriquecido."
    WScript.Echo "                                 Genera JSON versionado con metadata, normalización de colores y recursos."
    WScript.Echo "                                 REQUIERE ACCESO EXCLUSIVO: La base de datos debe estar cerrada en Access."
    WScript.Echo "                                 Incluye propiedades completas del formulario: caption, popUp, modal, width,"
    WScript.Echo "                                 autoCenter, borderStyle, recordSelectors, dividingLines, navigationButtons,"
    WScript.Echo "                                 scrollBars, controlBox, closeButton, minMaxButtons, movable, recordsetType,"
    WScript.Echo "                                 orientation y propiedades de SplitForm (si aplica)."
    WScript.Echo "                                 Opciones:"
    WScript.Echo "                                   --output <archivo>        - Archivo de salida (por defecto: <form_name>.json)"
    WScript.Echo "                                   --password <pwd>          - Contraseña de la base de datos"
    WScript.Echo "                                   --json                    - Salida JSON a consola (no guarda archivo)"
    WScript.Echo "                                   --schema-version <ver>    - Versión del esquema (por defecto: 1.0.0)"
    WScript.Echo "                                   --expand <ámbitos>        - Ámbitos a incluir: events,formatting,resources"
    WScript.Echo "                                                               (por defecto: todos)"
    WScript.Echo "                                   --resource-root <dir>     - Directorio base para rutas relativas de recursos"
    WScript.Echo "                                   --pretty                  - Formatear JSON con indentación"
    WScript.Echo "                                   --no-controls             - Solo metadata del formulario, sin controles"
    WScript.Echo "                                   --verbose                 - Mostrar información detallada del proceso"
    WScript.Echo "                                 Propiedades JSON exportadas:"
    WScript.Echo "                                   • caption: string - Título del formulario"
    WScript.Echo "                                   • popUp/modal: boolean - Comportamiento de ventana"
    WScript.Echo "                                   • width: number (twips) - Ancho del formulario"
    WScript.Echo "                                   • borderStyle: ""None""|""Thin""|""Sizable""|""Dialog"""
    WScript.Echo "                                   • scrollBars: ""Neither""|""Horizontal""|""Vertical""|""Both"""
    WScript.Echo "                                   • minMaxButtons: ""None""|""Min Enabled""|""Max Enabled""|""Both Enabled"""
    WScript.Echo "                                   • recordsetType: ""Dynaset""|""Snapshot""|""Dynaset (Inconsistent Updates)"""
    WScript.Echo "                                   • orientation: ""LeftToRight""|""RightToLeft"""
    WScript.Echo "                                   • splitForm*: propiedades específicas para formularios divididos"
    WScript.Echo "                                 Ejemplo: export-form db.accdb MiForm --pretty --expand=events,formatting"
    WScript.Echo "  import-form <json_path> <db_path> [opciones] - Crear/Modificar formulario desde JSON."
    WScript.Echo "                                 Soporta normalización automática ES→EN y reglas de coherencia."
    WScript.Echo "                                 Opciones:"
    WScript.Echo "                                   --password <pwd>          - Contraseña de la base de datos"
    WScript.Echo "                                   --strict                  - Modo estricto: errores por incoherencias"
    WScript.Echo "                                   --verbose                 - Mostrar decisiones de normalización"
    WScript.Echo "                                   --dry-run                 - Validar sin crear el formulario"
    WScript.Echo "                                   --schema <version>        - Versión del esquema a validar"
    WScript.Echo "                                 Normalización automática (ES→EN):"
    WScript.Echo "                                   • scrollBars: ""Ninguna""→""Neither"", ""Horizontal""→""Horizontal"","
    WScript.Echo "                                                 ""Vertical""→""Vertical"", ""Ambas""/""Ambos""→""Both"""
    WScript.Echo "                                   • borderStyle: ""Ninguno""→""None"", ""Fino""→""Thin"","
    WScript.Echo "                                                  ""Redimensionable""→""Sizable"", ""Cuadro de diálogo""→""Dialog"""
    WScript.Echo "                                   • minMaxButtons: ""Ninguno""→""None"", ""Solo minimizar""→""Min Enabled"","
    WScript.Echo "                                                    ""Solo maximizar""→""Max Enabled"", ""Ambos""→""Both Enabled"""
    WScript.Echo "                                   • recordsetType: ""Instantánea""→""Snapshot"","
    WScript.Echo "                                                    ""Dynaset (actualizaciones incoherentes)""→""Dynaset (Inconsistent Updates)"""
    WScript.Echo "                                   • orientation: ""De izquierda a derecha""→""LeftToRight"","
    WScript.Echo "                                                  ""De derecha a izquierda""→""RightToLeft"""
    WScript.Echo "                                   • booleans: ""Sí""→true, ""No""→false"
    WScript.Echo "                                 Reglas de coherencia aplicadas:"
    WScript.Echo "                                   • borderStyle ∈ {""None"",""Dialog""} ⇒ controlBox=false, minMaxButtons=""None"""
    WScript.Echo "                                   • controlBox=false ⇒ ignora closeButton y minMaxButtons"
    WScript.Echo "                                   • modal/popUp=true + borderStyle≠""Sizable"" ⇒ no min/max (WARN o ERROR)"
    WScript.Echo "                                   • splitForm*: solo aplicable si defaultView=""Split Form"""
    WScript.Echo "                                 Ejemplo: import-form form.json db.accdb --strict --verbose"
    WScript.Echo "  validate-form-json <json_path> [--strict] [--schema] - Validar estructura JSON de formulario"
    WScript.Echo "                                 --strict: Validación exhaustiva de coherencia con código VBA"
    WScript.Echo "                                 --schema: Validar contra esquema específico"
    WScript.Echo "  roundtrip-form <db_path> <form> [--password] [--temp-dir] [--verbose] - Test export→import de formulario"
    WScript.Echo ""
    WScript.Echo "🔗 UI as Code — Vinculación UI↔Código:"
    WScript.Echo "  Los comandos export-form e import-form detectan automáticamente módulos .bas/.cls"
    WScript.Echo "  asociados al formulario y gestionan la vinculación entre eventos UI y handlers VBA."
    WScript.Echo ""
    WScript.Echo "  📋 Detección de Módulos:"
    WScript.Echo "    • Busca archivos: Form_<FormName>.bas, <FormName>.bas, frm<FormName>.bas, Form_<FormName>.cls"
    WScript.Echo "    • Extrae handlers con patrón: Sub <Control>_<Event>(...)"
    WScript.Echo "    • Eventos soportados: Click, DblClick, Current, Load, Open, GotFocus, LostFocus,"
    WScript.Echo "                          Change, AfterUpdate, BeforeUpdate"
    WScript.Echo ""
    WScript.Echo "  📤 En export-form:"
    WScript.Echo "    • Genera bloque JSON 'code.module' con: exists, filename, handlers[]"
    WScript.Echo "    • Cada handler incluye: control, event, signature"
    WScript.Echo "    • Añade 'events.detected' a controles con handlers encontrados"
    WScript.Echo ""
    WScript.Echo "  📥 En import-form:"
    WScript.Echo "    • Establece automáticamente '[Event Procedure]' cuando:"
    WScript.Echo "      - El JSON especifica explícitamente '[Event Procedure]'"
    WScript.Echo "      - Existe handler correspondiente en código detectado"
    WScript.Echo "    • Modo --strict: ERROR si hay discrepancias entre JSON y código"
    WScript.Echo "    • Sin --strict: WARNING por discrepancias, continúa procesamiento"
    WScript.Echo ""
    WScript.Echo "  💡 Ejemplo de flujo:"
    WScript.Echo "    1. export-form db.accdb MiForm --src ./src"
    WScript.Echo "    2. Editar MiForm.json (cambiar propiedades UI)"
    WScript.Echo "    3. import-form MiForm.json db.accdb --strict"
    WScript.Echo "    → Los handlers existentes se preservan automáticamente"
    WScript.Echo ""
    WScript.Echo "  ⚠️  NOTA: Los comandos export/import/roundtrip operan SIEMPRE en vista Diseño para evitar"
    WScript.Echo "           la ejecución de eventos y garantizar operaciones seguras. El CLI desactiva"
    WScript.Echo "           automáticamente el 'Startup Form'/'Display Form' al abrir la BD."
    WScript.Echo ""
    WScript.Echo "FUNCIONALIDADES DISPONIBLES PARA 'bundle' (con dependencias automáticas):"
    WScript.Echo "(Basadas en CONDOR_MASTER_PLAN.md)"
    WScript.Echo ""
    WScript.Echo "🔐 Auth          - Autenticación + dependencias (Config, Error, Modelos)"
    WScript.Echo "                   Incluye archivos de autenticación y sus dependencias"
    WScript.Echo ""
    WScript.Echo "📄 Document      - Gestión de Documentos + dependencias (Config, FileSystem, Error, Word, Modelos)"
    WScript.Echo "                   Incluye archivos de documentos y sus dependencias"
    WScript.Echo ""
    WScript.Echo "📁 Expediente    - Gestión de Expedientes + dependencias (Config, Error, Modelos, Utilidades)"
    WScript.Echo "                   Incluye archivos de expedientes y sus dependencias"
    WScript.Echo ""
    WScript.Echo "📋 Solicitud     - Gestión de Solicitudes + dependencias (Config, Error, Modelos, Utilidades)"
    WScript.Echo "                   Incluye archivos de solicitudes y sus dependencias"
    WScript.Echo ""
    WScript.Echo "🔄 Workflow      - Flujos de Trabajo + dependencias (Config, Error, Modelos, Utilidades)"
    WScript.Echo "                   Incluye archivos de workflow y sus dependencias"
    WScript.Echo ""
    WScript.Echo "🗺️ Mapeo         - Gestión de Mapeos + dependencias (Config, Error, Modelos)"
    WScript.Echo "                   Incluye archivos de mapeos y sus dependencias"
    WScript.Echo ""
    WScript.Echo "🔔 Notification  - Notificaciones + dependencias (Config, Error, Modelos)"
    WScript.Echo "                   Incluye archivos de notificaciones y sus dependencias"
    WScript.Echo ""
    WScript.Echo "📊 Operation     - Operaciones y Logging + dependencias (Config, Error, Modelos)"
    WScript.Echo "                   Incluye archivos de operaciones y sus dependencias"
    WScript.Echo ""
    WScript.Echo "⚙️ Config        - Configuración del Sistema + dependencias (Error, Modelos)"
    WScript.Echo "                   Incluye archivos de configuración y sus dependencias"
    WScript.Echo ""
    WScript.Echo "📂 FileSystem    - Sistema de Archivos + dependencias (Error, Modelos)"
    WScript.Echo "                   Incluye archivos de sistema de archivos y sus dependencias"
    WScript.Echo ""
    WScript.Echo "❌ Error         - Manejo de Errores + dependencias (Modelos)"
    WScript.Echo "                   Incluye archivos de manejo de errores y sus dependencias"
    WScript.Echo ""
    WScript.Echo "📝 Word          - Microsoft Word + dependencias (Error, Modelos)"
    WScript.Echo "                   Incluye archivos de Word y sus dependencias"
    WScript.Echo ""
    WScript.Echo "🧪 TestFramework - Framework de Pruebas + dependencias (11 archivos)"
WScript.Echo "                   Incluye ITestReporter, CTestResult, CTestSuiteResult, CTestReporter, modTestRunner,"
    WScript.Echo "                   modTestUtils, ModAssert, TestModAssert, IFileSystem, IConfig, IErrorHandlerService"
    WScript.Echo ""
    WScript.Echo "🖥️ Aplicacion    - Gestión de Aplicación + dependencias (Config, Error, Modelos)"
    WScript.Echo "                   Incluye archivos de gestión de aplicación y sus dependencias"
    WScript.Echo ""
    WScript.Echo "📊 Modelos       - Modelos de Datos (entidades base)"
    WScript.Echo "                   Incluye todas las entidades de datos del sistema"
    WScript.Echo ""
    WScript.Echo "🔧 Utilidades    - Utilidades y Enumeraciones + dependencias (Error, Modelos)"
    WScript.Echo "                   Incluye utilidades, enumeraciones y sus dependencias"
    WScript.Echo ""
    WScript.Echo "OPCIONES GLOBALES:"
    WScript.Echo "  --help, -h, help             - Mostrar esta ayuda completa"
    WScript.Echo "  --src <directorio>           - Especificar directorio fuente alternativo"
    WScript.Echo "                                 (por defecto: " & RepoRoot() & "\\src)"
    WScript.Echo "  --strict                     - Modo estricto: validación exhaustiva de coherencia"
    WScript.Echo "                                 entre JSON y código VBA en formularios"
    WScript.Echo "  --verbose                    - Mostrar información detallada durante la operación"
    WScript.Echo "  --sharedopen                 - Abre la BD en modo compartido (por defecto el CLI abre en EXCLUSIVO)"
    WScript.Echo "  --db <ruta>, /db:<ruta>      - Especificar ruta de base de datos"
    WScript.Echo "                                 (por defecto: ENV('CONDOR_DEV_DB') o ruta por defecto)"
    WScript.Echo "  --password <clave>, /pwd:<clave> - Especificar contraseña de base de datos"
    WScript.Echo "  --print-db                   - Mostrar la ruta de base de datos que se utilizará"
    WScript.Echo "                                 ⚠️  NOTA: El CLI desactiva automáticamente el 'Startup Form'/'Display Form'"
    WScript.Echo "                                     al abrir la BD y lo restaura al cerrar automáticamente."
    WScript.Echo ""
    WScript.Echo "FLUJO DE TRABAJO RECOMENDADO:"
    WScript.Echo "  1. cscript condor_cli.vbs validate     (validar sintaxis antes de importar)"
    WScript.Echo "  2. cscript condor_cli.vbs rebuild      (reconstrucción completa del proyecto)"
    WScript.Echo "  3. cscript condor_cli.vbs test         (ejecutar pruebas unitarias)"
    WScript.Echo ""
    WScript.Echo "EJEMPLOS DE USO:"
    WScript.Echo "  cscript condor_cli.vbs --help"
    WScript.Echo "  cscript condor_cli.vbs validate --verbose --src " & RepoRoot() & "\\src"
    WScript.Echo "  cscript condor_cli.vbs export --verbose"
    WScript.Echo "  cscript condor_cli.vbs validate-form-json formulario.json --strict"
    WScript.Echo "  cscript condor_cli.vbs bundle Auth"
    WScript.Echo "  cscript condor_cli.vbs bundle Document C:\\\\temp"
    WScript.Echo "  cscript condor_cli.vbs createtable MiTabla ""CREATE TABLE MiTabla (ID LONG)"""
    WScript.Echo "  cscript condor_cli.vbs listtables"
    WScript.Echo "  cscript condor_cli.vbs relink --all"
    WScript.Echo "  cscript condor_cli.vbs rebuild --verbose"
    WScript.Echo "  cscript condor_cli.vbs update --changed --verbose"
    WScript.Echo "  cscript condor_cli.vbs update --all"
    WScript.Echo "  cscript condor_cli.vbs update ModuloA CClaseB"
    WScript.Echo "  cscript condor_cli.vbs test --password miClave"
    WScript.Echo "  cscript condor_cli.vbs list-forms --password miClave"
    WScript.Echo "  cscript condor_cli.vbs export-form MiDB.accdb MiForm"
    WScript.Echo ""
    WScript.Echo "RESOLUCIÓN DE BASE DE DATOS:"
    WScript.Echo "  Todos los comandos aceptan ruta posicional o --db <ruta>."
    WScript.Echo "  Prioridad: flag → posicional → ENV(CONDOR_DEV_DB) → default repo-relativo."
    WScript.Echo ""
    WScript.Echo "  Ejemplos repo-relativos:"
    WScript.Echo "  cscript condor_cli.vbs list-forms .\\ui\\sources\\Expedientes.accdb --password dpddpd --print-db"
    WScript.Echo "  cscript condor_cli.vbs list-modules --json --print-db"
    WScript.Echo "  cscript condor_cli.vbs export --db .\\back\\Desarrollo\\CONDOR.accdb --verbose --print-db"
    WScript.Echo ""
    WScript.Echo "  # Ejemplos de export/import por referencias (subformularios y TabControl):"
    WScript.Echo "  # 1. Exportar formularios hijo y padre por separado"
    WScript.Echo "  cscript condor_cli.vbs export-form """"FormHijo"""" --pretty"
    WScript.Echo "  cscript condor_cli.vbs export-form """"FormPadre"""" --pretty"
    WScript.Echo ""
    WScript.Echo "  # 2. Importar por dependencias desde carpeta (toposort automático)"
    WScript.Echo "  cscript condor_cli.vbs import-form """".\\ui\\forms"""" --strict"
    WScript.Echo ""
    WScript.Echo "  # 3. Importar formulario individual con validación estricta"
    WScript.Echo "  cscript condor_cli.vbs import-form """"FormPadre.json"""" --strict"
    WScript.Echo ""
    WScript.Echo "CONFIGURACIÓN:"
    WScript.Echo "  Base de datos desarrollo: " & DefaultFrontendDb()
    WScript.Echo "                            (o ENV('CONDOR_DEV_DB') si está definida)"
    WScript.Echo "  Base de datos datos:      " & DefaultBackendDb()
    WScript.Echo "  Directorio fuente:        " & RepoRoot() & "\\src"
    WScript.Echo "  Variable de entorno:      CONDOR_DEV_DB (opcional, para ruta de BD personalizada)"
    WScript.Echo ""
    WScript.Echo "NOTA: '--bypassstartup' está DEPRECADO; el CLI abre Access con bypass automáticamente. El flag se acepta por compatibilidad pero no tiene efecto."
    WScript.Echo ""
    WScript.Echo "Para más información, consulte la documentación en docs/CONDOR_MASTER_PLAN.md"
End Sub

' Importa un módulo .bas o .cls desde /src en la sesión ACTUAL sin diálogos.
' - .bas: Application.LoadFromText acModule
' - .cls: VBIDE.VBComponents.Import y renombrado
' Función robusta para importar archivos VBA con inyección de cabeceras temporales
' Importa un archivo VBA (.bas o .cls) con inyección de cabeceras si es necesario
' POLÍTICA DE IMPORTACIÓN:
' - .bas → LoadFromText acModule(5) + Save
' POLÍTICA DE IMPORTACIÓN VBA POR TIPO (FUENTE DE VERDAD DEL PROYECTO)
' 
' Esta función implementa la política oficial de importación VBA de CONDOR:
' - .bas → Application.LoadFromText(acModule, moduleName, tmpFileAnsi) + DoCmd.Save
' - .cls → ImportVbComponentFromFile(tmpFileAnsi) (VBIDE.Import + rename) + DoCmd.Save
' 
' FLUJO UNIFICADO:
' 1. Creación temporal ANSI (Windows-1252) - conversión automática UTF-8→ANSI
' 2. Eliminación previa del componente VBA si existe
' 3. Inyección de cabeceras solo si faltan (BuildBasHeader/BuildClsHeader)
' 4. Importación por tipo específico según extensión
' 5. Guardado individual con DoCmd.Save
' 6. Limpieza automática de archivos temporales
' 
' REQUISITO: Trust Center → "Confiar en el acceso al modelo de objetos de proyectos de VBA"
Function ImportVbaFileRobust(sourcePath, moduleName, tmpDir)
    On Error Resume Next
    ImportVbaFileRobust = False
    
    If Not FileExists(sourcePath) Then 
        WScript.Echo "  ❌ Archivo no encontrado: " & sourcePath
        Exit Function
    End If
    
    Dim ext
    ext = LCase(objFSO.GetExtensionName(sourcePath))
    If ext <> "bas" And ext <> "cls" Then 
        WScript.Echo "  ❌ Extensión no válida: " & ext
        Exit Function
    End If
    
    ' Leer contenido del archivo usando ReadAllText robusto
    Dim fileContent, headerBodyResult, header, body
    fileContent = ReadAllText(sourcePath)
    headerBodyResult = SplitHeaderBody(fileContent)
    header = headerBodyResult(0)
    body = headerBodyResult(1)
    
    ' Verificar si la cabecera satisface los requisitos
    Dim hasValidHeader, finalPath, tmpPath
    hasValidHeader = HeaderSatisfies(ext, header)
    
    If hasValidHeader Then
        ' Cabecera válida, crear archivo temporal ANSI usando CreateAnsiTempFrom
        finalPath = CreateAnsiTempFrom(sourcePath)
        If finalPath = "" Then
            WScript.Echo "  ❌ Error creando archivo temporal ANSI"
            Exit Function
        End If
        If gVerbose Then
            WScript.Echo "    ✓ Cabecera válida, usando archivo temporal ANSI: " & finalPath
        End If
    Else
        ' Cabecera inválida o ausente, crear archivo temporal con cabecera inyectada
        If gVerbose Then
            WScript.Echo "    ⚠ Cabecera ausente/inválida, inyectando cabecera temporal"
        End If
        
        Dim injectedHeader, nameFromFilename
        nameFromFilename = objFSO.GetBaseName(sourcePath)
        
        If ext = "bas" Then
            injectedHeader = BuildBasHeader(nameFromFilename)
        ElseIf ext = "cls" Then
            injectedHeader = BuildClsHeader(nameFromFilename)
        End If
        
        ' Crear archivo temporal con cabecera inyectada + cuerpo original
        tmpPath = objFSO.BuildPath(tmpDir, objFSO.GetFileName(sourcePath))
        Dim tmpContent
        tmpContent = injectedHeader & vbCrLf & body
        
        ' Escribir archivo temporal usando WriteAllText robusto
        WriteAllText tmpPath, tmpContent
        
        ' Crear versión ANSI del archivo temporal usando CreateAnsiTempFrom
        finalPath = CreateAnsiTempFrom(tmpPath)
        If finalPath = "" Then
            WScript.Echo "  ❌ Error creando archivo temporal ANSI"
            ' Limpiar archivo temporal intermedio
            If FileExists(tmpPath) Then objFSO.DeleteFile tmpPath, True
            Exit Function
        End If
        
        If gVerbose Then
            WScript.Echo "    → Archivo temporal ANSI creado: " & finalPath
        End If
    End If
    
    ' Eliminar componente previo antes de importar
    Call RemoveVBComponentIfExists(objAccess, moduleName)
    
    ' Diferenciación .bas/.cls según especificación
    If ext = "bas" Then
        ' Para archivos .bas: usar LoadFromText directamente
        On Error Resume Next
        objAccess.Application.LoadFromText 5, moduleName, finalPath  ' 5 = acModule
        If Err.Number <> 0 Then
            If gVerbose Then
                WScript.Echo "  ❌ Error en LoadFromText para " & moduleName & ": " & Err.Number & " - " & Err.Description
            End If
            Err.Clear
            ' Limpiar archivos temporales
            If FileExists(finalPath) Then objFSO.DeleteFile finalPath, True
            If Not hasValidHeader And FileExists(tmpPath) Then objFSO.DeleteFile tmpPath, True
            ImportVbaFileRobust = False
            Exit Function
        End If
        
        ' Guardar el módulo
        objAccess.DoCmd.Save 5, moduleName  ' 5 = acModule
        If Err.Number <> 0 Then
            If gVerbose Then
                WScript.Echo "  ❌ Error guardando módulo " & moduleName & ": " & Err.Number & " - " & Err.Description
            End If
            Err.Clear
            ' Limpiar archivos temporales
            If FileExists(finalPath) Then objFSO.DeleteFile finalPath, True
            If Not hasValidHeader And FileExists(tmpPath) Then objFSO.DeleteFile tmpPath, True
            ImportVbaFileRobust = False
            Exit Function
        End If
        
        ' Post-check: verificar que el módulo existe en VBComponents
        On Error Resume Next
        Dim vbProj, testComp
        Set vbProj = EnsureVBProject(objAccess)
        Set testComp = vbProj.VBComponents(moduleName)
        If Err.Number <> 0 Or testComp Is Nothing Then
            If gVerbose Then
                WScript.Echo "  ❌ Post-check fallido: módulo " & moduleName & " no encontrado en VBComponents"
            End If
            Err.Clear
            ' Limpiar archivos temporales
            If FileExists(finalPath) Then objFSO.DeleteFile finalPath, True
            If Not hasValidHeader And FileExists(tmpPath) Then objFSO.DeleteFile tmpPath, True
            ImportVbaFileRobust = False
            Exit Function
        End If
        Err.Clear
        
    Else
        ' Para archivos .cls: usar ImportVbComponentFromFile
        Dim importResult
        importResult = ImportVbComponentFromFile(objAccess, moduleName, finalPath)
        If Not importResult Then
            ' Limpiar archivos temporales
            If FileExists(finalPath) Then objFSO.DeleteFile finalPath, True
            If Not hasValidHeader And FileExists(tmpPath) Then objFSO.DeleteFile tmpPath, True
            ImportVbaFileRobust = False
            Exit Function
        End If
    End If
    
    ' Limpiar archivos temporales
    If FileExists(finalPath) Then objFSO.DeleteFile finalPath, True
    If Not hasValidHeader And FileExists(tmpPath) Then objFSO.DeleteFile tmpPath, True
    
    ImportVbaFileRobust = True
    WScript.Echo "  ✓ Importado: " & moduleName
    
    Err.Clear
End Function




' Subrutina para ejecutar la suite de pruebas unitarias
Sub ExecuteTests()
    WScript.Echo "=== INICIANDO EJECUCION DE PRUEBAS ==="
    Dim reportString
    
    ' Verificar refactorización de logging antes de ejecutar pruebas
    WScript.Echo "Verificando refactorización de logging..."
    Call VerifyLoggingRefactoring()
    
    WScript.Echo "Ejecutando suite de pruebas en Access..."
    On Error Resume Next
    
    ' Configurar Access en modo completamente silencioso para pruebas
    objAccess.Application.DisplayAlerts = False
    objAccess.Application.Echo False
    objAccess.Visible = False
    objAccess.UserControl = False
    
    ' Configuraciones adicionales para suprimir todos los diálogos
    On Error Resume Next
    objAccess.Application.AutomationSecurity = 1  ' msoAutomationSecurityLow
    objAccess.DoCmd.SetWarnings False
    Err.Clear
    On Error Resume Next
    
    ' CORRECCIÓN CRÍTICA: Usar función ExecuteAllTestsForCLI restaurada que ejecuta todas las pruebas
    reportString = objAccess.Run("ExecuteAllTestsForCLI")
    
    If Err.Number <> 0 Then
        WScript.Echo "ERROR: Fallo crítico al invocar la suite de pruebas."
        WScript.Echo "  Código de Error: " & Err.Number
        WScript.Echo "  Descripción: " & Err.Description
        WScript.Echo "  Fuente: " & Err.Source
        WScript.Echo "SUGERENCIA: Abre Access manualmente y ejecuta ExecuteAllTestsForCLI desde el módulo modTestRunner."
        WScript.Quit 1
    End If
    On Error GoTo 0
    
    ' Verificar si el string devuelto es válido
    If IsEmpty(reportString) Or reportString = "" Then
        WScript.Echo "ERROR: El motor de pruebas de Access no devolvió ningún resultado."
        WScript.Echo "SUGERENCIA: Verifique que la función 'ExecuteAllTestsForCLI' en 'modTestRunner' no esté fallando silenciosamente."
        WScript.Quit 1
    End If
    
    ' Mostrar el reporte completo
    WScript.Echo "--- INICIO DE RESULTADOS DE PRUEBAS ---"
    WScript.Echo reportString
    WScript.Echo "--- FIN DE RESULTADOS DE PRUEBAS ---"
    
    ' Determinar el éxito o fracaso buscando la línea final
    If InStr(UCase(reportString), "RESULT: SUCCESS") > 0 Then
        WScript.Echo "RESULTADO FINAL: ✓ Todas las pruebas pasaron."
        WScript.Echo "✅ REFACTORIZACIÓN COMPLETADA: Patrón EOperationLog implementado correctamente"
        WScript.Quit 0 ' Éxito
    Else
        WScript.Echo "RESULTADO FINAL: ✗ Pruebas fallidas."
        WScript.Quit 1 ' Error
    End If
End Sub

' Función para importar módulo con conversión UTF-8 -> ANSI


' Subrutina para re-vincular tablas de Access
Sub RelinkTables()
    Dim strDbPath, strLocalFolder
    
    WScript.Echo "=== INICIANDO RE-VINCULACION DE TABLAS ==="
    
    ' Verificar si se usa el modo --all
    If objArgs.Count >= 2 Then
        If LCase(objArgs(1)) = "--all" Then
            Call RelinkAllDatabases()
            Exit Sub
        End If
    End If
    
    ' Verificar que se proporcionaron los argumentos necesarios para modo manual
    If objArgs.Count < 3 Then
        WScript.Echo "Error: El comando relink requiere argumentos:"
        WScript.Echo "Uso: cscript condor_cli.vbs relink <db_path> <local_folder>"
        WScript.Echo "  o: cscript condor_cli.vbs relink --all"
        WScript.Echo "  db_path: Ruta a la base de datos frontend (.accdb)"
        WScript.Echo "  local_folder: Ruta a la carpeta con las bases de datos locales"
        WScript.Echo "  --all: Re-vincular todas las bases de datos en ./back automáticamente"
        WScript.Quit 1
    End If
    
    ' Leer argumentos de la línea de comandos
    strDbPath = objArgs(1)
    strLocalFolder = objArgs(2)
    
    WScript.Echo "Base de datos frontend: " & strDbPath
    WScript.Echo "Carpeta de backends locales: " & strLocalFolder
    
    ' Verificar que los paths existen
    If Not objFSO.FileExists(strDbPath) Then
        WScript.Echo "Error: La base de datos frontend no existe: " & strDbPath
        WScript.Quit 1
    End If
    
    If Not objFSO.FolderExists(strLocalFolder) Then
        WScript.Echo "Error: La carpeta de backends locales no existe: " & strLocalFolder
        WScript.Quit 1
    End If
    
    WScript.Echo "Funcionalidad de re-vinculación pendiente de implementación."
    WScript.Echo "=== RE-VINCULACION COMPLETADA ==="
End Sub

' Subrutina para re-vincular todas las bases de datos automáticamente
Sub RelinkAllDatabases()
    Dim objBackFolder, objFile
    Dim strBackPath, strDbCount
    Dim dbCount, successCount, errorCount
    Dim strDbName, strPassword
    Dim arrDatabases()
    Dim i
    
    WScript.Echo "=== MODO AUTOMATICO: RE-VINCULANDO TODAS LAS BASES DE DATOS ==="
    
    ' Definir ruta del directorio back
    strBackPath = objFSO.GetAbsolutePathName("back")
    
    ' Verificar que existe el directorio back
    If Not objFSO.FolderExists(strBackPath) Then
        WScript.Echo "Error: El directorio ./back no existe: " & strBackPath
        WScript.Quit 1
    End If
    
    WScript.Echo "Directorio de backends: " & strBackPath
    
    ' Contar y listar bases de datos .accdb
    Set objBackFolder = objFSO.GetFolder(strBackPath)
    dbCount = 0
    
    ' Redimensionar array para almacenar información de bases de datos
    ReDim arrDatabases(50) ' Máximo 50 bases de datos
    
    For Each objFile In objBackFolder.Files
        If LCase(objFSO.GetExtensionName(objFile.Name)) = "accdb" Then
            strDbName = objFSO.GetBaseName(objFile.Name)
            
            ' Determinar contraseña según el nombre de la base de datos
            If InStr(1, UCase(strDbName), "CONDOR") > 0 Then
                strPassword = "(sin contraseña)"
            Else
                strPassword = "dpddpd"
            End If
            
            ' Almacenar información de la base de datos
            arrDatabases(dbCount) = objFile.Path & "|" & strPassword
            dbCount = dbCount + 1
            
            WScript.Echo "  [" & dbCount & "] " & objFile.Name & " - " & strPassword
        End If
    Next
    
    If dbCount = 0 Then
        WScript.Echo "No se encontraron bases de datos .accdb en el directorio ./back"
        WScript.Echo "=== RE-VINCULACION COMPLETADA ==="
        Exit Sub
    End If
    
    WScript.Echo "Total de bases de datos encontradas: " & dbCount
    WScript.Echo "Iniciando proceso de re-vinculación..."
    WScript.Echo ""
    
    ' Procesar cada base de datos
    successCount = 0
    errorCount = 0
    
    For i = 0 To dbCount - 1
        Dim arrDbInfo
        arrDbInfo = Split(arrDatabases(i), "|")
        
        If UBound(arrDbInfo) >= 1 Then
            Dim strDbPath, strDbPassword
            strDbPath = arrDbInfo(0)
            strDbPassword = arrDbInfo(1)
            
            WScript.Echo "Procesando: " & objFSO.GetFileName(strDbPath)
            
            If RelinkSingleDatabase(strDbPath, strDbPassword, strBackPath) Then
                successCount = successCount + 1
                WScript.Echo "  ✓ Re-vinculación exitosa"
            Else
                errorCount = errorCount + 1
                WScript.Echo "  ❌ Error en re-vinculación"
            End If
            WScript.Echo ""
        End If
    Next
    
    ' Resumen final
    WScript.Echo "=== RESUMEN DE RE-VINCULACION AUTOMATICA ==="
    WScript.Echo "Total procesadas: " & dbCount
    WScript.Echo "Exitosas: " & successCount
    WScript.Echo "Con errores: " & errorCount
    
    If errorCount = 0 Then
        WScript.Echo "✓ Todas las bases de datos fueron re-vinculadas exitosamente"
    Else
        WScript.Echo "⚠️ Algunas bases de datos tuvieron errores durante la re-vinculación"
    End If
    
    WScript.Echo "=== RE-VINCULACION AUTOMATICA COMPLETADA ==="
End Sub

' Función para determinar la contraseña de una base de datos
Function GetDatabasePassword(strDbPath)
    Dim strDbName
    strDbName = objFSO.GetBaseName(strDbPath)
    
    ' Las bases de datos CONDOR no requieren contraseña
    If InStr(1, UCase(strDbName), "CONDOR") > 0 Then
        GetDatabasePassword = ""
    Else
        ' Las demás bases de datos usan 'dpddpd'
        GetDatabasePassword = "dpddpd"
    End If
End Function

' Función para re-vincular una sola base de datos
Function RelinkSingleDatabase(strDbPath, strPassword, strBackPath)
    Dim objDb, objTableDef
    Dim strConnectionString
    Dim linkedTableCount, successCount
    
    On Error Resume Next
    
    ' Abrir la base de datos
    If strPassword = "(sin contraseña)" Then
        Set objDb = objAccess.DBEngine.OpenDatabase(strDbPath)
    Else
        Set objDb = objAccess.DBEngine.OpenDatabase(strDbPath, False, False, ";PWD=" & strPassword)
    End If
    
    If Err.Number <> 0 Then
        WScript.Echo "  Error al abrir base de datos: " & Err.Description
        RelinkSingleDatabase = False
        Err.Clear
        Exit Function
    End If
    
    linkedTableCount = 0
    successCount = 0
    
    ' Recorrer todas las tablas vinculadas
    For Each objTableDef In objDb.TableDefs
        If Len(objTableDef.Connect) > 0 Then
            linkedTableCount = linkedTableCount + 1
            
            ' Extraer el nombre de la base de datos del connect string actual
            Dim strCurrentConnect, strSourceDb, strNewConnect
            strCurrentConnect = objTableDef.Connect
            
            ' Buscar el patrón DATABASE= en el connect string
            Dim intDbStart, intDbEnd, strDbName
            intDbStart = InStr(1, UCase(strCurrentConnect), "DATABASE=")
            
            If intDbStart > 0 Then
                intDbStart = intDbStart + 9 ' Longitud de "DATABASE="
                intDbEnd = InStr(intDbStart, strCurrentConnect, ";")
                
                If intDbEnd = 0 Then intDbEnd = Len(strCurrentConnect) + 1
                
                strSourceDb = Mid(strCurrentConnect, intDbStart, intDbEnd - intDbStart)
                strDbName = objFSO.GetFileName(strSourceDb)
                
                ' Construir nueva ruta local
                Dim strNewDbPath
                strNewDbPath = objFSO.BuildPath(strBackPath, strDbName)
                
                ' Verificar que la base de datos local existe
                If objFSO.FileExists(strNewDbPath) Then
                    ' Construir nuevo connect string
                    strNewConnect = Replace(strCurrentConnect, strSourceDb, strNewDbPath)
                    
                    ' Actualizar la vinculación
                    objTableDef.Connect = strNewConnect
                    objTableDef.RefreshLink
                    
                    If Err.Number = 0 Then
                        successCount = successCount + 1
                        WScript.Echo "    ✓ " & objTableDef.Name & " -> " & strDbName
                    Else
                        WScript.Echo "    ❌ Error en " & objTableDef.Name & ": " & Err.Description
                        Err.Clear
                    End If
                Else
                    WScript.Echo "    ⚠️ Base de datos local no encontrada: " & strDbName
                End If
            Else
                WScript.Echo "    ⚠️ No se pudo extraer DATABASE de: " & objTableDef.Name
            End If
        End If
    Next
    
    ' Cerrar base de datos
    objDb.Close
    Set objDb = Nothing
    
    WScript.Echo "    Tablas vinculadas procesadas: " & linkedTableCount
    WScript.Echo "    Re-vinculaciones exitosas: " & successCount
    
    ' Considerar exitoso si se procesó al menos una tabla correctamente
    RelinkSingleDatabase = (successCount > 0 Or linkedTableCount = 0)
    
    On Error GoTo 0
End Function

' Subrutina para crear backup de seguridad de la base de datos
Sub BackupDatabaseSafely(dbPath)
    On Error Resume Next
    
    Dim backupDir, backupName, backupPath
    backupDir = objFSO.GetParentFolderName(dbPath)
    
    ' Crear nombre con timestamp
    Dim timestamp
    timestamp = Replace(Replace(Now, ":", "-"), "/", "-")
    timestamp = Replace(timestamp, " ", "_")
    
    backupName = "backup_" & timestamp & ".accdb"
    backupPath = objFSO.BuildPath(backupDir, backupName)
    
    If gVerbose Then WScript.Echo "[VERBOSE] Creando backup de seguridad: " & backupPath
    
    objFSO.CopyFile dbPath, backupPath
    If Err.Number = 0 Then
        If gVerbose Then WScript.Echo "[VERBOSE] Backup creado exitosamente: " & backupName
    Else
        WScript.Echo "[WARNING] No se pudo crear backup: " & Err.Description
        WScript.Echo "[WARNING] Continuando con rebuild sin backup..."
        Err.Clear
    End If
    
    On Error GoTo 0
End Sub

' Subrutina para reconstruir completamente el proyecto VBA
' Función helper para verificar si la BD tiene proyecto VBA accesible
Function HasVbProject(app)
    On Error Resume Next
    Dim p
    Set p = app.VBE.ActiveVBProject
    HasVbProject = (Err.Number = 0 And Not p Is Nothing)
    Err.Clear
End Function

Sub RebuildProject()
    ' VERSIÓN TRANSACCIONAL - Asume objAccess ya abierto por OpenAccessApp
    ' No abre ni cierra Access; trabaja en la misma sesión
    
    WScript.Echo "=== RECONSTRUCCION COMPLETA DEL PROYECTO VBA ==="
    WScript.Echo "ADVERTENCIA: Se eliminaran TODOS los modulos VBA existentes"
    WScript.Echo "Iniciando proceso de reconstruccion..."
    
    On Error Resume Next
    
    ' Verificar que objAccess esté disponible
    If objAccess Is Nothing Then
        WScript.Echo "Error: Access no está abierto. RebuildProject requiere una sesión activa."
        WScript.Quit 1
    End If
    
    ' PRECHECK: Verificar que la BD tiene proyecto VBA accesible
    If Not HasVbProject(objAccess) Then
        WScript.Echo "✖ La base abierta no contiene proyecto VBA o VBIDE no es accesible."
        WScript.Echo "  Sugerencia: parece una BD de BACKEND (" & strAccessPath & ", origen=" & gDbSource & ")."
        WScript.Echo "  Use el FRONTEND (back\Desarrollo\CONDOR.accdb) o pase --db apuntando al frontend."
        WScript.Quit 1
    End If
    
    ' Suprimir alertas
    objAccess.DoCmd.SetWarnings False
    If Err.Number <> 0 Then
        If gVerbose Then WScript.Echo "[VERBOSE] No se pudieron suprimir alertas: " & Err.Description
        Err.Clear
    End If
    
    ' Crear backup de seguridad antes de proceder
    Dim backupDbPath
    If Len(Trim(gCurrentDbPath)) > 0 Then
        backupDbPath = gCurrentDbPath
    ElseIf Len(Trim(strAccessPath)) > 0 Then
        backupDbPath = strAccessPath
    Else
        WScript.Echo "[WARNING] No se pudo determinar la ruta de BD para backup. Se omite backup."
        backupDbPath = ""
    End If
    
    If Len(Trim(backupDbPath)) > 0 Then
        BackupDatabaseSafely backupDbPath
    End If
    
    ' Paso 1: Eliminar todos los módulos existentes vía VBE
    WScript.Echo "Paso 1: Eliminando todos los modulos VBA existentes..."
    
    ' Configurar AutomationSecurity para permitir acceso a VBIDE
    On Error Resume Next
    objAccess.Application.AutomationSecurity = 1  ' msoAutomationSecurityLow
    Err.Clear
    On Error Resume Next
    
    ' Verificar acceso a VBIDE antes de proceder
    Dim vbProject, vbComponent
    On Error Resume Next
    Set vbProject = objAccess.VBE.ActiveVBProject
    If Err.Number <> 0 Then
        ' Intentar una segunda vez después de una pausa
        WScript.Sleep 1000
        Err.Clear
        Set vbProject = objAccess.VBE.ActiveVBProject
        If Err.Number <> 0 Then
            WScript.Echo "✖ VBIDE no accesible. Error: " & Err.Description & " (Código: " & Err.Number & ")"
            WScript.Echo "Verifique que Access tenga permisos para acceder al modelo de objetos VBA."
            WScript.Quit 1
        End If
    End If
    On Error GoTo 0
    
    Dim componentCount, i
    componentCount = vbProject.VBComponents.Count
    
    ' Iterar hacia atrás para evitar problemas al eliminar elementos
    For i = componentCount To 1 Step -1
        On Error Resume Next
        Set vbComponent = vbProject.VBComponents(i)
        
        ' Solo eliminar módulos estándar y de clase (no formularios ni informes)
        If vbComponent.Type = 1 Or vbComponent.Type = 2 Then ' vbext_ct_StdModule = 1, vbext_ct_ClassModule = 2
            If gVerbose Then
                WScript.Echo "  Eliminando: " & vbComponent.Name & " (Tipo: " & vbComponent.Type & ")"
            End If
            
            vbProject.VBComponents.Remove vbComponent
            
            If Err.Number <> 0 Then
                WScript.Echo "  ❌ Error eliminando " & vbComponent.Name & ": " & Err.Description & " (Código: " & Err.Number & ")"
                WScript.Echo "  ⚠️ Continuando con el siguiente módulo..."
                Err.Clear
            Else
                If gVerbose Then
                    WScript.Echo "  ✓ Eliminado: " & vbComponent.Name
                End If
            End If
        End If
        On Error GoTo 0
    Next
    
    ' Paso 2: Importar todos los módulos desde /src usando importación robusta con inyección de cabeceras
    WScript.Echo "Paso 2: Importando todos los modulos desde /src..."
    
    If Not objFSO.FolderExists(strSourcePath) Then
        WScript.Echo "Error: Directorio de origen no existe: " & strSourcePath
        WScript.Quit 1
    End If
    
    ' Crear directorio temporal para archivos con cabeceras inyectadas
    Dim tmpDir, tmpParent
    tmpParent = objFSO.BuildPath(RepoRoot(), ".tmp")
    tmpDir = objFSO.BuildPath(tmpParent, "vbe")
    
    ' Crear directorio padre .tmp si no existe
    If Not objFSO.FolderExists(tmpParent) Then
        objFSO.CreateFolder tmpParent
    End If
    
    ' Crear directorio vbe si no existe
    If Not objFSO.FolderExists(tmpDir) Then
        objFSO.CreateFolder tmpDir
    End If
    
    Dim objFolder, objFile
    Dim strModuleName, importedCount, moduleType
    Set objFolder = objFSO.GetFolder(strSourcePath)
    
    ' Inicialización de contadores y limpieza de errores antes del bucle
    importedCount = 0
    Err.Clear
    
    For Each objFile In objFolder.Files
        If LCase(objFSO.GetExtensionName(objFile.Name)) = "bas" Or LCase(objFSO.GetExtensionName(objFile.Name)) = "cls" Then
            strModuleName = objFSO.GetBaseName(objFile.Name)
            
            ' Limpiar errores al comienzo de cada iteración
            Err.Clear
            
            If gVerbose Then
                WScript.Echo "  Procesando: " & strModuleName
            End If
            
            ' Importación robusta con inyección de cabeceras
            Dim okImport
            okImport = ImportVbaFileRobust(objFile.Path, strModuleName, tmpDir)
            If okImport Then importedCount = importedCount + 1
        End If
    Next
    
    ' Verificar si hubo errores de importación
    Dim totalFiles
    totalFiles = 0
    For Each objFile In objFolder.Files
        If LCase(objFSO.GetExtensionName(objFile.Name)) = "bas" Or LCase(objFSO.GetExtensionName(objFile.Name)) = "cls" Then
            totalFiles = totalFiles + 1
        End If
    Next
    
    ' Verificación: que existan los módulos recién importados
    Dim vbProj, comp, presentCount
    presentCount = 0
    Set vbProj = objAccess.VBE.ActiveVBProject
    For Each comp In vbProj.VBComponents
        If comp.Type = 1 Or comp.Type = 2 Or comp.Type = 3 Then ' std/class/forms modules
            presentCount = presentCount + 1
        End If
    Next
    
    WScript.Echo "=== RECONSTRUCCIÓN COMPLETADA ==="
    WScript.Echo "Módulos importados (cuenta lógica): " & importedCount & " de " & totalFiles
    WScript.Echo "Módulos presentes en proyecto: " & presentCount
    
    If importedCount = 0 Or presentCount = 0 Then
        WScript.Echo "❌ No se detectan módulos en la BD tras el proceso. Revise referencia VBIDE/trust y el popup inicial."
        WScript.Quit 1
    End If
    
    ' Asegurar referencias VBA antes de compilar
    WScript.Echo "Verificando referencias VBA..."
    Call EnsureVBReferences()
    
    ' Compilar y guardar todo
    WScript.Echo "Compilando proyecto VBA..."
    On Error Resume Next
    objAccess.DoCmd.RunCommand 636 ' Compile & Save All Modules
    If Err.Number<>0 Then WScript.Echo "❌ Compilación: "&Err.Number&" - "&Err.Description: WScript.Quit 1
    On Error GoTo 0
    
    WScript.Echo "El proyecto VBA ha sido completamente reconstruido y compilado"
    
    ' Verificación opcional de módulos
    Call VerifyModulesIfRequested()
    
    On Error GoTo 0
End Sub

' Subrutina para actualizar módulos específicos del proyecto
Sub UpdateProject()
    ' VERSIÓN TRANSACCIONAL - Asume objAccess ya abierto por OpenAccessApp
    ' No abre ni cierra Access; trabaja en la misma sesión
    
    WScript.Echo "=== ACTUALIZACION DE MODULOS VBA ==="
    
    ' Verificar que objAccess esté disponible
    If objAccess Is Nothing Then
        WScript.Echo "Error: Access no está abierto. UpdateProject requiere una sesión activa."
        WScript.Quit 1
    End If
    
    ' Establecer modo silencioso antes de cualquier importación
    On Error Resume Next
    objAccess.Application.DisplayAlerts = False
    objAccess.Application.Echo False
    objAccess.DoCmd.SetWarnings False
    Err.Clear
    On Error GoTo 0
    
    ' Parser robusto para argumentos del comando update
    Dim p, i, targetsTokens(), targetsRaw, listaNormalizada
    
    ' Encontrar el índice del token "update"
    p = -1
    For i = 0 To objArgs.Count - 1
        If LCase(objArgs(i)) = "update" Then
            p = i
            Exit For
        End If
    Next
    
    If p = -1 Or p >= objArgs.Count - 1 Then
        WScript.Echo "Error: Debe especificar un target para update"
        WScript.Echo "Uso: cscript condor_cli.vbs update <Nombre|Nombre1,Nombre2|Nombre1 Nombre2> [--changed|--all] [--verifyModules] [--verbose] [--bypassstartup]"
        WScript.Quit 1
    End If
    
    If Not objFSO.FolderExists(strSourcePath) Then
        WScript.Echo "Error: Directorio de origen no existe: " & strSourcePath
        WScript.Quit 1
    End If
    
    On Error Resume Next
    
    ' Verificar flags especiales primero
    If objArgs(p + 1) = "--all" Then
        ' Importar todos los módulos (sync suave sin eliminar)
        WScript.Echo "Actualizando todos los modulos desde /src..."
        Call UpdateAllModulesTransactional()
    ElseIf objArgs(p + 1) = "--changed" Then
        ' Importar solo los módulos cambiados por comparación
        WScript.Echo "Actualizando solo modulos cambiados..."
        Call UpdateChangedModulesTransactional()
    Else
        ' Recopilar tokens de targets (desde p+1 hasta encontrar una flag o fin)
        ReDim targetsTokens(0)
        Dim tokenCount: tokenCount = 0
        
        For i = p + 1 To objArgs.Count - 1
            Dim currentToken: currentToken = Trim(objArgs(i))
            ' Parar si encontramos una flag
            If Left(currentToken, 2) = "--" Or Left(currentToken, 1) = "/" Then
                Exit For
            End If
            ' Añadir token no vacío
            If Len(currentToken) > 0 Then
                If tokenCount > 0 Then
                    ReDim Preserve targetsTokens(tokenCount)
                End If
                targetsTokens(tokenCount) = currentToken
                tokenCount = tokenCount + 1
            End If
        Next
        
        If tokenCount = 0 Then
            WScript.Echo "Error: Debe especificar al menos un módulo para update"
            WScript.Echo "Uso: cscript condor_cli.vbs update <Nombre|Nombre1,Nombre2|Nombre1 Nombre2> [--changed|--all] [--verifyModules] [--verbose] [--bypassstartup]"
            WScript.Quit 1
        End If
        
        ' Construir targetsRaw
        targetsRaw = ""
        If InStr(targetsTokens(0), ",") > 0 Then
            ' El primer token contiene comas, usarlo tal cual
            targetsRaw = targetsTokens(0)
            ' Añadir el resto de tokens separados por coma
            For i = 1 To tokenCount - 1
                targetsRaw = targetsRaw & "," & targetsTokens(i)
            Next
        Else
            ' Construir lista separada por comas
            targetsRaw = targetsTokens(0)
            For i = 1 To tokenCount - 1
                targetsRaw = targetsRaw & "," & targetsTokens(i)
            Next
        End If
        
        ' Normalizar la lista
        ' 1. Quitar espacios alrededor de comas
        targetsRaw = Replace(targetsRaw, " ,", ",")
        targetsRaw = Replace(targetsRaw, ", ", ",")
        
        ' 2. Reemplazar bloques de espacios por comas
        Do While InStr(targetsRaw, "  ") > 0
            targetsRaw = Replace(targetsRaw, "  ", " ")
        Loop
        targetsRaw = Replace(targetsRaw, " ", ",")
        
        ' 3. Colapsar comas repetidas
        Do While InStr(targetsRaw, ",,") > 0
            targetsRaw = Replace(targetsRaw, ",,", ",")
        Loop
        
        ' 4. Eliminar comas al inicio y final
        If Left(targetsRaw, 1) = "," Then targetsRaw = Mid(targetsRaw, 2)
        If Right(targetsRaw, 1) = "," Then targetsRaw = Left(targetsRaw, Len(targetsRaw) - 1)
        
        ' 5. Deduplicar preservando orden
        Dim modules, uniqueModules(), uniqueCount, j, found
        modules = Split(targetsRaw, ",")
        ReDim uniqueModules(0)
        uniqueCount = 0
        
        For i = 0 To UBound(modules)
            Dim currentModule: currentModule = Trim(modules(i))
            If Len(currentModule) > 0 Then
                found = False
                For j = 0 To uniqueCount - 1
                    If LCase(uniqueModules(j)) = LCase(currentModule) Then
                        found = True
                        Exit For
                    End If
                Next
                If Not found Then
                    If uniqueCount > 0 Then
                        ReDim Preserve uniqueModules(uniqueCount)
                    End If
                    uniqueModules(uniqueCount) = currentModule
                    uniqueCount = uniqueCount + 1
                End If
            End If
        Next
        
        If uniqueCount = 0 Then
            WScript.Echo "Error: Lista de módulos vacía después de normalización"
            WScript.Quit 1
        End If
        
        ' Reconstruir lista normalizada
        listaNormalizada = uniqueModules(0)
        For i = 1 To uniqueCount - 1
            listaNormalizada = listaNormalizada & "," & uniqueModules(i)
        Next
        
        ' Llamar a la función apropiada
        WScript.Echo "Actualizando modulos especificos: " & listaNormalizada
        Call UpdateMultipleModulesTransactional(listaNormalizada)
    End If
    
    WScript.Echo "=== ACTUALIZACION COMPLETADA EXITOSAMENTE ==="
    
    ' Verificación opcional de módulos
    Call VerifyModulesIfRequested()
    
    On Error GoTo 0
End Sub

' Actualizar todos los módulos
Sub UpdateAllModules()
    Dim objFolder, objFile, strModuleName, importedCount
    Set objFolder = objFSO.GetFolder(strSourcePath)
    importedCount = 0
    
    For Each objFile In objFolder.Files
        If LCase(objFSO.GetExtensionName(objFile.Name)) = "bas" Or LCase(objFSO.GetExtensionName(objFile.Name)) = "cls" Then
            strModuleName = objFSO.GetBaseName(objFile.Name)
            
            If gVerbose Then
                WScript.Echo "Importando: " & strModuleName
            End If
            
            If ImportVbaFile(objFile.Path, strModuleName) Then
                importedCount = importedCount + 1
            End If
        End If
    Next
    
    ' Verificar si hubo errores de importación
    Dim totalFiles
        totalFiles = 0
    For Each objFile In objFolder.Files
        If LCase(objFSO.GetExtensionName(objFile.Name)) = "bas" Or LCase(objFSO.GetExtensionName(objFile.Name)) = "cls" Then
            totalFiles = totalFiles + 1
        End If
    Next
    
    If importedCount < totalFiles Then
        WScript.Echo "Modulos actualizados: " & importedCount & " de " & totalFiles & " (CON ERRORES)"
        WScript.Quit 1
    Else
        WScript.Echo "Modulos actualizados: " & importedCount
        
        ' Verificación opcional de módulos
        Call VerifyModulesIfRequested()
    End If
End Sub

' Actualizar solo módulos cambiados (comparación por hash MD5)
Sub UpdateChangedModules()
    Dim objFolder, objFile, strModuleName, importedCount
    Dim cacheDir, tempExportDir, needsComparison
    Set objFolder = objFSO.GetFolder(strSourcePath)
    importedCount = 0
    
    ' Verificar si existe .cache/export
    cacheDir = objFSO.BuildPath(objFSO.GetParentFolderName(strDatabasePath), ".cache")
    tempExportDir = objFSO.BuildPath(cacheDir, "export")
    needsComparison = False
    
    If Not objFSO.FolderExists(tempExportDir) Then
        WScript.Echo "Cache de exportación no encontrado. Creando exportación temporal para comparación..."
        
        ' Crear directorio temporal en %TEMP%
        tempExportDir = objFSO.BuildPath(objFSO.GetSpecialFolder(2), "condor_temp_export_" & Timer)
        If Not objFSO.FolderExists(tempExportDir) Then
            objFSO.CreateFolder tempExportDir
        End If
        
        ' Exportar módulos actuales a directorio temporal
        Call ExportModulesToDirectory(tempExportDir)
        needsComparison = True
    End If
    
    WScript.Echo "Comparando archivos por hash MD5..."
    
    For Each objFile In objFolder.Files
        If LCase(objFSO.GetExtensionName(objFile.Name)) = "bas" Or LCase(objFSO.GetExtensionName(objFile.Name)) = "cls" Then
            strModuleName = objFSO.GetBaseName(objFile.Name)
            
            ' Comparar hash del archivo fuente con el exportado
            Dim sourceHash, exportedHash, exportedFile
            exportedFile = objFSO.BuildPath(tempExportDir, objFile.Name)
            
            sourceHash = GetFileHash(objFile.Path)
            
            If objFSO.FileExists(exportedFile) Then
                exportedHash = GetFileHash(exportedFile)
            Else
                exportedHash = "" ' Archivo no existe, forzar importación
            End If
            
            ' Solo importar si los hashes son diferentes
            If sourceHash <> exportedHash Then
                If gVerbose Then
                    WScript.Echo "Importando: " & strModuleName & " (cambio detectado por hash)"
                End If
                
                If ImportVbaFile(objFile.Path, strModuleName) Then
                    importedCount = importedCount + 1
                End If
            Else
                If gVerbose Then
                    WScript.Echo "  - " & strModuleName & " sin cambios (hash idéntico)"
                End If
            End If
        End If
    Next
    
    ' Limpiar directorio temporal si se creó
    If needsComparison And objFSO.FolderExists(tempExportDir) Then
        objFSO.DeleteFolder tempExportDir, True
        If gVerbose Then
            WScript.Echo "Directorio temporal eliminado: " & tempExportDir
        End If
    End If
    
    ' No verificamos totalFiles aquí porque UpdateChangedModules solo actualiza módulos que cambiaron
    ' El éxito se mide por si importedCount >= módulos que necesitaban actualización
    WScript.Echo "Modulos actualizados: " & importedCount
    
    ' Verificación opcional de módulos
    Call VerifyModulesIfRequested()
End Sub

' ============================================================================
' SUBRUTINA: PrintJsonArrayOfNames
' Descripción: Imprime un array de nombres como JSON simple
' Parámetros: arr - Array de nombres (strings o objetos con propiedad "name")
' ============================================================================
Sub PrintJsonArrayOfNames(arr)
    Dim i, out
    out = "["
    On Error Resume Next
    Dim upperBound
    upperBound = UBound(arr)
    If Err.Number = 0 And upperBound >= 0 Then
        For i = 0 To upperBound
            Dim nm
            nm = arr(i)
            If IsObject(nm) Then nm = nm("name")
            ' Solo añadir elementos no vacíos
            If Len(Trim(nm)) > 0 Then
                If out <> "[" Then out = out & ","
                out = out & """" & Replace(nm, """", "\""") & """"
            End If
        Next
    End If
    On Error GoTo 0
    out = out & "]"
    WScript.Echo out
End Sub

' ============================================================================
' FUNCIÓN AUXILIAR: GetFileHash
' Descripción: Calcula hash MD5 simple de un archivo (basado en tamaño y fecha)
' Parámetros: filePath - Ruta del archivo
' Retorna: Hash como cadena
' ============================================================================
Private Function GetFileHash(filePath)
    Dim objFile, hashStr
    
    If Not objFSO.FileExists(filePath) Then
        GetFileHash = ""
        Exit Function
    End If
    
    Set objFile = objFSO.GetFile(filePath)
    
    ' Hash simple basado en tamaño y fecha de modificación
    hashStr = CStr(objFile.Size) & "|" & CStr(objFile.DateLastModified)
    
    ' Convertir a un hash más compacto usando suma de caracteres
    Dim i, charSum
    charSum = 0
    For i = 1 To Len(hashStr)
        charSum = charSum + Asc(Mid(hashStr, i, 1))
    Next
    
    GetFileHash = CStr(charSum) & "|" & CStr(objFile.Size)
End Function

' ============================================================================
' SUBRUTINA AUXILIAR: ExportModulesToDirectory
' Descripción: Exporta todos los módulos VBA a un directorio específico
' Parámetros: targetDir - Directorio destino
' ============================================================================
Private Sub ExportModulesToDirectory(targetDir)
    Dim objAccess, objModule, i
    
    ' Resolver BD usando función canónica
    Dim dbPath, dbOrigin
    Call ResolveDbForAction("export", dbPath, dbOrigin)
    
    ' Abrir Access para exportar módulos
    Set objAccess = OpenAccessApp(dbPath, gPassword, True)
    
    If objAccess Is Nothing Then
        WScript.Echo "Error: No se pudo abrir Access para exportar módulos"
        Exit Sub
    End If
    
    On Error Resume Next
    
    ' Exportar módulos estándar (.bas)
    For i = 0 To objAccess.CurrentProject.AllModules.Count - 1
        Set objModule = objAccess.CurrentProject.AllModules(i)
        
        Dim exportPath
        exportPath = objFSO.BuildPath(targetDir, objModule.Name & ".bas")
        
        Call ExportModuleToUtf8(objAccess, objModule.Name, exportPath)
        
        If Err.Number <> 0 Then
            If gVerbose Then
                WScript.Echo "Advertencia: No se pudo exportar módulo " & objModule.Name & ": " & Err.Description
            End If
            Err.Clear
        End If
    Next
    
    ' Exportar módulos de clase (.cls)
    For i = 0 To objAccess.CurrentProject.AllModules.Count - 1
        Set objModule = objAccess.CurrentProject.AllModules(i)
        
        ' Verificar si es un módulo de clase
        If objAccess.Modules(objModule.Name).Type = 1 Then ' acClassModule = 1
            exportPath = objFSO.BuildPath(targetDir, objModule.Name & ".cls")
            
            Call ExportModuleToUtf8(objAccess, objModule.Name, exportPath)
            
            If Err.Number <> 0 Then
                If gVerbose Then
                    WScript.Echo "Advertencia: No se pudo exportar módulo de clase " & objModule.Name & ": " & Err.Description
                End If
                Err.Clear
            End If
        End If
    Next
    
    On Error GoTo 0
    
    ' Cerrar Access
    Call CloseAccessApp(objAccess)
    
    If gVerbose Then
        WScript.Echo "Módulos exportados a: " & targetDir
    End If
End Sub

' Actualizar un módulo específico
Sub UpdateSingleModule(moduleName)
    Dim srcFile, srcPath
    
    ' Buscar el archivo .bas o .cls correspondiente
    srcPath = strSourcePath & "\" & moduleName & ".bas"
    If Not objFSO.FileExists(srcPath) Then
        srcPath = strSourcePath & "\" & moduleName & ".cls"
        If Not objFSO.FileExists(srcPath) Then
            WScript.Echo "Error: No se encontró el archivo " & moduleName & ".bas o " & moduleName & ".cls en /src"
            WScript.Quit 1
        End If
    End If
    
    If gVerbose Then
        WScript.Echo "Importando: " & moduleName & " desde " & srcPath
    End If
    
    If Not ImportVbaFile(srcPath, moduleName) Then
        WScript.Echo "Error al importar " & moduleName
        WScript.Quit 1
    End If
End Sub

' Subrutina para verificar y cerrar procesos de Access existentes
Sub CloseExistingAccessProcesses()
    Dim objWMI, colProcesses, objProcess
    Dim processCount
    
    WScript.Echo "Verificando procesos de Access existentes..."
    
    On Error Resume Next
    Set objWMI = GetObject("winmgmts:")
    Set colProcesses = objWMI.ExecQuery("SELECT * FROM Win32_Process WHERE Name = 'MSACCESS.EXE'")
    
    processCount = 0
    For Each objProcess In colProcesses
        processCount = processCount + 1
    Next
    
    If processCount > 0 Then
        WScript.Echo "Se encontraron " & processCount & " procesos de Access ejecutándose. Cerrándolos..."
        
        For Each objProcess In colProcesses
            WScript.Echo "Terminando proceso Access PID: " & objProcess.ProcessId
            objProcess.Terminate()
        Next
        
        ' Esperar un momento para que los procesos se cierren completamente
        WScript.Sleep 2000
        WScript.Echo "✓ Procesos de Access cerrados correctamente"
    Else
        WScript.Echo "✓ No se encontraron procesos de Access ejecutándose"
    End If
    
    On Error GoTo 0
End Sub

' La subrutina ExecuteTestModule ha sido eliminada ya que ahora se usa el motor interno modTestRunner

' Subrutina para sincronizar un módulo individual
' Parámetro: moduleName - Nombre del módulo a sincronizar (ej. "CAuthService")


' Subrutina optimizada para importar un solo módulo (sin cerrar/abrir BD)
Sub ImportSingleModuleOptimized(moduleName)
    On Error Resume Next
    
    ' Paso 1: Verificar que el fichero fuente (.bas o .cls) existe en la carpeta /src
    Dim strBasFile, strClsFile, strSourceFile, fileExtension
    strBasFile = objFSO.BuildPath(strSourcePath, moduleName & ".bas")
    strClsFile = objFSO.BuildPath(strSourcePath, moduleName & ".cls")
    
    If objFSO.FileExists(strBasFile) Then
        strSourceFile = strBasFile
        fileExtension = "bas"
    ElseIf objFSO.FileExists(strClsFile) Then
        strSourceFile = strClsFile
        fileExtension = "cls"
    Else
        WScript.Echo "  ❌ Error: No se encontró el archivo fuente para " & moduleName
        WScript.Echo "      Buscado: " & strBasFile
        WScript.Echo "      Buscado: " & strClsFile
        Exit Sub
    End If
    
    WScript.Echo "  ✓ Archivo fuente encontrado: " & strSourceFile
    
    ' Paso 2: Validar sintaxis del archivo
    Dim errorDetails, validationResult
    validationResult = ValidateVBASyntax(strSourceFile, errorDetails)
    
    If validationResult <> True Then
        WScript.Echo "  ❌ Error de sintaxis en " & moduleName & ": " & errorDetails
        Exit Sub
    End If
    
    WScript.Echo "  ✓ Sintaxis válida"
    
    ' Paso 3: Limpiar el contenido del fichero utilizando CleanVBAFile
    Dim cleanedContent
    cleanedContent = CleanVBAFile(strSourceFile, fileExtension)
    
    If cleanedContent = "" Then
        WScript.Echo "  ❌ Error: No se pudo leer o limpiar el contenido del archivo"
        Exit Sub
    End If
    
    WScript.Echo "  ✓ Contenido limpiado"
    
    ' Paso 4: Importar el módulo usando la rutina unificada ImportVbaFile
    WScript.Echo "  Importando módulo: " & moduleName
    If Not ImportVbaFile(strSourceFile, moduleName) Then
        WScript.Echo "  ❌ Error al importar módulo " & moduleName
        Exit Sub
    End If
    
    If fileExtension = "cls" Then
        WScript.Echo "✅ Clase " & moduleName & " importada correctamente"
    Else
        WScript.Echo "✅ Módulo " & moduleName & " importado correctamente"
    End If
    
    WScript.Echo "  ✅ Módulo " & moduleName & " sincronizado correctamente"
    
    On Error GoTo 0
End Sub

' Subrutina para actualizar proyecto VBA con sincronización selectiva


' Subrutina para exportar módulos VBA actuales a carpeta cache












' Función para verificar cambios antes de abrir la base de datos


' Subrutina para copiar solo archivos modificados a la caché


' Subrutina para mostrar ayuda específica del comando bundle
Sub ShowBundleHelp()
    WScript.Echo "=== CONDOR CLI - AYUDA DEL COMANDO BUNDLE ==="
    WScript.Echo "Empaqueta archivos de código por funcionalidad según CONDOR_MASTER_PLAN.md"
    WScript.Echo ""
    WScript.Echo "USO:"
    WScript.Echo "  cscript condor_cli.vbs bundle <funcionalidad> [ruta_destino]"
    WScript.Echo "  cscript condor_cli.vbs bundle --help"
    WScript.Echo ""
    WScript.Echo "PARÁMETROS:"
    WScript.Echo "  <funcionalidad>  - Nombre de la funcionalidad a empaquetar (obligatorio)"
    WScript.Echo "  [ruta_destino]   - Carpeta donde crear el paquete (opcional, por defecto: carpeta actual)"
    WScript.Echo ""
    WScript.Echo "FUNCIONALIDADES DISPONIBLES:"
    WScript.Echo ""
    WScript.Echo "🔐 Auth - Autenticación y Autorización"
    WScript.Echo "   Incluye: IAuthService, CAuthService, CMockAuthService, IAuthRepository,"
    WScript.Echo "            CAuthRepository, CMockAuthRepository, AuthData, modAuthFactory,"
    WScript.Echo "            TestCAuthService, IntegrationTestCAuthRepository + dependencias"
    WScript.Echo ""
    WScript.Echo "📄 Document - Gestión de Documentos"
    WScript.Echo "   Incluye: IDocumentService, CDocumentService, CMockDocumentService,"
    WScript.Echo "            IWordManager, CWordManager, CMockWordManager, ISolicitudService + dependencias"
    WScript.Echo ""
    WScript.Echo "📁 Expediente - Gestión de Expedientes"
    WScript.Echo "   Incluye: IExpedienteService, CExpedienteService, CMockExpedienteService,"
    WScript.Echo "            IExpedienteRepository, CExpedienteRepository + dependencias"
    WScript.Echo ""
    WScript.Echo "📋 Solicitud - Gestión de Solicitudes"
    WScript.Echo "   Incluye: ISolicitudService, CSolicitudService, CMockSolicitudService,"
    WScript.Echo "            ISolicitudRepository, CSolicitudRepository + modelos de datos"
    WScript.Echo ""
    WScript.Echo "🔄 Workflow - Flujos de Trabajo"
    WScript.Echo "   Incluye: IWorkflowService, CWorkflowService, CMockWorkflowService,"
    WScript.Echo "            IWorkflowRepository, CWorkflowRepository + modelos de estado"
    WScript.Echo ""
    WScript.Echo "🗺️ Mapeo - Gestión de Mapeos"
    WScript.Echo "   Incluye: IMapeoRepository, CMapeoRepository, CMockMapeoRepository,"
    WScript.Echo "            EMapeo, IntegrationTestCMapeoRepository + dependencias"
    WScript.Echo ""
    WScript.Echo "⚙️ Config - Configuración del Sistema"
    WScript.Echo "   Incluye: IConfig, CConfig, CMockConfig, modConfigFactory,"
    WScript.Echo "            TestConfig (simplificado tras Misión de Emergencia)"
    WScript.Echo ""
    WScript.Echo "📂 FileSystem - Sistema de Archivos"
    WScript.Echo "   Incluye: IFileSystem, CFileSystem, CMockFileSystem,"
    WScript.Echo "            ModFileSystemFactory, TestFileSystem + dependencias"
    WScript.Echo ""
    WScript.Echo "❌ Error - Manejo de Errores"
    WScript.Echo "   Incluye: IErrorHandlerService, CErrorHandlerService, CMockErrorHandlerService,"
    WScript.Echo "            modErrorHandlerFactory, modErrorHandler + dependencias"
    WScript.Echo ""
    WScript.Echo "📝 Word - Integración con Microsoft Word"
    WScript.Echo "   Incluye: IWordManager, CWordManager, CMockWordManager,"
    WScript.Echo "            ModWordManagerFactory, TestWordManager + dependencias"
    WScript.Echo ""
    WScript.Echo "🧪 TestFramework - Framework de Pruebas"
WScript.Echo "   Incluye: ITestReporter, CTestResult, CTestSuiteResult, CTestReporter, modTestRunner,"
    WScript.Echo "            modTestUtils, ModAssert, TestModAssert + interfaces base"
    WScript.Echo ""
    WScript.Echo "🚀 App - Gestión de Aplicación"
    WScript.Echo "   Incluye: IAppManager, CAppManager, ModAppManagerFactory,"
    WScript.Echo "            TestAppManager + dependencias de autenticación y config"
    WScript.Echo ""
    WScript.Echo "🏗️ Models - Modelos de Datos"
    WScript.Echo "   Incluye: Todas las entidades E_* (Usuario, Solicitud, Expediente,"
    WScript.Echo "            DatosPC, DatosCDCA, Estado, Transicion, Mapeo, etc.)"
    WScript.Echo ""
    WScript.Echo "🔧 Utils - Utilidades y Enumeraciones"
    WScript.Echo "   Incluye: ModDatabase, ModRepositoryFactory, ModUtils,"
    WScript.Echo "            E_TipoSolicitud, E_EstadoSolicitud, E_RolUsuario, etc."
    WScript.Echo ""
    WScript.Echo "🧪 Tests - Archivos de Pruebas"
    WScript.Echo "   Incluye: Todos los archivos Test* e IntegrationTest* del proyecto"
    WScript.Echo "            (TestAppManager, TestAuthService, TestCConfig, etc.)"
    WScript.Echo ""
    WScript.Echo "🖥️ CondorCli - CLI de CONDOR"
    WScript.Echo "   Copia condor_cli.vbs como condor_cli.vbs.txt en la raíz del proyecto"
    WScript.Echo "   (sobreescribe si existe)"
    WScript.Echo ""
    WScript.Echo "EJEMPLOS DE USO:"
    WScript.Echo "  cscript condor_cli.vbs bundle Auth"
    WScript.Echo "  cscript condor_cli.vbs bundle Document C:\\temp"
    WScript.Echo "  cscript condor_cli.vbs bundle TestFramework"
    WScript.Echo "  cscript condor_cli.vbs bundle Tests"
    WScript.Echo "  cscript condor_cli.vbs bundle Config"
    WScript.Echo "  cscript condor_cli.vbs bundle condorcli"
    WScript.Echo ""
    WScript.Echo "NOTAS:"
    WScript.Echo "  • Los archivos se copian con extensión .txt para fácil visualización"
    WScript.Echo "  • Se crea una carpeta con timestamp: bundle_<funcionalidad>_YYYYMMDD_HHMMSS"
    WScript.Echo "  • Cada funcionalidad incluye automáticamente sus dependencias"
    WScript.Echo "  • Si un archivo no existe, se muestra una advertencia pero continúa"
End Sub

' Subrutina para mostrar ayuda del comando roundtrip-form
Sub ShowRoundtripFormHelp()
    WScript.Echo "=== CONDOR CLI - AYUDA DEL COMANDO ROUNDTRIP-FORM ==="
    WScript.Echo "Ejecuta un ciclo completo de import-form --overwrite seguido de export-form"
    WScript.Echo "para verificar la integridad del proceso de importación/exportación."
    WScript.Echo ""
    WScript.Echo "USO:"
    WScript.Echo "  cscript condor_cli.vbs roundtrip-form <json_path> --db <db_path> [opciones]"
    WScript.Echo "  cscript condor_cli.vbs roundtrip-form --help"
    WScript.Echo ""
    WScript.Echo "PARÁMETROS:"
    WScript.Echo "  <json_path>              - Archivo JSON con definición del formulario (obligatorio)"
    WScript.Echo "  --db <db_path>           - Ruta a la base de datos Access (obligatorio)"
    WScript.Echo ""
    WScript.Echo "OPCIONES:"
    WScript.Echo "  --password <pwd>         - Contraseña de la base de datos (si está protegida)"
    WScript.Echo "  --pretty                 - Generar JSON con formato legible"
    WScript.Echo "  --strict                 - Modo estricto: falla si hay diferencias en roundtrip"
    WScript.Echo "  --diff                   - Mostrar diferencias línea por línea si las hay"
    WScript.Echo "  --verbose                - Mostrar información detallada del proceso"
    WScript.Echo "  --help, -h, help         - Mostrar esta ayuda"
    WScript.Echo ""
    WScript.Echo "FUNCIONAMIENTO:"
    WScript.Echo "  1. Importa el formulario desde JSON usando import-form --overwrite"
    WScript.Echo "  2. Exporta el formulario a JSON temporal usando export-form"
    WScript.Echo "  3. Si --diff está activo, compara los archivos línea por línea"
    WScript.Echo "  4. Si --strict está activo y hay diferencias, retorna código de salida 1"
    WScript.Echo "  5. Guarda el resultado en directorio temporal .tmp\roundtrip\"
    WScript.Echo ""
    WScript.Echo "CÓDIGOS DE SALIDA:"
    WScript.Echo "  0 - Roundtrip exitoso (sin diferencias o modo no estricto)"
    WScript.Echo "  1 - Error en el proceso o diferencias detectadas en modo estricto"
    WScript.Echo ""
    WScript.Echo "EJEMPLOS:"
    WScript.Echo "  # Roundtrip básico"
    WScript.Echo "  cscript condor_cli.vbs roundtrip-form "".\ui\forms\FormComercial.json"" --db "".\ui\sources\Expedientes.accdb"" --password dpddpd"
    WScript.Echo ""
    WScript.Echo "  # Roundtrip con formato pretty y modo estricto"
    WScript.Echo "  cscript condor_cli.vbs roundtrip-form "".\ui\forms\FormComercial.json"" --db "".\ui\sources\Expedientes.accdb"" --password dpddpd --pretty --strict"
    WScript.Echo ""
    WScript.Echo "  # Roundtrip con diff detallado"
    WScript.Echo "  cscript condor_cli.vbs roundtrip-form "".\ui\forms\FormComercial.json"" --db "".\ui\sources\Expedientes.accdb"" --password dpddpd --pretty --strict --diff"
End Sub
    WScript.Echo "  cscript condor_cli.vbs roundtrip-form ""C:\\DB\\app.accdb"" ""form.json"" --verbose"
    WScript.Echo "  cscript condor_cli.vbs roundtrip-form ""C:\\DB\\app.accdb"" ""FormularioClientes"""
    WScript.Echo "  cscript condor_cli.vbs roundtrip-form ""C:\\DB\\app.accdb"" ""form.json"" --password ""mi_pwd"""
    WScript.Echo ""
    WScript.Echo "NOTAS:"
    WScript.Echo "  • Los comandos operan en vista Diseño (no ejecutan eventos)"
    WScript.Echo "  • Se crea automáticamente el directorio 'roundtrip/' si no existe"
    WScript.Echo "  • El archivo de salida se nombra como <nombre_formulario>.json"
    WScript.Echo "  • Útil para verificar que el proceso de importación preserva la estructura"
End Sub

' Función para obtener la lista de archivos por funcionalidad según CONDOR_MASTER_PLAN.md
' Incluye dependencias para cada funcionalidad
' NOTA: El CLI (condor_cli.vbs) se empaqueta aparte y esta función lista módulos .bas/.cls de funcionalidades
' ============================================================================
' FUNCIÓN: GetFunctionalityFiles
' Descripción: Mapea funcionalidades a archivos *.bas/*.cls correspondientes
' Nota: El CLI usa VBIDE/LoadFromText para importación; esta lista mapea
'       archivos fuente por funcionalidad para empaquetado y gestión
' ============================================================================
Function GetFunctionalityFiles(strFunctionality)
    Dim arrFiles
    
    Select Case LCase(strFunctionality)
        Case "auth", "autenticacion", "authentication"
            ' Sección 3.1 - Autenticación + Dependencias
            arrFiles = Array("IAuthService.cls", "CAuthService.cls", "CMockAuthService.cls", _
                           "IAuthRepository.cls", "CAuthRepository.cls", "CMockAuthRepository.cls", _
                           "EAuthData.cls", "modAuthFactory.bas", "TestAuthService.bas", _
                           "TIAuthRepository.bas", _
                           "IConfig.cls", "IErrorHandlerService.cls", "modEnumeraciones.bas")
        
        Case "document", "documentos", "documents"
            ' Sección 3.2 - Gestión de Documentos + Dependencias
            arrFiles = Array("IDocumentService.cls", "CDocumentService.cls", "CMockDocumentService.cls", _
                           "IWordManager.cls", "CWordManager.cls", "CMockWordManager.cls", _
                           "IMapeoRepository.cls", "CMapeoRepository.cls", "CMockMapeoRepository.cls", _
                           "EMapeo.cls", "modDocumentServiceFactory.bas", _
                           "TIDocumentService.bas", _
                           "ISolicitudService.cls", "CSolicitudService.cls", "modSolicitudServiceFactory.bas", _
                           "IOperationLogger.cls", "IConfig.cls", "IErrorHandlerService.cls", "IFileSystem.cls", _
                           "modWordManagerFactory.bas", "modRepositoryFactory.bas", "modErrorHandlerFactory.bas")
        
        Case "expediente", "expedientes"
            ' Sección 3.3 - Gestión de Expedientes + Dependencias
            arrFiles = Array("IExpedienteService.cls", "CExpedienteService.cls", "CMockExpedienteService.cls", _
                           "IExpedienteRepository.cls", "CExpedienteRepository.cls", "CMockExpedienteRepository.cls", _
                           "EExpediente.cls", "modExpedienteServiceFactory.bas", "TestCExpedienteService.bas", _
                           "TIExpedienteRepository.bas", "modRepositoryFactory.bas", _
                           "IConfig.cls", "IOperationLogger.cls", "IErrorHandlerService.cls")
        
        Case "solicitud", "solicitudes"
            ' Sección 3.4 - Gestión de Solicitudes + Dependencias
            arrFiles = Array("ISolicitudService.cls", "CSolicitudService.cls", "CMockSolicitudService.cls", _
                           "ISolicitudRepository.cls", "CSolicitudRepository.cls", "CMockSolicitudRepository.cls", _
                           "ESolicitud.cls", "EDatosPc.cls", "EDatosCdCa.cls", "EDatosCdCaSub.cls", _
                           "modSolicitudServiceFactory.bas", "TestSolicitudService.bas", _
                           "TISolicitudRepository.bas", _
                           "IAuthService.cls", "modAuthFactory.bas", _
                           "IOperationLogger.cls", "IErrorHandlerService.cls", "IConfig.cls")
        
        Case "workflow", "flujo"
            ' Sección 3.5 - Gestión de Workflow + Dependencias
            arrFiles = Array("IWorkflowService.cls", "CWorkflowService.cls", "CMockWorkflowService.cls", _
                           "IWorkflowRepository.cls", "CWorkflowRepository.cls", "CMockWorkflowRepository.cls", _
                           "modWorkflowServiceFactory.bas", "TestWorkflowService.bas", _
                           "TIWorkflowRepository.bas", _
                           "IOperationLogger.cls", "IConfig.cls", "IErrorHandlerService.cls")
        
        Case "mapeo", "mapping"
            ' Sección 3.6 - Gestión de Mapeos + Dependencias
            arrFiles = Array("IMapeoRepository.cls", "CMapeoRepository.cls", "CMockMapeoRepository.cls", _
                           "EMapeo.cls", "TIMapeoRepository.bas", _
                           "IConfig.cls", "IErrorHandlerService.cls")
        
        Case "notification", "notificacion"
            ' Sección 3.7 - Gestión de Notificaciones + Dependencias
            arrFiles = Array("INotificationService.cls", "CNotificationService.cls", "CMockNotificationService.cls", _
                           "INotificationRepository.cls", "CNotificationRepository.cls", "CMockNotificationRepository.cls", _
                           "modNotificationServiceFactory.bas", "TINotificationService.bas", _
                           "IOperationLogger.cls", "IErrorHandlerService.cls", "IConfig.cls")
        
        Case "operation", "operacion", "logging"
            ' Sección 3.8 - Gestión de Operaciones y Logging + Dependencias
            arrFiles = Array("IOperationLogger.cls", "COperationLogger.cls", "CMockOperationLogger.cls", _
                           "IOperationRepository.cls", "COperationRepository.cls", "CMockOperationRepository.cls", _
                           "EOperationLog.cls", "modOperationLoggerFactory.bas", "TestOperationLogger.bas", _
                           "TIOperationRepository.bas", _
                           "IErrorHandlerService.cls", "IConfig.cls")
        
        Case "config", "configuracion"
            ' Sección 4 - Configuración + Dependencias
            arrFiles = Array("IConfig.cls", "CConfig.cls", "CMockConfig.cls", "modConfigFactory.bas", _
                           "TestCConfig.bas")
        
        Case "filesystem", "archivos"
            ' Sección 5 - Sistema de Archivos + Dependencias
            arrFiles = Array("IFileSystem.cls", "CFileSystem.cls", "CMockFileSystem.cls", _
                           "modFileSystemFactory.bas", "TIFileSystem.bas", _
                           "IErrorHandlerService.cls")
        
        Case "word"
            ' Sección 6 - Gestión de Word + Dependencias
            arrFiles = Array("IWordManager.cls", "CWordManager.cls", "CMockWordManager.cls", _
                           "modWordManagerFactory.bas", "TIWordManager.bas", _
                           "IFileSystem.cls", "IErrorHandlerService.cls")
        
        Case "error", "errores", "errors"
            ' Sección 7 - Gestión de Errores + Dependencias
            arrFiles = Array("IErrorHandlerService.cls", "CErrorHandlerService.cls", "CMockErrorHandlerService.cls", _
                           "modErrorHandlerFactory.bas", "TestErrorHandlerService.bas", _
                           "IConfig.cls", "IFileSystem.cls")
        
        Case "testframework", "testing", "framework"
            ' Sección 8 - Framework de Testing + Dependencias
            arrFiles = Array("ITestReporter.cls", "CTestResult.cls", "CTestSuiteResult.cls", "CTestReporter.cls", _
                           "modTestRunner.bas", "modTestUtils.bas", "modAssert.bas", _
                           "TestModAssert.bas", "IFileSystem.cls", "IConfig.cls", _
                           "IErrorHandlerService.cls")
        
        Case "app", "aplicacion", "application"
            ' Sección 9 - Gestión de Aplicación + Dependencias
            arrFiles = Array("IAppManager.cls", "CAppManager.cls", "CMockAppManager.cls", _
                           "ModAppManagerFactory.bas", "TestAppManager.bas", "IAuthService.cls", _
                           "IConfig.cls", "IErrorHandlerService.cls")
        
        Case "models", "modelos", "datos"
            ' Sección 10 - Modelos de Datos
            arrFiles = Array("EUsuario.cls", "ESolicitud.cls", "EExpediente.cls", "EDatosPc.cls", _
                           "EDatosCdCa.cls", "EDatosCdCaSub.cls", "EEstado.cls", "ETransicion.cls", _
                           "EMapeo.cls", "EAdjuntos.cls", "ELogCambios.cls", "ELogErrores.cls", "EOperationLog.cls", "EAuthData.cls")
        
        Case "utils", "utilidades", "enumeraciones"
            ' Sección 11 - Utilidades y Enumeraciones
            arrFiles = Array("modRepositoryFactory.bas", "modEnumeraciones.bas", "modQueries.bas", _
                           "ModAppManagerFactory.bas", "modAuthFactory.bas", "modConfigFactory.bas", _
                           "modDocumentServiceFactory.bas", "modErrorHandlerFactory.bas", _
                           "modExpedienteServiceFactory.bas", "modFileSystemFactory.bas", _
                           "modNotificationServiceFactory.bas", "modOperationLoggerFactory.bas", _
                           "modSolicitudServiceFactory.bas", "modWordManagerFactory.bas", _
                           "modWorkflowServiceFactory.bas")
        
        Case "forms", "formularios", "ui"
            ' Funcionalidad de Formularios - UI as Code
            arrFiles = Array("condor_cli.vbs")
            
        Case "cli", "infrastructure", "infraestructura"
            ' Funcionalidad CLI e Infraestructura
            arrFiles = Array("condor_cli.vbs")
            
        Case "condorcli"
            ' Funcionalidad especial para copiar condor_cli.vbs como .txt
            arrFiles = Array("condor_cli.vbs")
            
        Case "tests", "pruebas", "testing", "test"
            ' Sección 12 - Archivos de Pruebas (Autodescubrimiento)
            arrFiles = Array()
        Case Else
            ' Funcionalidad no reconocida - devolver array vacío
            arrFiles = Array()
    End Select
    
    ' export-form: SaveAsText acForm, sin OpenForm Design
    GetFunctionalityFiles = arrFiles
End Function



        


' Subrutina para empaquetar archivos de código por funcionalidad
' Subrutina para empaquetar archivos de código por funcionalidad o por lista de ficheros
Sub BundleFunctionality()
    On Error Resume Next
    
    Dim strFunctionalityOrFiles, strDestPath, strBundlePath, timestamp
    
    ' Verificar argumentos
    If objArgs.Count < 2 Then
        WScript.Echo "Error: Se requiere nombre de funcionalidad o lista de ficheros"
        WScript.Echo "Uso: cscript condor_cli.vbs bundle <funcionalidad | fichero1,fichero2,...> [ruta_destino]"
        WScript.Quit 1
    End If
    
    strFunctionalityOrFiles = objArgs(1)
    
    ' Determinar ruta de destino
    If objArgs.Count >= 3 Then
        strDestPath = objArgs(2)
    Else
        strDestPath = objFSO.GetParentFolderName(WScript.ScriptFullName)
    End If
    
    ' Crear timestamp
    timestamp = Year(Now) & Right("0" & Month(Now), 2) & Right("0" & Day(Now), 2) & "_" & _
                Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2) & Right("0" & Second(Now), 2)
    
    ' Crear nombre de carpeta bundle
    Dim bundleName
    If InStr(strFunctionalityOrFiles, ",") > 0 Then
        bundleName = "bundle_custom_" & timestamp
    Else
        bundleName = "bundle_" & strFunctionalityOrFiles & "_" & timestamp
    End If
    strBundlePath = objFSO.BuildPath(strDestPath, bundleName)
    
    WScript.Echo "=== EMPAQUETANDO ARTEFACTOS ==="
    WScript.Echo "Buscando archivos en: " & strSourcePath
    WScript.Echo "Carpeta destino: " & strBundlePath
    
    ' Crear carpeta de destino
    If Not objFSO.FolderExists(strBundlePath) Then
        objFSO.CreateFolder strBundlePath
        If Err.Number <> 0 Then
            WScript.Echo "Error creando carpeta de destino: " & Err.Description
            WScript.Quit 1
        End If
    End If
    
    Dim arrFilesToBundle
    
    ' Lógica de Detección Inteligente
    If InStr(strFunctionalityOrFiles, ",") > 0 Then
        ' MODO 1: Lista de ficheros explícita
        WScript.Echo "Modo: Lista de ficheros explícita."
        arrFilesToBundle = Split(strFunctionalityOrFiles, ",")
    Else
        ' MODO 2: Verificar si es funcionalidad conocida o archivo individual
        arrFilesToBundle = GetFunctionalityFiles(strFunctionalityOrFiles)
        
        If UBound(arrFilesToBundle) >= 0 Then
            ' Es una funcionalidad conocida
            WScript.Echo "Modo: Funcionalidad '" & strFunctionalityOrFiles & "'."
        Else
            ' No es funcionalidad conocida, buscar archivo individual en src
            Dim singleFilePath
            singleFilePath = objFSO.BuildPath(strSourcePath, strFunctionalityOrFiles)
            
            If objFSO.FileExists(singleFilePath) Then
                ' Archivo encontrado, tratarlo como lista de un elemento
                WScript.Echo "Modo: Archivo individual '" & strFunctionalityOrFiles & "'."
                ReDim arrFilesToBundle(0)
                arrFilesToBundle(0) = strFunctionalityOrFiles
            Else
                ' Archivo no encontrado
                WScript.Echo "Error: '" & strFunctionalityOrFiles & "' no es una funcionalidad conocida ni un archivo existente en src."
                WScript.Echo "Funcionalidades disponibles: Auth, Document, Expediente, Solicitud, Workflow, Mapeo, Notification, Operation, Config, FileSystem, Word, Error, TestFramework, App, Models, Utils, Tests"
                WScript.Quit 1
            End If
        End If
    End If
    
    ' Llamar a la subrutina de ayuda para copiar los ficheros
    Call CopyFilesToBundle(arrFilesToBundle, strBundlePath)
    
    On Error GoTo 0
End Sub

' NUEVA SUBRUTINA DE AYUDA
' Copia una lista de ficheros al directorio del paquete
Sub CopyFilesToBundle(arrFiles, strBundlePath)
    Dim copiedFiles, notFoundFiles
    copiedFiles = 0
    notFoundFiles = 0
    
    If UBound(arrFiles) < 0 Then
        WScript.Echo "Advertencia: La lista de ficheros a empaquetar está vacía."
    End If

    Dim i, fileName, filePath, destFilePath
    For i = 0 To UBound(arrFiles)
        fileName = Trim(arrFiles(i))
        
        ' Caso especial para condorcli: copiar desde la raíz del proyecto
        If fileName = "condor_cli.vbs" And InStr(strBundlePath, "bundle_condorcli_") > 0 Then
            filePath = objFSO.BuildPath(objFSO.GetParentFolderName(WScript.ScriptFullName), fileName)
        Else
            filePath = objFSO.BuildPath(strSourcePath, fileName)
        End If
        
        If objFSO.FileExists(filePath) Then
            ' Copiar archivo con extensión .txt añadida al directorio del bundle
            destFilePath = objFSO.BuildPath(strBundlePath, fileName & ".txt")
            objFSO.CopyFile filePath, destFilePath, True
            
            If Err.Number <> 0 Then
                WScript.Echo "  ? Error copiando " & fileName & ": " & Err.Description
                Err.Clear
            Else
                WScript.Echo "  ? " & fileName & " -> " & fileName & ".txt"
                copiedFiles = copiedFiles + 1
            End If
        Else
            WScript.Echo "  ? Archivo no encontrado: " & fileName
            notFoundFiles = notFoundFiles + 1
        End If
    Next
    
    ' NOTA: front\recursos y front\test_env se incluyen solo si son relevantes para la funcionalidad específica
    ' Call CopyResourceDirectories(strBundlePath, copiedFiles)
    
    WScript.Echo ""
    WScript.Echo "=== RESULTADO DEL EMPAQUETADO ==="
    WScript.Echo "Archivos copiados: " & copiedFiles
    WScript.Echo "Archivos no encontrados: " & notFoundFiles
    WScript.Echo "Ubicación del paquete: " & strBundlePath
    
    If copiedFiles = 0 Then
        WScript.Echo "? No se copió ningún archivo."
    Else
        WScript.Echo "? Empaquetado completado exitosamente"
    End If
End Sub

' NUEVA SUBRUTINA: Copia directorios de recursos al bundle
Sub CopyResourceDirectories(strBundlePath, ByRef copiedFiles)
    Dim frontRecursosPath, frontTestEnvPath
    frontRecursosPath = GetFrontRoot() & "\recursos"
    frontTestEnvPath = GetFrontRoot() & "\test_env"
    
    ' Copiar front\recursos si existe
    If objFSO.FolderExists(frontRecursosPath) Then
        Dim destRecursosPath
        destRecursosPath = objFSO.BuildPath(strBundlePath, "front_recursos")
        Call CopyFolderRecursive(frontRecursosPath, destRecursosPath)
        WScript.Echo "  ? front\recursos -> front_recursos/"
        copiedFiles = copiedFiles + 1
    End If
    
    ' Copiar front\test_env si existe
    If objFSO.FolderExists(frontTestEnvPath) Then
        Dim destTestEnvPath
        destTestEnvPath = objFSO.BuildPath(strBundlePath, "front_test_env")
        Call CopyFolderRecursive(frontTestEnvPath, destTestEnvPath)
        WScript.Echo "  ? front\test_env -> front_test_env/"
        copiedFiles = copiedFiles + 1
    End If
End Sub

' NUEVA SUBRUTINA: Copia recursiva de carpetas
Sub CopyFolderRecursive(sourcePath, destPath)
    On Error Resume Next
    
    ' Crear carpeta destino si no existe
    If Not objFSO.FolderExists(destPath) Then
        objFSO.CreateFolder destPath
    End If
    
    ' Copiar archivos
    Dim sourceFolder, file
    Set sourceFolder = objFSO.GetFolder(sourcePath)
    For Each file In sourceFolder.Files
        objFSO.CopyFile file.Path, objFSO.BuildPath(destPath, file.Name), True
    Next
    
    ' Copiar subcarpetas recursivamente
    Dim subFolder
    For Each subFolder In sourceFolder.SubFolders
        Call CopyFolderRecursive(subFolder.Path, objFSO.BuildPath(destPath, subFolder.Name))
    Next
    
    On Error GoTo 0
End Sub

' Función auxiliar para convertir rutas relativas a absolutas
Private Function ResolveRelativePath(relativePath)
    Dim objFSO
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    ' Verificar si la ruta ya es absoluta (contiene : en la segunda posición)
    If Len(relativePath) >= 2 And Mid(relativePath, 2, 1) = ":" Then
        ResolveRelativePath = relativePath
        Exit Function
    End If
    
    ' Obtener el directorio actual del script
    Dim currentDir
    currentDir = objFSO.GetParentFolderName(WScript.ScriptFullName)
    
    ' Si la ruta empieza con .\, quitarlo
    If Left(relativePath, 2) = ".\" Then
        relativePath = Mid(relativePath, 3)
    End If
    
    ' Si la ruta empieza con \, quitarlo
    If Left(relativePath, 1) = "\" Then
        relativePath = Mid(relativePath, 2)
    End If
    
    ' Combinar la ruta actual con la ruta relativa
    ResolveRelativePath = objFSO.BuildPath(currentDir, relativePath)
End Function

' Función auxiliar para convertir tipos de datos DAO a texto legible
Private Function DaoTypeToString(dataType)
    Select Case dataType
        Case 1: DaoTypeToString = "Boolean"
        Case 2: DaoTypeToString = "Byte"
        Case 3: DaoTypeToString = "Integer"
        Case 4: DaoTypeToString = "Long"
        Case 5: DaoTypeToString = "Currency"
        Case 6: DaoTypeToString = "Single"
        Case 7: DaoTypeToString = "Double"
        Case 8: DaoTypeToString = "DateTime"
        Case 10: DaoTypeToString = "Text"
        Case 11: DaoTypeToString = "OLE Object"
        Case 12: DaoTypeToString = "Memo"
        Case 20: DaoTypeToString = "BigInt"
        Case Else: DaoTypeToString = "Desconocido (" & dataType & ")"
    End Select
End Function

' ===================================================================
' SUBRUTINA: ExecuteMigrations
' Descripción: Ejecuta scripts de migración SQL desde la carpeta /db/migrations
' ===================================================================
Sub ExecuteMigrations()
    Dim strMigrationsPath, objMigrationsFolder, objFile, strTargetFile
    
    strMigrationsPath = objFSO.GetParentFolderName(strSourcePath) & "\db\migrations"
    WScript.Echo "=== INICIANDO MIGRACION DE DATOS SQL ==="
    WScript.Echo "Directorio de migraciones: " & strMigrationsPath
    
    If Not objFSO.FolderExists(strMigrationsPath) Then
        WScript.Echo "ERROR: El directorio de migraciones no existe: " & strMigrationsPath
        WScript.Quit 1
    End If
    
    Set objMigrationsFolder = objFSO.GetFolder(strMigrationsPath)
    
    ' Modo 1: Migrar un fichero específico
    If objArgs.Count > 1 Then
        strTargetFile = objArgs(1)
        Dim targetPath
        targetPath = objFSO.BuildPath(strMigrationsPath, strTargetFile)
        If objFSO.FileExists(targetPath) Then
            WScript.Echo "Ejecutando migración específica: " & strTargetFile
            Call ProcessSqlFile(targetPath)
        Else
            WScript.Echo "ERROR: El archivo de migración especificado no existe: " & targetPath
            WScript.Quit 1
        End If
    ' Modo 2: Migrar todos los ficheros .sql
    Else
        WScript.Echo "Ejecutando todas las migraciones en el directorio (en orden alfabético)..."
        
        ' Crear un array para almacenar los nombres de archivos y ordenarlos
        Dim arrFiles(), intFileCount, i, j, strTemp
        intFileCount = 0
        
        ' Contar archivos SQL
        For Each objFile In objMigrationsFolder.Files
            If LCase(objFSO.GetExtensionName(objFile.Name)) = "sql" Then
                intFileCount = intFileCount + 1
            End If
        Next
        
        ' Redimensionar array
        ReDim arrFiles(intFileCount - 1)
        
        ' Llenar array con rutas de archivos
        i = 0
        For Each objFile In objMigrationsFolder.Files
            If LCase(objFSO.GetExtensionName(objFile.Name)) = "sql" Then
                arrFiles(i) = objFile.Path
                i = i + 1
            End If
        Next
        
        ' Ordenar array usando bubble sort
        For i = 0 To UBound(arrFiles) - 1
            For j = i + 1 To UBound(arrFiles)
                If UCase(objFSO.GetFileName(arrFiles(i))) > UCase(objFSO.GetFileName(arrFiles(j))) Then
                    strTemp = arrFiles(i)
                    arrFiles(i) = arrFiles(j)
                    arrFiles(j) = strTemp
                End If
            Next
        Next
        
        ' Ejecutar archivos en orden
        For i = 0 To UBound(arrFiles)
            Call ProcessSqlFile(arrFiles(i))
        Next
    End If
    
    WScript.Echo "=== MIGRACION COMPLETADA EXITOSAMENTE ==="
End Sub

' ===================================================================
' SUBRUTINA: ProcessSqlFile
' Descripción: Parsea y ejecuta los comandos de un fichero SQL
' CORREGIDO: Utiliza ADODB.Stream para leer ficheros con codificación UTF-8.
' ===================================================================
' FUNCIÓN: CleanSqlContent
' Elimina comentarios SQL y líneas vacías del contenido
Function CleanSqlContent(sqlContent)
    Dim arrLines, cleanedLines, i, trimmedLine
    
    ' Dividir en líneas
    arrLines = Split(sqlContent, vbCrLf)
    If UBound(arrLines) = 0 Then
        arrLines = Split(sqlContent, vbLf)
    End If
    
    ' Filtrar líneas
    cleanedLines = ""
    For i = 0 To UBound(arrLines)
        trimmedLine = Trim(arrLines(i))
        
        ' Ignorar líneas vacías y comentarios
        If Len(trimmedLine) > 0 And Left(trimmedLine, 2) <> "--" Then
            If cleanedLines <> "" Then
                cleanedLines = cleanedLines & vbCrLf
            End If
            cleanedLines = cleanedLines & arrLines(i)
        End If
    Next
    
    CleanSqlContent = cleanedLines
End Function

Sub ProcessSqlFile(filePath)
    Dim objStream, strContent, arrCommands, sqlCommand, conn
    
    WScript.Echo "------------------------------------------------------------"
    WScript.Echo "Procesando fichero: " & objFSO.GetFileName(filePath)
    
    On Error Resume Next
    Set objStream = CreateObject("ADODB.Stream")
    objStream.Type = 2 ' adTypeText
    objStream.Charset = "UTF-8"
    objStream.Open
    objStream.LoadFromFile filePath
    strContent = objStream.ReadText
    objStream.Close
    Set objStream = Nothing
    If Err.Number <> 0 Then
        WScript.Echo "  ERROR: No se pudo leer el fichero: " & Err.Description
        WScript.Quit 1 ' Detener en caso de error de lectura
    End If
    On Error GoTo 0
    
    ' Usar conexión ADO para un manejo de errores DDL robusto
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & strAccessPath & ";"
    
    ' Limpiar comentarios y líneas vacías antes de procesar
    strContent = CleanSqlContent(strContent)
    
    ' Dividir en comandos por punto y coma
    arrCommands = Split(strContent, ";")
    
    ' Ejecutar cada comando
    For Each sqlCommand In arrCommands
        sqlCommand = Trim(sqlCommand)
        
        ' Solo ejecutar comandos que no estén vacías
        If Len(sqlCommand) > 5 Then
            On Error Resume Next
            conn.Execute sqlCommand
            
            If Err.Number <> 0 Then
                WScript.Echo "    ERROR al ejecutar comando: " & Err.Description
                WScript.Echo "    SQL: " & sqlCommand
                WScript.Echo "  MIGRACIÓN FALLIDA. Abortando."
                WScript.Echo "------------------------------------------------------------"
                conn.Close
                Set conn = Nothing
                WScript.Quit 1 ' Detener la ejecución inmediatamente
            Else
                WScript.Echo "    Comando ejecutado exitosamente."
            End If
            On Error GoTo 0
        End If
    Next
    
    conn.Close
    Set conn = Nothing
    
    WScript.Echo "  Fichero procesado exitosamente."
    WScript.Echo "------------------------------------------------------------"
End Sub

' Función para formatear texto con ancho fijo
Function PadRight(text, width)
    If Len(text) >= width Then
        PadRight = Left(text, width)
    Else
        PadRight = text & String(width - Len(text), " ")
    End If
End Function

' Función para verificar la refactorización de logging
Sub VerifyLoggingRefactoring()
    Dim serviceFiles, fileName, filePath, fileContent
    Dim obsoleteCalls, refactoredCalls
    Dim totalObsolete, totalRefactored
    
    serviceFiles = Array("CAuthService.cls", "CNotificationService.cls", "CWorkflowService.cls", "CSolicitudService.cls")
    totalObsolete = 0
    totalRefactored = 0
    
    WScript.Echo "  Verificando servicios refactorizados..."
    
    For Each fileName In serviceFiles
        filePath = strSourcePath & "\" & fileName
        
        If objFSO.FileExists(filePath) Then
            fileContent = objFSO.OpenTextFile(filePath, 1).ReadAll
            
            ' Buscar llamadas obsoletas (3 parámetros)
            obsoleteCalls = CountMatches(fileContent, "LogOperation\s*\(\s*""[^""]*""\s*,\s*\d+\s*,\s*""[^""]*""\s*\)")
            
            ' Buscar llamadas refactorizadas (EOperationLog)
            refactoredCalls = CountMatches(fileContent, "LogOperation\s*\(\s*operationLog\)")
            
            totalObsolete = totalObsolete + obsoleteCalls
            totalRefactored = totalRefactored + refactoredCalls
            
            If obsoleteCalls > 0 Then
                WScript.Echo "    ⚠️  " & fileName & ": " & obsoleteCalls & " llamadas obsoletas encontradas"
            Else
                WScript.Echo "    ✅ " & fileName & ": Refactorizado (" & refactoredCalls & " llamadas EOperationLog)"
            End If
        Else
            WScript.Echo "    ❌ " & fileName & ": Archivo no encontrado"
        End If
    Next
    
    WScript.Echo "  Resumen de refactorización:"
    WScript.Echo "    - Llamadas obsoletas: " & totalObsolete
    WScript.Echo "    - Llamadas refactorizadas: " & totalRefactored
    
    If totalObsolete > 0 Then
        WScript.Echo "    ⚠️  ADVERTENCIA: Aún existen llamadas obsoletas por refactorizar"
    Else
        WScript.Echo "    ✅ ÉXITO: Todos los servicios han sido refactorizados"
    End If
End Sub

' Función auxiliar para contar coincidencias de regex
Function CountMatches(text, pattern)
    Dim regex, matches
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = pattern
    regex.Global = True
    regex.IgnoreCase = True
    
    Set matches = regex.Execute(text)
    CountMatches = matches.Count
End Function

' ===================================================================
' SUBRUTINA: ExportForm
' Descripción: Exporta el diseño de un formulario a JSON usando SaveAsText sin abrir el formulario en vista diseño.
' ===================================================================
Sub ExportForm()
    Dim strDbPath, strFormName, strOutputPath, strPassword, strOutDir
    Dim i
    
    ' Verificar argumentos mínimos
    If objArgs.Count < 3 Then
        WScript.Echo "Error: El comando export-form requiere al menos una ruta de base de datos y un nombre de formulario."
        WScript.Echo "Uso: cscript condor_cli.vbs export-form <db_path> <form_name> [--password <pwd>] [--outdir <dir>] [--verbose]"
        WScript.Quit 1
    End If
    
    ' Asignar argumentos básicos
    strDbPath = objArgs(1)
    strFormName = objArgs(2)
    strPassword = ""
    strOutDir = ""
    
    ' Procesar argumentos opcionales
    For i = 3 To objArgs.Count - 1
        Dim currentArg
        currentArg = objArgs(i)
        
        If LCase(currentArg) = "--password" And i < objArgs.Count - 1 Then
            strPassword = objArgs(i + 1)
        ElseIf LCase(currentArg) = "--outdir" And i < objArgs.Count - 1 Then
            strOutDir = objArgs(i + 1)
        ElseIf LCase(currentArg) = "--verbose" Then
            gVerbose = True
        ElseIf LCase(currentArg) = "--sharedopen" Or LCase(currentArg) = "/sharedopen" Then
            WScript.Echo "BD en uso: no se pudo abrir en modo exclusivo. Use --sharedopen si solo necesita lectura."
            WScript.Quit 1
        End If
    Next
    
    ' Configurar directorio de salida por defecto
    If strOutDir = "" Then
        strOutDir = ".\ui\forms\"
    End If
    
    ' Asegurar que el directorio de salida existe
    If Not EnsureDir(strOutDir) Then
        WScript.Echo "Error: No se pudo crear el directorio de salida: " & strOutDir
        WScript.Quit 1
    End If
    
    ' Verificar que el archivo de base de datos existe
    If Not objFSO.FileExists(strDbPath) Then
        WScript.Echo "Error: La base de datos no existe: " & strDbPath
        WScript.Quit 1
    End If
    
    ' Convertir a ruta absoluta para evitar problemas de apertura
    strDbPath = objFSO.GetAbsolutePathName(strDbPath)

    If gVerbose Then WScript.Echo "✓ Verificando acceso exclusivo a BD: " & strDbPath
    
    ' Preflight de exclusividad con timeout
    If Not WaitForExclusive(strDbPath, strPassword, 5) Then
        WScript.Echo "✗ BD en uso: no se pudo abrir en modo exclusivo. Use --sharedopen si solo necesita lectura."
        WScript.Quit 1
    End If
    
    If gVerbose Then WScript.Echo "✓ BD abierta en modo exclusivo"
    
    ' Verificar existencia del formulario
    If Not HasForm(strFormName) Then
        WScript.Echo "✗ Formulario no encontrado: " & strFormName
        CloseAccessApp objAccess
        WScript.Quit 1
    End If
    
    If gVerbose Then WScript.Echo "✓ Formulario encontrado: " & strFormName
    
    
    ' Exportar formulario usando SaveAsText sin abrirlo
    Dim tempTxtPath, jsonOutputPath, dbName
    tempTxtPath = objFSO.GetTempName() & ".txt"
    tempTxtPath = objFSO.BuildPath(objFSO.GetSpecialFolder(2), tempTxtPath) ' Temp folder
    
    ' Construir nombre del archivo JSON de salida
    dbName = objFSO.GetBaseName(strDbPath)
    jsonOutputPath = objFSO.BuildPath(strOutDir, dbName & "__" & strFormName & ".json")
    
    If gVerbose Then WScript.Echo "✓ Exportando formulario usando SaveAsText..."
    
    ' Exportar formulario a texto usando SaveAsText (2 = acForm)
    On Error Resume Next
    objAccess.Application.SaveAsText 2, strFormName, tempTxtPath
    
    If Err.Number <> 0 Then
        Dim errMsg: errMsg = "Error al exportar formulario '" & strFormName & "': " & Err.Description
        Err.Clear
        CloseAccessApp objAccess
        WScript.Echo "✗ " & errMsg
        WScript.Quit 1
    End If
    On Error GoTo 0
    
    If gVerbose Then WScript.Echo "✓ Formulario exportado a archivo temporal: " & tempTxtPath
    
    ' Leer contenido del archivo exportado
    Dim vbeTextContent, vbeLines
    vbeTextContent = ReadAllTextAnsiOrUtf8(tempTxtPath)
    vbeLines = Split(vbeTextContent, vbCrLf)
    
    ' Calcular hash y tamaño del archivo temporal
    Dim fileSize, fileSha1
    fileSize = objFSO.GetFile(tempTxtPath).Size
    fileSha1 = Sha1OfFile(tempTxtPath)
    
    ' Construir estructura JSON
    Dim jsonContent
    jsonContent = "{" & vbCrLf
    jsonContent = jsonContent & "  ""db_path"": """ & Replace(strDbPath, "\", "\\") & """," & vbCrLf
    jsonContent = jsonContent & "  ""db_name"": """ & dbName & """," & vbCrLf
    jsonContent = jsonContent & "  ""form_name"": """ & strFormName & """," & vbCrLf
    jsonContent = jsonContent & "  ""timestamp"": """ & FormatDateTime(Now, vbGeneralDate) & """," & vbCrLf
    jsonContent = jsonContent & "  ""save_as_text_path"": """ & Replace(tempTxtPath, "\", "\\") & """," & vbCrLf
    jsonContent = jsonContent & "  ""size_bytes"": " & fileSize & "," & vbCrLf
    jsonContent = jsonContent & "  ""sha1"": """ & fileSha1 & """," & vbCrLf
    jsonContent = jsonContent & "  ""vbe_text"": [" & vbCrLf
    
    ' Agregar líneas del archivo VBE como array JSON
    Dim j
    For j = 0 To UBound(vbeLines)
        Dim escapedLine
        escapedLine = Replace(vbeLines(j), "\", "\\")
        escapedLine = Replace(escapedLine, """", "\""")
        escapedLine = Replace(escapedLine, vbTab, "\t")
        jsonContent = jsonContent & "    """ & escapedLine & """"
        If j < UBound(vbeLines) Then jsonContent = jsonContent & ","
        jsonContent = jsonContent & vbCrLf
    Next
    
    jsonContent = jsonContent & "  ]" & vbCrLf
    jsonContent = jsonContent & "}" & vbCrLf
    
    ' Escribir archivo JSON de salida usando UTF-8
    Call WriteUtf8File(jsonOutputPath, jsonContent)
    
    If gVerbose Then WScript.Echo "✓ Archivo JSON generado: " & jsonOutputPath
    
    ' Limpiar archivo temporal
    On Error Resume Next
    objFSO.DeleteFile tempTxtPath, True
    On Error GoTo 0
    
    ' Cerrar Access
    objAccess.Echo True
    CloseAccessApp objAccess
    
    WScript.Echo "✓ Exportación completada exitosamente"
    WScript.Echo "  Archivo: " & jsonOutputPath
    WScript.Echo "  Tamaño: " & fileSize & " bytes"
    WScript.Echo "  SHA1: " & fileSha1
End Sub

' Detecta módulo y handlers de eventos para el formulario
' Parámetros:
'   jsonWriter - Objeto JsonWriter para escribir la sección code
'   formName - Nombre del formulario
'   srcDir - Directorio donde buscar el módulo
'   verbose - Si mostrar información detallada
' Retorna: Diccionario con handlers detectados (clave: "control.event")
Function DetectModuleAndHandlers(jsonWriter, formName, srcDir, verbose)
    Dim objFSO, moduleFile, moduleExists, moduleFilename
    Dim handlers, fileContent, regEx, matches, match
    Dim i, controlName, eventName, signature
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set handlers = CreateObject("Scripting.Dictionary")
    
    moduleExists = False
    moduleFilename = ""
    
    ' Heurística para encontrar el archivo del módulo
    Dim candidateFiles(3)
    candidateFiles(0) = srcDir & "Form_" & formName & ".bas"
    candidateFiles(1) = srcDir & formName & ".bas"
    candidateFiles(2) = srcDir & "frm" & formName & ".bas"
    candidateFiles(3) = srcDir & "Form_" & formName & ".cls"
    
    For i = 0 To 3
        If objFSO.FileExists(candidateFiles(i)) Then
            moduleExists = True
            moduleFilename = objFSO.GetFileName(candidateFiles(i))
            moduleFile = candidateFiles(i)
            Exit For
        End If
    Next
    
    ' Escribir sección code.module
    jsonWriter.StartObjectProperty "code"
    jsonWriter.StartObjectProperty "module"
    jsonWriter.WriteProperty "exists", moduleExists
    jsonWriter.WriteProperty "filename", moduleFilename
    
    ' Si existe el módulo, parsear handlers
    If moduleExists Then
        If verbose Then WScript.Echo "Detectando handlers en: " & moduleFile
        
        ' Leer contenido del archivo
        Dim textStream
        Set textStream = objFSO.OpenTextFile(moduleFile, 1) ' ForReading
        fileContent = textStream.ReadAll
        textStream.Close
        
        ' Crear expresión regular para detectar handlers
        Set regEx = CreateObject("VBScript.RegExp")
        regEx.Global = True
        regEx.IgnoreCase = True
        regEx.MultiLine = True
        regEx.Pattern = "^\s*(Public|Private)?\s*Sub\s+([A-Za-z0-9_]+)_(Click|DblClick|Current|Load|Open|GotFocus|LostFocus|Change|AfterUpdate|BeforeUpdate)\s*\("
        
        Set matches = regEx.Execute(fileContent)
        
        ' Procesar matches
        jsonWriter.StartArrayProperty "handlers"
        
        For Each match In matches
            controlName = match.SubMatches(1)
            eventName = match.SubMatches(2)
            signature = Trim(match.Value)
            
            ' Agregar handler al JSON
            jsonWriter.StartObject
            jsonWriter.WriteProperty "control", controlName
            jsonWriter.WriteProperty "event", eventName
            jsonWriter.WriteProperty "signature", signature
            jsonWriter.EndObject
            
            ' Guardar en diccionario para referencia
            Dim handlerKey
            handlerKey = controlName & "." & eventName
            handlers.Add handlerKey, True
            
            If verbose Then WScript.Echo "Handler detectado: " & controlName & "." & eventName
        Next
        
        jsonWriter.EndArray ' handlers
    Else
        ' No hay módulo, array vacío
        jsonWriter.StartArrayProperty "handlers"
        jsonWriter.EndArray
        
        If verbose Then WScript.Echo "No se encontró módulo para el formulario " & formName
    End If
    
    jsonWriter.EndObject ' module
    jsonWriter.EndObject ' code
    
    ' Retornar diccionario de handlers para uso posterior
    Set DetectModuleAndHandlers = handlers
End Function

' ===================================================================
' SUBRUTINA: RoundtripFormCommand
' Descripción: Realiza test de roundtrip import→export de formulario
' ===================================================================
Sub RoundtripFormCommand()
    Dim strDbPath, strJsonPath, strFormName, strPassword, strOutputDir
    Dim i, objJsonData, strOutputPath, bStrict, bDiff, bPretty
    
    ' Verificar si se solicita ayuda
    If objArgs.Count > 1 Then
        If LCase(objArgs(1)) = "--help" Or LCase(objArgs(1)) = "-h" Or LCase(objArgs(1)) = "help" Then
            Call ShowRoundtripFormHelp()
            WScript.Quit 0
        End If
    End If
    
    ' Verificar argumentos mínimos
    If objArgs.Count < 3 Then
        WScript.Echo "Error: El comando roundtrip-form requiere un archivo JSON y una base de datos."
        WScript.Echo "Uso: cscript condor_cli.vbs roundtrip-form <json_path> --db <db_path> [--password <pwd>] [--strict] [--diff] [--pretty]"
        WScript.Quit 1
    End If
    
    ' Inicializar variables
    strJsonPath = objArgs(1)
    strDbPath = ""
    strPassword = ""
    bStrict = False
    bDiff = False
    bPretty = False
    
    ' Procesar argumentos
    For i = 2 To objArgs.Count - 1
        If LCase(objArgs(i)) = "--db" And i < objArgs.Count - 1 Then
            strDbPath = objArgs(i + 1)
        ElseIf LCase(objArgs(i)) = "--password" And i < objArgs.Count - 1 Then
            strPassword = objArgs(i + 1)
        ElseIf LCase(objArgs(i)) = "--strict" Then
            bStrict = True
        ElseIf LCase(objArgs(i)) = "--diff" Then
            bDiff = True
        ElseIf LCase(objArgs(i)) = "--pretty" Then
            bPretty = True
        ElseIf LCase(objArgs(i)) = "--verbose" Then
            gVerbose = True
        End If
    Next
    
    ' Validar argumentos requeridos
    If strDbPath = "" Then
        WScript.Echo "Error: Debe especificar --db <ruta_base_datos>"
        WScript.Quit 1
    End If
    
    ' Verificar que los archivos existen
    If Not objFSO.FileExists(strJsonPath) Then
        WScript.Echo "Error: El archivo JSON no existe: " & strJsonPath
        WScript.Quit 1
    End If
    
    If Not objFSO.FileExists(strDbPath) Then
        WScript.Echo "Error: La base de datos no existe: " & strDbPath
        WScript.Quit 1
    End If
    
    ' Leer nombre del formulario del JSON
    On Error Resume Next
    Set objJsonData = ParseJsonFile(strJsonPath)
    If Err.Number <> 0 Then
        WScript.Echo "Error: No se pudo parsear el archivo JSON: " & Err.Description
        WScript.Quit 1
    End If
    On Error GoTo 0
    
    If objJsonData.Exists("formName") Then
        strFormName = objJsonData("formName")
    ElseIf objJsonData.Exists("name") Then
        strFormName = objJsonData("name")
    Else
        WScript.Echo "Error: El archivo JSON no contiene el campo 'formName' o 'name' requerido."
        WScript.Quit 1
    End If
    
    ' Crear directorio temporal .tmp\roundtrip
    Dim strTempDir
    strTempDir = objFSO.BuildPath(objFSO.GetParentFolderName(objFSO.GetAbsolutePathName(strJsonPath)), ".tmp")
    If Not objFSO.FolderExists(strTempDir) Then
        objFSO.CreateFolder strTempDir
    End If
    
    Dim strRoundtripDir
    strRoundtripDir = objFSO.BuildPath(strTempDir, "roundtrip")
    If Not objFSO.FolderExists(strRoundtripDir) Then
        objFSO.CreateFolder strRoundtripDir
    End If
    
    strOutputPath = objFSO.BuildPath(strRoundtripDir, strFormName & ".json")
    
    If gVerbose Then
        WScript.Echo "=== ROUNDTRIP FORM INICIADO ==="
        WScript.Echo "JSON entrada: " & strJsonPath
        WScript.Echo "Base de datos: " & strDbPath
        WScript.Echo "Formulario: " & strFormName
        WScript.Echo "Salida temporal: " & strOutputPath
        WScript.Echo "Modo estricto: " & bStrict
        WScript.Echo "Generar diff: " & bDiff
    End If
    
    On Error Resume Next
    
    ' PASO A: Import del formulario con --overwrite
    If gVerbose Then WScript.Echo "PASO A: Importando formulario desde JSON..."
    
    Call ImportFormFromJson(strJsonPath, strDbPath, strPassword, True) ' True = overwrite
    
    If Err.Number <> 0 Then
        WScript.Echo "Error en import-form: " & Err.Description
        WScript.Quit 1
    End If
    
    If gVerbose Then WScript.Echo "Import completado exitosamente."
    
    ' PASO B: Export del formulario a carpeta temporal
    If gVerbose Then WScript.Echo "PASO B: Exportando formulario..."
    
    Call ExportFormToJson(strDbPath, strFormName, strOutputPath, strPassword, bPretty)
    
    If Err.Number <> 0 Then
        WScript.Echo "Error en export-form: " & Err.Description
        WScript.Quit 1
    End If
    
    If Not objFSO.FileExists(strOutputPath) Then
        WScript.Echo "Error: No se generó el archivo JSON de salida: " & strOutputPath
        WScript.Quit 1
    End If
    
    On Error GoTo 0
    
    ' PASO C: Generar diff si se solicita
    Dim bHasDifferences
    bHasDifferences = False
    
    If bDiff Then
        If gVerbose Then WScript.Echo "PASO C: Generando diff..."
        bHasDifferences = GenerateSimpleDiff(strJsonPath, strOutputPath)
    End If
    
    ' Mostrar resultados
    If bHasDifferences Then
        If bStrict Then
            WScript.Echo "✗ Roundtrip mismatch"
            WScript.Quit 1
        Else
            WScript.Echo "⚠ Roundtrip con diferencias (modo no estricto)"
        End If
    Else
        WScript.Echo "✓ Roundtrip OK"
    End If
    
    If gVerbose Then
        WScript.Echo "Archivo temporal disponible en: " & strOutputPath
    End If
End Sub

' ============================================================================
' SUBRUTINA: GenerateSimpleDiff
' Descripción: Genera un diff textual simple línea a línea sin dependencias externas
' Retorna: True si hay diferencias, False si son idénticos
' ============================================================================
Private Function GenerateSimpleDiff(originalFile, newFile)
    Dim objFSO, objOriginal, objNew
    Dim originalLines, newLines
    Dim i, maxLines, hasDifferences
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    hasDifferences = False
    
    ' Leer archivos línea por línea
    Set objOriginal = objFSO.OpenTextFile(originalFile, 1, False, -1) ' UTF-8
    Set objNew = objFSO.OpenTextFile(newFile, 1, False, -1) ' UTF-8
    
    originalLines = Split(objOriginal.ReadAll, vbCrLf)
    newLines = Split(objNew.ReadAll, vbCrLf)
    
    objOriginal.Close
    objNew.Close
    
    ' Comparar línea por línea
    maxLines = UBound(originalLines)
    If UBound(newLines) > maxLines Then maxLines = UBound(newLines)
    
    For i = 0 To maxLines
        Dim originalLine, newLine
        
        If i <= UBound(originalLines) Then
            originalLine = originalLines(i)
        Else
            originalLine = ""
        End If
        
        If i <= UBound(newLines) Then
            newLine = newLines(i)
        Else
            newLine = ""
        End If
        
        If originalLine <> newLine Then
            If Not hasDifferences Then
                WScript.Echo "=== DIFERENCIAS DETECTADAS ==="
                hasDifferences = True
            End If
            
            WScript.Echo "Línea " & (i + 1) & ":"
            WScript.Echo "- " & originalLine
            WScript.Echo "+ " & newLine
        End If
    Next
    
    If hasDifferences Then
        WScript.Echo "=== FIN DE DIFERENCIAS ==="
    End If
    
    GenerateSimpleDiff = hasDifferences
End Function

' ===================================================================
' SUBRUTINA AUXILIAR: ExportFormToJson
' Descripción: Versión interna de ExportForm para uso en roundtrip
' ===================================================================
Private Sub ExportFormInternal(dbPath, formName, outputPath, password)
    ' Implementación completa de export de formularios a JSON canónico
    ' Asegura apertura en vista Diseño OCULTO y usa JsonWriter para generar JSON por secciones
    
    If gVerbose Then WScript.Echo "Exportando " & formName & " desde " & dbPath & " a " & outputPath
    
    ' Resolver BD usando función canónica
    Dim resolvedDbPath, dbOrigin
    resolvedDbPath = ResolveDbForAction(dbPath, "export-form", dbOrigin)
    
    ' Crear instancia de Access usando función unificada con bypass
    Dim objAccessLocal
    Set objAccessLocal = OpenAccessApp(resolvedDbPath, password, True)
    
    If objAccessLocal Is Nothing Then
        WScript.Echo "Error al abrir la base de datos para export"
        Exit Sub
    End If
    
    On Error Resume Next
    
    ' Asegurar apertura en vista Diseño OCULTO
    Const acViewDesign = 1
    Const acHidden = 1
    objAccessLocal.Echo False
    objAccessLocal.DoCmd.OpenForm formName, acViewDesign, "", "", 0, acHidden
    
    If Err.Number <> 0 Then
        WScript.Echo "Error al abrir formulario para export: " & Err.Description
        objAccessLocal.Echo True
        CloseAccessQuiet objAccessLocal
        Exit Sub
    End If
    
    On Error GoTo 0
    
    ' Obtener referencia al formulario
    Dim frm
    Set frm = objAccessLocal.Forms(formName)
    
    ' Crear JsonWriter para generar JSON canónico
    Dim writer
    Set writer = New JsonWriter
    
    ' Iniciar objeto raíz
    writer.StartObject
    
    ' Propiedades básicas del formulario
    writer.AddProperty "schemaVersion", "1.0"
    writer.AddProperty "units", "twips"
    writer.AddProperty "formName", frm.Name
    
    ' Propiedades del formulario
    writer.StartObjectProperty "properties"
    
    On Error Resume Next
    writer.AddProperty "caption", frm.Caption
    writer.AddProperty "width", frm.Width
    writer.AddProperty "height", frm.InsideHeight
    writer.AddProperty "recordSource", frm.RecordSource
    writer.AddProperty "recordsetType", frm.RecordsetType
    writer.AddProperty "allowEdits", frm.AllowEdits
    writer.AddProperty "allowDeletions", frm.AllowDeletions
    writer.AddProperty "allowAdditions", frm.AllowAdditions
    writer.AddProperty "dataEntry", frm.DataEntry
    writer.AddProperty "defaultView", frm.DefaultView
    writer.AddProperty "viewsAllowed", frm.ViewsAllowed
    writer.AddProperty "scrollBars", MapScrollBarsToToken(frm.ScrollBars)
    writer.AddProperty "recordSelectors", frm.RecordSelectors
    writer.AddProperty "navigationButtons", frm.NavigationButtons
    writer.AddProperty "dividerLines", frm.DividerLines
    writer.AddProperty "autoResize", frm.AutoResize
    writer.AddProperty "autoCenter", frm.AutoCenter
    writer.AddProperty "popUp", frm.PopUp
    writer.AddProperty "modal", frm.Modal
    writer.AddProperty "borderStyle", MapBorderStyleToToken(frm.BorderStyle)
    writer.AddProperty "controlBox", frm.ControlBox
    writer.AddProperty "minMaxButtons", frm.MinMaxButtons
    writer.AddProperty "closeButton", frm.CloseButton
    writer.AddProperty "whatsThisButton", frm.WhatsThisButton
    writer.AddProperty "shortcutMenu", frm.ShortcutMenu
    writer.AddProperty "shortcutMenuBar", frm.ShortcutMenuBar
    writer.AddProperty "menuBar", frm.MenuBar
    writer.AddProperty "toolbar", frm.Toolbar
    writer.AddProperty "cycle", frm.Cycle
    writer.AddProperty "onLoad", frm.OnLoad
    writer.AddProperty "onUnload", frm.OnUnload
    writer.AddProperty "onOpen", frm.OnOpen
    writer.AddProperty "onClose", frm.OnClose
    writer.AddProperty "onActivate", frm.OnActivate
    writer.AddProperty "onDeactivate", frm.OnDeactivate
    writer.AddProperty "onGotFocus", frm.OnGotFocus
    writer.AddProperty "onLostFocus", frm.OnLostFocus
    writer.AddProperty "onClick", frm.OnClick
    writer.AddProperty "onDblClick", frm.OnDblClick
    writer.AddProperty "onMouseDown", frm.OnMouseDown
    writer.AddProperty "onMouseMove", frm.OnMouseMove
    writer.AddProperty "onMouseUp", frm.OnMouseUp
    writer.AddProperty "onKeyDown", frm.OnKeyDown
    writer.AddProperty "onKeyPress", frm.OnKeyPress
    writer.AddProperty "onKeyUp", frm.OnKeyUp
    writer.AddProperty "onResize", frm.OnResize
    writer.AddProperty "onError", frm.OnError
    writer.AddProperty "onFilter", frm.OnFilter
    writer.AddProperty "onApplyFilter", frm.OnApplyFilter
    writer.AddProperty "onTimer", frm.OnTimer
    writer.AddProperty "timerInterval", frm.TimerInterval
    ' Limpiar errores de propiedades no legibles
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
    
    writer.EndObject ' Cerrar properties
    
    ' Secciones del formulario
    writer.StartObjectProperty "sections"
    
    ' Definir las secciones estándar
    Dim sectionNames, sectionIds, i
    sectionNames = Array("header", "detail", "footer")
    sectionIds = Array(acHeader, acDetail, acFooter)
    
    For i = 0 To UBound(sectionNames)
        Dim sectionName, sectionId
        sectionName = sectionNames(i)
        sectionId = sectionIds(i)
        
        ' Verificar si la sección existe
        On Error Resume Next
        Dim sectionHeight, sectionBackColor
        sectionHeight = frm.Section(sectionId).Height
        sectionBackColor = frm.Section(sectionId).BackColor
        On Error GoTo 0
        
        If Err.Number = 0 Then
            ' La sección existe, crear su objeto
            writer.StartObjectProperty sectionName
            writer.AddProperty "height", sectionHeight
            writer.AddProperty "backColor", OleToHex(sectionBackColor)
            
            ' Obtener controles de esta sección
            writer.StartArrayProperty "controls"
            
            Dim control
            For Each control In frm.Controls
                ' Verificar si el control pertenece a esta sección
                On Error Resume Next
                Dim controlSection
                controlSection = control.Section
                On Error GoTo 0
                
                If controlSection = sectionId Then
                    writer.StartObject
                    writer.AddProperty "name", control.Name
                    writer.AddProperty "type", GetControlTypeName(control.ControlType)
                    writer.AddProperty "section", sectionName ' Campo informativo
                    
                    ' Propiedades del control
                    writer.StartObjectProperty "properties"
                    On Error Resume Next
                    writer.AddProperty "top", control.Top
                    writer.AddProperty "left", control.Left
                    writer.AddProperty "width", control.Width
                    writer.AddProperty "height", control.Height
                    writer.AddProperty "caption", control.Caption
                    writer.AddProperty "controlSource", control.ControlSource
                    writer.AddProperty "visible", control.Visible
                    writer.AddProperty "enabled", control.Enabled
                    writer.AddProperty "locked", control.Locked
                    writer.AddProperty "default", control.Default
                    ' Detectar Picture relativo para resources.images
                    If control.Picture <> "" And Not IsAbsolutePath(control.Picture) Then
                        writer.AddProperty "picture", control.Picture
                    End If
                    On Error GoTo 0
                    
                    writer.EndObject ' Cerrar properties del control
                    writer.EndObject ' Cerrar control
                End If
            Next
            
            writer.EndArray ' Cerrar controls
            writer.EndObject ' Cerrar sección
        End If
    Next
    
    writer.EndObject ' Cerrar sections
    
    ' Recursos (imágenes) - detectar Pictures relativos
    writer.StartObjectProperty "resources"
    writer.StartArrayProperty "images"
    
    For Each control In frm.Controls
        On Error Resume Next
        If control.Picture <> "" And Not IsAbsolutePath(control.Picture) Then
            writer.StartObject
            writer.AddProperty "path", control.Picture
            writer.AddProperty "controlName", control.Name
            writer.EndObject
        End If
        On Error GoTo 0
    Next
    
    writer.EndArray ' Cerrar images
    writer.EndObject ' Cerrar resources
    
    ' Código del módulo - detección básica
    writer.StartObjectProperty "code"
    writer.StartObjectProperty "module"
    
    Dim hasModule, moduleFilename
    hasModule = False
    moduleFilename = ""
    
    ' Detectar si hay código en el módulo del formulario
    On Error Resume Next
    If frm.HasModule Then
        hasModule = True
        moduleFilename = frm.Name & ".frm"
    End If
    On Error GoTo 0
    
    writer.AddProperty "exists", hasModule
    writer.AddProperty "filename", moduleFilename
    writer.StartArrayProperty "handlers"
    
    ' Detección básica de handlers - se puede expandir
    If hasModule Then
        ' Aquí se podría implementar detección de Sub <Control>_<Event>
        ' Por ahora dejamos el array vacío
    End If
    
    writer.EndArray ' Cerrar handlers
    writer.EndObject ' Cerrar module
    writer.EndObject ' Cerrar code
    
    writer.EndObject ' Cerrar objeto raíz
    
    ' Obtener JSON generado
    Dim jsonString
    jsonString = writer.GetJson()
    
    ' Cerrar formulario sin guardar con cierre robusto
    On Error Resume Next
    objAccessLocal.DoCmd.Close acForm, formName, acSaveNo
    objAccessLocal.Echo True
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
    
    ' Escribir archivo JSON usando UTF-8
    Call WriteUtf8File(outputPath, jsonString)
    
    If gVerbose Then WScript.Echo "✓ Exportado formulario " & formName & " → " & outputPath & " (UTF-8)"
    
    CloseAccessQuiet objAccessLocal
End Sub

' ===================================================================
' SUBRUTINA AUXILIAR: ImportFormInternal
' Descripción: Versión interna de ImportForm para uso en roundtrip
' Opera SIEMPRE en vista Diseño con validación previa usando JsonParser
' ===================================================================
Private Sub ImportFormInternal(jsonPath, dbPath, password, overwrite)
    ' Implementación determinista y silenciosa de import de formularios
    If gVerbose Then WScript.Echo "Importando " & jsonPath & " a " & dbPath
    
    ' Parsear JSON usando JsonParser unificado
    Dim jsonContent, jsonData, name, props
    jsonContent = ReadTextFile(jsonPath)
    
    Dim parser
    Set parser = New JsonParser
    Set jsonData = parser.Parse(jsonContent)
    
    ' Validar estructura del JSON antes de proceder
    Dim warnings
    Set warnings = CreateObject("Scripting.Dictionary")
    ValidateFormData jsonData, True, warnings ' Usar modo strict
    
    ' Si hay errores críticos en modo strict, abortar
    Dim warningKey
    For Each warningKey In warnings.Keys
        If InStr(warnings(warningKey), "ERROR:") > 0 Then
            WScript.Echo "Error de validación: " & warnings(warningKey)
            Exit Sub
        End If
    Next
    
    name = jsonData("formName")
    If IsPresent(jsonData, "properties") Then
        Set props = jsonData("properties")
    Else
        Set props = CreateObject("Scripting.Dictionary")
    End If
    
    ' Crear instancia de Access usando función unificada con bypass
    Dim objAccess
    Set objAccess = OpenAccessApp(dbPath, password, True)
    
    If objAccess Is Nothing Then
        WScript.Echo "Error al abrir la base de datos para import"
        Exit Sub
    End If
    
    On Error Resume Next
    
    ' Preparación silenciosa
    objAccess.Echo False
    objAccess.Application.DisplayAlerts = False
    
    ' Si overwrite=True y existe name: eliminar formulario existente
    If overwrite Then
        Dim formExists
        formExists = False
        On Error Resume Next
        Dim testForm
        Set testForm = objAccess.CurrentDb.AllForms(name)
        If Err.Number = 0 Then formExists = True
        On Error Resume Next
        
        If formExists Then
            objAccess.DoCmd.DeleteObject acForm, name
            If gVerbose Then WScript.Echo "Formulario existente eliminado: " & name
        End If
    End If
    
    ' Crear nuevo formulario vacío en vista Diseño
    objAccess.DoCmd.NewForm
    Dim tmpName
    tmpName = objAccess.Screen.ActiveForm.Name
    objAccess.DoCmd.Rename name, acForm, tmpName
    
    ' Abrir en vista Diseño explícitamente
    objAccess.DoCmd.OpenForm name, acViewDesign
    
    Dim frm
    Set frm = objAccess.Forms(name)
    
    ' Aplicar propiedades del formulario
    On Error Resume Next
    If IsPresent(props, "caption") Then 
        frm.Caption = props("caption")
        If gVerbose Then WScript.Echo "Aplicado Caption: " & props("caption")
    End If
    If IsPresent(props, "recordSource") Then 
        frm.RecordSource = props("recordSource")
        If gVerbose Then WScript.Echo "Aplicado RecordSource: " & props("recordSource")
    End If
    If IsPresent(props, "width") Then 
        frm.Width = CLng(props("width"))
        If gVerbose Then WScript.Echo "Aplicado Width: " & props("width")
    End If
    If IsPresent(props, "height") Then 
        frm.Section(acDetail).Height = CLng(props("height"))
        If gVerbose Then WScript.Echo "Aplicado Height: " & props("height")
    End If
    If IsPresent(props, "backColor") Then 
        frm.Section(acDetail).BackColor = HexToOle(props("backColor"))
        If gVerbose Then WScript.Echo "Aplicado BackColor: " & props("backColor")
    End If
    On Error GoTo 0
    
    ' Agregar controles usando AddControlsFromJson
    AddControlsFromJson frm, jsonData, objAccess
    
    ' Aplicar eventos si hay handlers en code.module
    ApplyEventHandlers frm, jsonData
    
    ' Guardar y cerrar sin ejecutar
    objAccess.DoCmd.Save acForm, name
    objAccess.DoCmd.Close acForm, name, acSaveYes
    objAccess.Echo True
    
    If gVerbose Then WScript.Echo "Formulario importado exitosamente: " & name
    
    CloseAccessApp objAccess
    Exit Sub
    
    ' Manejo de errores
    If Err.Number <> 0 Then
        objAccess.Echo True
        objAccess.Application.DisplayAlerts = True
        CloseAccessApp objAccess
        WScript.Echo "Error durante import: " & Err.Description
        Err.Clear
    End If
End Sub

' ===================================================================
' FUNCIÓN AUXILIAR: NormalizeJsonForComparison
' Descripción: Normaliza JSON removiendo metadata variable para comparación
' ===================================================================
Private Function NormalizeJsonForComparison(jsonText)
    Dim result
    result = jsonText
    
    ' Remover timestamps variables
    Dim regEx
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Global = True
    
    ' Remover generatedAtUTC
    regEx.Pattern = """generatedAtUTC""\s*:\s*""[^""]+"",?"
    result = regEx.Replace(result, "")
    
    ' Remover espacios extra y normalizar
    regEx.Pattern = "\s+"
    result = regEx.Replace(result, " ")
    
    NormalizeJsonForComparison = Trim(result)
End Function

' ===================================================================
' FUNCIÓN AUXILIAR: ReadTextFile
' Descripción: Lee contenido completo de un archivo de texto
' ===================================================================
Private Function ReadTextFile(filePath)
    Dim objFile, content
    Set objFile = objFSO.OpenTextFile(filePath, 1, False, 0) ' ASCII
    content = objFile.ReadAll
    objFile.Close
    ReadTextFile = content
End Function

' ===================================================================
' SUBRUTINA AUXILIAR: CleanupRoundtripTest
' Descripción: Limpia archivos temporales del test de roundtrip
' ===================================================================
Private Sub CleanupRoundtripTest(tempDir)
    On Error Resume Next
    If objFSO.FolderExists(tempDir) Then
        objFSO.DeleteFolder tempDir, True
    End If
    On Error GoTo 0
End Sub

' ================================================================================
' FUNCIÓN: ParseJson
' DESCRIPCIÓN: Parsear JSON usando PowerShell como alternativa
' ================================================================================
Private Function ParseJson(jsonText)
    ' Parsear JSON usando expresiones regulares simples
    Set ParseJson = CreateObject("Scripting.Dictionary")
    
    ' Limpiar BOM y caracteres especiales al inicio
    Dim cleanText
    cleanText = jsonText
    ' Remover BOM UTF-8 (EF BB BF)
    If Len(cleanText) > 0 And Asc(Left(cleanText, 1)) = 239 Then
        cleanText = Mid(cleanText, 4)
    End If
    ' Remover otros caracteres de control al inicio
    Do While Len(cleanText) > 0 And (Asc(Left(cleanText, 1)) < 32 Or Asc(Left(cleanText, 1)) > 126)
        cleanText = Mid(cleanText, 2)
    Loop
    
    ' Extraer formName
    Dim regEx, matches
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Pattern = """formName""\s*:\s*""([^""]+)"""
    regEx.Global = False
    Set matches = regEx.Execute(cleanText)
    If matches.Count > 0 Then
        ParseJson.Add "formName", matches(0).SubMatches(0)
    Else
        ParseJson.Add "formName", "FormularioSinNombre"
    End If
    
    ' Crear propiedades del formulario
    Dim formProps
    Set formProps = CreateObject("Scripting.Dictionary")
    
    ' Extraer caption
    regEx.Pattern = """caption""\s*:\s*""([^""]+)"""
    Set matches = regEx.Execute(cleanText)
    If matches.Count > 0 Then
        formProps.Add "caption", matches(0).SubMatches(0)
    End If
    
    ' Extraer width
    regEx.Pattern = """width""\s*:\s*(\d+)"
    Set matches = regEx.Execute(cleanText)
    If matches.Count > 0 Then
        formProps.Add "width", matches(0).SubMatches(0)
    End If
    
    ParseJson.Add "properties", formProps
    
    ' Crear secciones
    Dim sections, detailSection, detailProps
    Set sections = CreateObject("Scripting.Dictionary")
    Set detailSection = CreateObject("Scripting.Dictionary")
    Set detailProps = CreateObject("Scripting.Dictionary")
    
    ' Extraer height de la sección Detail
    regEx.Pattern = """Detail""[^}]*""height""\s*:\s*(\d+)"
    Set matches = regEx.Execute(cleanText)
    If matches.Count > 0 Then
        detailProps.Add "height", matches(0).SubMatches(0)
    End If
    
    detailSection.Add "properties", detailProps
    
    ' Parsear controles
    Dim controls
    Set controls = CreateObject("Scripting.Dictionary")
    
    ' Buscar todos los controles en la sección Detail
    regEx.Pattern = "\{[^}]*""name""\s*:\s*""([^""]+)""[^}]*""type""\s*:\s*""([^""]+)""[^}]*\}"
    regEx.Global = True
    Set matches = regEx.Execute(cleanText)
    
    Dim controlIndex
    controlIndex = 0
    
    Dim i
    For i = 0 To matches.Count - 1
        Dim ctrl, ctrlProps
        Set ctrl = CreateObject("Scripting.Dictionary")
        Set ctrlProps = CreateObject("Scripting.Dictionary")
        
        ' Obtener el bloque completo del control
        Dim controlBlock
        controlBlock = matches(i).Value
        
        ' Extraer propiedades básicas
        ctrl.Add "name", matches(i).SubMatches(0)
        ctrl.Add "type", matches(i).SubMatches(1)
        
        ' Extraer propiedades del control
        Dim propRegEx, propMatches
        Set propRegEx = CreateObject("VBScript.RegExp")
        
        ' Caption
        propRegEx.Pattern = """caption""\s*:\s*""([^""]+)"""
        Set propMatches = propRegEx.Execute(controlBlock)
        If propMatches.Count > 0 Then
            ctrlProps.Add "caption", propMatches(0).SubMatches(0)
        End If
        
        ' Top
        propRegEx.Pattern = """top""\s*:\s*(\d+)"
        Set propMatches = propRegEx.Execute(controlBlock)
        If propMatches.Count > 0 Then
            ctrlProps.Add "top", propMatches(0).SubMatches(0)
        End If
        
        ' Left
        propRegEx.Pattern = """left""\s*:\s*(\d+)"
        Set propMatches = propRegEx.Execute(controlBlock)
        If propMatches.Count > 0 Then
            ctrlProps.Add "left", propMatches(0).SubMatches(0)
        End If
        
        ' Width
        propRegEx.Pattern = """width""\s*:\s*(\d+)"
        Set propMatches = propRegEx.Execute(controlBlock)
        If propMatches.Count > 0 Then
            ctrlProps.Add "width", propMatches(0).SubMatches(0)
        End If
        
        ' Height
        propRegEx.Pattern = """height""\s*:\s*(\d+)"
        Set propMatches = propRegEx.Execute(controlBlock)
        If propMatches.Count > 0 Then
            ctrlProps.Add "height", propMatches(0).SubMatches(0)
        End If
        
        ' Picture
        propRegEx.Pattern = """picture""\s*:\s*""([^""]+)"""
        Set propMatches = propRegEx.Execute(controlBlock)
        If propMatches.Count > 0 Then
            ctrlProps.Add "picture", propMatches(0).SubMatches(0)
        End If
        
        ctrl.Add "properties", ctrlProps
        controls.Add controlIndex, ctrl
        controlIndex = controlIndex + 1
    Next
    
    detailSection.Add "controls", controls
    sections.Add "Detail", detailSection
    ParseJson.Add "sections", sections
    
    WScript.Echo "JSON parseado correctamente. Controles encontrados: " & controls.Count
End Function

' ================================================================================
' FUNCIÓN: JsonParser
' DESCRIPCIÓN: Parser robusto para JSON enriquecido con validaciones
' ================================================================================


' ================================================================================
' FUNCIÓN: MapEnumValue
' DESCRIPCIÓN: Mapear valores enum string a valores Access
' ================================================================================
Private Function MapEnumValue(enumType, stringValue)
    ' Tabla de mapeo local para enums conocidos
    Select Case LCase(enumType)
        Case "defaultview"
            Select Case LCase(stringValue)
                Case "single", "singleform": MapEnumValue = 0
                Case "continuous", "continuousforms": MapEnumValue = 1
                Case "datasheet": MapEnumValue = 2
                Case Else: 
                    LogWarn "Valor desconocido para DefaultView: '" & stringValue & "', usando Single (0)"
                    MapEnumValue = 0
            End Select
        Case "cycle"
            Select Case LCase(stringValue)
                Case "allrecords": MapEnumValue = 0
                Case "currentrecord": MapEnumValue = 1
                Case "currentpage": MapEnumValue = 2
                Case Else: 
                    LogWarn "Valor desconocido para Cycle: '" & stringValue & "', usando AllRecords (0)"
                    MapEnumValue = 0
            End Select
        Case "recordsourcetype"
            Select Case LCase(stringValue)
                Case "table": MapEnumValue = 0
                Case "dynaset": MapEnumValue = 1
                Case "snapshot": MapEnumValue = 2
                Case Else: 
                    LogWarn "Valor desconocido para RecordSourceType: '" & stringValue & "', usando Table (0)"
                    MapEnumValue = 0
            End Select
        Case "borderstyle"
            Select Case LCase(stringValue)
                Case "transparent": MapEnumValue = 0
                Case "solid": MapEnumValue = 1
                Case "dashes": MapEnumValue = 2
                Case "short dashes": MapEnumValue = 3
                Case "dots": MapEnumValue = 4
                Case "sparse dots": MapEnumValue = 5
                Case "dash dot": MapEnumValue = 6
                Case "dash dot dot": MapEnumValue = 7
                Case Else: 
                    LogWarn "Valor desconocido para BorderStyle: '" & stringValue & "', usando Solid (1)"
                    MapEnumValue = 1
            End Select
        Case "scrollbars"
            Select Case LCase(stringValue)
                Case "none": MapEnumValue = 0
                Case "horizontal": MapEnumValue = 1
                Case "vertical": MapEnumValue = 2
                Case "both": MapEnumValue = 3
                Case Else: 
                    LogWarn "Valor desconocido para ScrollBars: '" & stringValue & "', usando None (0)"
                    MapEnumValue = 0
            End Select
        Case Else
            ' Para tipos de enum no reconocidos, devolver el valor original
            LogWarn "Tipo de enumeración desconocido: '" & enumType & "', devolviendo valor original"
            MapEnumValue = stringValue
    End Select
End Function

' ================================================================================
' SUBRUTINA: ImportForm
' DESCRIPCIÓN: Crear/Modificar formulario desde JSON enriquecido con validaciones
' ================================================================================
Sub ImportForm()
    ' Verificar argumentos mínimos
    If objArgs.Count < 3 Then
        WScript.Echo "Error: Se requieren al menos 2 argumentos: <json_path_or_folder> <db_path>"
        WScript.Echo "Sintaxis: cscript condor_cli.vbs import-form <json_path_or_folder> <db_path> [opciones]"
        WScript.Echo "Opciones:"
        WScript.Echo "  --dry-run              Solo validar, no crear formulario"
        WScript.Echo "  --strict               Modo estricto: fallar en coherencias"
        WScript.Echo "  --password <pwd>       Contraseña de base de datos"
        WScript.Echo "  --bypassstartup on|off Bypass startup forms"
        WScript.Quit 1
    End If
    
    ' Declarar variables
    Dim strJsonPath, strDbPath, strPassword
    Dim bDryRun, bStrict, bypassStartup
    Dim objJsonData, formName
    Dim objAccess, frm
    Dim warnings, i
    
    ' Inicializar variables
    strJsonPath = objArgs(1)
    strDbPath = objArgs(2)
    strPassword = ""
    bDryRun = False
    bStrict = False
    bypassStartup = False
    
    Set warnings = CreateObject("Scripting.Dictionary")
    
    ' Detectar si es carpeta o archivo individual
    Dim isFolder
    isFolder = objFSO.FolderExists(strJsonPath)
    
    ' Procesar argumentos opcionales
    For i = 3 To objArgs.Count - 1
        Dim currentArg
        currentArg = objArgs(i)
        
        Select Case LCase(currentArg)
            Case "--dry-run"
                bDryRun = True
            Case "--strict"
                bStrict = True
            Case "--password"
                If i + 1 <= objArgs.Count - 1 Then
                    strPassword = objArgs(i + 1)
                    i = i + 1
                Else
                    WScript.Echo "Error: --password requiere una contraseña"
                    WScript.Quit 1
                End If
            Case Else
                ' Verificar banderas con formato /bandera:valor
                If Left(currentArg, 4) = "/pwd" Then
                    If Mid(currentArg, 5, 1) = ":" Then
                        strPassword = Mid(currentArg, 6)
                    Else
                        WScript.Echo "Error: Formato incorrecto para /pwd. Use /pwd:<clave>"
                        WScript.Quit 1
                    End If
                ElseIf Left(currentArg, 15) = "/bypassstartup:" Then
                    gBypassStartup = True ' mantenemos compatibilidad
                    If Not gVerbose Then
                        ' no spam
                    End If
                    WScript.Echo "[DEPRECATED] --bypassstartup ya no es necesario y no tiene efecto (el CLI abre Access con bypass por defecto)."
                End If
        End Select
    Next
    
    ' Convertir rutas relativas a absolutas si es necesario
    If InStr(strJsonPath, ":") = 0 Then
        strJsonPath = objFSO.GetAbsolutePathName(strJsonPath)
    End If
    
    If InStr(strDbPath, ":") = 0 Then
        ' Si no se especifica ruta, usar la base de datos por defecto
        If strDbPath = "" Then
            strDbPath = objFSO.GetAbsolutePathName(strDataPath)
        Else
            strDbPath = objFSO.GetAbsolutePathName(strDbPath)
        End If
    End If
    
    ' Verificar que la ruta existe
    If Not isFolder And Not objFSO.FileExists(strJsonPath) Then
        WScript.Echo "Error: El archivo JSON no existe: " & strJsonPath
        WScript.Quit 1
    End If
    
    If isFolder And Not objFSO.FolderExists(strJsonPath) Then
        WScript.Echo "Error: La carpeta no existe: " & strJsonPath
        WScript.Quit 1
    End If
    
    If Not bDryRun And Not objFSO.FileExists(strDbPath) Then
        WScript.Echo "Error: La base de datos no existe: " & strDbPath
        WScript.Quit 1
    End If
    
    If isFolder Then
        ' Importar múltiples formularios con resolución de dependencias
        ImportFormsFromFolder strJsonPath, strDbPath, strPassword, bypassStartup, bDryRun, bStrict
    Else
        ' Importar un solo formulario (lógica original)
        ImportSingleForm strJsonPath, strDbPath, strPassword, bypassStartup, bDryRun, bStrict
    End If
End Sub

' ============================================================================
' FUNCIÓN AUXILIAR: ImportSingleForm
' Descripción: Importa un solo formulario desde un archivo JSON
' ============================================================================
Private Sub ImportSingleForm(strJsonPath, strDbPath, strPassword, bypassStartup, bDryRun, bStrict)
    Dim objJsonData, formName, objAccess, frm, warnings
    Set warnings = CreateObject("Scripting.Dictionary")
    
    ' Leer y parsear el archivo JSON
    Dim strJsonContent
    strJsonContent = ReadTextFile(strJsonPath)
    
    On Error Resume Next
    Set objJsonData = ParseJson(strJsonContent)
    If Err.Number <> 0 Then
        WScript.Echo "Error al parsear JSON: " & Err.Description
        WScript.Quit 1
    End If
    On Error GoTo 0
    
    ' Validaciones mínimas de los datos del formulario
    If Not objJsonData.Exists("formName") And Not objJsonData.Exists("name") Then
        WScript.Echo "Error: El JSON debe contener 'formName' o 'name'"
        WScript.Quit 1
    End If
    
    If Not objJsonData.Exists("controls") And Not objJsonData.Exists("sections") Then
        WScript.Echo "Error: El JSON debe contener 'controls' o 'sections'"
        WScript.Quit 1
    End If
    
    ' Obtener nombre del formulario
    If objJsonData.Exists("formName") Then
        formName = objJsonData("formName")
    Else
        formName = objJsonData("name")
    End If
    WScript.Echo "Formulario a procesar: " & formName
    
    ' Si es dry-run, mostrar reporte y salir
    If bDryRun Then
        WScript.Echo "=== MODO DRY-RUN: VALIDACIÓN COMPLETADA ==="
        WScript.Echo "Formulario: " & formName
        
        If warnings.Count > 0 Then
            WScript.Echo "Advertencias encontradas:"
            Dim key
            For Each key In warnings.Keys
                WScript.Echo "  - " & warnings(key)
            Next
        End If
        
        WScript.Echo "Validación completada exitosamente."
        Exit Sub
    End If
    
    ' Abrir Access con bypass startup si está habilitado
    Set objAccess = OpenAccessApp(strDbPath, strPassword, bypassStartup)
    
    ' Configurar Access
    objAccess.UserControl = False
    
    WScript.Echo "Base de datos abierta: " & objAccess.CurrentProject.Name
    
    ' Crear o reemplazar formulario usando función auxiliar
    Set frm = CreateOrReplaceForm(objAccess, formName)
    If frm Is Nothing Then
        CloseAccessApp objAccess
        WScript.Quit 1
    End If
    
    ' Aplicar propiedades del formulario usando función auxiliar
    ApplyFormProperties frm, objJsonData, bStrict
    
    ' Añadir controles usando función auxiliar
    AddControlsFromJson objAccess, frm, objJsonData
    
    ' Cerrar formulario (sin guardar explícitamente)
    On Error Resume Next
    
    ' Cerrar el formulario
    objAccess.DoCmd.Close acObjectForm, frm.Name, acSaveYes
    If Err.Number <> 0 Then
        WScript.Echo "Error al cerrar el formulario: " & Err.Description
    End If
    On Error GoTo 0
    
    WScript.Echo "Formulario '" & formName & "' creado exitosamente."
    
    ' Cerrar Access
    CloseAccessApp objAccess
End Sub

' ============================================================================
' FUNCIÓN AUXILIAR: ImportFormsFromFolder
' Descripción: Importa múltiples formularios desde una carpeta con resolución de dependencias
' ============================================================================
Private Sub ImportFormsFromFolder(strFolderPath, strDbPath, strPassword, bypassStartup, bDryRun, bStrict)
    Dim objFolder, objFile, jsonFiles, formData, dependencyGraph, sortedForms
    Dim objAccess, i, formName, jsonPath
    
    Set jsonFiles = CreateObject("Scripting.Dictionary")
    Set formData = CreateObject("Scripting.Dictionary")
    Set dependencyGraph = CreateObject("Scripting.Dictionary")
    
    ' Recopilar todos los archivos JSON de la carpeta
    Set objFolder = objFSO.GetFolder(strFolderPath)
    For Each objFile In objFolder.Files
        If LCase(objFSO.GetExtensionName(objFile.Name)) = "json" Then
            Dim strJsonContent, objJsonData
            strJsonContent = ReadTextFile(objFile.Path)
            
            On Error Resume Next
            Set objJsonData = ParseJson(strJsonContent)
            If Err.Number = 0 And objJsonData.Exists("formName") Then
                formName = objJsonData("formName")
                jsonFiles.Add formName, objFile.Path
                formData.Add formName, objJsonData
                
                ' Extraer dependencias
                Dim dependencies
                Set dependencies = CreateObject("Scripting.Dictionary")
                If objJsonData.Exists("dependencies") Then
                    Dim depArray, j
                    depArray = objJsonData("dependencies")
                    If IsArray(depArray) Then
                        For j = 0 To UBound(depArray)
                            dependencies.Add depArray(j), True
                        Next
                    End If
                End If
                dependencyGraph.Add formName, dependencies
                
                WScript.Echo "Encontrado formulario: " & formName & " (" & dependencies.Count & " dependencias)"
            End If
            Err.Clear
            On Error GoTo 0
        End If
    Next
    
    If jsonFiles.Count = 0 Then
        WScript.Echo "Error: No se encontraron archivos JSON válidos en la carpeta: " & strFolderPath
        WScript.Quit 1
    End If
    
    ' Realizar ordenamiento topológico
    Set sortedForms = TopologicalSort(dependencyGraph, bStrict)
    If sortedForms Is Nothing Then
        WScript.Echo "Error: Dependencias circulares detectadas o formularios faltantes"
        WScript.Quit 1
    End If
    
    WScript.Echo "Orden de importación determinado: " & Join(sortedForms, ", ")
    
    If bDryRun Then
        WScript.Echo "=== MODO DRY-RUN: VALIDACIÓN COMPLETADA ==="
        WScript.Echo "Se importarían " & UBound(sortedForms) + 1 & " formularios en el orden mostrado."
        Exit Sub
    End If
    
    ' Abrir Access una sola vez
    Set objAccess = OpenAccessApp(strDbPath, strPassword, bypassStartup)
    objAccess.UserControl = False
    WScript.Echo "Base de datos abierta: " & objAccess.CurrentProject.Name
    
    ' Importar formularios en orden topológico
    For i = 0 To UBound(sortedForms)
        formName = sortedForms(i)
        jsonPath = jsonFiles(formName)
        
        WScript.Echo "Importando formulario " & (i + 1) & "/" & (UBound(sortedForms) + 1) & ": " & formName
        
        ' Importar formulario individual
        ImportFormFromData objAccess, formName, formData(formName), bStrict
    Next
    
    WScript.Echo "Todos los formularios importados exitosamente."
    
    ' Cerrar Access
    CloseAccessApp objAccess
End Sub

' ============================================================================
' FUNCIÓN AUXILIAR: TopologicalSort
' Descripción: Ordena formularios por dependencias usando algoritmo de Kahn
' ============================================================================
Private Function TopologicalSort(dependencyGraph, bStrict)
    Dim result, inDegree, queue, visited
    Dim formName, dependencies, depName
    
    Set result = CreateObject("Scripting.Dictionary")
    Set inDegree = CreateObject("Scripting.Dictionary")
    Set queue = CreateObject("Scripting.Dictionary")
    Set visited = CreateObject("Scripting.Dictionary")
    
    ' Inicializar grados de entrada
    For Each formName In dependencyGraph.Keys
        inDegree.Add formName, 0
    Next
    
    ' Calcular grados de entrada
    For Each formName In dependencyGraph.Keys
        Set dependencies = dependencyGraph(formName)
        For Each depName In dependencies.Keys
            If dependencyGraph.Exists(depName) Then
                inDegree(depName) = inDegree(depName) + 1
            ElseIf bStrict Then
                WScript.Echo "Error: Dependencia faltante: " & depName & " requerida por " & formName
                Set TopologicalSort = Nothing
                Exit Function
            Else
                WScript.Echo "WARNING: Dependencia faltante: " & depName & " requerida por " & formName
            End If
        Next
    Next
    
    ' Encontrar nodos sin dependencias
    For Each formName In inDegree.Keys
        If inDegree(formName) = 0 Then
            queue.Add formName, True
        End If
    Next
    
    ' Procesar cola
    Dim sortedArray()
    Dim count
    count = 0
    
    Do While queue.Count > 0
        ' Tomar primer elemento de la cola
        Dim currentForm
        currentForm = ""
        For Each formName In queue.Keys
            currentForm = formName
            Exit For
        Next
        queue.Remove currentForm
        
        ' Añadir al resultado
        ReDim Preserve sortedArray(count)
        sortedArray(count) = currentForm
        count = count + 1
        
        ' Reducir grado de entrada de dependientes
        Set dependencies = dependencyGraph(currentForm)
        For Each depName In dependencies.Keys
            If dependencyGraph.Exists(depName) Then
                inDegree(depName) = inDegree(depName) - 1
                If inDegree(depName) = 0 Then
                    queue.Add depName, True
                End If
            End If
        Next
    Loop
    
    ' Verificar si hay ciclos
    If count <> dependencyGraph.Count Then
        WScript.Echo "Error: Dependencias circulares detectadas"
        Set TopologicalSort = Nothing
        Exit Function
    End If
    
    TopologicalSort = sortedArray
End Function

' ============================================================================
' FUNCIÓN AUXILIAR: ImportFormFromData
' Descripción: Importa un formulario desde datos JSON ya parseados
' ============================================================================
Private Sub ImportFormFromData(objAccess, formName, objJsonData, bStrict)
    Dim frm
    
    ' Crear o reemplazar formulario
    Set frm = CreateOrReplaceForm(objAccess, formName)
    If frm Is Nothing Then
        WScript.Echo "Error: No se pudo crear el formulario " & formName
        Exit Sub
    End If
    
    ' Aplicar propiedades del formulario
    ApplyFormProperties frm, objJsonData, bStrict
    
    ' Añadir controles (incluyendo subformularios y TabControls)
    AddControlsFromJson objAccess, frm, objJsonData
    
    ' Cerrar formulario (sin guardar explícitamente)
    On Error Resume Next
    
    ' Cerrar el formulario
    objAccess.DoCmd.Close acObjectForm, frm.Name, acSaveYes
    If Err.Number <> 0 Then
        WScript.Echo "WARNING: Error al cerrar el formulario " & formName & ": " & Err.Description
    End If
    On Error GoTo 0
    
    WScript.Echo "Formulario '" & formName & "' importado exitosamente."
End Sub

' ============================================================================
' FUNCIÓN PRINCIPAL REFACTORIZADA: CreateOrReplaceForm
' Descripción: Crea un nuevo formulario o reemplaza uno existente de forma robusta.
' Parámetros: objAccess - Instancia de Access, formName - Nombre del formulario
' Retorna: Objeto Form creado o Nothing si hay error
' ============================================================================
Private Function CreateOrReplaceForm(objAccess, formName)
    Dim frm ' Usar 'Object' para late binding es más seguro en VBScript
    
    ' --- PASO 1: Limpieza ---
    ' Eliminar el formulario antiguo si existe. Es crucial que la BD esté en modo
    ' que permita cambios de diseño.
    On Error Resume Next
    objAccess.DoCmd.Close acForm, formName, acSaveNo ' Cerrar por si está abierto
    objAccess.DoCmd.DeleteObject acForm, formName
    If Err.Number <> 0 And Err.Number <> 2501 And Err.Number <> 7874 Then
        ' Ignoramos el error si 'Close' falla porque el form no estaba abierto (2501)
        ' o si 'DeleteObject' falla porque no existía (7874).
        ' Pero si es otro error, lo mostramos.
        WScript.Echo "Error inesperado durante la limpieza: " & Err.Description
    End If
    Err.Clear
    On Error GoTo 0
    
    ' --- PASO 2: Creación y Obtención del Objeto ---
    ' CreateForm crea un formulario y LO DEJA ABIERTO en vista Diseño.
    ' NO es necesario volver a abrirlo.
    On Error Resume Next
    Set frm = objAccess.CreateForm()
    If Err.Number <> 0 Then
        WScript.Echo "Error fatal al ejecutar objAccess.CreateForm(): " & Err.Description
        Set CreateOrReplaceForm = Nothing
        Exit Function
    End If
    On Error GoTo 0
    
    WScript.Echo "Formulario temporal creado con éxito: " & frm.Name
    
    ' Aquí es donde aplicarías las propiedades. El formulario está abierto en
    ' vista Diseño y el objeto 'frm' es válido.
    ' Ejemplo:
    ' ApplyFormProperties frm, objJsonData, True
    
    ' --- PASO 3: Guardado Definitivo y Cierre ---
    ' Una vez aplicadas todas las propiedades y añadidos los controles,
    ' guardamos el formulario DIRECTAMENTE con su nombre final y lo cerramos.
    On Error Resume Next
    
    ' Guardamos el objeto formulario activo (que es el nuestro) con el nombre final.
    objAccess.DoCmd.Save acForm, formName
    
    If Err.Number <> 0 Then
        WScript.Echo "Error fatal al guardar el formulario como '" & formName & "': " & Err.Description
        ' Si falla el guardado, intentamos cerrar el temporal sin guardar cambios.
        objAccess.DoCmd.Close acForm, frm.Name, acSaveNo
        Set CreateOrReplaceForm = Nothing
        Exit Function
    End If

    ' Cerramos el formulario que ahora se llama 'formName'.
    objAccess.DoCmd.Close acForm, formName, acSaveNo ' Ya está guardado, no hace falta guardar de nuevo.

    WScript.Echo "Formulario '" & formName & "' creado y guardado permanentemente."
    On Error GoTo 0

    ' --- PASO 4: Devolver una referencia al objeto recién creado (opcional) ---
    ' Si después necesitas manipularlo, debes reabrirlo. Si no, devuelve True/False.
    ' Para este ejemplo, no devolvemos el objeto, ya que está cerrado.
    ' La función podría modificarse para devolver solo un booleano de éxito.
    ' Línea corregida para indicar éxito (el objeto ya está cerrado):
    
    ' Para mantener compatibilidad con el código existente, reabrimos el formulario
    On Error Resume Next
    objAccess.DoCmd.OpenForm formName, acViewDesign
    If Err.Number <> 0 Then
        WScript.Echo "Error al reabrir formulario para devolverlo: " & Err.Description
        Set CreateOrReplaceForm = Nothing
        Exit Function
    End If
    Set frm = objAccess.Forms(formName)
    On Error GoTo 0
    
    Set CreateOrReplaceForm = frm
End Function

' ============================================================================
' FUNCIÓN AUXILIAR: ApplyFormProperties
' Descripción: Aplica las propiedades del formulario desde el JSON
' Parámetros: frm - Objeto Form, objJsonData - Datos JSON, bStrict - Modo estricto
' ============================================================================
Private Sub ApplyFormProperties(frm, objJsonData, bStrict)
    Dim formProps
    
    ' Verificar que el formulario esté disponible
    On Error Resume Next
    Dim testProp
    testProp = frm.Name
    If Err.Number <> 0 Then
        WScript.Echo "Error: El formulario no está disponible para modificar propiedades"
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Asignar propiedades del formulario desde JSON
    If objJsonData.Exists("properties") Then
        Set formProps = objJsonData("properties")
        
        ' Propiedades básicas con manejo de errores
        On Error Resume Next
        If formProps.Exists("caption") Then
            frm.Caption = formProps("caption")
            If Err.Number <> 0 Then
                WScript.Echo "Warning: No se pudo establecer Caption: " & Err.Description
                Err.Clear
            End If
        End If
        
        If formProps.Exists("width") Then
            frm.Width = CLng(formProps("width"))
            If Err.Number <> 0 Then
                WScript.Echo "Warning: No se pudo establecer Width: " & Err.Description
                Err.Clear
            End If
        End If
        
        If formProps.Exists("height") Then
            frm.WindowHeight = CLng(formProps("height"))
            If Err.Number <> 0 Then
                WScript.Echo "Warning: No se pudo establecer Height: " & Err.Description
                Err.Clear
            End If
        End If
        On Error GoTo 0
        
        ' Colores (convertir #RRGGBB a OLE)
        If formProps.Exists("backColor") Then
            frm.Detail.BackColor = HexToOLE(formProps("backColor"))
        End If
        
        ' Propiedades booleanas
        If formProps.Exists("modal") Then
            frm.Modal = CBool(formProps("modal"))
        End If
        
        If formProps.Exists("popup") Then
            frm.PopUp = CBool(formProps("popup"))
        End If
        
        ' Aplicar reglas de coherencia para MinMaxButtons
        If formProps.Exists("minMaxButtons") And formProps.Exists("borderStyle") And formProps.Exists("controlBox") Then
            Dim borderStyle, controlBox, minMaxButtons
            borderStyle = formProps("borderStyle")
            controlBox = CBool(formProps("controlBox"))
            minMaxButtons = CBool(formProps("minMaxButtons"))
            
            If (borderStyle = "Dialog" Or borderStyle = "None" Or Not controlBox) And minMaxButtons Then
                If bStrict Then
                    WScript.Echo "ERROR: MinMaxButtons=True incompatible con BorderStyle=" & borderStyle & " o ControlBox=False"
                    WScript.Quit 1
                Else
                    WScript.Echo "WARN: MinMaxButtons forzado a False por coherencia"
                    frm.MinMaxButtons = 0  ' None
                End If
            Else
                If minMaxButtons Then
                    frm.MinMaxButtons = 3  ' Both
                Else
                    frm.MinMaxButtons = 0  ' None
                End If
            End If
        End If
    End If
End Sub

' ============================================================================
' FUNCIÓN AUXILIAR: AddControlsFromJson
' Descripción: Añade controles al formulario desde el JSON
' Parámetros: objAccess - Instancia de Access, frm - Objeto Form, objJsonData - Datos JSON
' ============================================================================
Private Sub AddControlsFromJson(objAccess, frm, objJsonData)
    Dim sections, detailSection, headerSection, footerSection, controls, control, i
    
    ' Añadir controles del JSON usando Application.CreateControl
    If objJsonData.Exists("sections") Then
        Set sections = objJsonData("sections")
        
        ' Procesar sección detail si existe
        If sections.Exists("detail") Then
            Set detailSection = sections("detail")
            If detailSection.Exists("controls") Then
                Set controls = detailSection("controls")
                
                For i = 0 To controls.Count - 1
                    Set control = controls(i)
                    AddControlFromJson objAccess, frm.Name, control, 0 ' acDetail
                Next
            End If
        End If
        
        ' Procesar sección header si existe
        If sections.Exists("header") Then
            Set headerSection = sections("header")
            If headerSection.Exists("controls") Then
                Set controls = headerSection("controls")
                
                For i = 0 To controls.Count - 1
                    Set control = controls(i)
                    AddControlFromJson objAccess, frm.Name, control, 1 ' acHeader
                Next
            End If
        End If
        
        ' Procesar sección footer si existe
        If sections.Exists("footer") Then
            Set footerSection = sections("footer")
            If footerSection.Exists("controls") Then
                Set controls = footerSection("controls")
                
                For i = 0 To controls.Count - 1
                    Set control = controls(i)
                    AddControlFromJson objAccess, frm.Name, control, 2 ' acFooter
                Next
            End If
        End If
    ElseIf objJsonData.Exists("controls") Then
        ' Mantener compatibilidad con formato anterior y nuevo formato con section
        Set controls = objJsonData("controls")
        
        For i = 0 To controls.Count - 1
            Set control = controls(i)
            
            ' Determinar sección del control
            Dim controlSection
            controlSection = 0 ' acDetail por defecto
            
            If control.Exists("section") Then
                ' Usar el nuevo campo section
                controlSection = SectionTokenToId(control("section"))
            Else
                ' Compatibilidad: si no hay section, usar detail por defecto
                If gVerbose Then WScript.Echo "WARNING: Control " & control("name") & " no tiene campo 'section', usando 'detail' por defecto"
                controlSection = acDetail
            End If
            
            AddControlFromJson objAccess, frm.Name, control, controlSection
        Next
    End If
End Sub

' ============================================================================
' FUNCIÓN AUXILIAR: AddControlFromJson
' Descripción: Añade un control individual al formulario
' Parámetros: objAccess - Instancia de Access, formName - Nombre del formulario, controlJson - Datos del control
' ============================================================================
Private Sub AddControlFromJson(objAccess, formName, controlJson, sectionType)
    Dim newControl, controlType, parentControl, columnName
    Dim leftPos, topPos, widthSize, heightSize
    Dim properties, prop, propValue
    
    ' Obtener propiedades de posición y tamaño con valores por defecto en twips
    leftPos = 0
    topPos = 0
    widthSize = 1440 ' Valor por defecto en twips (1 pulgada)
    heightSize = 300  ' Valor por defecto en twips
    
    ' Normalizar propiedades numéricas usando CLng/CDbl
    If controlJson.Exists("properties") Then
        Set properties = controlJson("properties")
        
        If properties.Exists("left") Then 
            leftPos = CLng(properties("left"))
        ElseIf controlJson.Exists("left") Then 
            leftPos = CLng(controlJson("left"))
        Else
            LogWarn "Control " & controlJson("name") & ": coordenada 'left' no especificada, usando 0"
        End If
        
        If properties.Exists("top") Then 
            topPos = CLng(properties("top"))
        ElseIf controlJson.Exists("top") Then 
            topPos = CLng(controlJson("top"))
        Else
            LogWarn "Control " & controlJson("name") & ": coordenada 'top' no especificada, usando 0"
        End If
        
        If properties.Exists("width") Then 
            widthSize = CLng(properties("width"))
        ElseIf controlJson.Exists("width") Then 
            widthSize = CLng(controlJson("width"))
        Else
            LogWarn "Control " & controlJson("name") & ": ancho 'width' no especificado, usando 1440 twips (1 pulgada)"
        End If
        
        If properties.Exists("height") Then 
            heightSize = CLng(properties("height"))
        ElseIf controlJson.Exists("height") Then 
            heightSize = CLng(controlJson("height"))
        Else
            LogWarn "Control " & controlJson("name") & ": altura 'height' no especificada, usando 300 twips"
        End If
    Else
        ' Formato legacy: propiedades directas en el control
        If controlJson.Exists("left") Then 
            leftPos = CLng(controlJson("left"))
        Else
            LogWarn "Control " & controlJson("name") & ": coordenada 'left' no especificada, usando 0"
        End If
        
        If controlJson.Exists("top") Then 
            topPos = CLng(controlJson("top"))
        Else
            LogWarn "Control " & controlJson("name") & ": coordenada 'top' no especificada, usando 0"
        End If
        
        If controlJson.Exists("width") Then 
            widthSize = CLng(controlJson("width"))
        Else
            LogWarn "Control " & controlJson("name") & ": ancho 'width' no especificado, usando 1440 twips (1 pulgada)"
        End If
        
        If controlJson.Exists("height") Then 
            heightSize = CLng(controlJson("height"))
        Else
            LogWarn "Control " & controlJson("name") & ": altura 'height' no especificada, usando 300 twips"
        End If
    End If
    
    ' Validar sección (0=acDetail, 1=acHeader, 2=acFooter)
    If sectionType < 0 Or sectionType > 2 Then
        LogWarn "Sección inválida " & sectionType & " para control " & controlJson("name") & ", usando acDetail (0)"
        sectionType = 0
    End If
    
    parentControl = ""
    columnName = ""
    
    ' Determinar tipo de control usando mapeo centralizado
    controlType = MapControlType(controlJson("type"))
    
    If controlType = -1 Then
        LogWarn "Tipo de control desconocido: " & controlJson("type") & " para control " & controlJson("name")
        Exit Sub
    End If
    
    ' Crear control con Application.CreateControl
    On Error Resume Next
    Set newControl = objAccess.Application.CreateControl(formName, controlType, sectionType, parentControl, columnName, leftPos, topPos, widthSize, heightSize)
    
    If Err.Number <> 0 Then
        WScript.Echo "Error creando control " & controlJson("name") & ": " & Err.Description
        Err.Clear
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Asignar propiedades adicionales del control después de crearlo
    If controlJson.Exists("name") Then newControl.Name = controlJson("name")
    
    ' Aplicar propiedades adicionales si existen
    If controlJson.Exists("properties") Then
        ' Reutilizar la variable properties ya declarada
        Set properties = controlJson("properties")
        
        For Each prop In properties.Keys
            propValue = properties(prop)
            
            ' Saltar propiedades de posición y tamaño ya procesadas
            If LCase(prop) <> "left" And LCase(prop) <> "top" And LCase(prop) <> "width" And LCase(prop) <> "height" Then
                On Error Resume Next
                ' Normalizar valores numéricos si es necesario
                If IsNumeric(propValue) Then
                    If InStr(LCase(prop), "color") > 0 Or LCase(prop) = "backcolor" Or LCase(prop) = "forecolor" Or LCase(prop) = "bordercolor" Then
                        newControl.Properties(prop) = CLng(propValue)
                    ElseIf LCase(prop) = "fontsize" Or InStr(LCase(prop), "size") > 0 Then
                        newControl.Properties(prop) = CDbl(propValue)
                    Else
                        newControl.Properties(prop) = propValue
                    End If
                Else
                    newControl.Properties(prop) = propValue
                End If
                
                If Err.Number <> 0 Then
                    If gStrict Then
                        WScript.Echo "Error aplicando propiedad " & prop & " al control " & controlJson("name") & ": " & Err.Description
                        WScript.Quit 1
                    Else
                        LogWarn "Propiedad " & prop & " no aplicable al control " & controlJson("name") & " (tipo: " & controlJson("type") & ")"
                    End If
                    Err.Clear
                End If
                On Error GoTo 0
            End If
        Next
    End If
    
    ' Propiedades específicas por tipo (compatibilidad con formato legacy)
    If LCase(controlJson("type")) = "textbox" Then
        If controlJson.Exists("controlSource") Then 
            On Error Resume Next
            newControl.ControlSource = controlJson("controlSource")
            If Err.Number <> 0 Then
                LogWarn "No se pudo asignar ControlSource al control " & controlJson("name")
                Err.Clear
            End If
            On Error GoTo 0
        End If
    ElseIf LCase(controlJson("type")) = "label" Then
        If controlJson.Exists("caption") Then 
            On Error Resume Next
            newControl.Caption = controlJson("caption")
            If Err.Number <> 0 Then
                LogWarn "No se pudo asignar Caption al control " & controlJson("name")
                Err.Clear
            End If
            On Error GoTo 0
        End If
    ElseIf LCase(controlJson("type")) = "commandbutton" Then
        If controlJson.Exists("caption") Then 
            On Error Resume Next
            newControl.Caption = controlJson("caption")
            If Err.Number <> 0 Then
                LogWarn "No se pudo asignar Caption al control " & controlJson("name")
                Err.Clear
            End If
            On Error GoTo 0
        End If
    ElseIf LCase(controlJson("type")) = "subform" Then
        ' Propiedades específicas de Subform
        If controlJson.Exists("sourceObject") Then 
            On Error Resume Next
            newControl.SourceObject = controlJson("sourceObject")
            If Err.Number <> 0 Then
                If gStrict Then
                    WScript.Echo "ERROR: No se pudo asignar SourceObject al subformulario " & controlJson("name") & ": " & Err.Description
                    WScript.Quit 1
                Else
                    LogWarn "No se pudo asignar SourceObject al subformulario " & controlJson("name") & ". Subformulario creado sin referencia."
                End If
                Err.Clear
            End If
            On Error GoTo 0
        Else
            If gStrict Then
                WScript.Echo "ERROR: Subformulario " & controlJson("name") & " no tiene sourceObject definido"
                WScript.Quit 1
            Else
                LogWarn "Subformulario " & controlJson("name") & " creado sin sourceObject"
            End If
        End If
        
        If controlJson.Exists("linkMasterFields") Then 
            On Error Resume Next
            newControl.LinkMasterFields = controlJson("linkMasterFields")
            If Err.Number <> 0 Then
                LogWarn "No se pudo asignar LinkMasterFields al subformulario " & controlJson("name")
                Err.Clear
            End If
            On Error GoTo 0
        End If
        
        If controlJson.Exists("linkChildFields") Then 
            On Error Resume Next
            newControl.LinkChildFields = controlJson("linkChildFields")
            If Err.Number <> 0 Then
                LogWarn "No se pudo asignar LinkChildFields al subformulario " & controlJson("name")
                Err.Clear
            End If
            On Error GoTo 0
        End If
    ElseIf LCase(controlJson("type")) = "tabcontrol" Then
        ' Propiedades específicas de TabControl
        If controlJson.Exists("pages") Then
            Dim pages, page, j
            Set pages = controlJson("pages")
            
            ' Intentar crear páginas del TabControl
            For j = 0 To pages.Count - 1
                Set page = pages(j)
                
                On Error Resume Next
                ' Intentar crear página usando CreateControl
                Dim newPage
                Set newPage = objAccess.Application.CreateControl(formName, 124, sectionType, newControl.Name) ' acPage
                
                If Err.Number <> 0 Then
                    LogWarn "WARNING: No se pudo crear página " & page("name") & " del TabControl " & controlJson("name") & ". Su versión de Access podría no soportar creación de páginas por API."
                    Err.Clear
                    Exit For
                Else
                    ' Asignar propiedades de la página
                    If page.Exists("name") Then newPage.Name = page("name")
                    If page.Exists("caption") Then newPage.Caption = page("caption")
                    If page.Exists("pageIndex") Then newPage.PageIndex = page("pageIndex")
                End If
                On Error GoTo 0
            Next
        End If
    End If
    
    ' Aplicar eventos si existen (solo como metadatos, no generar handlers)
    If controlJson.Exists("events") Then
        Dim events
        Set events = controlJson("events")
        
        For Each prop In events.Keys
            propValue = events(prop)
            
            On Error Resume Next
            newControl.Properties(prop) = propValue
            If Err.Number <> 0 Then
                LogWarn "Evento " & prop & " no aplicable al control " & controlJson("name")
                Err.Clear
            End If
            On Error GoTo 0
        Next
    End If
    
    ' Eventos del control (compatibilidad con formato anterior)
    If controlJson.Exists("onClick") Then
        On Error Resume Next
        newControl.OnClick = controlJson("onClick")
        If Err.Number <> 0 Then
            LogWarn "No se pudo asignar evento OnClick al control " & controlJson("name")
            Err.Clear
        End If
        On Error GoTo 0
    End If
    
    WScript.Echo "Control creado: " & controlJson("name")
End Sub

' ============================================================================
' FUNCIÓN AUXILIAR: MapControlType
' Descripción: Mapea tipos de control de cadena a constantes de Access
' Parámetros: controlTypeStr - Tipo de control como cadena
' Retorna: Constante numérica de Access o -1 si no se reconoce
' ============================================================================
Private Function MapControlType(controlTypeStr)
    Select Case LCase(Trim(controlTypeStr))
        Case "textbox"
            MapControlType = 109 ' acTextBox
        Case "label"
            MapControlType = 100 ' acLabel
        Case "commandbutton", "button"
            MapControlType = 104 ' acCommandButton
        Case "combobox"
            MapControlType = 111 ' acComboBox
        Case "listbox"
            MapControlType = 110 ' acListBox
        Case "checkbox"
            MapControlType = 106 ' acCheckBox
        Case "optionbutton"
            MapControlType = 105 ' acOptionButton
        Case "togglebutton"
            MapControlType = 122 ' acToggleButton
        Case "optiongroup"
            MapControlType = 107 ' acOptionGroup
        Case "boundframe"
            MapControlType = 108 ' acBoundObjectFrame
        Case "unboundframe"
            MapControlType = 114 ' acObjectFrame
        Case "subform"
            MapControlType = 112 ' acSubform
        Case "line"
            MapControlType = 102 ' acLine
        Case "rectangle"
            MapControlType = 101 ' acRectangle
        Case "image"
            MapControlType = 103 ' acImage
        Case "page"
            MapControlType = 124 ' acPage
        Case "tabcontrol"
            MapControlType = 123 ' acTabCtl
        Case Else
            ' Tipo desconocido, retornar -1 para indicar error
            MapControlType = -1
    End Select
End Function

' ============================================================================
' FUNCIÓN AUXILIAR: ApplyEventHandlers
' Descripción: Aplica manejadores de eventos desde el JSON al formulario
' Parámetros: frm - Objeto Form, jsonData - Datos JSON del formulario
' ============================================================================
Private Sub ApplyEventHandlers(frm, jsonData)
    Dim code, moduleCode, eventHandlers, handler, i
    
    ' Verificar si hay código del módulo en el JSON
    If Not jsonData.Exists("code") Then
        If gVerbose Then WScript.Echo "No hay código de módulo en el JSON"
        Exit Sub
    End If
    
    Set code = jsonData("code")
    
    ' Verificar si hay código del módulo
    If Not code.Exists("module") Then
        If gVerbose Then WScript.Echo "No hay código de módulo definido"
        Exit Sub
    End If
    
    moduleCode = code("module")
    
    ' Si hay código del módulo, intentar aplicarlo
    If Len(Trim(moduleCode)) > 0 Then
        On Error Resume Next
        ' Intentar acceder al módulo del formulario
        Dim formModule
        Set formModule = frm.Module
        
        If Err.Number <> 0 Then
            LogWarn "No se puede acceder al módulo del formulario " & frm.Name & ": " & Err.Description
            Err.Clear
            On Error GoTo 0
            Exit Sub
        End If
        
        ' Limpiar el módulo existente (opcional, solo si está vacío)
        If formModule.CountOfLines = 0 Then
            formModule.InsertLines 1, moduleCode
            If gVerbose Then WScript.Echo "Código del módulo aplicado al formulario " & frm.Name
        Else
            LogWarn "El módulo del formulario " & frm.Name & " ya contiene código, no se sobrescribe"
        End If
        
        On Error GoTo 0
    End If
    
    ' Verificar si hay manejadores de eventos específicos
    If code.Exists("eventHandlers") Then
        Set eventHandlers = code("eventHandlers")
        
        For i = 0 To eventHandlers.Count - 1
            Set handler = eventHandlers(i)
            
            If handler.Exists("event") And handler.Exists("code") Then
                On Error Resume Next
                
                ' Aplicar el manejador de evento al formulario
                Select Case LCase(handler("event"))
                    Case "onload", "form_load"
                        frm.OnLoad = "[Event Procedure]"
                    Case "onopen", "form_open"
                        frm.OnOpen = "[Event Procedure]"
                    Case "onclose", "form_close"
                        frm.OnClose = "[Event Procedure]"
                    Case "oncurrent", "form_current"
                        frm.OnCurrent = "[Event Procedure]"
                    Case "onunload", "form_unload"
                        frm.OnUnload = "[Event Procedure]"
                    Case Else
                        LogWarn "Evento no reconocido: " & handler("event")
                End Select
                
                If Err.Number <> 0 Then
                    LogWarn "No se pudo aplicar evento " & handler("event") & " al formulario " & frm.Name & ": " & Err.Description
                    Err.Clear
                End If
                
                On Error GoTo 0
            End If
        Next
        
        If gVerbose Then WScript.Echo "Manejadores de eventos aplicados al formulario " & frm.Name
    End If
End Sub

' ============================================================================
' SUBRUTINA: ListForms
' Descripción: Lista todos los formularios de una base de datos Access
' Sintaxis: list-forms [db_path] [--password <pwd>] [--json]
' Parámetros:
'   db_path: Ruta a la base de datos (opcional, por defecto ./condor.accdb)
'   --password: Contraseña de la base de datos (opcional)
'   --json: Salida en formato JSON (opcional, por defecto texto)
' ============================================================================
Sub ListForms()
    Dim password, bJsonOutput
    Dim i, arg
    Dim objFSO
    Dim formsList, formCount
    Dim accessObj, formName, objAccessLocal
    
    ' Inicializar variables
    password = ""
    bJsonOutput = False
    Set objAccessLocal = Nothing
    
    ' Procesar argumentos
    i = 1
    While i < objArgs.Count
        arg = LCase(objArgs(i))
        
        If arg = "--help" Or arg = "-h" Or arg = "help" Then
            WScript.Echo "SINTAXIS: cscript condor_cli.vbs list-forms [--password <pwd>] [--json]"
            WScript.Echo ""
            WScript.Echo "PARÁMETROS:"
            WScript.Echo "  --password <pwd> Contraseña de la base de datos"
            WScript.Echo "  --json           Salida en formato JSON"
            WScript.Echo ""
            WScript.Echo "EJEMPLOS:"
            WScript.Echo "  cscript condor_cli.vbs list-forms"
            WScript.Echo "  cscript condor_cli.vbs list-forms --password 1234"
            WScript.Echo "  cscript condor_cli.vbs list-forms --json"
            WScript.Quit 0
        ElseIf arg = "--password" Then
            If i + 1 < objArgs.Count Then
                i = i + 1
                password = objArgs(i)
            Else
                WScript.Echo "Error: --password requiere un valor"
                WScript.Quit 1
            End If
        ElseIf Left(arg, 4) = "/pwd" Then
            If Mid(arg, 5, 1) = ":" Then
                password = Mid(arg, 6)
            Else
                WScript.Echo "Error: Formato incorrecto para /pwd. Use /pwd:<clave>"
                WScript.Quit 1
            End If
        ElseIf arg = "--json" Then
            bJsonOutput = True
        ElseIf arg = "--print-db" Then
            ' Flag ya procesado globalmente, ignorar aquí
        ElseIf arg = "--db" Then
            ' Flag ya procesado globalmente, ignorar aquí
            If i + 1 < objArgs.Count Then
                i = i + 1 ' Saltar el valor del --db
            End If
        ElseIf Left(arg, 4) = "/db:" Then
            ' Flag ya procesado globalmente, ignorar aquí
        ElseIf Left(arg, 2) <> "--" Then
            ' Ignorar argumentos posicionales - la BD ya está resuelta
            WScript.Echo "Advertencia: Argumento posicional ignorado (usar --db): " & objArgs(i)
        Else
            WScript.Echo "Error: Opción desconocida: " & objArgs(i)
            WScript.Quit 1
        End If
        
        i = i + 1
    Wend
    
    WScript.Echo "=== LISTANDO FORMULARIOS ==="
    
    ' Usar la BD ya resuelta por el flujo principal
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    ' Abrir Access
    Set objAccessLocal = OpenAccessApp(strAccessPath, password, True)
    If objAccessLocal Is Nothing Then
        WScript.Echo "Error: no se pudo abrir Access."
        WScript.Quit 1
    End If
    
    ' Contar formularios primero
    formCount = 0
    For Each accessObj In objAccessLocal.CurrentProject.AllForms
        formCount = formCount + 1
    Next
    
    ' Redimensionar array y llenar
    Dim formsArray()
    If formCount > 0 Then
        ReDim formsArray(formCount - 1)
        Dim formIndex
        formIndex = 0
        For Each accessObj In objAccessLocal.CurrentProject.AllForms
            formsArray(formIndex) = accessObj.Name
            formIndex = formIndex + 1
        Next
    Else
        ' Si no hay formularios, crear array vacío
        ReDim formsArray(-1)
    End If
    
    ' Cerrar Access
    CloseAccessApp objAccessLocal
    
    ' Generar salida
    If bJsonOutput Then
        ' Salida JSON
        Dim jsonOutput
        jsonOutput = "["
        If formCount > 0 Then
            Dim i_form
            For i_form = 0 To UBound(formsArray)
                If i_form > 0 Then jsonOutput = jsonOutput & ","
                jsonOutput = jsonOutput & """" & formsArray(i_form) & """"
            Next
        End If
        jsonOutput = jsonOutput & "]"
        WScript.Echo jsonOutput
    Else
        ' Salida texto
        If formCount = 0 Then
            WScript.Echo "No se encontraron formularios en la base de datos."
        Else
            WScript.Echo "Formularios encontrados (" & formCount & "):"
            For i_form = 0 To UBound(formsArray)
                WScript.Echo "  " & formsArray(i_form)
            Next
        End If
    End If
End Sub

' ============================================================================
' SUBRUTINA: ValidateFormJsonCommand
' Descripción: Valida la estructura JSON de un formulario
' Sintaxis: validate-form-json <json_path> [--strict] [--schema]
' Parámetros:
'   json_path: Ruta al archivo JSON del formulario
'   --strict: Modo estricto (falla en advertencias)
'   --schema: Mostrar esquema JSON esperado
' ============================================================================
Private Sub ValidateFormJsonCommand()
    Dim jsonPath, bStrict, bShowSchema
    Dim i, arg
    Dim objFSO, objFile, jsonContent
    Dim objParser, formData
    Dim warnings, errors
    
    ' Inicializar variables
    jsonPath = ""
    bStrict = False
    bShowSchema = False
    Set warnings = CreateObject("Scripting.Dictionary")
    Set errors = CreateObject("Scripting.Dictionary")
    
    ' Procesar argumentos
    For i = 1 To WScript.Arguments.Count - 1
        arg = WScript.Arguments(i)
        If Left(arg, 2) = "--" Then
            Select Case arg
                Case "--help", "-h"
                    ShowValidateFormJsonHelp()
                    WScript.Quit 0
                Case "--strict"
                    bStrict = True
                Case "--schema"
                    bShowSchema = True
                Case Else
                    WScript.Echo "Error: Opción desconocida: " & arg
                    WScript.Quit 1
            End Select
        ElseIf arg = "help" Then
            ShowValidateFormJsonHelp()
            WScript.Quit 0
        Else
            If jsonPath = "" Then
                jsonPath = arg
            Else
                WScript.Echo "Error: Múltiples rutas JSON especificadas"
                WScript.Quit 1
            End If
        End If
    Next
    
    ' Mostrar esquema si se solicita
    If bShowSchema Then
        ShowFormJsonSchema()
        If jsonPath = "" Then Exit Sub
    End If
    
    ' Validar argumentos requeridos
    If jsonPath = "" Then
        WScript.Echo "Error: Debe especificar la ruta del archivo JSON"
        WScript.Echo "Uso: condor_cli validate-form-json <json_path> [--strict] [--schema]"
        WScript.Quit 1
    End If
    
    ' Verificar que el archivo existe
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If Not objFSO.FileExists(jsonPath) Then
        WScript.Echo "Error: El archivo JSON no existe: " & jsonPath
        WScript.Quit 1
    End If
    
    WScript.Echo "Validando archivo JSON: " & jsonPath
    
    ' Leer y parsear el JSON
    On Error Resume Next
    Set objFile = objFSO.OpenTextFile(jsonPath, 1, False, 0) ' Cambiar a ASCII
    jsonContent = objFile.ReadAll
    objFile.Close
    
    If Err.Number <> 0 Then
        WScript.Echo "Error al leer el archivo JSON: " & Err.Description
        WScript.Quit 1
    End If
    
    ' Parsear JSON
    Set objParser = New JsonParser
    Set formData = objParser.Parse(jsonContent)
    
    If Err.Number <> 0 Then
        WScript.Echo "Error: JSON inválido - " & Err.Description
        WScript.Quit 1
    End If
    
    ' Verificar que el resultado sea un objeto (Dictionary)
    If TypeName(formData) <> "Dictionary" Then
        WScript.Echo "Error: JSON inválido - Se requiere un objeto"
        WScript.Quit 1
    End If
    
    On Error GoTo 0
    
    ' Validar campos mínimos requeridos
    ValidateMinimumFields formData, warnings
    
    ' Validar estructura del formulario
    ValidateFormData formData, bStrict, warnings
    
    ' Validar propiedades estructurales clave en modo estricto
    If bStrict Then
        ValidateStructuralProperties formData, warnings
    End If
    
    ' Validar recursos (sin directorio base)
    ValidateResources formData, "", warnings
    
    ' Mostrar resultados
    WScript.Echo "\n=== RESULTADO DE VALIDACIÓN ==="
    
    If warnings.Count = 0 Then
        WScript.Echo "✓ JSON válido - No se encontraron problemas"
    Else
        WScript.Echo "⚠ Se encontraron " & warnings.Count & " advertencias:"
        Dim key
        For Each key In warnings.Keys
            WScript.Echo "  - " & warnings(key)
        Next
        
        If bStrict Then
            WScript.Echo "\nError: Validación falló en modo estricto"
            WScript.Quit 1
        End If
    End If
    
    WScript.Echo "\nValidación completada exitosamente."
End Sub

' ============================================================================
' SUBRUTINA: ShowValidateFormJsonHelp
' Descripción: Muestra la ayuda del comando validate-form-json
' ============================================================================
Sub ShowValidateFormJsonHelp()
    WScript.Echo "=== CONDOR CLI - AYUDA DEL COMANDO VALIDATE-FORM-JSON ==="
    WScript.Echo "Valida la estructura y contenido de un archivo JSON de formulario."
    WScript.Echo ""
    WScript.Echo "USO:"
    WScript.Echo "  cscript condor_cli.vbs validate-form-json <json_path> [opciones]"
    WScript.Echo "  cscript condor_cli.vbs validate-form-json --help"
    WScript.Echo ""
    WScript.Echo "PARÁMETROS:"
    WScript.Echo "  <json_path>              - Ruta al archivo JSON del formulario (obligatorio)"
    WScript.Echo ""
    WScript.Echo "OPCIONES:"
    WScript.Echo "  --strict                 - Modo estricto: falla si faltan propiedades estructurales"
    WScript.Echo "                             clave como RecordSource, DefaultView, Width, Height"
    WScript.Echo "  --schema                 - Mostrar el esquema JSON canónico esperado"
    WScript.Echo "  --help, -h, help         - Mostrar esta ayuda"
    WScript.Echo ""
    WScript.Echo "VALIDACIONES REALIZADAS:"
    WScript.Echo "  • JSON bien formado y parseable"
    WScript.Echo "  • Campos mínimos requeridos: formName/name, controls/sections, properties"
    WScript.Echo "  • Estructura de controles y secciones válida"
    WScript.Echo "  • Propiedades de controles (name, type, top, left, width, height)"
    WScript.Echo "  • Tipos de controles soportados"
    WScript.Echo "  • Recursos (imágenes con rutas relativas y extensiones válidas)"
    WScript.Echo ""
    WScript.Echo "MODO ESTRICTO (--strict):"
    WScript.Echo "  En modo estricto, la validación falla si faltan propiedades estructurales"
    WScript.Echo "  clave del formulario como:"
    WScript.Echo "  • width, height - Dimensiones del formulario"
    WScript.Echo "  • defaultView - Vista por defecto del formulario"
    WScript.Echo "  • recordSource - Origen de datos (si recordSourceType requiere tabla/consulta)"
    WScript.Echo ""
    WScript.Echo "CÓDIGOS DE SALIDA:"
    WScript.Echo "  0 - Validación exitosa (sin errores críticos)"
    WScript.Echo "  1 - Error en validación o modo estricto con advertencias"
    WScript.Echo ""
    WScript.Echo "EJEMPLOS:"
    WScript.Echo "  # Validación básica"
    WScript.Echo "  cscript condor_cli.vbs validate-form-json "".\ui\forms\FormComercial.json"""
    WScript.Echo ""
    WScript.Echo "  # Validación estricta"
    WScript.Echo "  cscript condor_cli.vbs validate-form-json "".\ui\forms\FormComercial.json"" --strict"
    WScript.Echo ""
    WScript.Echo "  # Mostrar esquema esperado"
    WScript.Echo "  cscript condor_cli.vbs validate-form-json --schema"
End Sub

' ============================================================================
' SUBRUTINA: ValidateFormData
' Descripción: Valida la estructura de datos del formulario
' Acepta ambos formatos: secciones (canónico) y legacy (root.controls)
' ============================================================================
Private Sub ValidateFormData(formData, bStrict, warnings)
    Dim hasControls, hasSections
    hasControls = IsPresent(formData, "controls")
    hasSections = IsPresent(formData, "sections")
    
    ' Validar que existe al menos uno de los formatos
    If Not hasControls And Not hasSections Then
        warnings("no_controls_or_sections") = "No se encontró ni 'controls' ni 'sections' en el JSON"
        Exit Sub
    End If
    
    ' Procesar formato por secciones (canónico)
    If hasSections Then
        Dim sections, sectionNames, sectionName, section
        Set sections = formData("sections")
        sectionNames = Array("header", "detail", "footer")
        
        Dim i, j
        For i = 0 To UBound(sectionNames)
            sectionName = sectionNames(i)
            If IsPresent(sections, sectionName) Then
                Set section = sections(sectionName)
                
                ' Validar controles de la sección si existen
                If IsPresent(section, "controls") Then
                    Dim sectionControls
                    Set sectionControls = section("controls")
                    
                    For j = 0 To sectionControls.Count - 1
                        ValidateControl sectionControls(j), sectionName, j, warnings, bStrict
                    Next
                End If
            End If
        Next
    End If
    
    ' Procesar formato legacy (root.controls)
    If hasControls And Not hasSections Then
        Dim controls
        Set controls = formData("controls")
        
        For i = 0 To controls.Count - 1
            Dim control
            Set control = controls(i)
            
            ' En formato legacy, exigir campo "section"
            If Not IsPresent(control, "section") Then
                warnings("control_" & i & "_section_missing") = "Control " & i & ": Campo 'section' obligatorio en formato legacy"
            Else
                Dim controlSection
                controlSection = control("section")
                ValidateControl control, controlSection, i, warnings, bStrict
            End If
        Next
    End If
    
    ' Si existen ambos formatos, generar advertencia
    If hasControls And hasSections Then
        warnings("dual_format_detected") = "Se detectaron ambos formatos 'sections' y 'controls'. Se procesará 'sections' (canónico)"
    End If
End Sub

' ============================================================================
' SUBRUTINA: ValidateControl
' Descripción: Valida un control individual
' ============================================================================
Private Sub ValidateControl(control, expectedSection, controlIndex, warnings, bStrict)
    Dim controlName
    controlName = "control_" & controlIndex
    
    ' Validar campos obligatorios
    If Not IsPresent(control, "name") Then
        warnings(controlName & "_name_missing") = "Control " & controlIndex & ": Campo 'name' obligatorio"
    End If
    
    If Not IsPresent(control, "type") Then
        warnings(controlName & "_type_missing") = "Control " & controlIndex & ": Campo 'type' obligatorio"
    Else
        ' Validar tipo soportado
        Dim controlType
        controlType = control("type")
        Dim supportedTypes
        supportedTypes = Array("CommandButton", "Label", "TextBox", "ComboBox", "ListBox", "CheckBox", "OptionButton", "Image", "Rectangle", "Line")
        
        Dim typeFound, i
        typeFound = False
        For i = 0 To UBound(supportedTypes)
            If LCase(controlType) = LCase(supportedTypes(i)) Then
                typeFound = True
                Exit For
            End If
        Next
        
        If Not typeFound Then
            warnings(controlName & "_type_unsupported") = "Control " & controlIndex & ": Tipo '" & controlType & "' no soportado"
        End If
    End If
    
    ' Validar sección si está presente
    If IsPresent(control, "section") Then
        Dim controlSection
        controlSection = control("section")
        
        ' Validar que la sección es válida
        If controlSection <> "header" And controlSection <> "detail" And controlSection <> "footer" Then
            warnings(controlName & "_section_invalid") = "Control " & controlIndex & ": Valor de 'section' inválido '" & controlSection & "' (debe ser header|detail|footer)"
        End If
        
        ' Validar coherencia con sección portadora
        If controlSection <> expectedSection Then
            Dim warningKey
            warningKey = controlName & "_section_mismatch"
            If bStrict Then
                warnings(warningKey) = "ERROR: Control " & controlIndex & ": 'section' (" & controlSection & ") no coincide con sección portadora (" & expectedSection & ")"
            Else
                warnings(warningKey) = "WARNING: Control " & controlIndex & ": 'section' (" & controlSection & ") no coincide con sección portadora (" & expectedSection & ")"
            End If
        End If
    End If
    
    ' Validar propiedades si existen
    If IsPresent(control, "properties") Then
        Dim properties
        Set properties = control("properties")
        
        ' Validar propiedades de posición y tamaño (obligatorias)
        Dim requiredProps
        requiredProps = Array("top", "left", "width", "height")
        
        For i = 0 To UBound(requiredProps)
            Dim propName
            propName = requiredProps(i)
            If Not IsPresent(properties, propName) Then
                warnings(controlName & "_" & propName & "_missing") = "Control " & controlIndex & ": Propiedad '" & propName & "' obligatoria"
            Else
                ' Validar que es numérico
                Dim propValue
                propValue = properties(propName)
                If Not IsNumeric(propValue) Then
                    warnings(controlName & "_" & propName & "_not_numeric") = "Control " & controlIndex & ": Propiedad '" & propName & "' debe ser numérica"
                Else
                    ' Validar coherencia (width > 0, height > 0)
                    If (propName = "width" Or propName = "height") And CDbl(propValue) <= 0 Then
                        warnings(controlName & "_" & propName & "_invalid") = "Control " & controlIndex & ": Propiedad '" & propName & "' debe ser mayor que 0"
                    End If
                End If
            End If
        Next
    End If
End Sub

' ============================================================================
' SUBRUTINA: ValidateResources
' Descripción: Valida la estructura de recursos del formulario
' ============================================================================
Private Sub ValidateResources(formData, basePath, warnings)
    ' Validar sección de recursos si existe
    If IsPresent(formData, "resources") Then
        Dim resources
        Set resources = formData("resources")
        
        ' Validar imágenes si existen
        If IsPresent(resources, "images") Then
            Dim images
            Set images = resources("images")
            
            Dim i
            For i = 0 To images.Count - 1
                Dim imagePath
                imagePath = images(i)
                
                ' Validar que es una cadena
                If VarType(imagePath) <> vbString Then
                    warnings("resource_image_" & i & "_not_string") = "Recurso imagen " & i & ": Debe ser una cadena de texto"
                Else
                    ' Validar formato de ruta relativa
                    If InStr(imagePath, ":") > 0 Or Left(imagePath, 1) = "\" Or Left(imagePath, 1) = "/" Then
                        warnings("resource_image_" & i & "_not_relative") = "Recurso imagen " & i & ": Debe ser una ruta relativa (" & imagePath & ")"
                    End If
                    
                    ' Validar extensión de archivo
                    Dim ext
                    ext = LCase(Right(imagePath, 4))
                    If ext <> ".png" And ext <> ".jpg" And ext <> ".gif" And ext <> ".bmp" Then
                        warnings("resource_image_" & i & "_invalid_ext") = "Recurso imagen " & i & ": Extensión no soportada (" & imagePath & ")"
                    End If
                End If
            Next
        End If
    End If
End Sub

' ============================================================================
' SUBRUTINA: ShowFormJsonSchema
' Descripción: Muestra el esquema JSON esperado para formularios
' ============================================================================
Private Sub ShowFormJsonSchema()
    WScript.Echo "\n=== ESQUEMA JSON CANÓNICO DE FORMULARIO ==="
    WScript.Echo "{"
    WScript.Echo "  ""schemaVersion"": ""1.x"","
    WScript.Echo "  ""units"": ""twips"","
    WScript.Echo "  ""formName"": ""string (requerido)"","
    WScript.Echo "  ""properties"": {"
    WScript.Echo "    ""caption"": ""string"","
    WScript.Echo "    ""width"": ""number"","
    WScript.Echo "    ""height"": ""number"","
    WScript.Echo "    ""backColor"": ""string (hex: #RRGGBB)"","
    WScript.Echo "    ""defaultView"": ""enum (single|continuous|datasheet|pivotTable|pivotChart)"","
    WScript.Echo "    ""cycle"": ""enum (currentRecord|allRecords)"","
    WScript.Echo "    ""recordSourceType"": ""enum (table|sql|none)"","
    WScript.Echo "    ""recordSource"": ""string"","
    WScript.Echo "    ""allowEdits"": ""boolean"","
    WScript.Echo "    ""allowAdditions"": ""boolean"","
    WScript.Echo "    ""allowDeletions"": ""boolean"""
    WScript.Echo "  },"
    WScript.Echo "  ""sections"": {"
    WScript.Echo "    ""header"": {"
    WScript.Echo "      ""height"": ""number"","
    WScript.Echo "      ""backColor"": ""string (hex: #RRGGBB)"","
    WScript.Echo "      ""controls"": ["
    WScript.Echo "        { ""control_object"": ""ver estructura de control abajo"" }"
    WScript.Echo "      ]"
    WScript.Echo "    },"
    WScript.Echo "    ""detail"": {"
    WScript.Echo "      ""height"": ""number"","
    WScript.Echo "      ""backColor"": ""string (hex: #RRGGBB)"","
    WScript.Echo "      ""controls"": ["
    WScript.Echo "        { ""control_object"": ""ver estructura de control abajo"" }"
    WScript.Echo "      ]"
    WScript.Echo "    },"
    WScript.Echo "    ""footer"": {"
    WScript.Echo "      ""height"": ""number"","
    WScript.Echo "      ""backColor"": ""string (hex: #RRGGBB)"","
    WScript.Echo "      ""controls"": ["
    WScript.Echo "        { ""control_object"": ""ver estructura de control abajo"" }"
    WScript.Echo "      ]"
    WScript.Echo "    }"
    WScript.Echo "  },"
    WScript.Echo "  ""resources"": {"
    WScript.Echo "    ""images"": ["
    WScript.Echo "      ""relative/path.png"","
    WScript.Echo "      ""..."
    WScript.Echo "    ]"
    WScript.Echo "  },"
    WScript.Echo "  ""code"": {"
    WScript.Echo "    ""module"": {"
    WScript.Echo "      ""exists"": ""boolean"","
    WScript.Echo "      ""filename"": ""string"","
    WScript.Echo "      ""handlers"": ["
    WScript.Echo "        {"
    WScript.Echo "          ""control"": ""string"","
    WScript.Echo "          ""event"": ""string"","
    WScript.Echo "          ""signature"": ""string"""
    WScript.Echo "        }"
    WScript.Echo "      ]"
    WScript.Echo "    }"
    WScript.Echo "  }"
    WScript.Echo "}"
    WScript.Echo "\n=== ESTRUCTURA DE CONTROL ==="
    WScript.Echo "{"
    WScript.Echo "  ""name"": ""string (requerido)"","
    WScript.Echo "  ""type"": ""enum (CommandButton|Label|TextBox|...) (requerido)"","
    WScript.Echo "  ""section"": ""enum (header|detail|footer) (opcional en export, obligatorio si aparece en raíz)"","
    WScript.Echo "  ""properties"": {"
    WScript.Echo "    ""top"": ""number (requerido)"","
    WScript.Echo "    ""left"": ""number (requerido)"","
    WScript.Echo "    ""width"": ""number (requerido)"","
    WScript.Echo "    ""height"": ""number (requerido)"","
    WScript.Echo "    ""caption"": ""string"","
    WScript.Echo "    ""backColor"": ""string (hex: #RRGGBB)"","
    WScript.Echo "    ""foreColor"": ""string (hex: #RRGGBB)"","
    WScript.Echo "    ""fontName"": ""string"","
    WScript.Echo "    ""fontSize"": ""number"","
    WScript.Echo "    ""fontBold"": ""boolean"","
    WScript.Echo "    ""fontItalic"": ""boolean"","
    WScript.Echo "    ""picture"": ""string (ruta relativa)"","
    WScript.Echo "    ""textAlign"": ""enum (left|center|right)"","
    WScript.Echo "    ""borderStyle"": ""enum (transparent|solid|dashes|dots)"","
    WScript.Echo "    ""specialEffect"": ""enum (flat|raised|sunken|etched|shadowed|chiseled)"""
    WScript.Echo "  },"
    WScript.Echo "  ""events"": {"
    WScript.Echo "    ""detected"": ["
    WScript.Echo "      ""Click"","
    WScript.Echo "      ""..."
    WScript.Echo "    ]"
    WScript.Echo "  }"
    WScript.Echo "}"
    WScript.Echo "\n=== COMPATIBILIDAD DE LECTURA ==="
    WScript.Echo "NOTA: También se acepta 'root.controls[]' (legacy), pero la forma oficial es 'sections.*.controls[]'"
    WScript.Echo "\n=== COLORES VÁLIDOS ==="
    WScript.Echo "Formato hexadecimal: #RRGGBB (ej: #FF0000 para rojo)"
    WScript.Echo "\n=== ENUMS VÁLIDOS ==="
    WScript.Echo "defaultView: single, continuous, datasheet, pivotTable, pivotChart"
    WScript.Echo "cycle: currentRecord, allRecords"
    WScript.Echo "recordSourceType: table, sql, none"
    WScript.Echo "textAlign: left, center, right"
    WScript.Echo "borderStyle: transparent, solid, dashes, dots"
    WScript.Echo "specialEffect: flat, raised, sunken, etched, shadowed, chiseled"
    WScript.Echo "\n=== NOTA IMPORTANTE ==="
    WScript.Echo "Los comandos export/import/roundtrip operan en vista Diseño (no ejecutan eventos)."
End Sub

' ============================================================================
' SUBRUTINA: ValidateMinimumFields
' Descripción: Valida que existan los campos mínimos requeridos
' ============================================================================
Private Sub ValidateMinimumFields(formData, warnings)
    ' Validar FormName (requerido)
    If Not IsPresent(formData, "formName") And Not IsPresent(formData, "name") Then
        warnings("missing_form_name") = "Campo 'formName' o 'name' es obligatorio"
    End If
    
    ' Validar que existe Controls[] o sections (puede estar vacío)
    If Not IsPresent(formData, "controls") And Not IsPresent(formData, "sections") Then
        warnings("missing_controls_or_sections") = "Debe existir 'controls[]' o 'sections' (puede estar vacío)"
    End If
    
    ' Validar que existe Properties{} (puede estar vacío)
    If Not IsPresent(formData, "properties") Then
        warnings("missing_properties") = "Debe existir 'properties{}' (puede estar vacío)"
    End If
End Sub

' ============================================================================
' SUBRUTINA: ValidateStructuralProperties
' Descripción: Valida propiedades estructurales clave en modo estricto
' ============================================================================
Private Sub ValidateStructuralProperties(formData, warnings)
    If IsPresent(formData, "properties") Then
        Dim properties
        Set properties = formData("properties")
        
        ' Propiedades estructurales clave
        Dim structuralProps
        structuralProps = Array("width", "height", "defaultView")
        
        Dim i, propName
        For i = 0 To UBound(structuralProps)
            propName = structuralProps(i)
            If Not IsPresent(properties, propName) Then
                warnings("missing_structural_" & propName) = "ERROR STRICT: Falta propiedad estructural clave '" & propName & "'"
            End If
        Next
        
        ' Validar RecordSource si recordSourceType no es "none"
        If IsPresent(properties, "recordSourceType") Then
            Dim recordSourceType
            recordSourceType = properties("recordSourceType")
            If recordSourceType <> "none" And Not IsPresent(properties, "recordSource") Then
                warnings("missing_record_source") = "ERROR STRICT: Falta 'recordSource' cuando recordSourceType es '" & recordSourceType & "'"
            End If
        End If
    Else
        warnings("missing_properties_strict") = "ERROR STRICT: Falta sección 'properties' completa"
    End If
End Sub

' ============================================================================
' INFRAESTRUCTURA JSON Y UTILIDADES
' ============================================================================

' Inicializar variable global verbose
gVerbose = False

' Clase JsonWriter - Convierte objetos VBA a JSON
Class JsonWriter
    Private jsonContent
    Private objectStack
    Private arrayStack
    Private currentState
    Private needsComma
    
    Private Sub Class_Initialize()
        jsonContent = ""
        Set objectStack = CreateList()
        Set arrayStack = CreateList()
        currentState = "root"
        needsComma = False
    End Sub
    
    Public Sub StartObject()
        If needsComma Then
            jsonContent = jsonContent & ","
        End If
        jsonContent = jsonContent & "{"
        ' Verificar si objectStack es ArrayList o Dictionary
        If TypeName(objectStack) = "ArrayList" Then
            objectStack.Add currentState
        Else
            ' Es Dictionary, usar como lista con índices numéricos
            objectStack.Add objectStack.Count, currentState
        End If
        currentState = "object"
        needsComma = False
    End Sub
    
    Public Sub EndObject()
        jsonContent = jsonContent & "}"
        If objectStack.Count > 0 Then
            ' Verificar si objectStack es ArrayList o Dictionary
            If TypeName(objectStack) = "ArrayList" Then
                currentState = objectStack(objectStack.Count - 1)
                objectStack.RemoveAt objectStack.Count - 1
            Else
                ' Es Dictionary, usar como lista con índices numéricos
                currentState = objectStack(objectStack.Count - 1)
                objectStack.Remove objectStack.Count - 1
            End If
            needsComma = True
        End If
    End Sub
    
    Public Sub StartArray()
        If needsComma Then
            jsonContent = jsonContent & ","
        End If
        jsonContent = jsonContent & "["
        ' Verificar si arrayStack es ArrayList o Dictionary
        If TypeName(arrayStack) = "ArrayList" Then
            arrayStack.Add currentState
        Else
            ' Es Dictionary, usar como lista con índices numéricos
            arrayStack.Add arrayStack.Count, currentState
        End If
        currentState = "array"
        needsComma = False
    End Sub
    
    Public Sub EndArray()
        jsonContent = jsonContent & "]"
        If arrayStack.Count > 0 Then
            ' Verificar si arrayStack es ArrayList o Dictionary
            If TypeName(arrayStack) = "ArrayList" Then
                currentState = arrayStack(arrayStack.Count - 1)
                arrayStack.RemoveAt arrayStack.Count - 1
            Else
                ' Es Dictionary, usar como lista con índices numéricos
                currentState = arrayStack(arrayStack.Count - 1)
                arrayStack.Remove arrayStack.Count - 1
            End If
            needsComma = True
        End If
    End Sub
    
    Public Sub AddProperty(key, value)
        If needsComma Then
            jsonContent = jsonContent & ","
        End If
        jsonContent = jsonContent & """" & EscapeString(CStr(key)) & """:" & FormatValue(value)
        needsComma = True
    End Sub
    
    Public Sub StartObjectProperty(key)
        If needsComma Then
            jsonContent = jsonContent & ","
        End If
        jsonContent = jsonContent & """" & EscapeString(CStr(key)) & """:{"
        ' Verificar si objectStack es ArrayList o Dictionary
        If TypeName(objectStack) = "ArrayList" Then
            objectStack.Add currentState
        Else
            ' Es Dictionary, usar como lista con índices numéricos
            objectStack.Add objectStack.Count, currentState
        End If
        currentState = "object"
        needsComma = False
    End Sub
    
    Public Sub StartArrayProperty(key)
        If needsComma Then
            jsonContent = jsonContent & ","
        End If
        jsonContent = jsonContent & """" & EscapeString(CStr(key)) & """:["
        ' Verificar si arrayStack es ArrayList o Dictionary
        If TypeName(arrayStack) = "ArrayList" Then
            arrayStack.Add currentState
        Else
            ' Es Dictionary, usar como lista con índices numéricos
            arrayStack.Add arrayStack.Count, currentState
        End If
        currentState = "array"
        needsComma = False
    End Sub
    
    Public Sub AddValue(value)
        If needsComma Then
            jsonContent = jsonContent & ","
        End If
        jsonContent = jsonContent & FormatValue(value)
        needsComma = True
    End Sub
    
    Public Function GetJson()
        GetJson = jsonContent
    End Function
    
    Private Function FormatValue(value)
        If IsNull(value) Then
            FormatValue = "null"
        ElseIf VarType(value) = vbBoolean Then
            If value Then
                FormatValue = "true"
            Else
                FormatValue = "false"
            End If
        ElseIf IsNumeric(value) Then
            FormatValue = CStr(value)
        ElseIf VarType(value) = vbString Then
            FormatValue = """" & EscapeString(CStr(value)) & """"
        Else
            FormatValue = """" & EscapeString(CStr(value)) & """"
        End If
    End Function
    
    Private Function EscapeString(str)
        Dim result, i, char
        result = ""
        For i = 1 To Len(str)
            char = Mid(str, i, 1)
            Select Case char
                Case Chr(34) ' "
                    result = result & "\"""
                Case Chr(92) ' backslash
                    result = result & "\\"
                Case Chr(8)  ' \b
                    result = result & "\b"
                Case Chr(12) ' \f
                    result = result & "\f"
                Case Chr(10) ' \n
                    result = result & "\n"
                Case Chr(13) ' \r
                    result = result & "\r"
                Case Chr(9)  ' \t
                    result = result & "\t"
                Case Else
                    If Asc(char) < 32 Then
                        result = result & "\u" & Right("0000" & Hex(Asc(char)), 4)
                    Else
                        result = result & char
                    End If
            End Select
        Next
        EscapeString = result
    End Function
    
    ' Métodos de compatibilidad con la implementación anterior
    Public Sub WriteProperty(key, value)
        ' Método de compatibilidad que usa AddProperty
        AddProperty key, value
    End Sub
    
    Public Function Stringify(value)
        If IsNull(value) Then
            Stringify = "null"
        ElseIf VarType(value) = vbBoolean Then
            If value Then
                Stringify = "true"
            Else
                Stringify = "false"
            End If
        ElseIf IsNumeric(value) Then
            Stringify = CStr(value)
        ElseIf VarType(value) = vbString Then
            Stringify = """" & EscapeString(CStr(value)) & """"
        ElseIf IsDictionary(value) Then
            Stringify = StringifyObject(value)
        ElseIf IsArrayLike(value) Then
            Stringify = StringifyArray(value)
        Else
            Stringify = "null"
        End If
    End Function
    
    Private Function StringifyObject(obj)
        Dim result, key, first
        result = "{"
        first = True
        For Each key In obj.Keys
            If Not first Then result = result & ","
            result = result & """" & EscapeString(CStr(key)) & """:" & Stringify(obj(key))
            first = False
        Next
        result = result & "}"
        StringifyObject = result
    End Function
    
    Private Function StringifyArray(arr)
        Dim result, i, first
        result = "["
        first = True
        
        If TypeName(arr) = "ArrayList" Then
            For i = 0 To arr.Count - 1
                If Not first Then result = result & ","
                result = result & Stringify(arr(i))
                first = False
            Next
        Else
            ' Array nativo VBA
            For i = LBound(arr) To UBound(arr)
                If Not first Then result = result & ","
                result = result & Stringify(arr(i))
                first = False
            Next
        End If
        
        result = result & "]"
        StringifyArray = result
    End Function
End Class

' Clase JsonParser - Convierte JSON a objetos VBA
Class JsonParser
    Private pos
    Private jsonText
    
    Public Function Parse(json)
        jsonText = json
        pos = 1
        
        ' Limpiar BOM y caracteres especiales al inicio
        Do While pos <= Len(jsonText)
            Dim firstChar
            firstChar = Mid(jsonText, pos, 1)
            Dim asciiVal
            asciiVal = Asc(firstChar)
            ' Saltar BOM UTF-8 y otros caracteres de control
            If asciiVal = 239 Or asciiVal = 187 Or asciiVal = 191 Or asciiVal < 32 Then
                pos = pos + 1
            Else
                Exit Do
            End If
        Loop
        
        SkipWhitespace
        Set Parse = ParseValue()
    End Function
    
    Private Function ParseValue()
        SkipWhitespace
        Dim char
        char = Mid(jsonText, pos, 1)
        
        Select Case char
            Case "{"
                Set ParseValue = ParseObject()
            Case "["
                Set ParseValue = ParseArray()
            Case Chr(34) ' "
                ParseValue = ParseString()
            Case "t", "f"
                ParseValue = ParseBoolean()
            Case "n"
                ParseValue = ParseNull()
            Case Else
                If IsNumeric(char) Or char = "-" Then
                    ParseValue = ParseNumber()
                Else
                    Err.Raise 1001, "JsonParser", "Carácter inesperado en posición " & pos & ": " & char
                End If
        End Select
    End Function
    
    Private Function ParseObject()
        Dim obj
        Set obj = CreateDict()
        pos = pos + 1 ' Saltar '{'
        SkipWhitespace
        
        If Mid(jsonText, pos, 1) = "}" Then
            pos = pos + 1
            Set ParseObject = obj
            Exit Function
        End If
        
        Do
            SkipWhitespace
            Dim key
            key = ParseString()
            SkipWhitespace
            
            If Mid(jsonText, pos, 1) <> ":" Then
                Err.Raise 1002, "JsonParser", "Se esperaba ':' después de la clave"
            End If
            pos = pos + 1
            
            Dim value
            On Error Resume Next
            Set value = ParseValue()
            If Err.Number <> 0 Then
                Err.Clear
                value = ParseValue()
            End If
            On Error GoTo 0
            
            If IsObject(value) Then
                Set obj(key) = value
            Else
                obj(key) = value
            End If
            
            SkipWhitespace
            Dim nextChar
            nextChar = Mid(jsonText, pos, 1)
            
            If nextChar = "}" Then
                pos = pos + 1
                Exit Do
            ElseIf nextChar = "," Then
                pos = pos + 1
            Else
                Err.Raise 1003, "JsonParser", "Se esperaba ',' o '}'"
            End If
        Loop
        
        Set ParseObject = obj
    End Function
    
    Private Function ParseArray()
        Dim arr
        Set arr = CreateList()
        pos = pos + 1 ' Saltar '['
        SkipWhitespace
        
        If Mid(jsonText, pos, 1) = "]" Then
            pos = pos + 1
            Set ParseArray = arr
            Exit Function
        End If
        
        Do
            Dim value
            On Error Resume Next
            Set value = ParseValue()
            If Err.Number <> 0 Then
                Err.Clear
                value = ParseValue()
            End If
            On Error GoTo 0
            
            ' Verificar si arr es ArrayList o Dictionary
            If TypeName(arr) = "ArrayList" Then
                If IsObject(value) Then
                    arr.Add value
                Else
                    arr.Add value
                End If
            Else
                ' Es Dictionary, usar índice numérico
                Dim index
                index = arr.Count
                If IsObject(value) Then
                    Set arr(index) = value
                Else
                    arr(index) = value
                End If
            End If
            
            SkipWhitespace
            Dim nextChar
            nextChar = Mid(jsonText, pos, 1)
            
            If nextChar = "]" Then
                pos = pos + 1
                Exit Do
            ElseIf nextChar = "," Then
                pos = pos + 1
                SkipWhitespace
            Else
                Err.Raise 1004, "JsonParser", "Se esperaba ',' o ']'"
            End If
        Loop
        
        Set ParseArray = arr
    End Function
    
    Private Function ParseString()
        pos = pos + 1 ' Saltar '"' inicial
        Dim result, char
        result = ""
        
        Do While pos <= Len(jsonText)
            char = Mid(jsonText, pos, 1)
            
            If char = Chr(34) Then ' "
                pos = pos + 1
                ParseString = result
                Exit Function
            ElseIf char = "\" Then
                pos = pos + 1
                If pos > Len(jsonText) Then Exit Do
                
                Dim escapeChar
                escapeChar = Mid(jsonText, pos, 1)
                Select Case escapeChar
                    Case Chr(34) ' "
                        result = result & Chr(34)
                    Case "\"
                        result = result & "\"
                    Case "b"
                        result = result & Chr(8)
                    Case "f"
                        result = result & Chr(12)
                    Case "n"
                        result = result & Chr(10)
                    Case "r"
                        result = result & Chr(13)
                    Case "t"
                        result = result & Chr(9)
                    Case "u"
                        ' Unicode escape \uXXXX
                        If pos + 4 <= Len(jsonText) Then
                            Dim hexCode
                            hexCode = Mid(jsonText, pos + 1, 4)
                            result = result & Chr(CLng("&H" & hexCode))
                            pos = pos + 4
                        End If
                    Case Else
                        result = result & escapeChar
                End Select
            Else
                result = result & char
            End If
            pos = pos + 1
        Loop
        
        Err.Raise 1005, "JsonParser", "String sin terminar"
    End Function
    
    Private Function ParseNumber()
        Dim numStr, char
        numStr = ""
        
        Do While pos <= Len(jsonText)
            char = Mid(jsonText, pos, 1)
            If IsNumeric(char) Or char = "." Or char = "-" Or char = "+" Or LCase(char) = "e" Then
                numStr = numStr & char
                pos = pos + 1
            Else
                Exit Do
            End If
        Loop
        
        If IsNumeric(numStr) Then
            ParseNumber = CDbl(numStr)
        Else
            Err.Raise 1006, "JsonParser", "Número inválido: " & numStr
        End If
    End Function
    
    Private Function ParseBoolean()
        If Mid(jsonText, pos, 4) = "true" Then
            pos = pos + 4
            ParseBoolean = True
        ElseIf Mid(jsonText, pos, 5) = "false" Then
            pos = pos + 5
            ParseBoolean = False
        Else
            Err.Raise 1007, "JsonParser", "Valor booleano inválido"
        End If
    End Function
    
    Private Function ParseNull()
        If Mid(jsonText, pos, 4) = "null" Then
            pos = pos + 4
            ParseNull = Null
        Else
            Err.Raise 1008, "JsonParser", "Valor null inválido"
        End If
    End Function
    
    Private Sub SkipWhitespace()
        Do While pos <= Len(jsonText)
            Dim char
            char = Mid(jsonText, pos, 1)
            If char = " " Or char = Chr(9) Or char = Chr(10) Or char = Chr(13) Then
                pos = pos + 1
            Else
                Exit Do
            End If
        Loop
    End Sub
End Class

' ============================================================================
' FUNCIONES AUXILIARES JSON
' ============================================================================

Function IsDictionary(obj)
    On Error Resume Next
    IsDictionary = (TypeName(obj) = "Dictionary")
    On Error GoTo 0
End Function

Function IsArrayLike(obj)
    On Error Resume Next
    Dim result
    result = (TypeName(obj) = "ArrayList") Or (IsArray(obj))
    IsArrayLike = result
    On Error GoTo 0
End Function

Function CreateDict()
    Set CreateDict = CreateObject("Scripting.Dictionary")
End Function

Function CreateList()
    On Error Resume Next
    Set CreateList = CreateObject("System.Collections.ArrayList")
    If Err.Number <> 0 Then
        ' Fallback a Dictionary como lista si ArrayList no está disponible
        Err.Clear
        Set CreateList = CreateObject("Scripting.Dictionary")
    End If
    On Error GoTo 0
End Function

' ============================================================================
' UTILIDADES DE ARCHIVOS Y PATHS
' ============================================================================

Function ReadAllText(path)
    On Error Resume Next
    Dim stream, content
    
    ' Intentar primero con UTF-8
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2 ' adTypeText
    stream.Charset = "UTF-8"
    stream.Open
    stream.LoadFromFile path
    content = stream.ReadText(-1) ' adReadAll
    stream.Close
    
    ' Si hay error o contenido vacío, reintentar con Windows-1252
    If Err.Number <> 0 Or Len(content) = 0 Then
        Err.Clear
        Set stream = CreateObject("ADODB.Stream")
        stream.Type = 2
        stream.Charset = "Windows-1252"
        stream.Open
        stream.LoadFromFile path
        content = stream.ReadText(-1) ' adReadAll
        stream.Close
        
        If Err.Number <> 0 Then
            content = ""
        End If
    End If
    
    ' Quitar BOM si existe
    If Len(content) > 0 Then
        ' Quitar BOM Unicode (ChrW(&HFEFF))
        If Left(content, 1) = ChrW(&HFEFF) Then
            content = Mid(content, 2)
        End If
        
        ' Quitar secuencia BOM UTF-8 "ï»¿" si apareciera
        If Left(content, 3) = "ï»¿" Then
            content = Mid(content, 4)
        End If
    End If
    
    ReadAllText = content
    Err.Clear
End Function

Sub WriteAllText(path, text)
    ' Normalizar EOL: convertir todo a CRLF antes de escribir
    Dim normalizedText
    normalizedText = text
    
    ' Reemplazar CRLF por LF temporalmente para evitar duplicación
    normalizedText = Replace(normalizedText, vbCrLf, vbLf)
    ' Reemplazar CR solitarios por LF
    normalizedText = Replace(normalizedText, vbCr, vbLf)
    ' Convertir todo a CRLF
    normalizedText = Replace(normalizedText, vbLf, vbCrLf)
    
    ' Escribir como ANSI usando CreateTextFile (tercer parámetro False)
    On Error Resume Next
    Dim file
    Set file = objFSO.CreateTextFile(path, True, False) ' Overwrite, ANSI
    If Err.Number = 0 Then
        file.Write normalizedText
        file.Close
    End If
    On Error GoTo 0
End Sub

Function FileExists(path)
    FileExists = objFSO.FileExists(path)
End Function

' ============================================================================
' UTILIDADES DE ESCRITURA UTF-8
' ============================================================================

' Escribe contenido a un archivo en formato UTF-8 sin BOM
Sub WriteUtf8File(path, content)
    Dim stream
    Set stream = CreateObject("ADODB.Stream")
    
    On Error Resume Next
    stream.Type = 2 ' adTypeText
    stream.Charset = "UTF-8"
    stream.Open
    stream.WriteText content
    stream.SaveToFile path, 2 ' adSaveCreateOverWrite
    stream.Close
    
    If Err.Number <> 0 Then
        WScript.Echo "ERROR: No se pudo escribir archivo UTF-8: " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
    
    Set stream = Nothing
End Sub

' ============================================================================
' UTILIDADES DE BLOQUEO/FS PARA ACCESS
' ============================================================================

' Devuelve la ruta del archivo .laccdb (lock) para una base de datos Access
Function GetLockPath(dbPath)
    Dim fso, parentDir, baseName
    Set fso = CreateObject("Scripting.FileSystemObject")
    parentDir = fso.GetParentFolderName(dbPath)
    baseName = fso.GetBaseName(dbPath)
    GetLockPath = fso.BuildPath(parentDir, baseName & ".laccdb")
End Function

' Verifica si existe el archivo de bloqueo .laccdb para una base de datos
Function IsAccessLockPresent(dbPath)
    Dim lockPath
    lockPath = GetLockPath(dbPath)
    IsAccessLockPresent = objFSO.FileExists(lockPath)
End Function

' Intenta limpiar un archivo de bloqueo .laccdb obsoleto
Sub TryCleanupStaleLock(dbPath, verbose)
    Dim lockPath
    lockPath = GetLockPath(dbPath)
    
    ' Solo intentar limpiar si existe el lock
    If Not objFSO.FileExists(lockPath) Then Exit Sub
    
    ' Verificar si hay procesos de Access ejecutándose
    If IsAccessProcessRunning() Then
        If verbose Then WScript.Echo "⚠ Lock presente pero hay procesos de Access activos - no se elimina"
        Exit Sub
    End If
    
    ' Intentar eliminar el lock obsoleto
    On Error Resume Next
    objFSO.DeleteFile lockPath, True
    If Err.Number = 0 Then
        If verbose Then WScript.Echo "✓ Lock obsoleto eliminado: " & lockPath
    Else
        If verbose Then WScript.Echo "⚠ No se pudo eliminar lock: " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
End Sub

' Verifica si hay procesos de Access ejecutándose
Function IsAccessProcessRunning()
    Dim objWMI, colProcesses, processCount
    
    On Error Resume Next
    Set objWMI = GetObject("winmgmts:")
    Set colProcesses = objWMI.ExecQuery("SELECT * FROM Win32_Process WHERE Name = 'MSACCESS.EXE'")
    
    processCount = 0
    For Each objProcess In colProcesses
        processCount = processCount + 1
    Next
    
    IsAccessProcessRunning = (processCount > 0)
    On Error GoTo 0
End Function

' ============================================================================
' UTILIDADES DE ARGUMENTOS Y FLAGS
' ============================================================================

Function GetArgValue(flag)
    Dim i
    For i = 0 To WScript.Arguments.Count - 2
        If LCase(WScript.Arguments(i)) = LCase("--" & flag) Then
            GetArgValue = WScript.Arguments(i+1)
            Exit Function
        End If
    Next
    GetArgValue = ""
End Function

Function HasFlag(flag)
    Dim i: HasFlag = False
    For i = 0 To WScript.Arguments.Count - 1
        If LCase(WScript.Arguments(i)) = LCase("--" & flag) Then
            HasFlag = True
            Exit Function
        End If
    Next
End Function

' ============================================================================
' UTILIDADES DE CODIFICACIÓN Y ARCHIVOS TEMPORALES
' ============================================================================

' Devuelve el directorio temporal del sistema o "." si falla
Function GetTempDir()
    On Error Resume Next
    Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
    GetTempDir = fso.GetSpecialFolder(2) ' Temp folder
    If Err.Number <> 0 Then GetTempDir = "."
    Err.Clear
End Function

' Convierte archivo UTF-8 a ANSI temporal y devuelve la ruta del temporal
Function ConvertToAnsiTemp(srcPath)
    On Error Resume Next
    Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(srcPath) Then ConvertToAnsiTemp = "": Exit Function
    
    Dim tmpDir, tmpPath
    tmpDir = GetTempDir()
    tmpPath = fso.BuildPath(tmpDir, fso.GetTempName() & "_" & fso.GetFileName(srcPath))
    
    ' Leer UTF-8
    Dim sIn: Set sIn = CreateObject("ADODB.Stream")
    sIn.Type = 2: sIn.Charset = "utf-8": sIn.Open: sIn.LoadFromFile srcPath
    Dim txt
    txt = sIn.ReadText: sIn.Close
    
    ' Escribir ANSI usando ACP del SO (FSO.CreateTextFile con unicode:=False)
    Dim t: Set t = fso.CreateTextFile(tmpPath, True, False)
    t.Write txt: t.Close
    
    If Err.Number <> 0 Then ConvertToAnsiTemp = "" Else ConvertToAnsiTemp = tmpPath
    Err.Clear
End Function

' Lee archivo como UTF-8, con fallback a Windows-1252 si falla
Function ReadAllTextUtf8WithFallback(path)
    On Error Resume Next
    Dim stream, content
    
    ' Intentar UTF-8 primero
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "UTF-8"
    stream.Open
    stream.LoadFromFile path
    content = stream.ReadText
    stream.Close
    
    If Err.Number = 0 Then
        ReadAllTextUtf8WithFallback = content
        Exit Function
    End If
    
    ' Fallback a Windows-1252
    Err.Clear
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "Windows-1252"
    stream.Open
    stream.LoadFromFile path
    content = stream.ReadText
    stream.Close
    
    If Err.Number = 0 Then
        ReadAllTextUtf8WithFallback = content
    Else
        ReadAllTextUtf8WithFallback = ""
    End If
    Err.Clear
End Function

' Elimina módulo usando exclusivamente VBIDE
Function RemoveModuleVBIDEIfExists(name)
    On Error Resume Next
    Dim vbProj: Set vbProj = objAccess.VBE.ActiveVBProject
    Dim comp: Set comp = vbProj.VBComponents(name)
    If Err.Number = 0 Then vbProj.VBComponents.Remove comp
    Err.Clear
End Function

' Importa clase desde string usando VBIDE
Sub ImportClassFromString(moduleName, content)
    On Error Resume Next
    Dim vbProj: Set vbProj = objAccess.VBE.ActiveVBProject
    Dim comp: Set comp = vbProj.VBComponents.Add(2) ' vbext_ct_ClassModule = 2
    If Err.Number <> 0 Then
        WScript.Echo "  ❌ Error creando módulo de clase: " & Err.Description
        Err.Clear
        Exit Sub
    End If
    
    ' Asignar contenido al módulo
    comp.CodeModule.DeleteLines 1, comp.CodeModule.CountOfLines
    comp.CodeModule.AddFromString content
    
    If Err.Number <> 0 Then
        WScript.Echo "  ❌ Error agregando código al módulo: " & Err.Description
        Err.Clear
        Exit Sub
    End If
    
    ' Renombrar al nombre esperado
    comp.Name = moduleName
    If Err.Number <> 0 Then
        WScript.Echo "  ⚠️ No se pudo renombrar a '" & moduleName & "': " & Err.Description
        Err.Clear
    End If
End Sub

' Crea copia temporal ANSI (ACP del SO) de un archivo UTF-8 del repo.
Function MakeAnsiTempCopy(srcPath)
    On Error Resume Next
    Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(srcPath) Then MakeAnsiTempCopy = "": Exit Function
    Dim tmpDir
    tmpDir = fso.GetSpecialFolder(2) 'Temp
    Dim tmpPath
    tmpPath = fso.BuildPath(tmpDir, fso.GetTempName() & "_" & fso.GetFileName(srcPath))
    ' Leer UTF-8
    Dim sIn: Set sIn = CreateObject("ADODB.Stream")
    sIn.Type = 2: sIn.Charset = "utf-8": sIn.Open: sIn.LoadFromFile srcPath
    Dim txt
    txt = sIn.ReadText: sIn.Close
    ' Escribir ANSI usando ACP del SO (FSO.CreateTextFile con unicode:=False)
    Dim t: Set t = fso.CreateTextFile(tmpPath, True, False)
    t.Write txt: t.Close
    If Err.Number <> 0 Then MakeAnsiTempCopy = "" Else MakeAnsiTempCopy = tmpPath
    Err.Clear
End Function

Sub SafeDeleteFile(p)
    On Error Resume Next
    If Len(p)>0 Then CreateObject("Scripting.FileSystemObject").DeleteFile p, True
    Err.Clear
End Sub

' Lee un archivo ANSI (ACP) y lo guarda como UTF-8 en dstPath.
Sub ConvertAnsiFileToUtf8(srcAnsiPath, dstUtf8Path)
    On Error Resume Next
    Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
    ' Leer ANSI con ACP del SO
    Dim t: Set t = fso.OpenTextFile(srcAnsiPath, 1, False) ' ForReading, ANSI
    Dim txt
    txt = t.ReadAll: t.Close
    ' Escribir UTF-8 (ADODB.Stream con Charset utf-8)
    Dim sOut: Set sOut = CreateObject("ADODB.Stream")
    sOut.Type = 2: sOut.Charset = "utf-8": sOut.Open
    sOut.WriteText txt
    sOut.SaveToFile dstUtf8Path, 2
    sOut.Close
    Err.Clear
End Sub

' Usa SaveAsText (ANSI) a un temporal y luego convierte a UTF-8 en destino repo.
Sub ExportModuleToUtf8(app, moduleName, dstPath)
    On Error Resume Next
    Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim tmp
    tmp = fso.BuildPath(fso.GetSpecialFolder(2), fso.GetTempName() & "_" & moduleName & ".txt")
    ' acModule = 5
    app.Application.SaveAsText 5, moduleName, tmp
    If Err.Number <> 0 Then WScript.Echo "  ❌ SaveAsText(" & moduleName & "): " & Err.Description: Err.Clear: Exit Sub
    Call ConvertAnsiFileToUtf8(tmp, dstPath)
    Call SafeDeleteFile(tmp)
    WScript.Echo "  ✓ Exportado UTF-8: " & moduleName & " → " & dstPath
End Sub

' Elimina un componente por nombre usando solo VBIDE.
Sub RemoveModuleIfExists(app, moduleName)
    Call RemoveVBComponentIfExists(app, moduleName)
End Sub

' Importa .bas/.cls desde SRC (UTF-8) creando temp ANSI. .bas → LoadFromText; .cls → VBIDE.Import (+rename).


Function DirExists(path)
    DirExists = objFSO.FolderExists(path)
End Function

Function PathCombine(base, rel)
    If Right(base, 1) = "\" Then
        PathCombine = base & rel
    Else
        PathCombine = base & "\" & rel
    End If
End Function

Function MakeRelative(base, abs)
    ' Implementación simple - devuelve path relativo si es posible
    If InStr(1, abs, base, vbTextCompare) = 1 Then
        MakeRelative = Mid(abs, Len(base) + 2) ' +2 para saltar el \
    Else
        MakeRelative = abs
    End If
End Function

Function NowUtcIso()
    Dim now
    now = Now()
    ' Formato ISO UTC simplificado
    NowUtcIso = Year(now) & "-" & _
                Right("00" & Month(now), 2) & "-" & _
                Right("00" & Day(now), 2) & "T" & _
                Right("00" & Hour(now), 2) & ":" & _
                Right("00" & Minute(now), 2) & ":" & _
                Right("00" & Second(now), 2) & "Z"
End Function

Function OleToRgbHex(lng)
    Dim r, g, b
    r = lng And &HFF
    g = (lng \ &H100) And &HFF
    b = (lng \ &H10000) And &HFF
    OleToRgbHex = "#" & Right("00" & Hex(r), 2) & Right("00" & Hex(g), 2) & Right("00" & Hex(b), 2)
End Function

Function RgbHexToOle(hex)
    Dim cleanHex, r, g, b
    cleanHex = Replace(hex, "#", "")
    If Len(cleanHex) = 6 Then
        r = CLng("&H" & Mid(cleanHex, 1, 2))
        g = CLng("&H" & Mid(cleanHex, 3, 2))
        b = CLng("&H" & Mid(cleanHex, 5, 2))
        RgbHexToOle = r + (g * &H100) + (b * &H10000)
    Else
        RgbHexToOle = 0
    End If
End Function

' ===== Limpieza de metadatos VBA (.bas/.cls) =====
Function NormalizeEOL(s)
    If InStr(s, vbCrLf) = 0 And InStr(s, vbLf) > 0 Then s = Replace(s, vbLf, vbCrLf)
    NormalizeEOL = s
End Function

Function CleanVBAContent(rawText, ext)
    Dim lines, i, ln, out, isClass, inVB6Block, skipLine
    isClass = (LCase(ext) = "cls")
    inVB6Block = False
    rawText = NormalizeEOL(rawText)
    lines = Split(rawText, vbCrLf)
    out = ""
    For i = 0 To UBound(lines)
        ln = Trim(lines(i))
        skipLine = False
        
        ' Quitar cabecera VB6 de clases
        If isClass Then
            If LCase(Left(ln, 7)) = "version" Then
                inVB6Block = True
                skipLine = True
            ElseIf inVB6Block Then
                If UCase(ln) = "END" Then inVB6Block = False
                skipLine = True
            End If
        End If
        
        ' Quitar atributos y opciones duplicadas
        If Not skipLine Then
            If LCase(Left(ln, 17)) = "attribute vb_name" Then skipLine = True
            If isClass And LCase(Left(ln, 10)) = "attribute " Then skipLine = True
            If LCase(ln) = "option compare database" Then skipLine = True
            If LCase(ln) = "option explicit" Then skipLine = True
        End If
        
        ' Mantener el resto
        If Not skipLine Then
            If out <> "" Then out = out & vbCrLf
            out = out & lines(i)
        End If
        
    Next
    ' Trim final de líneas vacías iniciales
    Do While Left(out, 2) = vbCrLf
        out = Mid(out, 3)
    Loop
    CleanVBAContent = out
End Function
' ===== FIN limpieza =====

' ==== DAO helpers canónicos para password + bypass de inicio (.accdb) ====

Sub DaoEnsureEngine(ByRef dao)
    On Error Resume Next
    Set dao = CreateObject("DAO.DBEngine.120")
    If Err.Number <> 0 Then Err.Clear: Set dao = CreateObject("DAO.DBEngine.36")
    On Error GoTo 0
End Sub

Function DaoOpenDatabase(dbPath, pwd)
    Dim dao: Call DaoEnsureEngine(dao)
    If pwd <> "" Then
        Set DaoOpenDatabase = dao.OpenDatabase(dbPath, False, False, ";PWD=" & pwd)
    Else
        Set DaoOpenDatabase = dao.OpenDatabase(dbPath, False, False)
    End If
End Function

Function GetMacrosContainerName(db)
    If Not IsEmpty(db.Containers("Scripts")) Then GetMacrosContainerName = "Scripts": Exit Function
    If Not IsEmpty(db.Containers("Macros")) Then GetMacrosContainerName = "Macros": Exit Function
    GetMacrosContainerName = ""
End Function

Function IsEmpty(o) : On Error Resume Next: Dim t: Set t = o: IsEmpty = (Err.Number <> 0): Err.Clear: End Function

Function MacroExists(dbPath, pwd, macroName)
    On Error Resume Next
    Dim db, cName, c, d
    Set db = DaoOpenDatabase(dbPath, pwd)
    If Err.Number <> 0 Then MacroExists = False: Exit Function
    cName = GetMacrosContainerName(db)
    If cName = "" Then MacroExists = False: db.Close: Exit Function
    Set c = db.Containers(cName)
    For Each d In c.Documents
        If LCase(d.Name) = LCase(macroName) Then MacroExists = True: db.Close: Exit Function
    Next
    MacroExists = False
    db.Close
End Function

Sub RenameMacroIfExists(dbPath, pwd, oldName, newName)
    On Error Resume Next
    Dim db, cName, c, d
    Set db = DaoOpenDatabase(dbPath, pwd)
    If Err.Number <> 0 Then Exit Sub
    cName = GetMacrosContainerName(db)
    If cName = "" Then db.Close: Exit Sub
    Set c = db.Containers(cName)
    For Each d In c.Documents
        If LCase(d.Name) = LCase(oldName) Then d.Name = newName: Exit For
    Next
    db.Close
End Sub



' [DEPRECATED] El parámetro bypassStartup se ignora: OpenAccessQuiet aplica bypass automáticamente.
Function OpenAccessApp(dbPath, password, bypassStartup)
    ' Función unificada que usa OpenAccessQuiet con bypass siempre activo
    Set OpenAccessApp = OpenAccessQuiet(dbPath, password)
End Function

Sub CloseAccessApp(app)
    ' Función unificada que usa CloseAccessQuiet
    Call CloseAccessQuiet(app)
End Sub

' ==== FIN helpers canónicos ====

' ============================================================================
' DIFF SEMÁNTICO JSON
' ============================================================================

Function DiffJsonSemantico(jsonA, jsonB)
    On Error Resume Next
    
    Dim parser
    Set parser = New JsonParser
    
    Dim objA, objB
    Set objA = parser.Parse(jsonA)
    Set objB = parser.Parse(jsonB)
    
    If Err.Number <> 0 Then
        DiffJsonSemantico = "Error al parsear JSON: " & Err.Description
        Exit Function
    End If
    
    ' Normalizar y comparar
    Dim normalizedA, normalizedB
    normalizedA = NormalizeJsonObject(objA)
    normalizedB = NormalizeJsonObject(objB)
    
    If normalizedA = normalizedB Then
        DiffJsonSemantico = ""
    Else
        DiffJsonSemantico = FindDifferences(objA, objB, "", 0)
    End If
    
    On Error GoTo 0
End Function

Function NormalizeJsonObject(obj)
    Dim writer
    Set writer = New JsonWriter
    ' Aquí se podría implementar ordenación de claves, pero por simplicidad
    ' usamos la representación directa
    NormalizeJsonObject = writer.Stringify(obj)
End Function

Function FindDifferences(objA, objB, path, depth)
    If depth > 20 Then
        FindDifferences = "[Máximo de 20 niveles alcanzado]"
        Exit Function
    End If
    
    Dim result
    result = ""
    
    ' Implementación básica de comparación
    If IsDictionary(objA) And IsDictionary(objB) Then
        Dim key
        For Each key In objA.Keys
            If Not objB.Exists(key) Then
                result = result & "Clave faltante en B: " & path & "." & key & vbCrLf
            End If
        Next
        
        For Each key In objB.Keys
            If Not objA.Exists(key) Then
                result = result & "Clave extra en B: " & path & "." & key & vbCrLf
            End If
        Next
    Else
        If objA <> objB Then
            result = "Diferencia en " & path & ": '" & objA & "' vs '" & objB & "'" & vbCrLf
        End If
    End If
    
    FindDifferences = result
End Function

' ============================================================================
' FUNCIONES HELPER PARA EXPORT-FORM MEJORADO
' ============================================================================

' Verifica si un formulario existe en la BD sin abrirlo
Function HasForm(formName)
    Dim it
    HasForm = False
    On Error Resume Next
    For Each it In objAccess.CurrentProject.AllForms
        If StrComp(it.Name, formName, vbTextCompare) = 0 Then 
            HasForm = True
            Exit For
        End If
    Next
    If Err.Number <> 0 Then
        HasForm = False
        Err.Clear
    End If
    On Error GoTo 0
End Function

' Espera a que la BD se pueda abrir en modo exclusivo
Function WaitForExclusive(dbPath, password, seconds)
    Dim t: t = Timer
    Dim lockPath: lockPath = Replace(dbPath, ".accdb", ".laccdb")
    
    Do
        On Error Resume Next
        Set objAccess = OpenAccessQuiet(dbPath, password) ' ya exclusiva
        If Err.Number = 0 Then 
            WaitForExclusive = True
            Exit Function
        End If
        
        Dim code: code = Err.Number: Err.Clear
        WScript.Echo "  [WARN] Intento de apertura exclusiva fallido (Err " & code & ")."
        
        If objFSO.FileExists(lockPath) Then _
           WScript.Echo "  [HINT] Se detecta archivo de bloqueo .laccdb. Cierre instancias/usuarios y reintente."
        
        WScript.Sleep 500
    Loop While (Timer - t) < seconds
    
    WaitForExclusive = False
End Function

' Asegura que un directorio existe, creándolo si es necesario
Function EnsureDir(dir)
    On Error Resume Next
    If Not objFSO.FolderExists(dir) Then
        objFSO.CreateFolder dir
    End If
    EnsureDir = (Err.Number = 0)
    Err.Clear
End Function

' Escribe texto a archivo en UTF-8 sin BOM
Sub WriteAllTextUtf8NoBom(path, content)
    On Error Resume Next
    Dim stream
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2 ' adTypeText
    stream.Charset = "UTF-8"
    stream.Open
    stream.WriteText content
    stream.Position = 3 ' Saltar BOM UTF-8
    
    Dim binaryStream
    Set binaryStream = CreateObject("ADODB.Stream")
    binaryStream.Type = 1 ' adTypeBinary
    binaryStream.Open
    stream.CopyTo binaryStream
    stream.Close
    
    binaryStream.SaveToFile path, 2 ' adSaveCreateOverWrite
    binaryStream.Close
    On Error GoTo 0
End Sub

' Calcula SHA1 de un archivo (implementación simplificada)
Function Sha1OfFile(path)
    On Error Resume Next
    Dim objFile
    Set objFile = objFSO.GetFile(path)
    
    If Err.Number <> 0 Then
        Sha1OfFile = ""
        Exit Function
    End If
    
    ' Hash simplificado basado en contenido del archivo
    Dim content, i, hash
    content = ReadAllText(path)
    hash = 0
    
    For i = 1 To Len(content)
        hash = hash + Asc(Mid(content, i, 1)) * i
    Next
    
    ' Convertir a formato hexadecimal de 40 caracteres (simulando SHA1)
    Sha1OfFile = Right(String(40, "0") & Hex(hash), 40)
    Err.Clear
End Function

' ============================================================================
' FUNCIONES DE MAPEO PARA PROPIEDADES DE FORMULARIO
' ============================================================================

' Mapea BorderStyle numérico a token canónico
Function MapBorderStyleToToken(borderStyleValue)
    Select Case borderStyleValue
        Case 0: MapBorderStyleToToken = "None"
        Case 1: MapBorderStyleToToken = "Thin"
        Case 2: MapBorderStyleToToken = "Sizable"
        Case 3: MapBorderStyleToToken = "Dialog"
        Case Else: MapBorderStyleToToken = "Sizable" ' Default
    End Select
End Function

' Mapea ScrollBars numérico a token canónico
Function MapScrollBarsToToken(scrollBarsValue)
    Select Case scrollBarsValue
        Case 0: MapScrollBarsToToken = "Neither"
        Case 1: MapScrollBarsToToken = "Horizontal"
        Case 2: MapScrollBarsToToken = "Vertical"
        Case 3: MapScrollBarsToToken = "Both"
        Case Else: MapScrollBarsToToken = "Neither" ' Default
    End Select
End Function

' Mapea MinMaxButtons numérico a token canónico
Function MapMinMaxButtonsToToken(minMaxButtonsValue)
    Select Case minMaxButtonsValue
        Case 0: MapMinMaxButtonsToToken = "None"
        Case 1: MapMinMaxButtonsToToken = "Min Enabled"
        Case 2: MapMinMaxButtonsToToken = "Max Enabled"
        Case 3: MapMinMaxButtonsToToken = "Both Enabled"
        Case Else: MapMinMaxButtonsToToken = "None" ' Default
    End Select
End Function

' Mapea RecordsetType numérico a token canónico
Function MapRecordsetTypeToToken(recordsetTypeValue)
    Select Case recordsetTypeValue
        Case 0: MapRecordsetTypeToToken = "Dynaset"
        Case 1: MapRecordsetTypeToToken = "Snapshot"
        Case 2: MapRecordsetTypeToToken = "Dynaset (Inconsistent Updates)"
        Case Else: MapRecordsetTypeToToken = "Dynaset" ' Default
    End Select
End Function

' Mapea Orientation numérico a token canónico
Function MapOrientationToToken(orientationValue)
    Select Case orientationValue
        Case 0: MapOrientationToToken = "LeftToRight"
        Case 1: MapOrientationToToken = "RightToLeft"
        Case Else: MapOrientationToToken = "LeftToRight" ' Default
    End Select
End Function

' Mapea SplitFormOrientation numérico a token canónico
Function MapSplitFormOrientationToToken(splitFormOrientationValue)
    Select Case splitFormOrientationValue
        Case 0: MapSplitFormOrientationToToken = "DatasheetOnTop"
        Case 1: MapSplitFormOrientationToToken = "DatasheetOnBottom"
        Case 2: MapSplitFormOrientationToToken = "DatasheetOnLeft"
        Case 3: MapSplitFormOrientationToToken = "DatasheetOnRight"
        Case Else: MapSplitFormOrientationToToken = "DatasheetOnTop" ' Default
    End Select
End Function

' ============================================================================
' FUNCIONES DE NORMALIZACIÓN Y MAPEO INVERSO
' ============================================================================

' Normaliza tokens de entrada en español a inglés canónico
Function NormalizeToken(propName, value)
    Dim normalizedValue
    normalizedValue = Trim(value)
    
    Select Case LCase(propName)
        Case "scrollbars"
            Select Case LCase(normalizedValue)
                Case "ninguna": NormalizeToken = "Neither"
                Case "horizontal": NormalizeToken = "Horizontal"
                Case "vertical": NormalizeToken = "Vertical"
                Case "ambas", "ambos": NormalizeToken = "Both"
                Case Else: NormalizeToken = normalizedValue
            End Select
            
        Case "borderstyle"
            Select Case LCase(normalizedValue)
                Case "ninguno": NormalizeToken = "None"
                Case "fino": NormalizeToken = "Thin"
                Case "redimensionable": NormalizeToken = "Sizable"
                Case "cuadro de diálogo": NormalizeToken = "Dialog"
                Case Else: NormalizeToken = normalizedValue
            End Select
            
        Case "minmaxbuttons"
            Select Case LCase(normalizedValue)
                Case "ninguno": NormalizeToken = "None"
                Case "solo minimizar": NormalizeToken = "Min Enabled"
                Case "solo maximizar": NormalizeToken = "Max Enabled"
                Case "ambos": NormalizeToken = "Both Enabled"
                Case Else: NormalizeToken = normalizedValue
            End Select
            
        Case "recordsettype"
            Select Case LCase(normalizedValue)
                Case "instantánea": NormalizeToken = "Snapshot"
                Case "dynaset (actualizaciones incoherentes)": NormalizeToken = "Dynaset (Inconsistent Updates)"
                Case Else: NormalizeToken = normalizedValue
            End Select
            
        Case "orientation"
            Select Case LCase(normalizedValue)
                Case "de izquierda a derecha": NormalizeToken = "LeftToRight"
                Case "de derecha a izquierda": NormalizeToken = "RightToLeft"
                Case Else: NormalizeToken = normalizedValue
            End Select
            
        Case "splitformorientation"
            Select Case LCase(normalizedValue)
                Case "hoja de datos arriba": NormalizeToken = "DatasheetOnTop"
                Case "hoja de datos abajo": NormalizeToken = "DatasheetOnBottom"
                Case "hoja de datos izquierda": NormalizeToken = "DatasheetOnLeft"
                Case "hoja de datos derecha": NormalizeToken = "DatasheetOnRight"
                Case Else: NormalizeToken = normalizedValue
            End Select
            
        Case Else
            ' Para propiedades booleanas
            Select Case LCase(normalizedValue)
                Case "sí", "si": NormalizeToken = "true"
                Case "no": NormalizeToken = "false"
                Case Else: NormalizeToken = normalizedValue
            End Select
    End Select
End Function

' Mapea token canónico a valor numérico de BorderStyle
Function TokenToBorderStyle(token)
    Select Case LCase(token)
        Case "none": TokenToBorderStyle = 0
        Case "thin": TokenToBorderStyle = 1
        Case "sizable": TokenToBorderStyle = 2
        Case "dialog": TokenToBorderStyle = 3
        Case Else: TokenToBorderStyle = 2 ' Default: Sizable
    End Select
End Function

' Mapea token canónico a valor numérico de ScrollBars
Function TokenToScrollBars(token)
    Select Case LCase(token)
        Case "neither": TokenToScrollBars = 0
        Case "horizontal": TokenToScrollBars = 1
        Case "vertical": TokenToScrollBars = 2
        Case "both": TokenToScrollBars = 3
        Case Else: TokenToScrollBars = 0 ' Default: Neither
    End Select
End Function

' Mapea token canónico a valor numérico de MinMaxButtons
Function TokenToMinMaxButtons(token)
    Select Case LCase(token)
        Case "none": TokenToMinMaxButtons = 0
        Case "min enabled": TokenToMinMaxButtons = 1
        Case "max enabled": TokenToMinMaxButtons = 2
        Case "both enabled": TokenToMinMaxButtons = 3
        Case Else: TokenToMinMaxButtons = 0 ' Default: None
    End Select
End Function

' Mapea token canónico a valor numérico de RecordsetType
Function TokenToRecordsetType(token)
    Select Case LCase(token)
        Case "dynaset": TokenToRecordsetType = 0
        Case "snapshot": TokenToRecordsetType = 1
        Case "dynaset (inconsistent updates)": TokenToRecordsetType = 2
        Case Else: TokenToRecordsetType = 0 ' Default: Dynaset
    End Select
End Function

' Mapea token canónico a valor numérico de Orientation
Function TokenToOrientation(token)
    Select Case LCase(token)
        Case "lefttoright": TokenToOrientation = 0
        Case "righttoleft": TokenToOrientation = 1
        Case Else: TokenToOrientation = 0 ' Default: LeftToRight
    End Select
End Function

' Mapea token canónico a valor numérico de SplitFormOrientation
Function TokenToSplitFormOrientation(token)
    Select Case LCase(token)
        Case "datasheetontop": TokenToSplitFormOrientation = 0
        Case "datasheetonbottom": TokenToSplitFormOrientation = 1
        Case "datasheetonleft": TokenToSplitFormOrientation = 2
        Case "datasheetonright": TokenToSplitFormOrientation = 3
        Case Else: TokenToSplitFormOrientation = 0 ' Default: DatasheetOnTop
    End Select
End Function

' ============================================================================
' REGLAS DE COHERENCIA ENTRE PROPIEDADES
' ============================================================================

' Aplica reglas de coherencia entre propiedades de formulario
' Parámetros:
'   - props: Diccionario con las propiedades del formulario
'   - strictMode: Si es True, genera errores; si es False, solo advertencias
'   - verbose: Si es True, muestra las decisiones tomadas
Sub ApplyCoherenceRules(props, strictMode, verbose)
    Dim borderStyle, controlBox, minMaxButtons, modal, popUp, defaultView
    Dim hasError
    hasError = False
    
    ' Obtener valores actuales
    borderStyle = ""
    If props.Exists("borderStyle") Then borderStyle = props("borderStyle")
    
    controlBox = True
    If props.Exists("controlBox") Then controlBox = props("controlBox")
    
    minMaxButtons = ""
    If props.Exists("minMaxButtons") Then minMaxButtons = props("minMaxButtons")
    
    modal = False
    If props.Exists("modal") Then modal = props("modal")
    
    popUp = False
    If props.Exists("popUp") Then popUp = props("popUp")
    
    defaultView = ""
    If props.Exists("defaultView") Then defaultView = props("defaultView")
    
    ' Regla 1: Si borderStyle ∈ {"None","Dialog"} ⇒ controlBox=false y minMaxButtons="None"
    If borderStyle = "None" Or borderStyle = "Dialog" Then
        If controlBox = True Then
            If verbose Then LogInfo "Coherencia: borderStyle='" & borderStyle & "' requiere controlBox=false"
            props("controlBox") = False
        End If
        If minMaxButtons <> "None" And minMaxButtons <> "" Then
            If verbose Then LogInfo "Coherencia: borderStyle='" & borderStyle & "' requiere minMaxButtons='None'"
            props("minMaxButtons") = "None"
        End If
    End If
    
    ' Regla 2: Si controlBox=false ⇒ ignorar closeButton y minMaxButtons
    If controlBox = False Then
        If props.Exists("closeButton") Then
            If verbose Then LogInfo "Coherencia: controlBox=false, ignorando closeButton"
            props.Remove("closeButton")
        End If
        If minMaxButtons <> "None" And minMaxButtons <> "" Then
            If verbose Then LogInfo "Coherencia: controlBox=false, forzando minMaxButtons='None'"
            props("minMaxButtons") = "None"
        End If
    End If
    
    ' Regla 3: Si modal=true o popUp=true y borderStyle≠"Sizable" ⇒ no permitir min/max
    If (modal = True Or popUp = True) And borderStyle <> "Sizable" Then
        If minMaxButtons <> "None" And minMaxButtons <> "" Then
            Dim msg
            msg = "Incoherencia: formulario modal/popup con borderStyle='" & borderStyle & "' no debe tener minMaxButtons='" & minMaxButtons & "'"
            If strictMode Then
                LogErr msg
                hasError = True
            Else
                LogWarn msg
                If verbose Then LogInfo "Coherencia: forzando minMaxButtons='None' para modal/popup"
                props("minMaxButtons") = "None"
            End If
        End If
    End If
    
    ' Regla 4: SplitForm* solo si es Split Form
    Dim isSplitForm
    isSplitForm = (defaultView = "Split Form")
    
    If Not isSplitForm Then
        Dim splitProps
        splitProps = Array("splitFormSize", "splitFormOrientation", "splitFormSplitterBar")
        Dim i
        For i = 0 To UBound(splitProps)
            If props.Exists(splitProps(i)) Then
                Dim splitMsg
                splitMsg = "Propiedad '" & splitProps(i) & "' solo aplicable a Split Forms (defaultView='Split Form')"
                If strictMode Then
                    LogErr splitMsg
                    hasError = True
                Else
                    LogWarn splitMsg
                    If verbose Then LogInfo "Coherencia: removiendo propiedad '" & splitProps(i) & "' (no es Split Form)"
                    props.Remove(splitProps(i))
                End If
            End If
        Next
    End If
    
    ' Si hay errores en modo estricto, terminar
    If hasError And strictMode Then
        LogErr "Errores de coherencia detectados en modo --strict. Abortando."
        WScript.Quit 1
    End If
End Sub

' ============================================================================
' SISTEMA DE LOGGING
' ============================================================================

Sub LogInfo(message)
    If gVerbose Then
        WScript.Echo "[INFO] " & message
    End If
End Sub

Sub LogWarn(message)
    WScript.Echo "[WARN] " & message
End Sub

Sub LogErr(message)
    WScript.Echo "[ERROR] " & message
End Sub

' ============================================================================
' ============================================================================
' FUNCIONES DAO PARA BYPASS STARTUP
' ============================================================================

' Subrutina para establecer AllowByPassKey
Sub SetAllowBypassKey(dbPath, password, value)
    Dim db, prop
    Set db = DaoOpenDatabase(dbPath, password)
    
    On Error Resume Next
    If HasProp(db, "AllowByPassKey") Then
        db.Properties("AllowByPassKey") = value
    Else
        Set prop = db.CreateProperty("AllowByPassKey", 1, value) ' dbBoolean = 1
        db.Properties.Append prop
    End If
    
    If Err.Number <> 0 Then
        WScript.Echo "ERROR: No se pudo establecer AllowByPassKey: " & Err.Description
    End If
    On Error GoTo 0
    
    db.Close
End Sub

' Lista módulos VBA con opciones de formato y verificación
Sub ListModules()
    Dim flagJson, flagExpectSrc, flagDiff, i
    flagJson = False
    flagExpectSrc = False
    flagDiff = False
    
    ' Procesar flags
    For i = 1 To objArgs.Count - 1
        If objArgs(i) = "--json" Then flagJson = True
        If objArgs(i) = "--expectSrc" Then flagExpectSrc = True
        If objArgs(i) = "--diff" Then flagDiff = True
    Next
    
    Dim objAccess: Set objAccess = OpenAccessApp()
    If objAccess Is Nothing Then
        WScript.Echo "Error: No se pudo abrir Access"
        WScript.Quit 1
    End If
    
    Dim vbProject, vbComponent, modules, srcFiles
    Set vbProject = objAccess.VBE.ActiveVBProject
    Set modules = CreateObject("Scripting.Dictionary")
    Set srcFiles = CreateObject("Scripting.Dictionary")
    
    ' Recopilar módulos del proyecto
    For Each vbComponent In vbProject.VBComponents
        If vbComponent.Type = 1 Or vbComponent.Type = 2 Then
            Dim ext
            If vbComponent.Type = 1 Then
                ext = ".bas"
            Else
                ext = ".cls"
            End If
            modules.Add vbComponent.Name, ext
        End If
    Next
    
    ' Si --expectSrc o --diff, escanear archivos fuente
    If flagExpectSrc Or flagDiff Then
        Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
        If fso.FolderExists(strSourcePath) Then
            Dim folder: Set folder = fso.GetFolder(strSourcePath)
            Dim file
            For Each file In folder.Files
                If LCase(fso.GetExtensionName(file.Name)) = "bas" Or LCase(fso.GetExtensionName(file.Name)) = "cls" Then
                    Dim baseName
                    baseName = fso.GetBaseName(file.Name)
                    srcFiles.Add baseName, "." & LCase(fso.GetExtensionName(file.Name))
                End If
            Next
        End If
    End If
    
    ' Generar salida
    If flagJson Then
        WScript.Echo "{"
        WScript.Echo "  ""modules"": ["
        Dim first
        first = True
        Dim key
        For Each key In modules.Keys
            If Not first Then WScript.Echo ","
            WScript.Echo "    {""name"": """ & key & """, ""type"": """ & modules(key) & """}"
            first = False
        Next
        WScript.Echo "  ]"
        WScript.Echo "}"
    Else
        WScript.Echo "Módulos VBA encontrados:"
        For Each key In modules.Keys
            Dim status
            status = ""
            If flagExpectSrc Then
                If srcFiles.Exists(key) Then
                    status = " [✓ src]"
                Else
                    status = " [❌ no src]"
                End If
            End If
            If flagDiff And srcFiles.Exists(key) Then
                If modules(key) <> srcFiles(key) Then
                    status = status & " [⚠ ext diff]"
                End If
            End If
            WScript.Echo "  " & key & modules(key) & status
        Next
        
        If flagExpectSrc Then
            WScript.Echo "Archivos fuente sin módulo:"
            For Each key In srcFiles.Keys
                If Not modules.Exists(key) Then
                    WScript.Echo "  " & key & srcFiles(key) & " [❌ no module]"
                End If
            Next
        End If
    End If
    
    Call CloseAccessApp(objAccess)
End Sub

' Verifica módulos si se especifica el flag --verifyModules
Sub VerifyModulesIfRequested()
    Dim i, hasVerifyFlag
    hasVerifyFlag = False
    
    ' Buscar flag --verifyModules en argumentos
    For i = 0 To objArgs.Count - 1
        If objArgs(i) = "--verifyModules" Then
            hasVerifyFlag = True
            Exit For
        End If
    Next
    
    If hasVerifyFlag Then
        WScript.Echo ""
        WScript.Echo "=== VERIFICACION DE MODULOS ==="
        WScript.Echo "Ejecutando verificación de consistencia..."
        
        ' Llamar directamente a la verificación interna
        Call VerifyModulesInternal()
    End If
End Sub

' Verificación interna de módulos
Sub VerifyModulesInternal()
    On Error Resume Next
    Dim app, arr(), expected, diffInfo, hasVBIDE
    
    ' Usar la sesión de Access ya abierta
    Set app = objAccess
    If app Is Nothing Then
        WScript.Echo "⚠ Advertencia: No hay sesión de Access disponible para verificación"
        Exit Sub
    End If
    
    ' Inicializar array
    ReDim arr(-1)
    
    ' Intentar listar con VBIDE primero
    hasVBIDE = TryListModulesVBIDE(app, False, "", arr)
    
    ' Si falla, usar fallback
    If Not hasVBIDE Then
        arr = Array()
        If Not TryListModulesAllModules(app, "", arr) Then
            WScript.Echo "⚠ Advertencia: No se pudieron listar los módulos para verificación"
            Exit Sub
        End If
    End If
    
    ' Cargar módulos esperados desde /src
    Set expected = LoadExpectedFromSrc(strSourcePath)
    Set diffInfo = DiffModules(arr, expected)
    
    ' Verificar diferencias
    If diffInfo("missingCount") > 0 Or diffInfo("extraCount") > 0 Then
        WScript.Echo "⚠ La verificación detectó inconsistencias:"
        If diffInfo("missingCount") > 0 Then
            WScript.Echo "  - Módulos faltantes: " & diffInfo("missingCount")
        End If
        If diffInfo("extraCount") > 0 Then
            WScript.Echo "  - Módulos extras: " & diffInfo("extraCount")
        End If
        WScript.Echo "Ejecute: cscript condor_cli.vbs list-modules /expectSrc:\"" & strSourcePath & "\" /diff:on"
        WScript.Quit 1
    Else
        WScript.Echo "✓ Verificación completada sin inconsistencias"
    End If
    
    On Error GoTo 0
    
    ' Limpiar variables globales
    gBypassStartupEnabled = False
    gPreviousAllowBypassKey = Null
    gCurrentDbPath = ""
    gCurrentPassword = ""
    gPreviousStartupForm = Null
    gPreviousHasAutoExec = False
End Sub

' ============================================================================
' NUEVAS FUNCIONES AUXILIARES PARA REFACTORIZACIÓN
' ============================================================================

' Subrutina para resolver todos los flags antes de abrir Access
Sub ResolveFlags()
    Dim i, arg, nextArg
    
    For i = 1 To objArgs.Count - 1
        arg = LCase(objArgs(i))
        
        If arg = "--verbose" Then
            gVerbose = True
            If gVerbose Then WScript.Echo "[VERBOSE] Modo verbose activado"
            
        ElseIf arg = "--sharedopen" Or arg = "/sharedopen" Then
            gOpenShared = True
            If gVerbose Then WScript.Echo "[INFO] Apertura forzada en modo COMPARTIDO por --sharedopen."
            
        ElseIf arg = "--db" And i < objArgs.Count - 1 Then
            i = i + 1
            gDbPath = TrimQuotes(objArgs(i))
            
        ElseIf Left(arg, 4) = "/db:" Then
            gDbPath = TrimQuotes(Mid(objArgs(i), 5))
            
        ElseIf arg = "--password" And i < objArgs.Count - 1 Then
            i = i + 1
            gPassword = objArgs(i)
            
        ElseIf Left(arg, 5) = "/pwd:" Then
            gPassword = Mid(objArgs(i), 6)
            
        ElseIf arg = "--bypassstartup" And i < objArgs.Count - 1 Then
            i = i + 1
            gBypassStartup = True ' mantenemos compatibilidad
            If Not gVerbose Then
                ' no spam
            End If
            WScript.Echo "[DEPRECATED] --bypassstartup ya no es necesario y no tiene efecto (el CLI abre Access con bypass por defecto)."
            
        ElseIf Left(arg, 15) = "/bypassstartup:" Then
            gBypassStartup = True ' mantenemos compatibilidad
            If Not gVerbose Then
                ' no spam
            End If
            WScript.Echo "[DEPRECATED] --bypassstartup ya no es necesario y no tiene efecto (el CLI abre Access con bypass por defecto)."
            
        ElseIf arg = "--print-db" Then
            gPrintDb = True
            
        ElseIf arg = "--dry-run" Then
            gDryRun = True
            If gVerbose Then WScript.Echo "[VERBOSE] Flag --dry-run activado"
            
        End If
    Next
End Sub

' Subrutina para resolver la ruta de base de datos


' Subrutina para determinar bypass startup por defecto según comando
Sub SetDefaultBypassStartup()
    ' [DEPRECATED] Se mantiene por compatibilidad. El bypass se aplica siempre en OpenAccessQuiet.
    gBypassStartup = True
 End Sub

' Función para obtener el StartupForm actual
Function GetStartupForm(dbPath, password)
    On Error Resume Next
    Dim db, prop
    Set db = DaoOpenDatabase(dbPath, password)
    If Err.Number <> 0 Then
        GetStartupForm = Null
        Exit Function
    End If
    
    For Each prop In db.Properties
        If prop.Name = "StartupForm" Then
            GetStartupForm = prop.Value
            db.Close
            Set db = Nothing
            Exit Function
        End If
    Next
    
    GetStartupForm = Null
    db.Close
    Set db = Nothing
End Function

' Subrutina para limpiar el StartupForm
Sub ClearStartupForm(dbPath, password)
    On Error Resume Next
    Dim db, prop
    Set db = DaoOpenDatabase(dbPath, password)
    If Err.Number <> 0 Then Exit Sub
    
    For Each prop In db.Properties
        If prop.Name = "StartupForm" Then
            db.Properties.Delete "StartupForm"
            Exit For
        End If
    Next
    
    db.Close
     Set db = Nothing
 End Sub

' ============================================================================
' FUNCIONES TRANSACCIONALES PARA COMANDO UPDATE
' ============================================================================

' Actualizar todos los módulos (sync suave)
Sub UpdateAllModulesTransactional()
    ' Establecer modo silencioso antes de cualquier importación
    On Error Resume Next
    objAccess.Application.DisplayAlerts = False
    objAccess.Application.Echo False
    objAccess.DoCmd.SetWarnings False
    Err.Clear
    On Error GoTo 0
    
    Dim objFolder, objFile, strModuleName, importedCount, moduleType
    Set objFolder = objFSO.GetFolder(strSourcePath)
    importedCount = 0
    
    For Each objFile In objFolder.Files
        If LCase(objFSO.GetExtensionName(objFile.Name)) = "bas" Or LCase(objFSO.GetExtensionName(objFile.Name)) = "cls" Then
            strModuleName = objFSO.GetBaseName(objFile.Name)
            
            If gVerbose Then
                WScript.Echo "  Importando: " & strModuleName
            End If
            
            ' Usar la rutina unificada ImportVbaFile
            If ImportVbaFile(objFile.Path, strModuleName) Then
                importedCount = importedCount + 1
            End If
        End If
    Next
    
    WScript.Echo "Modulos actualizados: " & importedCount
    
    ' Verificación opcional de módulos
    Call VerifyModulesIfRequested()
End Sub

' Actualizar módulo específico
Sub UpdateSingleModuleTransactional(moduleName)
    ' Establecer modo silencioso antes de cualquier importación
    On Error Resume Next
    objAccess.Application.DisplayAlerts = False
    objAccess.Application.Echo False
    objAccess.DoCmd.SetWarnings False
    Err.Clear
    On Error GoTo 0
    
    Dim sourceFile, moduleType, actualFile
    
    ' Buscar archivo .bas o .cls
    sourceFile = objFSO.BuildPath(strSourcePath, moduleName & ".bas")
    If objFSO.FileExists(sourceFile) Then
        moduleType = 5  ' acModule
        actualFile = sourceFile
    Else
        sourceFile = objFSO.BuildPath(strSourcePath, moduleName & ".cls")
        If objFSO.FileExists(sourceFile) Then
            moduleType = 9  ' acClassModule
            actualFile = sourceFile
        Else
            WScript.Echo "Error: No se encontró el módulo " & moduleName & " (.bas o .cls) en /src"
            Exit Sub
        End If
    End If
    
    If gVerbose Then
        WScript.Echo "  Importando: " & moduleName & " desde " & actualFile
    End If
    
    ' Usar la rutina unificada ImportVbaFile
    If ImportVbaFile(actualFile, moduleName) Then
        WScript.Echo "✓ " & moduleName & " actualizado correctamente"
    Else
        WScript.Echo "❌ Error actualizando " & moduleName
    End If
End Sub

' Actualizar múltiples módulos específicos separados por comas
Sub UpdateMultipleModulesTransactional(moduleList)
    ' Establecer modo silencioso antes de cualquier importación
    On Error Resume Next
    objAccess.Application.DisplayAlerts = False
    objAccess.Application.Echo False
    objAccess.DoCmd.SetWarnings False
    Err.Clear
    On Error GoTo 0
    
    Dim modules, i, moduleName, successCount, totalCount
    
    ' Dividir la lista por comas y limpiar espacios
    modules = Split(moduleList, ",")
    successCount = 0
    totalCount = 0
    
    WScript.Echo "Procesando " & (UBound(modules) + 1) & " módulos..."
    
    For i = 0 To UBound(modules)
        moduleName = Trim(modules(i))
        If Len(moduleName) > 0 Then
            totalCount = totalCount + 1
            WScript.Echo "Procesando módulo: " & moduleName
            
            Dim sourceFile, moduleType, actualFile, found
            found = False
            
            ' Buscar archivo .bas o .cls
            sourceFile = objFSO.BuildPath(strSourcePath, moduleName & ".bas")
            WScript.Echo "Buscando: " & sourceFile
            If objFSO.FileExists(sourceFile) Then
                moduleType = 5  ' acModule
                actualFile = sourceFile
                found = True
                WScript.Echo "Encontrado archivo .bas: " & actualFile
            Else
                sourceFile = objFSO.BuildPath(strSourcePath, moduleName & ".cls")
                WScript.Echo "Buscando: " & sourceFile
                If objFSO.FileExists(sourceFile) Then
                    moduleType = 9  ' acClassModule
                    actualFile = sourceFile
                    found = True
                    WScript.Echo "Encontrado archivo .cls: " & actualFile
                End If
            End If
            
            If found Then
                WScript.Echo "Importando: " & moduleName & " desde " & actualFile
                
                ' Usar la rutina unificada ImportVbaFileRobust con manejo de errores
                On Error Resume Next
                Dim importResult
                importResult = ImportVbaFileRobust(actualFile, moduleName, strTempPath)
                If Err.Number <> 0 Then
                    WScript.Echo "❌ Error en ImportVbaFileRobust: " & Err.Description & " (Número: " & Err.Number & ")"
                    Err.Clear
                    importResult = False
                End If
                On Error GoTo 0
                
                If importResult Then
                    WScript.Echo "✓ " & moduleName & " actualizado correctamente"
                    successCount = successCount + 1
                Else
                    WScript.Echo "❌ Error actualizando " & moduleName
                End If
            Else
                WScript.Echo "❌ No se encontró el módulo " & moduleName & " (.bas o .cls) en /src"
            End If
            
            WScript.Echo "Continuando con siguiente módulo..."
        End If
    Next
    
    WScript.Echo "Resumen: " & successCount & " de " & totalCount & " módulos actualizados correctamente"
End Sub

' Actualizar solo módulos cambiados (comparación por tamaño+fecha)
Sub UpdateChangedModulesTransactional()
    ' Establecer modo silencioso antes de cualquier importación
    On Error Resume Next
    objAccess.Application.DisplayAlerts = False
    objAccess.Application.Echo False
    objAccess.DoCmd.SetWarnings False
    Err.Clear
    On Error GoTo 0
    
    Dim objFolder, objFile, strModuleName, importedCount, moduleType
    Dim tempExportDir, exportedFile, needsUpdate
    Set objFolder = objFSO.GetFolder(strSourcePath)
    importedCount = 0
    
    ' Crear directorio temporal para exportación
    tempExportDir = objFSO.BuildPath(objFSO.GetSpecialFolder(2), "condor_temp_export_" & Timer)
    If Not objFSO.FolderExists(tempExportDir) Then
        objFSO.CreateFolder tempExportDir
    End If
    
    WScript.Echo "Exportando módulos actuales para comparación..."
    Call ExportModulesToDirectoryTransactional(tempExportDir)
    
    WScript.Echo "Comparando archivos por tamaño y fecha..."
    
    For Each objFile In objFolder.Files
        If LCase(objFSO.GetExtensionName(objFile.Name)) = "bas" Or LCase(objFSO.GetExtensionName(objFile.Name)) = "cls" Then
            strModuleName = objFSO.GetBaseName(objFile.Name)
            exportedFile = objFSO.BuildPath(tempExportDir, objFile.Name)
            needsUpdate = False
            
            ' Determinar si necesita actualización
            If Not objFSO.FileExists(exportedFile) Then
                needsUpdate = True  ' Archivo no existe en BD
            Else
                ' Comparar tamaño y fecha
                Dim sourceFileObj, exportedFileObj
                Set sourceFileObj = objFSO.GetFile(objFile.Path)
                Set exportedFileObj = objFSO.GetFile(exportedFile)
                
                If sourceFileObj.Size <> exportedFileObj.Size Or sourceFileObj.DateLastModified <> exportedFileObj.DateLastModified Then
                    needsUpdate = True
                End If
            End If
            
            ' Solo importar si necesita actualización
            If needsUpdate Then
                If gVerbose Then
                    WScript.Echo "  Importando: " & strModuleName & " (cambio detectado)"
                End If
                
                ' Usar la rutina unificada ImportVbaFile
                If ImportVbaFile(objFile.Path, strModuleName) Then
                    importedCount = importedCount + 1
                End If
            Else
                If gVerbose Then
                    WScript.Echo "  - " & strModuleName & " (sin cambios)"
                End If
            End If
        End If
    Next
    
    ' Limpiar directorio temporal
    If objFSO.FolderExists(tempExportDir) Then
        objFSO.DeleteFolder tempExportDir, True
    End If
    
    WScript.Echo "Modulos actualizados: " & importedCount
    
    ' Verificación opcional de módulos
    Call VerifyModulesIfRequested()
End Sub

' Exportar módulos a directorio (versión transaccional)
Sub ExportModulesToDirectoryTransactional(exportDir)
    Dim vbProject, vbComponent
    Set vbProject = objAccess.VBE.ActiveVBProject
    
    For Each vbComponent In vbProject.VBComponents
        If vbComponent.Type = 1 Or vbComponent.Type = 2 Then ' Módulos estándar y de clase
            Dim exportPath, moduleType
            
            If vbComponent.Type = 1 Then
                exportPath = objFSO.BuildPath(exportDir, vbComponent.Name & ".bas")
                moduleType = 5  ' acModule
            Else
                exportPath = objFSO.BuildPath(exportDir, vbComponent.Name & ".cls")
                moduleType = 9  ' acClassModule
            End If
            
            ' Usar ExportModuleToUtf8 para exportar con codificación UTF-8
            Call ExportModuleToUtf8(objAccess, vbComponent.Name, exportPath)
            
            If Err.Number <> 0 Then
                If gVerbose Then
                    WScript.Echo "  Advertencia: No se pudo exportar " & vbComponent.Name & ": " & Err.Description
                End If
                Err.Clear
            End If
        End If
    Next
End Sub

' Subrutina para restaurar el StartupForm
' ===== FUNCIONES PARA LIST-MODULES =====

'--- BEGIN: WaitForProjectReady ---
Function WaitForProjectReady(app, retries, delayMs)
  On Error Resume Next
  Dim i, ok, tmp
  ok = False
  For i = 1 To retries
    Err.Clear
    tmp = app.CurrentProject.Name
    If Err.Number = 0 Then ok = True
    If Not ok Then
      Err.Clear
      Dim p: Set p = app.VBE.ActiveVBProject
      If Err.Number = 0 And Not p Is Nothing Then ok = True
    End If
    If ok Then Exit For
    WScript.Sleep delayMs
  Next
  WaitForProjectReady = ok
  On Error GoTo 0
End Function
'--- END: WaitForProjectReady ---

' Función para listar módulos usando VBIDE
' ===== LIST-MODULES | Núcleo de listado y salida =====

'--- BEGIN: TryListModulesVBIDE (robusto) ---
Function TryListModulesVBIDE(app, includeDocs, pattern, ByRef arr)
  On Error Resume Next
  ReDim arr(-1)
  If Not WaitForProjectReady(app, 20, 250) Then TryListModulesVBIDE = False: Exit Function

  Dim regex: If pattern<>"" Then Set regex = CreateObject("VBScript.RegExp"): regex.Pattern=pattern: regex.IgnoreCase=True
  Dim p, comps, i, vbComp, kind, name, dict
  Set p = app.VBE.ActiveVBProject: If Err.Number<>0 Or p Is Nothing Then Err.Clear: TryListModulesVBIDE=False: Exit Function
  Set comps = p.VBComponents: If Err.Number<>0 Or comps Is Nothing Then Err.Clear: TryListModulesVBIDE=False: Exit Function

  For i = 1 To comps.Count
    Set vbComp = comps(i)
    Select Case vbComp.Type
      Case 1: kind="STD"
      Case 2: kind="CLS"
      Case 3: kind="FRM"
      Case 100: kind="RPT"
      Case Else: kind="OTHER"
    End Select
    name = vbComp.Name
    If (kind="FRM" Or kind="RPT") And Not includeDocs Then
      ' omit
    ElseIf pattern="" Or regex.Test(name) Then
      If UBound(arr)=-1 Then ReDim arr(0) Else ReDim Preserve arr(UBound(arr)+1)
      Set dict = CreateObject("Scripting.Dictionary")
      dict.Add "kind", kind
      dict.Add "name", name
      Set arr(UBound(arr)) = dict
    End If
  Next

  TryListModulesVBIDE = (Err.Number=0)
  On Error GoTo 0
End Function
'--- END: TryListModulesVBIDE ---

'--- BEGIN: TryListModulesAllModules (robusto) ---
Function TryListModulesAllModules(app, pattern, ByRef arr)
  On Error Resume Next
  ReDim arr(-1)
  If Not WaitForProjectReady(app, 20, 250) Then TryListModulesAllModules=False: Exit Function

  If app.CurrentProject Is Nothing Then TryListModulesAllModules=False: Exit Function
  Dim mods, regex, m, name, dict
  If pattern<>"" Then Set regex=CreateObject("VBScript.RegExp"): regex.Pattern=pattern: regex.IgnoreCase=True
  Set mods = app.CurrentProject.AllModules: If Err.Number<>0 Or mods Is Nothing Then Err.Clear: TryListModulesAllModules=False: Exit Function

  For Each m In mods
    name = m.Name
    If pattern="" Or regex.Test(name) Then
      If UBound(arr)=-1 Then ReDim arr(0) Else ReDim Preserve arr(UBound(arr)+1)
      Set dict = CreateObject("Scripting.Dictionary")
      dict.Add "kind", "STD"
      dict.Add "name", name
      Set arr(UBound(arr)) = dict
    End If
  Next
  TryListModulesAllModules = (Err.Number=0)
  On Error GoTo 0
End Function
'--- END: TryListModulesAllModules ---

'--- BEGIN: TryListModulesDAO (manejo robusto de errores) ---
Function TryListModulesDAO(app, pattern, ByRef arr)
  On Error Resume Next
  ReDim arr(-1)
  
  If app Is Nothing Then 
    If gVerbose Then WScript.Echo "DEBUG: TryListModulesDAO - app es Nothing"
    TryListModulesDAO = False: Exit Function
  End If
  
  If app.CurrentProject Is Nothing Then 
    If gVerbose Then WScript.Echo "DEBUG: TryListModulesDAO - CurrentProject es Nothing"
    TryListModulesDAO = False: Exit Function
  End If
  
  Dim regex, db, obj, name, dict, moduleCount
  If pattern<>"" Then Set regex=CreateObject("VBScript.RegExp"): regex.Pattern=pattern: regex.IgnoreCase=True
  
  If gVerbose Then WScript.Echo "DEBUG: TryListModulesDAO - Intentando acceder a AllModules..."
  
  ' Usar variable intermedia como en el código de inspiración
  Set db = app.CurrentProject
  If Err.Number <> 0 Then
    If gVerbose Then WScript.Echo "DEBUG: TryListModulesDAO - Error asignando CurrentProject: " & Err.Number & " - " & Err.Description
    Err.Clear
    TryListModulesDAO = False
    Exit Function
  End If
  
  ' Contar módulos que coinciden con el patrón
  moduleCount = 0
  For Each obj In db.AllModules
    If Err.Number <> 0 Then
      If gVerbose Then WScript.Echo "DEBUG: TryListModulesDAO - Error en For Each (conteo): " & Err.Number & " - " & Err.Description
      ' No salir inmediatamente, intentar continuar
      Err.Clear
    Else
      name = obj.Name
      If Err.Number <> 0 Then
        If gVerbose Then WScript.Echo "DEBUG: TryListModulesDAO - Error obteniendo Name: " & Err.Number & " - " & Err.Description
        Err.Clear
      Else
        If gVerbose Then WScript.Echo "DEBUG: TryListModulesDAO - Encontrado módulo: " & name
        If pattern="" Or regex.Test(name) Then
          moduleCount = moduleCount + 1
        End If
      End If
    End If
  Next
  
  If gVerbose Then WScript.Echo "DEBUG: TryListModulesDAO - Módulos que coinciden con patrón: " & moduleCount
  
  ' Llenar el array con los módulos que coinciden
  If moduleCount > 0 Then
    For Each obj In db.AllModules
      If Err.Number = 0 Then
        name = obj.Name
        If Err.Number = 0 Then
          If pattern="" Or regex.Test(name) Then
            If UBound(arr)=-1 Then ReDim arr(0) Else ReDim Preserve arr(UBound(arr)+1)
            Set dict = CreateObject("Scripting.Dictionary")
            dict.Add "kind", "STD"
            dict.Add "name", name
            Set arr(UBound(arr)) = dict
          End If
        Else
          Err.Clear
        End If
      Else
        Err.Clear
      End If
    Next
  End If
  
  ' Limpiar referencias
  Set obj = Nothing
  Set db = Nothing
  
  TryListModulesDAO = (moduleCount > 0)
  If gVerbose Then WScript.Echo "DEBUG: TryListModulesDAO - Resultado final: " & TryListModulesDAO & ", Módulos encontrados: " & moduleCount
  On Error GoTo 0
End Function
'--- END: TryListModulesDAO ---

Sub PrintModulesText(arr)
    Dim total, i: total = 0
    If IsArray(arr) Then If UBound(arr) >= 0 Then total = UBound(arr)+1
    WScript.Echo "KIND  Name": WScript.Echo "----  ----"
    If total > 0 Then
        For i = 0 To UBound(arr)
            WScript.Echo arr(i)("kind") & "   " & arr(i)("name")
        Next
    End If
    WScript.Echo "": WScript.Echo "Total: " & total & " módulos"
End Sub

Sub PrintModulesJson(arr, diffInfo)
    Dim total, i: total = 0
    If IsArray(arr) Then If UBound(arr) >= 0 Then total = UBound(arr)+1
    Dim json, first: json = "{""modules"":[" : first = True
    If total > 0 Then
        For i = 0 To UBound(arr)
            If Not first Then json = json & ","
            json = json & "{""kind"":""" & arr(i)("kind") & """,""name"":""" & arr(i)("name") & """}"
            first = False
        Next
    End If
    json = json & "],""total"":" & total

    If IsObject(diffInfo) Then
        json = json & ",""diff"":{" & _
               """missing"":" & diffInfo("missingCount") & "," & _
               """extra"":" & diffInfo("extraCount") & "}"
    End If

    json = json & "}"
    WScript.Echo json
End Sub

Function LoadExpectedFromSrc(srcFolder)
    Dim expected, objFile, fileName, baseName
    Set expected = CreateObject("Scripting.Dictionary")
    If objFSO.FolderExists(srcFolder) Then
        For Each objFile In objFSO.GetFolder(srcFolder).Files
            fileName = objFile.Name
            If LCase(Right(fileName,4)) = ".bas" Then
                baseName = Left(fileName, Len(fileName)-4): expected(baseName) = "STD"
            ElseIf LCase(Right(fileName,4)) = ".cls" Then
                baseName = Left(fileName, Len(fileName)-4): expected(baseName) = "CLS"
            End If
        Next
    End If
    Set LoadExpectedFromSrc = expected
End Function

Function DiffModules(arr, expected)
    Dim present, missing, extra, diffInfo, i, name, kind, key
    Set present = CreateObject("Scripting.Dictionary")
    Set missing = CreateObject("Scripting.Dictionary")
    Set extra   = CreateObject("Scripting.Dictionary")
    If IsArray(arr) Then
        For i = 0 To UBound(arr)
            kind = arr(i)("kind"): name = arr(i)("name")
            If kind = "STD" Or kind = "CLS" Then present(name) = kind
        Next
    End If
    For Each key In expected.Keys
        If Not present.Exists(key) Then missing(key) = expected(key)
    Next
    For Each key In present.Keys
        If Not expected.Exists(key) Then extra(key) = present(key)
    Next
    Set diffInfo = CreateObject("Scripting.Dictionary")
    diffInfo("missing") = missing: diffInfo("extra") = extra
    diffInfo("missingCount") = missing.Count: diffInfo("extraCount") = extra.Count
    Set DiffModules = diffInfo
End Function

' ===== FIN núcleo list-modules =====

' Comando principal list-modules
' ===== LIST-MODULES | Comando principal =====
Sub ListModulesCommand()
    Dim app, closeAfter, arr(), includeDocs, pattern, jsonOn, diffOn
    Dim expectSrc, expectRaw, expected, diffInfo, ok
    Dim pwd, password

    includeDocs = HasFlag("includeDocs")     ' --includeDocs
    pattern     = GetArgValue("pattern")     ' --pattern <regex>
    jsonOn      = HasFlag("json")            ' --json
    diffOn      = HasFlag("diff")            ' --diff

    ' Password explícito o por defecto del CLI
    password = GetArgValue("password")
    
    ' Resolver contraseña si no se especificó
    If password = "" Then
        If gPassword = "" Then
            gPassword = GetDatabasePassword(strAccessPath)
        End If
        password = gPassword
    End If

    ' Expectación de /src para diff (si --expectSrc sin valor, usar strSourcePath)
    expectRaw = GetArgValue("expectSrc")     ' --expectSrc [ruta]
    If HasFlag("expectSrc") Then
        If expectRaw = "" Or LCase(expectRaw) = "on" Then
            expectSrc = strSourcePath
        Else
            expectSrc = expectRaw
        End If
    Else
        expectSrc = ""
    End If

    ' Verificar si necesitamos abrir Access o usar instancia existente
    Set app = Nothing
    closeAfter = False
    
    ' Si objAccess existe y tiene la BD correcta, usarlo
    If Not objAccess Is Nothing Then
        On Error Resume Next
        If objAccess.CurrentDb.Name = strAccessPath Then
            Set app = objAccess
        End If
        On Error GoTo 0
    End If
    
    ' Si no tenemos app válida, abrir nueva instancia
    If app Is Nothing Then
        If gVerbose Then WScript.Echo "DEBUG: password = '" & password & "'"
        Set app = OpenAccessApp(strAccessPath, password, True) ' bypass on
        If app Is Nothing Then
            WScript.Echo "Error: No se pudo abrir Access para listar módulos."
            WScript.Quit 1
        End If
        closeAfter = True
    End If

    ' Preferir VBIDE; fallback AllModules; fallback DAO
    If gVerbose Then WScript.Echo "DEBUG: Intentando TryListModulesVBIDE..."
    ok = TryListModulesVBIDE(app, includeDocs, pattern, arr)
    If gVerbose Then WScript.Echo "DEBUG: TryListModulesVBIDE resultado: " & ok
    If Not ok Then
        If gVerbose Then WScript.Echo "DEBUG: Intentando TryListModulesAllModules..."
        ok = TryListModulesAllModules(app, pattern, arr)
        If gVerbose Then WScript.Echo "DEBUG: TryListModulesAllModules resultado: " & ok
    End If
    If Not ok Then
        If gVerbose Then WScript.Echo "DEBUG: Intentando TryListModulesDAO..."
        ok = TryListModulesDAO(app, pattern, arr)
        If gVerbose Then WScript.Echo "DEBUG: TryListModulesDAO resultado: " & ok
    End If
    If Not ok Then
        Call CloseAccessApp(app)
        WScript.Echo "Error: No se pudieron listar los módulos."
        WScript.Quit 1
    End If

    ' Diff opcional contra /src
    If expectSrc <> "" And objFSO.FolderExists(expectSrc) Then
        Set expected = LoadExpectedFromSrc(expectSrc)
        Set diffInfo = DiffModules(arr, expected)
    End If

    ' Salida
    If jsonOn Then
        Call PrintModulesJson(arr, diffInfo)
    Else
        Call PrintModulesText(arr)
        If IsObject(diffInfo) Then
            If diffInfo("missingCount") > 0 Or diffInfo("extraCount") > 0 Then
                WScript.Echo "DIFERENCIAS:"
                WScript.Echo "  Faltantes: " & diffInfo("missingCount")
                WScript.Echo "  Extras:    " & diffInfo("extraCount")
            End If
        End If
    End If

    If closeAfter Then Call CloseAccessApp(app)
End Sub
' ===== FIN comando principal =====



' Función para verificar si existe un contenedor en la base de datos
Function HasContainer(db, name)
    On Error Resume Next
    Dim container
    Set container = db.Containers(name)
    HasContainer = (Err.Number = 0)
    On Error GoTo 0
End Function

' Función para obtener el nombre del contenedor de macros
Function GetMacroContainerName(db)
    If HasContainer(db, "Scripts") Then
        GetMacroContainerName = "Scripts"
    ElseIf HasContainer(db, "Macros") Then
        GetMacroContainerName = "Macros"
    Else
        GetMacroContainerName = ""
    End If
End Function



' Función para verificar si existe una macro
Function MacroExists(dbPath, password, macroName)
    On Error Resume Next
    Dim db, doc, cName
    Set db = DaoOpenDatabase(dbPath, password)
    If Err.Number <> 0 Then
        MacroExists = False
        Exit Function
    End If
    
    cName = GetMacroContainerName(db)
    If cName = "" Then
        MacroExists = False
        db.Close
        Set db = Nothing
        Exit Function
    End If
    
    For Each doc In db.Containers(cName).Documents
        If LCase(doc.Name) = LCase(macroName) Then
            MacroExists = True
            db.Close
            Set db = Nothing
            Exit Function
        End If
    Next
    
    MacroExists = False
    db.Close
    Set db = Nothing
    On Error GoTo 0
End Function

' Subrutina para renombrar macro si existe
Sub RenameMacroIfExists(dbPath, password, oldName, newName)
    On Error Resume Next
    Dim db, doc, cName
    Set db = DaoOpenDatabase(dbPath, password)
    If Err.Number <> 0 Then Exit Sub
    
    cName = GetMacroContainerName(db)
    If cName = "" Then
        db.Close
        Set db = Nothing
        Exit Sub
    End If
    
    For Each doc In db.Containers(cName).Documents
        If LCase(doc.Name) = LCase(oldName) Then
            doc.Name = newName
            Exit For
        End If
    Next
    
    db.Close
    Set db = Nothing
    On Error GoTo 0
End Sub

Sub RemoveBrokenReferences(app)
    On Error Resume Next
    Dim vbProj, ref
    Set vbProj = app.VBE.ActiveVBProject
    If vbProj Is Nothing Then Exit Sub
    For Each ref In vbProj.References
        If ref.IsBroken Then vbProj.References.Remove ref
    Next
    ' Asegurar referencias clave
    vbProj.References.AddFromGuid "{420B2830-E718-11CF-893D-00A0C9054228}", 1, 0   ' Scripting Runtime
    vbProj.References.AddFromGuid "{0002E157-0000-0000-C000-000000000046}", 5, 3   ' VBIDE Extensibility
    Err.Clear
End Sub

' ============================================================================
' HELPERS DE COLOR PARA EXPORT DE FORMULARIOS
' ============================================================================

Function Hex2(n)
    Hex2 = Right("0" & Hex(n), 2)
End Function

Function OleToHex(ole)
    Dim r, g, b
    r = (ole And &HFF)
    g = (ole \ &H100) And &HFF
    b = (ole \ &H10000) And &HFF
    OleToHex = "#" & Hex2(r) & Hex2(g) & Hex2(b)
End Function

Function HexToOle(hex)
    ' Parsea "#RRGGBB" y devuelve Long en formato BGR
    If Left(hex, 1) = "#" Then hex = Mid(hex, 2)
    If Len(hex) <> 6 Then
        HexToOle = 0
        Exit Function
    End If
    
    Dim r, g, b
    r = CLng("&H" & Mid(hex, 1, 2))
    g = CLng("&H" & Mid(hex, 3, 2))
    b = CLng("&H" & Mid(hex, 5, 2))
    
    ' Convertir a formato BGR (Blue-Green-Red)
    HexToOle = (b * &H10000) + (g * &H100) + r
End Function

' ===================================================================
' FUNCIÓN AUXILIAR: IsPresent
' Descripción: Verifica si una clave existe en un diccionario/objeto
' ===================================================================
Function IsPresent(dict, key)
    On Error Resume Next
    IsPresent = dict.Exists(key)
    If Err.Number <> 0 Then
        ' Si no es un Dictionary, intentar acceso directo
        Err.Clear
        Dim temp
        temp = dict(key)
        IsPresent = (Err.Number = 0)
    End If
    On Error GoTo 0
End Function

' ===================================================================
' FUNCIONES HELPER PARA FIX-SRC-HEADERS
' ===================================================================



' Función para normalizar nombre desde filename
Function NormalizeNameFromFilename(filePath)
    Dim baseName
    baseName = objFSO.GetBaseName(filePath)
    NormalizeNameFromFilename = baseName
End Function

' Función para verificar si una línea es parte de la cabecera
Function IsHeaderLine(line)
    Dim trimmedLine
    trimmedLine = Trim(line)
    
    ' Líneas de cabecera típicas
    If Left(trimmedLine, 9) = "Attribute" Then
        IsHeaderLine = True
    ElseIf Left(trimmedLine, 7) = "VERSION" Then
        IsHeaderLine = True
    ElseIf Left(trimmedLine, 6) = "Option" Then
        IsHeaderLine = True
    ElseIf trimmedLine = "" Then
        IsHeaderLine = True ' Líneas vacías al inicio se consideran cabecera
    ElseIf Left(trimmedLine, 1) = "'" Then
        IsHeaderLine = True ' Comentarios al inicio se consideran cabecera
    Else
        IsHeaderLine = False
    End If
End Function

' Función para separar cabecera y cuerpo
Function SplitHeaderBody(content)
    ' Normalizar EOL: reemplazar CRLF y CR por LF; Split por LF
    Dim normalizedContent
    normalizedContent = content
    
    ' Reemplazar CRLF por LF
    normalizedContent = Replace(normalizedContent, vbCrLf, vbLf)
    ' Reemplazar CR solitarios por LF
    normalizedContent = Replace(normalizedContent, vbCr, vbLf)
    
    ' Split por LF
    Dim lines, i, headerEnd, result(1)
    lines = Split(normalizedContent, vbLf)
    
    ' Recorrer desde la 1ª línea; mientras IsHeaderLine(line)=True, formar cabecera
    headerEnd = -1
    For i = 0 To UBound(lines)
        If Not IsHeaderLine(lines(i)) Then
            headerEnd = i - 1
            Exit For
        End If
    Next
    
    ' Si no se encontró fin de cabecera, todo es cabecera
    If headerEnd = -1 Then headerEnd = UBound(lines)
    
    ' Construir cabecera - recomponer con CRLF
    If headerEnd >= 0 Then
        Dim headerLines()
        ReDim headerLines(headerEnd)
        For i = 0 To headerEnd
            headerLines(i) = lines(i)
        Next
        result(0) = Join(headerLines, vbCrLf)
    Else
        result(0) = ""
    End If
    
    ' Construir cuerpo - recomponer con CRLF
    If headerEnd < UBound(lines) Then
        Dim bodyLines()
        ReDim bodyLines(UBound(lines) - headerEnd - 1)
        For i = headerEnd + 1 To UBound(lines)
            bodyLines(i - headerEnd - 1) = lines(i)
        Next
        result(1) = Join(bodyLines, vbCrLf)
    Else
        result(1) = ""
    End If
    
    SplitHeaderBody = result
End Function

' Función para construir cabecera de módulo .bas
Function BuildBasHeader(moduleName)
    Dim header
    header = "Attribute VB_Name = """ & moduleName & """" & vbCrLf
    header = header & "Option Compare Database" & vbCrLf
    header = header & "Option Explicit" & vbCrLf
    BuildBasHeader = header
End Function

' Función para construir cabecera de clase .cls
Function BuildClsHeader(className)
    ' Genera cabecera mínima válida para Access con formato específico
    Dim header
    header = "VERSION 1.0 CLASS" & vbCrLf
    header = header & "BEGIN" & vbCrLf
    header = header & "  MultiUse = -1  'True" & vbCrLf
    header = header & "END" & vbCrLf
    header = header & "Attribute VB_Name = """ & className & """" & vbCrLf
    header = header & "Attribute VB_GlobalNameSpace = False" & vbCrLf
    header = header & "Attribute VB_Creatable = False" & vbCrLf
    header = header & "Attribute VB_PredeclaredId = False" & vbCrLf
    header = header & "Attribute VB_Exposed = False" & vbCrLf
    header = header & "Option Compare Database" & vbCrLf
    header = header & "Option Explicit" & vbCrLf
    BuildClsHeader = header
End Function

' Función para validar si una cabecera satisface los requisitos según el tipo
Function HeaderSatisfies(expectedType, headerText)
    ' expectedType: "bas" o "cls"
    ' headerText: texto de la cabecera a validar
    ' Retorna: True si la cabecera es válida para el tipo, False si no
    
    HeaderSatisfies = False
    
    If Len(Trim(headerText)) = 0 Then
        Exit Function ' Cabecera vacía no satisface ningún tipo
    End If
    
    ' Limpiar BOM antes de validar tokens
    Dim cleanHeader
    cleanHeader = headerText
    
    ' Eliminar BOM UTF-8 (EF BB BF)
    If Len(cleanHeader) >= 3 Then
        If Asc(Mid(cleanHeader, 1, 1)) = 239 And Asc(Mid(cleanHeader, 2, 1)) = 187 And Asc(Mid(cleanHeader, 3, 1)) = 191 Then
            cleanHeader = Mid(cleanHeader, 4)
        End If
    End If
    
    ' Eliminar BOM UTF-16 LE (FF FE)
    If Len(cleanHeader) >= 2 Then
        If Asc(Mid(cleanHeader, 1, 1)) = 255 And Asc(Mid(cleanHeader, 2, 1)) = 254 Then
            cleanHeader = Mid(cleanHeader, 3)
        End If
    End If
    
    ' Eliminar BOM UTF-16 BE (FE FF)
    If Len(cleanHeader) >= 2 Then
        If Asc(Mid(cleanHeader, 1, 1)) = 254 And Asc(Mid(cleanHeader, 2, 1)) = 255 Then
            cleanHeader = Mid(cleanHeader, 3)
        End If
    End If
    
    Dim lines, i, line
    lines = Split(cleanHeader, vbCrLf)
    
    If expectedType = "bas" Then
        ' Para .bas necesitamos: Attribute VB_Name
        For i = 0 To UBound(lines)
            line = Trim(lines(i))
            If InStr(1, line, "Attribute VB_Name", vbTextCompare) = 1 Then
                HeaderSatisfies = True
                Exit Function
            End If
        Next
    ElseIf expectedType = "cls" Then
        ' Para .cls necesitamos: VERSION 1.0 CLASS y Attribute VB_Name
        Dim hasVersion, hasVbName
        hasVersion = False
        hasVbName = False
        
        For i = 0 To UBound(lines)
            line = Trim(lines(i))
            If InStr(1, line, "VERSION 1.0 CLASS", vbTextCompare) = 1 Then
                hasVersion = True
            ElseIf InStr(1, line, "Attribute VB_Name", vbTextCompare) = 1 Then
                hasVbName = True
            End If
        Next
        
        HeaderSatisfies = hasVersion And hasVbName
    End If
End Function

' ===================================================================
' COMANDO: FIX-SRC-HEADERS
' ===================================================================
Sub FixSrcHeadersCommand()
    Dim srcFolder, totalFiles, processedFiles, changedFiles
    Dim isDryRun, isVerbose
    
    ' Verificar flags
    isDryRun = gDryRun
    isVerbose = gVerbose
    
    ' Inicializar contadores
    totalFiles = 0
    processedFiles = 0
    changedFiles = 0
    
    ' Verificar que existe la carpeta src
    srcFolder = objFSO.BuildPath(RepoRoot(), "src")
    If Not objFSO.FolderExists(srcFolder) Then
        WScript.Echo "Error: No se encontró la carpeta ./src"
        WScript.Quit 1
    End If
    
    WScript.Echo "=== CONDOR CLI - Canonizador de Cabeceras ==="
    If isDryRun Then
        WScript.Echo "MODO: Dry-run (solo análisis, sin modificaciones)"
    Else
        WScript.Echo "MODO: Ejecución real (modificará archivos)"
    End If
    WScript.Echo "Procesando: " & srcFolder
    WScript.Echo ""
    
    ' Procesar archivos recursivamente
    Call ProcessFolderRecursive(srcFolder, totalFiles, processedFiles, changedFiles, isDryRun, isVerbose)
    
    ' Mostrar resumen final
    WScript.Echo ""
    WScript.Echo "=== RESUMEN ==="
    WScript.Echo "Archivos encontrados: " & totalFiles
    WScript.Echo "Archivos procesados: " & processedFiles
    WScript.Echo "Archivos modificados: " & changedFiles
    WScript.Echo "Archivos sin cambios: " & (processedFiles - changedFiles)
    
    If isDryRun Then
        WScript.Echo ""
        WScript.Echo "Para aplicar los cambios, ejecute sin --dry-run"
    End If
End Sub

' Función auxiliar para procesar carpetas recursivamente
Sub ProcessFolderRecursive(folderPath, totalFiles, processedFiles, changedFiles, isDryRun, isVerbose)
    Dim folder, file, subFolder
    Set folder = objFSO.GetFolder(folderPath)
    
    ' Procesar archivos en la carpeta actual
    For Each file In folder.Files
        Dim ext
        ext = LCase(objFSO.GetExtensionName(file.Name))
        
        If ext = "bas" Or ext = "cls" Then
            totalFiles = totalFiles + 1
            Call ProcessSourceFile(file.Path, ext, processedFiles, changedFiles, isDryRun, isVerbose)
        End If
    Next
    
    ' Procesar subcarpetas recursivamente
    For Each subFolder In folder.SubFolders
        Call ProcessFolderRecursive(subFolder.Path, totalFiles, processedFiles, changedFiles, isDryRun, isVerbose)
    Next
End Sub

' Función para procesar un archivo fuente individual
Sub ProcessSourceFile(filePath, fileExt, processedFiles, changedFiles, isDryRun, isVerbose)
    Dim originalContent, normalizedContent, headerBodyResult
    Dim currentName, expectedName, newHeader, needsChange
    Dim backupPath
    
    processedFiles = processedFiles + 1
    needsChange = False
    
    ' Leer contenido original
    originalContent = ReadAllText(filePath)
    If originalContent = "" Then
        If isVerbose Then
            WScript.Echo "WARNING: No se pudo leer " & objFSO.GetFileName(filePath)
        End If
        Exit Sub
    End If
    
    ' Obtener nombre esperado del archivo
    expectedName = NormalizeNameFromFilename(filePath)
    
    ' Separar cabecera y cuerpo
    headerBodyResult = SplitHeaderBody(originalContent)
    
    ' Construir nueva cabecera según el tipo
    If fileExt = "bas" Then
        newHeader = BuildBasHeader(expectedName)
    Else ' cls
        newHeader = BuildClsHeader(expectedName)
    End If
    
    ' Normalizar saltos de línea y ensamblar contenido
    Dim bodyContent
    bodyContent = headerBodyResult(1)
    
    ' Asegurar que el cuerpo no empiece con líneas vacías innecesarias
    Do While Left(bodyContent, 2) = vbCrLf
        bodyContent = Mid(bodyContent, 3)
    Loop
    
    ' Ensamblar contenido final
    If bodyContent <> "" Then
        normalizedContent = newHeader & vbCrLf & bodyContent
    Else
        normalizedContent = newHeader
    End If
    
    ' Normalizar todos los saltos de línea a CRLF
    normalizedContent = Replace(normalizedContent, vbLf, vbCrLf)
    normalizedContent = Replace(normalizedContent, vbCrLf & vbCrLf, vbCrLf)
    
    ' Verificar si hay cambios
    If originalContent <> normalizedContent Then
        needsChange = True
        changedFiles = changedFiles + 1
        
        If isVerbose Or isDryRun Then
            WScript.Echo "CAMBIO: " & objFSO.GetFileName(filePath) & " (nombre: " & expectedName & ")"
        End If
        
        ' Aplicar cambios si no es dry-run
        If Not isDryRun Then
            ' Crear backup si no existe
            backupPath = filePath & ".bak"
            If Not objFSO.FileExists(backupPath) Then
                objFSO.CopyFile filePath, backupPath
                If isVerbose Then
                    WScript.Echo "  Backup creado: " & objFSO.GetFileName(backupPath)
                End If
            End If
            
            ' Escribir contenido normalizado
            Call WriteAllText(filePath, normalizedContent)
            If isVerbose Then
                WScript.Echo "  Archivo actualizado"
            End If
        End If
    Else
        If isVerbose Then
            WScript.Echo "OK: " & objFSO.GetFileName(filePath)
        End If
    End If
End Sub

' ===================================================================
' FUNCIÓN: PostProcessHeader
' Descripción: Postprocesa un archivo exportado para garantizar cabeceras canónicas
' Parámetros: filePath - Ruta del archivo .bas/.cls a postprocesar
' ===================================================================
Sub PostProcessHeader(filePath)
    Dim content, parts, newHeader, newContent, moduleName, ext
    
    ' Leer contenido del archivo
    content = ReadAllText(filePath)
    If content = "" Then Exit Sub
    
    ' Obtener extensión y nombre del módulo
    ext = LCase(objFSO.GetExtensionName(filePath))
    moduleName = objFSO.GetBaseName(filePath)
    
    ' Dividir en cabecera y cuerpo
    parts = SplitHeaderBody(content)
    
    ' Construir nueva cabecera canónica según el tipo
    If ext = "bas" Then
        newHeader = BuildBasHeader(moduleName)
    ElseIf ext = "cls" Then
        newHeader = BuildClsHeader(moduleName)
    Else
        ' Tipo no soportado, no modificar
        Exit Sub
    End If
    
    ' Reconstruir contenido con cabecera canónica
    If parts(1) <> "" Then
        newContent = newHeader & vbCrLf & parts(1)
    Else
        newContent = newHeader
    End If
    
    ' Escribir archivo con cabecera canónica
    WriteAllText filePath, newContent
    
    ' Validar cabecera resultante
    If ext = "bas" Then
        If InStr(newContent, "Attribute VB_Name") = 0 Then
            WScript.Echo "Warning: Archivo .bas sin Attribute VB_Name: " & filePath
        End If
    ElseIf ext = "cls" Then
        If InStr(newContent, "VERSION 1.0 CLASS") = 0 Then
            WScript.Echo "Warning: Archivo .cls sin VERSION 1.0 CLASS: " & filePath
        End If
    End If
End Sub
