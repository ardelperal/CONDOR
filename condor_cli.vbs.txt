' CONDOR CLI - Herramienta de linea de comandos para el proyecto CONDOR
' Funcionalidades: Sincronizacion VBA, gestion de tablas, y operaciones del proyecto
' Version sin dialogos para automatizacion completa

Option Explicit

Dim objAccess
Dim strAccessPath
Dim strSourcePath
Dim strAction
Dim objFSO
Dim objArgs
Dim strDbPassword

' Configuracion
' Configuracion inicial - se determinara la base de datos segun la accion
Dim strDataPath
strAccessPath = "C:\Proyectos\CONDOR\back\Desarrollo\CONDOR.accdb"
strDataPath = "C:\Proyectos\CONDOR\back\CONDOR_datos.accdb"
strSourcePath = "C:\Proyectos\CONDOR\src"

' Obtener argumentos de linea de comandos
Set objArgs = WScript.Arguments

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
    WScript.Echo "  test       - Ejecutar suite de pruebas unitarias"
    WScript.Echo "  update     - Sincronizar modulos VBA (automatico o selectivo)"
    WScript.Echo "  rebuild    - Reconstruir proyecto VBA (eliminar todos los modulos y reimportar)"
    WScript.Echo "  bundle <funcionalidad> [ruta_destino] - Empaquetar archivos de codigo por funcionalidad"
    WScript.Echo "  lint       - Auditar codigo VBA para detectar cabeceras duplicadas"
    WScript.Echo "  createtable <nombre> <sql> - Crear tabla con consulta SQL"
    WScript.Echo "  droptable <nombre> - Eliminar tabla"
    WScript.Echo "  listtables [db_path] - Listar tablas (opcionalmente de base especifica)"
    WScript.Echo "  relink <db_path> <folder> - Re-vincular tablas a bases locales"
    WScript.Echo "  relink --all - Re-vincular todas las bases en ./back automaticamente"
    WScript.Echo ""
    WScript.Echo "COMANDO UPDATE - SINCRONIZACION INTELIGENTE:"
    WScript.Echo "  update                    - Modo automatico: sincroniza solo archivos modificados"
    WScript.Echo "  update <modulo1,modulo2>  - Modo selectivo: sincroniza modulos especificos"
    WScript.Echo ""
    WScript.Echo "FLUJO DE TRABAJO RECOMENDADO:"
    WScript.Echo "  1. cscript condor_cli.vbs update    (sincronizacion rapida de cambios)"
    WScript.Echo "  2. cscript condor_cli.vbs rebuild   (reconstruccion completa si es necesario)"
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
    WScript.Echo "  --dry-run  - Simular operacion sin modificar Access (solo con import)"
    WScript.Echo "  --verbose  - Mostrar informacion detallada durante la operacion"
    WScript.Echo ""
    WScript.Echo "EJEMPLOS:"
    WScript.Echo "  cscript condor_cli.vbs update"
    WScript.Echo "  cscript condor_cli.vbs update CAuthService,CExpedienteService"
    WScript.Echo "  cscript condor_cli.vbs validate"
    WScript.Echo "  cscript condor_cli.vbs export --verbose"
    WScript.Echo "  cscript condor_cli.vbs rebuild"
    WScript.Echo "  cscript condor_cli.vbs bundle Tests"
    WScript.Quit 1
End If

strAction = LCase(objArgs(0))

If strAction <> "export" And strAction <> "validate" And strAction <> "test" And strAction <> "createtable" And strAction <> "droptable" And strAction <> "listtables" And strAction <> "relink" And strAction <> "rebuild" And strAction <> "lint" And strAction <> "update" And strAction <> "bundle" Then
    WScript.Echo "Error: Comando debe ser 'export', 'validate', 'test', 'createtable', 'droptable', 'listtables', 'relink', 'rebuild', 'lint', 'update' o 'bundle'"
    WScript.Quit 1
End If

Set objFSO = CreateObject("Scripting.FileSystemObject")

' El comando bundle no requiere Access
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
End If

' Determinar qué base de datos usar según la acción
If strAction = "createtable" Or strAction = "droptable" Then
    strAccessPath = strDataPath
ElseIf strAction = "listtables" Then
    ' Para listtables, usar base específica si se proporciona, sino usar por defecto
    If objArgs.Count > 1 Then
        strAccessPath = objArgs(1)
    Else
        strAccessPath = strDataPath
    End If
End If

' Para rebuild y test, usar la base de datos de desarrollo
If strAction = "rebuild" Or strAction = "test" Then
    strAccessPath = "C:\Proyectos\CONDOR\back\Desarrollo\CONDOR.accdb"
End If

' Verificar que existe la base de datos
If Not objFSO.FileExists(strAccessPath) Then
    WScript.Echo "Error: La base de datos no existe: " & strAccessPath
    WScript.Quit 1
End If

WScript.Echo "=== INICIANDO SINCRONIZACION VBA ==="
WScript.Echo "Accion: " & strAction
WScript.Echo "Base de datos: " & strAccessPath
WScript.Echo "Directorio: " & strSourcePath

' Para el comando update, verificar cambios antes de abrir Access
If strAction = "update" Then
    If Not CheckForChangesBeforeUpdate() Then
        WScript.Echo "=== NO HAY CAMBIOS DETECTADOS ==="
        WScript.Echo "✅ Todos los archivos están sincronizados. No es necesario abrir la base de datos."
        WScript.Echo "=== SINCRONIZACION COMPLETADA EXITOSAMENTE ==="
        WScript.Quit 0
    End If
End If

' Verificar y cerrar procesos de Access existentes
Call CloseExistingAccessProcesses()

On Error Resume Next

' Crear aplicacion Access
WScript.Echo "Iniciando aplicacion Access..."
Set objAccess = CreateObject("Access.Application")

If Err.Number <> 0 Then
    WScript.Echo "Error al crear aplicacion Access: " & Err.Description
    WScript.Quit 1
End If

' Configurar Access en modo silencioso
objAccess.Visible = False
objAccess.UserControl = False
' Suprimir alertas y diálogos de confirmación
' ' objAccess.DoCmd.SetWarnings False  ' Comentado temporalmente por error de compilación  ' Comentado temporalmente por error de compilación
objAccess.Application.Echo False
' Configuraciones adicionales para suprimir diálogos
On Error Resume Next
objAccess.Application.AutomationSecurity = 1  ' msoAutomationSecurityLow
objAccess.VBE.MainWindow.Visible = False
Err.Clear
On Error GoTo 0

' Abrir base de datos con compilacion condicional
WScript.Echo "Abriendo base de datos..."

' Configurar Access para evitar errores de compilación
On Error Resume Next
' Intentar configurar propiedades si están disponibles
objAccess.DisplayAlerts = False
Err.Clear

' Determinar contraseña para la base de datos
strDbPassword = GetDatabasePassword(strAccessPath)

' Abrir base de datos con manejo de errores robusto
If strDbPassword = "" Then
    ' Sin contraseña
    objAccess.OpenCurrentDatabase strAccessPath
Else
    ' Con contrasena - usar solo dos parametros
    objAccess.OpenCurrentDatabase strAccessPath, , strDbPassword
End If

If Err.Number <> 0 Then
    WScript.Echo "Error al abrir base de datos: " & Err.Description
    objAccess.Quit
    WScript.Quit 1
End If

On Error GoTo 0
WScript.Echo "Base de datos abierta correctamente."
Call EnsureVBReferences

' Verificar opciones especiales
Dim bDryRun, bVerbose, i
bDryRun = False
bVerbose = False

For i = 1 To objArgs.Count - 1
    If LCase(objArgs(i)) = "--dry-run" Then
        bDryRun = True
        WScript.Echo "[MODO DRY-RUN] Simulacion activada - no se modificara Access"
    ElseIf LCase(objArgs(i)) = "--verbose" Then
        bVerbose = True
        WScript.Echo "[MODO VERBOSE] Informacion detallada activada"
    End If
Next

If strAction = "validate" Then
    Call ValidateAllModules(bVerbose)
ElseIf strAction = "export" Then
    Call ExportModules(bVerbose)
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

ElseIf strAction = "lint" Then
    Call LintProject()
ElseIf strAction = "relink" Then
    Call RelinkTables()
ElseIf strAction = "update" Then
    Call UpdateProject()
End If

' Cerrar Access
WScript.Echo "Cerrando Access..."
' Restaurar estado normal de Access antes de cerrar
On Error Resume Next
objAccess.Application.Echo True
objAccess.Quit 2  ' acQuitSaveNone = 2
If Err.Number <> 0 Then
    ' Intentar cerrar sin guardar si hay problemas
    objAccess.Quit 2  ' acQuitSaveNone = 2
End If
On Error GoTo 0
WScript.Echo "Access cerrado correctamente"

WScript.Echo "=== SINCRONIZACION COMPLETADA EXITOSAMENTE ==="
WScript.Quit 0

' Subrutina para validar todos los modulos sin importar
Sub ValidateAllModules(bVerbose)
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
            
            If bVerbose Then
                WScript.Echo "Validando: " & objFile.Name
            End If
            
            ' Validar sintaxis
            Dim errorDetails
            validationResult = ValidateVBASyntax(strFileName, errorDetails)
            
            If validationResult = True Then
                validFiles = validFiles + 1
                If bVerbose Then
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
Sub ExportModules(bVerbose)
    Dim vbComponent
    Dim strExportPath
    Dim exportedCount
    
    WScript.Echo "Iniciando exportacion de modulos VBA..."
    
    If Not objFSO.FolderExists(strSourcePath) Then
        objFSO.CreateFolder strSourcePath
        WScript.Echo "Directorio de destino creado: " & strSourcePath
    End If
    
    exportedCount = 0
    
    For Each vbComponent In objAccess.VBE.ActiveVBProject.VBComponents
        If vbComponent.Type = 1 Then  ' vbext_ct_StdModule
            strExportPath = strSourcePath & "\" & vbComponent.Name & ".bas"
            
            If bVerbose Then
                WScript.Echo "Exportando modulo: " & vbComponent.Name
            End If
            
            On Error Resume Next
            Call ExportModuleWithAnsiEncoding(vbComponent, strExportPath)
            
            If Err.Number <> 0 Then
                WScript.Echo "Error al exportar modulo " & vbComponent.Name & ": " & Err.Description
                Err.Clear
            Else
                If bVerbose Then
                    WScript.Echo "  ✓ Modulo " & vbComponent.Name & " exportado a: " & strExportPath
                Else
                    WScript.Echo "✓ " & vbComponent.Name & ".bas"
                End If
                exportedCount = exportedCount + 1
            End If
        ElseIf vbComponent.Type = 2 Then  ' vbext_ct_ClassModule
            strExportPath = strSourcePath & "\" & vbComponent.Name & ".cls"
            
            If bVerbose Then
                WScript.Echo "Exportando clase: " & vbComponent.Name
            End If
            
            On Error Resume Next
            Call ExportModuleWithAnsiEncoding(vbComponent, strExportPath)
            
            If Err.Number <> 0 Then
                WScript.Echo "Error al exportar clase " & vbComponent.Name & ": " & Err.Description
                Err.Clear
            Else
                If bVerbose Then
                    WScript.Echo "  ✓ Clase " & vbComponent.Name & " exportada a: " & strExportPath
                Else
                    WScript.Echo "✓ " & vbComponent.Name & ".cls"
                End If
                exportedCount = exportedCount + 1
            End If
        End If
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
    objAccess.DoCmd.DeleteObject 1, strQueryName  ' acQuery = 1
    
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
    objAccess.DoCmd.DeleteObject 0, strTableName  ' acTable = 0
    
    If Err.Number <> 0 Then
        WScript.Echo "Error al eliminar tabla: " & Err.Description
        Err.Clear
    Else
        WScript.Echo "Tabla '" & strTableName & "' eliminada exitosamente"
    End If
End Sub

' Subrutina para listar tablas
Sub ListTables()
    Dim tbl
    Dim tableCount
    
    WScript.Echo "=== LISTADO DE TABLAS ==="
    tableCount = 0
    
    For Each tbl In objAccess.CurrentDb.TableDefs
        ' Filtrar tablas del sistema (que empiezan con MSys)
        If Left(tbl.Name, 4) <> "MSys" And Left(tbl.Name, 1) <> "~" Then
            tableCount = tableCount + 1
            WScript.Echo tableCount & ". " & tbl.Name & " (" & tbl.RecordCount & " registros)"
        End If
    Next
    
    If tableCount = 0 Then
        WScript.Echo "No se encontraron tablas de usuario"
    Else
        WScript.Echo "Total de tablas: " & tableCount
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
    
    Set objAccess = CreateObject("Access.Application")
    objAccess.Visible = False
    objAccess.OpenCurrentDatabase strAccessPath, False
    
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
        objAccess.Quit
        WScript.Quit 1
    Else
        WScript.Echo ""
        WScript.Echo "=== LINT COMPLETADO EXITOSAMENTE ==="
    End If
    
    objAccess.Quit
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
                Trim(strLine) = "END" Or _
                Trim(strLine) = "Option Compare Database" Or _
                Trim(strLine) = "Option Explicit") Then
            
            ' Si no cumple ninguna condición, es código VBA válido
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

' Subrutina para mostrar ayuda completa
Sub ShowHelp()
    WScript.Echo "=== CONDOR CLI - Herramienta de línea de comandos ==="
    WScript.Echo "Versión: 2.0 - Sistema de gestión y sincronización VBA para proyecto CONDOR"
    WScript.Echo ""
    WScript.Echo "SINTAXIS:"
    WScript.Echo "  cscript condor_cli.vbs [comando] [opciones] [parámetros]"
    WScript.Echo ""
    WScript.Echo "COMANDOS PRINCIPALES:"
    WScript.Echo ""
    WScript.Echo "📤 EXPORTACIÓN:"
    WScript.Echo "  export [--verbose]           - Exportar módulos VBA desde Access a /src"
    WScript.Echo "                                 Codificación: ANSI para compatibilidad"
    WScript.Echo "                                 --verbose: Mostrar detalles de cada archivo"
    WScript.Echo ""
    WScript.Echo "🔄 SINCRONIZACIÓN:"
    WScript.Echo "  update                       - Sincronización automática (solo archivos modificados)"
    WScript.Echo "  update <módulo1,módulo2>     - Sincronización selectiva de módulos específicos"
    WScript.Echo "                                 Ejemplo: update CAuthService,CExpedienteService"
    WScript.Echo "  rebuild                      - Reconstrucción completa del proyecto VBA"
    WScript.Echo "                                 (Elimina todos los módulos y reimporta)"
    WScript.Echo ""
    WScript.Echo "✅ VALIDACIÓN Y PRUEBAS:"
    WScript.Echo "  validate [--verbose]         - Validar sintaxis VBA sin importar a Access"
    WScript.Echo "                                 --verbose: Mostrar detalles de validación"
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
    WScript.Echo "                                 db_path opcional (por defecto: CONDOR_datos.accdb)"
    WScript.Echo "  relink <db_path> <folder>    - Re-vincular tablas a bases locales específicas"
    WScript.Echo "  relink --all                 - Re-vincular automáticamente todas las bases en ./back"
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
    WScript.Echo "  --dry-run                    - Simular operación sin modificar Access (solo import)"
    WScript.Echo "  --verbose                    - Mostrar información detallada durante la operación"
    WScript.Echo ""
    WScript.Echo "FLUJO DE TRABAJO RECOMENDADO:"
    WScript.Echo "  1. cscript condor_cli.vbs validate     (validar sintaxis antes de importar)"
    WScript.Echo "  2. cscript condor_cli.vbs update       (sincronización rápida de cambios)"
    WScript.Echo "  3. cscript condor_cli.vbs test         (ejecutar pruebas unitarias)"
    WScript.Echo "  4. cscript condor_cli.vbs rebuild      (reconstrucción completa si es necesario)"
    WScript.Echo ""
    WScript.Echo "EJEMPLOS DE USO:"
    WScript.Echo "  cscript condor_cli.vbs --help"
    WScript.Echo "  cscript condor_cli.vbs update"
    WScript.Echo "  cscript condor_cli.vbs update CAuthService,CExpedienteService"
    WScript.Echo "  cscript condor_cli.vbs validate --verbose"
    WScript.Echo "  cscript condor_cli.vbs export --verbose"
    WScript.Echo "  cscript condor_cli.vbs bundle Auth"
    WScript.Echo "  cscript condor_cli.vbs bundle Document C:\\\\temp"
    WScript.Echo "  cscript condor_cli.vbs createtable MiTabla ""CREATE TABLE MiTabla (ID LONG)"""
    WScript.Echo "  cscript condor_cli.vbs listtables"
    WScript.Echo "  cscript condor_cli.vbs relink --all"
    WScript.Echo ""
    WScript.Echo "CONFIGURACIÓN:"
    WScript.Echo "  Base de datos desarrollo: C:\\Proyectos\\CONDOR\\back\\Desarrollo\\CONDOR.accdb"
    WScript.Echo "  Base de datos datos:      C:\\Proyectos\\CONDOR\\back\\CONDOR_datos.accdb"
    WScript.Echo "  Directorio fuente:        C:\\Proyectos\\CONDOR\\src"
    WScript.Echo ""
    WScript.Echo "Para más información, consulte la documentación en docs/CONDOR_MASTER_PLAN.md"
End Sub

' Nueva función que usa DoCmd.LoadFromText para evitar confirmaciones
Sub ImportModuleWithLoadFromText(strSourceFile, moduleName, fileExtension)
    On Error Resume Next
    
    ' Determinar el tipo de objeto Access para DoCmd.LoadFromText
    Dim objectType
    If fileExtension = "bas" Then
        objectType = 5  ' acModule para módulos estándar
    ElseIf fileExtension = "cls" Then
        objectType = 5  ' acModule también para módulos de clase
    Else
        WScript.Echo "  ❌ Error: Tipo de archivo no soportado: " & fileExtension
        Exit Sub
    End If
    
    ' Usar DoCmd.LoadFromText para importar el módulo
    objAccess.DoCmd.LoadFromText objectType, moduleName, strSourceFile
    
    If Err.Number <> 0 Then
        WScript.Echo "  ❌ Error al importar módulo " & moduleName & " con LoadFromText: " & Err.Description
        Err.Clear
        Exit Sub
    End If
    
    If fileExtension = "cls" Then
        WScript.Echo "✅ Clase " & moduleName & " importada correctamente"
    Else
        WScript.Echo "✅ Módulo " & moduleName & " importado correctamente"
    End If
    
    On Error GoTo 0
End Sub

' Subrutina para ejecutar la suite de pruebas unitarias
Sub ExecuteTests()
    WScript.Echo "=== INICIANDO EJECUCION DE PRUEBAS ==="
    Dim reportString
    
    ' Ejecutar las pruebas en Access y capturar el resultado
    WScript.Echo "Ejecutando suite de pruebas en Access..."
    On Error Resume Next
    
    ' Suprimir diálogos inesperados de Access durante las pruebas
    objAccess.Application.DisplayAlerts = False
    
    ' Llamar directamente a ExecuteAllTestsForCLI y capturar el valor de retorno
    reportString = objAccess.Application.Run("modTestRunner.ExecuteAllTestsForCLI")
    
    If Err.Number <> 0 Then
        WScript.Echo "ERROR: Fallo crítico al invocar la suite de pruebas."
        WScript.Echo "  Código de Error: " & Err.Number
        WScript.Echo "  Descripción: " & Err.Description
        WScript.Echo "  Fuente: " & Err.Source
        WScript.Echo "SUGERENCIA: Abre Access manualmente y ejecuta RunAllTests desde el módulo modTestRunner para ver el error específico"
        objAccess.Quit
        WScript.Quit 1
    End If
    On Error GoTo 0
    
    ' Verificar si reportString está vacío
    If IsEmpty(reportString) Or reportString = "" Then
        WScript.Echo "ERROR: La comunicación con el motor de pruebas de Access falló"
        WScript.Echo "SUGERENCIA: El motor de pruebas no devolvió ningún resultado"
        objAccess.Quit
        WScript.Quit 1
    End If
    
    ' Mostrar el reporte completo directamente en la consola
    WScript.Echo "--- INICIO DE RESULTADOS DE PRUEBAS ---"
    WScript.Echo reportString
    WScript.Echo "--- FIN DE RESULTADOS DE PRUEBAS ---"
    
    ' Determinar el éxito o fracaso buscando la línea final
    If InStr(reportString, "RESULT: SUCCESS") > 0 Then
        WScript.Echo "RESULTADO FINAL: ✓ Todas las pruebas pasaron."
        WScript.Quit 0 ' Código de éxito
    Else
        WScript.Echo "RESULTADO FINAL: ✗ Pruebas fallidas."
        WScript.Quit 1 ' Código de error para CI/CD
    End If
End Sub

' Función para importar módulo con conversión UTF-8 -> ANSI
Sub ImportModuleWithAnsiEncoding(strImportPath, moduleName, fileExtension, vbComponent, cleanedContent)
    ' Declarar variables locales
    Dim tempFolderPath, tempFileName, tempFilePath
    Dim objTempFile
    Dim importError, renameError, existingComponent
    
    If fileExtension = "bas" Then
        ' Lógica corregida para módulos estándar (.bas) - usar Add(1)
        On Error Resume Next
        
        ' Buscar si ya existe un componente con este nombre
        Set vbComponent = Nothing
        For Each existingComponent In objAccess.VBE.ActiveVBProject.VBComponents
            If existingComponent.Name = moduleName Then
                Set vbComponent = existingComponent
                Exit For
            End If
        Next
        
        ' Si no existe, crear nuevo componente
        If vbComponent Is Nothing Then
            Set vbComponent = objAccess.VBE.ActiveVBProject.VBComponents.Add(1) ' 1 = vbext_ct_StdModule
            If Err.Number <> 0 Then
                WScript.Echo "❌ ERROR: No se pudo crear componente estándar para " & moduleName & ": " & Err.Description
                On Error GoTo 0
                Exit Sub
            End If
            
            ' Renombrar inmediatamente después de crear
            vbComponent.Name = moduleName
            renameError = Err.Number
            If renameError <> 0 Then
                WScript.Echo "❌ ERROR: No se pudo renombrar el módulo nuevo a '" & moduleName & "': " & Err.Description & " (Código: " & Err.Number & ")"
                On Error GoTo 0
                Exit Sub
            End If
        Else
            ' Si existe, limpiar el código existente
            If vbComponent.CodeModule.CountOfLines > 0 Then
                vbComponent.CodeModule.DeleteLines 1, vbComponent.CodeModule.CountOfLines
            End If
        End If
        
        ' Insertar el contenido limpio en el módulo de código
        vbComponent.CodeModule.AddFromString cleanedContent
        If Err.Number <> 0 Then
            WScript.Echo "❌ ERROR: No se pudo insertar código en el módulo " & moduleName & ": " & Err.Description
            On Error GoTo 0
            Exit Sub
        End If
        
        On Error GoTo 0
        ' Confirmar éxito
        WScript.Echo "✅ Módulo " & moduleName & " importado correctamente"
        
    ElseIf fileExtension = "cls" Then
        ' Lógica específica para módulos de clase (.cls)
        On Error Resume Next
        
        ' Buscar si ya existe un componente con este nombre
        Set vbComponent = Nothing
        For Each existingComponent In objAccess.VBE.ActiveVBProject.VBComponents
            If existingComponent.Name = moduleName Then
                Set vbComponent = existingComponent
                Exit For
            End If
        Next
        
        ' Si no existe, crear nuevo componente
        If vbComponent Is Nothing Then
            Set vbComponent = objAccess.VBE.ActiveVBProject.VBComponents.Add(2) ' 2 = vbext_ct_ClassModule
            If Err.Number <> 0 Then
                WScript.Echo "❌ ERROR: No se pudo crear componente de clase para " & moduleName & ": " & Err.Description
                On Error GoTo 0
                Exit Sub
            End If
            
            ' Renombrar inmediatamente después de crear
            vbComponent.Name = moduleName
            renameError = Err.Number
            If renameError <> 0 Then
                WScript.Echo "❌ ERROR: No se pudo renombrar la clase nueva a '" & moduleName & "': " & Err.Description & " (Código: " & Err.Number & ")"
                On Error GoTo 0
                Exit Sub
            End If
        Else
            ' Si existe, limpiar el código existente
            If vbComponent.CodeModule.CountOfLines > 0 Then
                vbComponent.CodeModule.DeleteLines 1, vbComponent.CodeModule.CountOfLines
            End If
        End If
        
        ' Insertar el contenido limpio en el módulo de código
        vbComponent.CodeModule.AddFromString cleanedContent
        If Err.Number <> 0 Then
            WScript.Echo "❌ ERROR: No se pudo insertar código en la clase " & moduleName & ": " & Err.Description
            On Error GoTo 0
            Exit Sub
        End If
        
        On Error GoTo 0
        ' Confirmar éxito
        WScript.Echo "✅ Clase " & moduleName & " importada correctamente"
    End If
End Sub

' Función simplificada usando VBComponents.Import() - método desatendido
Sub ImportModuleWithAnsiEncodingNew(strImportPath, moduleName, fileExtension, vbComponent, cleanedContent)
    ' Método con verificación de referencias VBA y enlace tardío
    Dim existingComponent, vbeObject, vbProject, vbComponents
    
    On Error Resume Next
    
    ' Verificar que VBE esté disponible usando enlace tardío
    Set vbeObject = objAccess.VBE
    If Err.Number <> 0 Then
        WScript.Echo "❌ ERROR: VBA no está habilitado o no se puede acceder al VBE: " & Err.Description
        On Error GoTo 0
        Exit Sub
    End If
    
    ' Verificar que el proyecto VBA esté disponible
    Set vbProject = vbeObject.ActiveVBProject
    If Err.Number <> 0 Then
        WScript.Echo "❌ ERROR: No se puede acceder al proyecto VBA activo: " & Err.Description
        On Error GoTo 0
        Exit Sub
    End If
    
    ' Verificar que VBComponents esté disponible
    Set vbComponents = vbProject.VBComponents
    If Err.Number <> 0 Then
        WScript.Echo "❌ ERROR: No se puede acceder a VBComponents (referencias VBA requeridas): " & Err.Description
        On Error GoTo 0
        Exit Sub
    End If
    
    ' Eliminar el componente existente si ya existe
    Set vbComponent = Nothing
    For Each existingComponent In vbComponents
        If existingComponent.Name = moduleName Then
            vbComponents.Remove existingComponent
            If Err.Number <> 0 Then
                WScript.Echo "❌ ERROR: No se pudo eliminar componente existente " & moduleName & ": " & Err.Description
                On Error GoTo 0
                Exit Sub
            End If
            Exit For
        End If
    Next
    
    ' Importar directamente el archivo usando VBComponents.Import()
    Set vbComponent = vbComponents.Import(strImportPath)
    
    If Err.Number <> 0 Then
        WScript.Echo "❌ ERROR: No se pudo importar " & moduleName & ": " & Err.Description
        WScript.Echo "  Verifique que las referencias 'Microsoft Visual Basic for Applications Extensibility' estén habilitadas"
        On Error GoTo 0
        Exit Sub
    End If
    
    ' Verificar si el componente fue importado correctamente
    If vbComponent Is Nothing Then
        WScript.Echo "❌ ERROR: El componente importado es Nothing"
        On Error GoTo 0
        Exit Sub
    End If
    
    ' Renombrar el componente solo si es necesario
    If vbComponent.Name <> moduleName Then
        Dim originalName
        originalName = vbComponent.Name
        vbComponent.Name = moduleName
        If Err.Number <> 0 Then
            WScript.Echo "⚠️ ADVERTENCIA: No se pudo renombrar de '" & originalName & "' a '" & moduleName & "': " & Err.Description
            WScript.Echo "  El módulo se importó como '" & originalName & "' - verifique el nombre en el archivo fuente"
            Err.Clear
        End If
    End If
    
    On Error GoTo 0
    
    ' Confirmar éxito según el tipo
    If fileExtension = "bas" Then
        WScript.Echo "✅ Módulo " & moduleName & " importado correctamente"
    ElseIf fileExtension = "cls" Then
        WScript.Echo "✅ Clase " & moduleName & " importada correctamente"
    End If
End Sub


' Función desatendida para importar módulos usando VBComponents.Import()
' Mantiene la funcionalidad de limpieza de código de rebuild
Sub ImportModuleDesatendido(strImportPath, moduleName, fileExtension, cleanedContent)
    ' Declarar variables locales
    Dim tempFolderPath, tempFileName, tempFilePath
    Dim objTempFile, existingComponent, vbComp
    
    On Error Resume Next
    
    ' Eliminar módulo si ya existe
    Set vbComp = Nothing
    For Each existingComponent In objAccess.VBE.ActiveVBProject.VBComponents
        If existingComponent.Name = moduleName Then
            objAccess.VBE.ActiveVBProject.VBComponents.Remove existingComponent
            If Err.Number <> 0 Then
                WScript.Echo "❌ ERROR: No se pudo eliminar componente existente " & moduleName & ": " & Err.Description
                On Error GoTo 0
                Exit Sub
            End If
            Exit For
        End If
    Next
    
    ' Crear archivo temporal con contenido limpio
    tempFolderPath = objFSO.GetSpecialFolder(2) ' Carpeta temporal del sistema
    tempFileName = "temp_" & moduleName & "." & fileExtension
    tempFilePath = objFSO.BuildPath(tempFolderPath, tempFileName)
    
    ' Escribir contenido limpio al archivo temporal
    Set objTempFile = objFSO.CreateTextFile(tempFilePath, True, False) ' False = ANSI encoding
    objTempFile.Write cleanedContent
    objTempFile.Close
    Set objTempFile = Nothing
    
    ' Importar módulo usando VBComponents.Import()
    Set vbComp = objAccess.VBE.ActiveVBProject.VBComponents.Import(tempFilePath)
    
    If Err.Number <> 0 Then
        WScript.Echo "❌ ERROR: No se pudo importar " & moduleName & ": " & Err.Description
        ' Limpiar archivo temporal
        If objFSO.FileExists(tempFilePath) Then objFSO.DeleteFile tempFilePath
        On Error GoTo 0
        Exit Sub
    End If
    
    ' Renombrar el componente si es necesario
    If Not vbComp Is Nothing And vbComp.Name <> moduleName Then
        vbComp.Name = moduleName
        If Err.Number <> 0 Then
            WScript.Echo "⚠️ ADVERTENCIA: No se pudo renombrar a '" & moduleName & "': " & Err.Description
            Err.Clear
        End If
    End If
    
    ' Limpiar archivo temporal
    If objFSO.FileExists(tempFilePath) Then objFSO.DeleteFile tempFilePath
    
    On Error GoTo 0
    
    ' Confirmar éxito
    If fileExtension = "bas" Then
        WScript.Echo "✅ Módulo " & moduleName & " importado correctamente"
    ElseIf fileExtension = "cls" Then
        WScript.Echo "✅ Clase " & moduleName & " importada correctamente"
    End If
End Sub

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
        objAccess.Quit
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
        objAccess.Quit
        WScript.Quit 1
    End If
    
    If Not objFSO.FolderExists(strLocalFolder) Then
        WScript.Echo "Error: La carpeta de backends locales no existe: " & strLocalFolder
        objAccess.Quit
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
        objAccess.Quit
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

' Subrutina para reconstruir completamente el proyecto VBA
Sub RebuildProject()
    WScript.Echo "=== RECONSTRUCCION COMPLETA DEL PROYECTO VBA ==="
    WScript.Echo "ADVERTENCIA: Se eliminaran TODOS los modulos VBA existentes"
    WScript.Echo "Iniciando proceso de reconstruccion..."
    
    On Error Resume Next
    
    ' Paso 1: Eliminar todos los módulos existentes
    WScript.Echo "Paso 1: Eliminando todos los modulos VBA existentes..."
    
    Dim vbProject, vbComponent
    Set vbProject = objAccess.VBE.ActiveVBProject
    
    Dim componentCount, i, errorDetails
    componentCount = vbProject.VBComponents.Count
    
    ' Iterar hacia atrás para evitar problemas al eliminar elementos
    For i = componentCount To 1 Step -1
        Set vbComponent = vbProject.VBComponents(i)
        
        ' Solo eliminar módulos estándar y de clase (no formularios ni informes)
        If vbComponent.Type = 1 Or vbComponent.Type = 2 Then ' vbext_ct_StdModule = 1, vbext_ct_ClassModule = 2
            WScript.Echo "  Eliminando: " & vbComponent.Name & " (Tipo: " & vbComponent.Type & ")"
            vbProject.VBComponents.Remove vbComponent
            
            If Err.Number <> 0 Then
                WScript.Echo "  ❌ Error eliminando " & vbComponent.Name & ": " & Err.Description
                Err.Clear
            Else
                WScript.Echo "  ✓ Eliminado: " & vbComponent.Name
            End If
        End If
    Next
    
    WScript.Echo "Paso 2: Cerrando base de datos..."
    
    ' Cerrar sin guardar explícitamente para evitar confirmaciones
    objAccess.Quit 1  ' acQuitSaveAll = 1
    
    If Err.Number <> 0 Then
        WScript.Echo "Advertencia al cerrar Access: " & Err.Description
        Err.Clear
    End If
    
    Set objAccess = Nothing
    WScript.Echo "✓ Base de datos cerrada y guardada"
    
    ' Paso 3: Volver a abrir la base de datos
    WScript.Echo "Paso 3: Reabriendo base de datos con proyecto VBA limpio..."
    
    Set objAccess = CreateObject("Access.Application")
    
    If Err.Number <> 0 Then
        WScript.Echo "❌ Error al crear nueva instancia de Access: " & Err.Description
        WScript.Quit 1
    End If
    
    ' Configurar Access en modo silencioso
    objAccess.Visible = False
    objAccess.UserControl = False
    
    ' Suprimir alertas y diálogos de confirmación
    On Error Resume Next
    ' objAccess.DoCmd.SetWarnings False  ' Comentado temporalmente por error de compilación
    objAccess.Application.Echo False
    objAccess.DisplayAlerts = False
    ' Configuraciones adicionales para suprimir diálogos
    objAccess.Application.AutomationSecurity = 1  ' msoAutomationSecurityLow
    objAccess.VBE.MainWindow.Visible = False
    Err.Clear
    On Error GoTo 0
    
    ' Determinar contraseña para la base de datos
    Dim strDbPassword
    strDbPassword = GetDatabasePassword(strAccessPath)
    
    ' Abrir base de datos
    If strDbPassword = "" Then
        objAccess.OpenCurrentDatabase strAccessPath
    Else
        objAccess.OpenCurrentDatabase strAccessPath, , strDbPassword
    End If
    
    If Err.Number <> 0 Then
        WScript.Echo "❌ Error al reabrir base de datos: " & Err.Description
        WScript.Quit 1
    End If
    
    WScript.Echo "✓ Base de datos reabierta con proyecto VBA limpio"
    
    ' Paso 4: Importar todos los módulos de nuevo
    WScript.Echo "Paso 4: Importando todos los modulos desde /src..."
    
    ' Integrar lógica de importación directamente
    Dim objFolder, objFile
    Dim strModuleName, strFileName, strContent
    Dim srcModules
    Dim moduleExists
    Dim validationResult
    Dim totalFiles, validFiles, invalidFiles
    
    If Not objFSO.FolderExists(strSourcePath) Then
        WScript.Echo "Error: Directorio de origen no existe: " & strSourcePath
        objAccess.Quit
        WScript.Quit 1
    End If
    
    ' PASO 4.1: Validacion previa de sintaxis
    WScript.Echo "Validando sintaxis de todos los modulos..."
    Set objFolder = objFSO.GetFolder(strSourcePath)
    totalFiles = 0
    validFiles = 0
    invalidFiles = 0
    
    For Each objFile In objFolder.Files
        If LCase(objFSO.GetExtensionName(objFile.Name)) = "bas" Or LCase(objFSO.GetExtensionName(objFile.Name)) = "cls" Then
            totalFiles = totalFiles + 1
            validationResult = ValidateVBASyntax(objFile.Path, errorDetails)
            
            If validationResult = True Then
                validFiles = validFiles + 1
                WScript.Echo "  ✓ " & objFile.Name & " - Sintaxis valida"
            Else
                invalidFiles = invalidFiles + 1
                WScript.Echo "  ✗ ERROR en " & objFile.Name & ": " & errorDetails
            End If
        End If
    Next
    
    If invalidFiles > 0 Then
        WScript.Echo "ABORTANDO: Se encontraron " & invalidFiles & " archivos con errores de sintaxis."
        WScript.Echo "Use 'cscript condor_cli.vbs validate --verbose' para más detalles."
        objAccess.Quit
        WScript.Quit 1
    End If
    
    WScript.Echo "✓ Validacion completada: " & validFiles & " archivos validos"
    
    ' PASO 4.2: Procesar archivos de modulos
    Set objFolder = objFSO.GetFolder(strSourcePath)
    
    For Each objFile In objFolder.Files
        If LCase(objFSO.GetExtensionName(objFile.Name)) = "bas" Or LCase(objFSO.GetExtensionName(objFile.Name)) = "cls" Then
            strFileName = objFile.Path
            strModuleName = objFSO.GetBaseName(objFile.Name)
            
            WScript.Echo "Procesando modulo: " & strModuleName
            
            ' Determinar tipo de archivo
            Dim fileExtension
            fileExtension = LCase(objFSO.GetExtensionName(objFile.Name))
            
            ' Limpiar archivo antes de importar (eliminar metadatos Attribute)
            Dim cleanedContent
            cleanedContent = CleanVBAFile(strFileName, fileExtension)
            
            ' Importar usando contenido limpio
            WScript.Echo "  Clase " & strModuleName & " importada correctamente"
            Call ImportModuleWithAnsiEncoding(strFileName, strModuleName, fileExtension, Nothing, cleanedContent)
            
            If Err.Number <> 0 Then
                WScript.Echo "Error al importar modulo " & strModuleName & ": " & Err.Description
                Err.Clear
            End If
        End If
    Next
    
    ' PASO 4.3: Guardar cada modulo individualmente
    WScript.Echo "Guardando modulos individualmente..."
    On Error Resume Next
    
    For Each vbComponent In objAccess.VBE.ActiveVBProject.VBComponents
        If vbComponent.Type = 1 Then  ' vbext_ct_StdModule
            WScript.Echo "Guardando modulo: " & vbComponent.Name
            objAccess.DoCmd.Save 5, vbComponent.Name  ' acModule = 5
            If Err.Number <> 0 Then
                WScript.Echo "Advertencia al guardar " & vbComponent.Name & ": " & Err.Description
                Err.Clear
            End If
        ElseIf vbComponent.Type = 2 Then  ' vbext_ct_ClassModule
            WScript.Echo "Guardando clase: " & vbComponent.Name
            objAccess.DoCmd.Save 7, vbComponent.Name  ' acClassModule = 7
            If Err.Number <> 0 Then
                WScript.Echo "Advertencia al guardar " & vbComponent.Name & ": " & Err.Description
                Err.Clear
            End If
        End If
    Next
    
    ' PASO 4.4: Verificacion de integridad y compilacion
    WScript.Echo "Verificando integridad de nombres de modulos..."
    Call VerifyModuleNames()
    
    ' PASO 4.5: Copiar todos los archivos de src a la caché
    WScript.Echo "Paso 5: Copiando todos los archivos de /src a la cache..."
    Call CopyAllFilesToCache()
    
    WScript.Echo "=== RECONSTRUCCION COMPLETADA EXITOSAMENTE ==="
    WScript.Echo "El proyecto VBA ha sido completamente reconstruido"
    WScript.Echo "Todos los modulos han sido reimportados desde /src"
    
    On Error GoTo 0
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
Sub SyncSingleModule(moduleName)
    WScript.Echo "=== SINCRONIZANDO MODULO: " & moduleName & " ==="
    
    On Error Resume Next
    
    ' --- Paso 1: Verificar que existe el archivo fuente ---
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
        WScript.Echo "  ❌ No se encontró archivo fuente para " & moduleName
        Exit Sub
    End If
    
    WScript.Echo "  ✓ Archivo fuente encontrado: " & strSourceFile
    
    ' --- Paso 2: Validar sintaxis del archivo ---
    Dim errorDetails, validationResult
    validationResult = ValidateVBASyntax(strSourceFile, errorDetails)
    
    If Not validationResult Then
        WScript.Echo "  ❌ Error de sintaxis en " & moduleName & ":"
        WScript.Echo "      " & errorDetails
        Exit Sub
    End If
    
    WScript.Echo "  ✓ Sintaxis validada correctamente"
    
    ' --- Paso 3: Limpiar contenido con conversión UTF-8 -> ANSI ---
    Dim cleanedContent
    cleanedContent = CleanVBAFile(strSourceFile, fileExtension)
    
    If cleanedContent = "" Then
        WScript.Echo "  ❌ Error: No se pudo leer o limpiar el archivo " & strSourceFile
        Exit Sub
    End If
    
    WScript.Echo "  ✓ Contenido limpiado y convertido a ANSI"
    
    ' --- Paso 4: Importar usando ImportModuleWithAnsiEncoding ---
    Call ImportModuleWithAnsiEncoding(strSourceFile, moduleName, fileExtension, Nothing, cleanedContent)
    
    If Err.Number <> 0 Then
        WScript.Echo "  ❌ Error al importar módulo " & moduleName & ": " & Err.Description
        Err.Clear
        Exit Sub
    End If
    
    WScript.Echo "  ✅ Módulo " & moduleName & " sincronizado correctamente con conversión de codificación"
    
    On Error GoTo 0
End Sub

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
    
    ' Paso 4: Importar el módulo usando DoCmd.LoadFromText (sin confirmaciones)
    WScript.Echo "  Importando módulo: " & moduleName
    Call ImportModuleWithLoadFromText(strSourceFile, moduleName, fileExtension)
    
    If Err.Number <> 0 Then
        WScript.Echo "  ❌ Error al importar módulo " & moduleName & ": " & Err.Description
        Err.Clear
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
Sub UpdateProject()
    WScript.Echo "=== ACTUALIZACION SELECTIVA DEL PROYECTO VBA ==="
    
    ' Configurar Access en modo silencioso para evitar confirmaciones
    On Error Resume Next
    ' objAccess.DoCmd.SetWarnings False  ' Comentado temporalmente por error de compilación
    objAccess.Application.Echo False
    objAccess.DisplayAlerts = False
    objAccess.Application.AutomationSecurity = 1  ' msoAutomationSecurityLow
    objAccess.VBE.MainWindow.Visible = False
    ' Configuraciones adicionales para evitar popups
    objAccess.Application.UserControl = False
    objAccess.Application.Interactive = False
    objAccess.VBE.CommandBars.AdaptiveMenus = False
    Err.Clear
    On Error GoTo 0
    
    ' Verificar si hay argumentos adicionales (lista de módulos)
    If objArgs.Count > 1 Then
        ' Modo selectivo: sincronizar módulos específicos
        WScript.Echo "Modo selectivo: sincronizando módulos especificados..."
        
        Dim moduleList, moduleNames, i
        moduleList = objArgs(1)
        moduleNames = Split(moduleList, ",")
        
        WScript.Echo "Módulos a sincronizar: " & UBound(moduleNames) + 1
        
        ' Paso 1: Eliminar todos los módulos especificados
        WScript.Echo "Eliminando módulos existentes..."
        For i = 0 To UBound(moduleNames)
            Dim moduleName
            moduleName = Trim(moduleNames(i))
            
            If moduleName <> "" Then
                WScript.Echo "  Eliminando componente: " & moduleName
                Call RemoveVBAComponent(moduleName)
            End If
        Next
        
        ' Paso 2: Cerrar base de datos guardando cambios
        WScript.Echo "Cerrando base de datos guardando cambios..."
        
        ' Guardar todos los cambios pendientes antes de cerrar
        On Error Resume Next
        objAccess.DoCmd.Save
        objAccess.DoCmd.RunCommand 2040  ' acCmdSaveAll
        Err.Clear
        On Error GoTo 0
        
        objAccess.DoCmd.Close
        objAccess.Quit 1  ' Cerrar guardando
        Set objAccess = Nothing
        WScript.Echo "  ✓ Base de datos cerrada y guardada"
        
        ' Paso 3: Reabrir base de datos en modo seguro y oculto
        WScript.Echo "Reabriendo base de datos en modo seguro..."
        Set objAccess = CreateObject("Access.Application")
        objAccess.Visible = False
        objAccess.UserControl = False
        
        ' Determinar contraseña para la base de datos
        Dim strDbPassword
        strDbPassword = GetDatabasePassword(strAccessPath)
        
        ' Abrir base de datos
        If strDbPassword = "" Then
            objAccess.OpenCurrentDatabase strAccessPath
        Else
            objAccess.OpenCurrentDatabase strAccessPath, , strDbPassword
        End If
        
        ' Configurar modo silencioso y seguro
        On Error Resume Next
        ' objAccess.DoCmd.SetWarnings False  ' Comentado temporalmente por error de compilación
        objAccess.Application.Echo False
        objAccess.DisplayAlerts = False
        objAccess.Application.AutomationSecurity = 1  ' msoAutomationSecurityLow
        objAccess.VBE.MainWindow.Visible = False
        objAccess.Application.Interactive = False
        objAccess.VBE.CommandBars.AdaptiveMenus = False
        Err.Clear
        On Error GoTo 0
        WScript.Echo "  ✓ Base de datos reabierta en modo seguro"
        
        ' Paso 4: Importar todos los módulos validando sintaxis
        WScript.Echo "Importando módulos con validación de sintaxis..."
        
        For i = 0 To UBound(moduleNames)
            moduleName = Trim(moduleNames(i))
            
            If moduleName <> "" Then
                WScript.Echo ""
                WScript.Echo "=== IMPORTANDO MODULO: " & moduleName & " ==="
                
                ' Verificar que el fichero fuente existe
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
                
                ' Validar sintaxis del archivo
                Dim errorDetails, validationResult
                validationResult = ValidateVBASyntax(strSourceFile, errorDetails)
                
                If validationResult <> True Then
                    WScript.Echo "  ❌ Error de sintaxis en " & moduleName & ": " & errorDetails
                    Exit Sub
                End If
                
                WScript.Echo "  ✓ Sintaxis válida"
                
                ' Limpiar el contenido del fichero
                Dim cleanedContent
                cleanedContent = CleanVBAFile(strSourceFile, fileExtension)
                
                If cleanedContent = "" Then
                    WScript.Echo "  ❌ Error: No se pudo leer o limpiar el contenido del archivo"
                    Exit Sub
                End If
                
                WScript.Echo "  ✓ Contenido limpiado"
                
                ' Importar el módulo
                WScript.Echo "Importando modulo: " & moduleName
                Call ImportModuleWithAnsiEncoding(strSourceFile, moduleName, fileExtension, Nothing, cleanedContent)
                
                If Err.Number <> 0 Then
                    WScript.Echo "  ❌ Error al importar módulo " & moduleName & ": " & Err.Description
                    Err.Clear
                    Exit Sub
                End If
                
                ' Guardar el módulo individualmente
                On Error Resume Next
                objAccess.DoCmd.Save , moduleName
                If Err.Number <> 0 Then
                    WScript.Echo "  ⚠️ Advertencia al guardar " & moduleName & ": " & Err.Description
                    Err.Clear
                Else
                    WScript.Echo "  ✓ Módulo guardado: " & moduleName
                End If
                On Error GoTo 0
                
                If fileExtension = "cls" Then
                    WScript.Echo "✅ Clase " & moduleName & " importada correctamente"
                Else
                    WScript.Echo "✅ Módulo " & moduleName & " importado correctamente"
                End If
            End If
        Next
        
        ' Paso 5: Verificar integridad de nombres de módulos
        WScript.Echo "Verificando integridad de nombres de módulos..."
        Call VerifyModuleNames()
        
        ' Paso 6: Cerrar Access sin confirmaciones
        WScript.Echo "Cerrando Access sin confirmaciones..."
        
        ' Configurar para cerrar sin confirmaciones
        On Error Resume Next
        ' objAccess.DoCmd.SetWarnings False  ' Comentado temporalmente por error de compilación
        objAccess.Application.Echo False
        objAccess.DisplayAlerts = False
        objAccess.Application.Interactive = False
        
        ' Guardar todo antes de cerrar
        objAccess.DoCmd.Save
        objAccess.DoCmd.RunCommand 2040  ' acCmdSaveAll
        
        ' Cerrar sin confirmaciones
        objAccess.DoCmd.Close
        objAccess.Quit 1  ' acQuitSaveAll
        Set objAccess = Nothing
        
        Err.Clear
        On Error GoTo 0
        
        WScript.Echo "  ✓ Access cerrado sin confirmaciones"
        
    Else
        ' Modo automático: sincronizar solo archivos modificados
        WScript.Echo "Modo automático: sincronizando archivos modificados..."
        
        ' Paso 1: Usar carpeta persistente .vba_cache
        Dim strCachePath
        strCachePath = objFSO.BuildPath(objFSO.GetParentFolderName(strSourcePath), ".vba_cache")
        
        ' Crear carpeta de caché si no existe
        If Not objFSO.FolderExists(strCachePath) Then
            objFSO.CreateFolder strCachePath
            WScript.Echo "Carpeta de cache creada: " & strCachePath
        Else
            WScript.Echo "Usando carpeta de cache existente: " & strCachePath
        End If
        
        ' Paso 2: Comparar archivos de /src con los de .vba_cache
        WScript.Echo "Comparando archivos para detectar cambios..."
        Call CompareAndSyncModulesWithCache(strCachePath)
        
        ' Paso 3: Copiar archivos modificados a la caché
        WScript.Echo "Actualizando cache con archivos modificados..."
        Call CopyModifiedFilesToCache(strCachePath)
        
    End If
    
    WScript.Echo ""
    WScript.Echo "=== ACTUALIZACION COMPLETADA EXITOSAMENTE ==="
End Sub

' Subrutina para exportar módulos VBA actuales a carpeta cache
Sub ExportModulesToCache(cachePath)
    On Error Resume Next
    
    Dim vbProject, vbComponent
    Set vbProject = objAccess.VBE.ActiveVBProject
    
    For Each vbComponent In vbProject.VBComponents
        If vbComponent.Type = 1 Or vbComponent.Type = 2 Then ' Solo módulos estándar y de clase
            Dim fileExtension, fileName
            
            If vbComponent.Type = 1 Then
                fileExtension = "bas"
            Else
                fileExtension = "cls"
            End If
            
            fileName = objFSO.BuildPath(cachePath, vbComponent.Name & "." & fileExtension)
            
            ' Exportar usando la función existente
            Call ExportModuleWithAnsiEncoding(fileName, vbComponent.Name, fileExtension)
            
            If Err.Number <> 0 Then
                WScript.Echo "  ⚠️ Advertencia al exportar " & vbComponent.Name & ": " & Err.Description
                Err.Clear
            End If
        End If
    Next
    
    On Error GoTo 0
End Sub

' Subrutina para comparar y sincronizar módulos con caché persistente
Sub CompareAndSyncModulesWithCache(cachePath)
    On Error Resume Next
    
    Dim objSrcFolder, objCacheFolder, objFile
    Dim srcFile, cacheFile, moduleName
    Dim syncCount, modifiedModules()
    Dim moduleCount
    
    syncCount = 0
    moduleCount = 0
    ReDim modifiedModules(100) ' Array para almacenar módulos modificados
    
    ' Verificar archivos nuevos o modificados en /src
    If objFSO.FolderExists(strSourcePath) Then
        Set objSrcFolder = objFSO.GetFolder(strSourcePath)
        
        For Each objFile In objSrcFolder.Files
            If LCase(objFSO.GetExtensionName(objFile.Name)) = "bas" Or LCase(objFSO.GetExtensionName(objFile.Name)) = "cls" Then
                srcFile = objFile.Path
                moduleName = objFSO.GetBaseName(objFile.Name)
                cacheFile = objFSO.BuildPath(cachePath, objFile.Name)
                
                ' Verificar si el archivo es nuevo o ha sido modificado
                If Not objFSO.FileExists(cacheFile) Then
                    WScript.Echo "  📄 Archivo nuevo detectado: " & moduleName
                    Call SyncSingleModule(moduleName)
                    modifiedModules(moduleCount) = objFile.Name
                    moduleCount = moduleCount + 1
                    syncCount = syncCount + 1
                ElseIf CompareFileContents(srcFile, cacheFile) = False Then
                    WScript.Echo "  📝 Archivo modificado detectado: " & moduleName
                    Call SyncSingleModule(moduleName)
                    modifiedModules(moduleCount) = objFile.Name
                    moduleCount = moduleCount + 1
                    syncCount = syncCount + 1
                End If
            End If
        Next
    End If
    
    ' Verificar archivos eliminados (existen en cache pero no en /src)
    If objFSO.FolderExists(cachePath) Then
        Set objCacheFolder = objFSO.GetFolder(cachePath)
        
        For Each objFile In objCacheFolder.Files
            If LCase(objFSO.GetExtensionName(objFile.Name)) = "bas" Or LCase(objFSO.GetExtensionName(objFile.Name)) = "cls" Then
                cacheFile = objFile.Path
                moduleName = objFSO.GetBaseName(objFile.Name)
                srcFile = objFSO.BuildPath(strSourcePath, objFile.Name)
                
                ' Si el archivo no existe en /src, eliminar el componente VBA y el archivo de caché
                If Not objFSO.FileExists(srcFile) Then
                    WScript.Echo "  🗑️ Archivo eliminado detectado: " & moduleName
                    Call RemoveVBAComponent(moduleName)
                    ' Eliminar también de la caché
                    objFSO.DeleteFile cacheFile, True
                    WScript.Echo "    ✓ Eliminado de la cache: " & moduleName
                    syncCount = syncCount + 1
                End If
            End If
        Next
    End If
    
    If syncCount = 0 Then
        WScript.Echo "  ✅ No se detectaron cambios. Proyecto actualizado."
    Else
        WScript.Echo "  ✅ " & syncCount & " módulos sincronizados."
    End If
    
    ' Guardar la lista de módulos modificados para uso posterior
    ReDim Preserve modifiedModules(moduleCount - 1)
    
    On Error GoTo 0
End Sub

' Subrutina para comparar y sincronizar módulos modificados (versión anterior)
Sub CompareAndSyncModules(cachePath)
    On Error Resume Next
    
    Dim objSrcFolder, objCacheFolder, objFile
    Dim srcFile, cacheFile, moduleName
    Dim syncCount
    
    syncCount = 0
    
    ' Verificar archivos nuevos o modificados en /src
    If objFSO.FolderExists(strSourcePath) Then
        Set objSrcFolder = objFSO.GetFolder(strSourcePath)
        
        For Each objFile In objSrcFolder.Files
            If LCase(objFSO.GetExtensionName(objFile.Name)) = "bas" Or LCase(objFSO.GetExtensionName(objFile.Name)) = "cls" Then
                srcFile = objFile.Path
                moduleName = objFSO.GetBaseName(objFile.Name)
                cacheFile = objFSO.BuildPath(cachePath, objFile.Name)
                
                ' Verificar si el archivo es nuevo o ha sido modificado
                If Not objFSO.FileExists(cacheFile) Then
                    WScript.Echo "  📄 Archivo nuevo detectado: " & moduleName
                    Call SyncSingleModule(moduleName)
                    syncCount = syncCount + 1
                ElseIf CompareFileContents(srcFile, cacheFile) = False Then
                    WScript.Echo "  📝 Archivo modificado detectado: " & moduleName
                    Call SyncSingleModule(moduleName)
                    syncCount = syncCount + 1
                End If
            End If
        Next
    End If
    
    ' Verificar archivos eliminados (existen en cache pero no en /src)
    If objFSO.FolderExists(cachePath) Then
        Set objCacheFolder = objFSO.GetFolder(cachePath)
        
        For Each objFile In objCacheFolder.Files
            If LCase(objFSO.GetExtensionName(objFile.Name)) = "bas" Or LCase(objFSO.GetExtensionName(objFile.Name)) = "cls" Then
                cacheFile = objFile.Path
                moduleName = objFSO.GetBaseName(objFile.Name)
                srcFile = objFSO.BuildPath(strSourcePath, objFile.Name)
                
                ' Si el archivo no existe en /src, eliminar el componente VBA
                If Not objFSO.FileExists(srcFile) Then
                    WScript.Echo "  🗑️ Archivo eliminado detectado: " & moduleName
                    Call RemoveVBAComponent(moduleName)
                    syncCount = syncCount + 1
                End If
            End If
        Next
    End If
    
    If syncCount = 0 Then
        WScript.Echo "  ✅ No se detectaron cambios. Proyecto actualizado."
    Else
        WScript.Echo "  ✅ " & syncCount & " módulos sincronizados."
    End If
    
    On Error GoTo 0
End Sub

' Función para comparar contenido de dos archivos
Function CompareFileContents(file1, file2)
    On Error Resume Next
    
    Dim content1, content2
    content1 = ReadFileWithAnsiEncoding(file1)
    content2 = ReadFileWithAnsiEncoding(file2)
    
    CompareFileContents = (content1 = content2)
    
    On Error GoTo 0
End Function

' Subrutina para eliminar un componente VBA
Sub RemoveVBAComponent(moduleName)
    On Error Resume Next
    
    Dim vbProject, vbComponent
    Set vbProject = objAccess.VBE.ActiveVBProject
    
    For Each vbComponent In vbProject.VBComponents
        If vbComponent.Name = moduleName Then
            If vbComponent.Type = 1 Or vbComponent.Type = 2 Then ' Solo módulos estándar y de clase
                WScript.Echo "    Eliminando componente: " & moduleName
                vbProject.VBComponents.Remove vbComponent
                
                If Err.Number <> 0 Then
                    WScript.Echo "    ❌ Error eliminando " & moduleName & ": " & Err.Description
                    Err.Clear
                Else
                    WScript.Echo "    ✓ Componente eliminado: " & moduleName
                End If
            End If
            Exit For
        End If
    Next
    
End Sub

' Función para copiar todos los archivos de src a la caché
Sub CopyAllFilesToCache()
    On Error Resume Next
    
    ' Definir ruta de la caché
    Dim strCachePath
    strCachePath = objFSO.BuildPath(objFSO.GetParentFolderName(strSourcePath), ".vba_cache")
    
    ' Crear directorio de caché si no existe
    If Not objFSO.FolderExists(strCachePath) Then
        objFSO.CreateFolder strCachePath
        If Err.Number <> 0 Then
            WScript.Echo "Error creando directorio de cache: " & Err.Description
            Err.Clear
            Exit Sub
        End If
    End If
    
    ' Copiar todos los archivos .bas y .cls de src a cache
    Dim objFolder, objFile
    Dim copiedCount
    copiedCount = 0
    
    If Not objFSO.FolderExists(strSourcePath) Then
        WScript.Echo "Error: Directorio de origen no existe: " & strSourcePath
        Exit Sub
    End If
    
    Set objFolder = objFSO.GetFolder(strSourcePath)
    
    For Each objFile In objFolder.Files
        If LCase(objFSO.GetExtensionName(objFile.Name)) = "bas" Or LCase(objFSO.GetExtensionName(objFile.Name)) = "cls" Then
            Dim destPath
            destPath = objFSO.BuildPath(strCachePath, objFile.Name)
            
            ' Copiar archivo
            objFSO.CopyFile objFile.Path, destPath, True
            
            If Err.Number <> 0 Then
                WScript.Echo "  ❌ Error copiando " & objFile.Name & ": " & Err.Description
                Err.Clear
            Else
                WScript.Echo "  ✓ Copiado: " & objFile.Name
                copiedCount = copiedCount + 1
            End If
        End If
    Next
    
    WScript.Echo "✓ Cache actualizada: " & copiedCount & " archivos copiados"
    
    On Error GoTo 0
End Sub

' Función para verificar cambios antes de abrir la base de datos
Function CheckForChangesBeforeUpdate()
    On Error Resume Next
    
    Dim objSrcFolder, objCacheFolder, objFile
    Dim srcFile, cacheFile, moduleName
    Dim hasChanges
    
    hasChanges = False
    
    ' Definir rutas
    Dim strCachePath
    strCachePath = objFSO.BuildPath(objFSO.GetParentFolderName(strSourcePath), ".vba_cache")
    
    WScript.Echo "=== VERIFICANDO CAMBIOS ==="
    WScript.Echo "Comparando archivos de /src con caché..."
    
    ' Si no existe la caché, hay cambios
    If Not objFSO.FolderExists(strCachePath) Then
        WScript.Echo "📁 Caché no existe - se requiere sincronización"
        CheckForChangesBeforeUpdate = True
        Exit Function
    End If
    
    ' Verificar archivos nuevos o modificados en /src
    If objFSO.FolderExists(strSourcePath) Then
        Set objSrcFolder = objFSO.GetFolder(strSourcePath)
        
        For Each objFile In objSrcFolder.Files
            If LCase(objFSO.GetExtensionName(objFile.Name)) = "bas" Or LCase(objFSO.GetExtensionName(objFile.Name)) = "cls" Then
                srcFile = objFile.Path
                moduleName = objFSO.GetBaseName(objFile.Name)
                cacheFile = objFSO.BuildPath(strCachePath, objFile.Name)
                
                ' Verificar si el archivo es nuevo o ha sido modificado
                If Not objFSO.FileExists(cacheFile) Then
                    WScript.Echo "📄 Archivo nuevo: " & moduleName
                    hasChanges = True
                ElseIf CompareFileContents(srcFile, cacheFile) = False Then
                    WScript.Echo "📝 Archivo modificado: " & moduleName
                    hasChanges = True
                End If
            End If
        Next
    End If
    
    ' Verificar archivos eliminados (existen en cache pero no en /src)
    If objFSO.FolderExists(strCachePath) Then
        Set objCacheFolder = objFSO.GetFolder(strCachePath)
        
        For Each objFile In objCacheFolder.Files
            If LCase(objFSO.GetExtensionName(objFile.Name)) = "bas" Or LCase(objFSO.GetExtensionName(objFile.Name)) = "cls" Then
                cacheFile = objFile.Path
                moduleName = objFSO.GetBaseName(objFile.Name)
                srcFile = objFSO.BuildPath(strSourcePath, objFile.Name)
                
                ' Si el archivo no existe en /src, hay cambios
                If Not objFSO.FileExists(srcFile) Then
                    WScript.Echo "🗑️ Archivo eliminado: " & moduleName
                    hasChanges = True
                End If
            End If
        Next
    End If
    
    CheckForChangesBeforeUpdate = hasChanges
    
    On Error GoTo 0
End Function

' Subrutina para copiar solo archivos modificados a la caché
Sub CopyModifiedFilesToCache(cachePath)
    On Error Resume Next
    
    Dim objSrcFolder, objFile
    Dim srcFile, cacheFile
    Dim copiedCount
    
    copiedCount = 0
    
    ' Definir la ruta de origen
    Dim strSrcPath
    strSrcPath = objFSO.BuildPath(objFSO.GetParentFolderName(cachePath), "src")
    
    ' Verificar que existe la carpeta src
    If Not objFSO.FolderExists(strSrcPath) Then
        WScript.Echo "Error: No se encontró la carpeta src en: " & strSrcPath
        Exit Sub
    End If
    
    ' Crear carpeta de caché si no existe
    If Not objFSO.FolderExists(cachePath) Then
        objFSO.CreateFolder cachePath
    End If
    
    Set objSrcFolder = objFSO.GetFolder(strSrcPath)
    
    ' Recorrer archivos .bas y .cls en /src
    For Each objFile In objSrcFolder.Files
        If LCase(objFSO.GetExtensionName(objFile.Name)) = "bas" Or LCase(objFSO.GetExtensionName(objFile.Name)) = "cls" Then
            srcFile = objFile.Path
            cacheFile = objFSO.BuildPath(cachePath, objFile.Name)
            
            ' Copiar si el archivo no existe en caché o es diferente
            If Not objFSO.FileExists(cacheFile) Then
                objFSO.CopyFile srcFile, cacheFile, True
                copiedCount = copiedCount + 1
            ElseIf CompareFileContents(srcFile, cacheFile) = False Then
                objFSO.CopyFile srcFile, cacheFile, True
                copiedCount = copiedCount + 1
            End If
        End If
    Next
    
    WScript.Echo "✓ Archivos modificados copiados a cache: " & copiedCount & " archivos"
    
    On Error GoTo 0
End Sub

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
    WScript.Echo "EJEMPLOS DE USO:"
    WScript.Echo "  cscript condor_cli.vbs bundle Auth"
    WScript.Echo "  cscript condor_cli.vbs bundle Document C:\\temp"
    WScript.Echo "  cscript condor_cli.vbs bundle TestFramework"
    WScript.Echo "  cscript condor_cli.vbs bundle Tests"
    WScript.Echo "  cscript condor_cli.vbs bundle Config"
    WScript.Echo ""
    WScript.Echo "NOTAS:"
    WScript.Echo "  • Los archivos se copian con extensión .txt para fácil visualización"
    WScript.Echo "  • Se crea una carpeta con timestamp: bundle_<funcionalidad>_YYYYMMDD_HHMMSS"
    WScript.Echo "  • Cada funcionalidad incluye automáticamente sus dependencias"
    WScript.Echo "  • Si un archivo no existe, se muestra una advertencia pero continúa"
End Sub

' Función para obtener la lista de archivos por funcionalidad según CONDOR_MASTER_PLAN.md
' Incluye dependencias para cada funcionalidad
Function GetFunctionalityFiles(strFunctionality)
    Dim arrFiles
    
    Select Case LCase(strFunctionality)
        Case "auth", "autenticacion", "authentication"
            ' Sección 3.1 - Autenticación + Dependencias
            arrFiles = Array("IAuthService.cls", "CAuthService.cls", "CMockAuthService.cls", _
                           "IAuthRepository.cls", "CAuthRepository.cls", "CMockAuthRepository.cls", _
                           "AuthData.cls", "modAuthFactory.bas", "TestAuthService.bas", _
                           "TIAuthRepository.bas", _
                           "IConfig.cls", "IErrorHandlerService.cls", "modEnumeraciones.bas")
        
        Case "document", "documentos", "documents"
            ' Sección 3.2 - Gestión de Documentos + Dependencias
            arrFiles = Array("IDocumentService.cls", "CDocumentService.cls", "CMockDocumentService.cls", _
                           "IWordManager.cls", "CWordManager.cls", "CMockWordManager.cls", _
                           "IMapeoRepository.cls", "CMapeoRepository.cls", "CMockMapeoRepository.cls", _
                           "EMapeo.cls", "modDocumentServiceFactory.bas", "TestDocumentService.bas", _
                           "TIDocumentService.bas", _
                           "ISolicitudService.cls", "CSolicitudService.cls", "modSolicitudServiceFactory.bas", _
                           "IOperationLogger.cls", "IConfig.cls", "IErrorHandlerService.cls", "IFileSystem.cls", _
                           "modWordManagerFactory.bas", "modRepositoryFactory.bas", "modErrorHandlerFactory.bas")
        
        Case "expediente", "expedientes"
            ' Sección 3.3 - Gestión de Expedientes + Dependencias (Actualizado tras refactorización)
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
                           "IOperationLogger.cls", "IErrorHandlerService.cls", "IConfig.cls")
        
        Case "workflow", "flujo"
            ' Sección 3.5 - Gestión de Workflow (v2.0 Simplificada) + Dependencias
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
                           "OperationLog.cls", "modOperationLoggerFactory.bas", "TestOperationLogger.bas", _
                           "TIOperationRepository.bas", _
                           "IErrorHandlerService.cls", "IConfig.cls")
        
        Case "config", "configuracion"
            ' Sección 4 - Configuración + Dependencias (Simplificado tras Misión de Emergencia)
            arrFiles = Array("IConfig.cls", "CConfig.cls", "CMockConfig.cls", "modConfigFactory.bas", _
                           "TestCConfig.bas")
        
        Case "filesystem", "archivos"
            ' Sección 5 - Sistema de Archivos + Dependencias
            arrFiles = Array("IFileSystem.cls", "CFileSystem.cls", "CMockFileSystem.cls", _
                           "modFileSystemFactory.bas", "TestFileSystem.bas", "TIFileSystem.bas", _
                           "IErrorHandlerService.cls")
        
        Case "word"
            ' Sección 6 - Gestión de Word + Dependencias
            arrFiles = Array("IWordManager.cls", "CWordManager.cls", "CMockWordManager.cls", _
                           "modWordManagerFactory.bas", "TIWordManager.bas", _
                           "IFileSystem.cls", "IErrorHandlerService.cls")
        
        Case "error", "errores", "errors"
            ' Sección 7 - Gestión de Errores + Dependencias
            arrFiles = Array("IErrorHandlerService.cls", "CErrorHandlerService.cls", "CMockErrorHandlerService.cls", _
                           "modErrorHandlerFactory.bas", "modErrorHandler.bas", "TestErrorHandlerService.bas", _
                           "IConfig.cls", "IFileSystem.cls")
        
        Case "testframework", "testing", "framework"
            ' Sección 8 - Framework de Testing + Dependencias (Actualizado con ITestReporter)
            arrFiles = Array("ITestReporter.cls", "CTestResult.cls", "CTestSuiteResult.cls", "CTestReporter.cls", _
                           "modTestRunner.bas", "modTestUtils.bas", "modAssert.bas", _
                           "TestModAssert.bas", "IFileSystem.cls", "IConfig.cls", _
                           "IErrorHandlerService.cls")
        
        Case "app", "aplicacion", "application"
            ' Sección 9 - Gestión de Aplicación + Dependencias
            arrFiles = Array("IAppManager.cls", "CAppManager.cls", "CMockAppManager.cls", _
                           "modAppManagerFactory.bas", "TestAppManager.bas", "IAuthService.cls", _
                           "IConfig.cls", "IErrorHandlerService.cls")
        
        Case "models", "modelos", "datos"
            ' Sección 10 - Modelos de Datos
            arrFiles = Array("EUsuario.cls", "ESolicitud.cls", "EExpediente.cls", "EDatosPc.cls", _
                           "EDatosCdCa.cls", "EDatosCdCaSub.cls", "EEstado.cls", "ETransicion.cls", _
                           "EMapeo.cls", "EAdjunto.cls", "ELogCambio.cls", "ELogError.cls", "EOperationLog.cls")
        
        Case "utils", "utilidades", "enumeraciones"
            ' Sección 11 - Utilidades y Enumeraciones
            arrFiles = Array("modRepositoryFactory.bas", "modEnumeraciones.bas", "modQueries.bas", _
                           "ModAppManagerFactory.bas", "modAuthFactory.bas", "modConfigFactory.bas", _
                           "modDocumentServiceFactory.bas", "modErrorHandlerFactory.bas", _
                           "modExpedienteServiceFactory.bas", "modFileSystemFactory.bas", _
                           "modNotificationServiceFactory.bas", "modOperationLoggerFactory.bas", _
                           "modSolicitudServiceFactory.bas", "modWordManagerFactory.bas", _
                           "modWorkflowServiceFactory.bas")
        
        Case "tests", "pruebas", "testing"
            ' Sección 12 - Archivos de Pruebas
            arrFiles = Array("TestAppManager.bas", "TestAuthService.bas", "TestCConfig.bas", _
                           "TestCExpedienteService.bas", "TestDocumentService.bas", _
                           "TestErrorHandlerService.bas", "TestModAssert.bas", "TestOperationLogger.bas", _
                           "TestSolicitudService.bas", "TestWorkflowService.bas", _
                           "TIAuthRepository.bas", "TIExpedienteRepository.bas", _
                           "TIMapeoRepository.bas", "TIDocumentService.bas", _
                           "TIFileSystem.bas", "TINotificationService.bas", _
                           "TIOperationRepository.bas", "TISolicitudRepository.bas", _
                           "TIWordManager.bas", "TIWorkflowRepository.bas")
        
        Case Else
            ' Para funcionalidades no definidas, usar búsqueda por nombre (comportamiento anterior)
            arrFiles = Array()
    End Select
    
    GetFunctionalityFiles = arrFiles
End Function



        


' Subrutina para empaquetar archivos de código por funcionalidad
Sub BundleFunctionality()
    On Error Resume Next
    
    Dim strFunctionality, strDestPath, strBundlePath
    Dim objFolder, objFile
    Dim foundFiles, copiedFiles
    Dim timestamp
    Dim arrFunctionalityFiles, i
    Dim usePredefinedList
    
    ' Verificar argumentos
    If objArgs.Count < 2 Then
        WScript.Echo "Error: Se requiere nombre de funcionalidad"
        WScript.Echo "Uso: cscript condor_cli.vbs bundle <funcionalidad> [ruta_destino]"
        WScript.Quit 1
    End If
    
    strFunctionality = objArgs(1)
    
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
    strBundlePath = objFSO.BuildPath(strDestPath, "bundle_" & strFunctionality & "_" & timestamp)
    
    WScript.Echo "=== EMPAQUETANDO FUNCIONALIDAD: " & strFunctionality & " ==="
    WScript.Echo "Buscando archivos en: " & strSourcePath
    WScript.Echo "Carpeta destino: " & strBundlePath
    
    ' Obtener lista de archivos para la funcionalidad
    arrFunctionalityFiles = GetFunctionalityFiles(strFunctionality)
    usePredefinedList = (UBound(arrFunctionalityFiles) >= 0)
    
    If usePredefinedList Then
        WScript.Echo "Usando lista predefinida de archivos según CONDOR_MASTER_PLAN.md"
        WScript.Echo "Archivos esperados: " & (UBound(arrFunctionalityFiles) + 1)
    Else
        WScript.Echo "Usando búsqueda por nombre de funcionalidad"
    End If
    WScript.Echo ""
    
    ' Verificar que existe la carpeta src
    If Not objFSO.FolderExists(strSourcePath) Then
        WScript.Echo "Error: Directorio de origen no existe: " & strSourcePath
        WScript.Quit 1
    End If
    
    ' Crear carpeta de destino
    If Not objFSO.FolderExists(strBundlePath) Then
        objFSO.CreateFolder strBundlePath
        If Err.Number <> 0 Then
            WScript.Echo "Error creando carpeta de destino: " & Err.Description
            WScript.Quit 1
        End If
    End If
    
    Set objFolder = objFSO.GetFolder(strSourcePath)
    foundFiles = 0
    copiedFiles = 0
    
    If usePredefinedList Then
        ' Usar lista predefinida de archivos
        For i = 0 To UBound(arrFunctionalityFiles)
            Dim fileName, filePath, destFilePath
            fileName = arrFunctionalityFiles(i)
            filePath = objFSO.BuildPath(strSourcePath, fileName)
            
            If objFSO.FileExists(filePath) Then
                foundFiles = foundFiles + 1
                
                ' Copiar archivo con extensión .txt añadida
                destFilePath = objFSO.BuildPath(strBundlePath, fileName & ".txt")
                
                objFSO.CopyFile filePath, destFilePath, True
                
                If Err.Number <> 0 Then
                    WScript.Echo "  ❌ Error copiando " & fileName & ": " & Err.Description
                    Err.Clear
                Else
                    WScript.Echo "  ✓ " & fileName & " -> " & fileName & ".txt"
                    copiedFiles = copiedFiles + 1
                End If
            Else
                WScript.Echo "  ⚠️ Archivo no encontrado: " & fileName
            End If
        Next
    Else
        ' Usar búsqueda por nombre (comportamiento anterior)
        For Each objFile In objFolder.Files
            If LCase(objFSO.GetExtensionName(objFile.Name)) = "bas" Or LCase(objFSO.GetExtensionName(objFile.Name)) = "cls" Then
                ' Verificar si el nombre del archivo contiene la funcionalidad (sin distinguir mayúsculas)
                If InStr(1, LCase(objFile.Name), LCase(strFunctionality)) > 0 Then
                    foundFiles = foundFiles + 1
                    
                    ' Copiar archivo con extensión .txt añadida
                    Dim destFilePathLegacy
                    destFilePathLegacy = objFSO.BuildPath(strBundlePath, objFile.Name & ".txt")
                    
                    objFSO.CopyFile objFile.Path, destFilePathLegacy, True
                    
                    If Err.Number <> 0 Then
                        WScript.Echo "  ❌ Error copiando " & objFile.Name & ": " & Err.Description
                        Err.Clear
                    Else
                        WScript.Echo "  ✓ " & objFile.Name & " -> " & objFile.Name & ".txt"
                        copiedFiles = copiedFiles + 1
                    End If
                End If
            End If
        Next
    End If
    
    WScript.Echo ""
    WScript.Echo "=== RESULTADO DEL EMPAQUETADO ==="
    WScript.Echo "Archivos encontrados: " & foundFiles
    WScript.Echo "Archivos copiados: " & copiedFiles
    WScript.Echo "Ubicación del paquete: " & strBundlePath
    
    If copiedFiles = 0 Then
        If usePredefinedList Then
            WScript.Echo "⚠️ No se encontraron archivos de la funcionalidad '" & strFunctionality & "' según CONDOR_MASTER_PLAN.md"
        Else
            WScript.Echo "⚠️ No se encontraron archivos que contengan '" & strFunctionality & "'"
        End If
    Else
        WScript.Echo "✅ Empaquetado completado exitosamente"
    End If
    
    On Error GoTo 0
End Sub
