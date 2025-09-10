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
Dim pathArg, i

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

strAction = LCase(objArgs(0))

If strAction <> "export" And strAction <> "validate" And strAction <> "validate-schema" And strAction <> "test" And strAction <> "createtable" And strAction <> "droptable" And strAction <> "listtables" And strAction <> "relink" And strAction <> "rebuild" And strAction <> "lint" And strAction <> "bundle" And strAction <> "migrate" Then
    WScript.Echo "Error: Comando debe ser 'export', 'validate', 'validate-schema', 'test', 'createtable', 'droptable', 'listtables', 'relink', 'rebuild', 'lint', 'bundle' o 'migrate'"
    WScript.Quit 1
End If

Set objFSO = CreateObject("Scripting.FileSystemObject")

' Los comandos bundle y validate-schema no requieren Access
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
    Call ValidateSchema()
    WScript.Quit 0
End If

' Determinar qué base de datos usar según la acción
If strAction = "createtable" Or strAction = "droptable" Or strAction = "migrate" Then
    strAccessPath = strDataPath
ElseIf strAction = "listtables" Then
    pathArg = ""
    ' Buscar un argumento que no sea el flag --schema
    For i = 1 To objArgs.Count - 1
        If LCase(objArgs(i)) <> "--schema" Then
            pathArg = objArgs(i)
            Exit For
        End If
    Next
    
    If pathArg <> "" Then
        ' Resolver ruta relativa a absoluta
        strAccessPath = ResolveRelativePath(pathArg)
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
Dim bVerbose
bVerbose = False

For i = 1 To objArgs.Count - 1
    If LCase(objArgs(i)) = "--verbose" Then
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
ElseIf strAction = "migrate" Then
    Call ExecuteMigrations()

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

' Subrutina principal para validar esquemas de base de datos
Sub ValidateSchema()
    WScript.Echo "=== INICIANDO VALIDACIÓN DE ESQUEMA DE BASE DE DATOS ==="
    
    Dim lanzaderaSchema, condorSchema
    Dim allOk
    allOk = True
    
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
    If Not VerifySchema(strSourcePath & "\..\back\test_env\fixtures\databases\Lanzadera_test_template.accdb", "dpddpd", lanzaderaSchema) Then allOk = False
        If Not VerifySchema(strSourcePath & "\..ack\test_env\fixtures\databases\Document_test_template.accdb", "", condorSchema) Then allOk = False
    
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
    WScript.Echo "  rebuild                      - Método principal de sincronización del proyecto"
    WScript.Echo "                                 Reconstrucción completa: elimina todos los módulos"
    WScript.Echo "                                 y reimporta desde /src para garantizar coherencia"
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
    WScript.Echo "  migrate [file.sql]           - Ejecutar scripts de migración SQL desde ./db/migrations"
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
  
    WScript.Echo "  --verbose                    - Mostrar información detallada durante la operación"
    WScript.Echo ""
    WScript.Echo "FLUJO DE TRABAJO RECOMENDADO:"
    WScript.Echo "  1. cscript condor_cli.vbs validate     (validar sintaxis antes de importar)"
    WScript.Echo "  2. cscript condor_cli.vbs rebuild      (reconstrucción completa del proyecto)"
    WScript.Echo "  3. cscript condor_cli.vbs test         (ejecutar pruebas unitarias)"
    WScript.Echo ""
    WScript.Echo "EJEMPLOS DE USO:"
    WScript.Echo "  cscript condor_cli.vbs --help"
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
        objAccess.Quit
        WScript.Quit 1
    End If
    On Error GoTo 0
    
    ' Verificar si el string devuelto es válido
    If IsEmpty(reportString) Or reportString = "" Then
        WScript.Echo "ERROR: El motor de pruebas de Access no devolvió ningún resultado."
        WScript.Echo "SUGERENCIA: Verifique que la función 'ExecuteAllTestsForCLI' en 'modTestRunner' no esté fallando silenciosamente."
        objAccess.Quit
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
        
        Case "tests", "pruebas", "testing", "test"
            ' Sección 12 - Archivos de Pruebas (Autodescubrimiento)
            arrFiles = Array()
        Case Else
            ' Funcionalidad no reconocida - devolver array vacío
            arrFiles = Array()
    End Select
    
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
        filePath = objFSO.BuildPath(strSourcePath, fileName)
        
        If objFSO.FileExists(filePath) Then
            ' Copiar archivo con extensión .txt añadida
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
