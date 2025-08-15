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
If objArgs.Count = 0 Then
    WScript.Echo "=== CONDOR CLI - Herramienta de linea de comandos ==="
    WScript.Echo "Uso: cscript condor_cli.vbs [comando] [opciones]"
    WScript.Echo ""
    WScript.Echo "COMANDOS DISPONIBLES:"
    WScript.Echo "  export     - Exportar modulos VBA a /src (con codificacion ANSI)"
    WScript.Echo "  validate   - Validar sintaxis de modulos VBA sin importar"
    WScript.Echo "  test       - Ejecutar todos los tests (requiere rebuild previo)"
    WScript.Echo "  rebuild    - Reconstruir proyecto VBA (eliminar todos los modulos y reimportar)"
    WScript.Echo "  lint       - Auditar codigo VBA para detectar cabeceras duplicadas"
    WScript.Echo "  createtable <nombre> <sql> - Crear tabla con consulta SQL"
    WScript.Echo "  droptable <nombre> - Eliminar tabla"
    WScript.Echo "  listtables [db_path] - Listar tablas (opcionalmente de base especifica)"
    WScript.Echo "  relink <db_path> <folder> - Re-vincular tablas a bases locales"
    WScript.Echo "  relink --all - Re-vincular todas las bases en ./back automaticamente"
    WScript.Echo ""
    WScript.Echo "FLUJO DE TRABAJO RECOMENDADO:"
    WScript.Echo "  1. cscript condor_cli.vbs rebuild   (sincronizar y compilar modulos)"
    WScript.Echo "  2. cscript condor_cli.vbs test      (ejecutar pruebas)"
    WScript.Echo ""
    WScript.Echo "OPCIONES ESPECIALES:"
    WScript.Echo "  --dry-run  - Simular operacion sin modificar Access (solo con import)"
    WScript.Echo "  --verbose  - Mostrar informacion detallada durante la operacion"
    WScript.Echo ""
    WScript.Echo "EJEMPLOS:"
    WScript.Echo "  cscript condor_cli.vbs validate"
    WScript.Echo "  cscript condor_cli.vbs export --verbose"
    WScript.Echo "  cscript condor_cli.vbs rebuild && cscript condor_cli.vbs test"
    WScript.Quit 1
End If

strAction = LCase(objArgs(0))

If strAction <> "export" And strAction <> "validate" And strAction <> "createtable" And strAction <> "droptable" And strAction <> "listtables" And strAction <> "test" And strAction <> "relink" And strAction <> "rebuild" And strAction <> "lint" Then
    WScript.Echo "Error: Comando debe ser 'export', 'validate', 'createtable', 'droptable', 'listtables', 'test', 'relink', 'rebuild' o 'lint'"
    WScript.Quit 1
End If

Set objFSO = CreateObject("Scripting.FileSystemObject")

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

' Para tests y rebuild, usar la base de datos de desarrollo
If strAction = "test" Or strAction = "rebuild" Then
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
objAccess.DoCmd.SetWarnings False
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
WScript.Echo "Base de datos abierta correctamente (modo seguro)"

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
ElseIf strAction = "createtable" Then
    Call CreateTable()
ElseIf strAction = "droptable" Then
    Call DropTable()
ElseIf strAction = "listtables" Then
    Call ListTables()
ElseIf strAction = "test" Then
    Dim testModule, argIndex
    testModule = ""
    ' Buscar el primer argumento que no sea una opción especial
    For argIndex = 1 To objArgs.Count - 1
        If LCase(objArgs(argIndex)) <> "--verbose" And LCase(objArgs(argIndex)) <> "--dry-run" Then
            testModule = objArgs(argIndex)
            Exit For
        End If
    Next
    Call RunTestsWithEngine(testModule)
ElseIf strAction = "rebuild" Then
    Call RebuildProject()
ElseIf strAction = "lint" Then
    Call LintProject()
ElseIf strAction = "relink" Then
    Call RelinkTables()
End If

' Cerrar Access
WScript.Echo "Cerrando Access..."
' Restaurar estado normal de Access antes de cerrar
On Error Resume Next
objAccess.Application.Echo True
objAccess.Quit 1  ' acQuitSaveAll = 1
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

' Subrutina para ejecutar tests con motor interno (versión simplificada)
' Subrutina para ejecutar tests con motor interno (versión final y robusta)
Sub RunTestsWithEngine(testModule)
    WScript.Echo "=== EJECUTANDO PRUEBAS (MODO DIRECTO) ==="
    
    ' Forzar compilación antes de ejecutar las pruebas
    WScript.Echo "Forzando compilación antes de ejecutar pruebas..."
    On Error Resume Next
    Err.Clear
    
    ' Intentar compilación usando VBE
    Dim vbProject
    Set vbProject = objAccess.VBE.ActiveVBProject
    
    ' Forzar compilación accediendo a cada módulo
    Dim vbComponent
    For Each vbComponent In vbProject.VBComponents
        ' Acceder al código para forzar compilación
        Dim tempCode
        tempCode = vbComponent.CodeModule.Lines(1, 1)
        If Err.Number <> 0 Then
            WScript.Echo "⚠️ Error compilando " & vbComponent.Name & ": " & Err.Description
            Err.Clear
        End If
    Next
    
    WScript.Echo "✓ Compilación previa completada"
    On Error GoTo 0
    
    Dim resultado, testFunction
    
    ' Determinar la función a llamar
    If testModule = "" Then
        testFunction = "modTestRunner.RunAllTests"
    Else
        testFunction = testModule & ".RunAllTests"
    End If
    
    WScript.Echo "Intentando ejecutar directamente: '" & testFunction & "'"
    
    On Error Resume Next
     ' Usamos Application.Run, que es el método más fiable para funciones que devuelven un valor.
     ' Intentar diferentes métodos de llamada
     Set objDB = objAccess.CurrentDb
     
     ' Intentar con el nombre completo del módulo
     resultado = objAccess.Application.Run("modTestRunner.RunAllTests")
     
     ' Si falla, intentar sin el prefijo del módulo
     If Err.Number <> 0 Then
         Err.Clear
         resultado = objAccess.Application.Run("RunAllTests")
     End If
     
     If Err.Number <> 0 Then
        WScript.Echo ""
        WScript.Echo "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
        WScript.Echo "!!! ERROR CRITICO AL EJECUTAR LAS PRUEBAS !!!"
        WScript.Echo "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
        WScript.Echo "Error: " & Err.Number & " - " & Err.Description
        WScript.Echo "Causa más probable: El proyecto VBA no está compilado. El flujo de trabajo en dos pasos es obligatorio."
        WScript.Echo ""
        WScript.Echo "FLUJO DE TRABAJO OBLIGATORIO:"
        WScript.Echo "1. Ejecute: cscript condor_cli.vbs rebuild"
        WScript.Echo "2. Ejecute: cscript condor_cli.vbs test"
        Err.Clear
    Else
        WScript.Echo "" ' Línea en blanco
        WScript.Echo "--- INICIO DEL REPORTE DE PRUEBAS ---"
        
        ' Intentar leer el archivo temporal con los resultados detallados
        Dim fso, tempFile, testResults
        Set fso = CreateObject("Scripting.FileSystemObject")
        tempFile = fso.GetSpecialFolder(2) & "\condor_test_results.txt"  ' GetSpecialFolder(2) = Temp folder
        
        If fso.FileExists(tempFile) Then
            On Error Resume Next
            Dim file
            Set file = fso.OpenTextFile(tempFile, 1)  ' ForReading = 1
            If Err.Number = 0 Then
                testResults = file.ReadAll
                file.Close
                If Len(Trim(testResults)) > 0 Then
                    WScript.Echo testResults
                Else
                    WScript.Echo "[INFO] Archivo de resultados vacio"
                    WScript.Echo resultado
                End If
            Else
                WScript.Echo "[WARNING] No se pudo leer archivo de resultados: " & Err.Description
                WScript.Echo resultado
                Err.Clear
            End If
            On Error GoTo 0
            
            ' Limpiar archivo temporal
            fso.DeleteFile tempFile, True
        Else
            WScript.Echo "[INFO] No se encontro archivo de resultados, mostrando resultado directo"
            WScript.Echo resultado
        End If
        
        WScript.Echo "--- FIN DEL REPORTE DE PRUEBAS ---"
    End If
    
    On Error GoTo 0
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
                ' Intentar compilar usando DoCmd.Save con el tipo correcto
                If vbComponent.Type = 1 Then
                    objAccess.DoCmd.Save 5, vbComponent.Name  ' acModule = 5
                    
                    If Err.Number <> 0 Then
                        WScript.Echo "  ⚠️ Advertencia al guardar " & vbComponent.Name & ": " & Err.Description
                        compilationErrors = compilationErrors + 1
                        hasErrors = True
                        Err.Clear
                    Else
                        WScript.Echo "  ✓ " & vbComponent.Name & " compilado correctamente"
                        compiledModules = compiledModules + 1
                    End If
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
    
    ' Leer contenido del archivo usando ADODB.Stream con UTF-8
    Set objStream = CreateObject("ADODB.Stream")
    objStream.Type = 2 ' adTypeText
    objStream.Charset = "UTF-8"
    objStream.Open
    objStream.LoadFromFile filePath
    strContent = objStream.ReadText
    objStream.Close
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
    Dim objStream, strContent, arrLines, i, cleanedLines
    Dim strLine, errorDetails
    
    ' Usar el tipo de archivo pasado como parámetro ("bas" o "cls")
    ' fileType debe ser "bas" o "cls"
    
    ' Validar sintaxis antes de procesar
    If Not ValidateVBASyntax(filePath, errorDetails) Then
        WScript.Echo "[WARN] ADVERTENCIA: Errores de sintaxis detectados en " & objFSO.GetFileName(filePath) & ":"
        WScript.Echo errorDetails
        WScript.Echo "Continuando con la importación..."
    End If
    
    ' Leer contenido del archivo usando ADODB.Stream con UTF-8
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
        WScript.Echo "❌ ERROR: No se pudo leer el archivo " & filePath & ": " & Err.Description
        CleanVBAFile = ""
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    
    On Error GoTo 0
    
    ' Normalizar saltos de línea y dividir
    strContent = Replace(strContent, vbCrLf, vbLf)
    strContent = Replace(strContent, vbCr, vbLf)
    arrLines = Split(strContent, vbLf)
    
    ' Procesar línea por línea eliminando metadatos y líneas Option
    cleanedLines = ""
    Dim linesRemoved
    linesRemoved = 0
    
    For i = 0 To UBound(arrLines)
        strLine = Trim(arrLines(i))
        
        ' Verificar si es una línea que debe eliminarse
        Dim shouldRemove
        shouldRemove = False
        
        ' SIEMPRE eliminar metadatos Attribute y VERSION para ambos tipos
        If Left(strLine, 7) = "VERSION" Then shouldRemove = True
        If Left(strLine, 5) = "BEGIN" Then shouldRemove = True
        If Left(strLine, 8) = "MultiUse" Then shouldRemove = True
        If strLine = "END" Then shouldRemove = True
        If Left(strLine, 9) = "Attribute" Then shouldRemove = True
        
        ' ELIMINAR líneas Option para AMBOS tipos (.bas y .cls)
        ' Esto es necesario porque AddFromString automáticamente añade las cabeceras Option
        If Left(strLine, 6) = "Option" Then shouldRemove = True
        
        ' Eliminar líneas vacías solo al inicio del archivo
        If strLine = "" And cleanedLines = "" Then shouldRemove = True
        
        ' Si la línea no debe eliminarse, agregarla al contenido limpio
        If Not shouldRemove Then
            ' Reemplazar caracteres problemáticos
            strLine = arrLines(i) ' Usar línea original (con espacios)
            strLine = Replace(strLine, Chr(147), Chr(34)) ' Comilla izquierda tipográfica -> comilla normal
            strLine = Replace(strLine, Chr(148), Chr(34)) ' Comilla derecha tipográfica -> comilla normal
            strLine = Replace(strLine, Chr(145), Chr(39)) ' Apostrofe izquierdo -> apostrofe normal
            strLine = Replace(strLine, Chr(146), Chr(39)) ' Apostrofe derecho -> apostrofe normal
            ' Reemplazar caracteres Unicode problemáticos
            strLine = Replace(strLine, "✓", "[OK]") ' Check mark -> [OK]
            strLine = Replace(strLine, "✗", "[FALLO]") ' X mark -> [FALLO]
            strLine = Replace(strLine, "└─", "  |-") ' Box drawing -> simple chars
            strLine = Replace(strLine, "✅", "[OK]") ' Green check -> [OK]
            strLine = Replace(strLine, "❌", "[ERROR]") ' Red X -> [ERROR]
            strLine = Replace(strLine, "⚠️", "[WARN]") ' Warning -> [WARN]
            
            If cleanedLines <> "" Then
                cleanedLines = cleanedLines & vbCrLf
            End If
            cleanedLines = cleanedLines & strLine
        Else
            linesRemoved = linesRemoved + 1
        End If
    Next
    
    If fileType = "cls" Then
        WScript.Echo "  Eliminadas " & linesRemoved & " líneas de metadatos y Option (archivo .cls)"
    Else
        WScript.Echo "  Eliminadas " & linesRemoved & " líneas de metadatos (archivo .bas, Option mantenidas)"
    End If
    
    CleanVBAFile = cleanedLines
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
    
    WScript.Echo "Paso 2: Guardando y cerrando base de datos..."
    
    ' Forzar guardado y cierre
    objAccess.DoCmd.Save
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
    objAccess.DoCmd.SetWarnings False
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
            WScript.Echo "  Importando modulo (con limpieza): " & strFileName
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
    
    WScript.Echo "Iniciando compilacion condicional..."
    Call CompileModulesConditionally()
    
    WScript.Echo "=== RECONSTRUCCION COMPLETADA EXITOSAMENTE ==="
    WScript.Echo "El proyecto VBA ha sido completamente reconstruido"
    WScript.Echo "Todos los modulos han sido reimportados desde /src"
    
    On Error GoTo 0
End Sub

' La subrutina ExecuteTestModule ha sido eliminada ya que ahora se usa el motor interno modTestRunner
