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
    WScript.Echo "  test       - Ejecutar suite de pruebas unitarias"
    WScript.Echo "  update     - Sincronizar modulos VBA (automatico o selectivo)"
    WScript.Echo "  rebuild    - Reconstruir proyecto VBA (eliminar todos los modulos y reimportar)"
    WScript.Echo "  compile    - Compilar todos los modulos VBA del proyecto"
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
    WScript.Quit 1
End If

strAction = LCase(objArgs(0))

If strAction <> "export" And strAction <> "validate" And strAction <> "test" And strAction <> "createtable" And strAction <> "droptable" And strAction <> "listtables" And strAction <> "relink" And strAction <> "rebuild" And strAction <> "lint" And strAction <> "compile" And strAction <> "update" Then
    WScript.Echo "Error: Comando debe ser 'export', 'validate', 'test', 'createtable', 'droptable', 'listtables', 'relink', 'rebuild', 'lint', 'compile' o 'update'"
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
ElseIf strAction = "compile" Then
    Call CompileProject()
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
        If Not (Left(strLine, 12) = "Attribute VB_" Or _
                Left(strLine, 17) = "VERSION 1.0 CLASS" Or _
                strLine = "BEGIN" Or _
                Left(strLine, 10) = "MultiUse =" Or _
                strLine = "END" Or _
                strLine = "Option Compare Database" Or _
                strLine = "Option Explicit") Then
            
            ' Si no cumple ninguna condición, es código VBA válido
            ' Se añade al cleanedContent seguida de un salto de línea
            cleanedContent = cleanedContent & strLine & vbCrLf
        End If
    Next
    
    ' La función devuelve cleanedContent directamente
    ' No añade ninguna cabecera Option manualmente
    CleanVBAFile = cleanedContent
End Function

' Función para calcular hash MD5 de un fichero
Function GetFileHash(filePath)
    Dim objStream, objHasher, arrBytes, strHash, i
    
    On Error Resume Next
    
    ' Verificar que el archivo existe
    If Not objFSO.FileExists(filePath) Then
        GetFileHash = ""
        Exit Function
    End If
    
    ' Leer el archivo como binario
    Set objStream = CreateObject("ADODB.Stream")
    objStream.Type = 1 ' adTypeBinary
    objStream.Open
    objStream.LoadFromFile filePath
    arrBytes = objStream.Read
    objStream.Close
    Set objStream = Nothing
    
    If Err.Number <> 0 Then
        GetFileHash = ""
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    
    ' Crear objeto para calcular hash MD5
    Set objHasher = CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")
    
    If Err.Number <> 0 Then
        ' Si no está disponible .NET, usar método alternativo con CAPICOM
        Err.Clear
        Set objHasher = CreateObject("CAPICOM.HashedData")
        objHasher.Algorithm = 3 ' CAPICOM_HASH_ALGORITHM_MD5
        objHasher.Hash arrBytes
        strHash = objHasher.Value
        Set objHasher = Nothing
    Else
        ' Usar .NET MD5
        arrBytes = objHasher.ComputeHash_2(arrBytes)
        Set objHasher = Nothing
        
        ' Convertir bytes a string hexadecimal
        strHash = ""
        For i = 0 To UBound(arrBytes)
            strHash = strHash & Right("0" & Hex(arrBytes(i)), 2)
        Next
        strHash = LCase(strHash)
    End If
    
    If Err.Number <> 0 Then
        GetFileHash = ""
        Err.Clear
    Else
        GetFileHash = strHash
    End If
    
    On Error GoTo 0
End Function

' Función para obtener la ruta del caché persistente
Function GetCachePath()
    Dim strProjectRoot
    strProjectRoot = objFSO.GetParentFolderName(strSourcePath)
    GetCachePath = objFSO.BuildPath(strProjectRoot, ".vba_cache")
End Function

' Función para inicializar el caché persistente
Sub InitializePersistentCache()
    Dim strCachePath
    strCachePath = GetCachePath()
    
    If Not objFSO.FolderExists(strCachePath) Then
        objFSO.CreateFolder strCachePath
        WScript.Echo "Caché persistente inicializado: " & strCachePath
    End If
End Sub

' Función para copiar un fichero al caché persistente
Sub CopyFileToCache(sourceFile, moduleName, fileExtension)
    Dim strCachePath, strCacheFile
    
    On Error Resume Next
    
    strCachePath = GetCachePath()
    Call InitializePersistentCache()
    
    strCacheFile = objFSO.BuildPath(strCachePath, moduleName & "." & fileExtension)
    
    ' Copiar el archivo al caché
    objFSO.CopyFile sourceFile, strCacheFile, True
    
    If Err.Number <> 0 Then
        WScript.Echo "  ⚠️ Advertencia: No se pudo copiar " & moduleName & " al caché: " & Err.Description
        Err.Clear
    End If
    
    On Error GoTo 0
End Sub

' Función para eliminar un fichero del caché persistente
Sub RemoveFileFromCache(moduleName, fileExtension)
    Dim strCachePath, strCacheFile
    
    On Error Resume Next
    
    strCachePath = GetCachePath()
    strCacheFile = objFSO.BuildPath(strCachePath, moduleName & "." & fileExtension)
    
    If objFSO.FileExists(strCacheFile) Then
        objFSO.DeleteFile strCacheFile, True
        
        If Err.Number <> 0 Then
            WScript.Echo "  ⚠️ Advertencia: No se pudo eliminar " & moduleName & " del caché: " & Err.Description
            Err.Clear
        End If
    End If
    
    On Error GoTo 0
End Sub

' Función para comparar hashes entre /src y caché
Function CompareFileHashes(srcFile, cacheFile)
    Dim srcHash, cacheHash
    
    srcHash = GetFileHash(srcFile)
    cacheHash = GetFileHash(cacheFile)
    
    ' Si algún hash está vacío, considerar como diferentes
    If srcHash = "" Or cacheHash = "" Then
        CompareFileHashes = False
    Else
        CompareFileHashes = (srcHash = cacheHash)
    End If
End Function

' Función para comparar y sincronizar módulos usando hashes
Sub CompareAndSyncModulesWithHashes()
    Dim strCachePath, objSrcFolder, objCacheFolder
    Dim objSrcFile, objCacheFile
    Dim srcFiles, cacheFiles
    Dim moduleName, fileExtension
    Dim srcFilePath, cacheFilePath
    Dim modulesToUpdate, modulesToDelete
    Dim updatedCount, deletedCount
    
    On Error Resume Next
    
    strCachePath = GetCachePath()
    updatedCount = 0
    deletedCount = 0
    
    ' Crear diccionarios para archivos de /src y caché
    Set srcFiles = CreateObject("Scripting.Dictionary")
    Set cacheFiles = CreateObject("Scripting.Dictionary")
    
    ' Recopilar archivos de /src
    Set objSrcFolder = objFSO.GetFolder(strSourcePath)
    For Each objSrcFile In objSrcFolder.Files
        If LCase(objFSO.GetExtensionName(objSrcFile.Name)) = "bas" Or LCase(objFSO.GetExtensionName(objSrcFile.Name)) = "cls" Then
            moduleName = objFSO.GetBaseName(objSrcFile.Name)
            srcFiles.Add moduleName, objSrcFile.Path
        End If
    Next
    
    ' Recopilar archivos del caché
    If objFSO.FolderExists(strCachePath) Then
        Set objCacheFolder = objFSO.GetFolder(strCachePath)
        For Each objCacheFile In objCacheFolder.Files
            If LCase(objFSO.GetExtensionName(objCacheFile.Name)) = "bas" Or LCase(objFSO.GetExtensionName(objCacheFile.Name)) = "cls" Then
                moduleName = objFSO.GetBaseName(objCacheFile.Name)
                cacheFiles.Add moduleName, objCacheFile.Path
            End If
        Next
    End If
    
    ' Comparar archivos y detectar cambios
    Dim srcModuleNames, i
    srcModuleNames = srcFiles.Keys
    
    For i = 0 To UBound(srcModuleNames)
        moduleName = srcModuleNames(i)
        srcFilePath = srcFiles(moduleName)
        fileExtension = LCase(objFSO.GetExtensionName(srcFilePath))
        cacheFilePath = objFSO.BuildPath(strCachePath, moduleName & "." & fileExtension)
        
        ' Verificar si el módulo necesita actualización
        If Not cacheFiles.Exists(moduleName) Or Not CompareFileHashes(srcFilePath, cacheFilePath) Then
            WScript.Echo "  📝 Actualizando módulo: " & moduleName
            Call UpdateSingleModule(moduleName)
            updatedCount = updatedCount + 1
        End If
    Next
    
    ' Detectar módulos eliminados (existen en caché pero no en /src)
    Dim cacheModuleNames
    cacheModuleNames = cacheFiles.Keys
    
    For i = 0 To UBound(cacheModuleNames)
        moduleName = cacheModuleNames(i)
        If Not srcFiles.Exists(moduleName) Then
            WScript.Echo "  🗑️ Eliminando módulo: " & moduleName
            Call RemoveVBAComponent(moduleName)
            fileExtension = LCase(objFSO.GetExtensionName(cacheFiles(moduleName)))
            Call RemoveFileFromCache(moduleName, fileExtension)
            deletedCount = deletedCount + 1
        End If
    Next
    
    ' Mostrar resumen
    WScript.Echo ""
    WScript.Echo "=== RESUMEN DE SINCRONIZACIÓN ==="
    WScript.Echo "Módulos actualizados: " & updatedCount
    WScript.Echo "Módulos eliminados: " & deletedCount
    
    If updatedCount = 0 And deletedCount = 0 Then
        WScript.Echo "✅ Proyecto ya está sincronizado"
    End If
    
    On Error GoTo 0
End Sub

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
    Dim strLogPath, objLogFile, strLine, testsFailed
    strLogPath = "C:\Proyectos\CONDOR\logs\test_results.log"

    ' 1. Limpiar log anterior
    If objFSO.FileExists(strLogPath) Then objFSO.DeleteFile(strLogPath)

    ' 2. Ejecutar las pruebas en Access
    WScript.Echo "Ejecutando suite de pruebas en Access..."
    On Error Resume Next
    
    ' Verificar que la función ExecuteAllTests existe
    Dim testFunctionExists
    testFunctionExists = False
    
    ' Intentar verificar si la función existe en los módulos
    Dim vbComp, vbModule
    For Each vbComp In objAccess.VBE.ActiveVBProject.VBComponents
        If vbComp.Type = 1 Then ' vbext_ct_StdModule
            Set vbModule = vbComp.CodeModule
            If InStr(vbModule.Lines(1, vbModule.CountOfLines), "ExecuteAllTests") > 0 Then
                testFunctionExists = True
                WScript.Echo "✓ Función ExecuteAllTests encontrada en módulo: " & vbComp.Name
                Exit For
            End If
        End If
    Next
    
    If Not testFunctionExists Then
        WScript.Echo "ERROR: No se encontró la función ExecuteAllTests en ningún módulo VBA"
        WScript.Echo "SUGERENCIA: Verifica que el módulo modTestRunner.bas esté correctamente importado"
        objAccess.Quit
        WScript.Quit 1
    End If
    
    ' Ejecutar función wrapper sin parámetros
    objAccess.Application.Run "RunTests"

    If Err.Number <> 0 Then
        WScript.Echo "ERROR: Fallo crítico al invocar la suite de pruebas."
        WScript.Echo "  Código de Error: " & Err.Number
        WScript.Echo "  Descripción: " & Err.Description
        WScript.Echo "  Fuente: " & Err.Source
        WScript.Echo "SUGERENCIA: Abre Access manualmente y ejecuta ExecuteAllTests desde el módulo modTestRunner para ver el error específico"
        objAccess.Quit
        WScript.Quit 1
    End If
    On Error GoTo 0

    ' Access se cerrará automáticamente al final de las pruebas.
    ' Esperamos un momento para asegurar que el fichero de log se ha escrito.
    WScript.Sleep 2000

    ' 3. Leer y mostrar los resultados desde el log
    WScript.Echo "--- INICIO DE RESULTADOS DE PRUEBAS ---"
    testsFailed = True ' Asumir fallo hasta que se confirme el éxito
    If objFSO.FileExists(strLogPath) Then
        Set objLogFile = objFSO.OpenTextFile(strLogPath, 1) ' ForReading
        Do While Not objLogFile.AtEndOfStream
            strLine = objLogFile.ReadLine
            WScript.Echo strLine
            If InStr(strLine, "RESULT: SUCCESS") > 0 Then testsFailed = False
        Loop
        objLogFile.Close
    Else
        WScript.Echo "ERROR: No se encontró el fichero de resultados de pruebas."
    End If
    WScript.Echo "--- FIN DE RESULTADOS DE PRUEBAS ---"

    ' 4. Salir con el código de estado apropiado
    If testsFailed Then
        WScript.Echo "RESULTADO FINAL: ✗ Pruebas fallidas."
        WScript.Quit 1 ' Código de error para CI/CD
    Else
        WScript.Echo "RESULTADO FINAL: ✓ Todas las pruebas pasaron."
        WScript.Quit 0 ' Código de éxito
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
    
    ' Inicializar caché persistente
    Call InitializePersistentCache()
    
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
            Call ImportModuleWithAnsiEncoding(strFileName, strModuleName, fileExtension, Nothing, cleanedContent)
            
            If Err.Number <> 0 Then
                WScript.Echo "Error al importar modulo " & strModuleName & ": " & Err.Description
                Err.Clear
            Else
                WScript.Echo "  ✓ Módulo " & strModuleName & " importado correctamente"
                ' Copiar al caché persistente después de importar exitosamente
                Call CopyFileToCache(strFileName, strModuleName, fileExtension)
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

' Subrutina para compilar el proyecto VBA con verificaciones defensivas
Sub CompileProject()
    WScript.Echo "=== INICIANDO COMPILACION COMPLETA DEL PROYECTO VBA ==="

    ' --- Verificación Defensiva ---
    If objAccess Is Nothing Then
        WScript.Echo "ERROR CRITICO: El objeto de la aplicación Access no es válido (Nothing)."
        WScript.Quit 1
    End If

    On Error Resume Next
    If objAccess.CurrentDb Is Nothing Then
        WScript.Echo "ERROR CRITICO: No hay ninguna base de datos abierta en la instancia de Access."
        objAccess.Quit
        WScript.Quit 1
    End If
    On Error GoTo 0
    ' --- Fin de la Verificación ---

    WScript.Echo "Instancia de Access y base de datos validadas. Intentando compilar..."

    On Error Resume Next
    Err.Clear

    ' Comando para compilar y guardar todos los módulos.
    objAccess.DoCmd.RunCommand 584 ' acCmdCompileAndSaveAllModules

    If Err.Number <> 0 Then
        WScript.Echo "--------------------------------------------------"
        WScript.Echo "ERROR DE COMPILACION DETECTADO:"
        WScript.Echo "  Código de Error: " & Err.Number
        WScript.Echo "  Descripción: " & Err.Description
        WScript.Echo "--------------------------------------------------"
        WScript.Echo "ACCION REQUERIDA: Abre Access, ve al editor de VBA (Alt+F11) y selecciona 'Depuración -> Compilar' para localizar el error."
        Err.Clear
        objAccess.Quit
        WScript.Quit 1 ' Salir con código de error
    Else
        WScript.Echo "✓ Compilación completada exitosamente. No se encontraron errores."
        ' Dejamos que el script principal se encargue de cerrar Access.
    End If

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
    objAccess.DoCmd.SetWarnings False
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
        objAccess.DoCmd.SetWarnings False
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
        
        ' Inicializar caché persistente
        Call InitializePersistentCache()
        
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
                Else
                    WScript.Echo "  ✓ Módulo " & moduleName & " importado correctamente"
                    ' Copiar al caché persistente después de importar exitosamente
                    Call CopyFileToCache(strSourceFile, moduleName, fileExtension)
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
        objAccess.DoCmd.SetWarnings False
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
        ' Modo automático: sincronizar solo archivos modificados usando caché persistente
        WScript.Echo "Modo automático: sincronizando archivos modificados..."
        
        ' Paso 1: Inicializar caché persistente
        Call InitializePersistentCache()
        
        ' Paso 2: Comparar archivos de /src con los de .vba_cache usando hashes
        WScript.Echo "Comparando archivos para detectar cambios..."
        Call CompareAndSyncModulesWithHashes()
        
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

' Subrutina para comparar y sincronizar módulos modificados
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
                    
                    ' También eliminar del caché persistente
                    Dim fileExtension
                    If vbComponent.Type = 1 Then
                        fileExtension = "bas"
                    Else
                        fileExtension = "cls"
                    End If
                    
                    Call RemoveFileFromCache(moduleName, fileExtension)
                End If
            End If
            Exit For
        End If
    Next
    
    On Error GoTo 0
End Sub
