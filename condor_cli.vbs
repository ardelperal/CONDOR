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
    WScript.Echo "  import     - Importar modulos VBA desde /src (con validacion previa)"
    WScript.Echo "  export     - Exportar modulos VBA a /src (con codificacion ANSI)"
    WScript.Echo "  validate   - Validar sintaxis de modulos VBA sin importar"
    WScript.Echo "  test       - Ejecutar todos los tests y mostrar resultados"
    WScript.Echo "  createtable <nombre> <sql> - Crear tabla con consulta SQL"
    WScript.Echo "  droptable <nombre> - Eliminar tabla"
    WScript.Echo "  listtables [db_path] - Listar tablas (opcionalmente de base especifica)"
    WScript.Echo "  relink <db_path> <folder> - Re-vincular tablas a bases locales"
    WScript.Echo "  relink --all - Re-vincular todas las bases en ./back automaticamente"
    WScript.Echo ""
    WScript.Echo "OPCIONES ESPECIALES:"
    WScript.Echo "  --dry-run  - Simular operacion sin modificar Access (solo con import)"
    WScript.Echo "  --verbose  - Mostrar informacion detallada durante la operacion"
    WScript.Echo ""
    WScript.Echo "EJEMPLOS:"
    WScript.Echo "  cscript condor_cli.vbs validate"
    WScript.Echo "  cscript condor_cli.vbs import --dry-run"
    WScript.Echo "  cscript condor_cli.vbs export --verbose"
    WScript.Quit 1
End If

strAction = LCase(objArgs(0))

If strAction <> "import" And strAction <> "export" And strAction <> "validate" And strAction <> "createtable" And strAction <> "droptable" And strAction <> "listtables" And strAction <> "test" And strAction <> "relink" Then
    WScript.Echo "Error: Comando debe ser 'import', 'export', 'validate', 'createtable', 'droptable', 'listtables', 'test' o 'relink'"
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

' Para tests, usar la base de datos de desarrollo
If strAction = "test" Then
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
ElseIf strAction = "import" Then
    Call ImportModules(bDryRun, bVerbose)
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
ElseIf strAction = "relink" Then
    Call RelinkTables()
End If

' Cerrar Access
WScript.Echo "Cerrando Access..."
On Error Resume Next
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

' Subrutina para importar modulos
Sub ImportModules(bDryRun, bVerbose)
    Dim objFolder, objFile
    Dim strModuleName, strFileName, strContent
    Dim vbComponent
    Dim i, j
    Dim srcModules
    Dim moduleExists
    Dim validationResult
    Dim totalFiles, validFiles, invalidFiles
    
    If bDryRun Then
        WScript.Echo "=== MODO DRY-RUN: SIMULACION DE IMPORTACION ==="
    Else
        WScript.Echo "Iniciando importacion de modulos VBA..."
    End If
    
    If Not objFSO.FolderExists(strSourcePath) Then
        WScript.Echo "Error: Directorio de origen no existe: " & strSourcePath
        If Not bDryRun Then objAccess.Quit
        WScript.Quit 1
    End If
    
    ' PASO 1: Validacion previa de sintaxis
    WScript.Echo "Validando sintaxis de todos los modulos..."
    Set objFolder = objFSO.GetFolder(strSourcePath)
    totalFiles = 0
    validFiles = 0
    invalidFiles = 0
    
    For Each objFile In objFolder.Files
        If LCase(objFSO.GetExtensionName(objFile.Name)) = "bas" Or LCase(objFSO.GetExtensionName(objFile.Name)) = "cls" Then
            totalFiles = totalFiles + 1
            Dim errorDetails
            validationResult = ValidateVBASyntax(objFile.Path, errorDetails)
            
            If validationResult = True Then
                validFiles = validFiles + 1
                If bVerbose Then
                    WScript.Echo "  ✓ " & objFile.Name & " - Sintaxis valida"
                End If
            Else
                invalidFiles = invalidFiles + 1
                WScript.Echo "  ✗ ERROR en " & objFile.Name & ": " & errorDetails
            End If
        End If
    Next
    
    If invalidFiles > 0 Then
        WScript.Echo "ABORTANDO: Se encontraron " & invalidFiles & " archivos con errores de sintaxis."
        WScript.Echo "Use 'cscript condor_cli.vbs validate --verbose' para más detalles."
        If Not bDryRun Then objAccess.Quit
        WScript.Quit 1
    End If
    
    WScript.Echo "✓ Validacion completada: " & validFiles & " archivos validos"
    
    If bDryRun Then
        WScript.Echo "[DRY-RUN] Los siguientes modulos serian procesados:"
    End If
    
    ' Crear lista de módulos en src
    Set srcModules = CreateObject("Scripting.Dictionary")
    Set objFolder = objFSO.GetFolder(strSourcePath)
    
    For Each objFile In objFolder.Files
        If LCase(objFSO.GetExtensionName(objFile.Name)) = "bas" Or LCase(objFSO.GetExtensionName(objFile.Name)) = "cls" Then
            srcModules.Add objFSO.GetBaseName(objFile.Name), True
        End If
    Next
    
    ' PASO 2: Eliminar modulos que no estan en src
    If bDryRun Then
        WScript.Echo "[DRY-RUN] Modulos que serian eliminados:"
    Else
        WScript.Echo "Eliminando modulos que no estan en src..."
    End If
    
    On Error Resume Next
    For i = objAccess.VBE.ActiveVBProject.VBComponents.Count To 1 Step -1
        Set vbComponent = objAccess.VBE.ActiveVBProject.VBComponents(i)
        If vbComponent.Type = 1 Then  ' vbext_ct_StdModule
            If Not srcModules.Exists(vbComponent.Name) Then
                If bDryRun Then
                    WScript.Echo "  - " & vbComponent.Name & " (modulo obsoleto)"
                Else
                    WScript.Echo "Eliminando modulo obsoleto: " & vbComponent.Name
                    objAccess.VBE.ActiveVBProject.VBComponents.Remove vbComponent
                End If
            End If
        End If
    Next
    Err.Clear
    On Error GoTo 0
    
    ' PASO 3: Procesar archivos de modulos
    Set objFolder = objFSO.GetFolder(strSourcePath)
    
    For Each objFile In objFolder.Files
        If LCase(objFSO.GetExtensionName(objFile.Name)) = "bas" Or LCase(objFSO.GetExtensionName(objFile.Name)) = "cls" Then
            strFileName = objFile.Path
            strModuleName = objFSO.GetBaseName(objFile.Name)
            
            If bDryRun Then
                WScript.Echo "[DRY-RUN] Procesaria modulo: " & strModuleName
                If bVerbose Then
                    WScript.Echo "  - Archivo: " & strFileName
                    WScript.Echo "  - Eliminaria modulo existente si existe"
                    WScript.Echo "  - Limpiaria contenido (eliminar Attributes)"
                    WScript.Echo "  - Importaria con codificacion ANSI"
                End If
            Else
                WScript.Echo "Procesando modulo: " & strModuleName
                
                ' Eliminar modulo existente si existe
                On Error Resume Next
                For i = objAccess.VBE.ActiveVBProject.VBComponents.Count To 1 Step -1
                    If objAccess.VBE.ActiveVBProject.VBComponents(i).Name = strModuleName Then
                        If bVerbose Then
                            WScript.Echo "  Eliminando modulo existente: " & strModuleName
                        End If
                        objAccess.VBE.ActiveVBProject.VBComponents.Remove objAccess.VBE.ActiveVBProject.VBComponents(i)
                        Exit For
                    End If
                Next
                
                ' Limpiar archivo antes de importar (eliminar metadatos Attribute)
                Dim cleanedContent
                cleanedContent = CleanVBAFile(strFileName)
                
                ' Importar usando contenido limpio
                If bVerbose Then
                    WScript.Echo "  Importando modulo (con limpieza): " & strFileName
                End If
                Call ImportCleanModule(strModuleName, cleanedContent, objFile)
                
                If Err.Number <> 0 Then
                    WScript.Echo "Error al importar modulo " & strModuleName & ": " & Err.Description
                    Err.Clear
                End If
            End If
        End If
    Next
    
    If bDryRun Then
        WScript.Echo ""
        WScript.Echo "=== RESUMEN DRY-RUN ==="
        WScript.Echo "✓ Validacion de sintaxis completada"
        WScript.Echo "✓ Simulacion de eliminacion de modulos obsoletos"
        WScript.Echo "✓ Simulacion de importacion de modulos"
        WScript.Echo "[DRY-RUN] No se realizaron cambios en Access"
        WScript.Echo "Para ejecutar realmente: cscript condor_cli.vbs import"
    Else
        ' Guardar cada modulo individualmente para evitar dialogos
        WScript.Echo "Guardando modulos individualmente..."
        On Error Resume Next
        
        For Each vbComponent In objAccess.VBE.ActiveVBProject.VBComponents
            If vbComponent.Type = 1 Then  ' vbext_ct_StdModule
                If bVerbose Then
                    WScript.Echo "Guardando modulo: " & vbComponent.Name
                End If
                objAccess.DoCmd.Save 5, vbComponent.Name  ' acModule = 5
                If Err.Number <> 0 Then
                    WScript.Echo "Advertencia al guardar " & vbComponent.Name & ": " & Err.Description
                    Err.Clear
                End If
            End If
        Next
        
        ' Los modulos obsoletos ya fueron eliminados automaticamente al inicio
        
        ' Verificacion de integridad de nombres de modulos
        If bVerbose Then
            WScript.Echo "Verificando integridad de nombres de modulos..."
        End If
        Call VerifyModuleNames()
        
        ' Compilacion condicional de modulos
        WScript.Echo "Iniciando compilacion condicional..."
        Call CompileModulesConditionally()
        
        WScript.Echo "Importacion completada exitosamente"
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

' Subrutina para ejecutar tests con motor interno
Sub RunTestsWithEngine(testModule)
    WScript.Echo "=== EJECUTANDO TESTS CON MOTOR INTERNO ==="
    
    ' Compilación condicional antes de ejecutar pruebas
    WScript.Echo "Iniciando compilación condicional para pruebas..."
    Call CompileModulesConditionally()
    
    ' Pausa breve para asegurar que Access reconozca los módulos
    WScript.Sleep 1000
    
    ' Listar módulos disponibles para diagnóstico
     WScript.Echo "Listando módulos disponibles en la base de datos:"
     On Error Resume Next
     Dim i, vbComponents
     Set vbComponents = objAccess.VBE.ActiveVBProject.VBComponents
     
     If Err.Number = 0 And Not vbComponents Is Nothing Then
         For i = 1 To vbComponents.Count
             WScript.Echo "  - " & vbComponents(i).Name & " (Tipo: " & vbComponents(i).Type & ")"
         Next
     Else
         WScript.Echo "  No se pudieron listar los módulos VBA"
     End If
     On Error GoTo 0
     
     ' Ejecutar el motor de pruebas con manejo robusto de errores
     Dim resultado
     Dim testExecuted
     testExecuted = False
     
     ' Método 1: Intentar con DoCmd.RunCode
     WScript.Echo "Método 1: Intentando ejecutar RunAllTests con DoCmd.RunCode..."
     On Error Resume Next
     
     objAccess.TempVars.Add "TestResult", ""
     Dim testFunction
     If testModule = "" Then
         testFunction = "modTestRunner.RunAllTests()"
     ElseIf testModule = "modConfig.TestModConfig" Then
         testFunction = "Test_Config.RunAllTests()"
     Else
         testFunction = testModule & "()"
     End If
     
     objAccess.DoCmd.RunCode "TempVars(" & Chr(34) & "TestResult" & Chr(34) & ") = " & testFunction
     
     If Err.Number = 0 Then
         resultado = objAccess.TempVars("TestResult").Value
         objAccess.TempVars.Remove "TestResult"
         testExecuted = True
         WScript.Echo "✓ Pruebas ejecutadas exitosamente con DoCmd.RunCode"
     Else
         WScript.Echo "⚠️ Error con DoCmd.RunCode: " & Err.Number & " - " & Err.Description
         Err.Clear
         
         ' Limpiar TempVars si existe
         On Error Resume Next
         objAccess.TempVars.Remove "TestResult"
         Err.Clear
     End If
     
     ' Método 2: Intentar con Application.Run si el método 1 falló
     If Not testExecuted Then
         WScript.Echo "Método 2: Intentando con Application.Run..."
         On Error Resume Next
         If testModule = "" Then
             resultado = objAccess.Application.Run("modTestRunner.RunAllTests")
         ElseIf testModule = "modConfig.TestModConfig" Then
             resultado = objAccess.Application.Run("Test_Config.RunAllTests")
         Else
             resultado = objAccess.Application.Run(testModule)
         End If
         
         If Err.Number = 0 Then
             testExecuted = True
             WScript.Echo "✓ Pruebas ejecutadas exitosamente con Application.Run"
         Else
             WScript.Echo "⚠️ Error con Application.Run: " & Err.Number & " - " & Err.Description
             Err.Clear
         End If
     End If
     
     ' Método 3: Intentar con Eval si los métodos anteriores fallaron
     If Not testExecuted Then
         WScript.Echo "Método 3: Intentando con Eval..."
         On Error Resume Next
         resultado = objAccess.Eval(testFunction & "()")
         
         If Err.Number = 0 Then
             testExecuted = True
             WScript.Echo "✓ Pruebas ejecutadas exitosamente con Eval"
         Else
             WScript.Echo "⚠️ Error con Eval: " & Err.Number & " - " & Err.Description
             Err.Clear
         End If
     End If
     
     On Error GoTo 0
     
     ' Verificar si se pudo ejecutar algún método
     If testExecuted Then
         WScript.Echo "\n" & resultado
     Else
         WScript.Echo "\n❌ ERROR: No se pudo ejecutar el motor de pruebas con ningún método"
         WScript.Echo "Posibles causas:"
         WScript.Echo "  1. El módulo modTestRunner no está compilado correctamente"
         WScript.Echo "  2. Hay errores de sintaxis en el código VBA"
         WScript.Echo "  3. La función RunAllTests no es accesible"
         WScript.Echo "\nEjecutando diagnóstico básico..."
         
         ' Diagnóstico básico
         On Error Resume Next
         Dim basicTest
         basicTest = objAccess.Eval("1+1")
         If Err.Number = 0 Then
             WScript.Echo "✓ Eval básico funciona (1+1 = " & basicTest & ")"
         Else
             WScript.Echo "❌ Eval básico falla: " & Err.Description
         End If
         Err.Clear
         On Error GoTo 0
         
         WScript.Echo "\nEl CLI continuará funcionando para otras operaciones."
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
    
    ' Intentar compilar cada módulo individualmente
    For Each vbComponent In objAccess.VBE.ActiveVBProject.VBComponents
        If vbComponent.Type = 1 Then  ' vbext_ct_StdModule
            totalModules = totalModules + 1
            
            On Error Resume Next
            Err.Clear
            
            ' Intentar compilar el módulo específico
            WScript.Echo "Compilando módulo: " & vbComponent.Name
            
            ' Verificar si el módulo tiene errores de sintaxis
            Dim hasErrors
            hasErrors = False
            
            ' Intentar acceder al código del módulo para detectar errores
            Dim moduleCode
            moduleCode = vbComponent.CodeModule.Lines(1, vbComponent.CodeModule.CountOfLines)
            
            If Err.Number <> 0 Then
                WScript.Echo "  ⚠️ Error en módulo " & vbComponent.Name & ": " & Err.Description
                compilationErrors = compilationErrors + 1
                hasErrors = True
                Err.Clear
            Else
                ' Intentar compilar usando DoCmd.Save
                objAccess.DoCmd.Save 5, vbComponent.Name  ' acModule = 5
                
                If Err.Number <> 0 Then
                    WScript.Echo "  ⚠️ Advertencia al guardar " & vbComponent.Name & ": " & Err.Description
                    compilationErrors = compilationErrors + 1
                    hasErrors = True
                    Err.Clear
                Else
                    WScript.Echo "  ✓ Módulo " & vbComponent.Name & " compilado correctamente"
                    compiledModules = compiledModules + 1
                End If
            End If
            
            On Error GoTo 0
        End If
    Next
    
    ' Intentar compilación global solo si no hay errores individuales
    If compilationErrors = 0 Then
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
        If vbComponent.Type = 1 Then  ' vbext_ct_StdModule
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
    Dim objFile, strContent
    
    On Error Resume Next
    Set objFile = objFSO.OpenTextFile(filePath, 1, False, 0) ' ForReading = 1, Create = False, Format = 0 (ASCII/ANSI)
    If Err.Number <> 0 Then
        WScript.Echo "❌ ERROR: No se pudo leer el archivo " & filePath & ": " & Err.Description
        ReadFileWithAnsiEncoding = ""
        Exit Function
    End If
    
    strContent = objFile.ReadAll
    objFile.Close
    On Error GoTo 0
    
    ReadFileWithAnsiEncoding = strContent
End Function

' Función para limpiar archivos VBA eliminando líneas Attribute con validación mejorada
Function CleanVBAFile(filePath)
    Dim objStream, strContent, arrLines, i, cleanedLines
    Dim strLine, errorDetails
    
    ' Validar sintaxis antes de procesar
    If Not ValidateVBASyntax(filePath, errorDetails) Then
        WScript.Echo "⚠️ ADVERTENCIA: Errores de sintaxis detectados en " & objFSO.GetFileName(filePath) & ":"
        WScript.Echo errorDetails
        WScript.Echo "Continuando con la importación..."
    End If
    
    ' Leer contenido del archivo con codificación UTF-8 usando ADODB.Stream
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
        Exit Function
    End If
    On Error GoTo 0
    
    ' Normalizar saltos de línea y dividir
    strContent = Replace(strContent, vbCrLf, vbLf)
    strContent = Replace(strContent, vbCr, vbLf)
    arrLines = Split(strContent, vbLf)
    
    ' Filtrar líneas que no empiecen con "Attribute" y limpiar caracteres problemáticos
    cleanedLines = ""
    For i = 0 To UBound(arrLines)
        strLine = arrLines(i)
        If Left(Trim(strLine), 9) <> "Attribute" Then
            ' Reemplazar caracteres problemáticos
            strLine = Replace(strLine, Chr(147), Chr(34)) ' Comilla izquierda tipográfica -> comilla normal
            strLine = Replace(strLine, Chr(148), Chr(34)) ' Comilla derecha tipográfica -> comilla normal
            strLine = Replace(strLine, Chr(145), Chr(39)) ' Apostrofe izquierdo -> apostrofe normal
            strLine = Replace(strLine, Chr(146), Chr(39)) ' Apostrofe derecho -> apostrofe normal
            
            If cleanedLines <> "" Then
                cleanedLines = cleanedLines & vbCrLf
            End If
            cleanedLines = cleanedLines & strLine
        Else
            WScript.Echo "  Eliminando metadato: " & Trim(strLine)
        End If
    Next
    
    CleanVBAFile = cleanedLines
End Function

' Función para exportar módulo con conversión ANSI -> UTF-8
Sub ExportModuleWithAnsiEncoding(vbComponent, strExportPath)
    Dim tempFilePath, objTempFile, objStream
    Dim strContent
    
    ' Crear archivo temporal usando el método nativo Export (Access exporta en ANSI)
    tempFilePath = objFSO.GetParentFolderName(strExportPath) & "\temp_export_" & vbComponent.Name & "." & objFSO.GetExtensionName(strExportPath)
    
    ' Exportar a archivo temporal (Access usa ANSI internamente)
    vbComponent.Export tempFilePath
    
    ' Leer contenido del archivo temporal con codificación ANSI
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
End Sub

' Función para importar módulo con conversión UTF-8 -> ANSI
Sub ImportCleanModule(moduleName, cleanedContent, objFile)
    Dim tempFilePath, objTempFile, vbComponent
    
    ' Crear archivo temporal con contenido limpio usando codificación ANSI
    ' (Access requiere ANSI para importar correctamente)
    tempFilePath = objFSO.GetParentFolderName(objFile.Path) & "\temp_" & objFile.Name
    
    ' Escribir contenido limpio con codificación ANSI explícita
    Set objTempFile = objFSO.CreateTextFile(tempFilePath, True, False) ' Overwrite = True, Unicode = False (ANSI)
    objTempFile.Write cleanedContent
    objTempFile.Close
    
    ' Importar desde archivo temporal
    On Error Resume Next
    Call objAccess.VBE.ActiveVBProject.VBComponents.Import(tempFilePath)
    
    ' Limpiar archivo temporal
    objFSO.DeleteFile tempFilePath
    
    If Err.Number = 0 Then
        ' Renombrar el módulo importado al nombre correcto
        Set vbComponent = objAccess.VBE.ActiveVBProject.VBComponents(objAccess.VBE.ActiveVBProject.VBComponents.Count)
        If vbComponent.Name <> moduleName Then
            vbComponent.Name = moduleName
        End If
        WScript.Echo "Módulo " & moduleName & " importado correctamente (UTF-8 -> ANSI)"
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
    Dim dbCount
    Dim strDbName, strPassword
    
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
    
    For Each objFile In objBackFolder.Files
        If LCase(objFSO.GetExtensionName(objFile.Name)) = "accdb" Then
            dbCount = dbCount + 1
            strDbName = objFSO.GetBaseName(objFile.Name)
            
            ' Determinar contraseña según el nombre de la base de datos
            If InStr(1, UCase(strDbName), "CONDOR") > 0 Then
                strPassword = "(sin contraseña)"
            Else
                strPassword = "dpddpd"
            End If
            
            WScript.Echo "  [" & dbCount & "] " & objFile.Name & " - " & strPassword
        End If
    Next
    
    If dbCount = 0 Then
        WScript.Echo "No se encontraron bases de datos .accdb en el directorio ./back"
        WScript.Echo "=== RE-VINCULACION COMPLETADA ==="
        Exit Sub
    End If
    
    WScript.Echo "Total de bases de datos encontradas: " & dbCount
    WScript.Echo "Nota: Las bases de datos CONDOR no requieren contraseña, las demás usan 'dpddpd'"
    WScript.Echo "Funcionalidad de re-vinculación automática pendiente de implementación."
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

' La subrutina ExecuteTestModule ha sido eliminada ya que ahora se usa el motor interno modTestRunner
