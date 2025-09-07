Attribute VB_Name = "modTestUtils"
Option Compare Database
Option Explicit

' ===================================================================
' MÓDULO: modTestUtils
' DESCRIPCIÓN: Utilidades centrales para el framework de testing.
'              ÚNICA FUENTE DE VERDAD para la gestión del entorno de pruebas.
' ARQUITECTURA: back/test_env/ con /fixtures/ y /workspace/
' ===================================================================

Private Const FIXTURES_DATABASES_PATH As String = "back\test_env\fixtures\databases\"
Private Const WORKSPACE_PATH As String = "back\test_env\workspace\"

' Devuelve el nombre de la BD real de ENTORNO para una activa del workspace.
Private Function MapActiveToEnvBase(ByVal activeDbName As String) As String
    ' Activa del workspace -> BD real de ENTORNO
    Dim n As String: n = UCase$(activeDbName)
    ' Normaliza sufijos de test
    n = Replace(n, "_INTEGRATION_TEST.ACCDB", ".ACCDB")
    n = Replace(n, "_WORKSPACE_TEST.ACCDB", ".ACCDB")
    
    ' --- Mapeos por dominio ---
    ' Solicitud debe venir de Solicitud.accdb (no de CONDOR), con fallback posterior si no existiera
    If InStr(n, "SOLICITUD.ACCDB") > 0 Then MapActiveToEnvBase = "Solicitud.accdb": Exit Function
    If InStr(n, "DOCUMENT.ACCDB") > 0 Then MapActiveToEnvBase = "Document.accdb": Exit Function
    If InStr(n, "EXPEDIENTES.ACCDB") > 0 Then MapActiveToEnvBase = "Expedientes.accdb": Exit Function
    If InStr(n, "WORKFLOW.ACCDB") > 0 Then MapActiveToEnvBase = "Workflow.accdb": Exit Function
    If InStr(n, "LANZADERA.ACCDB") > 0 Then MapActiveToEnvBase = "Lanzadera.accdb": Exit Function
    If InStr(n, "CORREOS.ACCDB") > 0 Then MapActiveToEnvBase = "correos.accdb": Exit Function
    
    ' Por defecto: nombre normalizado
    MapActiveToEnvBase = Replace(n, ".ACCDB", ".accdb")
End Function

' Mapea nombre base a la clave de config *_DATA_PATH
Private Function MapDbKey(ByVal envBaseName As String) As String
    Dim n As String: n = UCase$(envBaseName)
    If InStr(n, "SOLICITUD.ACCDB")  > 0 Then MapDbKey = "SOLICITUD_DATA_PATH": Exit Function
    If InStr(n, "CONDOR.ACCDB")     > 0 Then MapDbKey = "CONDOR_DATA_PATH": Exit Function
    If InStr(n, "DOCUMENT.ACCDB")   > 0 Then MapDbKey = "DOCUMENT_DATA_PATH": Exit Function
    If InStr(n, "EXPEDIENTES.ACCDB")> 0 Then MapDbKey = "EXPEDIENTES_DATA_PATH": Exit Function
    If InStr(n, "WORKFLOW.ACCDB")   > 0 Then MapDbKey = "WORKFLOW_DATA_PATH": Exit Function
    If InStr(n, "LANZADERA.ACCDB")  > 0 Then MapDbKey = "LANZADERA_DATA_PATH": Exit Function
    If InStr(n, "CORREOS.ACCDB")    > 0 Then MapDbKey = "CORREOS_DATA_PATH": Exit Function
    MapDbKey = ""
End Function

' Devuelve la ruta de origen (ENTORNO -> Desarrollo -> Fixture)
Private Function ResolveEnvDbSource(ByVal templateDbName As String, ByVal activeDbName As String, ByRef why As String) As String
    On Error GoTo ExitFn
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim project As String: project = GetProjectPath()
    Dim envBase As String, src As String: why = "": src = ""
    
    ' 0) Activa de test -> BD real de ENTORNO (p.ej., Solicitud_* -> Solicitud.accdb)
    envBase = MapActiveToEnvBase(activeDbName)
    
    ' 1) ENTORNO vía IConfig (clave *_DATA_PATH)
    Dim cfg As IConfig: Set cfg = modTestContext.GetTestConfig()
    Dim key As String: key = MapDbKey(envBase)
    Dim cfgPath As String: cfgPath = ""
    On Error Resume Next
    If Not cfg Is Nothing And Len(key) > 0 Then cfgPath = cfg.GetValue(key)
    On Error GoTo 0
    If Len(cfgPath) > 0 And fso.FileExists(cfgPath) Then
        src = cfgPath: why = "entorno(" & key & ")": GoTo ExitOk
    End If
    
    ' 2) Desarrollo local (back\Desarrollo\<envBase>)
    Dim devPath As String: devPath = JoinPath(project, "back\Desarrollo\" & envBase)
    If fso.FileExists(devPath) Then
        src = devPath: why = "desarrollo": GoTo ExitOk
    End If
    
    ' 3) Fixture (último recurso): back\test_env\fixtures\databases\<template>
    Dim fixturePath As String: fixturePath = JoinPath(project, "back\test_env\fixtures\databases\" & templateDbName)
    If fso.FileExists(fixturePath) Then
        src = fixturePath: why = "fixture": GoTo ExitOk
    End If
    
    ' 4) Fallback final a CONDOR (por compatibilidad si no hay Solicitud.accdb en ningún sitio)
    Dim condorDev As String: condorDev = JoinPath(project, "back\Desarrollo\CONDOR.accdb")
    If fso.FileExists(condorDev) Then
        src = condorDev: why = "fallback(CONDOR.desarrollo)": GoTo ExitOk
    End If
    If Not cfg Is Nothing Then
        On Error Resume Next
        cfgPath = cfg.GetCondorDataPath()
        On Error GoTo 0
        If Len(cfgPath) > 0 And fso.FileExists(cfgPath) Then
            src = cfgPath: why = "fallback(Condor.config)": GoTo ExitOk
        End If
    End If
    
ExitOk:
    ResolveEnvDbSource = src
ExitFn:
    On Error Resume Next: Set fso = Nothing
End Function

Public Function GetProjectPath() As String
    ' Devuelve la ruta raíz del proyecto.
    GetProjectPath = Left(CurrentProject.FullName, InStrRev(CurrentProject.FullName, "\back\") - 1)
End Function

Public Function GetDatabaseFixturesPath() As String
    ' Devuelve la ruta a las plantillas maestras de BD para los tests.
    GetDatabaseFixturesPath = GetProjectPath() & "\" & FIXTURES_DATABASES_PATH
End Function

Public Function GetWorkspacePath() As String
    ' Devuelve la ruta al espacio de trabajo volátil para los tests.
    GetWorkspacePath = GetProjectPath() & "\" & WORKSPACE_PATH
End Function

' Devuelve la ruta de origen para la BD solicitada.
Private Function ResolveDbSourcePath(ByVal templateDbName As String, ByRef why As String) As String
    On Error GoTo ExitFn
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim project As String: project = GetProjectPath()
    Dim fixtures As String: fixtures = JoinPath(project, "back\test_env\fixtures\databases\" & templateDbName)
    Dim src As String: src = ""
    why = ""

    ' 1) Opción A: Fixture oficial
    If fso.FileExists(fixtures) Then
        src = fixtures: why = "fixture"
        GoTo ExitOk
    End If

    ' Derivar nombre base: Workflow_test_template.accdb -> Workflow.accdb
    Dim base As String: base = templateDbName
    If LCase$(Right$(base, 20)) = "_test_template.accdb" Then
        base = Left$(base, Len(base) - 20) & ".accdb"
    End If

    ' 2) Opción B: Carpeta de Desarrollo
    Dim devCandidates As Variant
    devCandidates = Array( _
        JoinPath(project, "back\Desarrollo\" & base), _
        JoinPath(project, "back\Datos\" & base), _
        JoinPath(project, "back\databases\" & base) _
    )
    Dim i As Long
    For i = LBound(devCandidates) To UBound(devCandidates)
        If fso.FileExists(devCandidates(i)) Then
            src = devCandidates(i): why = "desarrollo(" & CStr(i + 1) & ")"
            GoTo ExitOk
        End If
    Next i

    ' 3) Opción C: Ruta desde IConfig (*_DATA_PATH)
    On Error Resume Next
    Dim cfg As IConfig: Set cfg = modTestContext.GetTestConfig()
    Dim key As String: key = UCase$(Replace(base, ".accdb", "")) & "_DATA_PATH"
    Dim cfgPath As String: cfgPath = ""
    If Not cfg Is Nothing Then cfgPath = cfg.GetValue(key)
    On Error GoTo ExitFn
    If Len(cfgPath) > 0 And fso.FileExists(cfgPath) Then
        src = cfgPath: why = "config(" & key & ")"
        GoTo ExitOk
    End If

ExitOk:
    ResolveDbSourcePath = src
ExitFn:
    On Error Resume Next: Set fso = Nothing
End Function

Public Sub PrepareTestDatabase(ByVal templateDbName As String, ByVal activeDbName As String)
    On Error GoTo Fail
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim ws As String: ws = GetWorkspacePath(): EnsureFolder ws
    
    Dim why As String, src As String
    src = ResolveEnvDbSource(templateDbName, activeDbName, why)
    If Len(src) = 0 Then
        Err.Raise vbObjectError + 1001, "PrepareTestDatabase", _
                  "No se encontró la BD de origen para '" & templateDbName & "' en entorno/desarrollo/fixture."
    End If
    
    Dim dst As String: dst = JoinPath(ws, activeDbName)
    If fso.FileExists(dst) Then fso.DeleteFile dst, True
    fso.CopyFile src, dst, True
    
    On Error Resume Next: Debug.Print "[PrepareTestDatabase] " & activeDbName & " <- " & src & " [" & why & "]"
    On Error GoTo 0
    Set fso = Nothing: Exit Sub
Fail:
    Err.Raise Err.Number, "modTestUtils.PrepareTestDatabase", Err.Description
End Sub

' ===================================================================
' PROVISIONING SIMPLIFICADO - Helpers privados
' ===================================================================

Private Function IsUnderPath(ByVal child As String, ByVal base As String) As Boolean
    Dim c As String, b As String
    c = UCase$(Replace(child, "/", "\")): b = UCase$(Replace(base, "/", "\"))
    If Right$(b, 1) <> "\" Then b = b & "\"
    IsUnderPath = (Left$(c, Len(b)) = b)
End Function

Private Function ResolveEnvPath(ByVal cfg As IConfig, ByVal project As String, ByVal ws As String, ByVal envBase As String, ByVal configKey As String, ByRef why As String) As String
    On Error GoTo ExitFn
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim p As String, i As Long
    ' 1) ENTORNO (IConfig) si no apunta al workspace
    p = "": If Not cfg Is Nothing And Len(configKey) > 0 Then p = cfg.GetValue(configKey)
    If Len(p) > 0 And Not IsUnderPath(p, ws) And fso.FileExists(p) Then why = "entorno(" & configKey & ")": ResolveEnvPath = p: GoTo ExitFn
    ' 2) Desarrollo (varios candidatos)
    Dim dev() As String: dev = Split("back\Desarrollo|back\Datos|back\databases|back", "|")
    For i = LBound(dev) To UBound(dev)
        p = JoinPath(project, dev(i) & "\" & envBase)
        If fso.FileExists(p) Then why = "desarrollo(" & CStr(i + 1) & ")": ResolveEnvPath = p: GoTo ExitFn
    Next i
    ' 3) Fixture (último recurso): <NombreBase>_test_template.accdb
    p = JoinPath(project, "back\test_env\fixtures\databases\" & Replace(envBase, ".accdb", "_test_template.accdb"))
    If fso.FileExists(p) Then why = "fixture": ResolveEnvPath = p
ExitFn:
    On Error Resume Next: Set fso = Nothing
End Function

' ===================================================================
' PROVISIONING SIMPLIFICADO - Orquestador público
' ===================================================================

Public Sub PrepareCoreTestDatabases()
    On Error GoTo Fail
    Dim project As String: project = GetProjectPath()
    Dim ws As String: ws = GetWorkspacePath(): EnsureFolder ws
    Dim cfg As IConfig: Set cfg = modTestContext.GetTestConfig()
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    ' 4 BDs que sí se copian
    Dim entries As Variant
    entries = Array( _
        "Lanzadera_workspace_test.accdb|Lanzadera.accdb|LANZADERA_DATA_PATH", _
        "Expedientes_integration_test.accdb|Expedientes.accdb|EXPEDIENTES_DATA_PATH", _
        "correos_integration_test.accdb|correos.accdb|CORREOS_DATA_PATH", _
        "CondorFront_integration_test.accdb|CondorFront.accdb|CONDOR_FRONT_DATA_PATH" _
    )
    Dim i As Long, parts As Variant, src As String, dst As String, why As String
    For i = LBound(entries) To UBound(entries)
        parts = Split(CStr(entries(i)), "|")
        why = "": src = ResolveEnvPath(cfg, project, ws, parts(1), parts(2), why)
        If Len(src) = 0 Then Err.Raise vbObjectError + 2001, "PrepareCoreTestDatabases", "Sin origen para " & parts(0)
        dst = JoinPath(ws, parts(0))
        If fso.FileExists(dst) Then fso.DeleteFile dst, True
        fso.CopyFile src, dst, True
        On Error Resume Next: Debug.Print "[PrepareCoreTestDatabase] " & parts(0) & " <- " & src & " [" & why & "]": On Error GoTo Fail
    Next i
    Set fso = Nothing: Exit Sub
Fail:
    Err.Raise Err.Number, "modTestUtils.PrepareCoreTestDatabases", Err.Description
End Sub

' === Path Helpers ===
' Devuelve el directorio padre de una ruta de archivo o carpeta.
' Soporta rutas con barra final y UNC. Si no hay padre (p.ej. "C:"), devuelve "".
Public Function GetParentDirectory(ByVal fullPath As String) As String
    On Error GoTo ExitFn
    Dim fso As Object
    Dim p As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    p = Trim$(fullPath)
    If Len(p) = 0 Then GoTo ExitFn
    ' Normaliza: quita barra final si la hay
    If Right$(p, 1) = "\" Or Right$(p, 1) = "/" Then
        p = Left$(p, Len(p) - 1)
    End If
    GetParentDirectory = fso.GetParentFolderName(p)
ExitFn:
    On Error Resume Next
    Set fso = Nothing
End Function

Public Sub CleanupTestDatabase(ByVal activeDbName As String)
    ' Orquesta la limpieza de una BD de prueba del workspace.
    ' Es la contraparte simétrica de PrepareTestDatabase.
    On Error Resume Next ' Usamos Resume Next porque el objetivo es asegurar que no quede, sin importar si ya no estaba.

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim activePath As String
    activePath = GetWorkspacePath & activeDbName

    If fso.FileExists(activePath) Then
        fso.DeleteFile activePath, True
    End If
    
    Set fso = Nothing
End Sub

Public Sub CleanupTestFolder(ByVal folderName As String)
    ' Limpia una carpeta completa dentro del workspace.
    ' Usado por tests que crean múltiples ficheros (Word, DocumentService, etc.).
    ' NOTA: folderName debe ser relativo al workspace.
    On Error Resume Next

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim folderPath As String
    folderPath = GetWorkspacePath & folderName

    If fso.FolderExists(folderPath) Then
        fso.DeleteFolder folderPath, True
    End If
    
    Set fso = Nothing
End Sub

Public Sub EnsureFolder(ByVal folderPath As String)
    On Error GoTo ErrorHandler
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(folderPath) Then GoTo ExitOk
    Dim parent As String: parent = fso.GetParentFolderName(folderPath)
    If Len(parent) > 0 And Not fso.FolderExists(parent) Then
        EnsureFolder parent
    End If
    If Not fso.FolderExists(folderPath) Then fso.CreateFolder folderPath
ExitOk:
    Set fso = Nothing
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, "modTestUtils.EnsureFolder", "Error al crear: " & folderPath & " - " & Err.Description
End Sub

Public Function JoinPath(ByVal basePath As String, ByVal relativePath As String) As String
    Dim b As String: b = Trim$(basePath)
    Dim r As String: r = Trim$(relativePath)
    If Len(b) = 0 Then JoinPath = r: Exit Function
    If Len(r) = 0 Then JoinPath = b: Exit Function
    If Right$(b, 1) = "\" Or Right$(b, 1) = "/" Then b = Left$(b, Len(b) - 1)
    If Left$(r, 1) = "\" Or Left$(r, 1) = "/" Then r = Mid$(r, 2)
    JoinPath = b & "\" & r
End Function

Public Function VerifyAllTemplates() As Boolean
    ' Verifica la existencia de todas las plantillas de BD con el nuevo naming estándar.
    ' Devuelve True si todas las plantillas existen, False en caso contrario.
    On Error GoTo ErrorHandler
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim fixturesPath As String
    fixturesPath = GetDatabaseFixturesPath()
    
    ' Lista de plantillas requeridas según el nuevo estándar de naming
    Dim templates As Variant
    templates = Array( _
        "Document_test_template.accdb", _
        "Mapeo_test_template.accdb", _
        "Operation_test_template.accdb", _
        "Solicitud_test_template.accdb", _
        "Workflow_test_template.accdb", _
        "Lanzadera_test_template.accdb", _
        "Expedientes_test_template.accdb", _
        "correos_test_template.accdb" _
    )
    
    Dim i As Integer
    Dim templatePath As String
    Dim allExist As Boolean
    allExist = True
    
    For i = 0 To UBound(templates)
        templatePath = fixturesPath & templates(i)
        If Not fso.FileExists(templatePath) Then
            Debug.Print "❌ PLANTILLA FALTANTE: " & templatePath
            allExist = False
        Else
            Debug.Print "✓ PLANTILLA ENCONTRADA: " & templatePath
        End If
    Next i
    
    VerifyAllTemplates = allExist
    Set fso = Nothing
    Exit Function
    
ErrorHandler:
    Debug.Print "❌ ERROR EN VERIFICACIÓN DE PLANTILLAS: " & Err.Description
    VerifyAllTemplates = False
    Set fso = Nothing
End Function

' ===================================================================
' UTILIDADES DE LIMPIEZA DE WORD PARA TESTING
' ===================================================================

'
' Cierra la instancia COM principal de Word si existe (sin guardar).
Public Sub CloseAllWordInstancesForTesting()
    On Error Resume Next
    Dim wd As Object
    Set wd = GetObject(, "Word.Application")
    If Not wd Is Nothing Then
        Do While wd.Documents.Count > 0
            wd.Documents(1).Close False
        Loop
        wd.Quit 0
    End If
    Set wd = Nothing
End Sub

'
' Convierte CreationDate WMI (yyyymmddHHMMSS.mmmmmms+zzz) a Date VBA
Private Function WmiCreationToDate(ByVal s As String) As Date
    On Error Resume Next
    Dim y As Integer, m As Integer, d As Integer, hh As Integer, nn As Integer, ss As Integer
    y = CInt(Left$(s, 4)): m = CInt(Mid$(s, 5, 2)): d = CInt(Mid$(s, 7, 2))
    hh = CInt(Mid$(s, 9, 2)): nn = CInt(Mid$(s, 11, 2)): ss = CInt(Mid$(s, 13, 2))
    WmiCreationToDate = DateSerial(y, m, d) + TimeSerial(hh, nn, ss)
End Function

'
' Mata procesos WINWORD.EXE recientes (solo en modo testing) como último recurso.
Public Sub KillRecentWordProcesses(Optional ByVal maxAgeMinutes As Long = 10, Optional ByVal onlyWhenTestMode As Boolean = True)
    On Error Resume Next
    If onlyWhenTestMode Then
        If Environ$("CONDOR_TEST_MODE") <> "1" Then Exit Sub
    End If
    Dim svc As Object, procs As Object, p As Object
    Set svc = GetObject("winmgmts:root\cimv2")
    Set procs = svc.ExecQuery("SELECT ProcessId,CreationDate FROM Win32_Process WHERE Name='WINWORD.EXE'")
    For Each p In procs
        If DateDiff("n", WmiCreationToDate(p.CreationDate), Now) <= maxAgeMinutes Then
            svc.Get("Win32_Process.Handle='" & p.ProcessId & "'").Terminate ' ignora errores
        End If
    Next
End Sub