Option Compare Database
Option Explicit
' Macro VBA para reconstrucción completa del proyecto
' Esta macro puede ser ejecutada desde VBScript usando objAccess.Run "RebuildProject"
Public Sub RebuildProject()
    On Error GoTo ErrorHandler
    
    DoCmd.SetWarnings False
    Application.DisplayAlerts = False
    
    Debug.Print "=== RECONSTRUCCION COMPLETA DEL PROYECTO VBA ==="
    Debug.Print "ADVERTENCIA: Se eliminaran TODOS los modulos VBA existentes"
    Debug.Print "Iniciando proceso de reconstruccion con limpieza total..."
    
    ' FASE 1: LIMPIEZA TOTAL
    Debug.Print "FASE 1: LIMPIEZA TOTAL - Eliminando todos los modulos VBA existentes..."
    
    Dim vbProject As Object
    Dim vbComponent As Object
    Set vbProject = Application.VBE.ActiveVBProject
    
    Dim componentCount As Integer
    Dim i As Integer
    componentCount = vbProject.VBComponents.Count
    
    ' Iterar sobre todos los componentes en Application.VBE.ActiveVBProject.VBComponents
    For i = componentCount To 1 Step -1
        Set vbComponent = vbProject.VBComponents(i)
        
        ' Eliminar únicamente módulos estándar (vbext_ct_StdModule) y de clase (vbext_ct_ClassModule)
        If vbComponent.Type = 1 Or vbComponent.Type = 2 Then ' vbext_ct_StdModule = 1, vbext_ct_ClassModule = 2
            Debug.Print "  Eliminando: " & vbComponent.Name & " (Tipo: " & GetComponentTypeName(vbComponent.Type) & ")"
            vbProject.VBComponents.Remove vbComponent
        End If
    Next i
    
    ' FASE 2: VERIFICACION POST-LIMPIEZA
    Debug.Print "FASE 2: VERIFICACION POST-LIMPIEZA..."
    Dim remainingModules As Integer
    remainingModules = 0
    
    For Each vbComponent In vbProject.VBComponents
        If vbComponent.Type = 1 Or vbComponent.Type = 2 Then
            remainingModules = remainingModules + 1
        End If
    Next vbComponent
    
    Debug.Print "Módulos restantes después de limpieza: " & remainingModules
    
    ' FASE 3: IMPORTACION
    Debug.Print "FASE 3: IMPORTACION - Importando archivos desde /src..."
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim strSourcePath As String
    strSourcePath = CurrentProject.path & "\src"
    
    If Not fso.FolderExists(strSourcePath) Then
        Debug.Print "Error: Directorio de origen no existe: " & strSourcePath
        Exit Sub
    End If
    
    Dim objFolder As Object
    Dim objFile As Object
    Set objFolder = fso.GetFolder(strSourcePath)
    
    Dim importedCount As Integer
    importedCount = 0
    
    For Each objFile In objFolder.Files
        If LCase(fso.GetExtensionName(objFile.Name)) = "bas" Or LCase(fso.GetExtensionName(objFile.Name)) = "cls" Then
            Debug.Print "Importando: " & objFile.Name
            
            ' Importar el módulo
            If LCase(fso.GetExtensionName(objFile.Name)) = "bas" Then
                Application.LoadFromText acModule, fso.GetBaseName(objFile.Name), objFile.path
            Else
                Application.LoadFromText acClassModule, fso.GetBaseName(objFile.Name), objFile.path
            End If
            
            importedCount = importedCount + 1
            Debug.Print "? " & objFile.Name & " importado correctamente"
        End If
    Next objFile
    
    ' FASE 4: VERIFICACION POST-IMPORTACION
    Debug.Print "FASE 4: VERIFICACION POST-IMPORTACION..."
    
    Dim finalModuleCount As Integer
    finalModuleCount = 0
    
    For Each vbComponent In vbProject.VBComponents
        If vbComponent.Type = 1 Or vbComponent.Type = 2 Then
            finalModuleCount = finalModuleCount + 1
        End If
    Next vbComponent
    
    Debug.Print "=== RESUMEN DE RECONSTRUCCION ==="
    Debug.Print "Archivos importados: " & importedCount
    Debug.Print "Módulos VBA finales: " & finalModuleCount
    
    If importedCount = finalModuleCount Then
        Debug.Print "? RECONSTRUCCION COMPLETADA EXITOSAMENTE"
    Else
        Debug.Print "? ADVERTENCIA: Discrepancia en el conteo de módulos"
    End If
    
    DoCmd.SetWarnings True
    Application.DisplayAlerts = True
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error en RebuildProject: " & Err.Description
    DoCmd.SetWarnings True
    Application.DisplayAlerts = True
End Sub

' Función auxiliar para obtener el nombre del tipo de componente
Private Function GetComponentTypeName(componentType As Integer) As String
    Select Case componentType
        Case 1: GetComponentTypeName = "Módulo Estándar"
        Case 2: GetComponentTypeName = "Módulo de Clase"
        Case 3: GetComponentTypeName = "Formulario"
        Case 100: GetComponentTypeName = "Documento"
        Case Else: GetComponentTypeName = "Desconocido (" & componentType & ")"
    End Select
End Function






