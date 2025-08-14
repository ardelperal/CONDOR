Attribute VB_Name = "Test_Config"

' =============================================================================
' MODULO DE PRUEBAS PARA CONFIGURACION Y ENTORNOS
' =============================================================================
' Proposito: Verificar el correcto funcionamiento del modulo modConfig
' =============================================================================

' Prueba principal que ejecuta todas las verificaciones
Public Sub Test_ConfigModule()
    Debug.Print "=== INICIANDO PRUEBAS DE CONFIGURACION ==="
    
    ' Llamar a InitializeEnvironment
    Call modConfig.InitializeEnvironment
    
    ' Verificar que la configuracion se inicializo
    If Not modConfig.IsConfigInitialized() = True Then
        Debug.Print "Error: Configuracion no inicializada"
        Stop
    End If
    Debug.Print "OK Configuracion inicializada correctamente"
    
    ' Verificar informacion del entorno
    Dim envInfo As String
    envInfo = modConfig.GetEnvironmentInfo()
    If Not envInfo = "DESARROLLO (Local)" Then
        Debug.Print "Error: Entorno incorrecto - " & envInfo
        Stop
    End If
    Debug.Print "OK Entorno detectado: " & envInfo
    
    ' Verificar rutas de desarrollo (DEV_MODE = 1)
    Call Test_DevelopmentPaths
    
    ' Verificar funciones de utilidad
    Call Test_UtilityFunctions
    
    Debug.Print "=== TODAS LAS PRUEBAS DE CONFIGURACION PASARON ==="
End Sub

' Prueba las rutas de desarrollo
Private Sub Test_DevelopmentPaths()
    Debug.Print "--- Verificando rutas de desarrollo ---"
    
    ' Verificar ruta de base de datos CONDOR
    Dim condorPath As String
    condorPath = modConfig.GetCondorDbPath()
    If Not condorPath = "./back/CONDOR_datos.accdb" Then
        Debug.Print "Error: Ruta CONDOR incorrecta - " & condorPath
        Stop
    End If
    Debug.Print "OK Ruta CONDOR: " & condorPath
    
    ' Verificar ruta de base de datos Expedientes
    Dim expedientesPath As String
    expedientesPath = modConfig.GetExpedientesDbPath()
    If Not expedientesPath = "./back/Expedientes_Local.accdb" Then
        Debug.Print "Error: Ruta Expedientes incorrecta - " & expedientesPath
        Stop
    End If
    Debug.Print "OK Ruta Expedientes: " & expedientesPath
    
    ' Verificar ruta de plantillas
    Dim plantillasPath As String
    plantillasPath = modConfig.GetPlantillasPath()
    If Not plantillasPath = "./docs/Plantillas/" Then
        Debug.Print "Error: Ruta Plantillas incorrecta - " & plantillasPath
        Stop
    End If
    Debug.Print "OK Ruta Plantillas: " & plantillasPath
    
    ' Verificar ruta de logs
    Dim logsPath As String
    logsPath = modConfig.GetLogsPath()
    If Not logsPath = "./logs/" Then
        Debug.Print "Error: Ruta Logs incorrecta - " & logsPath
        Stop
    End If
    Debug.Print "OK Ruta Logs: " & logsPath
    
    ' Verificar ruta de backup
    Dim backupPath As String
    backupPath = modConfig.GetBackupPath()
    If Not backupPath = "./backup/" Then
        Debug.Print "Error: Ruta Backup incorrecta - " & backupPath
        Stop
    End If
    Debug.Print "OK Ruta Backup: " & backupPath
    
    ' Verificar ruta temporal
    Dim tempPath As String
    tempPath = modConfig.GetTempPath()
    If Not tempPath = "./temp/" Then
        Debug.Print "Error: Ruta Temp incorrecta - " & tempPath
        Stop
    End If
    Debug.Print "OK Ruta Temp: " & tempPath
End Sub

' Prueba las funciones de utilidad
Private Sub Test_UtilityFunctions()
    Debug.Print "--- Verificando funciones de utilidad ---"
    
    ' Verificar reset de configuracion
    Call modConfig.ResetConfiguration
    If Not modConfig.IsConfigInitialized() = True Then
        Debug.Print "Error: Reset no funciono correctamente"
        Stop
    End If
    Debug.Print "OK Reset de configuracion funciona correctamente"
    
    ' Verificar que las rutas siguen siendo correctas despues del reset
    If Not modConfig.GetCondorDbPath() = "./back/CONDOR_datos.accdb" Then
        Debug.Print "Error: Ruta CONDOR incorrecta despues del reset"
        Stop
    End If
    Debug.Print "OK Rutas mantienen consistencia despues del reset"
End Sub

' Prueba de inicializacion automatica
Public Sub Test_AutoInitialization()
    Debug.Print "=== PRUEBA DE INICIALIZACION AUTOMATICA ==="
    
    ' Resetear configuracion para simular estado inicial
    Call modConfig.ResetConfiguration
    
    ' Llamar directamente a una funcion getter sin inicializar explicitamente
    Dim path As String
    path = modConfig.GetCondorDbPath()
    
    ' Verificar que se inicializo automaticamente
    If Not modConfig.IsConfigInitialized() = True Then
        Debug.Print "Error: No se inicializo automaticamente"
        Stop
    End If
    If Not path = "./back/CONDOR_datos.accdb" Then
        Debug.Print "Error: Ruta incorrecta en inicializacion automatica"
        Stop
    End If
    
    Debug.Print "OK Inicializacion automatica funciona correctamente"
    Debug.Print "=== PRUEBA DE INICIALIZACION AUTOMATICA COMPLETADA ==="
End Sub

' Prueba de consistencia de todas las rutas
Public Sub Test_PathConsistency()
    Debug.Print "=== PRUEBA DE CONSISTENCIA DE RUTAS ==="
    
    Call modConfig.InitializeEnvironment
    
    ' Obtener todas las rutas multiples veces
    Dim i As Integer
    For i = 1 To 3
        If Not modConfig.GetCondorDbPath() = "./back/CONDOR_datos.accdb" Then
            Debug.Print "Error: Inconsistencia en ruta CONDOR - iteracion " & i
            Stop
        End If
        If Not modConfig.GetExpedientesDbPath() = "./back/Expedientes_Local.accdb" Then
            Debug.Print "Error: Inconsistencia en ruta Expedientes - iteracion " & i
            Stop
        End If
        If Not modConfig.GetPlantillasPath() = "./docs/Plantillas/" Then
            Debug.Print "Error: Inconsistencia en ruta Plantillas - iteracion " & i
            Stop
        End If
    Next i
    
    Debug.Print "OK Todas las rutas mantienen consistencia en multiples llamadas"
    Debug.Print "=== PRUEBA DE CONSISTENCIA COMPLETADA ==="
End Sub