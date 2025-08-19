# Documentación de Despliegue y Configuración - Sistema CONDOR

## Introducción

Este documento describe los procedimientos de despliegue, configuración y administración del Sistema CONDOR, basado en la arquitectura cliente-servidor definida en la especificación funcional del proyecto.

---

## 1. Arquitectura de Despliegue

### 1.1 Componentes del Sistema

```
┌─────────────────────────────────────────────────────────────────┐
│                        SERVIDOR CENTRAL                        │
│  ┌─────────────────┐  ┌─────────────────┐  ┌─────────────────┐ │
│  │   CONDOR.accde  │  │ CONDOR_datos.   │  │   Lanzadera     │ │
│  │   (Frontend)    │  │    accdb        │  │ condor_cli.vbs  │ │
│  │                 │  │  (Backend)      │  │                 │ │
│  └─────────────────┘  └─────────────────┘  └─────────────────┘ │
│           │                     │                     │         │
│           └─────────────────────┼─────────────────────┘         │
│                                 │                               │
└─────────────────────────────────┼─────────────────────────────────┘
                                  │
                    ┌─────────────┴─────────────┐
                    │                           │
          ┌─────────▼─────────┐       ┌─────────▼─────────┐
          │   CLIENTE 1       │       │   CLIENTE N       │
          │                   │       │                   │
          │ • Acceso directo  │  ...  │ • Acceso directo  │
          │   al servidor     │       │   al servidor     │
          │ • Actualización   │       │ • Actualización   │
          │   automática      │       │   automática      │
          └───────────────────┘       └───────────────────┘
```

### 1.2 Ubicaciones de Archivos

#### Servidor Central
```
\\SERVIDOR\CONDOR\
├── CONDOR.accde              # Frontend compilado (solo lectura)
├── CONDOR_datos.accdb        # Base de datos backend
├── condor_cli.vbs           # Script de lanzadera
├── config\
│   ├── configuracion.ini     # Configuración del sistema
│   └── entornos.ini         # Configuración de entornos
├── plantillas\
│   ├── PC_template.docx     # Plantilla Cambio de Precio
│   ├── CDCA_template.docx   # Plantilla Concesión/Desviación
│   └── CDCASUB_template.docx # Plantilla Sub-Concesión
├── logs\
│   ├── sistema.log          # Log del sistema
│   ├── errores.log          # Log de errores
│   └── accesos.log          # Log de accesos
└── temp\
    └── documentos\          # Documentos generados temporalmente
```

#### Cliente Local
```
C:\CONDOR_LOCAL\
├── CONDOR.accde             # Copia local del frontend
├── config_local.ini         # Configuración local
├── cache\
│   └── expedientes\         # Cache de expedientes consultados
└── logs\
    └── cliente.log          # Log del cliente
```

---

## 2. Procedimientos de Instalación

### 2.1 Instalación del Servidor

#### Prerrequisitos
- Windows Server 2016 o superior
- Microsoft Access Runtime 2016 o superior
- Permisos de administrador en el servidor
- Carpeta compartida configurada con permisos apropiados

#### Pasos de Instalación

1. **Crear Estructura de Directorios**
   ```batch
   mkdir \\SERVIDOR\CONDOR
   mkdir \\SERVIDOR\CONDOR\config
   mkdir \\SERVIDOR\CONDOR\plantillas
   mkdir \\SERVIDOR\CONDOR\logs
   mkdir \\SERVIDOR\CONDOR\temp
   mkdir \\SERVIDOR\CONDOR\temp\documentos
   ```

2. **Copiar Archivos del Sistema**
   ```batch
   copy CONDOR.accde \\SERVIDOR\CONDOR\
   copy CONDOR_datos.accdb \\SERVIDOR\CONDOR\
   copy condor_cli.vbs \\SERVIDOR\CONDOR\
   ```

3. **Configurar Permisos de Carpeta**
   ```
   \\SERVIDOR\CONDOR\
   ├── Usuarios CONDOR: Lectura y Ejecución
   ├── Administradores: Control Total
   └── Sistema: Control Total
   
   \\SERVIDOR\CONDOR\logs\
   ├── Usuarios CONDOR: Modificar
   
   \\SERVIDOR\CONDOR\temp\
   ├── Usuarios CONDOR: Modificar
   ```

4. **Configurar Base de Datos Backend**
   - Abrir CONDOR_datos.accdb como administrador
   - Ejecutar script de inicialización de tablas
   - Configurar usuarios y permisos de base de datos
   - Crear índices para optimización

### 2.2 Configuración del Cliente

#### Instalación Automática via Lanzadera

1. **Crear Acceso Directo en Escritorio del Usuario**
   ```
   Destino: \\SERVIDOR\CONDOR\condor_cli.vbs
   Nombre: Sistema CONDOR
   Icono: \\SERVIDOR\CONDOR\condor.ico
   ```

2. **Primera Ejecución**
   - El script `condor_cli.vbs` detecta primera ejecución
   - Crea estructura local en `C:\CONDOR_LOCAL`
   - Copia `CONDOR.accde` localmente
   - Configura variables de entorno
   - Registra cliente en el servidor

#### Instalación Manual (Modo Desarrollo)

1. **Crear Carpeta Local**
   ```batch
   mkdir C:\CONDOR_LOCAL
   mkdir C:\CONDOR_LOCAL\cache
   mkdir C:\CONDOR_LOCAL\cache\expedientes
   mkdir C:\CONDOR_LOCAL\logs
   ```

2. **Copiar Archivos**
   ```batch
   copy \\SERVIDOR\CONDOR\CONDOR.accde C:\CONDOR_LOCAL\
   ```

3. **Configurar Modo Desarrollo**
   ```ini
   ; Archivo: C:\CONDOR_LOCAL\config_local.ini
   [DESARROLLO]
   DEV_MODE=true
   SERVIDOR_DESARROLLO=\\SERVIDOR-DEV\CONDOR
   LOG_LEVEL=DEBUG
   MOCK_EXPEDIENTES=true
   ```

---

## 3. Archivos de Configuración

### 3.1 Configuración Principal (configuracion.ini)

```ini
; Archivo: \\SERVIDOR\CONDOR\config\configuracion.ini

[SISTEMA]
VERSION=1.0.0
NOMBRE_APLICACION=Sistema CONDOR
FECHA_VERSION=2024-12-20
MODO_MANTENIMIENTO=false

[BASE_DATOS]
RUTA_BACKEND=\\SERVIDOR\CONDOR\CONDOR_datos.accdb
TIMEOUT_CONEXION=30
MAX_CONEXIONES_SIMULTANEAS=50
BACKUP_AUTOMATICO=true
INTERVALO_BACKUP_HORAS=24

[SEGURIDAD]
AUTENTICACION_WINDOWS=true
TIMEOUT_SESION_MINUTOS=120
LOG_ACCESOS=true
ENCRIPTAR_COMUNICACION=false

[INTEGRACION]
SERVICIO_EXPEDIENTES_URL=http://servidor-expedientes/api/
SERVICIO_RAC_URL=http://servidor-rac/api/
TIMEOUT_SERVICIOS_SEGUNDOS=30
REINTENTOS_MAXIMOS=3

[DOCUMENTOS]
RUTA_PLANTILLAS=\\SERVIDOR\CONDOR\plantillas\
RUTA_TEMP=\\SERVIDOR\CONDOR\temp\documentos\
FORMATO_SALIDA=DOCX
ELIMINAR_TEMP_DIAS=7

[NOTIFICACIONES]
SMTP_SERVIDOR=smtp.empresa.com
SMTP_PUERTO=587
SMTP_USUARIO=condor@empresa.com
SMTP_SSL=true
REMITENTE_DEFECTO=Sistema CONDOR <condor@empresa.com>

[LOGGING]
NIVEL_LOG=INFO
MAX_TAMAÑO_LOG_MB=10
MAX_ARCHIVOS_LOG=5
LOG_ROTACION=true
```

### 3.2 Configuración de Entornos (entornos.ini)

```ini
; Archivo: \\SERVIDOR\CONDOR\config\entornos.ini

[PRODUCCION]
ACTIVO=true
RUTA_BACKEND=\\SERVIDOR-PROD\CONDOR\CONDOR_datos.accdb
SERVICIO_EXPEDIENTES=http://prod-expedientes/api/
SERVICIO_RAC=http://prod-rac/api/
LOG_LEVEL=INFO
DEBUG_MODE=false

[DESARROLLO]
ACTIVO=false
RUTA_BACKEND=\\SERVIDOR-DEV\CONDOR\CONDOR_datos.accdb
SERVICIO_EXPEDIENTES=http://dev-expedientes/api/
SERVICIO_RAC=http://dev-rac/api/
LOG_LEVEL=DEBUG
DEBUG_MODE=true
MOCK_SERVICIOS=true

[TESTING]
ACTIVO=false
RUTA_BACKEND=C:\CONDOR_TEST\CONDOR_datos.accdb
SERVICIO_EXPEDIENTES=MOCK
SERVICIO_RAC=MOCK
LOG_LEVEL=DEBUG
DEBUG_MODE=true
MOCK_SERVICIOS=true
DATA_RESET_ON_START=true
```

### 3.3 Configuración Local del Cliente (config_local.ini)

```ini
; Archivo: C:\CONDOR_LOCAL\config_local.ini

[CLIENTE]
ID_CLIENTE=PC-{NOMBRE_EQUIPO}-{USUARIO}
ULTIMA_ACTUALIZACION=2024-12-20 10:30:00
VERSION_LOCAL=1.0.0
MODO_OFFLINE=false

[CACHE]
HABILITAR_CACHE=true
TAMAÑO_MAX_CACHE_MB=100
TIEMPO_EXPIRACION_HORAS=24
LIMPIEZA_AUTOMATICA=true

[DESARROLLO]
DEV_MODE=false
SERVIDOR_DESARROLLO=\\SERVIDOR-DEV\CONDOR
LOG_LEVEL=INFO
MOCK_EXPEDIENTES=false
MOCK_RAC=false

[UI]
TEMA=Claro
IDIOMA=ES
TAMAÑO_FUENTE=10
MOSTRAR_TOOLTIPS=true
ANIMACIONES=true
```

---

## 4. Sistema de Lanzadera (condor_cli.vbs)

### 4.1 Funcionalidades del Script

```vbscript
' Archivo: condor_cli.vbs
' Responsabilidades:
' - Verificar actualizaciones del frontend
' - Gestionar instalación local
' - Configurar variables de entorno
' - Lanzar aplicación con parámetros correctos
' - Manejar errores de conexión

' Estructura principal:
Sub Main()
    ' 1. Verificar prerrequisitos
    If Not VerificarPrerrequisitos() Then Exit Sub
    
    ' 2. Verificar/crear instalación local
    If Not VerificarInstalacionLocal() Then
        If Not CrearInstalacionLocal() Then Exit Sub
    End If
    
    ' 3. Verificar actualizaciones
    If HayActualizacionDisponible() Then
        If ActualizarAplicacion() Then
            MsgBox "Aplicación actualizada exitosamente"
        Else
            MsgBox "Error al actualizar. Contacte al administrador"
            Exit Sub
        End If
    End If
    
    ' 4. Configurar entorno
    ConfigurarEntorno
    
    ' 5. Lanzar aplicación
    LanzarAplicacion
End Sub
```

### 4.2 Verificación de Actualizaciones

```vbscript
Function HayActualizacionDisponible()
    Dim fso, archivoServidor, archivoLocal
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Comparar fechas de modificación
    Set archivoServidor = fso.GetFile(RUTA_SERVIDOR & "\CONDOR.accde")
    
    If fso.FileExists(RUTA_LOCAL & "\CONDOR.accde") Then
        Set archivoLocal = fso.GetFile(RUTA_LOCAL & "\CONDOR.accde")
        HayActualizacionDisponible = (archivoServidor.DateLastModified > archivoLocal.DateLastModified)
    Else
        HayActualizacionDisponible = True
    End If
End Function

Function ActualizarAplicacion()
    On Error Resume Next
    
    ' Cerrar instancias existentes
    CerrarInstanciasAccess
    
    ' Copiar nuevo archivo
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    fso.CopyFile RUTA_SERVIDOR & "\CONDOR.accde", RUTA_LOCAL & "\CONDOR.accde", True
    
    If Err.Number = 0 Then
        ActualizarAplicacion = True
        RegistrarActualizacion
    Else
        ActualizarAplicacion = False
        LogError "Error al actualizar: " & Err.Description
    End If
End Function
```

---

## 5. Gestión de Entornos

### 5.1 Detección Automática de Entorno

```vba
' Módulo: ModConfiguracion.bas
' Función para detectar entorno actual

Public Function DetectarEntornoActual() As String
    Dim entorno As String
    
    ' Verificar variable de entorno del sistema
    entorno = Environ("CONDOR_ENTORNO")
    
    If Len(entorno) = 0 Then
        ' Detectar por ubicación del servidor
        If InStr(UCase(GetRutaServidor()), "PROD") > 0 Then
            entorno = "PRODUCCION"
        ElseIf InStr(UCase(GetRutaServidor()), "DEV") > 0 Then
            entorno = "DESARROLLO"
        ElseIf InStr(UCase(GetRutaServidor()), "TEST") > 0 Then
            entorno = "TESTING"
        Else
            entorno = "PRODUCCION" ' Por defecto
        End If
    End If
    
    DetectarEntornoActual = UCase(entorno)
End Function

Public Sub ConfigurarEntorno(entorno As String)
    Dim rutaConfig As String
    rutaConfig = GetRutaServidor() & "\config\entornos.ini"
    
    ' Cargar configuración específica del entorno
    CargarConfiguracionEntorno rutaConfig, entorno
    
    ' Configurar variables globales
    ConfigurarVariablesGlobales entorno
    
    ' Configurar logging según entorno
    ConfigurarLogging entorno
End Sub
```

### 5.2 Configuración por Entorno

#### Producción
- Logging mínimo (solo errores e información crítica)
- Sin modo debug
- Conexiones reales a servicios externos
- Validaciones completas habilitadas
- Cache optimizado para rendimiento

#### Desarrollo
- Logging detallado (debug)
- Modo debug habilitado
- Posibilidad de usar mocks
- Validaciones relajadas para testing
- Recarga automática de configuración

#### Testing
- Logging completo para análisis
- Datos de prueba automáticos
- Mocks habilitados por defecto
- Reset de datos al iniciar
- Métricas de rendimiento

---

## 6. Monitoreo y Mantenimiento

### 6.1 Logs del Sistema

#### Estructura de Logs
```
\\SERVIDOR\CONDOR\logs\
├── sistema_YYYYMMDD.log      # Log general del sistema
├── errores_YYYYMMDD.log      # Log específico de errores
├── accesos_YYYYMMDD.log      # Log de accesos de usuarios
├── performance_YYYYMMDD.log  # Log de métricas de rendimiento
└── integracion_YYYYMMDD.log  # Log de integraciones externas
```

#### Formato de Log
```
[YYYY-MM-DD HH:MM:SS] [NIVEL] [MODULO] [USUARIO] [ACCION] - MENSAJE

Ejemplos:
[2024-12-20 10:30:15] [INFO] [SolicitudService] [juan.perez] [CrearSolicitud] - Solicitud SOL-2024-001 creada exitosamente
[2024-12-20 10:31:22] [ERROR] [ExpedienteService] [maria.garcia] [ConsultarExpediente] - Error de conexión: Timeout al consultar EXP-2024-100
[2024-12-20 10:32:05] [DEBUG] [DocumentoService] [admin] [GenerarDocumento] - Plantilla PC cargada en 250ms
```

### 6.2 Monitoreo Automático

#### Script de Monitoreo (monitor_condor.vbs)
```vbscript
' Verificaciones automáticas cada 15 minutos

Sub VerificarEstadoSistema()
    ' 1. Verificar accesibilidad del servidor
    If Not VerificarConexionServidor() Then
        EnviarAlerta "Servidor CONDOR no accesible"
    End If
    
    ' 2. Verificar tamaño de logs
    If TamañoLogsExcesivo() Then
        RotarLogs
    End If
    
    ' 3. Verificar espacio en disco
    If EspacioDiscoInsuficiente() Then
        EnviarAlerta "Espacio en disco insuficiente en servidor CONDOR"
    End If
    
    ' 4. Verificar integridad de base de datos
    If Not VerificarIntegridadBD() Then
        EnviarAlerta "Posible corrupción en base de datos CONDOR"
    End If
End Sub
```

### 6.3 Backup y Recuperación

#### Backup Automático Diario
```batch
@echo off
REM Script: backup_condor.bat
REM Ejecutar diariamente a las 02:00 AM

set FECHA=%date:~6,4%%date:~3,2%%date:~0,2%
set RUTA_BACKUP=\\SERVIDOR-BACKUP\CONDOR\%FECHA%
set RUTA_ORIGEN=\\SERVIDOR\CONDOR

REM Crear carpeta de backup
mkdir "%RUTA_BACKUP%"

REM Backup de base de datos
copy "%RUTA_ORIGEN%\CONDOR_datos.accdb" "%RUTA_BACKUP%\CONDOR_datos_%FECHA%.accdb"

REM Backup de configuración
xcopy "%RUTA_ORIGEN%\config" "%RUTA_BACKUP%\config\" /E /I

REM Backup de plantillas
xcopy "%RUTA_ORIGEN%\plantillas" "%RUTA_BACKUP%\plantillas\" /E /I

REM Comprimir logs del mes anterior
7z a "%RUTA_BACKUP%\logs_%FECHA%.7z" "%RUTA_ORIGEN%\logs\*" -mx9

REM Limpiar backups antiguos (mantener 30 días)
forfiles /p "\\SERVIDOR-BACKUP\CONDOR" /m *.* /d -30 /c "cmd /c rmdir /s /q @path"

echo Backup completado: %date% %time% >> "%RUTA_ORIGEN%\logs\backup.log"
```

#### Procedimiento de Recuperación

1. **Recuperación Completa**
   ```batch
   REM Detener acceso al sistema
   net share CONDOR /delete
   
   REM Restaurar archivos desde backup
   copy "\\SERVIDOR-BACKUP\CONDOR\20241220\CONDOR_datos_20241220.accdb" "\\SERVIDOR\CONDOR\CONDOR_datos.accdb"
   xcopy "\\SERVIDOR-BACKUP\CONDOR\20241220\config\" "\\SERVIDOR\CONDOR\config\" /E /Y
   
   REM Verificar integridad
   msaccess.exe "\\SERVIDOR\CONDOR\CONDOR_datos.accdb" /compact
   
   REM Reactivar acceso
   net share CONDOR="\\SERVIDOR\CONDOR" /grant:everyone,full
   ```

2. **Recuperación de Solo Datos**
   ```sql
   -- Importar datos desde backup usando Access
   -- 1. Abrir CONDOR_datos.accdb actual
   -- 2. Archivo > Datos externos > Access
   -- 3. Seleccionar backup: CONDOR_datos_YYYYMMDD.accdb
   -- 4. Importar tablas específicas según necesidad
   ```

---

## 7. Troubleshooting

### 7.1 Problemas Comunes

#### Error: "No se puede acceder al servidor"
**Síntomas:**
- Aplicación no inicia
- Mensaje de error de conexión
- Timeout al acceder a archivos

**Soluciones:**
1. Verificar conectividad de red: `ping SERVIDOR`
2. Verificar permisos de carpeta compartida
3. Verificar que el servicio de archivos esté activo
4. Comprobar firewall y antivirus

#### Error: "Base de datos bloqueada"
**Síntomas:**
- No se pueden guardar cambios
- Mensaje "The database has been placed in a state by user 'Admin' on machine..."

**Soluciones:**
1. Cerrar todas las instancias de Access en todos los clientes
2. Eliminar archivo `.laccdb` en el servidor
3. Verificar que no haya procesos colgados de Access
4. Reiniciar servicio de archivos si es necesario

#### Error: "Plantilla no encontrada"
**Síntomas:**
- Error al generar documentos
- Mensaje sobre archivo de plantilla no accesible

**Soluciones:**
1. Verificar existencia de archivos en `\\SERVIDOR\CONDOR\plantillas\`
2. Comprobar permisos de lectura en carpeta de plantillas
3. Verificar configuración en `configuracion.ini`
4. Restaurar plantillas desde backup si es necesario

### 7.2 Herramientas de Diagnóstico

#### Script de Diagnóstico (diagnostico_condor.vbs)
```vbscript
Sub EjecutarDiagnostico()
    Dim reporte As String
    reporte = "=== DIAGNÓSTICO SISTEMA CONDOR ===" & vbCrLf & vbCrLf
    
    ' Verificar conectividad
    reporte = reporte & "1. CONECTIVIDAD" & vbCrLf
    reporte = reporte & "   Servidor accesible: " & IIf(PingServidor(), "SÍ", "NO") & vbCrLf
    reporte = reporte & "   Carpeta compartida: " & IIf(AccesoCarpeta(), "SÍ", "NO") & vbCrLf
    
    ' Verificar archivos
    reporte = reporte & vbCrLf & "2. ARCHIVOS" & vbCrLf
    reporte = reporte & "   Frontend (CONDOR.accde): " & IIf(ExisteArchivo("CONDOR.accde"), "SÍ", "NO") & vbCrLf
    reporte = reporte & "   Backend (CONDOR_datos.accdb): " & IIf(ExisteArchivo("CONDOR_datos.accdb"), "SÍ", "NO") & vbCrLf
    reporte = reporte & "   Lanzadera (condor_cli.vbs): " & IIf(ExisteArchivo("condor_cli.vbs"), "SÍ", "NO") & vbCrLf
    
    ' Verificar configuración
    reporte = reporte & vbCrLf & "3. CONFIGURACIÓN" & vbCrLf
    reporte = reporte & "   Archivo configuracion.ini: " & IIf(ExisteArchivo("config\configuracion.ini"), "SÍ", "NO") & vbCrLf
    reporte = reporte & "   Entorno detectado: " & DetectarEntorno() & vbCrLf
    
    ' Verificar servicios externos
    reporte = reporte & vbCrLf & "4. SERVICIOS EXTERNOS" & vbCrLf
    reporte = reporte & "   Servicio Expedientes: " & VerificarServicioExpedientes() & vbCrLf
    reporte = reporte & "   Servicio RAC: " & VerificarServicioRAC() & vbCrLf
    
    ' Mostrar reporte
    MsgBox reporte, vbInformation, "Diagnóstico CONDOR"
    
    ' Guardar reporte en archivo
    GuardarReporte reporte
End Sub
```

---

## 8. Procedimientos de Actualización

### 8.1 Actualización del Frontend

1. **Preparación**
   - Notificar a usuarios sobre mantenimiento
   - Realizar backup completo
   - Compilar nueva versión de CONDOR.accde

2. **Despliegue**
   ```batch
   REM Detener acceso temporal
   ren "\\SERVIDOR\CONDOR\CONDOR.accde" "CONDOR.accde.old"
   
   REM Copiar nueva versión
   copy "CONDOR_v1.1.0.accde" "\\SERVIDOR\CONDOR\CONDOR.accde"
   
   REM Verificar integridad
   if errorlevel 1 (
       ren "\\SERVIDOR\CONDOR\CONDOR.accde.old" "CONDOR.accde"
       echo "Error en actualización - Rollback ejecutado"
   ) else (
       del "\\SERVIDOR\CONDOR\CONDOR.accde.old"
       echo "Actualización completada exitosamente"
   )
   ```

3. **Verificación Post-Actualización**
   - Probar funcionalidades críticas
   - Verificar logs por errores
   - Confirmar con usuarios piloto

### 8.2 Actualización de Base de Datos

1. **Scripts de Migración**
   ```sql
   -- Archivo: migracion_v1.0_a_v1.1.sql
   
   -- Agregar nueva columna
   ALTER TABLE Tb_Solicitudes ADD COLUMN FechaAprobacion DATETIME;
   
   -- Crear nueva tabla
   CREATE TABLE TbAuditoria (
       AuditoriaId AUTOINCREMENT PRIMARY KEY,
       TablaAfectada TEXT(50),
       Operacion TEXT(10),
       UsuarioOperacion TEXT(50),
       FechaOperacion DATETIME,
       DatosAnteriores MEMO,
       DatosNuevos MEMO
   );
   
   -- Actualizar versión
   UPDATE TbConfiguracion SET Valor = '1.1.0' WHERE Clave = 'VERSION_BD';
   ```

2. **Procedimiento de Migración**
   ```vba
   Public Sub EjecutarMigracion(versionOrigen As String, versionDestino As String)
       On Error GoTo ErrorHandler
       
       ' Verificar versión actual
       If GetVersionBD() <> versionOrigen Then
           MsgBox "Versión de BD no coincide. Migración cancelada."
           Exit Sub
       End If
       
       ' Backup automático antes de migración
       If Not CrearBackupMigracion() Then
           MsgBox "Error al crear backup. Migración cancelada."
           Exit Sub
       End If
       
       ' Ejecutar scripts de migración
       EjecutarScriptMigracion versionOrigen, versionDestino
       
       ' Verificar integridad post-migración
       If VerificarIntegridadPostMigracion() Then
           MsgBox "Migración completada exitosamente"
       Else
           RestaurarBackupMigracion
           MsgBox "Error en migración. Sistema restaurado."
       End If
       
       Exit Sub
       
   ErrorHandler:
       RestaurarBackupMigracion
       MsgBox "Error crítico en migración: " & Err.Description
   End Sub
   ```

---

## 9. Checklist de Despliegue

### 9.1 Pre-Despliegue

- [ ] Backup completo realizado y verificado
- [ ] Entorno de testing validado
- [ ] Scripts de migración probados
- [ ] Usuarios notificados sobre mantenimiento
- [ ] Ventana de mantenimiento programada
- [ ] Plan de rollback preparado
- [ ] Herramientas de monitoreo configuradas

### 9.2 Durante el Despliegue

- [ ] Acceso de usuarios restringido
- [ ] Archivos copiados correctamente
- [ ] Migraciones de BD ejecutadas
- [ ] Configuraciones actualizadas
- [ ] Permisos verificados
- [ ] Servicios reiniciados si es necesario

### 9.3 Post-Despliegue

- [ ] Funcionalidades críticas probadas
- [ ] Logs revisados por errores
- [ ] Rendimiento verificado
- [ ] Usuarios piloto confirmados
- [ ] Documentación actualizada
- [ ] Acceso de usuarios restaurado
- [ ] Monitoreo activo por 24 horas

---

*Documento generado según la Especificación Funcional y Arquitectura CONDOR*  
*Versión: 1.0*  
*Fecha: Diciembre 2024*