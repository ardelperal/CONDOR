# CONDOR

## Resumen
CONODOR es una aplicación para gestionar el ciclo de vida de solicitudes de cambio, desviación o concesión en expedientes de contratos públicos. Está desarrollada en Microsoft Access con VBA y orientada a usuarios de Calidad y Técnico.

## Arquitectura

CONDOR sigue una arquitectura de 3 capas con integración a sistema existente:

- **Capa de Presentación**: Formularios de Access
- **Capa de Negocio**: Módulos VBA con lógica de negocio + ExpedienteService (interfaz con app existente)
- **Capa de Datos**: Base de datos Access + Integración con aplicación de expedientes existente por IDExpediente

### 🏗️ Patrones de Diseño Implementados

El proyecto CONDOR implementa una arquitectura modular basada en patrones de diseño:

- **Patrón Repository**: Para acceso a datos
- **Patrón Factory**: Para creación de servicios
- **Inyección de Dependencias**: Para desacoplamiento
- **Interfaces**: Para contratos bien definidos
- **Separación de Responsabilidades**: Cada clase tiene una función específica
- **Gestión Segura de Conexiones**: Uso de OpenDatabase con cierre explícito
- **Configuración Centralizada**: Acceso a backend a través de modConfig.GetDataPath()

### Integración con Sistema Existente
CONDOR se conecta con la aplicación de expedientes existente para obtener:
- Nemotécnico del expediente
- Responsable de calidad
- Jefe de proyecto
- Información de contratista principal

### Características Técnicas
- Despliegue centralizado mediante una lanzadera.
- Front-end y back-end separados.
- Actualización automática de versiones.
- Funciona en modo oficina (producción) y local (desarrollo/test) sin cambiar el código.

## Gestión de Entornos (Local vs. Remoto)

CONDOR implementa un sistema avanzado de gestión de entornos que permite a los desarrolladores trabajar tanto con datos locales como remotos de manera flexible y controlada.

### Diferencia entre Modo de Compilación y Entorno de Ejecución

**Modo de Compilación (DEV_MODE):**
- Es una constante de compilación que determina si la aplicación se compila en modo desarrollo o producción
- Se define mediante directivas de compilación condicional (`#If DEV_MODE Then`)
- Controla qué código se incluye en la compilación final

**Entorno de Ejecución (Local vs. Remoto):**
- Determina qué rutas de bases de datos y recursos utiliza la aplicación en tiempo de ejecución
- **Local**: Utiliza rutas del directorio de desarrollo (`C:\Proyectos\CONDOR\...`)
- **Remoto**: Utiliza rutas de la red corporativa (`\\datoste\aplicaciones_dys\...`)

### Constante ENTORNO_FORZADO

Dentro del módulo `modConfig.bas` existe una constante privada `ENTORNO_FORZADO` que permite a los desarrolladores forzar un entorno específico independientemente del modo de compilación:

```vba
Private Enum E_EnvironmentOverride
    ForzarNinguno = 0 ' Elige automáticamente basado en DEV_MODE
    ForzarLocal = 1   ' Fuerza el uso de rutas locales
    ForzarRemoto = 2  ' Fuerza el uso de rutas remotas
End Enum
Private Const ENTORNO_FORZADO As E_EnvironmentOverride = ForzarNinguno
```

### Lógica de Decisión de Entorno

La función `InitializeEnvironment()` utiliza la siguiente lógica para determinar el entorno:

```vba
Select Case ENTORNO_FORZADO
    Case ForzarLocal
        usarRutasLocales = True
        g_AppConfig.EntornoActivo = "Local (Forzado)"
    Case ForzarRemoto
        usarRutasLocales = False
        g_AppConfig.EntornoActivo = "Remoto (Forzado)"
    Case ForzarNinguno
        ' Comportamiento por defecto: depende del modo de compilación
        usarRutasLocales = IsDevelopmentMode()
        If usarRutasLocales Then
            g_AppConfig.EntornoActivo = "Local (DEV_MODE)"
        Else
            g_AppConfig.EntornoActivo = "Remoto (Producción)"
        End If
End Select
```

### Cómo Cambiar el Entorno para Desarrollo

Para cambiar el entorno de ejecución durante el desarrollo:

1. **Abrir** el archivo `src/modConfig.bas`
2. **Localizar** la línea con `Private Const ENTORNO_FORZADO`
3. **Cambiar** el valor según necesidad:
   - `ForzarLocal`: Para trabajar con datos locales de desarrollo
   - `ForzarRemoto`: Para depurar con datos reales de la red
   - `ForzarNinguno`: Para usar el comportamiento automático
4. **Actualizar** el proyecto: `cscript condor_cli.vbs update` (o `rebuild` si hay problemas)

### Casos de Uso Típicos

**Desarrollo Normal:**
```vba
Private Const ENTORNO_FORZADO As E_EnvironmentOverride = ForzarNinguno
```
- Usa datos locales en modo desarrollo
- Usa datos remotos en producción

**Debug con Datos Reales:**
```vba
Private Const ENTORNO_FORZADO As E_EnvironmentOverride = ForzarRemoto
```
- Permite depurar desde el entorno de desarrollo usando las bases de datos de la red
- Esencial para reproducir errores que solo ocurren con datos reales

**Pruebas Aisladas:**
```vba
Private Const ENTORNO_FORZADO As E_EnvironmentOverride = ForzarLocal
```
- Garantiza el uso de datos locales incluso en compilaciones de producción
- Útil para pruebas controladas

### ⚠️ Importante: Antes del Commit Final

**SIEMPRE** devolver la constante a `ForzarNinguno` antes de hacer commit:

```vba
Private Const ENTORNO_FORZADO As E_EnvironmentOverride = ForzarNinguno
```

Esto garantiza que:
- El comportamiento por defecto se mantenga en el repositorio
- Otros desarrolladores no hereden configuraciones específicas
- La aplicación funcione correctamente en todos los entornos

### Verificación del Entorno Activo

Puede verificarse qué entorno está activo mediante:
```vba
Debug.Print GetActiveEnvironment()
```

Esto mostrará valores como:
- "Local (DEV_MODE)"
- "Remoto (Producción)"
- "Local (Forzado)"
- "Remoto (Forzado)"

## Gestión de Usuarios y Roles
- Login integrado con el sistema central.
- Roles: Calidad, Técnico, Administrador, y actores externos (solo reciben documentos).

## Flujo de Trabajo
- Fase interna: Preparación y revisión de solicitudes por Calidad y Técnico.
- Fase externa: Generación y envío de documentos a actores externos, recepción y cierre.

## Arquitectura de Código
- Separación en capas: Presentación, Negocio, Acceso a Datos y Servicios Externos.
- Uso de interfaces para facilitar tests unitarios.
- Sistema de manejo de errores centralizado (`modErrorHandler.bas`) que registra todos los errores en la tabla `Tb_Log_Errores` de la base de datos.

### Servicios y Factories Adicionales

**Sistema de Logging de Operaciones:**
Se ha implementado un nuevo sistema para registrar las operaciones importantes del usuario y del sistema, proporcionando trazabilidad y soporte para auditorías.
- **Interfaz:** `IOperationLogger.cls`
- **Implementación:** `COperationLogger.cls` (registra operaciones en `Tb_Operaciones_Log`)
- **Mock:** `CMockOperationLogger.cls` (para pruebas)
- **Factory:** `modOperationLoggerFactory.bas` (gestiona la creación de instancias del logger)

**Factory de Configuración:**
Se ha añadido un factory específico para la gestión de servicios de configuración, siguiendo el patrón de inyección de dependencias.
- **Factory:** `modConfigFactory.bas` (proporciona instancias de `IConfig`)

**Servicio de Documentos:**
Implementación completa del servicio de generación y lectura de documentos Word con arquitectura de inyección de dependencias.
- **Interfaz:** `IDocumentService.cls` (contrato para operaciones de documentos)
- **Implementación:** `CDocumentService.cls` (lógica principal de generación y lectura)
- **Mock:** `CMockDocumentService.cls` (para pruebas unitarias aisladas)
- **Factory:** `modDocumentServiceFactory.bas` (inyección de IConfig, ISolicitudRepository, IOperationLogger, IWordManager)

**Gestión de Word:**
Abstracción completa para el manejo de documentos Word, eliminando dependencias directas de Word.Application.
- **Interfaz:** `IWordManager.cls` (contrato para operaciones con Word)
- **Implementación:** `CWordManager.cls` (encapsula Word.Application)
- **Mock:** `CMockWordManager.cls` (simula operaciones de Word para pruebas)

**Características del DocumentService:**
- **Aislamiento Total:** Sin dependencias directas de Word o base de datos en las pruebas
- **Operaciones:** GenerarDocumento (plantilla → documento final), LeerDocumento (documento → base de datos)
- **Pruebas Unitarias:** Test_DocumentService.bas con cobertura completa usando mocks
- **Manejo de Errores:** Integración completa con modErrorHandler y IOperationLogger

## Estructura de Datos
- Tablas principales: Expedientes, Solicitudes, Datos específicos y Mapeo de campos.

## Herramienta CLI de Desarrollo

CONDOR incluye una herramienta de línea de comandos (`condor_cli.vbs`) que facilita el desarrollo y mantenimiento del código VBA.

### Comandos Disponibles

#### Actualización Selectiva de Módulos (Recomendado)
```bash
# Actualizar un solo módulo
cscript condor_cli.vbs update CAuthService

# Actualizar múltiples módulos específicos
cscript condor_cli.vbs update CAuthService,modUtils,CConfig

# Sincronización automática optimizada (solo abre BD si hay cambios)
cscript condor_cli.vbs update
```
- **Comando optimizado** para sincronización discrecional de archivos
- **Optimización de rendimiento**: El comando `update` sin parámetros verifica cambios antes de abrir la base de datos
- **Conversión automática**: Incluye conversión UTF-8 a ANSI para soporte completo de caracteres especiales (ñ, tildes)
- Permite actualizar módulos específicos sin afectar el resto del proyecto
- Solo procesa los módulos especificados, no toda la base de datos
- Elimina e importa únicamente los módulos indicados usando `DoCmd.LoadFromText`
- **Sintaxis**: Los nombres de módulos se separan con comas (sin espacios)
- **Nota**: No incluir extensiones (.bas/.cls) en los nombres
- **Ventaja**: Mucho más rápido que `rebuild` para cambios específicos
- **Flexibilidad**: Ideal para desarrollo iterativo y correcciones puntuales

#### Exportación de Módulos
```bash
cscript condor_cli.vbs export
```
- Exporta todos los módulos VBA desde la base de datos Access hacia archivos `.bas` en el directorio `src/`
- Útil para sincronizar cambios realizados directamente en Access hacia el control de versiones
- Mantiene la estructura del código y comentarios

#### Reconstrucción Completa del Proyecto
```bash
cscript condor_cli.vbs rebuild
```
- Elimina todos los módulos VBA existentes de la base de datos Access
- Importa todos los archivos `.bas` del directorio `src/` hacia la base de datos Access
- Compila automáticamente los módulos después de la importación
- Garantiza un estado 100% limpio y compilado
- **Usar solo cuando `update` no sea suficiente** (problemas de sincronización graves)
- Muestra advertencias de compilación si las hay
- **Conversión automática**: Incluye conversión UTF-8 a ANSI para soporte completo de caracteres especiales (ñ, tildes)

#### Ayuda de Comandos
```bash
cscript condor_cli.vbs help
```
- Muestra una lista detallada de todos los comandos disponibles y su descripción.

#### Flujo de Trabajo de Verificación Manual (Post-push de la IA)

Después de que el agente autónomo complete una tarea y suba los cambios, el supervisor humano debe realizar el siguiente proceso de control de calidad para validar el trabajo:

**Paso 1: Sincronizar el Repositorio Local**
- Abrir una terminal en la raíz del proyecto.
- Ejecutar `git pull` para descargar los últimos cambios.

**Paso 2: Actualizar la Base de Datos de Desarrollo**
- **Opción A (Recomendada)**: Actualización selectiva si conoces los módulos modificados:
```bash
cscript //nologo condor_cli.vbs update CAuthService,modUtils
```
- **Opción B**: Actualización completa (más lenta pero segura):
```bash
cscript //nologo condor_cli.vbs update
```
- **Opción C**: Solo usar si hay problemas graves de sincronización:
```bash
cscript //nologo condor_cli.vbs rebuild
```

**Paso 3: Verificar la Compilación (Paso de Calidad Crítico)**
- Abrir el fichero `CONDOR.accdb` de la carpeta de desarrollo.
- Abrir el editor de VBA (`Alt + F11`).
- Ir al menú **Depurar > Compilar [Nombre del Proyecto]**.
- Si aparece un error de compilación, la verificación falla. Se debe notificar a la IA con una captura de pantalla del error. No se debe continuar al siguiente paso.
- Si no ocurre nada, la compilación es exitosa.

**Paso 4: Ejecutar la Suite de Pruebas Automatizadas**
- Con el editor de VBA abierto, mostrar la Ventana Inmediato (`Ctrl + G`).
- Para ejecutar todas las pruebas, escribir en la Ventana Inmediato y pulsar Enter:
```vb
EJECUTAR_TODAS_LAS_PRUEBAS
```

**Paso 5: Analizar los Resultados y Notificar**
- Revisar el informe de pruebas que aparece en la Ventana Inmediato.
- Si todas las pruebas pasan, notificar a la IA que el trabajo ha sido validado y que puede proceder con la siguiente tarea del `PLAN_DE_ACCION.MD`.
- Si alguna prueba falla, copiar el log de error y proporcionárselo a la IA para que inicie un nuevo ciclo de depuración.

### Conversión Automática de Codificación

La herramienta CLI maneja automáticamente la conversión entre las diferentes codificaciones de caracteres utilizadas por VS Code y Access VBA:

#### Problema de Codificación
- **VS Code**: Utiliza codificación UTF-8 (estándar moderno) que puede representar cualquier carácter
- **Access VBA**: Utiliza codificación ANSI/Windows-1252 (estándar legacy) con conjunto limitado de caracteres
- **Conflicto**: Los caracteres especiales (tildes, eñes) se representan de forma diferente en cada codificación

#### Solución Automática
La CLI actúa como traductor inteligente en ambas direcciones:

**Durante la Exportación (Access → src/):**
- Lee módulos VBA desde Access (formato ANSI interno)
- Convierte automáticamente a UTF-8 al escribir archivos en `/src`
- Preserva todos los caracteres especiales correctamente

**Durante la Importación (src/ → Access):**
- Lee archivos UTF-8 desde el directorio `/src`
- Convierte automáticamente a ANSI antes de importar a Access
- Elimina metadatos "Attribute VB_" durante el proceso
- Garantiza compatibilidad total con el editor VBA

#### Beneficios
- **Transparente**: Los desarrolladores no necesitan preocuparse por la codificación
- **Preserva Caracteres**: Mantiene tildes, eñes y caracteres especiales intactos
- **Sin Mojibake**: Evita caracteres corruptos como "Ã¡" o "�"
- **Código Limpio**: El código fuente nunca se modifica, solo se traduce la codificación

### Sistema de Pruebas

CONDOR incluye un motor de pruebas unitarias integrado que permite validar la funcionalidad del código VBA.

#### Estructura de Pruebas

- **Motor de Ejecución**: `modTestRunner.bas` - Ejecuta todas las pruebas registradas aplicando el Principio de Responsabilidad Única (SRP)
- **Generador de Informes**: `CTestReporter.cls` - Clase especializada en generar informes consolidados de resultados
- **Gestión de Resultados**: `CTestSuiteResult.cls` - Clase que encapsula los resultados de cada suite de pruebas
- **Módulos de Prueba**: Archivos que contienen funciones de prueba (ej: `Test_Ejemplo.bas`, `Test_Integracion.bas`, `Test_OperationLogger.bas`)
- **Convención de Nombres**: Las funciones de prueba deben comenzar con `Test_`

#### Tipos de Pruebas

**Pruebas Unitarias:**
- Validan funciones individuales
- Prueban lógica de negocio aislada
- Verifican cálculos y validaciones básicas

**Pruebas de Integración:**
- Validan la interacción entre capas del sistema
- Prueban flujos completos de trabajo
- Verifican la integración con servicios externos
- Incluyen escenarios de recuperación de errores

#### Formato de Salida

```
=============================================
        INICIANDO PRUEBAS DE CONDOR
=============================================

>> Ejecutando: Test_SumaBasica...
   [OK] PASO
>> Ejecutando: Test_ConcatenacionTexto...
   [OK] PASO
>> Ejecutando: Test_PruebaQueFalla...
   [X] FALLO - Error 13: No coinciden los tipos

---------------------------------------------
Resumen de Pruebas: 3 total, 2 pasaron, 1 fallaron.
Tiempo total de ejecucion: 0.02 segundos.
=============================================

ATENCION! HUBO ERRORES EN LAS PRUEBAS.
```

#### Características del Framework de Pruebas

**Motor de Ejecución (`modTestRunner.bas`):**
- **Ejecución Silenciosa**: Desactiva avisos de Access durante las pruebas
- **Registro de Suites**: Sistema centralizado para registrar y ejecutar todas las suites de pruebas
- **Manejo de Errores**: Captura y reporta errores específicos con número y descripción
- **Medición de Tiempo**: Calcula y muestra el tiempo total de ejecución

**Generador de Informes (`CTestReporter.cls`):**
- **Responsabilidad Única**: Clase dedicada exclusivamente a la generación de informes
- **Formato Visual**: Utiliza caracteres ASCII compatibles con consolas Windows
- **Resumen Detallado**: Muestra estadísticas completas de éxito/fallo
- **Arquitectura Orientada a Objetos**: Implementación limpia siguiendo principios SOLID

#### Pruebas de Integración Implementadas

El módulo `Test_Integracion.bas` incluye las siguientes pruebas:

**Capa de Presentación:**
- `Test_IntegracionFormularioExpediente`: Integración formulario-negocio
- `Test_IntegracionValidacionDatos`: Validación entre capas

**Capa de Negocio:**
- `Test_IntegracionReglasNegocio`: Aplicación de reglas de negocio
- `Test_IntegracionFlujoTrabajo`: Flujos completos de trabajo

**Capa de Datos:**
- `Test_IntegracionBaseDatos`: Operaciones CRUD
- `Test_IntegracionTransacciones`: Manejo de transacciones

**Servicios Externos:**
- `Test_IntegracionGeneracionDocumentos`: Generación de PDFs
- `Test_IntegracionEnvioEmail`: Envío de correos

**Escenarios Complejos:**
- `Test_IntegracionEscenarioCompleto`: Flujo end-to-end
- `Test_IntegracionManejoConcurrencia`: Acceso concurrente
- `Test_IntegracionRecuperacionErrores`: Recuperación ante fallos

#### Pruebas de Workflow de Estados

El módulo `Test_Solicitud.bas` incluye pruebas específicas para el sistema de workflow:

**Transiciones de Estado:**
- `Test_ChangeState_ValidTransition_ReturnsTrue`: Valida transiciones permitidas entre estados
- `Test_ChangeState_InvalidTransition_ReturnsFalse`: Verifica rechazo de transiciones no válidas

**Arquitectura del Workflow:**
- **IWorkflowRepository.cls**: Interfaz que define el contrato para gestión de reglas de transición
- **CWorkflowRepository.cls**: Implementación real con inyección de dependencia IConfig para acceso a datos
- **CMockWorkflowRepository.cls**: Implementación mock para pruebas con Collection interna y método AddRule
- **CSolicitudService.cls**: Servicio que integra la validación de workflow en el método ChangeState
- **modWorkflowRepositoryFactory.bas**: Factory que maneja la creación e inyección de dependencias

**Pruebas de Integración del Workflow:**
El módulo `Test_WorkflowRepository.bas` contiene pruebas de integración:
- `Test_WorkflowRepository_ValidTransition_ReturnsTrue`: Prueba transiciones válidas con datos reales
- `Test_WorkflowRepository_InvalidTransition_ReturnsFalse`: Prueba transiciones inválidas
- `Test_WorkflowRepository_NonExistentType_ReturnsFalse`: Prueba tipos de solicitud inexistentes
- `Test_WorkflowRepository_InactiveTransition_ReturnsFalse`: Prueba transiciones inactivas

#### Creación de Pruebas

Para crear nuevas pruebas:

1. Crear un módulo VBA (ej: `Test_MiModulo.bas`)
2. Definir funciones que comiencen con `Test_`
3. Usar `On Error Resume Next` para manejo de errores
4. Crear una función `Test_MiModulo_RunAll()` que ejecute todas las pruebas del módulo
5. Registrar la nueva suite en `modTestRunner.RegisterTestSuites()` siguiendo el patrón existente

**Nota**: El framework utiliza un sistema de registro centralizado que facilita la integración de nuevas suites de pruebas.

Ejemplo de función de prueba:
```vba
Public Sub Test_MiFuncion()
    On Error Resume Next
    Err.Clear
    
    Dim resultado As Integer
    resultado = MiFuncion(5, 3)
    
    If resultado <> 8 Then
        Err.Raise 1001, , "Se esperaba 8 pero se obtuvo " & resultado
    End If
End Sub
```

### 3.4. Scripts Auxiliares

CONDOR incluye scripts auxiliares que facilitan la configuración inicial y el mantenimiento del sistema.

#### populate_mappings.vbs

**Propósito:**
Este script lee la configuración de mapeo de campos definida en la Sección 9 de la Especificación Funcional y la inserta en la tabla `TbMapeo_Campos` de la base de datos `CONDOR_datos.accdb`. Los mapeos definen la correspondencia entre los campos de las tablas de datos y los marcadores en las plantillas Word para la generación automática de documentos.

**Cuándo Usarlo:**
- **Configuración inicial**: Se debe ejecutar una vez al inicio del proyecto para poblar la tabla con los mapeos iniciales
- **Sincronización**: Cada vez que haya cambios en la especificación de los mapeos (Sección 9 del documento funcional) para asegurar que la base de datos esté sincronizada con la documentación
- **Restauración**: Cuando sea necesario restaurar los mapeos a su configuración original

**Comando de Ejecución:**
```bash
cscript //nologo populate_mappings.vbs
```

**Funcionalidad:**
- Limpia la tabla `TbMapeo_Campos` antes de insertar los nuevos registros
- Procesa los mapeos para las tres plantillas principales:
  - **PC**: Propuesta de Cambio (F4203.11)
  - **CDCA**: Desviación/Concesión (F4203.10)
  - **CDCASUB**: Desviación/Concesión Sub-suministrador (F4203.101)
- Inserta aproximadamente 116 registros de mapeo
- Proporciona un reporte detallado del proceso de inserción

## Planificación del Proyecto

Consulta el **[Plan de Acción](PLAN_DE_ACCION.md)** para ver el roadmap completo del proyecto, incluyendo:
- Estado actual de desarrollo
- Funcionalidades pendientes organizadas por prioridad
- Sistema de checkboxes para seguimiento del progreso
- Próximos pasos inmediatos

### Flujo de Desarrollo

**Flujo Recomendado (Sincronización Discrecional):**
1. **Desarrollo Local**: Modificar archivos `.bas` en el directorio `src/`
2. **Sincronización Selectiva**: `cscript condor_cli.vbs update [módulos]` para sincronizar solo los archivos modificados
3. **Verificación Manual**: Abrir `CONDOR.accdb`, ejecutar macro `EJECUTAR_TODAS_LAS_PRUEBAS` (Alt+F8) y revisar resultados en Ventana Inmediato (Ctrl+G)
4. **Exportación**: `cscript condor_cli.vbs export` para sincronizar cambios desde Access (opcional)

**Comandos de Sincronización:**
```bash
# Sincronización selectiva (recomendado para desarrollo iterativo)
cscript condor_cli.vbs update CAuthService,modUtils

# Sincronización completa (cuando hay múltiples cambios)
cscript condor_cli.vbs update

# Reconstrucción completa (solo si hay problemas de sincronización)
cscript condor_cli.vbs rebuild
```

**Ventajas de la Sincronización Discrecional:**
- **Eficiencia**: Solo actualiza los módulos que realmente han cambiado
- **Velocidad**: Evita procesar toda la base de datos innecesariamente
- **Flexibilidad**: Permite trabajar en módulos específicos sin afectar otros
- **Desarrollo Iterativo**: Ideal para ciclos rápidos de desarrollo y prueba

**⚠️ Importante**: Las pruebas requieren que los módulos estén compilados. Siempre ejecute `update` o `rebuild` antes de ejecutar las pruebas manualmente para garantizar un estado limpio y compilado.

### Configuración

- **Base de Datos**: `back/Desarrollo/CONDOR.accdb`
- **Directorio de Código**: `src/`
- **Herramienta CLI**: `condor_cli.vbs`

---
Este README se actualiza continuamente según se implementan nuevas funcionalidades.
