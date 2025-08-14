# CONDOR

## Resumen
CONODOR es una aplicación para gestionar el ciclo de vida de solicitudes de cambio, desviación o concesión en expedientes de contratos públicos. Está desarrollada en Microsoft Access con VBA y orientada a usuarios de Calidad y Técnico.

## Arquitectura

CONDOR sigue una arquitectura de 3 capas con integración a sistema existente:

- **Capa de Presentación**: Formularios de Access
- **Capa de Negocio**: Módulos VBA con lógica de negocio + ExpedienteService (interfaz con app existente)
- **Capa de Datos**: Base de datos Access + Integración con aplicación de expedientes existente por IDExpediente

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

## Gestión de Usuarios y Roles
- Login integrado con el sistema central.
- Roles: Calidad, Técnico, Administrador, y actores externos (solo reciben documentos).

## Flujo de Trabajo
- Fase interna: Preparación y revisión de solicitudes por Calidad y Técnico.
- Fase externa: Generación y envío de documentos a actores externos, recepción y cierre.

## Arquitectura de Código
- Separación en capas: Presentación, Negocio, Acceso a Datos y Servicios Externos.
- Uso de interfaces para facilitar tests unitarios.

## Estructura de Datos
- Tablas principales: Expedientes, Solicitudes, Datos específicos y Mapeo de campos.

## Herramienta CLI de Desarrollo

CONDOR incluye una herramienta de línea de comandos (`condor_cli.vbs`) que facilita el desarrollo y mantenimiento del código VBA.

### Comandos Disponibles

#### Exportación de Módulos
```bash
cscript condor_cli.vbs export
```
- Exporta todos los módulos VBA desde la base de datos Access hacia archivos `.bas` en el directorio `src/`
- Útil para sincronizar cambios realizados directamente en Access hacia el control de versiones
- Mantiene la estructura del código y comentarios

#### Importación de Módulos
```bash
cscript condor_cli.vbs import
```
- Importa todos los archivos `.bas` del directorio `src/` hacia la base de datos Access
- Reemplaza los módulos existentes con las versiones actualizadas
- Compila automáticamente los módulos después de la importación
- Muestra advertencias de compilación si las hay

#### Ejecución de Pruebas
```bash
cscript condor_cli.vbs test
```
- Ejecuta el motor de pruebas interno de CONDOR
- Importa automáticamente los módulos antes de ejecutar las pruebas
- Proporciona un informe detallado de resultados

### Sistema de Pruebas

CONDOR incluye un motor de pruebas unitarias integrado que permite validar la funcionalidad del código VBA.

#### Estructura de Pruebas

- **Motor de Pruebas**: `modTestRunner.bas` - Ejecuta y reporta resultados de todas las pruebas
- **Módulos de Prueba**: Archivos que contienen funciones de prueba (ej: `Test_Ejemplo.bas`, `Test_Integracion.bas`)
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

#### Características del Motor de Pruebas

- **Ejecución Silenciosa**: Desactiva avisos de Access durante las pruebas
- **Manejo de Errores**: Captura y reporta errores específicos con número y descripción
- **Medición de Tiempo**: Calcula y muestra el tiempo total de ejecución
- **Formato Visual**: Utiliza caracteres ASCII compatibles con consolas Windows
- **Resumen Detallado**: Muestra estadísticas completas de éxito/fallo

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

#### Creación de Pruebas

Para crear nuevas pruebas:

1. Crear un módulo VBA (ej: `Test_MiModulo.bas`)
2. Definir funciones que comiencen con `Test_`
3. Usar `On Error Resume Next` para manejo de errores
4. Agregar la llamada en `modTestRunner.RunAllTests()`

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

## Planificación del Proyecto

Consulta el **[Plan de Acción](PLAN_DE_ACCION.md)** para ver el roadmap completo del proyecto, incluyendo:
- Estado actual de desarrollo
- Funcionalidades pendientes organizadas por prioridad
- Sistema de checkboxes para seguimiento del progreso
- Próximos pasos inmediatos

### Flujo de Desarrollo

1. **Desarrollo Local**: Modificar archivos `.bas` en el directorio `src/`
2. **Importación**: `cscript condor_cli.vbs import` para aplicar cambios a Access
3. **Pruebas**: `cscript condor_cli.vbs test` para validar funcionalidad
4. **Exportación**: `cscript condor_cli.vbs export` para sincronizar cambios desde Access

### Configuración

- **Base de Datos**: `back/Desarrollo/CONDOR.accdb`
- **Directorio de Código**: `src/`
- **Herramienta CLI**: `condor_cli.vbs`

---
Este README se actualiza continuamente según se implementan nuevas funcionalidades.
