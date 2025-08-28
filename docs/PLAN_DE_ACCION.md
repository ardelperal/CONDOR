# Plan de Acción - Aplicación CONDOR

---

### **PRINCIPIOS DE ARQUITECTURA DE CÓDIGO**

Para garantizar que la aplicación CONDOR sea robusta, mantenible y testeable, todo el código VBA debe adherirse a los siguientes principios de arquitectura:

1. **Arquitectura en 3 Capas:** El código se organizará en tres capas lógicas:

   * **Capa de Presentación:** Formularios. Su única responsabilidad es mostrar datos y capturar la interacción del usuario. Contiene el mínimo código posible.
   * **Capa de Lógica de Negocio:** Clases. Representan las entidades del negocio (ej. una Solicitud). Contienen las reglas y la lógica de negocio.
   * **Capa de Servicios:** Módulos. Proporcionan servicios de bajo nivel a otras capas (ej. acceso a la base de datos, logging, gestión de ficheros).
2. **Inversión de Dependencias mediante Interfaces:** Este es nuestro principio más importante para la calidad del código.

   * **Regla:** Las clases de alto nivel no deben depender directamente de otras clases concretas; deben depender de **Interfaces**.
   * **Objetivo:** Lograr un **bajo acoplamiento** y, fundamentalmente, permitir la **creación de Mocks para pruebas unitarias**.
   * **Implementación Práctica:**
     * Para cualquier servicio o entidad compleja (como `CExpedienteService`), **primero se debe definir una Interfaz** (ej. `IExpedienteService.cls`).
     * La clase concreta **debe implementar esa interfaz** (ej. `CExpedienteService` implementa `IExpedienteService`).
     * Otras partes del código que necesiten este servicio deberían, en la medida de lo posible, usar variables del tipo de la Interfaz, no de la clase concreta.
3. **Convención de Nomenclatura:**

   * **Interfaces:** Deben empezar con el prefijo `I` (ej. `IExpedienteService`).
   * **Clases:** Deben empezar con el prefijo `C` (ej. `CExpedienteService`).
   * **Módulos:** Deben empezar con el prefijo `mod` (ej. `modConfig`).
   * **Miembros (Propiedades, Funciones, Subrutinas):** Los nombres de miembros deben usar CamelCase. El uso de guiones bajos está prohibido para evitar bugs del compilador.
4. **Principio de Pruebas Unitarias: Programar contra la Interfaz.**

   * **Regla Inquebrantable:** Dentro de cualquier módulo de pruebas (ficheros Test_*.bas), las variables que referencian a nuestras clases de negocio (C*) DEBEN ser declaradas del tipo de su interfaz (I*), no de su clase concreta.
   * **Ejemplo Correcto:** `Dim configSvc As IConfig`
   * **Ejemplo Incorrecto:** `Dim configSvc As CConfig`
   * **Objetivo:** Forzar el desacoplamiento total en el entorno de pruebas. Esto garantiza que los tests solo dependan del contrato público definido en la interfaz, lo cual es esencial para el mocking y la prevención de errores de compilación como "método no encontrado".

---

### **CICLOS DE TRABAJO DE DESARROLLO (MODO TDD AUTÓNOMO)**

El desarrollo en CONDOR se rige por dos modos operativos, dependiendo del estado de compilación del proyecto. El objetivo final es operar siempre en el "Modo Estándar".

---
#### **MODO 1: CICLO DE ESTABILIZACIÓN (Prioridad: Compilar)**

*Este protocolo se activa automáticamente siempre que el proyecto **no compile**. Su único objetivo es alcanzar un estado de compilación exitoso.*

1.  **Reporte de Error (Supervisor):** El Supervisor reporta el error de compilación exacto al Arquitecto mediante una captura de pantalla.
2.  **Prompt de Corrección (Arquitecto):** El Arquitecto genera un prompt específico para la corrección del error. Si detecta un patrón, puede generar un prompt de refactorización proactiva.
3.  **Sincronización (IA):** La IA ejecuta la corrección y sincroniza con `cscript //nologo condor_cli.vbs rebuild`.
4.  **Verificación de Compilación (Supervisor):** El Supervisor compila manualmente en Access (`Depuración -> Compilar`).
5.  **Bucle de Corrección:** Si la compilación falla, se vuelve al paso 1. **El comando `test` está prohibido durante este ciclo.**
6.  **Fin del Ciclo:** Cuando el Supervisor informa **"Compilación exitosa"**, este modo se desactiva y se pasa al Modo 2 para la siguiente tarea.

---
#### **MODO 2: CICLO DE TRABAJO ESTÁNDAR (Prioridad: Funcionalidad y Pruebas)**

*Este es el flujo de trabajo por defecto para desarrollar nuevas funcionalidades o realizar cambios en un proyecto **que ya compila**.*

1.  **Iniciativa (Arquitecto):** El Arquitecto proporciona un prompt detallado para la nueva tarea.
2.  **Desarrollo y Sincronización (IA):** La IA implementa la funcionalidad (idealmente TDD) y sincroniza con `cscript //nologo condor_cli.vbs update`.
3.  **Verificación de Compilación (Supervisor):** El Supervisor realiza la compilación manual en Access para validar la integración del nuevo código. Si falla, se entra en el **MODO 1**.
4.  **Ejecución de Pruebas (IA):** Tras la confirmación de compilación, la IA ejecuta la suite de pruebas completa con `cscript //nologo condor_cli.vbs test`.
5.  **Commit (IA):** Si las pruebas pasan, la IA prepara y realiza el commit.

---

### **SINCRONIZACIÓN DISCRECIONAL DE ARCHIVOS**

CONDOR implementa un sistema avanzado de sincronización que permite actualizar módulos específicos sin afectar el proyecto completo, optimizando el flujo de desarrollo.

**Características Principales:**

1. **Actualización Selectiva:** Permite sincronizar únicamente los módulos que han sido modificados
2. **Eficiencia Mejorada:** Reduce significativamente el tiempo de sincronización al evitar procesar módulos innecesarios
3. **Estabilidad:** Minimiza el riesgo de introducir errores en módulos no relacionados con los cambios
4. **Flexibilidad de Desarrollo:** Facilita el trabajo en funcionalidades específicas sin impactar otras áreas del proyecto

**Comandos de Sincronización:**

```bash
# Sincronización de módulo único
cscript //nologo condor_cli.vbs update CAuthService

# Sincronización de múltiples módulos específicos
cscript //nologo condor_cli.vbs update CAuthService,modConfig,CValidationService

# Sincronización completa (todos los módulos)
cscript //nologo condor_cli.vbs update

# Reconstrucción completa (solo para problemas graves)
cscript //nologo condor_cli.vbs rebuild
```

**Casos de Uso Recomendados:**

- **Desarrollo Iterativo:** Usar sincronización selectiva durante el desarrollo de funcionalidades específicas
- **Correcciones Puntuales:** Actualizar solo los módulos afectados por un bug fix
- **Pruebas Incrementales:** Sincronizar módulos individuales para pruebas focalizadas
- **Integración Continua:** Reducir el tiempo de sincronización en ciclos de desarrollo rápidos

**Implementación Técnica:**

La funcionalidad utiliza `DoCmd.LoadFromText` para importar módulos específicos, eliminando la necesidad de cerrar y reabrir la base de datos Access, lo que resulta en:

- **Mayor Velocidad:** Eliminación del overhead de apertura/cierre de la base de datos
- **Mejor Estabilidad:** Reducción de posibles errores de sincronización
- **Proceso Simplificado:** Menos pasos en el flujo de actualización

---

## Estado del Proyecto

- [X] Estructura base del proyecto
- [X] Herramienta CLI (condor_cli.vbs)
- [X] Sistema de pruebas unitarias
- [X] Sistema de pruebas de integración
- [X] Refactorización del sistema de pruebas (comando test CLI reactivado y robustecido)
- [X] Documentación inicial (README.md)
- [X] Arquitectura en 3 capas implementada
- [X] Sistema de interfaces y mocks para testing
- [X] Servicios de autenticación y configuración
- [X] Framework de testing completo con reportes
- [X] Ejecución automatizada de pruebas desde CLI (comando test con códigos de salida estándar)
- [X] Sistema de sincronización discrecional de archivos (comando update optimizado)
- [X] Sistema de logging de operaciones
- [X] Factory para servicios de configuración


## 1. ARQUITECTURA Y ESTRUCTURA BASE

### 1.1 Capa de Datos

- [X] Diseño de base de datos completa
- [X] Tablas principales (Solicitudes, Estados, Seguimiento)
- [X] Tablas de configuración (TipoSolicitud, EstadoInterno, etc.)
- [X] Clase/Interfaz de conexión con aplicación de Expedientes existente
- [X] Relaciones y constraints
- [X] Índices para optimización
- [X] Procedimientos almacenados básicos

### 1.2 Capa de Negocio

- [X] Interfaces y clases base (IAuthService, CAuthService, CMockAuthService)
- [X] Clase ExpedienteService (interfaz con aplicación existente)
- [X] Módulo de gestión de solicitudes (ISolicitud, CSolicitudPC, ModSolicitudFactory)
- [ ] Módulo de workflow y estados
- [ ] Módulo de validaciones de negocio
- [ ] Módulo de cálculos y reglas
- [ ] Módulo de notificaciones

### 1.3 Capa de Presentación

- [X] Sistema de gestión de aplicaciones (ModAppManager)
- [ ] Formulario principal de navegación
- [ ] Formulario de consulta de expedientes (solo lectura desde app existente)
- [ ] Formulario de gestión de solicitudes
- [ ] Formularios de configuración
- [ ] Reportes y consultas
- [ ] Interfaz de usuario responsive

## 2. FUNCIONALIDADES CORE

### 2.1 Integración con Expedientes Existentes

- [X] Interfaz de consulta de expedientes por IDExpediente
- [X] Obtener datos del expediente (nemotécnico, responsable calidad, jefe proyecto)
- [X] Verificar si somos contratista principal
- [ ] Buscar y filtrar expedientes desde aplicación externa
- [ ] Cache local de datos de expedientes consultados
- [ ] Sincronización con aplicación de expedientes

### 2.2 Gestión de Solicitudes

- [X] Crear nueva solicitud (Factory Pattern implementado)
- [X] Interfaz común ISolicitud para todos los tipos de solicitud
- [X] Implementación CSolicitudPC para solicitudes de PC
- [X] Estructura de datos E_Datos_PC, E_Datos_CD_CA, E_Datos_CD_CA_SUB
- [X] Pruebas unitarias completas para módulo de solicitudes
- [X] Cambio de estados de solicitud (CSolicitudService.ChangeState con validación de workflow)
- [ ] Vincular solicitud a expediente
- [ ] Seguimiento de plazos
- [ ] Generación de documentos
- [ ] Notificaciones automáticas

### 2.3 Workflow y Estados

- [X] Definición de flujos de trabajo (IWorkflowRepository.cls)
- [X] Transiciones de estado automáticas (CSolicitudService.ChangeState con validación)
- [X] Validaciones por estado (CMockWorkflowRepository con reglas configurables)
- [X] Auditoría de cambios (IOperationLogger integrado en ChangeState)
- [X] Pruebas TDD para workflow (Test_ChangeState_ValidTransition y Test_ChangeState_InvalidTransition)
- [ ] Alertas de vencimiento
- [ ] Escalado automático

## 3. FUNCIONALIDADES AVANZADAS

### 3.1 Reportes y Analytics

- [ ] Reporte de expedientes por estado
- [ ] Reporte de solicitudes pendientes
- [ ] Dashboard de métricas
- [ ] Exportación a Excel/PDF
- [ ] Gráficos de tendencias
- [ ] Indicadores KPI

### 3.2 Integración y Comunicación

- [ ] Integración con email
- [ ] Generación automática de documentos
- [ ] Importación/exportación de datos
- [ ] API para integraciones externas
- [ ] Sincronización con sistemas externos
- [ ] Backup automático

### 3.3 Configuración y Administración

- [X] Sistema de configuración base (ModConfig)
- [X] Gestión de usuarios y permisos (AuthService)
- [ ] Configuración de tipos de expediente
- [ ] Configuración de estados y transiciones
- [ ] Plantillas de documentos
- [ ] Configuración de notificaciones
- [X] Logs del sistema

## 4. CALIDAD Y TESTING

### 4.1 Pruebas

- [X] Framework de pruebas unitarias
- [X] Pruebas de integración básicas
- [X] Pruebas unitarias para módulo de solicitudes (Test_Solicitudes)
- [X] Integración de pruebas de solicitudes en ModTestRunner
- [X] Auditoría y corrección completa de Test_CSolicitudPC.bas
- [X] Creación de stubs para funciones de prueba faltantes en CSolicitudPC
- [X] Integración de Test_CSolicitudPC_RunAll en batería completa de pruebas
- [X] Implementación completa de tests CSolicitudPC (Properties_SetAndGet, Load_Success, Save_Success, ChangeState_Success, DatosPC_SetAndGet)
- [X] Corrección de tipos de retorno y propiedades en CSolicitudPC.cls (Property Set/Get para objetos)
- [X] Validación completa: 38/38 tests pasan exitosamente
- [X] Pruebas TDD para workflow de estados (Test_ChangeState_ValidTransition_ReturnsTrue, Test_ChangeState_InvalidTransition_ReturnsFalse)
- [X] Arquitectura de pruebas para workflow (IWorkflowRepository, CMockWorkflowRepository)
- [X] Pruebas de integración para WorkflowRepository (Test_WorkflowRepository.bas)
- [ ] Pruebas de rendimiento
- [ ] Pruebas de seguridad
- [ ] Pruebas de usabilidad
- [ ] Pruebas de regresión

### 4.2 Documentación

- [X] README.md básico
- [ ] Manual de usuario
- [ ] Documentación técnica
- [ ] Guía de instalación
- [ ] Guía de mantenimiento
- [ ] Casos de uso detallados

## 5. DESPLIEGUE Y MANTENIMIENTO

### 5.1 Preparación para Producción

- [ ] Optimización de rendimiento
- [ ] Configuración de seguridad
- [ ] Scripts de instalación
- [ ] Migración de datos
- [ ] Plan de rollback
- [ ] Monitoreo del sistema

### 5.2 Mantenimiento

- [ ] Procedimientos de backup
- [ ] Actualización de versiones
- [ ] Mantenimiento de base de datos
- [ ] Limpieza de logs
- [ ] Optimización periódica
- [ ] Soporte técnico

## 6. PRÓXIMOS PASOS INMEDIATOS

### Prioridad Alta (Próximas 2 semanas)

- [ ] Diseñar esquema de base de datos completo (sin tabla expedientes)
- [ ] Crear clase/interfaz ExpedienteService para conectar con app existente
- [ ] Crear formulario principal de navegación
- [ ] Implementar consulta de expedientes desde aplicación externa
- [ ] Crear pruebas para funcionalidades básicas

### Prioridad Media (Próximo mes)

- [ ] Implementar gestión de solicitudes
- [ ] Desarrollar sistema de workflow
- [ ] Crear reportes básicos
- [ ] Implementar validaciones de negocio

### Prioridad Baja (Próximos 3 meses)

- [ ] Funcionalidades avanzadas de reporting
- [ ] Integraciones externas
- [ ] Optimizaciones de rendimiento
- [ ] Documentación completa

---

## Notas de Progreso

### Última actualización: Enero 2025

**Completado:** 38/85+ tareas (~45%)

### Próxima revisión: Enero 2025

**Responsable:** CONDOR-Expert

### Comentarios:

- ✅ **Arquitectura base completada:** Implementada arquitectura en 3 capas con interfaces
- ✅ **Capa de datos completa:** Base de datos, tablas principales, configuración, relaciones e índices
- ✅ **Sistema de testing robusto:** Framework completo con pruebas unitarias e integración
- ✅ **Servicios fundamentales:** AuthService y Config implementados con mocks
- ✅ **Herramientas de desarrollo:** CLI funcional con importación y testing automatizado
- ✅ **ExpedienteService implementado:** Interfaz IExpedienteService, clase CExpedienteService, mock CMockExpedienteService y pruebas completas
- ✅ **Integración con BD Expedientes:** Consulta SQL compleja implementada con conexión a base de datos externa
- ✅ **Type E_Expediente:** Estructura de datos definida para manejar información completa de expedientes
- ✅ **Módulo de Solicitudes implementado:** ISolicitud, CSolicitudPC, modSolicitudFactory con Factory Pattern
- ✅ **Estructuras de datos de solicitudes:** E_Datos_PC, E_Datos_CD_CA, E_Datos_CD_CA_SUB implementadas
- ✅ **Pruebas de solicitudes:** Test_Solicitudes con cobertura completa del módulo
- ✅ **Sistema de testing automatizado:** Comando test CLI reactivado con ExecuteAllTestsForCLI() y códigos de salida estándar
- [X] Sistema de manejo de errores centralizado:** ModErrorHandler.bas implementado con función LogError
- ✅ **Integración de manejo de errores:** Refactorizado CAuthService, CExpedienteService y ModDatabase para usar sistema centralizado
- ✅ **Pruebas de manejo de errores:** Test_ErrorHandler.bas con cobertura completa del sistema de errores
- [X] Sistema de pruebas completo:** 23 módulos de prueba integrados en ModTestRunner con 38 pruebas ejecutándose exitosamente
- [X] Test_CSolicitudPC.bas implementado:** Suite completa con 7 funciones de prueba (Test_CSolicitudPC_Properties_SetAndGet, Test_CSolicitudPC_Load_Success, Test_CSolicitudPC_Save_Success, Test_CSolicitudPC_ChangeState_Success, Test_CSolicitudPC_DatosPC_SetAndGet) integrada en ModTestRunner.bas
- ✅ **Ciclo de Trabajo Asistido completado:** Tests de CSolicitudPC implementados completamente. Sistema con 38 tests ejecutándose exitosamente, garantizando la estabilidad del proyecto
- 🔧 **Próximo objetivo:** Implementar workflow y estados de solicitudes
- 📋 **Decisión arquitectónica:** Uso de interfaces para permitir mocking y testing efectivo
- ✅ Sistema de Logging de Operaciones implementado (IOperationLogger, COperationLogger, CMockOperationLogger, ModOperationLoggerFactory, Test_OperationLogger)
- ✅ Factory de Configuración implementado (ModConfigFactory)

---

*Este documento se actualiza regularmente para reflejar el progreso del proyecto CONDOR.*
