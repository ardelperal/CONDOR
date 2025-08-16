# Plan de Acción - Aplicación CONDOR
---
### **PRINCIPIOS DE ARQUITECTURA DE CÓDIGO**

Para garantizar que la aplicación CONDOR sea robusta, mantenible y testeable, todo el código VBA debe adherirse a los siguientes principios de arquitectura:

1.  **Arquitectura en 3 Capas:** El código se organizará en tres capas lógicas:
    *   **Capa de Presentación:** Formularios. Su única responsabilidad es mostrar datos y capturar la interacción del usuario. Contiene el mínimo código posible.
    *   **Capa de Lógica de Negocio:** Clases. Representan las entidades del negocio (ej. una Solicitud). Contienen las reglas y la lógica de negocio.
    *   **Capa de Servicios:** Módulos. Proporcionan servicios de bajo nivel a otras capas (ej. acceso a la base de datos, logging, gestión de ficheros).

2.  **Inversión de Dependencias mediante Interfaces:** Este es nuestro principio más importante para la calidad del código.
    *   **Regla:** Las clases de alto nivel no deben depender directamente de otras clases concretas; deben depender de **Interfaces**.
    *   **Objetivo:** Lograr un **bajo acoplamiento** y, fundamentalmente, permitir la **creación de Mocks para pruebas unitarias**.
    *   **Implementación Práctica:**
        *   Para cualquier servicio o entidad compleja (como `CExpedienteService`), **primero se debe definir una Interfaz** (ej. `IExpedienteService.cls`).
        *   La clase concreta **debe implementar esa interfaz** (ej. `CExpedienteService` implementa `IExpedienteService`).
        *   Otras partes del código que necesiten este servicio deberían, en la medida de lo posible, usar variables del tipo de la Interfaz, no de la clase concreta.

3.  **Convención de Nomenclatura:**
    *   **Interfaces:** Deben empezar con el prefijo `I` (ej. `IExpedienteService`).
    *   **Clases:** Deben empezar con el prefijo `C` (ej. `CExpedienteService`).
    *   **Módulos:** Deben empezar con el prefijo `mod` (ej. `modConfig`).
    *   **Miembros (Propiedades, Funciones, Subrutinas):** Los nombres de miembros deben usar CamelCase. El uso de guiones bajos está prohibido para evitar bugs del compilador.

4.  **Principio de Pruebas Unitarias: Programar contra la Interfaz.**
    *   **Regla Inquebrantable:** Dentro de cualquier módulo de pruebas (ficheros Test_*.bas), las variables que referencian a nuestras clases de negocio (C*) DEBEN ser declaradas del tipo de su interfaz (I*), no de su clase concreta.
    *   **Ejemplo Correcto:** `Dim configSvc As IConfig`
    *   **Ejemplo Incorrecto:** `Dim configSvc As CConfig`
    *   **Objetivo:** Forzar el desacoplamiento total en el entorno de pruebas. Esto garantiza que los tests solo dependan del contrato público definido en la interfaz, lo cual es esencial para el mocking y la prevención de errores de compilación como "método no encontrado".

---
### **CICLO DE TRABAJO ASISTIDO (SUPERVISOR + IA)**

**Objetivo:** Combinar la velocidad de la IA con la supervisión humana en el punto crítico de la compilación para garantizar la estabilidad del proyecto.

**Proceso a Seguir para Cada Tarea:**

1.  **Iniciativa (CONDOR-Architect):** El Arquitecto le proporciona al Supervisor un prompt detallado para la tarea.

2.  **Ejecución de IA (Tu Rol, Copilot):**
    a. Recibes el prompt del Supervisor.
    b. Generas el código necesario y ejecutas `cscript //nologo condor_cli.vbs rebuild` para sincronizar el proyecto.
    c. Pausa y espera la confirmación del Supervisor.

3.  **Verificación Manual (Rol del Supervisor):**
    a. El Supervisor abre CONDOR.accdb y ejecuta Depuración -> Compilar Proyecto.
    b. Si hay un error, el ciclo se detiene y el Supervisor lo reporta al Arquitecto.
    c. Si el proyecto compila, el Supervisor te da la orden para continuar.

4.  **Finalización de IA (Tu Rol, Copilot):**
    a. Tras la confirmación del Supervisor, ejecutas la secuencia final: `cscript //nologo condor_cli.vbs compile` y luego `cscript //nologo condor_cli.vbs test`.
    b. Si las pruebas pasan, preparas el commit y ejecutas el push.

5.  **Informe Final (Tu Rol, Copilot):** Notificas al Supervisor que la tarea se ha completado con éxito.
---
## Estado del Proyecto
- [x] Estructura base del proyecto
- [x] Herramienta CLI (condor_cli.vbs)
- [x] Sistema de pruebas unitarias
- [x] Sistema de pruebas de integración
- [x] Refactorización del sistema de pruebas (eliminación comando test CLI, método manual implementado)
- [x] Documentación inicial (README.md)
- [x] Arquitectura en 3 capas implementada
- [x] Sistema de interfaces y mocks para testing
- [x] Servicios de autenticación y configuración
- [x] Framework de testing completo con reportes
- [x] Método manual de ejecución de pruebas (_EJECUTAR_TODAS_LAS_PRUEBAS)

## 1. ARQUITECTURA Y ESTRUCTURA BASE

### 1.1 Capa de Datos
- [x] Diseño de base de datos completa
- [x] Tablas principales (Solicitudes, Estados, Seguimiento)
- [x] Tablas de configuración (TipoSolicitud, EstadoInterno, etc.)
- [x] Clase/Interfaz de conexión con aplicación de Expedientes existente
- [x] Relaciones y constraints
- [x] Índices para optimización
- [x] Procedimientos almacenados básicos

### 1.2 Capa de Negocio
- [x] Interfaces y clases base (IAuthService, CAuthService, CMockAuthService)
- [x] Clase ExpedienteService (interfaz con aplicación existente)
- [x] Módulo de gestión de solicitudes (ISolicitud, CSolicitudPC, modSolicitudFactory)
- [ ] Módulo de workflow y estados
- [ ] Módulo de validaciones de negocio
- [ ] Módulo de cálculos y reglas
- [ ] Módulo de notificaciones

### 1.3 Capa de Presentación
- [x] Sistema de gestión de aplicaciones (modAppManager)
- [ ] Formulario principal de navegación
- [ ] Formulario de consulta de expedientes (solo lectura desde app existente)
- [ ] Formulario de gestión de solicitudes
- [ ] Formularios de configuración
- [ ] Reportes y consultas
- [ ] Interfaz de usuario responsive

## 2. FUNCIONALIDADES CORE

### 2.1 Integración con Expedientes Existentes
- [x] Interfaz de consulta de expedientes por IDExpediente
- [x] Obtener datos del expediente (nemotécnico, responsable calidad, jefe proyecto)
- [x] Verificar si somos contratista principal
- [ ] Buscar y filtrar expedientes desde aplicación externa
- [ ] Cache local de datos de expedientes consultados
- [ ] Sincronización con aplicación de expedientes

### 2.2 Gestión de Solicitudes
- [x] Crear nueva solicitud (Factory Pattern implementado)
- [x] Interfaz común ISolicitud para todos los tipos de solicitud
- [x] Implementación CSolicitudPC para solicitudes de PC
- [x] Estructura de datos T_Datos_PC, T_Datos_CD_CA, T_Datos_CD_CA_SUB
- [x] Pruebas unitarias completas para módulo de solicitudes
- [ ] Vincular solicitud a expediente
- [ ] Cambio de estados de solicitud
- [ ] Seguimiento de plazos
- [ ] Generación de documentos
- [ ] Notificaciones automáticas

### 2.3 Workflow y Estados
- [ ] Definición de flujos de trabajo
- [ ] Transiciones de estado automáticas
- [ ] Validaciones por estado
- [ ] Alertas de vencimiento
- [ ] Escalado automático
- [ ] Auditoría de cambios

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
- [x] Sistema de configuración base (modConfig)
- [x] Gestión de usuarios y permisos (AuthService)
- [ ] Configuración de tipos de expediente
- [ ] Configuración de estados y transiciones
- [ ] Plantillas de documentos
- [ ] Configuración de notificaciones
- [ ] Logs del sistema

## 4. CALIDAD Y TESTING

### 4.1 Pruebas
- [x] Framework de pruebas unitarias
- [x] Pruebas de integración básicas
- [x] Pruebas unitarias para módulo de solicitudes (Test_Solicitudes)
- [x] Integración de pruebas de solicitudes en modTestRunner
- [ ] Pruebas de rendimiento
- [ ] Pruebas de seguridad
- [ ] Pruebas de usabilidad
- [ ] Pruebas de regresión

### 4.2 Documentación
- [x] README.md básico
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
**Completado:** 35/85+ tareas (~41%)

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
- ✅ **Type T_Expediente:** Estructura de datos definida para manejar información completa de expedientes
- ✅ **Módulo de Solicitudes implementado:** ISolicitud, CSolicitudPC, modSolicitudFactory con Factory Pattern
- ✅ **Estructuras de datos de solicitudes:** T_Datos_PC, T_Datos_CD_CA, T_Datos_CD_CA_SUB implementadas
- ✅ **Pruebas de solicitudes:** Test_Solicitudes con cobertura completa del módulo
- ✅ **Refactorización del sistema de testing:** Eliminado comando test de CLI, implementado método manual _EJECUTAR_TODAS_LAS_PRUEBAS
- ✅ **Documentación actualizada:** README.md y PLAN_DE_ACCION.md actualizados con nuevo flujo de pruebas
- ✅ **Sistema de manejo de errores centralizado:** modErrorHandler.bas implementado con función LogError
- ✅ **Integración de manejo de errores:** Refactorizado CAuthService, CExpedienteService y modDatabase para usar sistema centralizado
- ✅ **Pruebas de manejo de errores:** Test_ErrorHandler.bas con cobertura completa del sistema de errores
- 🔧 **Próximo objetivo:** Implementar workflow y estados de solicitudes
- 📋 **Decisión arquitectónica:** Uso de interfaces para permitir mocking y testing efectivo

---

*Este documento se actualiza regularmente para reflejar el progreso del proyecto CONDOR.*