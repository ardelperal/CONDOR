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

---
### **CICLO DE TRABAJO DE DESARROLLO (MODO AUTÓNOMO)**

**Objetivo:** Automatizar el ciclo de desarrollo completo para cada tarea.

**Capacidades Asumidas:** Tienes acceso a la terminal y puedes ejecutar comandos como `cscript`, `git add`, `git commit` y `git push`.

**Proceso a Seguir para Cada Tarea:**

1.  **Iniciativa (Supervisor Humano):** El Supervisor te indicará la siguiente tarea a realizar del plan de acción.

2.  **Desarrollo y Pruebas (Tu Bucle Autónomo):**
    a. **Generar Código Protegido:** Escribe el código VBA para la funcionalidad y sus pruebas. Ambos deben estar protegidos por el bloque de compilación condicional `#If DEV_MODE Then`.
    b. **Importar a Access con Limpieza:** Ejecuta `cscript //nologo condor_cli.vbs import`. **CRÍTICO:** Durante la importación, el sistema debe "limpiar" automáticamente cada archivo .bas/.cls eliminando todas las líneas que empiecen con "Attribute" antes de usar AddFromString en Access.
    c. **Ejecutar Pruebas:** Ejecuta `cscript //nologo condor_cli.vbs test`.
    d. **Analizar Resultado:**
        *   **Si las pruebas fallan:** Analiza el log de error, corrige el código VBA, y **repite desde el paso 2b**. Continúa en este bucle hasta que todas las pruebas pasen.
        *   **Si las pruebas pasan:** Procede al siguiente paso.

3.  **Finalización y Despliegue (Tu Secuencia Final Autónoma):**
    a. **Liberar Código:** Reescribe los archivos de código VBA (funcionalidad y pruebas) eliminando los bloques `#If DEV_MODE Then`.
    b. **Importar Código Final con Limpieza:** Ejecuta `cscript //nologo condor_cli.vbs import`. **CRÍTICO:** El sistema debe aplicar la misma lógica de limpieza de metadatos "Attribute" durante esta importación final.
    c. **Actualizar Documentación:** Si es necesario, modifica el `README.md` para reflejar la nueva funcionalidad.
    d. **Confirmar Cambios:** Ejecuta la secuencia de Git: `git add .`, `git commit -m "..."` (con un mensaje descriptivo), y `git push`.

4.  **Informe Final (Tu Notificación al Supervisor):** Una vez completado el `push`, notifícame que la tarea se ha completado con éxito y proporciona un resumen de lo que se ha hecho.
---
## Estado del Proyecto
- [x] Estructura base del proyecto
- [x] Herramienta CLI (condor_cli.vbs)
- [x] Sistema de pruebas unitarias
- [x] Sistema de pruebas de integración
- [x] Documentación inicial (README.md)

## 1. ARQUITECTURA Y ESTRUCTURA BASE

### 1.1 Capa de Datos
- [ ] Diseño de base de datos completa
- [ ] Tablas principales (Solicitudes, Estados, Seguimiento)
- [ ] Tablas de configuración (TipoSolicitud, EstadoInterno, etc.)
- [ ] Clase/Interfaz de conexión con aplicación de Expedientes existente
- [ ] Relaciones y constraints
- [ ] Índices para optimización
- [ ] Procedimientos almacenados básicos

### 1.2 Capa de Negocio
- [ ] Clase ExpedienteService (interfaz con aplicación existente)
- [ ] Módulo de gestión de solicitudes
- [ ] Módulo de workflow y estados
- [ ] Módulo de validaciones de negocio
- [ ] Módulo de cálculos y reglas
- [ ] Módulo de notificaciones

### 1.3 Capa de Presentación
- [ ] Formulario principal de navegación
- [ ] Formulario de consulta de expedientes (solo lectura desde app existente)
- [ ] Formulario de gestión de solicitudes
- [ ] Formularios de configuración
- [ ] Reportes y consultas
- [ ] Interfaz de usuario responsive

## 2. FUNCIONALIDADES CORE

### 2.1 Integración con Expedientes Existentes
- [ ] Interfaz de consulta de expedientes por IDExpediente
- [ ] Obtener datos del expediente (nemotécnico, responsable calidad, jefe proyecto)
- [ ] Verificar si somos contratista principal
- [ ] Buscar y filtrar expedientes desde aplicación externa
- [ ] Cache local de datos de expedientes consultados
- [ ] Sincronización con aplicación de expedientes

### 2.2 Gestión de Solicitudes
- [ ] Crear nueva solicitud
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
- [ ] Gestión de usuarios y permisos
- [ ] Configuración de tipos de expediente
- [ ] Configuración de estados y transiciones
- [ ] Plantillas de documentos
- [ ] Configuración de notificaciones
- [ ] Logs del sistema

## 4. CALIDAD Y TESTING

### 4.1 Pruebas
- [x] Framework de pruebas unitarias
- [x] Pruebas de integración básicas
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

### Última actualización: [Fecha]
**Completado:** X/Y tareas (X%)

### Próxima revisión: [Fecha]
**Responsable:** [Nombre]

### Comentarios:
- [Agregar comentarios sobre el progreso]
- [Obstáculos encontrados]
- [Decisiones tomadas]

---

*Este documento se actualiza regularmente para reflejar el progreso del proyecto CONDOR.*