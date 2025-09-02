# PLAN DE ACCIÓN - PROYECTO CONDOR

## Estado Actual del Proyecto

**Fecha de última actualización:** $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")

El proyecto ha superado la fase de estabilización estructural del backend. Las capas de Infraestructura, Datos y Lógica de Negocio están completas a nivel de código y compilación. La fase actual se centra en la depuración final de las pruebas de integración y la planificación del desarrollo de la Capa de Presentación.

## Checklist de Progreso por Componente Arquitectónico

### ✅ 1. Capa de Infraestructura y Servicios Centrales
**Estado:** COMPLETADA Y ESTABLE.

### ✅ 2. Capa de Datos (Backend)
**Estado:** ESTRUCTURALMENTE COMPLETA.

**Tareas Pendientes:**
- [PENDIENTE] Finalizar la depuración de las pruebas de integración (TI*.bas).

### ✅ 3. Capa de Lógica de Negocio (Backend)
**Estado:** ESTRUCTURALMENTE COMPLETA.

**Tareas Pendientes:**
- [PENDIENTE] Finalizar la depuración de las pruebas unitarias (Test*.bas).

### 📋 4. Capa de Presentación (Frontend)
**Estado:** PENDIENTE DE DESARROLLO.

**Checklist de Componentes a Implementar:**

#### 4.1. Formulario de Arranque (frmSplash)
```code
- [ ] **Propósito:** Pantalla de carga inicial. No visible para el usuario.
- [ ] **Lógica de `OnLoad`:**
    - [ ] Llamar a `CAppManager.StartApplication` para inicializar todos los servicios.
    - [ ] Obtener el rol del usuario actual (`CAppManager.GetCurrentUserRole`).
    - [ ] Según el rol, abrir el formulario principal correspondiente (`frmPanelCalidad` o `frmPanelTecnico`).
    - [ ] Gestionar errores de arranque (ej. no se puede conectar a la BD) y cerrar la aplicación de forma segura.
```

#### 4.2. Panel Principal - Calidad/Admin (frmPanelCalidad)
```code
- [ ] **Propósito:** Vista principal para los roles de Calidad y Administrador. Centro de operaciones.
- [ ] **Componentes:**
    - [ ] **Botón "Nueva Solicitud":** Inicia el asistente de creación.
    - [ ] **Subformulario de Vista General (`subfrmSolicitudesGrid`):**
        - [ ] Muestra una lista de **todas** las solicitudes del sistema.
        - [ ] Columnas clave: `Código Solicitud`, `Tipo`, `Expediente`, `Estado`, `Fecha Creación`, `Usuario Creación`.
        - [ ] Permitir filtrar y ordenar por cualquier columna.
        - [ ] Doble clic en una fila abre `frmDetalleSolicitud` con los datos de esa solicitud.
    - [ ] **Filtros Avanzados:** Campos para buscar por rango de fechas, tipo, estado o expediente.
```

#### 4.3. Panel Principal - Técnico (frmPanelTecnico)
```code
- [ ] **Propósito:** Vista simplificada y enfocada para el rol Técnico.
- [ ] **Componentes:**
    - [ ] **Subformulario de Tareas Pendientes (`subfrmSolicitudesGridTecnico`):**
        - [ ] Muestra una lista de solicitudes asignadas al técnico actual que están en estado "En Fase Técnica".
        - [ ] Columnas clave: `Código Solicitud`, `Tipo`, `Expediente`, `Fecha Pase a Técnico`.
        - [ ] Doble clic en una fila abre `frmDetalleSolicitud` en modo de edición técnica.
```

#### 4.4. Formulario de Detalle de Solicitud (frmDetalleSolicitud)
```code
- [ ] **Propósito:** Vista y edición de una única solicitud. Es el formulario más complejo.
- [ ] **Lógica de `OnLoad`:**
    - [ ] Recibe un `idSolicitud` al abrirse.
    - [ ] Llama a `CSolicitudService.ObtenerSolicitudPorId` para cargar el objeto `ESolicitud` completo.
    - [ ] Rellena todos los controles del formulario con los datos del objeto.
    - [ ] Habilita/deshabilita controles según el rol del usuario y el estado de la solicitud.
- [ ] **Componentes:**
    - [ ] **Cabecera:** Campos comunes (`Código Solicitud`, `Expediente`, `Estado Actual`, etc.). No editables.
    - [ ] **Control de Pestañas (`tabDatosSolicitud`):**
        - [ ] **Pestaña 1 - Datos PC:** Contiene el subformulario `subfrmDatosPC`. Visible solo si `tipoSolicitud = "PC"`.
        - [ ] **Pestaña 2 - Datos CD/CA:** Contiene el subformulario `subfrmDatosCDCA`. Visible solo si `tipoSolicitud = "CD/CA"`.
        - [ ] **Pestaña 3 - Datos CD/CA-SUB:** Contiene el subformulario `subfrmDatosCDCASUB`. Visible solo si `tipoSolicitud = "CD/CA-SUB"`.
    - [ ] **Sección de Adjuntos (`subfrmAdjuntos`):**
        - [ ] Lista los ficheros adjuntos a la solicitud.
        - [ ] Botones para "Añadir Fichero" y "Abrir Fichero".
    - [ ] **Sección de Historial (`subfrmHistorial`):**
        - [ ] Muestra el log de cambios de estado y modificaciones.
    - [ ] **Barra de Acciones (Botones):**
        - [ ] **Guardar:** Llama a `CSolicitudService.SaveSolicitud`.
        - [ ] **Pasar a Técnico:** (Visible para Calidad) Cambia el estado y llama a `CNotificationService`.
        - [ ] **Finalizar Tarea Técnica:** (Visible para Técnico) Cambia el estado y llama a `CNotificationService`.
        - [ ] **Generar Documento:** Llama a `CDocumentService.GenerarDocumento`.
        - [ ] **Sincronizar desde Documento:** Abre un selector de archivos y llama a `CDocumentService.LeerDocumento`.
```

#### 4.5. Subformularios de Datos Específicos
```code
- [ ] **`subfrmDatosPC`:** Contiene todos los campos de la entidad `EDatosPc`.
- [ ] **`subfrmDatosCDCA`:** Contiene todos los campos de la entidad `EDatosCdCa`.
- [ ] **`subfrmDatosCDCASUB`:** Contiene todos los campos de la entidad `EDatosCdCaSub`.
```

#### 4.6. Asistente de Nueva Solicitud (frmAsistenteNuevaSolicitud)
```code
- [ ] **Propósito:** Guía paso a paso para crear una nueva solicitud.
- [ ] **Paso 1:** Selección del tipo de solicitud (PC, CD/CA, CD/CA-SUB).
- [ ] **Paso 2:** Selección o creación del expediente asociado.
- [ ] **Paso 3:** Cumplimentación de datos específicos según el tipo.
- [ ] **Paso 4:** Confirmación y creación de la solicitud.
- [ ] **Lógica de finalización:** Llama a `CSolicitudService.CreateSolicitud` y abre `frmDetalleSolicitud`.
```

#### 4.7. Formularios de Gestión de Expedientes
```code
- [ ] **`frmBuscarExpediente`:** Búsqueda y selección de expedientes existentes.
- [ ] **`frmNuevoExpediente`:** Creación de nuevos expedientes.
- [ ] **`frmDetalleExpediente`:** Vista detallada de un expediente con sus solicitudes asociadas.
```

#### 4.8. Formularios de Administración (Solo para Administradores)
```code
- [ ] **`frmGestionUsuarios`:** Gestión de usuarios y roles del sistema.
- [ ] **`frmConfiguracion`:** Configuración de parámetros del sistema.
- [ ] **`frmLogViewer`:** Visualización de logs de errores y operaciones.
```

#### 4.9. Componentes Transversales
```code
- [ ] **`subfrmAdjuntos`:** Subformulario reutilizable para gestión de archivos adjuntos.
- [ ] **`subfrmHistorial`:** Subformulario reutilizable para mostrar el historial de cambios.
- [ ] **`subfrmSolicitudesGrid`:** Subformulario reutilizable para listas de solicitudes.
- [ ] **Controles personalizados:** DatePicker, ComboBox con búsqueda, etc.
```

## Notas Técnicas de Implementación

### Arquitectura de Inyección de Dependencias
El sistema utiliza un patrón de **Factorías con Parámetros Opcionales**:
- **En Producción:** Se llama a las factorías sin argumentos.
- **En Pruebas:** Se llama a las factorías inyectando mocks.

### Sistema de Pruebas
El proyecto utiliza un sistema de **Auto-aprovisionamiento** para todas las pruebas de integración (TI*.bas), garantizando entornos de prueba limpios y reproducibles.

### Convenciones de Nomenclatura para Formularios
- **frm**: Formularios principales
- **subfrm**: Subformularios
- **tab**: Controles de pestañas
- **btn**: Botones
- **txt**: Campos de texto
- **cmb**: ComboBox
- **lst**: ListBox

---

**Responsable:** CONDOR-Architect  
**Próxima Revisión:** Al iniciar el desarrollo de la Capa de Presentación.