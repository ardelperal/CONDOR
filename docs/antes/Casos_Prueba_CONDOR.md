# Casos de Prueba - Sistema CONDOR

## Introducción

Este documento contiene los casos de prueba para la validación funcional del Sistema CONDOR, basados en las reglas de negocio, flujos de trabajo y requisitos definidos en la especificación funcional del proyecto.

---

## 1. Estrategia de Pruebas

### 1.1 Tipos de Pruebas

- **Pruebas Unitarias**: Validación de componentes individuales
- **Pruebas de Integración**: Validación de interacciones entre componentes
- **Pruebas Funcionales**: Validación de requisitos de negocio
- **Pruebas de Usuario**: Validación de flujos completos de usuario
- **Pruebas de Regresión**: Validación de funcionalidades existentes tras cambios

### 1.2 Criterios de Aceptación

- Todas las reglas de negocio deben cumplirse
- Los permisos por rol deben respetarse
- Las integraciones externas deben funcionar correctamente
- La generación de documentos debe ser precisa
- Los logs deben registrar todas las operaciones críticas

---

## 2. Casos de Prueba por Módulo

### 2.1 Módulo de Autenticación y Roles

#### CP-AUTH-001: Login Exitoso con Usuario Válido
**Objetivo**: Verificar que un usuario válido puede acceder al sistema

**Precondiciones**:
- Sistema CONDOR iniciado
- Usuario existe en TbUsuarios
- Usuario tiene rol asignado

**Datos de Prueba**:
```
Usuario: juan.perez@empresa.com
Rol: Calidad
Estado: Activo
```

**Pasos**:
1. Iniciar aplicación CONDOR
2. El sistema detecta automáticamente el usuario Windows
3. Verificar que se muestra la interfaz principal
4. Verificar que el menú refleja los permisos del rol Calidad

**Resultado Esperado**:
- Usuario autenticado exitosamente
- Interfaz cargada con opciones según rol Calidad
- Log de acceso registrado en TbLog_Accesos

**Criterios de Aceptación**:
- [ ] Login automático funciona
- [ ] Permisos de rol aplicados correctamente
- [ ] Log de acceso creado

---

#### CP-AUTH-002: Acceso Denegado para Usuario Inactivo
**Objetivo**: Verificar que usuarios inactivos no pueden acceder

**Precondiciones**:
- Usuario existe en TbUsuarios
- Usuario tiene Estado = "Inactivo"

**Datos de Prueba**:
```
Usuario: maria.garcia@empresa.com
Rol: Ingeniería
Estado: Inactivo
```

**Pasos**:
1. Iniciar aplicación CONDOR
2. Sistema detecta usuario inactivo

**Resultado Esperado**:
- Mensaje de error: "Usuario inactivo. Contacte al administrador"
- Aplicación se cierra
- Log de intento de acceso denegado

**Criterios de Aceptación**:
- [ ] Acceso denegado correctamente
- [ ] Mensaje de error apropiado
- [ ] Log de seguridad registrado

---

#### CP-AUTH-003: Verificación de Permisos por Rol
**Objetivo**: Verificar que cada rol tiene acceso solo a sus funciones permitidas

**Datos de Prueba**:
```
Roles a probar:
- Calidad: Crear/Editar/Consultar solicitudes
- Ingeniería: Solo consultar solicitudes
- Administrador: Todas las funciones
```

**Matriz de Permisos a Validar**:

| Función | Calidad | Ingeniería | Administrador |
|---------|---------|------------|---------------|
| Crear Solicitud PC | ✓ | ✗ | ✓ |
| Crear Solicitud CD_CA | ✓ | ✗ | ✓ |
| Editar Solicitud Propia | ✓ | ✗ | ✓ |
| Editar Cualquier Solicitud | ✗ | ✗ | ✓ |
| Consultar Solicitudes | ✓ | ✓ | ✓ |
| Generar Documentos | ✓ | ✗ | ✓ |
| Gestionar Usuarios | ✗ | ✗ | ✓ |

**Criterios de Aceptación**:
- [ ] Cada rol accede solo a funciones permitidas
- [ ] Botones/menús deshabilitados según permisos
- [ ] Mensajes de error apropiados para accesos no permitidos

---

### 2.2 Módulo de Gestión de Solicitudes

#### CP-SOL-001: Crear Solicitud PC Válida
**Objetivo**: Verificar creación exitosa de solicitud de Cambio de Precio

**Precondiciones**:
- Usuario con rol Calidad autenticado
- Expediente existe en sistema externo

**Datos de Prueba**:
```
Tipo Solicitud: PC
Número Expediente: EXP-2024-001
Descripción: "Cambio de precio por incremento de materiales según índice ICCP"
Justificación: "Incremento del 15% en acero estructural"
Importe Original: 1,500,000.00
Importe Nuevo: 1,725,000.00
Porcentaje Variación: 15.00%
```

**Pasos**:
1. Acceder a "Nueva Solicitud"
2. Seleccionar tipo "PC"
3. Ingresar número de expediente
4. Completar campos obligatorios
5. Hacer clic en "Guardar"

**Resultado Esperado**:
- Solicitud creada con ID único
- Estado inicial: "Borrador"
- Campos calculados automáticamente (porcentaje, fechas)
- Notificación enviada a Ingeniería
- Log de creación registrado

**Criterios de Aceptación**:
- [ ] Solicitud guardada en Tb_Solicitudes
- [ ] ID único generado correctamente
- [ ] Campos calculados son precisos
- [ ] Estado inicial correcto
- [ ] Notificación enviada
- [ ] Log registrado

---

#### CP-SOL-002: Validación de Campos Obligatorios
**Objetivo**: Verificar que no se puede crear solicitud sin campos obligatorios

**Datos de Prueba**:
```
Casos a probar:
1. Sin número de expediente
2. Sin descripción
3. Sin justificación
4. Con expediente inexistente
```

**Pasos para cada caso**:
1. Acceder a "Nueva Solicitud"
2. Completar formulario omitiendo campo específico
3. Intentar guardar

**Resultado Esperado**:
- Mensaje de validación específico para cada campo
- Solicitud no se guarda
- Foco se posiciona en campo con error

**Criterios de Aceptación**:
- [ ] Validaciones funcionan para todos los campos obligatorios
- [ ] Mensajes de error son claros y específicos
- [ ] No se crean registros inválidos

---

#### CP-SOL-003: Editar Solicitud en Estado Borrador
**Objetivo**: Verificar que se puede editar solicitud en estado Borrador

**Precondiciones**:
- Solicitud existe en estado "Borrador"
- Usuario es el creador de la solicitud

**Datos de Prueba**:
```
Solicitud ID: SOL-2024-001
Estado Actual: Borrador
Campo a Modificar: Descripción
Nuevo Valor: "Descripción actualizada con más detalles"
```

**Pasos**:
1. Buscar solicitud SOL-2024-001
2. Hacer clic en "Editar"
3. Modificar descripción
4. Guardar cambios

**Resultado Esperado**:
- Cambios guardados exitosamente
- FechaUltimaModificacion actualizada
- Log de modificación registrado

**Criterios de Aceptación**:
- [ ] Edición permitida en estado Borrador
- [ ] Cambios persistidos correctamente
- [ ] Metadatos de modificación actualizados
- [ ] Log de auditoría creado

---

#### CP-SOL-004: Restricción de Edición en Estado Enviado
**Objetivo**: Verificar que no se puede editar solicitud enviada

**Precondiciones**:
- Solicitud existe en estado "Enviado"

**Datos de Prueba**:
```
Solicitud ID: SOL-2024-002
Estado Actual: Enviado
```

**Pasos**:
1. Buscar solicitud SOL-2024-002
2. Intentar hacer clic en "Editar"

**Resultado Esperado**:
- Botón "Editar" deshabilitado o no visible
- Si se intenta editar: mensaje "No se puede editar solicitud en estado Enviado"

**Criterios de Aceptación**:
- [ ] Edición bloqueada para estados no editables
- [ ] Interfaz refleja restricciones
- [ ] Mensaje de error apropiado

---

### 2.3 Módulo de Búsqueda y Consulta

#### CP-BUS-001: Búsqueda por Número de Expediente
**Objetivo**: Verificar búsqueda exitosa por número de expediente

**Datos de Prueba**:
```
Criterio: Número de Expediente
Valor: EXP-2024-001
Resultados Esperados: 2 solicitudes
```

**Pasos**:
1. Acceder a "Buscar Solicitudes"
2. Ingresar "EXP-2024-001" en campo Expediente
3. Hacer clic en "Buscar"

**Resultado Esperado**:
- Lista con 2 solicitudes relacionadas al expediente
- Información básica mostrada (ID, Tipo, Estado, Fecha)
- Opción de ver detalles disponible

**Criterios de Aceptación**:
- [ ] Búsqueda retorna resultados correctos
- [ ] Información mostrada es precisa
- [ ] Performance aceptable (<2 segundos)

---

#### CP-BUS-002: Búsqueda Avanzada con Múltiples Criterios
**Objetivo**: Verificar búsqueda con combinación de criterios

**Datos de Prueba**:
```
Tipo Solicitud: PC
Estado: Borrador
Fecha Desde: 01/12/2024
Fecha Hasta: 31/12/2024
Usuario Creador: juan.perez@empresa.com
```

**Pasos**:
1. Acceder a "Búsqueda Avanzada"
2. Completar múltiples criterios
3. Ejecutar búsqueda

**Resultado Esperado**:
- Resultados filtrados según todos los criterios
- Posibilidad de exportar resultados
- Paginación si hay muchos resultados

**Criterios de Aceptación**:
- [ ] Filtros combinados funcionan correctamente
- [ ] Resultados son precisos
- [ ] Exportación funciona
- [ ] Paginación implementada

---

#### CP-BUS-003: Búsqueda Sin Resultados
**Objetivo**: Verificar comportamiento cuando no hay resultados

**Datos de Prueba**:
```
Criterio: Número de Expediente
Valor: EXP-INEXISTENTE-999
```

**Pasos**:
1. Buscar expediente inexistente
2. Ejecutar búsqueda

**Resultado Esperado**:
- Mensaje: "No se encontraron solicitudes con los criterios especificados"
- Sugerencias para refinar búsqueda
- Opción de limpiar criterios

**Criterios de Aceptación**:
- [ ] Mensaje informativo apropiado
- [ ] No se muestran datos erróneos
- [ ] Interfaz mantiene usabilidad

---

### 2.4 Módulo de Generación de Documentos

#### CP-DOC-001: Generar Documento PC
**Objetivo**: Verificar generación correcta de documento de Cambio de Precio

**Precondiciones**:
- Solicitud PC en estado "Aprobado"
- Plantilla PC_template.docx disponible
- Usuario con permisos de generación

**Datos de Prueba**:
```
Solicitud ID: SOL-2024-001
Tipo: PC
Expediente: EXP-2024-001
Contratista: "CONSTRUCTORA ABC S.A."
Importe Original: 1,500,000.00
Importe Nuevo: 1,725,000.00
```

**Pasos**:
1. Seleccionar solicitud SOL-2024-001
2. Hacer clic en "Generar Documento"
3. Confirmar generación
4. Verificar documento generado

**Resultado Esperado**:
- Documento Word generado correctamente
- Todos los campos reemplazados con datos reales
- Formato y estructura mantenidos
- Archivo guardado en ubicación temporal

**Criterios de Aceptación**:
- [ ] Documento generado sin errores
- [ ] Todos los marcadores reemplazados
- [ ] Formato preservado
- [ ] Cálculos correctos en el documento

---

#### CP-DOC-002: Mapeo de Campos en Plantilla
**Objetivo**: Verificar que todos los campos se mapean correctamente

**Campos a Verificar**:
```
{{NumeroExpediente}} → EXP-2024-001
{{TipoSolicitud}} → PC
{{DescripcionSolicitud}} → Descripción completa
{{FechaCreacion}} → 20/12/2024
{{ImporteOriginal}} → $1,500,000.00
{{ImporteNuevo}} → $1,725,000.00
{{PorcentajeVariacion}} → 15.00%
{{Contratista}} → CONSTRUCTORA ABC S.A.
{{JustificacionTecnica}} → Justificación detallada
```

**Criterios de Aceptación**:
- [ ] Todos los campos mapeados correctamente
- [ ] Formato de números y fechas apropiado
- [ ] Caracteres especiales manejados correctamente

---

#### CP-DOC-003: Error en Plantilla Faltante
**Objetivo**: Verificar manejo de error cuando falta plantilla

**Precondiciones**:
- Plantilla PC_template.docx no existe o no es accesible

**Pasos**:
1. Intentar generar documento PC
2. Sistema detecta plantilla faltante

**Resultado Esperado**:
- Mensaje de error: "Plantilla no encontrada: PC_template.docx"
- Log de error registrado
- Proceso cancelado sin generar archivo corrupto

**Criterios de Aceptación**:
- [ ] Error manejado graciosamente
- [ ] Mensaje de error informativo
- [ ] Log de error registrado
- [ ] No se generan archivos corruptos

---

### 2.5 Módulo de Integración Externa

#### CP-INT-001: Consulta Exitosa de Expediente
**Objetivo**: Verificar integración con ExpedienteService

**Precondiciones**:
- Servicio ExpedienteService disponible
- Expediente existe en sistema externo

**Datos de Prueba**:
```
Número Expediente: EXP-2024-001
Respuesta Esperada:
{
  "numeroExpediente": "EXP-2024-001",
  "contratista": "CONSTRUCTORA ABC S.A.",
  "objeto": "Construcción de puente vehicular",
  "importeContrato": 1500000.00,
  "fechaInicio": "2024-01-15",
  "estado": "En Ejecución"
}
```

**Pasos**:
1. Ingresar número de expediente en formulario
2. Sistema consulta ExpedienteService automáticamente
3. Verificar datos cargados en formulario

**Resultado Esperado**:
- Datos del expediente cargados automáticamente
- Campos del formulario poblados
- Tiempo de respuesta < 5 segundos

**Criterios de Aceptación**:
- [ ] Integración funciona correctamente
- [ ] Datos mapeados apropiadamente
- [ ] Performance aceptable
- [ ] Manejo de timeout implementado

---

#### CP-INT-002: Expediente No Encontrado
**Objetivo**: Verificar manejo cuando expediente no existe

**Datos de Prueba**:
```
Número Expediente: EXP-INEXISTENTE-999
Respuesta Esperada: HTTP 404 Not Found
```

**Pasos**:
1. Ingresar número de expediente inexistente
2. Sistema intenta consultar ExpedienteService

**Resultado Esperado**:
- Mensaje: "Expediente no encontrado en el sistema"
- Campos del formulario permanecen vacíos
- Usuario puede continuar con entrada manual

**Criterios de Aceptación**:
- [ ] Error 404 manejado correctamente
- [ ] Mensaje de error claro
- [ ] Formulario sigue siendo usable

---

#### CP-INT-003: Timeout de Servicio Externo
**Objetivo**: Verificar manejo de timeout en servicios externos

**Precondiciones**:
- Servicio ExpedienteService lento o no disponible
- Timeout configurado en 30 segundos

**Pasos**:
1. Configurar servicio para simular lentitud
2. Intentar consultar expediente
3. Esperar timeout

**Resultado Esperado**:
- Después de 30 segundos: mensaje de timeout
- Opción de reintentar
- Opción de continuar sin datos externos

**Criterios de Aceptación**:
- [ ] Timeout configurado funciona
- [ ] Mensaje de timeout apropiado
- [ ] Opciones de recuperación disponibles
- [ ] Aplicación no se cuelga

---

### 2.6 Módulo de Notificaciones

#### CP-NOT-001: Notificación de Nueva Solicitud
**Objetivo**: Verificar envío de notificación al crear solicitud

**Precondiciones**:
- Solicitud PC creada
- Configuración SMTP válida
- Usuarios de Ingeniería configurados

**Datos de Prueba**:
```
Solicitud: SOL-2024-001 (PC)
Destinatarios: ingenieria@empresa.com
Asunto: "Nueva solicitud PC - EXP-2024-001"
```

**Pasos**:
1. Crear nueva solicitud PC
2. Cambiar estado a "Enviado"
3. Verificar envío de notificación

**Resultado Esperado**:
- Email enviado a grupo Ingeniería
- Contenido incluye datos básicos de solicitud
- Link para acceder al sistema
- Log de notificación registrado

**Criterios de Aceptación**:
- [ ] Notificación enviada correctamente
- [ ] Contenido del email es apropiado
- [ ] Destinatarios correctos
- [ ] Log de envío registrado

---

#### CP-NOT-002: Notificación de Cambio de Estado
**Objetivo**: Verificar notificación cuando cambia estado de solicitud

**Datos de Prueba**:
```
Solicitud: SOL-2024-001
Estado Anterior: Enviado
Estado Nuevo: Aprobado
Destinatario: Creador de la solicitud
```

**Pasos**:
1. Cambiar estado de solicitud a "Aprobado"
2. Verificar notificación al creador

**Resultado Esperado**:
- Email enviado al creador original
- Asunto indica cambio de estado
- Contenido explica el cambio

**Criterios de Aceptación**:
- [ ] Notificación de cambio de estado funciona
- [ ] Destinatario correcto (creador)
- [ ] Contenido informativo apropiado

---

### 2.7 Módulo de Auditoría y Logging

#### CP-AUD-001: Registro de Operaciones Críticas
**Objetivo**: Verificar que todas las operaciones críticas se registran

**Operaciones a Verificar**:
```
1. Login de usuario
2. Creación de solicitud
3. Modificación de solicitud
4. Cambio de estado
5. Generación de documento
6. Consulta de expediente externo
7. Envío de notificación
8. Errores del sistema
```

**Criterios de Aceptación**:
- [ ] Todas las operaciones críticas se registran
- [ ] Logs incluyen timestamp, usuario, acción y resultado
- [ ] Logs son legibles y estructurados
- [ ] Rotación de logs funciona correctamente

---

#### CP-AUD-002: Trazabilidad de Cambios
**Objetivo**: Verificar trazabilidad completa de cambios en solicitudes

**Datos de Prueba**:
```
Solicitud: SOL-2024-001
Cambios a realizar:
1. Modificar descripción
2. Cambiar estado a Enviado
3. Aprobar solicitud
4. Generar documento
```

**Pasos**:
1. Realizar cada cambio secuencialmente
2. Verificar registro en logs de auditoría
3. Consultar historial de la solicitud

**Resultado Esperado**:
- Cada cambio registrado con detalle
- Historial completo disponible
- Posibilidad de reconstruir estado en cualquier momento

**Criterios de Aceptación**:
- [ ] Trazabilidad completa implementada
- [ ] Historial de cambios accesible
- [ ] Información suficiente para auditoría

---

## 3. Casos de Prueba de Integración

### 3.1 Flujo Completo de Solicitud PC

#### CP-FLOW-001: Flujo Completo Exitoso
**Objetivo**: Verificar flujo completo desde creación hasta documento final

**Pasos del Flujo**:
1. **Login** (Usuario Calidad)
2. **Crear Solicitud PC**
   - Expediente: EXP-2024-001
   - Descripción completa
   - Importes y justificación
3. **Enviar para Revisión**
   - Cambiar estado a "Enviado"
   - Verificar notificación a Ingeniería
4. **Revisión Técnica** (Usuario Ingeniería)
   - Login como Ingeniería
   - Consultar solicitud
   - Aprobar solicitud
5. **Generación de Documento** (Usuario Calidad)
   - Generar documento oficial
   - Verificar contenido
6. **Finalización**
   - Cambiar estado a "Finalizado"
   - Verificar notificaciones finales

**Criterios de Aceptación**:
- [ ] Flujo completo ejecuta sin errores
- [ ] Todos los estados se transicionan correctamente
- [ ] Notificaciones enviadas en momentos apropiados
- [ ] Documento generado es correcto
- [ ] Auditoría completa registrada

---

### 3.2 Pruebas de Carga y Performance

#### CP-PERF-001: Múltiples Usuarios Concurrentes
**Objetivo**: Verificar comportamiento con múltiples usuarios simultáneos

**Escenario**:
```
Usuarios Concurrentes: 10
Operaciones por Usuario:
- 5 consultas de expedientes
- 2 creaciones de solicitudes
- 3 búsquedas
- 1 generación de documento

Duración: 30 minutos
```

**Criterios de Aceptación**:
- [ ] Tiempo de respuesta < 3 segundos para operaciones normales
- [ ] No hay bloqueos de base de datos
- [ ] Todas las operaciones completan exitosamente
- [ ] Memoria y CPU del servidor dentro de límites

---

#### CP-PERF-002: Volumen de Datos
**Objetivo**: Verificar performance con gran volumen de datos

**Datos de Prueba**:
```
Solicitudes en BD: 10,000
Expedientes: 5,000
Usuarios: 100
Logs: 50,000 registros
```

**Operaciones a Probar**:
- Búsquedas complejas
- Reportes con muchos registros
- Consultas de auditoría
- Backup y restauración

**Criterios de Aceptación**:
- [ ] Búsquedas completan en < 10 segundos
- [ ] Reportes generan en tiempo razonable
- [ ] Base de datos mantiene performance
- [ ] Backup/restore funciona correctamente

---

## 4. Casos de Prueba de Seguridad

### 4.1 Validación de Permisos

#### CP-SEC-001: Intento de Acceso No Autorizado
**Objetivo**: Verificar que usuarios no pueden acceder a funciones no permitidas

**Escenarios**:
```
1. Usuario Ingeniería intenta crear solicitud
2. Usuario Calidad intenta gestionar usuarios
3. Usuario inactivo intenta acceder al sistema
4. Acceso directo a formularios sin autenticación
```

**Criterios de Aceptación**:
- [ ] Todos los intentos no autorizados son bloqueados
- [ ] Mensajes de error apropiados
- [ ] Intentos registrados en logs de seguridad
- [ ] No hay bypass de seguridad posible

---

### 4.2 Validación de Datos

#### CP-SEC-002: Inyección de Código
**Objetivo**: Verificar protección contra inyección de código malicioso

**Datos de Prueba Maliciosos**:
```
SQL Injection:
- '; DROP TABLE Tb_Solicitudes; --
- ' OR '1'='1

Script Injection:
- <script>alert('XSS')</script>
- javascript:alert('XSS')

Path Traversal:
- ../../../windows/system32/
- \\..\\..\\config.ini
```

**Criterios de Aceptación**:
- [ ] Todos los inputs maliciosos son sanitizados
- [ ] No se ejecuta código no autorizado
- [ ] Datos se almacenan de forma segura
- [ ] Intentos de inyección registrados

---

## 5. Casos de Prueba de Regresión

### 5.1 Funcionalidades Core

#### CP-REG-001: Regresión Post-Actualización
**Objetivo**: Verificar que funcionalidades existentes siguen funcionando tras actualizaciones

**Funcionalidades Críticas a Verificar**:
```
1. Login y autenticación
2. Creación de solicitudes (todos los tipos)
3. Búsqueda y consulta
4. Generación de documentos
5. Integración con servicios externos
6. Notificaciones por email
7. Auditoría y logging
8. Gestión de usuarios (Admin)
```

**Criterios de Aceptación**:
- [ ] Todas las funcionalidades core funcionan
- [ ] Performance no se ha degradado
- [ ] Datos existentes siguen siendo accesibles
- [ ] Configuraciones se mantienen

---

## 6. Automatización de Pruebas

### 6.1 Suite de Pruebas Automatizadas

```vba
' Archivo: TestSuite_CONDOR.bas
' Suite principal de pruebas automatizadas

Public Sub EjecutarSuiteCompleta()
    Dim resultados As New Collection
    
    ' Pruebas de Autenticación
    resultados.Add EjecutarPruebasAuth()
    
    ' Pruebas de Solicitudes
    resultados.Add EjecutarPruebasSolicitudes()
    
    ' Pruebas de Búsqueda
    resultados.Add EjecutarPruebasBusqueda()
    
    ' Pruebas de Documentos
    resultados.Add EjecutarPruebasDocumentos()
    
    ' Pruebas de Integración
    resultados.Add EjecutarPruebasIntegracion()
    
    ' Generar reporte final
    GenerarReportePruebas resultados
End Sub

Public Function EjecutarPruebasAuth() As T_ResultadoPrueba
    Dim resultado As New T_ResultadoPrueba
    resultado.Modulo = "Autenticación"
    
    ' CP-AUTH-001: Login exitoso
    resultado.Casos.Add EjecutarCP_AUTH_001()
    
    ' CP-AUTH-002: Usuario inactivo
    resultado.Casos.Add EjecutarCP_AUTH_002()
    
    ' CP-AUTH-003: Permisos por rol
    resultado.Casos.Add EjecutarCP_AUTH_003()
    
    Set EjecutarPruebasAuth = resultado
End Function
```

### 6.2 Datos de Prueba

```sql
-- Script: datos_prueba.sql
-- Crear datos de prueba para testing

-- Usuarios de prueba
INSERT INTO TbUsuarios (UsuarioId, Email, Nombre, Rol, Estado) VALUES
('USR-TEST-001', 'test.calidad@empresa.com', 'Usuario Test Calidad', 'Calidad', 'Activo'),
('USR-TEST-002', 'test.ingenieria@empresa.com', 'Usuario Test Ingeniería', 'Ingeniería', 'Activo'),
('USR-TEST-003', 'test.admin@empresa.com', 'Usuario Test Admin', 'Administrador', 'Activo'),
('USR-TEST-004', 'test.inactivo@empresa.com', 'Usuario Test Inactivo', 'Calidad', 'Inactivo');

-- Solicitudes de prueba
INSERT INTO Tb_Solicitudes (SolicitudId, NumeroExpediente, TipoSolicitud, DescripcionSolicitud, EstadoInterno, UsuarioCreador, FechaCreacion) VALUES
('SOL-TEST-001', 'EXP-TEST-001', 'PC', 'Solicitud de prueba PC', 'Borrador', 'test.calidad@empresa.com', #2024-12-20#),
('SOL-TEST-002', 'EXP-TEST-002', 'CD_CA', 'Solicitud de prueba CD_CA', 'Enviado', 'test.calidad@empresa.com', #2024-12-19#),
('SOL-TEST-003', 'EXP-TEST-003', 'CD_CA_SUB', 'Solicitud de prueba CD_CA_SUB', 'Aprobado', 'test.calidad@empresa.com', #2024-12-18#);

-- Configuración de prueba
INSERT INTO TbConfiguracion (Clave, Valor, Descripcion) VALUES
('TEST_MODE', 'true', 'Modo de pruebas activado'),
('MOCK_EXPEDIENTES', 'true', 'Usar mock para servicio de expedientes'),
('LOG_LEVEL', 'DEBUG', 'Nivel de logging para pruebas');
```

---

## 7. Criterios de Finalización

### 7.1 Criterios de Éxito

- **Cobertura de Pruebas**: Mínimo 90% de funcionalidades cubiertas
- **Tasa de Éxito**: Mínimo 95% de casos de prueba exitosos
- **Performance**: Todas las operaciones dentro de tiempos esperados
- **Seguridad**: Cero vulnerabilidades críticas o altas
- **Usabilidad**: Flujos de usuario completables sin asistencia

### 7.2 Entregables de Pruebas

- [ ] Casos de prueba ejecutados y documentados
- [ ] Reporte de resultados con evidencias
- [ ] Lista de defectos encontrados y su estado
- [ ] Recomendaciones para mejoras
- [ ] Suite de pruebas automatizadas funcional
- [ ] Datos de prueba configurados
- [ ] Documentación de procedimientos de testing

---

*Documento generado según la Especificación Funcional y Arquitectura CONDOR*  
*Versión: 1.0*  
*Fecha: Diciembre 2024*