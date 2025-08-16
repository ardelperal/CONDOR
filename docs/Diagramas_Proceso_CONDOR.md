# Diagramas de Proceso Visual - Sistema CONDOR

## Introducción

Este documento presenta los diagramas de proceso visual para los flujos de trabajo principales del Sistema CONDOR, basados en la especificación funcional del proyecto.

---

## 1. Flujo de Creación de Solicitud

### Descripción
Proceso completo desde el inicio de una nueva solicitud hasta su registro en el sistema.

### Diagrama de Flujo

```svg
<svg width="800" height="600" xmlns="http://www.w3.org/2000/svg">
  <!-- Título -->
  <text x="400" y="30" text-anchor="middle" font-size="18" font-weight="bold" fill="#2c3e50">Flujo de Creación de Solicitud CONDOR</text>
  
  <!-- Inicio -->
  <ellipse cx="400" cy="70" rx="60" ry="25" fill="#27ae60" stroke="#1e8449" stroke-width="2"/>
  <text x="400" y="77" text-anchor="middle" font-size="12" fill="white">INICIO</text>
  
  <!-- Login y Autenticación -->
  <rect x="320" y="120" width="160" height="40" rx="5" fill="#3498db" stroke="#2980b9" stroke-width="2"/>
  <text x="400" y="143" text-anchor="middle" font-size="11" fill="white">Login y Autenticación</text>
  
  <!-- Verificar Rol -->
  <polygon points="350,190 450,190 470,220 450,250 350,250 330,220" fill="#f39c12" stroke="#e67e22" stroke-width="2"/>
  <text x="400" y="223" text-anchor="middle" font-size="10" fill="white">¿Rol Calidad/Admin?</text>
  
  <!-- Acceso Denegado -->
  <rect x="520" y="200" width="120" height="40" rx="5" fill="#e74c3c" stroke="#c0392b" stroke-width="2"/>
  <text x="580" y="223" text-anchor="middle" font-size="10" fill="white">Acceso Denegado</text>
  
  <!-- Seleccionar Expediente -->
  <rect x="320" y="290" width="160" height="40" rx="5" fill="#9b59b6" stroke="#8e44ad" stroke-width="2"/>
  <text x="400" y="313" text-anchor="middle" font-size="11" fill="white">Seleccionar Expediente</text>
  
  <!-- Tipo de Solicitud -->
  <polygon points="350,360 450,360 470,390 450,420 350,420 330,390" fill="#f39c12" stroke="#e67e22" stroke-width="2"/>
  <text x="400" y="393" text-anchor="middle" font-size="10" fill="white">Tipo Solicitud?</text>
  
  <!-- Tipos de Solicitud -->
  <rect x="150" y="460" width="100" height="30" rx="3" fill="#16a085" stroke="#138d75" stroke-width="1"/>
  <text x="200" y="478" text-anchor="middle" font-size="10" fill="white">PC</text>
  
  <rect x="350" y="460" width="100" height="30" rx="3" fill="#16a085" stroke="#138d75" stroke-width="1"/>
  <text x="400" y="478" text-anchor="middle" font-size="10" fill="white">CD_CA</text>
  
  <rect x="550" y="460" width="100" height="30" rx="3" fill="#16a085" stroke="#138d75" stroke-width="1"/>
  <text x="600" y="478" text-anchor="middle" font-size="10" fill="white">CD_CA_SUB</text>
  
  <!-- Guardar Solicitud -->
  <rect x="320" y="520" width="160" height="40" rx="5" fill="#2ecc71" stroke="#27ae60" stroke-width="2"/>
  <text x="400" y="543" text-anchor="middle" font-size="11" fill="white">Guardar Solicitud</text>
  
  <!-- Fin -->
  <ellipse cx="400" cy="590" rx="50" ry="20" fill="#e74c3c" stroke="#c0392b" stroke-width="2"/>
  <text x="400" y="596" text-anchor="middle" font-size="11" fill="white">FIN</text>
  
  <!-- Flechas -->
  <defs>
    <marker id="arrowhead" markerWidth="10" markerHeight="7" refX="9" refY="3.5" orient="auto">
      <polygon points="0 0, 10 3.5, 0 7" fill="#34495e"/>
    </marker>
  </defs>
  
  <!-- Conexiones -->
  <line x1="400" y1="95" x2="400" y2="120" stroke="#34495e" stroke-width="2" marker-end="url(#arrowhead)"/>
  <line x1="400" y1="160" x2="400" y2="190" stroke="#34495e" stroke-width="2" marker-end="url(#arrowhead)"/>
  <line x1="470" y1="220" x2="520" y2="220" stroke="#34495e" stroke-width="2" marker-end="url(#arrowhead)"/>
  <line x1="400" y1="250" x2="400" y2="290" stroke="#34495e" stroke-width="2" marker-end="url(#arrowhead)"/>
  <line x1="400" y1="330" x2="400" y2="360" stroke="#34495e" stroke-width="2" marker-end="url(#arrowhead)"/>
  <line x1="350" y1="390" x2="200" y2="460" stroke="#34495e" stroke-width="2" marker-end="url(#arrowhead)"/>
  <line x1="400" y1="420" x2="400" y2="460" stroke="#34495e" stroke-width="2" marker-end="url(#arrowhead)"/>
  <line x1="450" y1="390" x2="600" y2="460" stroke="#34495e" stroke-width="2" marker-end="url(#arrowhead)"/>
  <line x1="200" y1="490" x2="400" y2="520" stroke="#34495e" stroke-width="2" marker-end="url(#arrowhead)"/>
  <line x1="400" y1="490" x2="400" y2="520" stroke="#34495e" stroke-width="2" marker-end="url(#arrowhead)"/>
  <line x1="600" y1="490" x2="400" y2="520" stroke="#34495e" stroke-width="2" marker-end="url(#arrowhead)"/>
  <line x1="400" y1="560" x2="400" y2="570" stroke="#34495e" stroke-width="2" marker-end="url(#arrowhead)"/>
  
  <!-- Etiquetas -->
  <text x="485" y="215" font-size="9" fill="#e74c3c">NO</text>
  <text x="405" y="275" font-size="9" fill="#27ae60">SÍ</text>
</svg>
```

---

## 2. Flujo de Seguimiento de Solicitud

### Descripción
Proceso de seguimiento y actualización del estado de una solicitud existente.

### Diagrama de Flujo

```svg
<svg width="800" height="700" xmlns="http://www.w3.org/2000/svg">
  <!-- Título -->
  <text x="400" y="30" text-anchor="middle" font-size="18" font-weight="bold" fill="#2c3e50">Flujo de Seguimiento de Solicitud CONDOR</text>
  
  <!-- Inicio -->
  <ellipse cx="400" cy="70" rx="60" ry="25" fill="#27ae60" stroke="#1e8449" stroke-width="2"/>
  <text x="400" y="77" text-anchor="middle" font-size="12" fill="white">INICIO</text>
  
  <!-- Buscar Solicitud -->
  <rect x="320" y="120" width="160" height="40" rx="5" fill="#3498db" stroke="#2980b9" stroke-width="2"/>
  <text x="400" y="143" text-anchor="middle" font-size="11" fill="white">Buscar Solicitud</text>
  
  <!-- ¿Existe? -->
  <polygon points="350,190 450,190 470,220 450,250 350,250 330,220" fill="#f39c12" stroke="#e67e22" stroke-width="2"/>
  <text x="400" y="223" text-anchor="middle" font-size="10" fill="white">¿Existe?</text>
  
  <!-- No Encontrada -->
  <rect x="520" y="200" width="120" height="40" rx="5" fill="#e74c3c" stroke="#c0392b" stroke-width="2"/>
  <text x="580" y="223" text-anchor="middle" font-size="10" fill="white">No Encontrada</text>
  
  <!-- Verificar Permisos -->
  <polygon points="350,290 450,290 470,320 450,350 350,350 330,320" fill="#9b59b6" stroke="#8e44ad" stroke-width="2"/>
  <text x="400" y="323" text-anchor="middle" font-size="10" fill="white">¿Permisos?</text>
  
  <!-- Sin Permisos -->
  <rect x="520" y="300" width="120" height="40" rx="5" fill="#e74c3c" stroke="#c0392b" stroke-width="2"/>
  <text x="580" y="323" text-anchor="middle" font-size="10" fill="white">Sin Permisos</text>
  
  <!-- Estado Actual -->
  <polygon points="350,390 450,390 470,420 450,450 350,450 330,420" fill="#f39c12" stroke="#e67e22" stroke-width="2"/>
  <text x="400" y="423" text-anchor="middle" font-size="10" fill="white">Estado Actual?</text>
  
  <!-- Estados -->
  <rect x="100" y="490" width="80" height="30" rx="3" fill="#16a085" stroke="#138d75" stroke-width="1"/>
  <text x="140" y="508" text-anchor="middle" font-size="9" fill="white">Borrador</text>
  
  <rect x="220" y="490" width="80" height="30" rx="3" fill="#f39c12" stroke="#e67e22" stroke-width="1"/>
  <text x="260" y="508" text-anchor="middle" font-size="9" fill="white">En Revisión</text>
  
  <rect x="340" y="490" width="80" height="30" rx="3" fill="#3498db" stroke="#2980b9" stroke-width="1"/>
  <text x="380" y="508" text-anchor="middle" font-size="9" fill="white">Aprobado</text>
  
  <rect x="460" y="490" width="80" height="30" rx="3" fill="#9b59b6" stroke="#8e44ad" stroke-width="1"/>
  <text x="500" y="508" text-anchor="middle" font-size="9" fill="white">Enviado</text>
  
  <rect x="580" y="490" width="80" height="30" rx="3" fill="#27ae60" stroke="#1e8449" stroke-width="1"/>
  <text x="620" y="508" text-anchor="middle" font-size="9" fill="white">Cerrado</text>
  
  <!-- Actualizar Estado -->
  <rect x="320" y="550" width="160" height="40" rx="5" fill="#2ecc71" stroke="#27ae60" stroke-width="2"/>
  <text x="400" y="573" text-anchor="middle" font-size="11" fill="white">Actualizar Estado</text>
  
  <!-- Notificar -->
  <rect x="320" y="610" width="160" height="40" rx="5" fill="#e67e22" stroke="#d35400" stroke-width="2"/>
  <text x="400" y="633" text-anchor="middle" font-size="11" fill="white">Enviar Notificación</text>
  
  <!-- Fin -->
  <ellipse cx="400" cy="680" rx="50" ry="20" fill="#e74c3c" stroke="#c0392b" stroke-width="2"/>
  <text x="400" y="686" text-anchor="middle" font-size="11" fill="white">FIN</text>
  
  <!-- Flechas -->
  <line x1="400" y1="95" x2="400" y2="120" stroke="#34495e" stroke-width="2" marker-end="url(#arrowhead)"/>
  <line x1="400" y1="160" x2="400" y2="190" stroke="#34495e" stroke-width="2" marker-end="url(#arrowhead)"/>
  <line x1="470" y1="220" x2="520" y2="220" stroke="#34495e" stroke-width="2" marker-end="url(#arrowhead)"/>
  <line x1="400" y1="250" x2="400" y2="290" stroke="#34495e" stroke-width="2" marker-end="url(#arrowhead)"/>
  <line x1="470" y1="320" x2="520" y2="320" stroke="#34495e" stroke-width="2" marker-end="url(#arrowhead)"/>
  <line x1="400" y1="350" x2="400" y2="390" stroke="#34495e" stroke-width="2" marker-end="url(#arrowhead)"/>
  
  <!-- Conexiones a estados -->
  <line x1="350" y1="420" x2="140" y2="490" stroke="#34495e" stroke-width="1" marker-end="url(#arrowhead)"/>
  <line x1="370" y1="430" x2="260" y2="490" stroke="#34495e" stroke-width="1" marker-end="url(#arrowhead)"/>
  <line x1="400" y1="450" x2="380" y2="490" stroke="#34495e" stroke-width="1" marker-end="url(#arrowhead)"/>
  <line x1="430" y1="430" x2="500" y2="490" stroke="#34495e" stroke-width="1" marker-end="url(#arrowhead)"/>
  <line x1="450" y1="420" x2="620" y2="490" stroke="#34495e" stroke-width="1" marker-end="url(#arrowhead)"/>
  
  <!-- Convergencia -->
  <line x1="140" y1="520" x2="400" y2="550" stroke="#34495e" stroke-width="1" marker-end="url(#arrowhead)"/>
  <line x1="260" y1="520" x2="400" y2="550" stroke="#34495e" stroke-width="1" marker-end="url(#arrowhead)"/>
  <line x1="380" y1="520" x2="400" y2="550" stroke="#34495e" stroke-width="1" marker-end="url(#arrowhead)"/>
  <line x1="500" y1="520" x2="400" y2="550" stroke="#34495e" stroke-width="1" marker-end="url(#arrowhead)"/>
  <line x1="620" y1="520" x2="400" y2="550" stroke="#34495e" stroke-width="1" marker-end="url(#arrowhead)"/>
  
  <line x1="400" y1="590" x2="400" y2="610" stroke="#34495e" stroke-width="2" marker-end="url(#arrowhead)"/>
  <line x1="400" y1="650" x2="400" y2="660" stroke="#34495e" stroke-width="2" marker-end="url(#arrowhead)"/>
  
  <!-- Etiquetas -->
  <text x="485" y="215" font-size="9" fill="#e74c3c">NO</text>
  <text x="405" y="275" font-size="9" fill="#27ae60">SÍ</text>
  <text x="485" y="315" font-size="9" fill="#e74c3c">NO</text>
  <text x="405" y="375" font-size="9" fill="#27ae60">SÍ</text>
</svg>
```

---

## 3. Flujo de Cierre de Solicitud

### Descripción
Proceso de finalización y cierre de una solicitud, incluyendo generación de documentos finales.

### Diagrama de Flujo

```svg
<svg width="800" height="650" xmlns="http://www.w3.org/2000/svg">
  <!-- Título -->
  <text x="400" y="30" text-anchor="middle" font-size="18" font-weight="bold" fill="#2c3e50">Flujo de Cierre de Solicitud CONDOR</text>
  
  <!-- Inicio -->
  <ellipse cx="400" cy="70" rx="60" ry="25" fill="#27ae60" stroke="#1e8449" stroke-width="2"/>
  <text x="400" y="77" text-anchor="middle" font-size="12" fill="white">INICIO</text>
  
  <!-- Verificar Estado -->
  <polygon points="350,120 450,120 470,150 450,180 350,180 330,150" fill="#f39c12" stroke="#e67e22" stroke-width="2"/>
  <text x="400" y="153" text-anchor="middle" font-size="10" fill="white">¿Estado Enviado?</text>
  
  <!-- Estado Incorrecto -->
  <rect x="520" y="130" width="120" height="40" rx="5" fill="#e74c3c" stroke="#c0392b" stroke-width="2"/>
  <text x="580" y="153" text-anchor="middle" font-size="10" fill="white">Estado Incorrecto</text>
  
  <!-- Validar Documentos -->
  <rect x="320" y="220" width="160" height="40" rx="5" fill="#3498db" stroke="#2980b9" stroke-width="2"/>
  <text x="400" y="243" text-anchor="middle" font-size="11" fill="white">Validar Documentos</text>
  
  <!-- ¿Documentos OK? -->
  <polygon points="350,290 450,290 470,320 450,350 350,350 330,320" fill="#9b59b6" stroke="#8e44ad" stroke-width="2"/>
  <text x="400" y="323" text-anchor="middle" font-size="10" fill="white">¿Docs OK?</text>
  
  <!-- Documentos Faltantes -->
  <rect x="520" y="300" width="120" height="40" rx="5" fill="#e67e22" stroke="#d35400" stroke-width="2"/>
  <text x="580" y="323" text-anchor="middle" font-size="10" fill="white">Docs Faltantes</text>
  
  <!-- Generar Documento Final -->
  <rect x="320" y="390" width="160" height="40" rx="5" fill="#16a085" stroke="#138d75" stroke-width="2"/>
  <text x="400" y="413" text-anchor="middle" font-size="11" fill="white">Generar Doc Final</text>
  
  <!-- Archivar Solicitud -->
  <rect x="320" y="460" width="160" height="40" rx="5" fill="#8e44ad" stroke="#7d3c98" stroke-width="2"/>
  <text x="400" y="483" text-anchor="middle" font-size="11" fill="white">Archivar Solicitud</text>
  
  <!-- Actualizar Estado a Cerrado -->
  <rect x="320" y="530" width="160" height="40" rx="5" fill="#27ae60" stroke="#1e8449" stroke-width="2"/>
  <text x="400" y="553" text-anchor="middle" font-size="11" fill="white">Estado = Cerrado</text>
  
  <!-- Notificar Cierre -->
  <rect x="320" y="590" width="160" height="40" rx="5" fill="#e67e22" stroke="#d35400" stroke-width="2"/>
  <text x="400" y="613" text-anchor="middle" font-size="11" fill="white">Notificar Cierre</text>
  
  <!-- Fin -->
  <ellipse cx="400" cy="650" rx="50" ry="20" fill="#e74c3c" stroke="#c0392b" stroke-width="2"/>
  <text x="400" y="656" text-anchor="middle" font-size="11" fill="white">FIN</text>
  
  <!-- Flechas -->
  <line x1="400" y1="95" x2="400" y2="120" stroke="#34495e" stroke-width="2" marker-end="url(#arrowhead)"/>
  <line x1="470" y1="150" x2="520" y2="150" stroke="#34495e" stroke-width="2" marker-end="url(#arrowhead)"/>
  <line x1="400" y1="180" x2="400" y2="220" stroke="#34495e" stroke-width="2" marker-end="url(#arrowhead)"/>
  <line x1="400" y1="260" x2="400" y2="290" stroke="#34495e" stroke-width="2" marker-end="url(#arrowhead)"/>
  <line x1="470" y1="320" x2="520" y2="320" stroke="#34495e" stroke-width="2" marker-end="url(#arrowhead)"/>
  <line x1="400" y1="350" x2="400" y2="390" stroke="#34495e" stroke-width="2" marker-end="url(#arrowhead)"/>
  <line x1="400" y1="430" x2="400" y2="460" stroke="#34495e" stroke-width="2" marker-end="url(#arrowhead)"/>
  <line x1="400" y1="500" x2="400" y2="530" stroke="#34495e" stroke-width="2" marker-end="url(#arrowhead)"/>
  <line x1="400" y1="570" x2="400" y2="590" stroke="#34495e" stroke-width="2" marker-end="url(#arrowhead)"/>
  <line x1="400" y1="630" x2="400" y2="630" stroke="#34495e" stroke-width="2" marker-end="url(#arrowhead)"/>
  
  <!-- Etiquetas -->
  <text x="485" y="145" font-size="9" fill="#e74c3c">NO</text>
  <text x="405" y="205" font-size="9" fill="#27ae60">SÍ</text>
  <text x="485" y="315" font-size="9" fill="#e74c3c">NO</text>
  <text x="405" y="375" font-size="9" fill="#27ae60">SÍ</text>
</svg>
```

---

## 4. Matriz de Estados y Transiciones

### Descripción
Tabla que muestra los estados posibles de una solicitud y las transiciones permitidas según el rol del usuario.

| Estado Origen | Estado Destino | Rol Requerido | Condiciones |
|---------------|----------------|---------------|-------------|
| - | Borrador | Calidad/Admin | Nueva solicitud |
| Borrador | En Revisión | Calidad/Admin | Datos completos |
| En Revisión | Borrador | Calidad/Admin | Correcciones necesarias |
| En Revisión | Aprobado | Calidad/Admin | Revisión exitosa |
| Aprobado | Enviado | Calidad/Admin | Documentos generados |
| Enviado | Cerrado | Calidad/Admin | Proceso completado |
| Cualquiera | Cancelado | Admin | Cancelación autorizada |

---

## 5. Actores y Responsabilidades

### Roles del Sistema

#### Calidad
- **Responsabilidades**: Crear, gestionar y cerrar solicitudes
- **Permisos**: Acceso completo a todas las fases
- **Flujos**: Todos los diagramas anteriores

#### Ingeniería (Técnico)
- **Responsabilidades**: Completar detalles técnicos
- **Permisos**: Edición limitada según fase
- **Flujos**: Solo seguimiento y actualización

#### Administrador
- **Responsabilidades**: Configuración del sistema
- **Permisos**: Todos los permisos + configuración
- **Flujos**: Todos + gestión de usuarios

---

## 6. Integración con Sistemas Externos

### ExpedienteService
- **Propósito**: Obtener datos de expedientes
- **Datos**: Nemotécnico, responsables, contratista
- **Momento**: Al crear nueva solicitud

### Sistema de Lanzadera
- **Propósito**: Autenticación y actualización
- **Datos**: Email del usuario
- **Momento**: Al iniciar la aplicación

### Sistema RAC
- **Propósito**: Sincronización de estados
- **Datos**: Estado externo de solicitudes
- **Momento**: Según configuración

---

*Documento generado según la Especificación Funcional CONDOR*  
*Versión: 1.0*  
*Fecha: Diciembre 2024*