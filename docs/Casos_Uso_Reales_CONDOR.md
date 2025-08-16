# Casos de Uso Reales - Sistema CONDOR

## Introducción

Este documento presenta casos de uso reales del Sistema CONDOR con datos específicos del dominio de contratos públicos, basados en la especificación funcional del proyecto.

---

## Caso de Uso 1: Solicitud de Cambio de Precio (PC)

### Contexto del Expediente
- **Número de Expediente**: EXP-2024-INF-001
- **Título**: "Suministro e instalación de equipos informáticos para centros educativos"
- **Contratista Principal**: Tecnología Avanzada S.L.
- **Responsable de Calidad**: María García López (maria.garcia@empresa.com)
- **Jefe de Proyecto**: Carlos Rodríguez Martín
- **Valor del Contrato**: 450.000 €
- **Estado del Expediente**: En Ejecución

### Descripción del Caso
El contratista solicita un cambio de precio debido al incremento del coste de los componentes electrónicos por la crisis de semiconductores.

### Datos de la Solicitud PC

#### Información General
- **ID Solicitud**: SOL-2024-001
- **Tipo**: PC (Precio)
- **Fecha Creación**: 15/03/2024 09:30
- **Usuario Creador**: maria.garcia@empresa.com
- **Estado Interno**: Borrador
- **Estado RAC**: Pendiente

#### Datos Específicos PC (T_Datos_PC)
```json
{
  "DescripcionCambio": "Incremento del precio unitario de tablets educativas debido al aumento del coste de semiconductores",
  "JustificacionTecnica": "Los proveedores han notificado incrementos del 15% en componentes críticos (procesadores ARM, memoria flash) debido a la escasez global de semiconductores iniciada en Q4 2023",
  "JustificacionEconomica": "El incremento propuesto de 67.500€ representa el 15% del valor de las tablets (450 unidades x 100€ incremento unitario)",
  "ImporteOriginal": 450000.00,
  "ImporteNuevo": 517500.00,
  "DiferenciaImporte": 67500.00,
  "PorcentajeIncremento": 15.0,
  "FechaEfecto": "2024-04-01",
  "DocumentacionAdjunta": "Certificados proveedores, análisis mercado semiconductores, propuesta técnica revisada",
  "ObservacionesCalidad": "Solicitud justificada técnica y económicamente. Verificar disponibilidad presupuestaria",
  "ObservacionesIngenieria": "Componentes alternativos evaluados. No existen opciones viables con menor coste",
  "FechaRevisionTecnica": "2024-03-18",
  "ResponsableTecnico": "juan.martinez@empresa.com"
}
```

### Flujo del Proceso

1. **Creación (15/03/2024)**
   - María García recibe comunicación del contratista
   - Accede a CONDOR con rol Calidad
   - Selecciona expediente EXP-2024-INF-001
   - Crea nueva solicitud tipo PC
   - Completa datos iniciales y guarda como Borrador

2. **Revisión Técnica (16/03/2024)**
   - Sistema notifica a Juan Martínez (Ingeniería)
   - Juan completa campos técnicos y observaciones
   - Cambia estado a "En Revisión"

3. **Aprobación (20/03/2024)**
   - María revisa completitud de datos
   - Genera documento preliminar usando plantilla PC
   - Cambia estado a "Aprobado"

4. **Envío (22/03/2024)**
   - Genera documento final con mapeo de campos
   - Adjunta documentación de soporte
   - Cambia estado a "Enviado"
   - Sistema notifica a todas las partes

---

## Caso de Uso 2: Solicitud de Concesión/Desviación (CD_CA)

### Contexto del Expediente
- **Número de Expediente**: EXP-2024-OBR-045
- **Título**: "Construcción de centro de salud municipal"
- **Contratista Principal**: Construcciones Mediterráneo S.A.
- **Responsable de Calidad**: Ana Fernández Ruiz (ana.fernandez@empresa.com)
- **Jefe de Proyecto**: Miguel Ángel Torres
- **Valor del Contrato**: 2.850.000 €
- **Estado del Expediente**: En Ejecución

### Descripción del Caso
Se detecta una desviación en las especificaciones del sistema de climatización que requiere aprobación para continuar con la obra.

### Datos de la Solicitud CD_CA

#### Información General
- **ID Solicitud**: SOL-2024-015
- **Tipo**: CD_CA (Concesión/Desviación)
- **Fecha Creación**: 08/05/2024 14:15
- **Usuario Creador**: ana.fernandez@empresa.com
- **Estado Interno**: En Revisión
- **Estado RAC**: En Proceso

#### Datos Específicos CD_CA (T_Datos_CD_CA)
```json
{
  "TipoDesviacion": "Técnica",
  "ClasificacionRiesgo": "Medio",
  "DescripcionDesviacion": "Sustitución de sistema de climatización VRV por sistema centralizado con mayor eficiencia energética",
  "CausaRaiz": "Discontinuación del modelo especificado por el fabricante. Nuevo sistema ofrece mejor rendimiento energético",
  "ImpactoTecnico": "Mejora en eficiencia energética del 20%. Reducción de ruido operacional. Mantenimiento simplificado",
  "ImpactoEconomico": "Incremento inicial de 45.000€ compensado con ahorro operacional de 8.000€/año",
  "ImpactoTemporal": "Retraso de 2 semanas en instalación por cambio de diseño",
  "MedidasCorrectivas": "Actualización de planos, formación adicional a técnicos, revisión de cronograma",
  "MedidasPreventivas": "Verificación de disponibilidad de equipos antes de pedidos futuros",
  "FechaDeteccion": "2024-05-05",
  "FechaNotificacion": "2024-05-06",
  "ResponsableDeteccion": "supervisor.obra@construcciones.com",
  "EstadoImplementacion": "Pendiente Aprobación",
  "DocumentacionSoporte": "Certificados técnicos nuevo equipo, análisis comparativo, cronograma revisado",
  "AprobacionRequerida": true,
  "NivelAprobacion": "Dirección Técnica",
  "FechaLimiteRespuesta": "2024-05-20"
}
```

### Flujo del Proceso

1. **Detección y Notificación (05-06/05/2024)**
   - Supervisor de obra detecta discontinuación del equipo
   - Notifica a Ana Fernández (Calidad)
   - Ana evalúa impacto y necesidad de solicitud formal

2. **Creación de Solicitud (08/05/2024)**
   - Ana crea solicitud CD_CA en CONDOR
   - Completa análisis inicial de impactos
   - Solicita revisión técnica a Ingeniería

3. **Análisis Técnico (09-12/05/2024)**
   - Ingeniería evalúa alternativa propuesta
   - Valida mejoras de eficiencia energética
   - Confirma viabilidad técnica

4. **Revisión y Aprobación (En proceso)**
   - Pendiente aprobación de Dirección Técnica
   - Evaluación de impacto presupuestario
   - Decisión esperada antes del 20/05/2024

---

## Caso de Uso 3: Solicitud de Sub-Concesión (CD_CA_SUB)

### Contexto del Expediente
- **Número de Expediente**: EXP-2024-SER-012
- **Título**: "Servicios de limpieza y mantenimiento de edificios municipales"
- **Contratista Principal**: Servicios Integrales Norte S.L.
- **Responsable de Calidad**: Pedro Sánchez Vila (pedro.sanchez@empresa.com)
- **Jefe de Proyecto**: Laura Martín González
- **Valor del Contrato**: 180.000 €/año
- **Estado del Expediente**: En Ejecución

### Descripción del Caso
Solicitud de sub-concesión para modificar el protocolo de limpieza en época de pandemia, derivada de una concesión principal ya aprobada.

### Datos de la Solicitud CD_CA_SUB

#### Información General
- **ID Solicitud**: SOL-2024-023
- **Tipo**: CD_CA_SUB (Sub-Concesión)
- **Fecha Creación**: 12/06/2024 11:20
- **Usuario Creador**: pedro.sanchez@empresa.com
- **Estado Interno**: Aprobado
- **Estado RAC**: Aprobado

#### Datos Específicos CD_CA_SUB (T_Datos_CD_CA_SUB)
```json
{
  "SolicitudPadreId": "SOL-2024-018",
  "TipoSubconcesion": "Modificación Protocolo",
  "RelacionConPadre": "Extensión de concesión COVID-19 para incluir nuevos protocolos sanitarios",
  "DescripcionModificacion": "Implementación de protocolo de desinfección con ozono en áreas de alto tránsito",
  "JustificacionAdicional": "Nuevas recomendaciones sanitarias requieren desinfección adicional con ozono cada 4 horas",
  "ImpactoSobrePadre": "Complementa protocolo base sin interferir con actividades principales",
  "RecursosAdicionales": "1 equipo generador de ozono, 2 horas técnico especializado/día",
  "CostesAdicionales": 2400.00,
  "PeriodoImplementacion": "3 meses (julio-septiembre 2024)",
  "IndicadoresExito": "Reducción 95% carga viral superficies, cumplimiento protocolo 100%",
  "RiesgosIdentificados": "Posible irritación si no se respetan tiempos de ventilación",
  "MedidasMitigacion": "Señalización clara, formación personal, cronograma estricto",
  "AprobacionesRequeridas": "Servicio Prevención Riesgos Laborales, Dirección Servicios",
  "DocumentacionReferencia": "Protocolo sanitario municipal, fichas técnicas ozono, plan formación",
  "FechaInicioEfecto": "2024-07-01",
  "FechaFinEfecto": "2024-09-30",
  "ResponsableImplementacion": "coordinador.limpieza@servicios.com",
  "EstadoValidacion": "Validado por Prevención"
}
```

### Flujo del Proceso

1. **Identificación de Necesidad (10/06/2024)**
   - Nuevas directrices sanitarias municipales
   - Pedro evalúa necesidad de modificación
   - Verifica relación con concesión padre SOL-2024-018

2. **Creación y Documentación (12/06/2024)**
   - Pedro crea solicitud CD_CA_SUB
   - Vincula con solicitud padre
   - Completa análisis de impacto y recursos

3. **Validación Técnica (13-15/06/2024)**
   - Servicio de Prevención valida protocolo
   - Ingeniería confirma viabilidad técnica
   - Dirección de Servicios aprueba recursos

4. **Aprobación Final (18/06/2024)**
   - Todas las validaciones completadas
   - Estado cambiado a "Aprobado"
   - Implementación programada para julio

---

## Caso de Uso 4: Búsqueda y Consulta Avanzada

### Escenario
Ana Fernández necesita localizar todas las solicitudes relacionadas con el contratista "Construcciones Mediterráneo S.A." para preparar un informe de seguimiento.

### Criterios de Búsqueda
- **Contratista**: "Construcciones Mediterráneo S.A."
- **Período**: Enero 2024 - Junio 2024
- **Estados**: Todos excepto "Cancelado"
- **Tipos**: Todos (PC, CD_CA, CD_CA_SUB)

### Resultados Esperados
```json
{
  "total_solicitudes": 8,
  "solicitudes": [
    {
      "id": "SOL-2024-015",
      "expediente": "EXP-2024-OBR-045",
      "tipo": "CD_CA",
      "estado": "En Revisión",
      "fecha_creacion": "2024-05-08",
      "descripcion": "Desviación sistema climatización"
    },
    {
      "id": "SOL-2024-008",
      "expediente": "EXP-2024-OBR-032",
      "tipo": "PC",
      "estado": "Cerrado",
      "fecha_creacion": "2024-02-15",
      "descripcion": "Cambio precio materiales construcción"
    }
  ],
  "resumen_por_tipo": {
    "PC": 3,
    "CD_CA": 4,
    "CD_CA_SUB": 1
  },
  "resumen_por_estado": {
    "Cerrado": 5,
    "En Revisión": 2,
    "Aprobado": 1
  }
}
```

---

## Caso de Uso 5: Generación de Documentos

### Escenario
Generación automática de documento oficial para la solicitud SOL-2024-001 (Cambio de Precio).

### Proceso de Mapeo
El sistema utiliza la tabla `Tb_Mapeo_Campos` para mapear los datos de la solicitud a la plantilla Word correspondiente.

#### Mapeo para Plantilla PC
```json
{
  "mapeos_aplicados": [
    {
      "campo_origen": "NumeroExpediente",
      "marcador_plantilla": "{{NUMERO_EXPEDIENTE}}",
      "valor": "EXP-2024-INF-001"
    },
    {
      "campo_origen": "DescripcionCambio",
      "marcador_plantilla": "{{DESCRIPCION_CAMBIO}}",
      "valor": "Incremento del precio unitario de tablets educativas..."
    },
    {
      "campo_origen": "ImporteOriginal",
      "marcador_plantilla": "{{IMPORTE_ORIGINAL}}",
      "valor": "450.000,00 €"
    },
    {
      "campo_origen": "ImporteNuevo",
      "marcador_plantilla": "{{IMPORTE_NUEVO}}",
      "valor": "517.500,00 €"
    },
    {
      "campo_origen": "PorcentajeIncremento",
      "marcador_plantilla": "{{PORCENTAJE_INCREMENTO}}",
      "valor": "15,0%"
    }
  ],
  "documento_generado": "SOL-2024-001_Cambio_Precio_v1.0.docx",
  "ruta_almacenamiento": "\\servidor\condor\documentos\2024\03\SOL-2024-001_Cambio_Precio_v1.0.docx",
  "fecha_generacion": "2024-03-22 10:15:30",
  "usuario_generador": "maria.garcia@empresa.com"
}
```

---

## Caso de Uso 6: Notificaciones Automáticas

### Escenario
Cambio de estado de solicitud SOL-2024-015 de "Borrador" a "En Revisión".

### Notificaciones Generadas

#### Para Ingeniería (Revisión Técnica)
```json
{
  "destinatario": "juan.martinez@empresa.com",
  "asunto": "CONDOR: Nueva solicitud pendiente de revisión técnica - SOL-2024-015",
  "cuerpo": "Se ha creado una nueva solicitud que requiere su revisión técnica:\n\nID: SOL-2024-015\nTipo: CD_CA (Concesión/Desviación)\nExpediente: EXP-2024-OBR-045\nDescripción: Desviación sistema climatización\nCreada por: Ana Fernández\nFecha límite: 2024-05-20\n\nAcceda a CONDOR para completar la revisión.",
  "prioridad": "Alta",
  "fecha_envio": "2024-05-08 14:20:00"
}
```

#### Para Jefe de Proyecto (Información)
```json
{
  "destinatario": "miguel.torres@empresa.com",
  "asunto": "CONDOR: Solicitud en revisión - EXP-2024-OBR-045",
  "cuerpo": "Se ha iniciado el proceso de revisión para una solicitud en su expediente:\n\nExpediente: EXP-2024-OBR-045 - Construcción centro de salud\nSolicitud: SOL-2024-015\nTipo: Desviación técnica\nEstado: En Revisión\n\nPodrá consultar el progreso en CONDOR.",
  "prioridad": "Media",
  "fecha_envio": "2024-05-08 14:25:00"
}
```

---

## Caso de Uso 7: Integración con ExpedienteService

### Escenario
Creación de nueva solicitud requiere obtener datos del expediente desde sistema externo.

### Llamada al Servicio
```vba
' Ejemplo de uso del ExpedienteService
Dim expedienteService As IExpedienteService
Set expedienteService = New CExpedienteService

Dim expediente As T_Expediente
Set expediente = expedienteService.ObtenerExpediente("EXP-2024-INF-001")

If Not expediente Is Nothing Then
    ' Datos obtenidos exitosamente
    Debug.Print "Nemotécnico: " & expediente.Nemotecnico
    Debug.Print "Responsable: " & expediente.ResponsableCalidad
    Debug.Print "Contratista: " & expediente.ContratistaPrincipal
End If
```

### Datos Retornados
```json
{
  "numero_expediente": "EXP-2024-INF-001",
  "nemotecnico": "INF001-24",
  "titulo_expediente": "Suministro e instalación de equipos informáticos para centros educativos",
  "estado_expediente": "En Ejecución",
  "responsable_calidad": "María García López",
  "email_responsable": "maria.garcia@empresa.com",
  "jefe_proyecto": "Carlos Rodríguez Martín",
  "contratista_principal": "Tecnología Avanzada S.L.",
  "valor_contrato": 450000.00,
  "fecha_inicio": "2024-01-15",
  "fecha_fin_prevista": "2024-12-31",
  "ultima_actualizacion": "2024-03-10 16:30:00"
}
```

---

## Resumen de Patrones de Uso

### Frecuencia por Tipo de Solicitud
- **PC (Cambios de Precio)**: 35% - Principalmente por fluctuaciones de mercado
- **CD_CA (Concesiones/Desviaciones)**: 50% - Adaptaciones técnicas y normativas
- **CD_CA_SUB (Sub-Concesiones)**: 15% - Refinamientos de concesiones principales

### Tiempos Promedio de Proceso
- **PC**: 7-10 días laborables
- **CD_CA**: 10-15 días laborables
- **CD_CA_SUB**: 5-7 días laborables

### Actores Más Frecuentes
- **Calidad**: 100% participación (creación y gestión)
- **Ingeniería**: 85% participación (revisión técnica)
- **Administrador**: 10% participación (casos complejos)

---

*Documento generado según la Especificación Funcional CONDOR*  
*Casos basados en situaciones reales del dominio de contratos públicos*  
*Versión: 1.0*  
*Fecha: Diciembre 2024*