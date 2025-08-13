# Requerimientos de la Herramienta CONDOR (Borrador 1.0)

## 1. Objetivo Principal
Reducir el tiempo operativo en la creación, seguimiento, cierre y búsqueda de Cambios, Concesiones y Desviaciones, automatizando el proceso de gestión interna y la generación de documentación oficial.

## 2. Roles de Usuario
La aplicación debe tener al menos dos roles con permisos diferenciados:
*   **Calidad:**
    *   Inicia los registros (pre-registro).
    *   Revisa y aprueba/deniega internamente las solicitudes completadas por Ingeniería.
    *   Gestiona los estados de la solicitud.
    *   Exporta los formularios a formato oficial Word.
    *   Envía la documentación y registra las firmas externas (RAC, Cliente, etc.).
*   **Ingeniería:**
    *   Recibe notificaciones para completar los datos técnicos de una solicitud.
    *   Cumplimenta los campos técnicos obligatorios del formulario.
    *   Participa en la revisión de solicitudes denegadas.

## 3. Flujo de Trabajo y Gestión de Estados
La aplicación debe gestionar el ciclo de vida de una solicitud a través de los siguientes estados, notificando a los roles correspondientes en cada cambio:
1.  **Registrado:** `Calidad` crea la solicitud con datos básicos.
2.  **En Desarrollo:** `Ingeniería` es notificada para que complete los datos técnicos.
3.  **Modificación:** `Ingeniería` ha terminado y `Calidad` es notificada para que revise.
4.  **Revisión:** `Calidad` está revisando. Puede aprobar o denegar.
5.  **Validación:** La solicitud ha sido aprobada internamente por `Calidad`.
6.  **Formalización:** Se gestiona el envío y la recepción de firmas externas.
7.  **Aprobada:** El proceso ha finalizado con todas las firmas.
8.  **Rechazada / Cancelada:** (Estado a considerar) para solicitudes que no prosperan.

## 4. Gestión de Datos
La aplicación debe contar con un formulario único que agrupe todos los campos necesarios, basándose en la lista detallada en el punto `4.2` del documento `IN250000APPQ02...`. Algunos campos clave son:
*   Datos del Expediente (N.º, Objeto, etc.)
*   Datos de la Solicitud (Tipo, Asunto, Estado, Fechas, etc.)
*   Detalles Técnicos (Causa, Acciones, Especificaciones, etc.)
*   Trazabilidad (Registros vinculados, Documentos referenciados, etc.)
*   Ficheros Adjuntos.

## 5. Generación y Gestión de Documentos

La interacción con los ficheros `.docx` es una funcionalidad central. La aplicación utilizará librerías específicas para manipular estos archivos.

*   **5.1. Generación de Documentos (Escribir):**
    *   El sistema debe ser capaz de tomar los datos de un registro de la aplicación y rellenar una plantilla `.docx` predefinida (`/docs/Plantillas`).
    *   La plantilla contendrá marcadores de posición (ej. `{{NUMERO_EXPEDIENTE}}`) que la aplicación reemplazará con los datos reales del registro.
    *   El resultado se guardará como un nuevo archivo `.docx` en el sistema de ficheros del servidor.

*   **5.2. Almacenamiento y Vinculación (Guardar):**
    *   La aplicación no almacenará el fichero `.docx` directamente en la base de datos. 
    *   En su lugar, guardará la ruta al fichero generado en un campo asociado al registro correspondiente en la base de datos, creando un vínculo permanente.
    *   Esta misma lógica se aplica a cualquier otro fichero que se suba manualmente al registro (adjuntos).

*   **5.3. Importación de Datos (Leer):**
    *   **(Funcionalidad Avanzada)** Se contempla la capacidad de importar un fichero `.docx` para actualizar un registro existente.
    *   El sistema leería el documento, extraería la información relevante y la usaría para rellenar o modificar los campos correspondientes en la base de datos de la aplicación.

## 6. Búsqueda e Informes
*   **Búsqueda Avanzada:** Un sistema de búsqueda y filtrado potente para encontrar rápidamente cualquier solicitud por cualquiera de sus campos (N.º de expediente, estado, fechas, suministrador, etc.).
*   **Exportación a Excel:** Permitir la exportación de los datos de las solicitudes a formato Excel para análisis y cálculo de indicadores por parte del departamento de Calidad.
