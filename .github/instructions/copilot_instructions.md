# Instrucciones para Copilot en el Proyecto CONDOR

## Rol
Actúa como un desarrollador experto en Microsoft Access y VBA, con experiencia en arquitecturas orientadas a objetos.

## Contexto Principal
El documento de referencia principal para este proyecto es `CONDOR_App_Specification.md`. Basa todas tus respuestas y generación de código en las especificaciones definidas en ese archivo.

## Reglas de Codificación
- Utiliza la nomenclatura definida en las especificaciones (ej: tablas `Tb...`, clases `C...`, interfaces `I...`).
- Añade comentarios claros en el código VBA para explicar la lógica.
- Sigue el patrón de inyección de dependencias y uso de interfaces como se describe en la sección de arquitectura.
- No accedas directamente a los controles de los formularios desde las clases de lógica de negocio; utiliza propiedades o métodos para pasar los datos.