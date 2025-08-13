# CONDOR

## Resumen
CONODOR es una aplicación para gestionar el ciclo de vida de solicitudes de cambio, desviación o concesión en expedientes de contratos públicos. Está desarrollada en Microsoft Access con VBA y orientada a usuarios de Calidad y Técnico.

## Arquitectura
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

---
Este README es un resumen inicial. Se irá completando según se implementen las funcionalidades.
