# Arquitectura del Proyecto CONDOR

## Estructura de Bases de Datos

### Frontend (Desarrollo)
- **Archivo**: `c:\Proyectos\CONDOR\back\Desarrollo\CONDOR.accdb`
- **Propósito**: Base de datos de desarrollo que contiene:
  - Módulos VBA (.bas)
  - Clases VBA (.cls)
  - Formularios y otros objetos de Access
- **Operaciones**: Import/Export de código VBA

### Backend (Datos)
- **Archivo**: `c:\Proyectos\CONDOR\back\CONDOR_datos.accdb`
- **Propósito**: Base de datos de datos que contiene:
  - Tablas de datos del sistema
  - Estructura de datos del proyecto
- **Operaciones**: Creación, modificación y eliminación de tablas

## Herramienta condor_cli.vbs

El script `condor_cli.vbs` (CONDOR Command Line Interface) está configurado para:
- **Acciones de código** (`import`, `export`): Usar CONDOR.accdb (frontend)
- **Acciones de tablas** (`createtable`, `droptable`, `listtables`): Usar CONDOR_datos.accdb (backend)
- **Funcionalidades futuras**: Tests, validaciones y otras operaciones del proyecto

## Reglas de Trabajo

1. **Para operaciones de tablas**: Siempre trabajar con `CONDOR_datos.accdb`
2. **Para operaciones de código VBA**: Siempre trabajar con `CONDOR.accdb`
3. **Separación clara**: Frontend para lógica, Backend para datos

---
*Documentación actualizada para mantener la coherencia arquitectónica del proyecto CONDOR*