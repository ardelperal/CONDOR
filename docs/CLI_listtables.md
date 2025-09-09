# Comando listtables - Documentación

## Descripción

El comando `listtables` permite listar las tablas de una base de datos Access y mostrar información detallada sobre su estructura.

## Sintaxis

```bash
cscript condor_cli.vbs listtables [db_path] [--schema] [--output]
```

## Parámetros

- `db_path`: Ruta a la base de datos Access (.accdb). Si no se especifica, usa la base de datos por defecto.
- `--schema`: Muestra información detallada del esquema incluyendo campos, tipos y si son requeridos.
- `--output`: Exporta la salida al archivo `[nombre_bd]_listtables.txt` (donde nombre_bd es el nombre de la base de datos).

## Ejemplos de Uso

### Listar solo nombres de tablas
```bash
cscript condor_cli.vbs listtables ./back/CONDOR_datos.accdb
```

### Mostrar esquema completo con campos requeridos
```bash
cscript condor_cli.vbs listtables ./back/CONDOR_datos.accdb --schema
```

### Exportar esquema a archivo
```bash
cscript condor_cli.vbs listtables ./back/CONDOR_datos.accdb --schema --output
# Genera: CONDOR_datos_listtables.txt
```

## Formato de Salida con --schema

Cuando se usa la opción `--schema`, la salida incluye las siguientes columnas:

- **Campo**: Nombre del campo en la tabla
- **Tipo**: Tipo de datos del campo (Text, Long, DateTime, Boolean, etc.)
- **PK**: Indica si el campo es clave primaria ("PK" o vacío)
- **Requerido**: Indica si el campo es obligatorio ("true" o "false")

### Ejemplo de Salida

```
=== LISTADO DE TABLAS ===
Modo: Esquema Detallado

------------------------------------------------------------
1. tbSolicitudes (0 registros)
------------------------------------------------------------
Campo                    Tipo           PK      Requerido
------------------------------------------------------------
idSolicitud              Long           PK      true
idExpediente             Long                   true
tipoSolicitud            Text                   false
subTipoSolicitud         Text                   false
codigoSolicitud          Text                   false
idEstadoInterno          Long                   true
fechaCreacion            DateTime               true
usuarioCreacion          Text                   true
fechaPaseTecnico         DateTime               false
fechaModificacion        DateTime               false
usuarioModificacion      Text                   false
```

## Notas Técnicas

- El comando utiliza la función `PadRight` para alinear correctamente las columnas en la salida.
- Los campos requeridos se determinan mediante la propiedad `Required` de los objetos Field de Access.
- Las tablas del sistema (MSys*) se excluyen automáticamente del listado.
- La salida se formatea tanto para consola como para archivo de texto.

## Casos de Uso

1. **Documentación de Base de Datos**: Generar documentación actualizada del esquema.
2. **Análisis de Estructura**: Revisar la estructura de tablas antes de realizar migraciones.
3. **Validación de Campos**: Verificar qué campos son obligatorios en cada tabla.
4. **Desarrollo**: Consultar rápidamente la estructura de datos durante el desarrollo.

## Archivos Relacionados

- `condor_cli.vbs`: Implementación del comando
- `[nombre_bd]_listtables.txt`: Archivo de salida cuando se usa `--output` (ej: `Lanzadera_datos_listtables.txt`)
- `back/CONDOR_datos.accdb`: Base de datos principal del proyecto

## Historial de Cambios

### Versión Actual
- ✅ Agregada columna "Requerido" que muestra si los campos son obligatorios
- ✅ Valores "true"/"false" para compatibilidad con procesamiento programático
- ✅ Alineación mejorada de columnas con función `PadRight`
- ✅ Soporte para exportación a archivo de texto

### Versiones Anteriores
- Mostraba solo Campo, Tipo y PK
- Usaba tabulaciones que causaban desalineamiento
- Valores "Sí"/"No" para campos requeridos