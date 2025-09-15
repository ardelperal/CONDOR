# CONDOR - Plan Maestro de Desarrollo

## Contrato JSON de Formularios

### Estructura Principal

```json
{
  "name": "string (requerido)",
  "properties": {
    "caption": "string",
    "width": "number",
    "height": "number",
    "backColor": "string (hex: #RRGGBB)",
    "defaultView": "enum (single|continuous|datasheet|pivotTable|pivotChart)",
    "cycle": "enum (currentRecord|allRecords)",
    "recordSourceType": "enum (table|sql|none)",
    "recordSource": "string",
    "allowEdits": "boolean",
    "allowAdditions": "boolean",
    "allowDeletions": "boolean"
  },
  "sections": {
    "detail": {
      "height": "number",
      "backColor": "string (hex: #RRGGBB)"
    }
  },
  "controls": [
    {
      "name": "string (requerido)",
      "type": "enum (CommandButton|Label|TextBox) (requerido)",
      "properties": {
        "caption": "string",
        "top": "number (requerido)",
        "left": "number (requerido)",
        "width": "number (requerido)",
        "height": "number (requerido)",
        "backColor": "string (hex: #RRGGBB)",
        "foreColor": "string (hex: #RRGGBB)",
        "fontName": "string",
        "fontSize": "number",
        "fontBold": "boolean",
        "fontItalic": "boolean",
        "picture": "string (ruta relativa)",
        "textAlign": "enum (left|center|right)",
        "borderStyle": "enum (transparent|solid|dashes|dots)",
        "specialEffect": "enum (flat|raised|sunken|etched|shadowed|chiseled)"
      }
    }
  ]
}
```

### Tokens Válidos

#### recordSourceType
- **table**: Tabla de base de datos
- **sql**: Consulta SQL personalizada  
- **none**: Sin origen de datos

#### defaultView
- **single**: Vista de formulario único
- **continuous**: Vista continua
- **datasheet**: Vista de hoja de datos
- **pivotTable**: Vista de tabla dinámica
- **pivotChart**: Vista de gráfico dinámico

#### cycle
- **currentRecord**: Solo registro actual
- **allRecords**: Todos los registros

#### textAlign
- **left**: Alineación izquierda
- **center**: Alineación centrada
- **right**: Alineación derecha

#### borderStyle
- **transparent**: Sin borde
- **solid**: Borde sólido
- **dashes**: Borde con guiones
- **dots**: Borde punteado

#### specialEffect
- **flat**: Efecto plano
- **raised**: Efecto elevado
- **sunken**: Efecto hundido
- **etched**: Efecto grabado
- **shadowed**: Efecto con sombra
- **chiseled**: Efecto cincelado

### Comandos Disponibles

#### export-form
Exporta un formulario de Access a formato JSON.

#### import-form
Importa un formulario desde JSON a Access.

#### roundtrip-form
Ejecuta un ciclo completo de import-form --overwrite seguido de export-form para verificar la integridad del proceso.

### Notas Importantes

⚠️ **NOTA**: Los comandos export/import/roundtrip operan en vista Diseño (no ejecutan eventos).

### Colores Válidos

Formato hexadecimal: #RRGGBB (ej: #FF0000 para rojo)

## Estado del Proyecto

- ✅ Implementación de comando roundtrip-form
- ✅ Corrección de documentación recordSourceType
- ✅ Smoke tests completados
- ✅ Documentación actualizada
- ✅ **REFACTORING COMPLETADO (Enero 2025)**:
  - ResolveDbPath() unificado con DefaultFrontendDb/DefaultBackendDb
  - OpenAccessApp/CloseAccessApp con manejo correcto de variables globales
  - RebuildProject robustecido con validaciones VBIDE y backup fallbacks
  - Eliminación de rutas hardcodeadas y duplicados de código