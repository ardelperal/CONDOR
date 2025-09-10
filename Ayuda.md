

```

## Principio Arquitectónico: UI como Código

El proyecto CONDOR implementa el principio arquitectónico **"UI como Código"** para la gestión de formularios de Microsoft Access. Este principio establece que los formularios de Access deben ser tratados como código fuente, permitiendo su versionado, revisión y gestión a través de herramientas de control de versiones.

### Definición del Principio

Los formularios de Access se serializan como archivos JSON estructurados que contienen toda la información necesaria para recrear el formulario: propiedades, controles, posicionamiento, formato y configuración. Estos archivos JSON se consideran la **fuente de verdad** para los formularios de la aplicación.

### Estructura de Directorios Canónica

```
ui/
├── definitions/     ← Definiciones JSON de formularios (fuente de verdad)
│   ├── frmPrincipal.json
│   ├── frmSolicitudes.json
│   └── TestForm.json
├── assets/         ← Recursos gráficos (iconos, imágenes)
│   ├── Guardar1_25x25.png
│   ├── Cerrar1_25x25.png
│   └── ...
└── templates/      ← Plantillas base para nuevos formularios
```

### Herramientas CLI

El sistema proporciona dos comandos principales para implementar este principio:

#### Exportación de Formularios (`export-form`)

```bash
# Exportar formulario a JSON
cscript condor_cli.vbs export-form <db_path> <form_name> [--output] [--password]

# Ejemplos
cscript condor_cli.vbs export-form ./back/CONDOR.accdb frmPrincipal
cscript condor_cli.vbs export-form ./back/CONDOR.accdb frmPrincipal --output ./ui/definitions/
```

**Funcionalidades:**
- Extrae el diseño completo del formulario incluyendo propiedades, secciones y controles
- Genera archivos JSON legibles y versionables
- Captura todos los tipos de controles (TextBox, Label, CommandButton, etc.)
- Incluye propiedades detalladas: posición, tamaño, formato, fuentes
- Soporte para bases de datos protegidas con contraseña

#### Importación de Formularios (`import-form`)

```bash
# Crear/Modificar formulario desde JSON
cscript condor_cli.vbs import-form <json_path> <db_path> [--password]

# Ejemplos
cscript condor_cli.vbs import-form ./ui/definitions/frmPrincipal.json ./back/CONDOR.accdb
```

**Funcionalidades:**
- Crea formularios nuevos o reemplaza existentes basándose en la definición JSON
- Genera automáticamente todos los controles especificados
- Configura automáticamente posición, tamaño, formato y propiedades
- Mapeo automático de tipos de controles del JSON a objetos Access nativos
- Reemplazo seguro con eliminación previa de formularios existentes
- Validación de estructura del JSON antes de proceder

### Flujo de Trabajo Obligatorio

Para cualquier modificación de formularios en CONDOR, se debe seguir este flujo:

1. **Exportar**: Usar `export-form` para extraer el formulario actual a JSON
2. **Modificar**: Editar el archivo JSON con los cambios requeridos
3. **Versionar**: Confirmar los cambios en el control de versiones (Git)
4. **Importar**: Usar `import-form` para aplicar los cambios al formulario de Access
5. **Validar**: Verificar que el formulario funciona correctamente

### Ventajas del Principio

- **Versionado**: Los formularios pueden ser versionados como cualquier código fuente
- **Revisión de Código**: Los cambios en formularios pueden ser revisados mediante pull requests
- **Trazabilidad**: Historial completo de cambios en la interfaz de usuario
- **Colaboración**: Múltiples desarrolladores pueden trabajar en formularios sin conflictos
- **Automatización**: Posibilidad de generar formularios programáticamente
- **Backup y Restauración**: Los formularios están respaldados en el repositorio
- **Consistencia**: Garantiza que todos los entornos tengan la misma versión de formularios

### Consideraciones Técnicas

- Los archivos JSON deben mantener la estructura definida por el sistema de exportación
- Las rutas de imágenes en `assets/` deben ser relativas al directorio `ui/`
- Se recomienda usar imágenes PNG para compatibilidad con Access
- Los nombres de controles deben seguir las convenciones de nomenclatura de VBA
- Las propiedades de formularios deben ser válidas según la versión de Access utilizada

### Integración con el Ciclo de Desarrollo

Este principio se integra con el **Ciclo de Trabajo de Desarrollo** definido en el proyecto:

1. Los cambios de UI se realizan mediante modificación de archivos JSON
2. Los archivos JSON se incluyen en el proceso de revisión de código
3. La importación de formularios forma parte del proceso de despliegue
4. Las pruebas de UI se ejecutan contra los formularios importados desde JSON

**Nota**: Este principio es fundamental para mantener la coherencia y trazabilidad de la interfaz de usuario en el proyecto CONDOR, y su cumplimiento es obligatorio para todas las modificaciones de formularios.

Método Application.CreateControl (Access)
07/04/2023
El método CreateControl crea un control en un formulario abierto especificado. Por ejemplo, suponga que va a crear un asistente personalizado que permite a los usuarios generar de manera sencilla un determinado formulario. Use el método CreateControl en el asistente para agregar los controles adecuados al formulario.

Sintaxis
expresión. CreateControl (FormName, ControlType, Section, Parent, ColumnName, Left, Top, Width, Height)

expresión Variable que representa un objeto Application.

Parámetros
Nombre	Obligatorio/opcional	Tipo de datos	Descripción
FormName	Necesario	String	Nombre del formulario o informe en el que desea crear el control.
ControlType	Obligatorio	AcControlType	Constante AcControlType que representa el tipo de control que desea crear.
Section	Opcional	AcSection	Constante AcSection que identifica la sección que contendrá el nuevo control.
Parent	Opcional	Variant	Nombre del control principal de un control adjunto. En el caso de los controles que no tienen ningún control primario, use una cadena de longitud cero para este argumento o oódelo.
ColumnName	Opcional	Variant	Nombre del campo al que se enlazará el control si va a ser un control enlazado a datos.
Left,Top	Opcional	Variant	Coordenadas de la esquina superior izquierda del control en twips.
Width, Height	Opcional	Variant	Expresiones numéricas que indican el ancho y el alto del control en twips.
Valor devuelto
Control

Comentarios
Use los métodos CreateControl y CreateReportControl en un asistente personalizado para crear controles en un formulario o informe. Ambos métodos devuelven un objeto Control .

Use los métodos CreateControl y CreateReportControl solo en la vista Diseño del formulario o en la vista Diseño del informe, respectivamente.

Use el argumento Parent para identificar la relación entre un control principal y un control subordinado. Por ejemplo, si un cuadro de texto tiene una etiqueta adjunta, el cuadro de texto es el control principal y la etiqueta es el control subordinado. Al crear el control de etiqueta, establezca su argumento Parent en una cadena que identifique el nombre del control primario. Al crear el cuadro de texto, establezca su argumento Parent en una cadena de longitud cero.

También se establece el argumento Primario al crear casillas, botones de opción o botones de alternancia. Un grupo de opciones es el control principal de las casillas, botones de opción o botones de alternancia que contiene. Los únicos controles que pueden tener un control principal son una etiqueta, casilla, botón de opción o botón de alternancia. Todos estos controles pueden también crearse independientemente, sin un control principal.

Establezca el argumento ColumnName según el tipo de control que esté creando y si se enlazará o no a un campo de una tabla. Los controles que pueden depender de un campo incluyen el cuadro de texto, cuadro de lista, cuadro combinado, grupo de opciones y marco de objeto dependiente. Igualmente, los controles de botón de alternancia, de botón de opción y de casilla pueden depender de un campo si no están contenidos en un grupo de opciones.

Si especifica el nombre de un campo para el argumento ColumnName , cree un control que esté enlazado a ese campo. Todas las propiedades del control se establecen entonces automáticamente en los valores de las propiedades de campo correspondientes. Por ejemplo, el valor de la propiedad ValidationRule del control será el mismo que el valor de esa propiedad para el campo.

 Nota

Si el asistente crea controles en un formulario o informe nuevo o existente, debe abrir primero el formulario o informe en la vista Diseño.

Para quitar un control de un formulario o informe, use los métodos DeleteControl y DeleteReportControl .

Ejemplo:
En el ejemplo siguiente se crea primero un nuevo formulario basado en una tabla Pedidos. Después utiliza el método CreateControl para crear un control de cuadro de texto y un control de etiqueta adjunta en el formulario.

VB

Copiar
Sub NewControls() 
 Dim frm As Form 
 Dim ctlLabel As Control, ctlText As Control 
 Dim intDataX As Integer, intDataY As Integer 
 Dim intLabelX As Integer, intLabelY As Integer 
 
 ' Create new form with Orders table as its record source. 
 Set frm = CreateForm 
 frm.RecordSource = "Orders" 
 ' Set positioning values for new controls. 
 intLabelX = 100 
 intLabelY = 100 
 intDataX = 1000 
 intDataY = 100 
 ' Create unbound default-size text box in detail section. 
 Set ctlText = CreateControl(frm.Name, acTextBox, , "", "", _ 
 intDataX, intDataY) 
 ' Create child label control for text box. 
 Set ctlLabel = CreateControl(frm.Name, acLabel, , _ 
 ctlText.Name, "NewLabel", intLabelX, intLabelY) 
 ' Restore form. 
 DoCmd.Restore 
End Sub


Método Application.CreateForm (Access)
07/04/2023
El método CreateForm crea un formulario y devuelve un objeto de formulario.

Sintaxis
expresión. CreateForm (Database, FormTemplate)

expresión Variable que representa un objeto Application.

Parámetros
Nombre	Obligatorio/opcional	Tipo de datos	Descripción
Base de datos	Opcional	Variant	Nombre de la base de datos que contiene la plantilla de formulario que desea usar para crear un formulario. Si desea la base de datos activa, omita este argumento. Si desea utilizar una base de datos de biblioteca abierta, especifique la biblioteca de base de datos mediante este argumento.
FormTemplate	Opcional	Variant	Nombre del formulario que desea usar como plantilla para crear un formulario nuevo.
Valor devuelto
Formulario

Comentarios
Use el método CreateForm al diseñar un asistente que cree un nuevo formulario.

El método CreateForm abre un nuevo formulario minimizado en la vista Diseño del formulario.

Si el nombre que usa para el argumento FormTemplate no es válido, Visual Basic usa la plantilla de formulario especificada por la configuración Plantilla de formulario en la pestaña Formularios o informes del cuadro de diálogo Opciones .

Ejemplo:
En este ejemplo se crea un nuevo formulario en la base de datos de ejemplo Northwind basado en el formulario clientes y establece su propiedad RecordSource en la tabla Customers. Ejecute este código desde la base de datos de ejemplo Northwind.

VB

Copiar
Sub NewForm() 
 Dim frm As Form 
 
 ' Create form based on Customers form. 
 Set frm = CreateForm("Customers") 
 DoCmd.Restore 
 ' Set RecordSource property to Customers table. 
 frm.RecordSource = "Customers" 
End Sub

Método Application.LoadPicture (Access)
07/04/2023
El método LoadPicture carga un gráfico en un control ActiveX.

Sintaxis
expresión. LoadPicture (FileName)

expresión Variable que representa un objeto Application.

Parámetros
Nombre	Obligatorio/opcional	Tipo de datos	Descripción
FileName	Necesario	String	Nombre de archivo del gráfico que se va a cargar. El gráfico puede ser un archivo de mapa de bits (.bmp), un archivo de icono (.ico), un archivo codificado de longitud de ejecución (.rle) o un metarchivo (.wmf).
Valor devuelto
Objeto

Comentarios
Asigne el valor devuelto del método LoadPicture a la propiedad Picture de un control ActiveX para cargar dinámicamente un gráfico en el control. En el siguiente ejemplo se carga un mapa de bits en un control denominado OLECustomControl en un formulario Pedidos:

VB

Copiar
Set Forms!Orders!OLECustomControl.Picture = _ 
 LoadPicture("Stars.bmp")
 Nota

No se puede utilizar el método LoadPicture para establecer la propiedad Picture de un control de imagen. Este método sólo funciona con controles de ActiveX. Para establecer la propiedad Picture de un control de imagen, asigne a la propiedad una cadena especificando el nombre de archivo y la ruta de acceso del gráfico que desee.