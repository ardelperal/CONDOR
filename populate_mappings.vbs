' ===============================================================================
' Script: populate_mappings.vbs
' Propósito: Poblar la tabla TbMapeo_Campos con los mapeos definidos en la
'           Especificación Funcional del Sistema CONDOR (Sección 9)
' Autor: Sistema CONDOR
' Fecha: 2024
' ===============================================================================

Option Explicit

Dim db, strDBPath

' Configurar la ruta de la base de datos
strDBPath = "C:\Proyectos\CONDOR\back\CONDOR_datos.accdb"

' Verificar que el archivo existe
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
If Not fso.FileExists(strDBPath) Then
    WScript.Echo "ERROR: No se encuentra la base de datos en: " & strDBPath
    WScript.Quit 1
End If

On Error Resume Next

' Conectar a la base de datos usando DAO
Dim dbEngine
Set dbEngine = CreateObject("DAO.DBEngine.120")
Set db = dbEngine.OpenDatabase(strDBPath)

If Err.Number <> 0 Then
    WScript.Echo "ERROR: No se pudo conectar a la base de datos: " & Err.Description
    WScript.Quit 1
End If

WScript.Echo "=== INICIANDO POBLACIÓN DE TABLA TbMapeo_Campos ==="
WScript.Echo "Conectado exitosamente a: " & strDBPath

' Limpiar registros existentes para evitar duplicados
WScript.Echo "Eliminando registros existentes..."
db.Execute "DELETE * FROM TbMapeo_Campos"

If Err.Number <> 0 Then
    WScript.Echo "ERROR al limpiar tabla: " & Err.Description
    db.Close
    WScript.Quit 1
End If

WScript.Echo "Tabla limpiada correctamente."

' ===============================================================================
' PLANTILLA PC (F4203.11 - Propuesta de Cambio)
' ===============================================================================
WScript.Echo "Poblando mapeos para la plantilla PC..."

' Campos de texto simples (NULL)
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('PC', 'RefContratoInspeccionOficial', NULL, 'Parte0_1')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('PC', 'RefSuministrador', NULL, 'Parte0_2')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('PC', 'SuministradorNombreDir', NULL, 'Parte1_1')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('PC', 'ObjetoContrato', NULL, 'Parte1_2')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('PC', 'DescripcionMaterialAfectado', NULL, 'Parte1_3')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('PC', 'NumPlanoEspecificacion', NULL, 'Parte1_4')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('PC', 'DescripcionPropuestaCambio', NULL, 'Parte1_5')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('PC', 'DescripcionPropuestaCambio_Cont', NULL, 'Parte1_5Cont')"

' Campos booleanos de motivos (True)
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('PC', 'Motivo_CorregirDeficiencias', True, 'Parte1_6_1')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('PC', 'Motivo_MejorarCapacidad', True, 'Parte1_6_2')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('PC', 'Motivo_AumentarNacionalizacion', True, 'Parte1_6_3')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('PC', 'Motivo_MejorarSeguridad', True, 'Parte1_6_4')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('PC', 'Motivo_MejorarFiabilidad', True, 'Parte1_6_5')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('PC', 'Motivo_MejorarCosteEficacia', True, 'Parte1_6_6')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('PC', 'Motivo_Otros', True, 'Parte1_6_7')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('PC', 'Motivo_Otros_Detalle', NULL, 'Parte1_6_8')"

' Campos de incidencia con valores específicos
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('PC', 'IncidenciaCoste', 'Aumentará', 'Parte1_7a_1')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('PC', 'IncidenciaCoste', 'Disminuirá', 'Parte1_7a_2')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('PC', 'IncidenciaCoste', 'No variará', 'Parte1_7a_3')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('PC', 'IncidenciaPlazo', 'Aumentará', 'Parte1_7b_1')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('PC', 'IncidenciaPlazo', 'Disminuirá', 'Parte1_7b_2')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('PC', 'IncidenciaPlazo', 'No variará', 'Parte1_7b_3')"

' Campos booleanos de incidencias técnicas (True)
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('PC', 'Incidencia_Seguridad', True, 'Parte1_7c_1')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('PC', 'Incidencia_Fiabilidad', True, 'Parte1_7c_2')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('PC', 'Incidencia_Mantenibilidad', True, 'Parte1_7c_3')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('PC', 'Incidencia_Intercambiabilidad', True, 'Parte1_7c_4')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('PC', 'Incidencia_VidaUtilAlmacen', True, 'Parte1_7c_5')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('PC', 'Incidencia_FuncionamientoFuncion', True, 'Parte1_7c_6')"

' Campos booleanos de afectación de material (True)
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('PC', 'CambioAfecta_MaterialEntregado', True, 'Parte1_9_1')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('PC', 'CambioAfecta_MaterialPorEntregar', True, 'Parte1_9_2')"

' Campos de firmas y datos finales
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('PC', 'FirmaOficinaTecnica_Nombre', NULL, 'Parte1_10')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('PC', 'FirmaRepSuministrador_Nombre', NULL, 'Parte1_11')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('PC', 'ObservacionesRAC_Ref', NULL, 'Parte2_1')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('PC', 'RAC_Codigo', NULL, 'Parte2_2')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('PC', 'ObservacionesRAC', NULL, 'Parte2_3')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('PC', 'FechaFirmaRAC', NULL, 'Parte2_4')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('PC', 'ObsAprobacionAutoridadDiseno', NULL, 'Parte3_1')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('PC', 'FirmaAutoridadDiseno_NombreCargo', NULL, 'Parte3_2')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('PC', 'FechaFirmaAutoridadDiseno', NULL, 'Parte3_3')"

' Decisión final con valores específicos
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('PC', 'DecisionFinal', 'APROBADO', 'Parte3_2_1')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('PC', 'DecisionFinal', 'NO APROBADO', 'Parte3_2_2')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('PC', 'ObsDecisionFinal', NULL, 'Parte3_3_1')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('PC', 'CargoFirmanteFinal', NULL, 'Parte3_3_2')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('PC', 'FechaFirmaDecisionFinal', NULL, 'Parte3_3_3')"

' ===============================================================================
' PLANTILLA CDCA (F4203.10 - Desviación / Concesión)
' ===============================================================================
WScript.Echo "Poblando mapeos para la plantilla CDCA..."

' Campos de texto simples (NULL)
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('CDCA', 'RefSuministrador', NULL, 'Parte0_1')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('CDCA', 'NumContrato', NULL, 'Parte1_2')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('CDCA', 'IdentificacionMaterial', NULL, 'Parte1_3')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('CDCA', 'NumPlanoEspecificacion', NULL, 'Parte1_4')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('CDCA', 'CantidadPeriodo', NULL, 'Parte1_5a')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('CDCA', 'NumSerieLote', NULL, 'Parte1_5b')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('CDCA', 'DescripcionImpactoNC', NULL, 'Parte1_6')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('CDCA', 'RefDesviacionesPrevias', NULL, 'Parte1_7')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('CDCA', 'CausaNC', NULL, 'Parte1_8')"

' Campos de impacto coste con valores específicos
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('CDCA', 'ImpactoCoste', 'Increased / aumentado', 'Parte1_9_1')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('CDCA', 'ImpactoCoste', 'Decreased / disminuido', 'Parte1_9_2')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('CDCA', 'ImpactoCoste', 'Unchanged / sin cambio', 'Parte1_9_3')"

' Clasificación NC con valores específicos
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('CDCA', 'ClasificacionNC', 'Major / Mayor', 'Parte1_10_1')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('CDCA', 'ClasificacionNC', 'Minor / Menor', 'Parte1_10_2')"

' Campos booleanos y de texto
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('CDCA', 'RequiereModificacionContrato', True, 'Parte1_12_1')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('CDCA', 'EfectoFechaEntrega', NULL, 'Parte1_13')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('CDCA', 'IdentificacionAutoridadDiseno', NULL, 'Parte1_14')"

' Campos booleanos de autoridad de diseño
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('CDCA', 'EsSuministradorAD', True, 'Parte1_18_1')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('CDCA', 'EsSuministradorAD', False, 'Parte1_18_2')"

' Campos adicionales
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('CDCA', 'DescripcionImpactoNC_Cont', NULL, 'Parte1_20')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('CDCA', 'RAC_Ref', NULL, 'Parte2_21_1')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('CDCA', 'RAC_Codigo', NULL, 'Parte2_21_2')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('CDCA', 'ObservacionesRAC', NULL, 'Parte2_21_3')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('CDCA', 'FechaFirmaRAC', NULL, 'Parte2_22')"

' Decisión final con valores específicos
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('CDCA', 'DecisionFinal', 'APROBADO', 'Parte3_23_1')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('CDCA', 'DecisionFinal', 'NO APROBADO', 'Parte3_23_2')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('CDCA', 'ObservacionesFinales', NULL, 'Parte3_24_1')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('CDCA', 'FechaFirmaDecisionFinal', NULL, 'Parte3_24_2')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('CDCA', 'CargoFirmanteFinal', NULL, 'Parte3_24_4')"

' ===============================================================================
' PLANTILLA CDCASUB (F4203.101 - Desviación / Concesión Sub-suministrador)
' ===============================================================================
WScript.Echo "Poblando mapeos para la plantilla CDCASUB..."

' Campos de texto simples (NULL)
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('CDCASUB', 'RefSuministrador', NULL, 'Parte0_1')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('CDCASUB', 'RefSubSuministrador', NULL, 'Parte0_2')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('CDCASUB', 'SuministradorPrincipalNombreDir', NULL, 'Parte1_1')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('CDCASUB', 'SubSuministradorNombreDir', NULL, 'Parte1_2')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('CDCASUB', 'IdentificacionMaterial', NULL, 'Parte1_5')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('CDCASUB', 'NumPlanoEspecificacion', NULL, 'Parte1_6')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('CDCASUB', 'CantidadPeriodo', NULL, 'Parte1_7a')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('CDCASUB', 'NumSerieLote', NULL, 'Parte1_7b')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('CDCASUB', 'DescripcionImpactoNC', NULL, 'Parte1_8')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('CDCASUB', 'RefDesviacionesPrevias', NULL, 'Parte1_9')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('CDCASUB', 'CausaNC', NULL, 'Parte1_10')"

' Campos de impacto coste con valores específicos
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('CDCASUB', 'ImpactoCoste', 'Incrementado', 'Parte1_11_1')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('CDCASUB', 'ImpactoCoste', 'Sin cambio', 'Parte1_11_2')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('CDCASUB', 'ImpactoCoste', 'Disminuido', 'Parte1_11_3')"

' Clasificación NC con valores específicos
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('CDCASUB', 'ClasificacionNC', 'Mayor', 'Parte1_12_1')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('CDCASUB', 'ClasificacionNC', 'Menor', 'Parte1_12_2')"

' Campos booleanos de afectación (True)
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('CDCASUB', 'Afecta_Prestaciones', True, 'Parte1_13_1')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('CDCASUB', 'Afecta_Seguridad', True, 'Parte1_13_2')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('CDCASUB', 'Afecta_Fiabilidad', True, 'Parte1_13_3')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('CDCASUB', 'Afecta_VidaUtil', True, 'Parte1_13_4')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('CDCASUB', 'Afecta_Medioambiente', True, 'Parte1_13_5')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('CDCASUB', 'Afecta_Intercambiabilidad', True, 'Parte1_13_6')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('CDCASUB', 'Afecta_Mantenibilidad', True, 'Parte1_13_7')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('CDCASUB', 'Afecta_Apariencia', True, 'Parte1_13_8')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('CDCASUB', 'Afecta_Otros', True, 'Parte1_13_9')"

' Campos adicionales
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('CDCASUB', 'RequiereModificacionContrato', True, 'Parte1_14')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('CDCASUB', 'EfectoFechaEntrega', NULL, 'Parte1_15')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('CDCASUB', 'IdentificacionAutoridadDiseno', NULL, 'Parte1_16')"

' Campos booleanos de autoridad de diseño
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('CDCASUB', 'EsSubSuministradorAD', True, 'Parte1_20_1')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('CDCASUB', 'EsSubSuministradorAD', False, 'Parte1_20_2')"

' Campos finales
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('CDCASUB', 'NombreRepSubSuministrador', NULL, 'Parte1_21')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('CDCASUB', 'DescripcionImpactoNC_Cont', NULL, 'Parte1_22')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('CDCASUB', 'RAC_Ref', NULL, 'Parte2_23_1')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('CDCASUB', 'RAC_Codigo', NULL, 'Parte2_23_2')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('CDCASUB', 'ObservacionesRAC', NULL, 'Parte2_23_3')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('CDCASUB', 'FechaFirmaRAC', NULL, 'Parte2_25')"

' Decisión del suministrador principal con valores específicos
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('CDCASUB', 'DecisionSuministradorPrincipal', 'APROBADO', 'Parte3_26_1')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('CDCASUB', 'DecisionSuministradorPrincipal', 'NO APROBADO', 'Parte3_26_2')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('CDCASUB', 'ObsSuministradorPrincipal', NULL, 'Parte3_27_1')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('CDCASUB', 'FechaFirmaSuministradorPrincipal', NULL, 'Parte3_27_2')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('CDCASUB', 'FirmaSuministradorPrincipal_NombreCargo', NULL, 'Parte3_27_4')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('CDCASUB', 'ObsRACDelegador', NULL, 'Parte4_28')"
db.Execute "INSERT INTO TbMapeo_Campos (NombrePlantilla, NombreCampoTabla, ValorAsociado, NombreCampoWord) VALUES ('CDCASUB', 'FechaFirmaRACDelegador', NULL, 'Parte4_30')"

' ===============================================================================
' FINALIZACIÓN
' ===============================================================================

If Err.Number <> 0 Then
    WScript.Echo "ERROR durante la inserción: " & Err.Description
    db.Close
    WScript.Quit 1
End If

' Contar registros insertados
Dim rs, totalRegistros
Set rs = db.OpenRecordset("SELECT COUNT(*) AS Total FROM TbMapeo_Campos")
totalRegistros = rs.Fields("Total").Value
rs.Close

WScript.Echo "=== PROCESO COMPLETADO EXITOSAMENTE ==="
WScript.Echo "Tabla TbMapeo_Campos poblada con " & totalRegistros & " registros."
WScript.Echo "Plantillas procesadas:"
WScript.Echo "  - PC: Propuesta de Cambio (F4203.11)"
WScript.Echo "  - CDCA: Desviación/Concesión (F4203.10)"
WScript.Echo "  - CDCASUB: Desviación/Concesión Sub-suministrador (F4203.101)"

' Cerrar la base de datos
db.Close
Set db = Nothing
Set dbEngine = Nothing

WScript.Echo "Base de datos cerrada correctamente."
WScript.Echo "Script populate_mappings.vbs finalizado."
