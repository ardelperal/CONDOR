Attribute VB_Name = "Test_ExpedienteService"
Option Compare Database
Option Explicit

#If DEV_MODE Then

' ============================================================================
' Modulo: Test_ExpedienteService
' Descripcion: Pruebas unitarias para el servicio de expedientes
' Autor: CONDOR-Expert
' Fecha: Diciembre 2024
' ============================================================================

' Ejecuta todas las pruebas del ExpedienteService
Public Sub Test_ExpedienteService_All()
    Debug.Print "=== INICIANDO PRUEBAS DE EXPEDIENTE SERVICE ==="
    
    Test_GetExpedientePorID_ExpedienteExistente
    Test_GetExpedientePorID_ExpedienteSecundario
    Test_GetExpedientePorID_ExpedienteConCamposVacios
    Test_GetExpedientePorID_ExpedienteNoExistente
    
    Debug.Print "=== PRUEBAS DE EXPEDIENTE SERVICE COMPLETADAS ==="
End Sub

' Prueba obtener expediente existente (ID 1001)
Public Sub Test_GetExpedientePorID_ExpedienteExistente()
    Debug.Print "Ejecutando: Test_GetExpedientePorID_ExpedienteExistente"
    
    Dim svc As IExpedienteService
    Set svc = New CMockExpedienteService
    
    Dim resultado As T_Expediente
    resultado = svc.GetExpedientePorID(1001)
    
    ' Verificar que los datos coinciden con los hardcodeados en el Mock
    Debug.Assert resultado.IDExpediente = 1001
    Debug.Assert resultado.Nemotecnico = "TEST-001"
    Debug.Assert resultado.Titulo = "Expediente de Prueba Principal"
    Debug.Assert resultado.ResponsableCalidad = "Juan Perez"
    Debug.Assert resultado.ResponsableTecnico = "Maria Garcia"
    Debug.Assert resultado.Pecal = "SI"
    
    Debug.Print "Test_GetExpedientePorID_ExpedienteExistente PASADO"
End Sub

' Prueba obtener expediente secundario (ID 1002)
Public Sub Test_GetExpedientePorID_ExpedienteSecundario()
    Debug.Print "Ejecutando: Test_GetExpedientePorID_ExpedienteSecundario"
    
    Dim svc As IExpedienteService
    Set svc = New CMockExpedienteService
    
    Dim resultado As T_Expediente
    resultado = svc.GetExpedientePorID(1002)
    
    ' Verificar que los datos coinciden con los hardcodeados en el Mock
    Debug.Assert resultado.IDExpediente = 1002
    Debug.Assert resultado.Nemotecnico = "TEST-002"
    Debug.Assert resultado.Titulo = "Expediente de Prueba Secundario"
    Debug.Assert resultado.ResponsableCalidad = "Ana Lopez"
    Debug.Assert resultado.ResponsableTecnico = "Carlos Ruiz"
    Debug.Assert resultado.Pecal = "NO"
    
    Debug.Print "Test_GetExpedientePorID_ExpedienteSecundario PASADO"
End Sub

' Prueba obtener expediente con campos vacios (ID 1003)
Public Sub Test_GetExpedientePorID_ExpedienteConCamposVacios()
    Debug.Print "Ejecutando: Test_GetExpedientePorID_ExpedienteConCamposVacios"
    
    Dim svc As IExpedienteService
    Set svc = New CMockExpedienteService
    
    Dim resultado As T_Expediente
    resultado = svc.GetExpedientePorID(1003)
    
    ' Verificar que los datos coinciden con los hardcodeados en el Mock
    Debug.Assert resultado.IDExpediente = 1003
    Debug.Assert resultado.Nemotecnico = "TEST-003"
    Debug.Assert resultado.Titulo = "Expediente con Campos Vacios"
    Debug.Assert resultado.ResponsableCalidad = ""
    Debug.Assert resultado.ResponsableTecnico = ""
    Debug.Assert resultado.Pecal = ""
    
    Debug.Print "Test_GetExpedientePorID_ExpedienteConCamposVacios PASADO"
End Sub

' Prueba obtener expediente no existente (ID 9999)
Public Sub Test_GetExpedientePorID_ExpedienteNoExistente()
    Debug.Print "Ejecutando: Test_GetExpedientePorID_ExpedienteNoExistente"
    
    Dim svc As IExpedienteService
    Set svc = New CMockExpedienteService
    
    Dim resultado As T_Expediente
    resultado = svc.GetExpedientePorID(9999)
    
    ' Verificar que devuelve una estructura vacia para expediente no encontrado
    Debug.Assert resultado.IDExpediente = 0
    Debug.Assert resultado.Nemotecnico = ""
    Debug.Assert resultado.Titulo = ""
    Debug.Assert resultado.ResponsableCalidad = ""
    Debug.Assert resultado.ResponsableTecnico = ""
    Debug.Assert resultado.Pecal = ""
    
    Debug.Print "Test_GetExpedientePorID_ExpedienteNoExistente PASADO"
End Sub

#End If