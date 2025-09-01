VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CExpedienteRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Implements IExpedienteRepository

' Dependencias inyectadas
Private m_Config As IConfig
Private m_ErrorHandler As IErrorHandlerService

' ============================================================================
' MÉTODOS DE INICIALIZACIÓN
' ============================================================================

' Inicializa el repositorio con sus dependencias
Public Sub Initialize(ByVal config As IConfig, ByVal errorHandler As IErrorHandlerService)
    Set m_Config = config
    Set m_ErrorHandler = errorHandler
End Sub

' ============================================================================
' IMPLEMENTACIÓN DE LA INTERFAZ IExpedienteRepository
' ============================================================================

' Implementación de IExpedienteRepository.ObtenerExpedientePorId
Private Function IExpedienteRepository_ObtenerExpedientePorId(ByVal idExpediente As Long) As EExpediente
    On Error GoTo ErrorHandler
    
    ' Variables para la consulta parametrizada
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim rs As DAO.recordset
    Dim expediente As EExpediente
    
    ' Logging removido - responsabilidad de la capa de servicio
    
    ' Obtener conexión al Backend usando configuración
    Set db = DBEngine.OpenDatabase(m_Config.GetDataPath(), dbFailOnError, False, ";PWD=" & m_Config.GetDatabasePassword())
    
    ' Crear QueryDef con consulta desde modQueries
    Set qdf = db.CreateQueryDef("", GET_EXPEDIENTE_BY_ID)
    
    ' Asignar parámetro directo
    qdf.parameters("pIdExpediente") = idExpediente
    
    ' Ejecutar consulta parametrizada
    Set rs = qdf.OpenRecordset()
    
    ' Mapear recordset a objeto de dominio
    If Not rs.EOF Then
        Set expediente = New EExpediente
        expediente.idExpediente = rs!ID
        expediente.Nemotecnico = rs!Nemotecnico & ""
        expediente.Titulo = rs!Titulo & ""
        expediente.ResponsableCalidad = rs!ResponsableCalidad & ""
        expediente.ResponsableTecnico = rs!ResponsableTecnico & ""
        expediente.Pecal = rs!Pecal & ""
        expediente.NumeroExpediente = rs!NumeroExpediente & ""
        expediente.EstadoExpediente = rs!EstadoExpediente & ""
        expediente.EmailResponsable = rs!EmailResponsable & ""
        expediente.JefeProyecto = rs!JefeProyecto & ""
        expediente.ContratistaPrincipal = rs!ContratistaPrincipal & ""
        If Not IsNull(rs!FechaInicio) Then expediente.FechaInicio = rs!FechaInicio
        If Not IsNull(rs!FechaFinPrevista) Then expediente.FechaFinPrevista = rs!FechaFinPrevista
        If Not IsNull(rs!FechaFinContrato) Then expediente.FechaFinContrato = rs!FechaFinContrato
        If Not IsNull(rs!FechaFinGarantia) Then expediente.FechaFinGarantia = rs!FechaFinGarantia
    End If
    
    ' Limpiar recursos
    rs.Close
    Set rs = Nothing
    Set qdf = Nothing
    db.Close
    Set db = Nothing
    
    ' Devolver el objeto mapeado
    Set IExpedienteRepository_ObtenerExpedientePorId = expediente
    
    Exit Function
    
ErrorHandler:
    ' Limpiar recursos en caso de error
    If Not rs Is Nothing Then rs.Close
    If Not qdf Is Nothing Then Set qdf = Nothing
    If Not db Is Nothing Then db.Close: Set db = Nothing
    
    ' Log del error
    If Not m_ErrorHandler Is Nothing Then
        m_ErrorHandler.LogError Err.Number, Err.Description, "CExpedienteRepository.ObtenerExpedientePorId"
    End If
    
    ' Devolver Nothing en caso de error
    Set IExpedienteRepository_ObtenerExpedientePorId = Nothing
End Function

' Implementación de IExpedienteRepository.ObtenerExpedientePorNemotecnico
Private Function IExpedienteRepository_ObtenerExpedientePorNemotecnico(ByVal Nemotecnico As String) As EExpediente
    On Error GoTo ErrorHandler
    
    ' Variables para la consulta parametrizada
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim rs As DAO.recordset
    Dim expediente As EExpediente
    
    ' Logging removido - responsabilidad de la capa de servicio
    
    ' Obtener conexión al Backend usando configuración
    Set db = DBEngine.OpenDatabase(m_Config.GetDataPath(), dbFailOnError, False, ";PWD=" & m_Config.GetDatabasePassword())
    
    ' Crear QueryDef con consulta desde modQueries
    Set qdf = db.CreateQueryDef("", GET_EXPEDIENTE_BY_NEMOTECNICO)
    
    ' Asignar parámetro directo
    qdf.parameters("pNemotecnico") = Nemotecnico
    
    ' Ejecutar consulta parametrizada
    Set rs = qdf.OpenRecordset()
    
    ' Mapear recordset a objeto de dominio
    If Not rs.EOF Then
        Set expediente = New EExpediente
        expediente.idExpediente = rs!ID
        expediente.Nemotecnico = rs!Nemotecnico & ""
        expediente.Titulo = rs!Titulo & ""
        expediente.ResponsableCalidad = rs!ResponsableCalidad & ""
        expediente.ResponsableTecnico = rs!ResponsableTecnico & ""
        expediente.Pecal = rs!Pecal & ""
        expediente.NumeroExpediente = rs!NumeroExpediente & ""
        expediente.EstadoExpediente = rs!EstadoExpediente & ""
        expediente.EmailResponsable = rs!EmailResponsable & ""
        expediente.JefeProyecto = rs!JefeProyecto & ""
        expediente.ContratistaPrincipal = rs!ContratistaPrincipal & ""
        If Not IsNull(rs!FechaInicio) Then expediente.FechaInicio = rs!FechaInicio
        If Not IsNull(rs!FechaFinPrevista) Then expediente.FechaFinPrevista = rs!FechaFinPrevista
        If Not IsNull(rs!FechaFinContrato) Then expediente.FechaFinContrato = rs!FechaFinContrato
        If Not IsNull(rs!FechaFinGarantia) Then expediente.FechaFinGarantia = rs!FechaFinGarantia
    End If
    
    ' Limpiar recursos
    rs.Close
    Set rs = Nothing
    Set qdf = Nothing
    db.Close
    Set db = Nothing
    
    ' Devolver el objeto mapeado
    Set IExpedienteRepository_ObtenerExpedientePorNemotecnico = expediente
    
    Exit Function
    
ErrorHandler:
    ' Limpiar recursos en caso de error
    If Not rs Is Nothing Then rs.Close
    If Not qdf Is Nothing Then Set qdf = Nothing
    If Not db Is Nothing Then db.Close: Set db = Nothing
    
    ' Log del error
    If Not m_ErrorHandler Is Nothing Then
        m_ErrorHandler.LogError Err.Number, Err.Description, "CExpedienteRepository.ObtenerExpedientePorNemotecnico"
    End If
    
    ' Devolver Nothing en caso de error
    Set IExpedienteRepository_ObtenerExpedientePorNemotecnico = Nothing
End Function

' Implementación de IExpedienteRepository.ObtenerExpedientesActivosParaSelector
Private Function IExpedienteRepository_ObtenerExpedientesActivosParaSelector() As Scripting.Dictionary
    On Error GoTo ErrorHandler
    
    ' Variables para la consulta
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim rs As DAO.recordset
    Dim expedientes As Scripting.Dictionary
    Dim expediente As EExpediente
    
    ' Logging removido - responsabilidad de la capa de servicio
    
    ' Inicializar diccionario
    Set expedientes = New Scripting.Dictionary
    expedientes.CompareMode = TextCompare ' Opcional pero recomendado
    
    ' Obtener conexión al Backend usando configuración
    Set db = DBEngine.OpenDatabase(m_Config.GetDataPath(), dbFailOnError, False, ";PWD=" & m_Config.GetDatabasePassword())
    
    ' Crear QueryDef con consulta desde modQueries
    Set qdf = db.CreateQueryDef("", GET_EXPEDIENTES_ACTIVOS_SELECTOR)
    
    ' Esta consulta no requiere parámetros
    
    ' Ejecutar consulta
    Set rs = qdf.OpenRecordset()
    
    ' Mapear cada registro a un objeto EExpediente
    Do While Not rs.EOF
        Set expediente = New EExpediente
        expediente.idExpediente = rs!ID
        expediente.Nemotecnico = rs!Nemotecnico & ""
        expediente.Titulo = rs!Titulo & ""
        expediente.ResponsableCalidad = rs!ResponsableCalidad & ""
        expediente.ResponsableTecnico = rs!ResponsableTecnico & ""
        expediente.Pecal = rs!Pecal & ""
        expediente.NumeroExpediente = rs!NumeroExpediente & ""
        expediente.EstadoExpediente = rs!EstadoExpediente & ""
        expediente.EmailResponsable = rs!EmailResponsable & ""
        expediente.JefeProyecto = rs!JefeProyecto & ""
        expediente.ContratistaPrincipal = rs!ContratistaPrincipal & ""
        If Not IsNull(rs!FechaInicio) Then expediente.FechaInicio = rs!FechaInicio
        If Not IsNull(rs!FechaFinPrevista) Then expediente.FechaFinPrevista = rs!FechaFinPrevista
        If Not IsNull(rs!FechaFinContrato) Then expediente.FechaFinContrato = rs!FechaFinContrato
        If Not IsNull(rs!FechaFinGarantia) Then expediente.FechaFinGarantia = rs!FechaFinGarantia
        
        expedientes.Add Key:=expediente.idExpediente, Item:=expediente ' Usar ID como clave
        rs.MoveNext
    Loop
    
    ' Limpiar recursos
    rs.Close
    Set rs = Nothing
    Set qdf = Nothing
    db.Close
    Set db = Nothing
    
    ' Devolver la colección
    Set IExpedienteRepository_ObtenerExpedientesActivosParaSelector = expedientes
    
    Exit Function
    
ErrorHandler:
    ' Limpiar recursos en caso de error
    If Not rs Is Nothing Then rs.Close
    If Not qdf Is Nothing Then Set qdf = Nothing
    If Not db Is Nothing Then db.Close: Set db = Nothing
    
    ' Log del error
    If Not m_ErrorHandler Is Nothing Then
        m_ErrorHandler.LogError Err.Number, Err.Description, "CExpedienteRepository.ObtenerExpedientesActivosParaSelector"
    End If
    
    ' Devolver colección vacía en caso de error
    Set IExpedienteRepository_ObtenerExpedientesActivosParaSelector = New Collection
End Function

' ============================================================================
' MÉTODOS PÚBLICOS DE CONVENIENCIA
' ============================================================================
' Estos métodos exponen públicamente la funcionalidad de la interfaz
' para facilitar las pruebas y el uso directo de la clase concreta

' Método público de conveniencia para ObtenerExpedientePorId
Public Function ObtenerExpedientePorId(ByVal idExpediente As Long) As EExpediente
    Set ObtenerExpedientePorId = IExpedienteRepository_ObtenerExpedientePorId(idExpediente)
End Function

' Método público de conveniencia para ObtenerExpedientePorNemotecnico
Public Function ObtenerExpedientePorNemotecnico(ByVal Nemotecnico As String) As EExpediente
    Set ObtenerExpedientePorNemotecnico = IExpedienteRepository_ObtenerExpedientePorNemotecnico(Nemotecnico)
End Function

' Método público de conveniencia para ObtenerExpedientesActivosParaSelector
Public Function ObtenerExpedientesActivosParaSelector() As Collection
    Set ObtenerExpedientesActivosParaSelector = IExpedienteRepository_ObtenerExpedientesActivosParaSelector()
End Function


