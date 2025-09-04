Attribute VB_Name = "TestWorkflowService"
Option Compare Database
Option Explicit


Public Function TestWorkflowServiceRunAll() As CTestSuiteResult
    Set TestWorkflowServiceRunAll = New CTestSuiteResult
    TestWorkflowServiceRunAll.Initialize "TestWorkflowService"
    TestWorkflowServiceRunAll.AddResult TestValidateTransition_ValidCase()
End Function

Private Function TestValidateTransition_ValidCase() As CTestResult
    Set TestValidateTransition_ValidCase = New CTestResult
    TestValidateTransition_ValidCase.Initialize "ValidateTransition con mock válido debe pasar"
    
    Dim serviceImpl As CWorkflowService
    Dim mockRepo As CMockWorkflowRepository
    Dim mockLogger As CMockOperationLogger
    Dim mockErrorHandler As CMockErrorHandlerService
    Dim service As IWorkflowService
    
    On Error GoTo TestFail
    
    ' Arrange
    Set mockRepo = New CMockWorkflowRepository
    mockRepo.ConfigureIsValidTransition True
    
    Set mockLogger = New CMockOperationLogger
    Set mockErrorHandler = New CMockErrorHandlerService
    
    Set serviceImpl = New CWorkflowService
    serviceImpl.Initialize mockRepo, mockLogger, mockErrorHandler
    Set service = serviceImpl
    
    ' Act
    Dim result As Boolean
    result = service.ValidateTransition(1, "A", "B", "PC", "ADMIN")
    
    ' Assert
    modAssert.AssertTrue result, "La transición debería ser válida."
    
    TestValidateTransition_ValidCase.Pass
Cleanup:
    Exit Function
TestFail:
    TestValidateTransition_ValidCase.Fail "Error: " & Err.Description
    Resume Cleanup
End Function



