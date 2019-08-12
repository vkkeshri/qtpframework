'***********************************************************************************************************************************************************************************
'*				@				File Name      				                : 						SuitExecuter																																											      *
'*				@				Description         						: 						This is the executer file for this framework 																							          *
'*				@				Author              							 : 						Vinod Keshri																																																				 * 
'*				@				Date                							  : 					  01/07/2014																																																							 *
'*				@				Updated Date        				   : 					  -------																																															           *
'***********************************************************************************************************************************************************************************

'Check excel if opened then close
reportManager.killExcelProcess

'Check the suite name
suiteName = Environment.Value("SUITE_NAME")
If suiteName = "" Then
    halt "Error, the environment variable SUITE_NAME was not set."
End If

'check the project directory
projectDir = Environment.Value("PROJECT_DIR")
If projectDir = "" Then
    halt "Error, the environment variable PROJECT_DIR was not set."
End If

'Start the test run
Set errors = executionManager.loadSuite(suiteName, projectDir)
If errors.size() > 0 Then
    halt errors.toString()
End If

'Create the new report
result = reportManager.startReport(suiteName)
If result <> "" Then
    halt result
End If

'Run the actual suite of tests
Set errors = executionManager.executeSuite()
If errors.size() > 0 Then
   halt errors.toString()
End If

'Terminate reporting
reportManager.stopReport

' Send email notification
Set email = newEmail
email.send

'Support 
Private Sub halt(errorMessage)
    error "SuiteExecutor Error: " & errorMessage
    If Environment("ERROR_DIALOG_ON_END") = "True" Then
        MsgBox errorMessage, vbCritical, "SuiteExecutor Error"
        ExitActionIteration(errorMessage)
    End If
End Sub
'***********************************************************************************************************************************************************************************
