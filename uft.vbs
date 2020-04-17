Dim qtApp 
Dim qtTest 
Set qtApp = CreateObject("QuickTest.Application")
qtApp.Launch
qtApp.Visible = True
qtApp.Open "C:/Users/dea/.jenkins/workspace/ExportExecutionResult/DairyManagement1",True
Set qtTest = qtApp.Test
qtTest.Run 
qtApp.Quit
Set qtTest = Nothing' Release the Test object
Set qtApp = Nothing' Release the Application object