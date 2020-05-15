Attribute VB_Name = "devHelpers"
Option Explicit
Option Private Module

Private Sub TestAppRun()
    '* sometimes during development macros can become locked or disabled
    '* this is a quickie to test execution

    Call Application.Run(ThisWorkbook.Name & "!VBA_DO_POST", "\", Array("foo"))
End Sub


Private Sub PickupNewPythonScript()
    '# for development only to help pick up script changes we kill the python process
    Call CreateObject("WScript.Shell").Run("taskkill /f /im pythonw.exe", 0, True)
    
End Sub

