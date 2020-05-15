Attribute VB_Name = "modWebserverAndCallbacks"
Option Explicit
Option Private Module

'* VBA client of Python web server and implementations of DO_GET and DO_POST
'* brought to you by https://exceldevelopmentplatform.blogspot.com
'*


Dim mobjPythonWebServer As Object

Public Const PORT As Long = 80  '* probably best change this to differ from 80 as it will clash no doubt

Private Function TestPythonVBAWebserver_StartWebServer()
    '* create the Python web server
    '* is late bound, no type library available (not out of the box anyway,
    '*  my blog give a type library enabled )
    Set mobjPythonWebServer = VBA.CreateObject("PythonInVBA.PythonVBAWebserver")

    '* give the web server the details of the server address and port number
    '* as well as the callback routines for GET and POST HTTP calls
    Debug.Print mobjPythonWebServer.StartWebServer(Excel.Application, ThisWorkbook.Name & "!VBA_DO_GET", _
            ThisWorkbook.Name & "!VBA_DO_POST", "localhost", PORT)
End Function

Private Sub TestPythonVBAWebserver_StopWebServer()
    '* stops the web server
    If Not mobjPythonWebServer Is Nothing Then
        Debug.Print mobjPythonWebServer.StopWebServer
    End If
End Sub

Private Sub TestPythonVBAWebserver_StopLogging()
    '* stops the logging allowing the log file to be deleted and thus cleared
    If Not mobjPythonWebServer Is Nothing Then
        Debug.Print mobjPythonWebServer.StopLogging
    End If
End Sub

Public Function VBA_DO_GET(arg0, arg1)
    '* Entry point that services HTTP GET requests
    '* do little exccept tell the time
    
    Debug.Print "VBA_DO_GET running"
    VBA_DO_GET = "Time is now " + CStr(Now())
End Function


Public Function VBA_DO_POST(arg0, arg1)
    '* Entry point that services HTTP POST requests
    '* will take a byte array, deserialize it to a 2d grid
    
    Debug.Print "VBA_DO_POST running"

    If TypeName(arg1) = "Byte()" Then

        '* argument comes in as a variant wrapping a byte array but
        '* we strictly need a Byte() so just convert this easily
        Dim abyt() As Byte
        ReDim abyt(LBound(arg1) To UBound(arg1)) As Byte
        Dim lLoop As Long
        For lLoop = LBound(arg1) To UBound(arg1)
            abyt(lLoop) = arg1(lLoop)
        Next

        Dim oPipe As cPipedVariants
        Set oPipe = New cPipedVariants
        
        If oPipe.InitializePipe Then
    
            '* magic line of code which takes a Byte() and yields a Variant Array grid
            Dim vDestination As Variant
            oPipe.DeserializeFromBytes abyt, vDestination
        
            
            If Not IsArray(vDestination) Then
                '* no array, no problem , print the message
                Debug.Print vDestination
            Else
                '* for a 2d grid variant array let's write this to
                '* the first worksheet, Sheet1, (whilst unaffecting workbook's save status)
                Dim wsSheet1 As Excel.Worksheet
                Set wsSheet1 = ThisWorkbook.Worksheets.Item("Sheet1")
                
                Dim bSaved As Boolean
                bSaved = ThisWorkbook.Saved
                wsSheet1.Cells(1, 1).CurrentRegion.Clear
                
                '* get a range of the correct dimensions for the given 2d grid
                Dim rngPaste As Excel.Range
                Set rngPaste = wsSheet1.Cells(1, 1).Resize( _
                            UBound(vDestination, 1) - LBound(vDestination, 1) + 1, _
                            UBound(vDestination, 2) - LBound(vDestination, 2) + 1)
                            
                '* abracadabra, write that grid!   (no text wrap)
                rngPaste.Value = vDestination
                rngPaste.WrapText = False
                
                ThisWorkbook.Saved = bSaved   '* leave workbook's save status unaffected
            End If
        End If
    End If

SingleExit:
    VBA_DO_POST = "VBA_DO_POST ok"
End Function




