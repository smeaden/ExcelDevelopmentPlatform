Attribute VB_Name = "tstPipedVariants"
Option Explicit
Option Private Module

'* This module is test client code for the cPipedVariants class



Sub SamePipeForSerializeAndDeserialize()
    Dim oPipe As cPipedVariants
    Set oPipe = New cPipedVariants
    
    If oPipe.InitializePipe Then
        Dim vSource As Variant
        vSource = TestData


        Dim abytSerialized() As Byte

        Call oPipe.SerializeToBytes(vSource, abytSerialized)
        
        Stop '* at this point vSource is populated but vDestination is empty

        Dim vDestination As Variant
        oPipe.DeserializeFromBytes abytSerialized, vDestination
    
        Stop
    End If
End Sub

Function TestData() As Variant
    Dim vSource(1 To 2, 1 To 4) As Variant
    vSource(1, 1) = "Hello World"
    vSource(1, 2) = True
    vSource(1, 3) = False
    vSource(1, 4) = Null
    vSource(2, 1) = 65535
    vSource(2, 2) = 7.5
    vSource(2, 3) = CDate("12:00:00 16-Sep-1989") 'now()
    vSource(2, 4) = CVErr(xlErrNA)
    TestData = vSource
End Function

