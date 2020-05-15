VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cPipedVariants"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'* Pipe-based helper to serialize/deserialize VB-UDTs InMemory ... [Olaf Schmidt 2015]
'* Based on code by Olaf Schmidt 2015, http://www.vbforums.com/showthread.php?807205-VB6-pipe-based-UDT-serializing-deserializing-InMemory



'* https://docs.microsoft.com/en-us/windows/win32/api/winbase/nf-winbase-createnamedpipea
Private Declare Function CreateNamedPipeW& Lib "kernel32" (ByVal lpName As Long, ByVal dwOpenMode&, ByVal dwPipeMode&, _
            ByVal nMaxInstances&, ByVal nOutBufferSize&, ByVal nInBufferSize&, _
            ByVal nDefaultTimeOut&, ByVal lpSecurityAttributes&)

'* https://docs.microsoft.com/en-us/windows/win32/api/fileapi/nf-fileapi-writefile
Private Declare Function WriteFile& Lib "kernel32" (ByVal hFile&, lpBuffer As Any, _
            ByVal nNumberOfBytesToWrite&, lpNumberOfBytesWritten&, ByVal lpOverlapped&)
            
'* https://docs.microsoft.com/en-us/windows/win32/api/fileapi/nf-fileapi-readfile
Private Declare Function ReadFile& Lib "kernel32" (ByVal hFile&, lpBuffer As Any, _
            ByVal nNumberOfBytesToRead&, lpNumberOfBytesRead&, ByVal lpOverlapped&)
            
'* https://docs.microsoft.com/en-us/windows/win32/api/namedpipeapi/nf-namedpipeapi-peeknamedpipe
Private Declare Function PeekNamedPipe& Lib "kernel32" (ByVal hNamedPipe&, lpBuffer As Any, _
            ByVal nBufferSize&, lpBytesRead&, lpTotalBytesAvail&, lpBytesLeftThisMessage&)
            
'* https://docs.microsoft.com/en-us/windows/win32/api/namedpipeapi/nf-namedpipeapi-disconnectnamedpipe
Private Declare Function DisconnectNamedPipe& Lib "kernel32" (ByVal hPipe&)

Private Declare Function CloseHandle& Lib "kernel32" (ByVal hObject&)
 
Private mhPipe As Long
Private mlFileNumber As Long
Private mabytSerialized() As Byte

Private Enum eOpenMode
    PIPE_ACCESS_INBOUND = 1
    PIPE_ACCESS_OUTBOUND = 2
    PIPE_ACCESS_DUPLEX = 3
End Enum

Private Enum ePipeMode
    PIPE_TYPE_BYTE = 0
    PIPE_TYPE_MESSAGE = 4

    PIPE_READMODE_BYTE = 0
    PIPE_READMODE_MESSAGE = 2
    
    PIPE_WAIT = 0
    PIPE_NOWAIT = 1
End Enum

Private Enum ePipeInstances
    PIPE_UNLIMITED_INSTANCES = 255
End Enum

Public Function InitializePipe(Optional sPipeNameSuffix As String = "vbaPipedVariantArrays") As Boolean
    Const csPipeNamePrefix As String = "\\.\pipe\"
    CleanUp
    
    Dim sPipeName As String
    sPipeName = csPipeNamePrefix & sPipeNameSuffix
    
    '* Must call CreateNamedPipe first before calling Open <<pathname>> For Binary otherwise you get bad file number
    mhPipe = CreateNamedPipeW(StrPtr(sPipeName), PIPE_ACCESS_DUPLEX, PIPE_TYPE_BYTE + PIPE_READMODE_BYTE + PIPE_WAIT, _
            PIPE_UNLIMITED_INSTANCES, -1, -1, 0, 0)
            
    If mhPipe = -1 Then mhPipe = 0 'reset from InvalidHandleValue to "no Handle"
    
    If mhPipe Then
        '* only try to find a free VB-FileNumber when mhPipe is valid (i.e. pipe has been created)
        mlFileNumber = FreeFile
        If mlFileNumber Then
            Open sPipeName For Binary As mlFileNumber  'open only, when we got an mlFileNumber
        End If
    End If
    
    InitializePipe = mhPipe <> 0 And mlFileNumber <> 0
End Function

Public Function SerializeToBytes(ByRef vSrc As Variant, ByRef pabytSerialized() As Byte) As Long

    Dim lBytesAvail As Long

    Debug.Assert IsArray(vSrc)

    If mlFileNumber <> 0 Then
    
        '* this next line writes the Variant array to the pipe
        Put mlFileNumber, 1, vSrc
        
        '* we should now have some bytes to read out of the pipe, use PeekNamedPipe to verify there are bytes available
        PeekNamedPipe mhPipe, ByVal 0&, 0, ByVal 0&, lBytesAvail, 0
        
        If lBytesAvail > 0 Then
            
            '* so now we can dimension the byte array
            ReDim Preserve pabytSerialized(0 To lBytesAvail - 1)
            
            '* and now we can read the bytes out of the pipe and into the byte array
            ReadFile mhPipe, pabytSerialized(0), lBytesAvail, lBytesAvail, ByVal 0&
            
            '* return number of bytes as a courtesy (not actually required)
            SerializeToBytes = lBytesAvail
        End If
    End If
End Function

Public Function DeserializeFromVariantWrappedByteArray(ByRef vSerialized As Variant, ByRef pvDest As Variant) As Long

    If TypeName(vSerialized) = "Byte()" Then

        '* argument comes in as a variant wrapping a byte array but
        '* we strictly need a Byte() so just convert this easily
        Dim abyt() As Byte
        ReDim abyt(LBound(vSerialized) To UBound(vSerialized)) As Byte
        Dim lLoop As Long
        For lLoop = LBound(vSerialized) To UBound(vSerialized)
            abyt(lLoop) = vSerialized(lLoop)
        Next
        
        DeserializeFromVariantWrappedByteArray = Me.DeserializeFromBytes(abyt(), pvDest)
    Else
        Err.Raise vbObjectError, "", "Not appropriate to call this method we input parameter type!"
    End If

End Function

Public Function DeserializeFromBytes(ByRef abytSerialized() As Byte, ByRef pvDest As Variant) As Long
    
    Dim lBytesWritten As Long
    
    If mhPipe <> 0 And mlFileNumber <> 0 Then

        '* write the byte array to the pipe
        WriteFile mhPipe, abytSerialized(0), UBound(abytSerialized) + 1, lBytesWritten, 0
        
        If lBytesWritten = UBound(abytSerialized) + 1 Then
            '* the pipe contains a byte array serialization of a variant array
            '* we can use VBA's Get statement to read it directly into a variant array variable
            Get mlFileNumber, 1, pvDest
            
            '* report the amount of deserialized Bytes as a courtesy (not actually required)
            DeserializeFromBytes = Loc(mlFileNumber)
        End If
    End If
End Function

Private Sub CleanUp()
    If mlFileNumber Then Close mlFileNumber: mlFileNumber = 0
    If mhPipe Then DisconnectNamedPipe mhPipe
    If mhPipe Then CloseHandle mhPipe: mhPipe = 0
End Sub
 
Private Sub Class_Terminate()
    CleanUp
End Sub




