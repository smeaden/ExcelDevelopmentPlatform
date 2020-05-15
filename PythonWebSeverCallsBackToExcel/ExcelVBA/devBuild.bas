Attribute VB_Name = "devBuild"
Option Explicit
Option Private Module

Private Sub BuildWorkbook()
    '* START READING HERE!
    '* To avoid checking in an Excel workbook which will be too large and does not
    '* delta diff the source code modules we check in the code modules separately
    '* but then this necessitates a workbook build routine.  So this piece of code will
    '* import the necessary VBA code modules to build a workbook.
    
    '* import the other components
    ImportVBComponentsToThisWorkbook
    
    '* rename the VBA prokect
    ThisWorkbook.VBProject.Name = "vbaJobLeadsCache"
    
    '* check this workbook has been named/saved as expected
    If StrComp(ThisWorkbook.Name, "JobLeadsCache.xlsm", vbTextCompare) <> 0 Then
        Debug.Print "ThisWorkbook needs saving as JobLeadsCache.xlsm"
    End If
    

End Sub

Private Function VBComponentFileNames() As Variant
    VBComponentFileNames = Array("devBuild.bas", "devHelpers.bas", "modWebserverAndCallbacks.bas", _
                            "tstPipedVariants.bas", "cPipedVariants.cls")
End Function

Private Sub ImportVBComponentsToThisWorkbook()
    ImportVBComponents ThisWorkbook
End Sub

Private Sub ExportVBComponentsFromThisWorkbook()
    ExportVBComponents ThisWorkbook
End Sub

Private Sub RemoveVBComponentsFromThisWorkbook()
    RemoveVBComponents ThisWorkbook, Array("devBuild", "Sheet1", "ThisWorkbook")
End Sub


Private Sub ImportVBComponents(ByVal wbWorkbook As Excel.Workbook)
    
    Dim vLoop
    For Each vLoop In VBComponentFileNames
        Dim sPath As String:
        sPath = FSO.BuildPath(wbWorkbook.Path, vLoop)
        
        If Not FSO.FileExists(sPath) Then
            Debug.Print "Warning: could not find file  '" & sPath & "', import failed, continuing."
        Else
            Dim sCompName As String
            sCompName = SplitOffFileExt(vLoop)
            If VBComponentExists(wbWorkbook, sCompName) Then
                Debug.Print "Warning: module '" & sCompName & "' already exists, will not re-import.  Delete the module first if you want to import."
            Else
            
                wbWorkbook.VBProject.VBComponents.Import sPath
                Debug.Print "Imported VBComponent '" & sCompName & "' from " & sPath
            
            End If
        End If
    Next
End Sub


Private Sub RemoveVBComponents(ByVal wbWorkbook As Excel.Workbook, ByVal vExcept As Variant)
    Dim dicVBComponents As Object
    Set dicVBComponents = CollectionToDict(wbWorkbook.VBProject.VBComponents)
    
    Dim vExceptLoop
    For Each vExceptLoop In vExcept
        If dicVBComponents.Exists(vExceptLoop) Then dicVBComponents.Remove vExceptLoop
    Next
    
    Dim vLoop As Variant
    For Each vLoop In dicVBComponents.Keys
        
        Dim vbcLoop As Object
        Set vbcLoop = wbWorkbook.VBProject.VBComponents.Item(vLoop)
        wbWorkbook.VBProject.VBComponents.Remove vbcLoop
            
        Debug.Print "Removed VBComponent '" & vLoop & "' "
    Next
End Sub





Private Sub ExportVBComponents(ByVal wbWorkbook As Excel.Workbook)
    Dim vLoop
    For Each vLoop In VBComponentList
    
        If Not VBComponentExists(wbWorkbook, vLoop) Then
            Debug.Print "Warning could not find module '" & vLoop & "'"
        Else
            Dim vbcLoop As Object
            Set vbcLoop = GetVBComponentSafe(wbWorkbook, vLoop)
            
            Dim sPath As String:
            sPath = FSO.BuildPath(wbWorkbook.Path, vLoop & FileExt(vbcLoop))
            vbcLoop.Export sPath
            
            Debug.Print "Exported VBComponent '" & vLoop & "' to " & sPath
        End If
    Next
End Sub

Private Function CollectionToDict(col) As Object
    Dim dic As Object
    Set dic = VBA.CreateObject("Scripting.Dictionary")
    Dim v
    For Each v In col
        If Not dic.Exists(v.Name) Then dic.Add v.Name, ""
    Next

    Set CollectionToDict = dic
End Function

Private Function VBComponentList()
    Dim vCopy As Variant
    vCopy = VBComponentFileNames
    
    Dim idx As Long
    For idx = 0 To UBound(vCopy)
        vCopy(idx) = SplitOffFileExt(vCopy(idx))
    Next
    VBComponentList = vCopy
End Function

Private Function FSO() As Object
    Static stfso As Object
    If stfso Is Nothing Then Set stfso = VBA.CreateObject("Scripting.FileSystemObject")
    Set FSO = stfso
End Function

Private Function FileExt(vbc) As String
    FileExt = VBA.Switch(vbc.Type = 1, ".bas", vbc.Type = 2, ".cls")
End Function

Private Function GetVBComponentSafe(ByVal wbWorkbook As Excel.Workbook, ByVal sModuleName As String) As Object
    On Error Resume Next
    Set GetVBComponentSafe = wbWorkbook.VBProject.VBComponents.Item(sModuleName)
End Function

Private Function VBComponentExists(ByVal wbWorkbook As Excel.Workbook, ByVal sModuleName As String) As Boolean
    VBComponentExists = Not GetVBComponentSafe(wbWorkbook, sModuleName) Is Nothing
End Function

Private Function SplitOffFileExt(s)
    SplitOffFileExt = VBA.Split(s, ".")(0)
End Function
