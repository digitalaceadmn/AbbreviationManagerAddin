Sub ForceJSSDRibbonVisible()
    '
    ' ForceJSSDRibbonVisible Macro
    ' This macro forces the JSSD ribbon tab to be visible in newer Word versions
    ' To use: Press Alt+F11, paste this code in a new module, then run it
    '
    
    On Error Resume Next
    
    ' Method 1: Try to access the add-in directly
    Dim addin As Object
    Set addin = Application.COMAddIns("AbbreviationWordAddin").Object
    
    If Not addin Is Nothing Then
        ' Try different method names
        On Error Resume Next
        
        ' Try the main method
        addin.ForceJSSDRibbonVisible
        If Err.Number = 0 Then
            MsgBox "JSSD ribbon refresh attempted via add-in method (ForceJSSDRibbonVisible).", vbInformation, "Ribbon Fix"
            Exit Sub
        End If
        Err.Clear
        
        ' Try accessing the ribbon directly via Globals.Ribbons
        Dim ribbonTab As Object
        Set ribbonTab = addin.Ribbons.AbbreviationRibbon
        If Not ribbonTab Is Nothing Then
            ' Try the simple test method first
            ribbonTab.TestRibbonAccess
            If Err.Number = 0 Then
                Exit Sub
            End If
            Err.Clear
            
            ' Try the main method
            ribbonTab.ForceJSSDTabVisible
            If Err.Number = 0 Then
                MsgBox "JSSD ribbon refresh attempted via ribbon object.", vbInformation, "Ribbon Fix"
                Exit Sub
            End If
        End If
        
        On Error GoTo 0
    End If
    
    ' Method 2: Try alternative approach
    Dim ribbon As Object
    Set ribbon = Application.CommandBars.ActiveMenuBar
    
    ' Force ribbon invalidation
    Application.CommandBars.ExecuteMso "TabAddIns"
    
    MsgBox "JSSD ribbon refresh attempted. If the tab is still not visible, please:" & vbCrLf & _
           "1. Restart Microsoft Word" & vbCrLf & _
           "2. Check if the add-in is enabled in File > Options > Add-ins" & vbCrLf & _
           "3. Try reinstalling the AbbreviationWordAddin", vbInformation, "Ribbon Fix Instructions"
    
End Sub

Sub CheckAddInStatus()
    '
    ' CheckAddInStatus Macro  
    ' This macro checks if the AbbreviationWordAddin is properly loaded
    '
    
    Dim addin As COMAddIn
    Dim found As Boolean
    found = False
    
    For Each addin In Application.COMAddIns
        If InStr(addin.ProgId, "AbbreviationWordAddin") > 0 Or InStr(addin.Description, "AbbreviationWordAddin") > 0 Then
            found = True
            MsgBox "Add-in found:" & vbCrLf & _
                   "Name: " & addin.Description & vbCrLf & _
                   "Connected: " & addin.Connect & vbCrLf & _
                   "Path: " & addin.ProgId, vbInformation, "Add-in Status"
            Exit For
        End If
    Next addin
    
    If Not found Then
        MsgBox "AbbreviationWordAddin not found in COM Add-ins list." & vbCrLf & _
               "Please check if the add-in is properly installed and enabled in:" & vbCrLf & _
               "File > Options > Add-ins > COM Add-ins", vbExclamation, "Add-in Not Found"
    End If
    
End Sub