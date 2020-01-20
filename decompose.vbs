Call decomposeAll

Sub decomposeAll()
'export excel macro module
    
    Dim isFixedMode
    isFixedMode = True
    
    Dim oApp
    Dim oFso
    
    Dim module
    Dim modules
    Dim ext
    
    Dim parentPath
    Dim sourcePath
    Dim targetPath
    Dim sFilePath
    Dim TargetBook
    
    Dim bn
    Dim xn
    
    Set oApp = CreateObject("Excel.Application")
    Set oFso = CreateObject("Scripting.FileSystemObject")
    Set oShl = CreateObject("Shell.Application")
    
    
    If isFixedMode Then
        tmp = getFixedPath
        parentPath = tmp(0)
        sourcePath = tmp(1)
        targetPath = tmp(2)
    Else
        targetPath = getFilePath
        
        If targetPath = "" Then
            MsgBox "exit this script"
            Exit Sub
        End If
        
		prn= oFso.GetParentFolderName(targetPath)
        bn = oFso.GetBaseName(targetPath)
        xn = oFso.GetExtensionName(targetPath)
        
        If Left(xn, 3) <> "xls" Then
            MsgBox "this file is not Excel File"
            Exit Sub
        End If
        
        parentPath =oFso.buildPath( prn , bn)
        sourcePath =oFso.buildPath( parentPath ,"src")
    End If
    
    If Not oFso.FolderExists(parentPath) Then oFso.createFolder (parentPath)
    If Not oFso.FolderExists(sourcePath) Then oFso.createFolder (sourcePath)
    
    
    vbext_ct_StdModule = 1
    vbext_ct_ClassModule = 2
    vbext_ct_MSForm = 3
 
       Set TargetBook = oApp.Workbooks.Open(targetPath)
    
    Set modules = TargetBook.VBProject.VBComponents
    
    For Each module In modules
        ext = ""
        If (module.Type = vbext_ct_ClassModule) Then
            ext = "cls"
        ElseIf (module.Type = vbext_ct_MSForm) Then
            ext = "frm"
        ElseIf (module.Type = vbext_ct_StdModule) Then
            ext = "bas"
        End If
        
        If ext <> "" Then
            sFilePath = oFso.buildPath(sourcePath , module.Name & "." & ext)
            Call module.Export(sFilePath)
            
        End If
    Next
    TargetBook.Close
	oApp.Quit
    MsgBox "Complete!"
End Sub

Function getFilePath()
    
    Dim oShl
    Dim oBrw
    Dim strPath
    On Error Resume Next
    Set oShl = WScript.CreateObject("Shell.Application")
    Set oBrw = oShl.BrowseForFolder(0, "Select Excel macro file", &H4000)
    
    If (oBrw Is Nothing) Then
        Err.Clear
        getFilePath = ""
    Else
        getFilePath = oBrw.Items.Item.Path
    End If
    
    Set oShl = Nothing
    Set oBrw = Nothing
    Err.Clear
    On Error GoTo 0
    
End Function

Function getFixedPath()
    Dim oFso
    Dim scriptPath
    Dim targetPath
    Dim sorcePath
    Dim parentPath
    
    Set oFso = CreateObject("Scripting.FileSystemObject")
    parentPath = Replace(WScript.ScriptFullName, WScript.ScriptName, "")
    parentName = oFso.getFilename(parentPath)

    sourcePath =oFso.buildPath( parentPath ,"src")
    targetPath =oFso.buildPath( parentPath , "bin" & "\" & parentName & ".xlsm")

    getFixedPath = Array(parentPath, sourcePath, targetPath)
End Function

