Call composeAll

Sub composeAll()
'import excel macro module
    
    Dim isFixedMode
    isFixedMode = True
    
    On Error Resume Next
    Dim oApp
    Dim oFso
    Dim sArModule()
    Dim sModule
    Dim sExt
    Dim sourcePath
    Dim parentPath
    Dim targetName
    Dim targetPath
    Dim targetBook
    
    
    xlExcel9795 = 43                   ' //.xls 97-2003 format in Excel 2003 or prev
    xlExcel8 = 56                      ' //.xls 97-2003 format in Excel 2007
    xlTemplate = 17                    ' //.xlt
    xlAddIn = 18                       ' //.xla
    xlExcel12 = 50                     ' //.xlsb
    xlOpenXMLWorkbookMacroEnabled = 52 ' //.xlsm
    xlOpenXMLTemplateMacroEnabled = 53 ' //.xltm
    xlOpenXMLAddIn = 55                ' //.xlam
    
    Set oApp = CreateObject("Excel.Application")
    Set oFso = CreateObject("Scripting.FileSystemObject")
    
    
    If isFixedMode Then
        tmp = getFixedPath
        
        parentPath = tmp(0)
        sourcePath = tmp(1)
        'targetPath = tmp(2)
    Else
        sourcePath = getFolderPath
        If sourcePath = "" Then
            MsgBox "exit this script"
            Exit Sub
        End If
        parentPath = oFso.getParentFolderName(sourcePath)
        
    End If
    targetName = oFso.getFilename(parentPath) & ".xlsm"
    binPath = oFso.buildPath(parentPath, "bin")
    If Not oFso.FolderExists(binPath) Then oFso.createFolder (binPath)
    targetPath = oFso.buildPath(binPath, targetName)
    
    If oFso.FileExists(targetPath) Then
		Call cleanAll(targetPath)
	    Set targetBook = oApp.Workbooks(targetPath).Open
		targetBook.save
    Else
        Set targetBook = oApp.Workbooks.Add
        Call targetBook.SaveAs(targetPath, xlOpenXMLWorkbookMacroEnabled)
    End If
    
    Set oSorceFdr = oFso.getFolder(sourcePath)
    
    
    For Each fl In oSorceFdr.Files
        pn = fl.path
        sExt = LCase(oFso.GetExtensionName(pn))
        
        If (sExt = "cls" Or sExt = "frm" Or sExt = "bas") Then
            Call lfToCrlf(pn)
            Call targetBook.VBProject.VBComponents.Import(pn)
        End If
    Next
    targetBook.Save
    targetBook.Close
	oApp.Quit
    MsgBox "complete!"
End Sub

Function getFolderPath()
'folder picker dialog
    Dim ret
    Dim oShl
    Dim oBrw
    Dim strPath
    On Error Resume Next
    Set oShl = WScript.CreateObject("Shell.Application")
    Set oBrw = oShl.BrowseForFolder(0, "Select sorce folder", &H10)
    If (oBrw Is Nothing) Then
        Err.Clear
        ret = ""
    Else
        ret = oBrw.Items.Item.path
    End If
    Set oShl = Nothing
    Set oBrw = Nothing
    Err.Clear
    On Error GoTo 0
'msgbox "folderPath=" & ret
    getFolderPath = ret
End Function

Sub lfToCrlf(pn)
'change LF to CRLF in the file pn
    Dim oFso
    Dim oStm
    Set oFso = CreateObject("Scripting.FileSystemObject")
    Set oStm = oFso.openTextfile(pn)
    str0 = oStm.readAll
    oStm.Close
    txts = Split(str0, Chr(10))
    Set oStm = oFso.createtextfile(pn)
    For Each txt In txts
        If Right(txt, 1) = Chr(13) Then
            txt = Left(txt, Len(txt) - 1)
            oStm.writeline (txt)
        End If
    Next
    oStm.Close
End Sub

Function getFixedPath()
    Dim oFso
    Dim scriptPath
    Dim targetPath
    Dim sorcePath
    Dim parentPath
    
    Set oFso = CreateObject("Scripting.FileSystemObject")
    parentPath = Replace(WScript.ScriptFullName, WScript.ScriptName, "")
    parentName = oFso.getFilename(parentPath)
    
    sourcePath = oFso.buildPath(parentPath, "src")
    targetPath = oFso.buildPath(parentPath, "bin" & "\" & parentName & ".xlsm")
    
    getFixedPath = Array(parentPath, sourcePath, targetPath)
End Function


Sub cleanAll(targetPath)
	dim oApp
    Set oApp = CreateObject("Excel.Application")
    Set targetBook=oApp.workbooks.open(targetPath)
    vbext_ct_StdModule = 1
    vbext_ct_ClassModule = 2
    vbext_ct_MSForm = 3
    vbext_ct_Document = 100
    
    On Error Resume Next
    Set cmps = targetBook.VBProject.VBComponents
    For Each cmp In cmps
        cn = cmp.Name
        If cmp.Type = vbext_ct_Document Then
            
            Call cmp.CodeModule.DeleteLines(1, cmp.CodeModule.CountOfLines)
        Else
            cmps.Remove (cmp)
        End If
    Next
	targetBook.save
	targetBook.close
	targetBook=Nothing
	oApp.Quit
    On Error GoTo 0
End Sub
