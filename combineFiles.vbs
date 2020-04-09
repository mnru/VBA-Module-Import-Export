call main
Sub main()
  Set app = CreateObject("Excel.Application")
	app.visible=true
  fls = app.GetOpenFilename("all files,*.*", , "select files to combine", , True)
  sfn = app.GetSaveAsFilename("combined.txt","all files,*.*", , "set save file name")
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set saveStm = fso.CreateTextFile(sfn)
  For Each fl In fls
    Set eachStm = fso.OpenTextFile(fl)
    txt = eachStm.ReadAll
    saveStm.WriteLine (txt)
    eachStm.Close
  Next
  saveStm.Close
End Sub