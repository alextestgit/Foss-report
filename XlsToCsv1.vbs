Dim vExcel, vCSV, vFolder
Set vExcel = CreateObject("Excel.Application")

' get path to folder with XLS files
vFolder = Wscript.Arguments.Item(0)

' get list of XLS in folder
Set fso = CreateObject("Scripting.FileSystemObject")
Set folder = fso.GetFolder(vFolder)
Set files = folder.Files

For Each file In files
    If LCase(Right(file.Name, 4)) = ".xls" Then
        ' Open XLS
        Set vCSV = vExcel.Workbooks.Open(file.Path)

        ' save XLS in format CSV.
        vCSV.SaveAs Replace(file.Path, ".xls", ".csv"), 6

        ' Close XLS.
        vCSV.Close False
    End If
Next

' exit from Excel.
vExcel.Quit