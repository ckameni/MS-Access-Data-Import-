# MS-Access-Data-Import-



Private Sub btn_Browse_Click()
'microsoft office object library
Dim diag As Office.FileDialog
Dim item As Variant
 
Set diag = Application.FileDialog(msoFileDialogFilePicker)

On Error GoTo Fehler
diag.AllowMultiSelect = False ' only select one file
diag.Title = "Please select an excel Spreadsheet"
diag.Filters.Clear
diag.Filters.Add "Excel SpreadSheets", "*.xls, *.xlsx"
 
'diag.Show ' return a long value indicating how many items where selected
 
If diag.Show Then
For Each item In diag.SelectedItems
Me.txtFilename = item
Next item
End If
Quit:
    Exit Sub
Fehler:
    MsgBox Err.Description
Resume Quit
End Sub

Private Sub btnDatentImport_Click()

'Verweis=>: Mirosoft scripting Runtime
Dim FSO As New FileSystemObject
'Dim tableName As String
On Error GoTo errorHandler
If Nz(Me.txtFilename, "") = "" Then
MsgBox "Please select a file"
Exit Sub
End If
 
If FSO.FileExists(Nz(Me.txtFilename, "")) Then
'tableName = FSO.GetFileName(Me.txtFilename) ' i give my Table the name of the Excel File
Call mdlDataImport.MainImport(Me.txtFilename)

Else
MsgBox "File Not Found"
End If
 
ExitSub:
Exit Sub
errorHandler:
MsgBox "The File you tried to import was not an excel spreadsheet"
Me.txtFilename = vbNullString
Resume ExitSub
End Sub

