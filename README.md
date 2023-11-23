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


'This function add
Public Sub ImportExcelSpreadsheet(Filename As String, tableName As String)
 DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12, tableName, Filename, True
End Sub

'This Function is used only once!!
'The first time you add data in the data base

Public Sub ImportFromMultipleExcelSpreadsheets()
'Dim excelapp As New Excel.Application
Dim excelApp As Object
Set excelApp = CreateObject("Excel.Application")
'Dim excelbook As New Excel.Workbook
Dim excelbook As Object
Set excelbook = excelApp.Workbooks.Add
'Dim excelsheet As New Excel.Worksheet
'Dim excelsheet As Object
'Set excelsheet = excelbook.Sheets
Dim intNoOfSheets As Integer, intCounter As Integer
Dim strFilePath As String, strLastDataColumn As String
Dim strLastDataRow As String, strLastDataCell As String
Dim tableName As String
Dim rangeName(15) As String
Dim i As Integer

On Error GoTo errorHandler
    tableName = "Auslieferung"
    strFilePath = "\\DC\Users\c.kameni\Eigene Dateien\Programmierung\VBA\Excel\Auslieferungen TJ_MASTER.xlsb"
    
    Set excelbook = excelApp.Workbooks.Open(strFilePath)
    
     intNoOfSheets = excelbook.Worksheets.Count
    
     Dim CurrSheetName As String
    
    For intCounter = 1 To intNoOfSheets
        rangeName(intCounter) = excelbook.Worksheets(intCounter).Name & "!" & Replace(excelbook.Worksheets(intCounter).UsedRange.Address, "$", "")
        Debug.Print rangeName(intCounter)
    Next
    excelbook.Close
    excelApp.Quit
    Set excelApp = Nothing
        
    For i = 1 To intNoOfSheets
        If rangeName(i) = "" Then Exit For
        DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel9, tableName, strFilePath, True, rangeName(i)
    Next i

ExitSub:
 
    Exit Sub
errorHandler:
    MsgBox Err.Description
    Resume ExitSub
End Sub

