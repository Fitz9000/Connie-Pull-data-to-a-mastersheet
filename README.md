Option Explicit
Option Compare Binary

'This sub procedure pulls data from a single cell on an excel sheet and pastes it onto the sheet you run this from
Sub import()
    
    Dim masterFile As String
    Dim destinationTabMasterFile As String
    Dim destinationOutputColumnMasterFile As String
    Dim lastRowMasterFile As Long
    
    Dim workbookUse As Boolean
    Dim importTab As String
    Dim importRange As String
    Dim folderPath As String
    Dim importFile As String
    Dim importPath As String
    
    'Master File variables
    masterFile = ThisWorkbook.Name
    destinationTabMasterFile = "Put Values Here"                'tab you want to put data in
    destinationOutputColumnMasterFile = "A"                     'column to pull out put data once pulled
    
    lastRowMasterFile = LastRow(destinationOutputColumnMasterFile, destinationTabMasterFile)
    
    'Import File variables
    folderPath = "C:\Users\cHo\some folder1\some folder2"       'folder path of the file where you want to get data from
    importFile = "some spreadsheet.xlsx"                        'excel file that you want to grab data from
    importTab = "Some tab name"                                 'tab that you want to get data from
    importRange = "A1"                                          'range where you want to get data from
    
    
    importPath = folderPath & importFile
        
        'Check if import path exists
        If Dir(importPath, vbDirectory) <> "" Then
        
            'Check if workbook is already open/in use
            If IsWorkBookOpen(importPath) Then
                workbookUse = True
            Else
                Workbooks.Open importPath
                workbookUse = False
            End If
            
            'Open workbook
            Workbooks(importFile).Sheets(importTab).Activate
            
            'Transfer Document Number data
            Workbooks(importFile).Worksheets(importTab).Range(importRange).Copy
            Workbooks(masterFile).Worksheets(destinationTabMasterFile).Range(destinationOutputColumnMasterFile & lastRowMasterFile).Offset(1, 0).PasteSpecial xlPasteValuesAndNumberFormats
            
            'If workbook is in use: don't close workbook
            If workbookUse = False Then
                Workbooks(importFile).Close
            End If
        
        Else
            MsgBox "Import path: " & importPath & " could not be found" _
                , vbExclamation, "Import Error"
        End If

End Sub

'Returns "TRUE" if file is open
Private Function IsWorkBookOpen(FileName As String) As Boolean

    Dim ff As Long, ErrNo As Long

    On Error Resume Next
    ff = FreeFile()
    Open FileName For Input Lock Read As #ff
    Close ff
    ErrNo = Err
    On Error GoTo 0
    
    Select Case ErrNo
    Case 0:    IsWorkBookOpen = False
    Case 70:   IsWorkBookOpen = True
    Case Else: Error ErrNo
    End Select
    
End Function

'Returns the number of the last row in a column
Private Function LastRow(lookupCol As String, lookupSheet As String) As Integer
    
    'Returns the number of the last row in a column
    LastRow = Worksheets(lookupSheet).Range(lookupCol & Rows.Count).End(xlUp).Row

End Function
