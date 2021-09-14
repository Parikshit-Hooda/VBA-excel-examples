Sub getsheetname()


   Dim xlFilePath, xlFileName, xlFilePathFolder As String
   Dim colIdx As Long

    xlFilePath = ThisWorkbook.FullName 'folderPathContainingFile/fileName
    xlFileName = ThisWorkbook.Name ' filename
    xlFilePathFolder = ThisWorkbook.Path 'PathName of the file - folderfolderPathContainingFile
    
    'adding values in the currently selected sheet
    Sheets("Sheet1").Select
    Range("A1").Select
    ActiveCell.Value = "Current Excel File Path" 'fills A1
    ActiveCell.Offset(0, 1).Value = xlFilePath 'fills B1
    Range("A2").Select
    'ActiveCell.Offset(1, -1).Select
    ActiveCell.Value = "Current Excel File Name" 'fills A2
    ActiveCell.Offset(0, 1).Value = xlFileName ' fills B2
    Range("A3").Select
    'ActiveCell.Offset(1, -1).Select
    ActiveCell.Value = "Current Excel File containing folder Path" 'fills A2
    ActiveCell.Offset(0, 1).Value = xlFilePathFolder 'fills B3
    
    'Autofit relevant columns to make every column visible
    Columns("A:B").EntireColumn.AutoFit 'autofits column to make the entire
    'text visible in columns from A to B of the currently selected sheet

    'autofit columns to a cetain max width. Enable if needed
    'For longIdx = 1 To 2
    'If Columns(longIdx).ColumnWidth <= 9 Then 'set min column width
     '   Columns(longIdx).ColumnWidth = 9
    'ElseIf Columns(longIdx).ColumnWidth > 20 Then 'set max column width
     '   Columns(longIdx).ColumnWidth = 20
    'End If
'Next


End Sub
