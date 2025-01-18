Sub ExtractColumnsToOutput()
    Dim wsSource As Worksheet
    Dim wsOutput As Worksheet
    Dim lastRowSource As Long
    Dim lastRowOutput As Long
    Dim outputRow As Long
    Dim cell As Range

    ' Set worksheets
    Set wsSource = ThisWorkbook.Sheets("NSF IPPT")
    Set wsOutput = ThisWorkbook.Sheets("output")
    
    ' Find the last row in column A of the source sheet
    lastRowSource = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row
    
    ' Find the last row in the output sheet to append data
    lastRowOutput = wsOutput.Cells(wsOutput.Rows.Count, "A").End(xlUp).Row + 1
    
    ' Initialize the output row to start from row 2 (to preserve headers)
    outputRow = 2
    
    ' Loop through each row in the source sheet and copy required columns
    For Each cell In wsSource.Range("A2:A" & lastRowSource)
        ' Copy the data from each specified column
        wsOutput.Cells(outputRow, 1).Value = cell.Value          ' Column A
        wsOutput.Cells(outputRow, 2).Value = cell.Offset(0, 1).Value ' Column B
        wsOutput.Cells(outputRow, 3).Value = cell.Offset(0, 2).Value ' Column C
        wsOutput.Cells(outputRow, 4).Value = cell.Offset(0, 3).Value ' Column D
        wsOutput.Cells(outputRow, 5).Value = cell.Offset(0, 4).Value ' Column E
        wsOutput.Cells(outputRow, 6).Value = cell.Offset(0, 5).Value ' Column F
        wsOutput.Cells(outputRow, 7).Value = cell.Offset(0, 6).Value ' Column G
        wsOutput.Cells(outputRow, 8).Value = cell.Offset(0, 8).Value ' Column M
        
        ' Move to the next row in the output sheet
        outputRow = outputRow + 1
    Next cell
    
    ' Notify user of completion
    MsgBox "Data has been successfully extracted to the 'Output' sheet.", vbInformation
End Sub

