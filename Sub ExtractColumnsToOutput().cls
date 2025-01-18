Sub ExtractColumnsToOutputWithConditions()
    Dim wsSource As Worksheet
    Dim wsOutput As Worksheet
    Dim lastRowSource As Long
    Dim outputRow As Long
    Dim cell As Range

    ' Set worksheets
    Set wsSource = ThisWorkbook.Sheets("NSF IPPT")
    Set wsOutput = ThisWorkbook.Sheets("Output")
    
    ' Find the last row in column A of the source sheet
    lastRowSource = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row
    
    ' Initialize the output row to start from row 2 (to preserve headers)
    outputRow = 2
    
    ' Loop through each row in the source sheet starting from row 2
    For Each cell In wsSource.Range("A2:A" & lastRowSource)
        ' Check pre-conditions:
        ' Column M must be "Year 1" or "Year 2"
        If wsSource.Cells(cell.Row, "M").Value = "Year 1" Or wsSource.Cells(cell.Row, "M").Value = "Year 2" Then
            ' Column U must be "IPPT"
            If wsSource.Cells(cell.Row, "U").Value = "IPPT" Then
                ' Column V must be "FAIL" or "NOT TAKEN"
                If wsSource.Cells(cell.Row, "V").Value = "FAIL" Or wsSource.Cells(cell.Row, "V").Value = "NOT TAKEN" Then
                    ' Copy columns A, B, C, D, E, F, G, and M to the output sheet
                    wsOutput.Cells(outputRow, 1).Value = wsSource.Cells(cell.Row, "A").Value ' Column A
                    wsOutput.Cells(outputRow, 2).Value = wsSource.Cells(cell.Row, "B").Value ' Column B
                    wsOutput.Cells(outputRow, 3).Value = wsSource.Cells(cell.Row, "C").Value ' Column C
                    wsOutput.Cells(outputRow, 4).Value = wsSource.Cells(cell.Row, "D").Value ' Column D
                    wsOutput.Cells(outputRow, 5).Value = wsSource.Cells(cell.Row, "E").Value ' Column E
                    wsOutput.Cells(outputRow, 6).Value = wsSource.Cells(cell.Row, "F").Value ' Column F
                    wsOutput.Cells(outputRow, 7).Value = wsSource.Cells(cell.Row, "G").Value ' Column G
                    wsOutput.Cells(outputRow, 8).Value = wsSource.Cells(cell.Row, "M").Value ' Column M
                    
                    ' Move to the next output row
                    outputRow = outputRow + 1
                End If
            End If
        End If
    Next cell
    
    ' Notify user of completion
    MsgBox "Data has been successfully extracted to the 'Output' sheet.", vbInformation
End Sub

