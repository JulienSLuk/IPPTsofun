Sub ExtractColumnsToFilteredWithConditions()
    Dim wsSource As Worksheet
    Dim wsFiltered As Worksheet
    Dim lastRowSource As Long
    Dim filteredRow As Long
    Dim cell As Range

    ' Set worksheets
    Set wsSource = ThisWorkbook.Sheets("NSF IPPT")
    Set wsFiltered = ThisWorkbook.Sheets("filtered")
    
    ' Find the last row in column A of the source sheet
    lastRowSource = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row
    
    ' Initialize the filtered row to start from row 2 (to preserve headers)
    filteredRow = 2
    
    ' Loop through each row in the source sheet starting from row 2
    For Each cell In wsSource.Range("A2:A" & lastRowSource)
        ' Check pre-conditions:
        ' Column M must be "Year 1" or "Year 2"
        If wsSource.Cells(cell.Row, "M").Value = "Year 1" Or wsSource.Cells(cell.Row, "M").Value = "Year 2" Then
            ' Column U must be "IPPT"
            If wsSource.Cells(cell.Row, "U").Value = "IPPT" Then
                ' Column V must be "FAIL" or "NOT TAKEN"
                If wsSource.Cells(cell.Row, "V").Value = "FAIL" Or wsSource.Cells(cell.Row, "V").Value = "NOT TAKEN" Then
                    ' Copy columns A, ... to the filtered sheet
                    wsFiltered.Cells(filteredRow, 1).Value = wsSource.Cells(cell.Row, "A").Value ' ID
                    wsFiltered.Cells(filteredRow, 2).Value = wsSource.Cells(cell.Row, "B").Value ' NAME
                    wsFiltered.Cells(filteredRow, 3).Value = wsSource.Cells(cell.Row, "C").Value ' DEPT
                    wsFiltered.Cells(filteredRow, 4).Value = wsSource.Cells(cell.Row, "D").Value ' AGE
                    wsFiltered.Cells(filteredRow, 5).Value = wsSource.Cells(cell.Row, "E").Value ' ADDRESS
                    wsFiltered.Cells(filteredRow, 6).Value = wsSource.Cells(cell.Row, "F").Value ' PHONE
                    wsFiltered.Cells(filteredRow, 7).Value = wsSource.Cells(cell.Row, "G").Value ' EMAIL
                    wsFiltered.Cells(filteredRow, 8).Value = wsSource.Cells(cell.Row, "M").Value ' YEAR TYPE
                    wsFiltered.Cells(filteredRow, 9).Value = wsSource.Cells(cell.Row, "I").Value ' START DATE
                    wsFiltered.Cells(filteredRow, 10).Value = wsSource.Cells(cell.Row, "J").Value ' END DATE
                    wsFiltered.Cells(filteredRow, 12).Value = wsSource.Cells(cell.Row, "V").Value ' REMARKS
                    wsFiltered.Cells(filteredRow, 13).Value = wsSource.Cells(cell.Row, "U").Value ' STATUS

                    
                    ' Move to the next filtered row
                    filteredRow = filteredRow + 1
                End If
            End If
        End If
    Next cell
    
    ' Notify user of completion
    MsgBox "Data has been successfully extracted to the 'filtered' sheet.", vbInformation
End Sub

