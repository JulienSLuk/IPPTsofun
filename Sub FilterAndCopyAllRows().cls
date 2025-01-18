Sub FilterAndCopyAllRows()

    Dim wsFiltered As Worksheet
    Dim wsMaster As Worksheet
    Dim wsTraining As Worksheet
    Dim lastRow As Long, trainingLastRow As Long
    Dim m2Date As Date
    Dim i As Long
    Dim yearType As String
    Dim compareDate As Date
    
    ' Set references to sheets
    Set wsFiltered = ThisWorkbook.Sheets("Filtered")
    Set wsMaster = ThisWorkbook.Sheets("master")
    
    ' Get date from cell M2 of master sheet
    m2Date = wsMaster.Range("M2").Value
    
    ' Check if "Training Phase" sheet exists, create if it doesn't
    On Error Resume Next
    Set wsTraining = ThisWorkbook.Sheets("Training Phase")
    If wsTraining Is Nothing Then
        Set wsTraining = ThisWorkbook.Sheets.Add
        wsTraining.Name = "Training Phase"
    End If
    On Error GoTo 0
    
    ' Clear the Training Phase sheet
    wsTraining.Cells.Clear
    
    ' Copy headers from Filtered to Training Phase
    wsFiltered.Rows(1).Copy Destination:=wsTraining.Rows(1)
    trainingLastRow = 1 ' Start after headers
    
    ' Get the last row in the Filtered sheet
    lastRow = wsFiltered.Cells(wsFiltered.Rows.Count, "A").End(xlUp).Row
    
    ' Loop through rows in Filtered sheet
    For i = 2 To lastRow ' Start from row 2 (assuming headers in row 1)
        yearType = wsFiltered.Cells(i, "H").Value ' Column H
        
        ' Default value for Column K
        Dim newValue As String
        newValue = yearType ' Keep the original value
        
        If yearType = "Year 1" Then
            compareDate = wsFiltered.Cells(i, "I").Value ' Column I
            If compareDate < m2Date Then
                newValue = "Year 2" ' Update value if condition is met
            End If
        ElseIf yearType = "Year 2" Then
            compareDate = wsFiltered.Cells(i, "J").Value ' Column J
            If compareDate < m2Date Then
                newValue = "ORD" ' Update value if condition is met
            End If
        End If
        
        ' Copy row to Training Phase
        trainingLastRow = trainingLastRow + 1
        wsFiltered.Rows(i).Copy Destination:=wsTraining.Rows(trainingLastRow)
        wsTraining.Cells(trainingLastRow, "K").Value = newValue ' Update Column K
    Next i
    
    ' Delete Columns I and J in Training Phase sheet
    wsTraining.Columns("I:J").Delete
    
    MsgBox "Data has been filtered, copied, and columns I & J have been deleted in 'Training Phase' sheet.", vbInformation
    
End Sub

