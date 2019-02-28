Function generateReport(reportName As String, searchedValue As String)

    Dim newSheet As Worksheet
    Dim currDate As String, newName As String
    Dim newSheetRow As Long, lastRow As Long, lastColumn As Long, i As Long, j As Long, nameIndex As Long, phoneIndex As Long
	
    newName = Left((reportName & "_" & Format(Now, "mdhsm")), 31)
    Sheets.Add(After:=Sheets(Sheets.Count)).Name = newName
    Set newSheet = Sheets(newName)
    newSheet.Range("A1:B1").Font.Bold = True
    newSheet.Cells(1, 1).Value = "Employee Name"
    newSheet.Cells(1, 2).Value = "Phone Number"
	latestNewSheetRow = newSheet.Range("A" & newSheet.Rows.Count).End(xlUp).Row + 1
	
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    lastColumn = Cells(1, Columns.Count).End(xlToLeft).Column

    For i = 1 To lastRow
        For j = 1 To lastColumn
            If Cells(i, j) Like searchedValue Then
                newSheet.Range("A" & newSheetRow).Value = Cells(i, nameIndex)
                newSheet.Range("B" & newSheetRow).Value = Cells(i, phoneIndex)
                latestNewSheetRow = latestNewSheetRow + 1
            ElseIf Cells(i, j) Like "*name*" Or Cells(i, j) Like "*Name*" Then
                nameIndex = j
            ElseIf Cells(i, j) Like "*phone*" Or Cells(i, j) Like "*Phone*" Then
                phoneIndex = j
            End If
        Next j
    Next i
    newSheet.Columns("A:B").AutoFit

End Function
