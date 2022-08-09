Attribute VB_Name = "ProcessButton"
Function getMaxRow(col As Integer) As Double
    getMaxRow = ActiveSheet.Cells(Rows.Count, col).End(xlUp).row
End Function

Function getMaxCol(row As Integer) As Double
    getMaxCol = ActiveSheet.Cells(row, Columns.Count).End(xlToLeft).Column
End Function

Function getFilenameFromPath(ByVal strPath As String) As String
    If Right$(strPath, 1) <> "\" And Len(strPath) > 0 Then
        getFilenameFromPath = getFilenameFromPath(Left$(strPath, Len(strPath) - 1)) + Right$(strPath, 1)
    End If
End Function

Function findCellInColumn(row As Integer, str As String) As Double
    Dim i As Double
    i = 1
    Dim m As Double
    m = getMaxCol(row)
    While Application.WorksheetFunction.Trim(LCase(ActiveSheet.Cells(row, i).Value)) <> Application.WorksheetFunction.Trim(LCase(str)) And i <= m
        i = i + 1
    Wend
    findCellInColumn = i
End Function

Sub Process()
    Dim file1 As String
    Dim file2 As String
    file1 = Range("B1").Value
    file2 = Range("B5").Value
    Workbooks.Open Filename:=file1, UpdateLinks:=0
    
    Dim path_file1 As String
    path_file1 = ActiveWorkbook.Path

    Dim max_row As Double
    Dim max_col As Double
    max_row = getMaxRow(1)
    max_col = getMaxCol(1)

    Workbooks.Open Filename:=file2, UpdateLinks:=0
    ActiveWorkbook.Sheets("BLP & WH Stock (LX02)").Activate
    Dim source As String
    source = path_file1 + "\[" + getFilenameFromPath(file1) + "]Sheet1!C1:C10"
    
    Dim pt_name As String
    file2 = Range("B5").Value
    For Each pt In ActiveSheet.PivotTables
        pt_name = pt.Name
    Next pt
    ActiveSheet.PivotTables(pt_name).ChangePivotCache ActiveWorkbook. _
        PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        source, Version:=6)
    With ActiveSheet.PivotTables(pt_name).PivotFields("Storage Type")
        .Orientation = xlColumnField
        .Position = 1
        .PivotItems("K12").Position = 1
        .PivotItems("O01").Position = 2
        .PivotItems("K24").Position = 3
        .PivotItems("R00").Position = 4
        .PivotItems("902").Position = 5
        .PivotItems("921").Position = 6
        .PivotItems("PD2").Position = 7
        .PivotItems("K61").Position = 8
        .PivotItems("R03").Position = 9
    End With

    ActiveWorkbook.Sheets("Update").Activate
    ActiveSheet.ShowAllData
    
    Range(Cells(8, 5), Cells(getMaxRow(1), 13)).Select
    Selection.ClearContents
    
    Dim j As Double
    j = 9
    For i = 5 To 13 Step 1
        Cells(8, i).FormulaR1C1 = "=VLOOKUP(RC[" + CStr(i - j) + "],'BLP & WH Stock (LX02)'!C[" + CStr(i - j) + "]:C[" + CStr(-3) + "]," + CStr(i - 3) + ",0)"
        j = j + 2
    Next
    
    Range("E8:M8").Select
    Selection.AutoFill Destination:=Range(Cells(8, 5), Cells(getMaxRow(1), 13))

    ActiveSheet.Range("$A$7:$AR$" + CStr(getMaxRow(1))).AutoFilter Field:=5, Criteria1:="#N/A"
    
    Range("E8").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("E8").Select
    Selection.AutoFill Destination:=Range("E8:M8"), Type:=xlFillDefault
    Range("E8:M8").Select
    Selection.Copy
    Range(Selection, Selection.End(xlDown)).Select
    ActiveSheet.Paste
End Sub
