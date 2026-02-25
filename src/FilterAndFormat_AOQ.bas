Attribute VB_Name = "Module2"
Sub FilterAndFormatCompanyNames()
    Dim ws As Worksheet, wsFiltered As Worksheet
    Dim lastRow As Long
    Dim i As Long, k
    Dim keywords As Variant
    Dim cellValue As String
    Dim matchCol As Long
    Dim filterRange As Range
    Dim newSheetName As String
    
    ' Set original worksheet
    On Error GoTo SheetError
    Set ws = ThisWorkbook.Sheets("ARRIVALLLANDSCAPE_LETTER.RPT")
    On Error GoTo 0

    ' Get last row of data
    lastRow = ws.Cells(ws.Rows.Count, "S").End(xlUp).Row

    ' Keywords (prefix match)
    keywords = Array("CQB", "Hotembeds", "Coajiang", "KLTAO")

    ' === SAFE FIX 1: clear old MatchHelper column values (prevents stale MATCH) ===
    ws.Range("Z1").Value = "MatchHelper"
    ws.Range("Z2:Z" & lastRow).ClearContents

    For i = 2 To lastRow
        If Not IsError(ws.Cells(i, "S").Value) Then
            
            ' === SAFE FIX 2: Trim spaces to prevent missed matches ===
            cellValue = Trim(CStr(ws.Cells(i, "S").Value))
            
            ' === SAFE FIX 3 (optional but safe): normalize double spaces ===
            Do While InStr(cellValue, "  ") > 0
                cellValue = Replace(cellValue, "  ", " ")
            Loop
            
            For Each k In keywords
                ' Match only if the cell starts with the keyword (case-insensitive)
                If LCase(Left(cellValue, Len(k))) = LCase(k) Then
                    ws.Cells(i, "Z").Value = "MATCH"
                    Exit For
                End If
            Next k
        End If
    Next i

    ' === SAFE FIX 4: avoid ShowAllData error ===
    On Error Resume Next
    If ws.AutoFilterMode Then ws.AutoFilter.ShowAllData
    On Error GoTo 0

    ' Apply filter
    Set filterRange = ws.Range("A1:Z" & lastRow)
    filterRange.AutoFilter Field:=26, Criteria1:="MATCH"

    ' Copy visible rows to new worksheet
    newSheetName = "Filtered_Results"
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets(newSheetName).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set wsFiltered = ThisWorkbook.Sheets.Add(After:=ws)
    wsFiltered.Name = newSheetName

    ' Copy visible filtered rows
    On Error Resume Next
    filterRange.SpecialCells(xlCellTypeVisible).Copy Destination:=wsFiltered.Range("A1")
    On Error GoTo 0

    ' Turn off autofilter in original sheet
    ws.AutoFilterMode = False

    ' === Clean & format in new sheet only ===
    With wsFiltered
        Dim col As Variant
        Dim colsToDelete As Variant
        colsToDelete = Array("AB", "AA", "Y", "X", "W", "V", "Q", "O", "N", "M", "L", "J", "I", "G", "F", "E", "D")
        For Each col In colsToDelete
            On Error Resume Next
            .Columns(col).Delete
            On Error GoTo 0
        Next col

        ' Delete MatchHelper column
        Dim cell As Range
        matchCol = 0
        For Each cell In .Rows(1).Cells
            If Not IsError(cell.Value) And Trim(CStr(cell.Value)) = "MatchHelper" Then
                matchCol = cell.Column
                Exit For
            End If
        Next cell
        If matchCol > 0 Then .Columns(matchCol).Delete

        ' Final adjustments
        .Columns("G").Cut
        .Columns("K").Insert Shift:=xlToRight
        Application.CutCopyMode = False
        .Columns("A").Delete
    End With

    MsgBox "Filtered data copied to 'Filtered_Results' and formatted.", vbInformation
    Exit Sub

SheetError:
    MsgBox "Sheet 'ARRIVALLLANDSCAPE_LETTER.RPT' not found. Please check the name again.", vbCritical
End Sub
