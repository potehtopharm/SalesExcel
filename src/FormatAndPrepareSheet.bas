Attribute VB_Name = "Module1"
Sub FormatAndPrepareSheet()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' === Column resizing and wrapping ===
    ws.Columns("A").ColumnWidth = ws.Columns("A").ColumnWidth / 3
    ws.Columns("B").ColumnWidth = ws.Columns("B").ColumnWidth * 2
    ws.Columns("B").WrapText = True
    ws.Columns("H").ColumnWidth = ws.Columns("H").ColumnWidth / 2
    ws.Columns("K").ColumnWidth = ws.Columns("K").ColumnWidth / 3
    ws.Columns("S").ColumnWidth = ws.Columns("S").ColumnWidth * 2
    ws.Columns("S").WrapText = True
    ws.Columns("Y").ColumnWidth = ws.Columns("Y").ColumnWidth * 2
    ws.Columns("Y").WrapText = True
    ws.Columns("AA").ColumnWidth = ws.Columns("AA").ColumnWidth * 4
    ws.Columns("AA").WrapText = True

    ' === Delete specified columns ===
    Dim colsToDelete As Variant
    colsToDelete = Array("AB", "Z", "X", "W", "V", "Q", "O", "N", "M", "L", "J", "I", "G", "F", "E", "D")
    Dim col As Variant
    For Each col In colsToDelete
        ws.Columns(col).Delete
    Next col

    ' === Add horizontal borders A to L, every row with data ===
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    Dim rng As Range
    Set rng = ws.Range("A1:L" & lastRow)
    
    With rng.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = vbBlack
    End With
    With rng.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = vbBlack
    End With

    ' === Clear specific header cells ===
    ws.Range("A1:G1").ClearContents
    ws.Range("J1").ClearContents

    ' === Overwrite specific header values ===
    ws.Range("B2").Value = "Name"
    ws.Range("C2").Value = "HH"
    ws.Range("D2").Value = "RM#"
    ws.Range("E2").Value = "#NTS"
    ws.Range("F2").Value = "RATE"
    ws.Range("G2").Value = "CODE"
    ws.Range("H2").Value = "Company"
    ws.Range("I2").Value = "CONF"
    ws.Range("J2").Value = "RMtype"
    ws.Range("K2").Value = "disc"
    ws.Range("L2").Value = "comments"
    ws.Range("B3").Value = "ERROR (ONQ report glitch)"

    ' === Bold Row 2 ===
    ws.Rows(2).Font.Bold = True

    MsgBox "Formatting complete!", vbInformation
End Sub

