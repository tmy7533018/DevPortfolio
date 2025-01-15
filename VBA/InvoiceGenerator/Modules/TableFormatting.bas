Attribute VB_Name = "TableFormatting"
Option Explicit

Sub TableFormat(ws As Worksheet, topmostRow As Integer, leftmostCol As String, lastRow As Long, lastCol As Long)
    
    With ws.Range(ws.Cells(topmostRow, leftmostCol), ws.Cells(lastRow, lastCol))
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
    
End Sub
