Attribute VB_Name = "TableFormatting"
Option Explicit

Sub TableFormat(ws As Worksheet, topmostRow_ As Integer, leftmostCol_ As String, lastRow As Long, lastCol As Long)
    
    With ws.Range(ws.Cells(topmostRow_, leftmostCol_), ws.Cells(lastRow, lastCol))
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
    
End Sub
