Attribute VB_Name = "GenerateInvoice"
Option Explicit


Public Type BillingParams
    
    CustomerName As String
    SingleYear As String
    SingleMonth As String
    StartYear As String
    StartMonth As String
    LastYear As String
    LastMonth As String
    CustomerListSheet As Worksheet
        
End Type


Function GenerateInvoiceSheet(params As BillingParams, flag As String)
    
    Dim lastRow As Long
    Dim lastCol As Long
    Dim newInvoiceSheet As Worksheet
    Dim singleYearMonth As String
    Dim startYearMonth As String
    Dim lastYearMonth As String
    Dim customerDataAry() As Variant
    
    lastRow = params.CustomerListSheet.Cells(params.CustomerListSheet.Rows.Count, leftmostCol).End(xlUp).Row
    lastCol = params.CustomerListSheet.Cells(topmostRow, params.CustomerListSheet.Cells.Columns.Count).End(xlToLeft).Column

    
    Dim i As Long
    Dim cnt As Long
    cnt = 0
    Select Case flag
        
        Case "PageSingle"
        
            singleYearMonth = params.SingleYear & "/" & Format(params.SingleMonth, "00")
            
            If params.CustomerName = "" Then
                '¿‹“ú‚¾‚¯
                For i = topmostRow + 1 To lastRow
                    If params.CustomerListSheet.Cells(i, billMonthCol).Value = singleYearMonth Then
                        ReDim customerDataAry(0 To cnt)
                        customerDataAry(cnt) = params.CustomerListSheet.Range(i, lastCol).Value
                        cnt = cnt + 1
                    End If
                Next i
                
            Else
                '¿‹“ú‚Æ–¼‘O
                For i = topmostRow + 1 To lastRow
                    If params.CustomerListSheet.Cells(i, billMonthCol).Value = singleYearMonth & params.CustomerListSheet.Cells(i, nameCol).Value = params.CustomerName Then
                        ReDim customerDataAry(0 To cnt)
                        customerDataAry(cnt) = params.CustomerListSheet.Range(i, lastCol).Value
                        cnt = cnt + 1
                    End If
                Next i
                
            End If
            
        Case "PageRange"
            
            startYearMonth = params.StartYear & "/" & params.StartMonth & "/01"
            lastYearMonth = params.LastYear & "/" & params.LastMonth & "/01"   'DateŒ^‚Æ‚µ‚Äˆµ‚¤‚½‚ß‚É‰¼‚Ì“ú•t‚ð’Ç‰Á
            
            If params.CustomerName = "" Then
                '¿‹”ÍˆÍ‚¾‚¯
                For i = topmostRow + 1 To lastRow
                    If CDate(params.CustomerListSheet.Cells(i, billMonthCol)) >= startYearMonth & CDate(params.CustomerListSheet.Cells(i, billMonthCol)) <= lastYearMonth Then
                        ReDim customerDataAry(0 To cnt)
                        customerDataAry(cnt) = params.CustomerListSheet.Range(i, lastCol).Value
                        cnt = cnt + 1
                    End If
                Next i
                
            Else
                '¿‹”ÍˆÍ‚Æ–¼‘O
                For i = topmostRow + 1 To lastRow
                    If CDate(params.CustomerListSheet.Cells(i, billMonthCol)) >= startYearMonth & CDate(params.CustomerListSheet.Cells(i, billMonthCol)) <= lastYearMonth & params.CustomerListSheet.Cells(i, nameCol) = CustomerName Then
                        ReDim customerDataAry(0 To cnt)
                        customerDataAry(cnt) = params.CustomerListSheet.Range(i, lastCol).Value
                        cnt = cnt + 1
                    End If
                Next i
                
            End If
            
    End Select
    
End Function

