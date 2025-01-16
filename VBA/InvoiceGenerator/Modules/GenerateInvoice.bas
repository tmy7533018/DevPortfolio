Attribute VB_Name = "GenerateInvoice"
Option Explicit


Public Type BillingParams
    
    customerName As String
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
    Dim isMach As Boolean
    Dim i As Long
    Dim cnt As Long
    
    lastRow = params.CustomerListSheet.Cells(params.CustomerListSheet.Rows.Count, leftmostCol).End(xlUp).Row
    lastCol = params.CustomerListSheet.Cells(topmostRow, params.CustomerListSheet.Cells.Columns.Count).End(xlToLeft).Column
    

    cnt = 0
    Select Case flag
        
        Case "PageSingle"
        
            singleYearMonth = params.SingleYear & "/" & Format(params.SingleMonth, "00")
            
            If params.customerName = "" Then
                '請求日だけ
                For i = topmostRow + 1 To lastRow
                    If params.CustomerListSheet.Cells(i, billMonthCol).Value = singleYearMonth Then
                        ReDim customerDataAry(0 To cnt)
                        customerDataAry(cnt) = params.CustomerListSheet.Range(i, lastCol).Value
                        cnt = cnt + 1
                    End If
                Next i
                
            Else
                '請求日と名前
                For i = topmostRow + 1 To lastRow
                    isMach = params.CustomerListSheet.Cells(i, billMonthCol).Value = singleYearMonth _
                            & params.CustomerListSheet.Cells(i, nameCol).Value = params.customerName
                            
                    If isMach Then
                        ReDim customerDataAry(0 To cnt)
                        customerDataAry(cnt) = params.CustomerListSheet.Range(i, lastCol).Value
                        cnt = cnt + 1
                    End If
                Next i
                
            End If
            
        Case "PageRange"
            
            startYearMonth = params.StartYear & "/" & params.StartMonth & "/01"
            lastYearMonth = params.LastYear & "/" & params.LastMonth & "/01"   'Date型として扱うために仮の日付を追加
            
            If params.customerName = "" Then
                '請求範囲だけ
                For i = topmostRow + 1 To lastRow
                    isMach = CDate(params.CustomerListSheet.Cells(i, billMonthCol)) >= startYearMonth _
                            & CDate(params.CustomerListSheet.Cells(i, billMonthCol)) <= lastYearMonth
                            
                    If isMach Then
                        ReDim customerDataAry(0 To cnt)
                        customerDataAry(cnt) = params.CustomerListSheet.Range(i, lastCol).Value
                        cnt = cnt + 1
                    End If
                Next i
                
            Else
                '請求範囲と名前
                For i = topmostRow + 1 To lastRow
                    isMach = CDate(params.CustomerListSheet.Cells(i, billMonthCol)) >= startYearMonth _
                            & CDate(params.CustomerListSheet.Cells(i, billMonthCol)) <= lastYearMonth _
                            & params.CustomerListSheet.Cells(i, nameCol) = customerName
                            
                    If isMach Then
                        ReDim customerDataAry(0 To cnt)
                        customerDataAry(cnt) = params.CustomerListSheet.Range(i, lastCol).Value
                        cnt = cnt + 1
                    End If
                Next i
                
            End If
            
    End Select
    
    
    Dim rowData As Variant
    Dim customerName As String
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' 配列をループして顧客ごとにデータを分ける
    For Each rowData In customerDataAry
        customerName = rowData(2) ' 顧客名の列を取得（配列の2列目）
    
        ' 顧客名がDictionaryに存在しない場合、新しいエントリを作成
        If Not dict.Exists(customerName) Then
            dict.Add customerName, Array() ' 空の配列を初期化
        End If
    
        ' 既存の顧客データに新しい行を追加
        Dim tempArray As Variant
        tempArray = dict(customerName) ' 現在の顧客データを取得
        ReDim Preserve tempArray(LBound(tempArray) To UBound(tempArray) + 1) ' 配列を拡張
        tempArray(UBound(tempArray)) = rowData ' 新しい行データを追加
        dict(customerName) = tempArray ' 拡張された配列を戻す
    Next
    
    
    'dict(customerName)に、顧客名ごとにデータを分けた
    'それを使って以下に請求書生成アルゴリズムを書く
    
End Function

