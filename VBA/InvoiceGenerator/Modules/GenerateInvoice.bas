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


Function InvoiceGenerator(params As BillingParams, flag As String) As Boolean
    
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
    
    On Error GoTo errorhandler
    
    lastRow = params.CustomerListSheet.Cells(params.CustomerListSheet.Rows.Count, leftmostCol).End(xlUp).row
    lastCol = params.CustomerListSheet.Cells(topmostRow, params.CustomerListSheet.Cells.Columns.Count).End(xlToLeft).Column
    
    cnt = 0
    Select Case flag
        
        Case "PageSingle"
        
            singleYearMonth = params.SingleYear & "/" & Format(params.SingleMonth, "00")
            
            If params.customerName = "" Then
                '請求日だけ
                For i = topmostRow + 1 To lastRow
                    If params.CustomerListSheet.Cells(i, billMonthCol).Value = singleYearMonth Then
                        ReDim Preserve customerDataAry(0 To cnt)
                        customerDataAry(cnt) = params.CustomerListSheet.Range(params.CustomerListSheet.Cells(i, leftmostCol), params.CustomerListSheet.Cells(i, lastCol)).Value
                        cnt = cnt + 1
                    End If
                Next i
                
            Else
                '請求日と名前
                For i = topmostRow + 1 To lastRow
                    isMach = params.CustomerListSheet.Cells(i, billMonthCol).Value = singleYearMonth _
                            And params.CustomerListSheet.Cells(i, nameCol).Value = params.customerName
                            
                    If isMach Then
                        ReDim Preserve customerDataAry(0 To cnt)
                        customerDataAry(cnt) = params.CustomerListSheet.Range(params.CustomerListSheet.Cells(i, leftmostCol), params.CustomerListSheet.Cells(i, lastCol)).Value
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
                            And CDate(params.CustomerListSheet.Cells(i, billMonthCol)) <= lastYearMonth
                            
                    If isMach Then
                        ReDim Preserve customerDataAry(0 To cnt)
                        customerDataAry(cnt) = params.CustomerListSheet.Range(params.CustomerListSheet.Cells(i, leftmostCol), params.CustomerListSheet.Cells(i, lastCol)).Value
                        cnt = cnt + 1
                    End If
                Next i
                
            Else
                '請求範囲と名前
                For i = topmostRow + 1 To lastRow
                    isMach = CDate(params.CustomerListSheet.Cells(i, billMonthCol)) >= startYearMonth _
                            And CDate(params.CustomerListSheet.Cells(i, billMonthCol)) <= lastYearMonth _
                            And params.CustomerListSheet.Cells(i, nameCol) = params.customerName
                            
                    If isMach Then
                        ReDim Preserve customerDataAry(0 To cnt)
                        customerDataAry(cnt) = params.CustomerListSheet.Range(params.CustomerListSheet.Cells(i, leftmostCol), params.CustomerListSheet.Cells(i, lastCol)).Value
                        cnt = cnt + 1
                    End If
                Next i
                
            End If
            
    End Select
    
    
    
    Dim rowData As Variant
    Dim customerName As String
    Dim dict As Object
    Dim dictKeys() As String
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' 配列をループして顧客ごとにデータを分ける
    cnt = 0
    For Each rowData In customerDataAry
    
        ' rowDataのインデックス
        ' (1, 1): No
        ' (1, 2): 名前
        ' (1, 3): 住所
        ' (1, 4): 商品名
        ' (1, 5): 数量
        ' (1, 6): 価格
        ' (1, 7): 請求月
        ' (1, 8): 登録日
        
        customerName = rowData(1, 2) ' 顧客名の列を取得
    
        ' 顧客名がDictionaryに存在しない場合、新しいエントリを作成
        If Not dict.Exists(customerName) Then
            dict.Add customerName, Array() ' 空の配列を初期化
            
            ReDim Preserve dictKeys(0 To cnt)
            dictKeys(cnt) = customerName  'キーを格納しておく
            cnt = cnt + 1
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
    
    Dim key As Variant
    Dim Data() As Variant
    
    For Each key In dictKeys
        
        Data = dict(key)
        Call BuildCustomerInvoiceSheet(CStr(key), Data)
        
    Next
    
    InvoiceGenerator = True
    Exit Function
    
errorhandler:
    Dim isInvoiceGenerated As Boolean
    isInvoiceGenerated = False
    
End Function


Sub BuildCustomerInvoiceSheet(customerName As String, customerData As Variant)

    Dim ws As Worksheet
    Dim i As Long, row As Long
    Dim totalAmount As Double
    
    ' 新しいシートを作成
    Set ws = ThisWorkbook.Sheets.Add
    ws.name = "請求書_" & customerName
    ws.Activate
    ActiveWindow.Zoom = 85
    
    
    'タイトルやヘッダーなど
    With ws
        .Range("A1:E2").Merge
        .Range("A1").HorizontalAlignment = xlCenter
        .Range("A1").Font.Size = 20
        .Range("A1").Value = "請求書"
        
        .Range("D3").Value = "請求番号"
        .Range("D4").Value = "請求日"
        .Range("A5").Value = "氏名"
        .Range("A6").Value = "住所"
        
        .Range("A9:B9").Merge
        .Range("A9").HorizontalAlignment = xlCenter
        .Range("A9").Value = "下記の通りご請求申し上げます。"
        
        .Range("A10:B10").Borders.LineStyle = xlContinuous
        .Range("A10").Value = "ご請求金額"
        
        .Columns("A").ColumnWidth = 10
        .Columns("B").ColumnWidth = 22
        .Columns("C").ColumnWidth = 8
        .Columns("D").ColumnWidth = 13
        .Columns("E").ColumnWidth = 13
        .Range("A13").Value = "日付"
        .Range("B13").Value = "商品名"
        .Range("C13").Value = "数量"
        .Range("D13").Value = "単価"
        .Range("E13").Value = "金額"
        
        
        .Range("A34:E34").Borders(xlEdgeTop).LineStyle = xlContinuous
        .Range("A36:E36").Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Range("A34").Value = "備考："
    End With
    
    '請求書番号
    Dim timeStamp As String
    timeStamp = "I-" & Format(Now, "yyyyMMddHHmmss")
    ws.Range("E3").Value = timeStamp
    
    '請求日
    ws.Range("E4").Value = Format(customerData(0)(1, 7), "yyyy年MM月")
    
    '氏名
    ws.Range("B5").Value = customerName
    
    '住所
    ws.Range("B6").Value = customerData(0)(1, 3)
    
    '明細データ
    row = 14
    totalAmount = 0
    For i = LBound(customerData) To UBound(customerData)
        ws.Cells(row, "A").Value = Format(customerData(i)(1, 8), "mm/dd")
        ws.Cells(row, "B").Value = customerData(i)(1, 4) '商品名
        ws.Cells(row, "C").Value = customerData(i)(1, 5) '数量
        ws.Cells(row, "D").Value = customerData(i)(1, 6) '単価
        ws.Cells(row, "E").Formula = "=" & ws.Cells(row, "C").address & "*" & ws.Cells(row, "D").address '金額
        totalAmount = totalAmount + ws.Cells(row, "E").Value
        row = row + 1
    Next i
    ws.Range(Cells(13, "A"), Cells(row - 1, "E")).HorizontalAlignment = xlCenter
    Call TableFormat(ws, 13, "A", row - 1, 5)
    
    ws.Cells(row + 1, "D").Value = "合計金額"
    
    With ws.Cells(row + 1, "E")
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlMedium
        .HorizontalAlignment = xlCenter
        .Value = totalAmount
    End With
    
    '請求金額
    ws.Range("B10").HorizontalAlignment = xlCenter
    ws.Range("B10").Value = totalAmount & "円"
    
End Sub

