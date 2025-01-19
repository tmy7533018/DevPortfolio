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
                '����������
                For i = topmostRow + 1 To lastRow
                    If params.CustomerListSheet.Cells(i, billMonthCol).Value = singleYearMonth Then
                        ReDim Preserve customerDataAry(0 To cnt)
                        customerDataAry(cnt) = params.CustomerListSheet.Range(params.CustomerListSheet.Cells(i, leftmostCol), params.CustomerListSheet.Cells(i, lastCol)).Value
                        cnt = cnt + 1
                    End If
                Next i
                
            Else
                '�������Ɩ��O
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
            lastYearMonth = params.LastYear & "/" & params.LastMonth & "/01"   'Date�^�Ƃ��Ĉ������߂ɉ��̓��t��ǉ�
            
            If params.customerName = "" Then
                '�����͈͂���
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
                '�����͈͂Ɩ��O
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
    
    ' �z������[�v���Čڋq���ƂɃf�[�^�𕪂���
    cnt = 0
    For Each rowData In customerDataAry
    
        ' rowData�̃C���f�b�N�X
        ' (1, 1): No
        ' (1, 2): ���O
        ' (1, 3): �Z��
        ' (1, 4): ���i��
        ' (1, 5): ����
        ' (1, 6): ���i
        ' (1, 7): ������
        ' (1, 8): �o�^��
        
        customerName = rowData(1, 2) ' �ڋq���̗���擾
    
        ' �ڋq����Dictionary�ɑ��݂��Ȃ��ꍇ�A�V�����G���g�����쐬
        If Not dict.Exists(customerName) Then
            dict.Add customerName, Array() ' ��̔z���������
            
            ReDim Preserve dictKeys(0 To cnt)
            dictKeys(cnt) = customerName  '�L�[���i�[���Ă���
            cnt = cnt + 1
        End If
        
        ' �����̌ڋq�f�[�^�ɐV�����s��ǉ�
        Dim tempArray As Variant
        tempArray = dict(customerName) ' ���݂̌ڋq�f�[�^���擾
        ReDim Preserve tempArray(LBound(tempArray) To UBound(tempArray) + 1) ' �z����g��
        tempArray(UBound(tempArray)) = rowData ' �V�����s�f�[�^��ǉ�
        dict(customerName) = tempArray ' �g�����ꂽ�z���߂�
    Next
    
    
    
    'dict(customerName)�ɁA�ڋq�����ƂɃf�[�^�𕪂���
    '������g���Ĉȉ��ɐ����������A���S���Y��������
    
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
    
    ' �V�����V�[�g���쐬
    Set ws = ThisWorkbook.Sheets.Add
    ws.name = "������_" & customerName
    ws.Activate
    ActiveWindow.Zoom = 85
    
    
    '�^�C�g����w�b�_�[�Ȃ�
    With ws
        .Range("A1:E2").Merge
        .Range("A1").HorizontalAlignment = xlCenter
        .Range("A1").Font.Size = 20
        .Range("A1").Value = "������"
        
        .Range("D3").Value = "�����ԍ�"
        .Range("D4").Value = "������"
        .Range("A5").Value = "����"
        .Range("A6").Value = "�Z��"
        
        .Range("A9:B9").Merge
        .Range("A9").HorizontalAlignment = xlCenter
        .Range("A9").Value = "���L�̒ʂ育�����\���グ�܂��B"
        
        .Range("A10:B10").Borders.LineStyle = xlContinuous
        .Range("A10").Value = "���������z"
        
        .Columns("A").ColumnWidth = 10
        .Columns("B").ColumnWidth = 22
        .Columns("C").ColumnWidth = 8
        .Columns("D").ColumnWidth = 13
        .Columns("E").ColumnWidth = 13
        .Range("A13").Value = "���t"
        .Range("B13").Value = "���i��"
        .Range("C13").Value = "����"
        .Range("D13").Value = "�P��"
        .Range("E13").Value = "���z"
        
        
        .Range("A34:E34").Borders(xlEdgeTop).LineStyle = xlContinuous
        .Range("A36:E36").Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Range("A34").Value = "���l�F"
    End With
    
    '�������ԍ�
    Dim timeStamp As String
    timeStamp = "I-" & Format(Now, "yyyyMMddHHmmss")
    ws.Range("E3").Value = timeStamp
    
    '������
    ws.Range("E4").Value = Format(customerData(0)(1, 7), "yyyy�NMM��")
    
    '����
    ws.Range("B5").Value = customerName
    
    '�Z��
    ws.Range("B6").Value = customerData(0)(1, 3)
    
    '���׃f�[�^
    row = 14
    totalAmount = 0
    For i = LBound(customerData) To UBound(customerData)
        ws.Cells(row, "A").Value = Format(customerData(i)(1, 8), "mm/dd")
        ws.Cells(row, "B").Value = customerData(i)(1, 4) '���i��
        ws.Cells(row, "C").Value = customerData(i)(1, 5) '����
        ws.Cells(row, "D").Value = customerData(i)(1, 6) '�P��
        ws.Cells(row, "E").Formula = "=" & ws.Cells(row, "C").address & "*" & ws.Cells(row, "D").address '���z
        totalAmount = totalAmount + ws.Cells(row, "E").Value
        row = row + 1
    Next i
    ws.Range(Cells(13, "A"), Cells(row - 1, "E")).HorizontalAlignment = xlCenter
    Call TableFormat(ws, 13, "A", row - 1, 5)
    
    ws.Cells(row + 1, "D").Value = "���v���z"
    
    With ws.Cells(row + 1, "E")
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlMedium
        .HorizontalAlignment = xlCenter
        .Value = totalAmount
    End With
    
    '�������z
    ws.Range("B10").HorizontalAlignment = xlCenter
    ws.Range("B10").Value = totalAmount & "�~"
    
End Sub

