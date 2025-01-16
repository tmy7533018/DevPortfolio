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
                '����������
                For i = topmostRow + 1 To lastRow
                    If params.CustomerListSheet.Cells(i, billMonthCol).Value = singleYearMonth Then
                        ReDim customerDataAry(0 To cnt)
                        customerDataAry(cnt) = params.CustomerListSheet.Range(i, lastCol).Value
                        cnt = cnt + 1
                    End If
                Next i
                
            Else
                '�������Ɩ��O
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
            lastYearMonth = params.LastYear & "/" & params.LastMonth & "/01"   'Date�^�Ƃ��Ĉ������߂ɉ��̓��t��ǉ�
            
            If params.customerName = "" Then
                '�����͈͂���
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
                '�����͈͂Ɩ��O
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
    
    ' �z������[�v���Čڋq���ƂɃf�[�^�𕪂���
    For Each rowData In customerDataAry
        customerName = rowData(2) ' �ڋq���̗���擾�i�z���2��ځj
    
        ' �ڋq����Dictionary�ɑ��݂��Ȃ��ꍇ�A�V�����G���g�����쐬
        If Not dict.Exists(customerName) Then
            dict.Add customerName, Array() ' ��̔z���������
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
    
End Function

