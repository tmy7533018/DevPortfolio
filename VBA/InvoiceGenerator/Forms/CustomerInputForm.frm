VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CustomerInputForm 
   Caption         =   "�ڋq�����̓t�H�[��"
   ClientHeight    =   8730.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15735
   OleObjectBlob   =   "CustomerInputForm.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "CustomerInputForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub CommandButtonConfirm_Click()
    
    Dim CustomerListSheet As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim targetRow As Long
    Dim name As String
    Dim address As String
    Dim item As String
    Dim quantity As Long
    Dim price As Long
    Dim regDate As Date
    Dim billMonth As String
    Dim prompt As String
    Dim result As VbMsgBoxResult
    
    
    Set CustomerListSheet = ThisWorkbook.Sheets("�ڋq���")
    
    lastRow = CustomerListSheet.Cells(CustomerListSheet.Rows.Count, leftmostCol).End(xlUp).Row
    lastCol = CustomerListSheet.Cells(topmostRow, CustomerListSheet.Cells.Columns.Count).End(xlToLeft).Column
    targetRow = lastRow + 1
    
    regDate = Date
    
    prompt = ""
    prompt = prompt & ErrorCheck(LabelName.caption, TextBoxName.text, 1)
    prompt = prompt & ErrorCheck(LabelAddress.caption, TextBoxAddress.text, 1)
    prompt = prompt & ErrorCheck(LabelItem.caption, TextBoxItem.text, 1)
    prompt = prompt & ErrorCheck(LabelQuantity.caption, TextBoxQuantity.text, 2)
    prompt = prompt & ErrorCheck(LabelPrice.caption, TextBoxPrice.text, 2)

    If Not prompt = "" Then
    
        Beep
        LabelMessage.ForeColor = RGB(255, 0, 0)
        LabelMessage.caption = prompt
        
        Exit Sub
    
    End If
    
    
    name = TextBoxName.text
    address = TextBoxAddress.text
    item = TextBoxItem.text
    quantity = TextBoxQuantity.text
    price = TextBoxPrice.text
    
    result = MsgBox("�o�^���܂����H", vbOKCancel, "�m�F���")
    If result = 2 Then Exit Sub
    
    If CustomerListSheet.Cells(lastRow, numberCol).Value = "No" Then
        CustomerListSheet.Cells(targetRow, numberCol).Value = 1
        
    Else
        CustomerListSheet.Cells(targetRow, numberCol).Value = CustomerListSheet.Cells(lastRow, numberCol).Value + 1
        
    End If
    
    
    CustomerListSheet.Cells(targetRow, nameCol).Value = name
    CustomerListSheet.Cells(targetRow, addressCol).Value = address
    CustomerListSheet.Cells(targetRow, itemCol).Value = item
    CustomerListSheet.Cells(targetRow, quantityCol).Value = quantity
    CustomerListSheet.Cells(targetRow, priceCol).Value = price
    CustomerListSheet.Cells(targetRow, regDateCol).Value = regDate
    
    billMonth = Format(DateAdd("m", 1, regDate), "yyyy/mm")
    CustomerListSheet.Cells(targetRow, billMonthCol).Value = billMonth
    
    lastRow = CustomerListSheet.Cells(CustomerListSheet.Rows.Count, leftmostCol).End(xlUp).Row
    lastCol = CustomerListSheet.Cells(topmostRow, CustomerListSheet.Cells.Columns.Count).End(xlToLeft).Column
    Call TableFormat(CustomerListSheet, topmostRow, leftmostCol, lastRow, lastCol)
    
    
    LabelMessage.ForeColor = RGB(0, 0, 0)
    LabelMessage.caption = "�o�^���������܂����B" & vbCrLf & vbCrLf & "�����ēo�^��Ƃ��ł��܂��B"
    Me.TextBoxName.SetFocus
    TextBoxName.text = ""
    TextBoxAddress.text = ""
    TextBoxItem.text = ""
    TextBoxQuantity.text = ""
    TextBoxPrice.text = ""
    
End Sub


Private Sub CommandButtonExit_Click()

    Dim result As VbMsgBoxResult
    
    result = MsgBox("�I�����܂����H", vbOKCancel, "�I�����")
    If result = 1 Then Unload Me
    
    Exit Sub
    
End Sub


Private Sub UserForm_Initialize()
    
    Me.TextBoxName.SetFocus
    LabelRegDate.caption = Date
    
End Sub


Private Function ErrorCheck(labelDetails As String, userInput As String, checkMode As Integer) As String
    
    Dim prompt As String
    prompt = ""
    
    Select Case checkMode
        
        Case 1:
        
            If userInput = "" Then
                prompt = "�u" & labelDetails & "�v" & " �����͂���Ă��܂���B" & vbCrLf & vbCrLf
                
            Else
                Exit Function
                
            End If
            
            ErrorCheck = prompt
        
        Case 2:
        
            If userInput = "" Then
                prompt = "�u" & labelDetails & "�v" & " �����͂���Ă��܂���B" & vbCrLf & vbCrLf
                
            ElseIf Not IsNumeric(userInput) Then
                prompt = "�u" & labelDetails & "�v" & " �ɂ͐��l����͂��Ă��������B" & vbCrLf & vbCrLf
            
            ElseIf CLng(userInput) <> CDbl(userInput) Then
                prompt = "�u" & labelDetails & "�v" & " �ɂ͐�������͂��Ă��������B" & vbCrLf & vbCrLf
                
            ElseIf CLng(userInput) < 0 Then
                prompt = "�u" & labelDetails & "�v" & " �ɕ��̒l�͖����ł��B" & vbCrLf & vbCrLf
            
            Else
                Exit Function
            
            End If
            
            ErrorCheck = prompt
            
    End Select
    
End Function


Private Sub UserForm_Click()

End Sub
