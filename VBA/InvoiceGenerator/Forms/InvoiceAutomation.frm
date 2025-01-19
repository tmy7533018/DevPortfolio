VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} InvoiceAutomation 
   Caption         =   "UserForm1"
   ClientHeight    =   9105.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7920
   OleObjectBlob   =   "InvoiceAutomation.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "InvoiceAutomation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub CommandButtonGenerate_Click()
    
    Dim CustomerListSheet As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim selectedPageName As String
    Dim SingleYear As String
    Dim SingleMonth As String
    Dim rangeStartYear As String
    Dim rangeStartMonth As String
    Dim rangeLastYear As String
    Dim rangeLastMonth As String
    Dim prompt As String
    Dim GenerateInvoiceFlag As String
    Dim invoiceSheet As Worksheet
    Dim result As VbMsgBoxResult
    Dim params As BillingParams
    
    
    With params
        .customerName = ""
        .SingleYear = ""
        .SingleMonth = ""
        .StartYear = ""
        .StartMonth = ""
        .LastYear = ""
        .LastMonth = ""
        Set .CustomerListSheet = Nothing
    End With
    
    GenerateInvoiceFlag = MultiPageBillingOptions.Pages(MultiPageBillingOptions.Value).name
    
        
    Select Case GenerateInvoiceFlag
        Case "PageSingle"
            GenerateInvoiceFlag = "PageSingle"
            prompt = ErrorCheck_YearMonth(LabelSingle.caption, TextBoxSingleYear.text, TextBoxSingleMonth.text)
            If CheckBox1_Name.Value = True Then prompt = prompt & ErrorCheck_Str(CheckBox1_Name.caption, TextBox1_Name.text)
            
            If Not prompt = "" Then
                Call ErrorNotification(prompt)
                Exit Sub
            End If
            
            With params
                .customerName = TextBox1_Name.text
                .SingleYear = TextBoxSingleYear.text
                .SingleMonth = TextBoxSingleMonth.text
                Set .CustomerListSheet = ThisWorkbook.Worksheets("�ڋq���")
            End With
                
        Case "PageRange"
            GenerateInvoiceFlag = "PageRange"
            prompt = ErrorCheck_YearMonth(LabelRangeStart.caption, TextBoxStartYear.text, TextBoxStartMonth.text)
            prompt = prompt & ErrorCheck_YearMonth(LabelRangeLast.caption, TextBoxLastYear.text, TextBoxLastMonth.text)
            If CheckBox2_Name.Value = True Then prompt = prompt & ErrorCheck_Str(CheckBox2_Name.caption, TextBox2_Name.text)
            
            If Not prompt = "" Then
                Call ErrorNotification(prompt)
                Exit Sub
            End If
            
            With params
                .customerName = TextBox2_Name.text
                .StartYear = TextBoxStartYear.text
                .StartMonth = TextBoxStartMonth.text
                .LastYear = TextBoxLastYear.text
                .LastMonth = TextBoxLastMonth.text
                Set .CustomerListSheet = ThisWorkbook.Worksheets("�ڋq���")
            End With
            
    End Select
    
    
    result = MsgBox("�������𐶐����܂����H", vbOKCancel, "�m�F���")
    If result = 2 Then Exit Sub
    
    '�����������֐��Ăяo��
    Dim isInvoiceGenerated As Boolean
    isInvoiceGenerated = InvoiceGenerator(params, GenerateInvoiceFlag)
    
    If Not isInvoiceGenerated Then
        Beep
        MsgBox "�������������ł��܂���ł����B"
        LabelMessage.caption = "�������������ł��܂���ł����B" & vbCrLf & vbCrLf & "�������l����͂��Ă��������B"
        LabelMessage.ForeColor = RGB(255, 0, 0)
        Exit Sub
    End If
    
    MsgBox "�������𐶐����܂����B"
    LabelMessage.ForeColor = RGB(0, 0, 0)
    LabelMessage.caption = "�������𐶐����܂����B" & vbCrLf & vbCrLf & "�����Đ������ł��܂��B"
    TextBoxSingleYear.text = ""
    TextBoxSingleMonth.text = ""
    TextBox1_Name.text = ""
    TextBoxStartYear.text = ""
    TextBoxStartMonth.text = ""
    TextBoxLastYear.text = ""
    TextBoxLastMonth.text = ""
    TextBox2_Name.text = ""
    
End Sub


Private Sub CommandButtonExit_Click()
    
    Dim result As VbMsgBoxResult
    
    result = MsgBox("�I�����܂����H", vbOKCancel, "�I�����")
    If result = 1 Then Unload Me
    
    Exit Sub
    
End Sub


Private Sub MultiPageBillingOptions_Change()
    
    Dim selectedPageName As String
    
    selectedPageName = MultiPageBillingOptions.Pages(MultiPageBillingOptions.Value).name
    LabelMessage.ForeColor = RGB(0, 0, 0)
    
    If selectedPageName = "PageSingle" Then
        Me.TextBoxSingleYear.SetFocus
        LabelMessage.caption = "�����N������͂��Ă��������B" & vbCrLf & vbCrLf & "�ڋq���w�肷��ꍇ�̓`�F�b�N�{�b�N�X�Ƀ`�F�b�N�����Ă��������B"
        
    ElseIf selectedPageName = "PageRange" Then
        Me.TextBoxStartYear.SetFocus
        LabelMessage.caption = "�J�n�N���E�I���N������͂��Ă��������B" & vbCrLf & vbCrLf & "�ڋq���w�肷��ꍇ�̓`�F�b�N�{�b�N�X�Ƀ`�F�b�N�����Ă��������B"
        
    End If
    
End Sub


Private Sub UserForm_Initialize()
    
    Me.TextBoxSingleYear.SetFocus
    
End Sub


Private Sub UserForm_Click()

    LabelMessage.ForeColor = RGB(0, 0, 0)
    LabelMessage.caption = "�P���w��܂��͔͈͎w��Ő������𐶐����܂��B" & vbCrLf & vbCrLf & "�ڋq���w�肷��ꍇ�̓`�F�b�N�{�b�N�X�Ƀ`�F�b�N�����Ă��������B"
    
End Sub


Private Function ErrorCheck_YearMonth(labelText As String, userInputYear As String, userInputMonth As String) As String
    
    Dim prompt As String
    prompt = ""
    
    If userInputYear = "" Then
        prompt = labelText & "�N�����͂���Ă��܂���B" & vbCrLf & vbCrLf
        
    ElseIf Not IsNumeric(userInputYear) Or Len(userInputYear) <> 4 Then
        prompt = labelText & "�N��4���̐��l�œ��͂��Ă��������B" & vbCrLf & vbCrLf
        
    ElseIf CInt(userInputYear) <> CDbl(userInputYear) Or Val(userInputYear) < 0 Then
        prompt = labelText & "�N�ɖ����Ȓl���܂܂�Ă��܂��B" & vbCrLf & vbCrLf
    
    End If
    
    If userInputMonth = "" Then
        prompt = prompt & labelText & "�������͂���Ă��܂���B" & vbCrLf & vbCrLf
    
    ElseIf Not IsNumeric(userInputMonth) Or Val(userInputMonth) < 1 Or Val(userInputMonth) > 12 Then
        prompt = prompt & labelText & "����1�`12�̐��l�œ��͂��Ă��������B" & vbCrLf & vbCrLf
        
    ElseIf CInt(userInputMonth) <> CDbl(userInputMonth) Then
        prompt = prompt & labelText & "���ɖ����Ȓl���܂܂�Ă��܂��B" & vbCrLf & vbCrLf
        
    End If
    
    ErrorCheck_YearMonth = prompt
    
End Function


Private Function ErrorCheck_Str(caption As String, text As String) As String
    
    Dim prompt As String
    prompt = ""
    
    If text = "" Then
        prompt = prompt & caption & "���O�����͂���Ă��܂���B"
    End If
    
    ErrorCheck_Str = prompt
    
End Function


Private Function ErrorNotification(prompt As String)
    
    Beep
    LabelMessage.ForeColor = RGB(230, 0, 0)
    LabelMessage.caption = prompt
    
End Function

