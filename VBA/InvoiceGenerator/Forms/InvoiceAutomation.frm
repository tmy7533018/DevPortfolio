VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} InvoiceAutomation 
   Caption         =   "UserForm1"
   ClientHeight    =   9105.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7920
   OleObjectBlob   =   "InvoiceAutomation.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
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
                Set .CustomerListSheet = ThisWorkbook.Worksheets("顧客情報")
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
                Set .CustomerListSheet = ThisWorkbook.Worksheets("顧客情報")
            End With
            
    End Select
    
    
    result = MsgBox("請求書を生成しますか？", vbOKCancel, "確認画面")
    If result = 2 Then Exit Sub
    
    '請求書生成関数呼び出し
    Dim isInvoiceGenerated As Boolean
    isInvoiceGenerated = InvoiceGenerator(params, GenerateInvoiceFlag)
    
    If Not isInvoiceGenerated Then
        Beep
        MsgBox "請求書が生成できませんでした。"
        LabelMessage.caption = "請求書が生成できませんでした。" & vbCrLf & vbCrLf & "正しい値を入力してください。"
        LabelMessage.ForeColor = RGB(255, 0, 0)
        Exit Sub
    End If
    
    MsgBox "請求書を生成しました。"
    LabelMessage.ForeColor = RGB(0, 0, 0)
    LabelMessage.caption = "請求書を生成しました。" & vbCrLf & vbCrLf & "続けて生成ができます。"
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
    
    result = MsgBox("終了しますか？", vbOKCancel, "終了画面")
    If result = 1 Then Unload Me
    
    Exit Sub
    
End Sub


Private Sub MultiPageBillingOptions_Change()
    
    Dim selectedPageName As String
    
    selectedPageName = MultiPageBillingOptions.Pages(MultiPageBillingOptions.Value).name
    LabelMessage.ForeColor = RGB(0, 0, 0)
    
    If selectedPageName = "PageSingle" Then
        Me.TextBoxSingleYear.SetFocus
        LabelMessage.caption = "請求年月を入力してください。" & vbCrLf & vbCrLf & "顧客を指定する場合はチェックボックスにチェックを入れてください。"
        
    ElseIf selectedPageName = "PageRange" Then
        Me.TextBoxStartYear.SetFocus
        LabelMessage.caption = "開始年月・終了年月を入力してください。" & vbCrLf & vbCrLf & "顧客を指定する場合はチェックボックスにチェックを入れてください。"
        
    End If
    
End Sub


Private Sub UserForm_Initialize()
    
    Me.TextBoxSingleYear.SetFocus
    
End Sub


Private Sub UserForm_Click()

    LabelMessage.ForeColor = RGB(0, 0, 0)
    LabelMessage.caption = "単月指定または範囲指定で請求書を生成します。" & vbCrLf & vbCrLf & "顧客を指定する場合はチェックボックスにチェックを入れてください。"
    
End Sub


Private Function ErrorCheck_YearMonth(labelText As String, userInputYear As String, userInputMonth As String) As String
    
    Dim prompt As String
    prompt = ""
    
    If userInputYear = "" Then
        prompt = labelText & "年が入力されていません。" & vbCrLf & vbCrLf
        
    ElseIf Not IsNumeric(userInputYear) Or Len(userInputYear) <> 4 Then
        prompt = labelText & "年は4桁の数値で入力してください。" & vbCrLf & vbCrLf
        
    ElseIf CInt(userInputYear) <> CDbl(userInputYear) Or Val(userInputYear) < 0 Then
        prompt = labelText & "年に無効な値が含まれています。" & vbCrLf & vbCrLf
    
    End If
    
    If userInputMonth = "" Then
        prompt = prompt & labelText & "月が入力されていません。" & vbCrLf & vbCrLf
    
    ElseIf Not IsNumeric(userInputMonth) Or Val(userInputMonth) < 1 Or Val(userInputMonth) > 12 Then
        prompt = prompt & labelText & "月は1～12の数値で入力してください。" & vbCrLf & vbCrLf
        
    ElseIf CInt(userInputMonth) <> CDbl(userInputMonth) Then
        prompt = prompt & labelText & "月に無効な値が含まれています。" & vbCrLf & vbCrLf
        
    End If
    
    ErrorCheck_YearMonth = prompt
    
End Function


Private Function ErrorCheck_Str(caption As String, text As String) As String
    
    Dim prompt As String
    prompt = ""
    
    If text = "" Then
        prompt = prompt & caption & "名前が入力されていません。"
    End If
    
    ErrorCheck_Str = prompt
    
End Function


Private Function ErrorNotification(prompt As String)
    
    Beep
    LabelMessage.ForeColor = RGB(230, 0, 0)
    LabelMessage.caption = prompt
    
End Function

