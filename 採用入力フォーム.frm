VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 採用入力フォーム 
   Caption         =   "採用入力フォーム"
   ClientHeight    =   6624
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6576
   OleObjectBlob   =   "採用入力フォーム.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "採用入力フォーム"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn開始_Click()
    Call ClearForm
    MsgBox "入力フォームを初期化しました。", vbInformation, "開始"
End Sub

Private Sub btn登録_Click()
    Dim ws As Worksheet
    Dim lastRow As Long
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    If txt応募者名.Value = "" Then
        MsgBox "応募者名を入力してください", vbExclamation
        txt応募者名.SetFocus
        Exit Sub
    End If
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row + 1

    ws.Cells(lastRow, 1).Value = txt応募者名.Value
    ws.Cells(lastRow, 2).Value = txt応募日.Value
    ws.Cells(lastRow, 3).Value = cmb職種.Value
    ws.Cells(lastRow, 4).Value = cmbステータス.Value
    ws.Cells(lastRow, 5).Value = txt面接日.Value
    ws.Cells(lastRow, 6).Value = cmb面接官.Value
    ws.Cells(lastRow, 7).Value = cmb評価.Value
    ws.Cells(lastRow, 8).Value = txt備考.Value

    MsgBox txt応募者名.Value & " さんを登録しました！", vbInformation, "完了"

    '--- 入力欄クリア ---
    Call ClearForm

End Sub

Private Sub btn閉じる_Click()
    Unload Me
End Sub
Private Sub txt応募日_Enter()
    Set frmCalendar.targetTextBox = Me.txt応募日
    Debug.Print frmCalendar.targetTextBox.Name   ' ← ここで確認
    frmCalendar.Show
End Sub

Private Sub txt面接日_Enter()
    Set frmCalendar.targetTextBox = Me.txt面接日
    Debug.Print frmCalendar.targetTextBox.Name   ' ← ここで確認
    frmCalendar.Show
End Sub

Private Sub UserForm_Initialize()

    '職種
    Dim rng職種 As Range
    Set rng職種 = Sheets("項目名").Range("A2:A21")
    Dim c As Range
    For Each c In rng職種
        If c.Value <> "" Then Me.cmb職種.AddItem c.Value
    Next c

    'ステータス
    Dim rngステータス As Range
    Set rngステータス = Sheets("項目名").Range("B2:B10")
    For Each c In rngステータス
        If c.Value <> "" Then Me.cmbステータス.AddItem c.Value
    Next c

    '面接官
    Dim rng面接官 As Range
    Set rng面接官 = Sheets("項目名").Range("C2:C10")
    For Each c In rng面接官
        If c.Value <> "" Then Me.cmb面接官.AddItem c.Value
    Next c

    '評価
    Dim rng評価 As Range
    Set rng評価 = Sheets("項目名").Range("D2:D10")
    For Each c In rng評価
        If c.Value <> "" Then Me.cmb評価.AddItem c.Value
    Next c

End Sub

Private Sub ClearForm()
    Dim ctrl As Control
    For Each ctrl In Me.Controls
        Select Case TypeName(ctrl)
            Case "TextBox", "ComboBox"
                ctrl.Value = ""
            Case "CheckBox", "OptionButton"
                ctrl.Value = False
        End Select
    Next ctrl

    ' 最初の入力欄にフォーカス戻す
    txt応募者名.SetFocus
End Sub

