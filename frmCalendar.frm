VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCalendar 
   Caption         =   "カレンダー"
   ClientHeight    =   5400
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6972
   OleObjectBlob   =   "frmCalendar.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public targetTextBox As MSForms.TextBox
Dim currentDate As Date
Dim DayButtons(1 To 42) As clsDayButton

' 親フォームから呼び出す入口
Public Sub ShowCalendar(ByVal txt As Object)

    Set targetTextBox = txt
    currentDate = Date
    Me.Show
End Sub
Private Sub CreateCalendarButtons()

    Dim btn As MSForms.CommandButton
    Dim i As Long
    Dim row As Long, col As Long
    Dim leftPos As Single, topPos As Single

    leftPos = 60
    topPos = 60

    For i = 1 To 42

        Set btn = Me.Controls.Add("Forms.CommandButton.1", "btnDay" & i)

        row = Int((i - 1) / 7)
        col = (i - 1) Mod 7

        btn.Width = 30
        btn.Height = 25
        btn.Left = leftPos + col * 32
        btn.Top = topPos + row * 27
        btn.Caption = ""

        Set DayButtons(i) = New clsDayButton
        Set DayButtons(i).btn = btn
        Set DayButtons(i).ParentForm = Me

    Next i

End Sub


Private Sub UserForm_Initialize()
    
   Me.BackColor = RGB(255, 182, 193)
   
   lblYearMonth.BackColor = RGB(255, 200, 210)   ' ← ここに追加
    
    Call CreateCalendarButtons
    currentDate = Date
    Call DrawCalendar

End Sub

' 前月
Private Sub btnPrev_Click()
    currentDate = DateAdd("m", -1, currentDate)
    Call DrawCalendar
End Sub

' 次月
Private Sub btnNext_Click()
    currentDate = DateAdd("m", 1, currentDate)
    Call DrawCalendar
End Sub

' カレンダー描画
Private Sub DrawCalendar()
    Dim firstDay As Date
    Dim startCell As Integer
    Dim lastDay As Integer
    Dim i As Integer

    firstDay = DateSerial(Year(currentDate), Month(currentDate), 1)
    startCell = Weekday(firstDay, vbSunday)
    lastDay = Day(DateSerial(Year(currentDate), Month(currentDate) + 1, 0))

    lblYearMonth.Caption = Year(currentDate) & "年 " & Month(currentDate) & "月"

    ' 全ボタン初期化
    For i = 1 To 42
        With Me.Controls("btnDay" & i)
            .Caption = ""
            .Enabled = False
        End With
    Next i

    ' 日付セット
    For i = 1 To lastDay
        With Me.Controls("btnDay" & (startCell - 1 + i))
            .Caption = i
            .Enabled = True
        End With
    Next i
End Sub

Public Sub SetDateFromButton(dayValue As String)
    If Not targetTextBox Is Nothing Then
        targetTextBox.Value = Year(currentDate) & "/" & _
                              Month(currentDate) & "/" & _
                              dayValue
    End If
    Unload Me
End Sub

