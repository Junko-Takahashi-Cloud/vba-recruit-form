VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDayButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents btn As MSForms.CommandButton
Attribute btn.VB_VarHelpID = -1
Public ParentForm As frmCalendar

Private Sub Btn_Click()
    If btn.Caption <> "" Then
        ParentForm.SetDateFromButton btn.Caption
    End If
End Sub

