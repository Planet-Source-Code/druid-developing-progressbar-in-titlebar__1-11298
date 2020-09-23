VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Load..."
   ClientHeight    =   2580
   ClientLeft      =   4680
   ClientTop       =   3675
   ClientWidth     =   7260
   LinkTopic       =   "Form1"
   ScaleHeight     =   2580
   ScaleWidth      =   7260
   Begin VB.Timer tmrProgress 
      Interval        =   50
      Left            =   1560
      Top             =   1440
   End
   Begin MSComctlLib.ProgressBar prInTitleBar 
      Height          =   255
      Left            =   1680
      TabIndex        =   0
      Top             =   600
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Call Init
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Terminate
End Sub

Private Sub tmrProgress_Timer()
    prInTitleBar.Value = prInTitleBar.Value + 2
    If prInTitleBar.Value > 99 Then
    'Stop the ProgressBar an make it invisible
    prInTitleBar.Value = 0
    prInTitleBar.Visible = False
    'Stop the timer
    tmrProgress.Enabled = False
    Me.Caption = "ProgressBar in Titlebar!"
    MsgBox "Did you like it? If yes, please vote for me!", vbQuestion, "Finished"
    End If
End Sub
