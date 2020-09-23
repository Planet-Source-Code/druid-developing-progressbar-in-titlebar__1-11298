Attribute VB_Name = "modMain"
Option Explicit

'*********************
'* API Declarations  *
'*********************
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As Rect) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook&, ByVal lpfn&, ByVal hmod&, ByVal dwThreadId&) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook&) As Long

'*********************
'* Type Declarations *
'*********************
Private Type Rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type CWPSTRUCT
    lParam As Long
    wParam As Long
    Message As Long
    hwnd As Long
End Type

'*********************
'* Consts            *
'*********************
Const WM_MOVE = &H3
Const WM_SETCURSOR = &H20
Const WM_NCPAINT = &H85
Const WM_COMMAND = &H111

Const SWP_FRAMECHANGED = &H20
Const GWL_EXSTYLE = -20

'*********************
'* Vars              *
'*********************
Private WHook&

Public Sub Init()
  'Initialize the window hooking for the button
  WHook = SetWindowsHookEx(4, AddressOf HookProc, 0, App.ThreadID)
  Call SetWindowLong(Form1.prInTitleBar.hwnd, GWL_EXSTYLE, &H80)
  Call SetParent(Form1.prInTitleBar.hwnd, GetParent(Form1.hwnd))
End Sub

Public Sub Terminate()
  'Terminate the window hooking
  Call UnhookWindowsHookEx(WHook)
  Call SetParent(Form1.prInTitleBar.hwnd, Form1.hwnd)
End Sub

Public Function HookProc&(ByVal nCode&, ByVal wParam&, Inf As CWPSTRUCT)
    Dim FormRect As Rect
    Static LastParam&
    If Inf.hwnd = GetParent(Form1.prInTitleBar.hwnd) Then
        ElseIf Inf.hwnd = Form1.hwnd Then
        If Inf.Message = WM_NCPAINT Or Inf.Message = WM_MOVE Then
            'Get the size of the Form
            Call GetWindowRect(Form1.hwnd, FormRect)
            'Place the button int the Titlebar
            Call SetWindowPos(Form1.prInTitleBar.hwnd, 0, FormRect.Left + 180, FormRect.Top + 6, FormRect.Right - 60 - (FormRect.Left + 180), 14, SWP_FRAMECHANGED)
        End If
    End If
End Function
