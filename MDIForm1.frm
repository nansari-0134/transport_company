VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   Moveable        =   0   'False
   ScrollBars      =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long

Const WS_CAPTION = &HC00000
Const WS_SYSMENU = &H80000
Const WS_MINIMIZEBOX = &H20000
Const WS_MAXIMIZEBOX = &H10000
Const GWL_STYLE = (-16)
 
Private Sub MDIForm_Load()
    Dim L As Long
    
    L = GetWindowLong(Me.hwnd, GWL_STYLE)
    L = L And Not (WS_MINIMIZEBOX)
    L = L And Not (WS_MAXIMIZEBOX)
    L = L Xor WS_CAPTION
    L = SetWindowLong(Me.hwnd, GWL_STYLE, L)
    Me.Top = -100
    Me.Left = -100
    Me.Height = Screen.Height * 95 / 100 + 200
    Me.Width = Screen.Width + 300
End Sub
