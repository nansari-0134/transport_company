VERSION 5.00
Begin VB.Form chpass 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8820
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   12450
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8820
   ScaleWidth      =   12450
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cancel 
      Height          =   550
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3600
      Width           =   1450
   End
   Begin VB.CommandButton cnfrm 
      Height          =   550
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3600
      Width           =   1450
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H80000004&
      Height          =   495
      Left            =   2880
      TabIndex        =   7
      Top             =   2760
      Width           =   3735
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000004&
      Height          =   495
      Left            =   2880
      TabIndex        =   6
      Top             =   2040
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000004&
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   2880
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1320
      Width           =   3735
   End
   Begin VB.Shape Shape1 
      Height          =   2175
      Left            =   240
      Top             =   1200
      Width           =   6495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Re-Type Password "
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   5
      Top             =   2760
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "New Password "
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   2040
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Current Password "
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   1320
      Width           =   2415
   End
   Begin VB.Label cl 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11520
      TabIndex        =   1
      Top             =   0
      Width           =   495
   End
   Begin VB.Label title 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Change Password"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12135
   End
End
Attribute VB_Name = "chpass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cl_Click()
menu.Enabled = True
Unload Me
End Sub
Private Sub Form_Load()
'form apearance
Me.Height = Screen.Height - Screen.Height * 5 / 100 - 450
Me.Width = Screen.Width * 80 / 100
Me.Top = 450
Me.Left = Screen.Width * 20 / 100
Me.BackColor = vbWhite
Me.Picture = LoadPicture(App.Path & "\appdata\images\back.jpg")
'title bar setting
Call settl
' label setting
Call setlabel
End Sub
Private Sub settl()
title.Width = Me.Width
title.Height = 450
cl.Top = 0
cl.Left = Me.Width - 495 - 50
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = vbCtrlMask And KeyCode = vbKeyX Then
     Call cl_Click
   End If
End Sub
Private Sub setlabel()
Shape1.Top = title.Height + 500
Shape1.Left = Me.Width / 2 - Shape1.Width / 2
Label1.Top = Shape1.Top + 120
Label1.Left = Shape1.Left + 120
Label2.Top = Label1.Top + Label1.Height + 120
Label2.Left = Shape1.Left + 120
Label3.Top = Label2.Top + Label2.Height + 120
Label3.Left = Shape1.Left + 120
Text1.Top = Label1.Top
Text2.Top = Label2.Top
Text3.Top = Label3.Top
Text1.Left = Label1.Left + Label1.Width + 100
Text2.Left = Label2.Left + Label2.Width + 100
Text3.Left = Label3.Left + Label3.Width + 100
cnfrm.Top = Shape1.Top + 500 + Shape1.Height
cnfrm.Left = Me.Width / 2 - (cnfrm.Width + 200 + cancel.Width) / 2
cancel.Top = cnfrm.Top
cancel.Left = cnfrm.Left + cnfrm.Width + 100
cnfrm.Picture = LoadPicture(App.Path & "\appdata\icons\cmd_cfnm.jpg")
cancel.Picture = LoadPicture(App.Path & "\appdata\icons\cmd_cancel.jpg")
End Sub
