VERSION 5.00
Begin VB.Form login 
   BorderStyle     =   0  'None
   Caption         =   "4343"
   ClientHeight    =   10125
   ClientLeft      =   0
   ClientTop       =   345
   ClientWidth     =   17805
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   10125
   ScaleWidth      =   17805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6120
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1080
      TabIndex        =   2
      Top             =   4680
      Width           =   4815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Password :"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   3960
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   840
      Left            =   14640
      Stretch         =   -1  'True
      Top             =   8640
      Width           =   1560
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "M.S.K. Yadu Transport "
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1095
      Left            =   960
      TabIndex        =   0
      Top             =   360
      Width           =   6375
   End
End
Attribute VB_Name = "login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
menu.Show
End Sub
Private Sub Command1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call Command1_Click
End If
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call Command1_Click
End If
End Sub
Private Sub Form_Load()
Me.Height = Screen.Height - Screen.Height * 5 / 100
Me.Width = Screen.Width
Me.Top = 0
Me.Left = 0
Me.BackColor = &HFFC0C0
Call setfrm
End Sub
Private Sub Image1_Click()
End
End Sub
Private Sub setfrm()
Label1.Top = 300
Label1.Left = 300
Label3.Top = Me.Height / 2 - (Label3.Height + 400 + Command1.Height + Text1.Height) / 2
Label3.Left = Label1.Left
Text1.Top = Label3.Top + Label3.Height + 200
Text1.Left = Label1.Left
Command1.Top = Text1.Top + Text1.Height + 200
Command1.Left = Text1.Left + Text1.Width - Command1.Width
Image1.Height = 840
Image1.Width = 1560
Image1.Left = Me.Width - Image1.Width - 450
Image1.Top = Me.Height - Image1.Height - 450
Me.Picture = LoadPicture(App.Path & "\appdata\images\lback.jpg")
Image1.Picture = LoadPicture(App.Path & "\appdata\images\exit.jpg")
Command1.Picture = LoadPicture(App.Path & "\appdata\icons\cmd_login.jpg")
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call Command1_Click
End If
End Sub
