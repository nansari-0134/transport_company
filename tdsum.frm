VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form tdsum 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9585
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   17460
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9585
   ScaleWidth      =   17460
   ShowInTaskbar   =   0   'False
   Begin MSDBGrid.DBGrid sdata 
      Height          =   2655
      Left            =   240
      OleObjectBlob   =   "tdsum.frx":0000
      TabIndex        =   18
      Top             =   6600
      Width           =   12255
   End
   Begin VB.Frame party 
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   600
      TabIndex        =   12
      Top             =   600
      Width           =   6975
      Begin VB.TextBox Text1 
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   495
         Left            =   2160
         TabIndex        =   17
         Top             =   720
         Width           =   4575
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H80000003&
         Caption         =   "Option1"
         Height          =   315
         Left            =   240
         TabIndex        =   14
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H80000003&
         Caption         =   "Option1"
         Height          =   315
         Left            =   240
         TabIndex        =   13
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "All Driver"
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
         Left            =   600
         TabIndex        =   16
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Driver Name"
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
         Left            =   600
         TabIndex        =   15
         Top             =   720
         Width           =   1455
      End
   End
   Begin VB.CommandButton cancel 
      Height          =   550
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6000
      Width           =   1450
   End
   Begin VB.CommandButton gen 
      Height          =   550
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6000
      Width           =   1450
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   2280
      TabIndex        =   5
      Top             =   4920
      Width           =   2535
      Begin VB.OptionButton Option5 
         Caption         =   "Option31"
         Height          =   495
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton Option6 
         Caption         =   "Option3"
         Height          =   495
         Left            =   1200
         TabIndex        =   6
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label6 
         Caption         =   "Print"
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
         TabIndex        =   8
         Top             =   0
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Preview"
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
         Left            =   1560
         TabIndex        =   7
         Top             =   0
         Width           =   855
      End
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2280
      TabIndex        =   3
      Top             =   4560
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Options"
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
      Left            =   600
      TabIndex        =   4
      Top             =   4920
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Transaction"
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
      Left            =   600
      TabIndex        =   2
      Top             =   4320
      Width           =   1455
   End
   Begin VB.Label cl 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
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
      Caption         =   "Trip Driver Summary"
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
Attribute VB_Name = "tdsum"
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
Me.Width = Screen.Width * 40 / 100
Me.Top = 450
Me.Left = Screen.Width * 20 / 100
Me.BackColor = vbWhite
Me.Picture = LoadPicture(App.Path & "\appdata\images\back.jpg")
 gen.Picture = LoadPicture(App.Path & "\appdata\icons\cmd_gen.jpg")
 cancel.Picture = LoadPicture(App.Path & "\appdata\icons\cmd_cancel.jpg")
'title bar setting
Call settl
'form setting
Call setfrm
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

Private Sub setfrm()
party.Top = title.Height + 100
party.Left = (Me.Width / 2) - (party.Width / 2)
Label4.Top = party.Top + party.Height + 100
Combo1.Top = Label4.Top
Combo1.Left = Label4.Left + Label4.Width + 100
Label5.Top = Label4.Top + Label4.Height + 150
Frame1.Top = Label5.Top
Frame1.Left = Label5.Left + Label5.Width + 100
gen.Top = Label5.Top + Label5.Height + 300
gen.Left = (Me.Width / 2) - ((gen.Width + 100 + cancel.Width) / 2)
cancel.Top = gen.Top
cancel.Left = gen.Left + gen.Width + 100
sdata.Width = Me.Width
sdata.Height = Me.Height - gen.Top - gen.Height - 1500
sdata.Top = Me.Height - sdata.Height
sdata.Left = 0
End Sub
