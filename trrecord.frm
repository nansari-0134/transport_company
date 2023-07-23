VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form trrecord 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9420
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   12720
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9420
   ScaleWidth      =   12720
   Begin VB.Frame Frame1 
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   3360
      TabIndex        =   22
      Top             =   5280
      Width           =   2535
      Begin VB.OptionButton Option6 
         BackColor       =   &H80000003&
         Caption         =   "Option3"
         Height          =   495
         Left            =   1200
         TabIndex        =   24
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H80000003&
         Caption         =   "Option31"
         Height          =   495
         Left            =   0
         TabIndex        =   23
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label7 
         BackColor       =   &H80000003&
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
         TabIndex        =   26
         Top             =   0
         Width           =   855
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000003&
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
         TabIndex        =   25
         Top             =   0
         Width           =   855
      End
   End
   Begin VB.CommandButton gen 
      Height          =   550
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   6720
      Width           =   1450
   End
   Begin VB.CommandButton cancel 
      Height          =   550
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   6720
      Width           =   1450
   End
   Begin VB.Frame dte 
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      Caption         =   """"
      Height          =   2175
      Left            =   360
      TabIndex        =   8
      Top             =   2520
      Width           =   6975
      Begin VB.OptionButton Option2 
         BackColor       =   &H80000003&
         Caption         =   "Option1"
         Height          =   315
         Left            =   240
         TabIndex        =   10
         Top             =   720
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H80000003&
         Caption         =   "Option1"
         Height          =   315
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   255
      End
      Begin MSComCtl2.DTPicker dfdate 
         Height          =   495
         Left            =   2400
         TabIndex        =   11
         Top             =   720
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   873
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   110624769
         CurrentDate     =   43579
      End
      Begin MSComCtl2.DTPicker dtdate 
         Height          =   495
         Left            =   4920
         TabIndex        =   12
         Top             =   720
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   873
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   110624769
         CurrentDate     =   43579
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "All Dates "
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
         TabIndex        =   17
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "From Date "
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
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "To"
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
         Left            =   4320
         TabIndex        =   15
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " This Week"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   3720
         TabIndex        =   14
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " This Month"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   5160
         TabIndex        =   13
         Top             =   1440
         Width           =   1215
      End
   End
   Begin VB.Frame party 
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   360
      TabIndex        =   2
      Top             =   960
      Width           =   6975
      Begin VB.OptionButton Option4 
         BackColor       =   &H80000003&
         Caption         =   "Option1"
         Height          =   315
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   255
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H80000003&
         Caption         =   "Option1"
         Height          =   315
         Left            =   240
         TabIndex        =   4
         Top             =   120
         Width           =   255
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   495
         Left            =   2400
         TabIndex        =   3
         Top             =   720
         Width           =   3975
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Vehicle Name"
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
         TabIndex        =   7
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "All Vehicles"
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
         TabIndex        =   6
         Top             =   120
         Width           =   1575
      End
   End
   Begin MSDBGrid.DBGrid sdata 
      Height          =   2655
      Left            =   0
      OleObjectBlob   =   "trrecord.frx":0000
      TabIndex        =   21
      Top             =   7320
      Width           =   11175
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
      Left            =   360
      TabIndex        =   18
      Top             =   5280
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
      Left            =   11640
      TabIndex        =   1
      Top             =   0
      Width           =   495
   End
   Begin VB.Label title 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "         Trip record register        "
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
Attribute VB_Name = "trrecord"
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
dte.Top = party.Top + party.Height + 50
dte.Left = party.Left
Label5.Top = dte.Top + dte.Height + 250
Label5.Left = dte.Left
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

Private Sub Label1_Click()
If Option1.Value = True Then
Option1.Value = False
Else
Option1.Value = True
End If
End Sub
Private Sub Label10_Click()
If Option3.Value = True Then
Option3.Value = False
Else
Option3.Value = True
End If
End Sub
