VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form trtig 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   10515
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   15855
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   10515
   ScaleWidth      =   15855
   ShowInTaskbar   =   0   'False
   Begin MSDBGrid.DBGrid sdata 
      Height          =   3975
      Left            =   0
      OleObjectBlob   =   "trtig.frx":0000
      TabIndex        =   19
      Top             =   5280
      Width           =   12615
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   1440
      TabIndex        =   14
      Top             =   3240
      Width           =   3375
      Begin VB.OptionButton Option5 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   1800
         MaskColor       =   &H00808080&
         TabIndex        =   17
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton Option4 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   0
         MaskColor       =   &H00808080&
         TabIndex        =   16
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Preview"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   2040
         TabIndex        =   18
         Top             =   0
         Width           =   1335
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Print"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   240
         TabIndex        =   15
         Top             =   0
         Width           =   1335
      End
   End
   Begin VB.CommandButton cancel 
      BackColor       =   &H80000004&
      Height          =   550
      Left            =   3840
      Picture         =   "trtig.frx":09D1
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CommandButton gen 
      BackColor       =   &H80000004&
      Height          =   550
      Left            =   2160
      Picture         =   "trtig.frx":4182
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4440
      Width           =   1455
   End
   Begin MSComCtl2.DTPicker fdate 
      Height          =   495
      Left            =   3240
      TabIndex        =   8
      Top             =   2040
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarForeColor=   4210752
      Format          =   112132097
      CurrentDate     =   43579
   End
   Begin VB.OptionButton Option3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   4560
      MaskColor       =   &H00808080&
      TabIndex        =   7
      Top             =   840
      Width           =   255
   End
   Begin VB.OptionButton Option2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   2520
      MaskColor       =   &H00808080&
      TabIndex        =   4
      Top             =   840
      Width           =   255
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   360
      MaskColor       =   &H00808080&
      TabIndex        =   2
      Top             =   840
      Width           =   255
   End
   Begin MSComCtl2.DTPicker tdate 
      Height          =   495
      Left            =   6840
      TabIndex        =   11
      Top             =   2040
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarForeColor=   4210752
      Format          =   112132097
      CurrentDate     =   43579
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "To Date "
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
      Left            =   5280
      TabIndex        =   10
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
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
      Left            =   960
      TabIndex        =   9
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   " This Month"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   4800
      TabIndex        =   6
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   " This Week"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   2760
      TabIndex        =   5
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   " By Date "
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   600
      TabIndex        =   3
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label cl 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
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
      Width           =   615
   End
   Begin VB.Label title 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Trip Pending To Invoice Generation"
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
Attribute VB_Name = "trtig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cancel_Click()
Call ivsblfrm
End Sub
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
'title bar setting
Call settl
'form setting
Call setfrm
'visibility false
Call ivsblfrm
End Sub
Private Sub settl()
title.Width = Me.Width
title.Height = 450
cl.Top = 0
cl.Left = Me.Width - 495 - 50
End Sub
Private Sub ivsblfrm()
fdate.Visible = False
Label4.Visible = False
tdate.Visible = False
Label5.Visible = False
Frame1.Visible = False
gen.Visible = False
cancel.Visible = False
End Sub
Private Sub vsblfrm()
fdate.Visible = True
Label4.Visible = True
tdate.Visible = True
Label5.Visible = True
Frame1.Visible = True
gen.Visible = True
cancel.Visible = True
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = vbCtrlMask And KeyCode = vbKeyX Then
     Call cl_Click
   End If
End Sub
Private Sub setfrm()
Option1.Top = 100 + title.Height
Option1.Left = 100
Label1.Top = 100 + title.Height
Label1.Left = Option1.Left + Option1.Width
Option2.Top = 100 + title.Height
Option2.Left = Label1.Left + Label1.Width + 200
Label2.Top = 100 + title.Height
Label2.Left = Option2.Left + Option2.Width
Option3.Top = 100 + title.Height
Option3.Left = Label2.Left + Label2.Width + 200
Label3.Top = 100 + title.Height
Label3.Left = Option3.Left + Option3.Width
Label4.Top = Label1.Top + Label1.Height + 500
Label4.Left = 200
fdate.Top = Label4.Top
fdate.Left = Label1.Left + Label1.Width + 100
Label5.Top = Label4.Top
Label5.Left = Me.Width - Label5.Width - 450 - tdate.Width
tdate.Top = Label4.Top
tdate.Left = Label5.Left + Label5.Width + 100
Frame1.Top = Label4.Top + Label4.Height + 300
Frame1.Left = fdate.Left
Frame1.Width = Option4.Width + Label6.Width + 200 + Option5.Width + Label7.Width
Option5.Left = Label6.Left + Label6.Width + 200
Label7.Left = Option5.Left + Option5.Width
gen.Top = Frame1.Top + Label6.Height + 800
gen.Left = (Me.Width / 2) - ((gen.Width + 100 + cancel.Width) / 2)
cancel.Top = gen.Top
cancel.Left = gen.Left + gen.Width + 100
sdata.Width = Me.Width
sdata.Height = Me.Height - gen.Top - gen.Height - 1500
sdata.Top = Me.Height - sdata.Height
sdata.Left = 0
 gen.Picture = LoadPicture(App.Path & "\appdata\icons\cmd_gen.jpg")
 cancel.Picture = LoadPicture(App.Path & "\appdata\icons\cmd_cancel.jpg")
End Sub

Private Sub Label1_Click()
Option1.Value = True
Call vsblfrm
End Sub

Private Sub Label2_Click()
Option2.Value = True
Call vsblfrm
End Sub

Private Sub Label3_Click()
Option3.Value = True
Call vsblfrm
End Sub

Private Sub Option1_Click()
Call vsblfrm
End Sub

Private Sub Option2_Click()
Call vsblfrm
End Sub

Private Sub Option3_Click()
Call vsblfrm
End Sub
