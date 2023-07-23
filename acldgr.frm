VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form acldgr 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   10170
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   13470
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   10170
   ScaleWidth      =   13470
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cancel 
      Height          =   550
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   8880
      Width           =   1450
   End
   Begin VB.CommandButton gen 
      Height          =   550
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   8880
      Width           =   1450
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   2160
      TabIndex        =   35
      Top             =   8160
      Width           =   2535
      Begin VB.OptionButton Option5 
         BackColor       =   &H80000003&
         Caption         =   "Option31"
         Height          =   495
         Left            =   0
         TabIndex        =   37
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton Option6 
         BackColor       =   &H80000003&
         Caption         =   "Option3"
         Height          =   495
         Left            =   1200
         TabIndex        =   36
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label16 
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
         TabIndex        =   39
         Top             =   0
         Width           =   855
      End
      Begin VB.Label Label15 
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
         TabIndex        =   38
         Top             =   0
         Width           =   855
      End
   End
   Begin MSComCtl2.DTPicker dt1 
      Height          =   495
      Left            =   2160
      TabIndex        =   33
      Top             =   6360
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   110624769
      CurrentDate     =   43583
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   3100
      Left            =   3360
      TabIndex        =   27
      Top             =   2520
      Width           =   4935
      Begin MSComctlLib.ListView lv 
         Height          =   2655
         Left            =   0
         TabIndex        =   30
         Top             =   480
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   4683
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         FlatScrollBar   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ind"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "name"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.CheckBox Check10 
         Caption         =   "Check1"
         Height          =   210
         Left            =   240
         TabIndex        =   28
         Top             =   120
         Width           =   210
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Accounts"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   29
         Top             =   0
         Width           =   1575
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H80000003&
         FillColor       =   &H80000003&
         FillStyle       =   0  'Solid
         Height          =   489
         Left            =   0
         Top             =   0
         Width           =   4935
      End
   End
   Begin VB.Frame actype 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3135
      Left            =   240
      TabIndex        =   10
      Top             =   2520
      Width           =   2535
      Begin VB.CheckBox Check9 
         Caption         =   "Check1"
         Height          =   210
         Left            =   120
         TabIndex        =   19
         Top             =   2760
         Width           =   210
      End
      Begin VB.CheckBox Check8 
         Caption         =   "Check1"
         Height          =   210
         Left            =   120
         TabIndex        =   18
         Top             =   2400
         Width           =   210
      End
      Begin VB.CheckBox Check7 
         Caption         =   "Check1"
         Height          =   210
         Left            =   120
         TabIndex        =   17
         Top             =   2040
         Width           =   210
      End
      Begin VB.CheckBox Check6 
         Caption         =   "Check1"
         Height          =   210
         Left            =   120
         TabIndex        =   16
         Top             =   1680
         Width           =   210
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Check1"
         Height          =   210
         Left            =   120
         TabIndex        =   15
         Top             =   1320
         Width           =   210
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Check1"
         Height          =   210
         Left            =   120
         TabIndex        =   14
         Top             =   960
         Width           =   210
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Check1"
         Height          =   210
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   210
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Check1"
         Height          =   210
         Left            =   120
         TabIndex        =   11
         Top             =   120
         Width           =   210
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         Height          =   3135
         Left            =   0
         Top             =   0
         Width           =   2535
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Stock Godown"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   26
         Top             =   2760
         Width           =   1695
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Vehicle"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   25
         Top             =   2400
         Width           =   1695
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "General"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   24
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Purchase Party"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   23
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Driver"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   22
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Khadan"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   21
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   20
         Top             =   600
         Width           =   1695
      End
      Begin VB.Line Line2 
         X1              =   375
         X2              =   375
         Y1              =   0
         Y2              =   3135
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Account Types"
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
         Left            =   480
         TabIndex        =   12
         Top             =   0
         Width           =   1695
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   2520
         Y1              =   480
         Y2              =   480
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   210
      Left            =   840
      TabIndex        =   8
      Top             =   2160
      Width           =   210
   End
   Begin VB.Frame party 
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   840
      TabIndex        =   2
      Top             =   480
      Width           =   6975
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
         Left            =   1920
         TabIndex        =   5
         Top             =   720
         Width           =   4695
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H80000003&
         Caption         =   "Option1"
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H80000003&
         Caption         =   "Option1"
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "All Areas"
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
         Left            =   480
         TabIndex        =   7
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Area Name"
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
         Left            =   480
         TabIndex        =   6
         Top             =   720
         Width           =   1455
      End
   End
   Begin MSComCtl2.DTPicker dt2 
      Height          =   495
      Left            =   2160
      TabIndex        =   34
      Top             =   7200
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   110624769
      CurrentDate     =   43583
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Option"
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
      TabIndex        =   40
      Top             =   8160
      Width           =   1575
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "To Date"
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
      TabIndex        =   32
      Top             =   7200
      Width           =   1575
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "From Date"
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
      TabIndex        =   31
      Top             =   6360
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Show only Account Balance"
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
      Left            =   1200
      TabIndex        =   9
      Top             =   2040
      Width           =   3015
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
      Caption         =   "Account Ledger"
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
Attribute VB_Name = "acldgr"
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
'label settings
Call setlabel
'title bar setting
Call settl
lv.ColumnHeaders(1).Width = lv.Width * 15 / 100
lv.ColumnHeaders(2).Width = lv.Width * 84 / 100
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
party.Top = title.Height + 200
party.Left = 200
Check1.Top = party.Top + party.Height + 200
Check1.Left = party.Left + 120
Label3.Top = Check1.Top - 90
Label3.Left = Check1.Left + Check1.Width + 120
actype.Left = party.Left
actype.Top = Label3.Height + Label3.Top + 100
Frame1.Width = Me.Width - 220 - actype.Width - actype.Left
Frame1.Top = actype.Top
Frame1.Left = actype.Left + actype.Width + 100
Shape1.Width = Frame1.Width
lv.Width = Frame1.Width
Label13.Left = 200
Label14.Left = 200
Label18.Left = 200
Label13.Top = actype.Top + actype.Height + 400
Label14.Top = Label13.Top + Label13.Height + 140
Label18.Top = Label14.Top + Label14.Height + 140
dt1.Left = Label13.Left + Label13.Width + 200
dt2.Left = Label14.Left + Label14.Width + 200
dt1.Top = Label13.Top
dt2.Top = Label14.Top
Frame2.Left = dt1.Left
Frame2.Top = Label18.Top
gen.Top = Frame2.Top + Frame2.Height + 900
cancel.Top = gen.Top
gen.Left = 300
cancel.Left = gen.Left + gen.Width + 200

End Sub

Private Sub Label1_Click()
If Option3.Value = False Then
 Option3.Value = True
Else
Option3.Value = False
End If
End Sub

Private Sub Label10_Click()
If Check8.Value = 0 Then
  Check8.Value = 1
Else
  Check8.Value = 0
End If
End Sub

Private Sub Label11_Click()
If Check9.Value = 0 Then
  Check9.Value = 1
Else
  Check9.Value = 0
End If
End Sub

Private Sub Label15_Click()
If Option6.Value = False Then
 Option6.Value = True
Else
Option6.Value = False
End If
End Sub

Private Sub Label16_Click()
If Option5.Value = False Then
 Option5.Value = True
Else
Option5.Value = False
End If
End Sub

Private Sub Label2_Click()
If Option4.Value = False Then
 Option4.Value = True
Else
Option4.Value = False
End If
End Sub

Private Sub Label3_Click()
If Check1.Value = 0 Then
  Check1.Value = 1
Else
  Check1.Value = 0
End If

End Sub

Private Sub Label4_Click()
If Check2.Value = 0 Then
  Check2.Value = 1
Else
  Check2.Value = 0
End If
End Sub

Private Sub Label5_Click()
If Check3.Value = 0 Then
  Check3.Value = 1
Else
  Check3.Value = 0
End If
End Sub

Private Sub Label6_Click()
If Check4.Value = 0 Then
  Check4.Value = 1
Else
  Check4.Value = 0
End If
End Sub

Private Sub Label7_Click()
If Check5.Value = 0 Then
  Check5.Value = 1
Else
  Check5.Value = 0
End If
End Sub

Private Sub Label8_Click()
If Check6.Value = 0 Then
  Check6.Value = 1
Else
  Check6.Value = 0
End If
End Sub

Private Sub Label9_Click()
If Check7.Value = 0 Then
  Check7.Value = 1
Else
  Check7.Value = 0
End If
End Sub
