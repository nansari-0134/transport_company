VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form vhldgr 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2.45655e5
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   12450
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2.45655e5
   ScaleWidth      =   12450
   Begin VB.Frame Frame2 
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   2160
      TabIndex        =   10
      Top             =   9000
      Width           =   2535
      Begin VB.OptionButton Option6 
         BackColor       =   &H80000003&
         Caption         =   "Option3"
         Height          =   495
         Left            =   1200
         TabIndex        =   12
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H80000003&
         Caption         =   "Option31"
         Height          =   495
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label15 
         BackColor       =   &H80000003&
         Caption         =   "Display"
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
         TabIndex        =   14
         Top             =   55
         Width           =   855
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
         TabIndex        =   13
         Top             =   55
         Width           =   855
      End
   End
   Begin VB.CommandButton gen 
      Height          =   550
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   9720
      Width           =   1450
   End
   Begin VB.CommandButton cancel 
      Height          =   550
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   9720
      Width           =   1450
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4665
      Left            =   360
      TabIndex        =   4
      Top             =   1920
      Width           =   8775
      Begin VB.CheckBox Check10 
         Caption         =   "Check1"
         Height          =   210
         Left            =   6720
         TabIndex        =   6
         Top             =   150
         Width           =   210
      End
      Begin MSComctlLib.ListView lv 
         Height          =   4095
         Left            =   0
         TabIndex        =   5
         Top             =   480
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   7223
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
            Text            =   "vname"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ind"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "  Vehicle List"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   0
         TabIndex        =   21
         Top             =   100
         Width           =   2895
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Select all"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7080
         TabIndex        =   20
         Top             =   100
         Width           =   1575
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H80000003&
         FillColor       =   &H80000003&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   0
         Top             =   0
         Width           =   8775
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
         TabIndex        =   7
         Top             =   0
         Width           =   1575
      End
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   3
      Top             =   600
      Width           =   2295
   End
   Begin MSComCtl2.DTPicker dt1 
      Height          =   495
      Left            =   2160
      TabIndex        =   15
      Top             =   7200
      Width           =   1935
      _ExtentX        =   3413
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
      Format          =   110886913
      CurrentDate     =   43583
   End
   Begin MSComCtl2.DTPicker dt2 
      Height          =   495
      Left            =   2160
      TabIndex        =   16
      Top             =   8040
      Width           =   1935
      _ExtentX        =   3413
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
      Format          =   110886913
      CurrentDate     =   43583
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " From Date"
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
      TabIndex        =   19
      Top             =   7200
      Width           =   1575
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " To Date"
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
      Top             =   8040
      Width           =   1575
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Option"
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
      TabIndex        =   17
      Top             =   9000
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Vehicle No."
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
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   1575
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
      Caption         =   "Vehicle Ledger"
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
Attribute VB_Name = "vhldgr"
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
'label settings
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
Label1.Top = 200 + title.Height
Label1.Left = 200
Text1.Top = Label1.Top
Text1.Left = Label1.Left + Label1.Width + 100
Frame1.Width = Me.Width - 400
Frame1.Top = Label1.Top + Label1.Height + 500
Frame1.Left = 200
Shape1.Width = Frame1.Width
lv.Width = Frame1.Width
lv.Height = Frame1.Height - Shape1.Height
lv.ColumnHeaders(1).Width = lv.Width * 74 / 100
lv.ColumnHeaders(2).Width = lv.Width * 25 / 100
Check10.Left = lv.Width - Check10.Width - lv.Width * 22 / 100
Label2.Left = Check10.Left + 100 + Check10.Width
Label13.Left = 200
Label14.Left = 200
Label18.Left = 200
Label13.Top = Frame1.Top + Frame1.Height + 500
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
