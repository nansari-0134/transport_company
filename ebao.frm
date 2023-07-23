VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form ebao 
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
   Begin MSComCtl2.DTPicker dt1 
      Height          =   495
      Left            =   2040
      TabIndex        =   39
      Top             =   6360
      Width           =   1695
      _ExtentX        =   2990
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
      Format          =   112525313
      CurrentDate     =   43584
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      ItemData        =   "ebao.frx":0000
      Left            =   2040
      List            =   "ebao.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   38
      Top             =   7200
      Width           =   2295
   End
   Begin VB.CommandButton cancel 
      Height          =   550
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   8880
      Width           =   1450
   End
   Begin VB.CommandButton gen 
      Height          =   550
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   8880
      Width           =   1450
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   2160
      TabIndex        =   29
      Top             =   8160
      Width           =   2535
      Begin VB.OptionButton Option5 
         BackColor       =   &H80000003&
         Caption         =   "Option31"
         Height          =   495
         Left            =   0
         TabIndex        =   31
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton Option6 
         BackColor       =   &H80000003&
         Caption         =   "Option3"
         Height          =   495
         Left            =   1200
         TabIndex        =   30
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
         TabIndex        =   33
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
         TabIndex        =   32
         Top             =   0
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   3100
      Left            =   3360
      TabIndex        =   25
      Top             =   2520
      Width           =   4935
      Begin VB.CheckBox Check10 
         Caption         =   "Check1"
         Height          =   210
         Left            =   240
         TabIndex        =   26
         Top             =   120
         Width           =   210
      End
      Begin MSComctlLib.ListView lv 
         Height          =   2655
         Left            =   0
         TabIndex        =   40
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
         TabIndex        =   27
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
      TabIndex        =   8
      Top             =   2520
      Width           =   2535
      Begin VB.CheckBox Check9 
         Caption         =   "Check1"
         Height          =   210
         Left            =   120
         TabIndex        =   17
         Top             =   2760
         Width           =   210
      End
      Begin VB.CheckBox Check8 
         Caption         =   "Check1"
         Height          =   210
         Left            =   120
         TabIndex        =   16
         Top             =   2400
         Width           =   210
      End
      Begin VB.CheckBox Check7 
         Caption         =   "Check1"
         Height          =   210
         Left            =   120
         TabIndex        =   15
         Top             =   2040
         Width           =   210
      End
      Begin VB.CheckBox Check6 
         Caption         =   "Check1"
         Height          =   210
         Left            =   120
         TabIndex        =   14
         Top             =   1680
         Width           =   210
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Check1"
         Height          =   210
         Left            =   120
         TabIndex        =   13
         Top             =   1320
         Width           =   210
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Check1"
         Height          =   210
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   210
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Check1"
         Height          =   210
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   210
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Check1"
         Height          =   210
         Left            =   120
         TabIndex        =   9
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
         TabIndex        =   24
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
         TabIndex        =   23
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
         TabIndex        =   22
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
         TabIndex        =   21
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
         TabIndex        =   20
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
         TabIndex        =   19
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
         TabIndex        =   18
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
         TabIndex        =   10
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
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Filter"
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
      TabIndex        =   37
      Top             =   7200
      Width           =   1095
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
      TabIndex        =   34
      Top             =   8160
      Width           =   1095
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Date"
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
      TabIndex        =   28
      Top             =   6360
      Width           =   1095
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
      Caption         =   "Entry Based Account Outstanding"
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
Attribute VB_Name = "ebao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim sshape As String
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
cn.Mode = adModeShareDenyNone
cn.Provider = "MSDATASHAPE"
cn.Open "Data Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & App.Path & "\try.accdb;Persist Security Info=False"
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
actype.Left = party.Left
actype.Top = party.Height + party.Top + 200
Frame1.Width = Me.Width - 220 - actype.Width - actype.Left
Frame1.Top = actype.Top
Frame1.Left = actype.Left + actype.Width + 100
Shape1.Width = Frame1.Width
lv.Width = Frame1.Width
Label13.Left = 200
Label3.Left = 200
Label18.Left = 200
Label13.Top = actype.Top + actype.Height + 400
Label3.Top = Label13.Top + Label13.Height + 140
Label18.Top = Label3.Top + Label3.Height + 140
dt1.Left = Label13.Left + Label13.Width + 200
Combo1.Left = Label3.Left + Label3.Width + 200
dt1.Top = Label13.Top
Combo1.Top = Label3.Top
Frame2.Left = dt1.Left
Frame2.Top = Label18.Top
gen.Top = Frame2.Top + Frame2.Height + 900
cancel.Top = gen.Top
gen.Left = 300
cancel.Left = gen.Left + gen.Width + 200
Combo1.AddItem "Debit Balance"
Combo1.AddItem "Credit Balance"
Combo1.AddItem "All"
Combo1.Text = "Debit Balance"
End Sub
Private Sub Form_Unload(cancel As Integer)
Set rs = Nothing
cn.Close
End Sub
Private Sub gen_Click()
'On Error GoTo errorhandler:
'sshape = "Shape {select trns,vid,narratn,dr from db} Append ({select dr,vid,adate,refno,cr,amt from db} As ch Relate vid To vid) As rec"
'sshape = "Shape {SELECT vid,dr,adate,refno,cr,trns,amt,narratn FROM db} Append({SELECT vid,dr,cr FROM db} Relate 'vid' To 'vid')"
'sshape = "Shape {select * from db} As DATA COMPUTE DATA By dr,cr,adate,trns,amt,narratn"
'sshape = "SHAPE APPEND {SELECT DISTINCT trns from db} ((SHAPE APPEND {select * from db}) As data RELATE trns to trns)"
'sshape = "SHAPE {select distinct trns from db} as parent append({select vid,adate,dr,cr,trns,amt,narratn from db} As data relate trns to trns) as rec"
'sshape = "SHAPE {select distinct vid from db} as parent " _
         & "APPEND({select vid,dr,cr,amt,narratn,adate,trns from db} As data RELATE vid to vid) as rec"
sshape = "SHAPE {select distinct vid from db} APPEND({SELECT * FROM db} as data relate vid To vid)"
Set obj = New PrinterControl
obj.ChngOrientationPortrait
rpt_ebao.Width = Printer.Width - 365
rs.Open sshape, cn, adOpenStatic
If rs.EOF = False Then
    Call setrptr
    rpt_ebao.Show
    'Unload Me
End If
Exit Sub
errorhandler: MsgBox err.Description
obj.ReSetOrientation
End Sub
Private Sub setrptr()
Printer.ScaleMode = vbInches 'shilu's coding
Printer.CurrentX = 1
Printer.CurrentY = 1
Printer.Zoom = 0.75
With rpt_ebao
  Set .DataSource = rs
  .DataMember = rs.DataMember
  .Left = 0
  .Top = 0
   With .Sections("section4")
      .Controls("label1").Left = rpt_ebao.Width / 2 - .Controls("label1").Width / 2 - 200
      .Controls("label3").Left = rpt_ebao.Width / 2 - .Controls("label3").Width / 2 - 200
   End With
   With .Sections("section2")
      .Controls("label2").Width = rpt_ebao.Width * 8 / 100 - 50
      .Controls("label4").Width = rpt_ebao.Width * 9 / 100 - 50
      .Controls("label5").Width = rpt_ebao.Width * 10 / 100 - 50
      .Controls("label6").Width = rpt_ebao.Width * 18 / 100 - 50
      .Controls("label7").Width = rpt_ebao.Width * 18 / 100 - 50
      .Controls("label8").Width = rpt_ebao.Width * 15 / 100 - 50
      .Controls("label9").Width = rpt_ebao.Width * 12 / 100 - 50
      .Controls("label10").Width = rpt_ebao.Width * 10 / 100 - 50
      .Controls("shape3").Left = 0
      .Controls("shape3").Top = 0
      .Controls("shape3").Width = rpt_ebao.Width
      .Controls("shape2").Left = 0
      .Controls("shape2").Top = .Height
      .Controls("shape2").Width = rpt_ebao.Width
      .Controls("shape6").Left = 0
      .Controls("shape6").Top = 0
      .Controls("label2").Left = 60
      .Controls("label4").Left = 60 + .Controls("label2").Left + .Controls("label2").Width
      .Controls("label5").Left = 60 + .Controls("label4").Left + .Controls("label4").Width
      .Controls("label6").Left = 60 + .Controls("label5").Left + .Controls("label5").Width
      .Controls("label7").Left = 60 + .Controls("label6").Left + .Controls("label6").Width
      .Controls("label8").Left = 60 + .Controls("label7").Left + .Controls("label7").Width
      .Controls("label9").Left = 60 + .Controls("label8").Left + .Controls("label8").Width
      .Controls("label10").Left = 60 + .Controls("label9").Left + .Controls("label9").Width
      .Controls("line1").Left = .Controls("label2").Left + .Controls("label2").Width
      .Controls("line2").Left = .Controls("label4").Left + .Controls("label4").Width
      .Controls("line4").Left = .Controls("label5").Left + .Controls("label5").Width
      .Controls("line5").Left = .Controls("label6").Left + .Controls("label6").Width
      .Controls("line6").Left = .Controls("label7").Left + .Controls("label7").Width
      .Controls("line7").Left = .Controls("label8").Left + .Controls("label8").Width
      .Controls("line8").Left = .Controls("label9").Left + .Controls("label9").Width
      .Controls("shape5").Left = rpt_ebao.Width
   End With
   With .Sections("section6")
      .Controls("shape13").Left = 0
      .Controls("shape14").Left = rpt_ebao.Sections("section2").Controls("shape5").Left
   End With
   With .Sections("section1")
      .Controls("text1").Width = rpt_ebao.Width * 8 / 100 - 50
      .Controls("text2").Width = rpt_ebao.Width * 9 / 100 - 50
      .Controls("text3").Width = rpt_ebao.Width * 10 / 100 - 50
      .Controls("text4").Width = rpt_ebao.Width * 18 / 100 - 50
      .Controls("text5").Width = rpt_ebao.Width * 18 / 100 - 50
      .Controls("text6").Width = rpt_ebao.Width * 15 / 100 - 50
      .Controls("text7").Width = rpt_ebao.Width * 12 / 100 - 50
      .Controls("text8").Width = rpt_ebao.Width * 10 / 100 - 50
      .Controls("line3").Left = 0
      .Controls("line3").Top = 0
      .Controls("line3").Width = rpt_ebao.Width
      .Controls("shape8").Left = 0
      .Controls("shape8").Top = 0
      .Controls("text1").Left = 60
      .Controls("text1").Top = 0
      .Controls("text2").Left = 60 + .Controls("text1").Left + .Controls("text1").Width
      .Controls("text2").Top = 0
      .Controls("text3").Left = 60 + .Controls("text2").Left + .Controls("text2").Width
      .Controls("text3").Top = 0
      .Controls("text4").Left = 60 + .Controls("text3").Left + .Controls("text3").Width
      .Controls("text4").Top = 0
      .Controls("text5").Left = 60 + .Controls("text4").Left + .Controls("text4").Width
      .Controls("text5").Top = 0
      .Controls("text6").Left = 60 + .Controls("text5").Left + .Controls("text5").Width
      .Controls("text6").Top = 0
      .Controls("text7").Left = 60 + .Controls("text6").Left + .Controls("text6").Width
      .Controls("text7").Top = 0
      .Controls("text8").Left = 60 + .Controls("text7").Left + .Controls("text7").Width
      .Controls("text8").Top = 0
      .Controls("line9").Left = .Controls("text1").Left + .Controls("text1").Width
      .Controls("line10").Left = .Controls("text2").Left + .Controls("text2").Width
      .Controls("line11").Left = .Controls("text3").Left + .Controls("text3").Width
      .Controls("line12").Left = .Controls("text4").Left + .Controls("text4").Width
      .Controls("line13").Left = .Controls("text5").Left + .Controls("text5").Width
      .Controls("line14").Left = .Controls("text6").Left + .Controls("text6").Width
      .Controls("line15").Left = .Controls("text7").Left + .Controls("text7").Width
      .Controls("shape7").Left = rpt_ebao.Width
   End With
   With .Sections("section7")
      .Controls("shape12").Top = 0
      .Controls("line17").Width = rpt_ebao.Sections("section2").Controls("shape5").Left
      .Controls("line28").Width = rpt_ebao.Sections("section2").Controls("shape5").Left
      .Controls("label18").Left = rpt_ebao.Sections("section2").Controls("label5").Left
      .Controls("line16").Left = rpt_ebao.Sections("section2").Controls("line4").Left
      .Controls("line18").Left = rpt_ebao.Sections("section2").Controls("line5").Left
      .Controls("line19").Left = rpt_ebao.Sections("section2").Controls("line6").Left
      .Controls("line20").Left = rpt_ebao.Sections("section2").Controls("line7").Left
      .Controls("line21").Left = rpt_ebao.Sections("section2").Controls("line8").Left
      .Controls("function1").Width = rpt_ebao.Sections("section2").Controls("label6").Width
      .Controls("function2").Width = rpt_ebao.Sections("section2").Controls("label7").Width
      .Controls("function3").Width = rpt_ebao.Sections("section2").Controls("label8").Width
      .Controls("function4").Width = rpt_ebao.Sections("section2").Controls("label9").Width
      .Controls("function5").Width = rpt_ebao.Sections("section2").Controls("label10").Width
      .Controls("function1").Left = rpt_ebao.Sections("section2").Controls("label6").Left
      .Controls("function2").Left = rpt_ebao.Sections("section2").Controls("label7").Left
      .Controls("function3").Left = rpt_ebao.Sections("section2").Controls("label8").Left
      .Controls("function4").Left = rpt_ebao.Sections("section2").Controls("label9").Left
      .Controls("function5").Left = rpt_ebao.Sections("section2").Controls("label10").Left
      .Controls("shape1").Left = rpt_ebao.Sections("section2").Controls("shape5").Left
   End With
   With .Sections("section5")
       .Controls("shape10").Left = 0
       .Controls("shape11").Left = rpt_ebao.Sections("section2").Controls("shape5").Left
       .Controls("shape4").Width = rpt_ebao.Sections("section2").Controls("shape5").Left
       .Controls("shape9").Width = rpt_ebao.Sections("section2").Controls("shape5").Left
       .Controls("label12").Left = rpt_ebao.Sections("section2").Controls("label5").Left
       .Controls("label13").Left = rpt_ebao.Sections("section2").Controls("label5").Left
       .Controls("line23").Left = rpt_ebao.Sections("section2").Controls("line4").Left
       .Controls("line24").Left = rpt_ebao.Sections("section2").Controls("line5").Left
       .Controls("line25").Left = rpt_ebao.Sections("section2").Controls("line6").Left
       .Controls("line26").Left = rpt_ebao.Sections("section2").Controls("line7").Left
       .Controls("line27").Left = rpt_ebao.Sections("section2").Controls("line8").Left
       .Controls("line22").Left = .Controls("line23").Left
       .Controls("line22").Width = rpt_ebao.Width - (rpt_ebao.Sections("section2").Controls("line4").Left)
       .Controls("function6").Width = rpt_ebao.Sections("section2").Controls("label6").Width
       .Controls("function7").Width = rpt_ebao.Sections("section2").Controls("label7").Width
       .Controls("function8").Width = rpt_ebao.Sections("section2").Controls("label8").Width
       .Controls("function9").Width = rpt_ebao.Sections("section2").Controls("label9").Width
       .Controls("function10").Width = rpt_ebao.Sections("section2").Controls("label10").Width
       .Controls("function11").Width = rpt_ebao.Sections("section2").Controls("label9").Width
       .Controls("function6").Left = rpt_ebao.Sections("section2").Controls("label6").Left
       .Controls("function7").Left = rpt_ebao.Sections("section2").Controls("label7").Left
       .Controls("function8").Left = rpt_ebao.Sections("section2").Controls("label8").Left
       .Controls("function9").Left = rpt_ebao.Sections("section2").Controls("label9").Left
       .Controls("function10").Left = rpt_ebao.Sections("section2").Controls("label10").Left
       .Controls("function11").Left = rpt_ebao.Sections("section2").Controls("label9").Left
   End With
End With
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

