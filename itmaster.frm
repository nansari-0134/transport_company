VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form itmaster 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   11430
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   15135
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "itmaster.frx":0000
   ScaleHeight     =   11430
   ScaleWidth      =   15135
   ShowInTaskbar   =   0   'False
   Begin VB.Frame sfrm 
      BackColor       =   &H8000000B&
      Height          =   6135
      Left            =   8400
      TabIndex        =   26
      Top             =   1320
      Visible         =   0   'False
      Width           =   5415
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1560
         TabIndex        =   31
         Top             =   600
         Width           =   3735
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Show"
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
         Left            =   4080
         TabIndex        =   27
         Top             =   5520
         Width           =   1215
      End
      Begin MSComctlLib.ListView inames 
         Height          =   4215
         Left            =   1560
         TabIndex        =   28
         Top             =   1200
         Visible         =   0   'False
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   7435
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483626
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Ac. Id "
            Object.Width           =   6588
         EndProperty
      End
      Begin VB.Label byagcan 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
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
         Left            =   5040
         TabIndex        =   32
         Top             =   0
         Width           =   375
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Item Name:"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   30
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label16 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Search "
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   161
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   0
         TabIndex        =   29
         Top             =   0
         Width           =   5415
      End
   End
   Begin MSComctlLib.ListView sdata 
      Height          =   2415
      Left            =   240
      TabIndex        =   25
      Top             =   7080
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   4260
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Item Name "
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Opening Stock"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Closing Stock"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Unit"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Sale Rate "
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Loading Rate"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame smnu 
      BorderStyle     =   0  'None
      Height          =   550
      Left            =   1440
      TabIndex        =   22
      Top             =   5640
      Visible         =   0   'False
      Width           =   2895
      Begin VB.CommandButton cancel 
         BackColor       =   &H80000004&
         Height          =   550
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   0
         Width           =   1455
      End
      Begin VB.CommandButton save 
         BackColor       =   &H80000004&
         Height          =   550
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   0
         Width           =   1455
      End
   End
   Begin VB.CommandButton last 
      BackColor       =   &H80000004&
      Height          =   550
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   6240
      Width           =   1455
   End
   Begin VB.CommandButton nxt 
      BackColor       =   &H80000004&
      Height          =   550
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   6240
      Width           =   1455
   End
   Begin VB.CommandButton prev 
      BackColor       =   &H80000004&
      Height          =   550
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   6240
      Width           =   1455
   End
   Begin VB.CommandButton first 
      BackColor       =   &H80000004&
      Height          =   550
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   6240
      Width           =   1455
   End
   Begin VB.CommandButton search 
      BackColor       =   &H80000004&
      Height          =   550
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton dlt 
      BackColor       =   &H80000004&
      Height          =   550
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton edt 
      BackColor       =   &H80000004&
      Height          =   550
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton add 
      BackColor       =   &H80000004&
      Height          =   550
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
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
      ForeColor       =   &H00000000&
      Height          =   540
      Left            =   7560
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   4080
      Width           =   1575
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
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
      ForeColor       =   &H00000000&
      Height          =   540
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   4800
      Width           =   1575
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
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
      ForeColor       =   &H00000000&
      Height          =   540
      Left            =   7560
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   3360
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
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
      ForeColor       =   &H00000000&
      Height          =   540
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   3240
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
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
      ForeColor       =   &H00000000&
      Height          =   540
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   4080
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
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
      ForeColor       =   &H00000000&
      Height          =   540
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   2640
      Width           =   6975
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sale Rate "
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   5640
      TabIndex        =   12
      Top             =   4080
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Loading Rate :"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   120
      TabIndex        =   10
      Top             =   4800
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Closing Stock "
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   5640
      TabIndex        =   8
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Opening Stock "
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   120
      TabIndex        =   6
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Unit "
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   120
      TabIndex        =   4
      Top             =   4080
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Item Name "
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   120
      TabIndex        =   2
      Top             =   2640
      Width           =   1815
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
      Left            =   11400
      TabIndex        =   1
      Top             =   0
      Width           =   495
   End
   Begin VB.Label title 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Item Master"
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
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12135
   End
End
Attribute VB_Name = "itmaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim k, s, a As Integer
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim n, gen As Integer
Dim rsrch As New ADODB.Recordset
Dim sql, sql1 As String
Private Sub add_Click()
Text1.Locked = False
Text2.Locked = False
Text3.Locked = False
Text4.Locked = False
Text5.Locked = False
Text6.Locked = False
n = 1
a = rs!itemid
rs.MoveLast
gen = rs!itemid + 1
Set rs = Nothing
Text1.Text = Clear
Text2.Text = Clear
Text3.Text = Clear
Text4.Text = Clear
Text5.Text = Clear
Text6.Text = Clear
edt.Visible = False
dlt.Visible = False
search.Visible = False
first.Visible = False
last.Visible = False
prev.Visible = False
nxt.Visible = False
smnu.Visible = True
add.Enabled = False
End Sub
Private Sub byagcan_Click()
sfrm.Visible = False
Text7.Text = ""
inames.Visible = False
End Sub
Private Sub cancel_Click()
If n = 1 Then
Set rs = Nothing
Call rec
While Not rs!itemid = a
rs.MoveNext
Wend
Text1.Locked = True
Text2.Locked = True
Text3.Locked = True
Text4.Locked = True
Text5.Locked = True
Text6.Locked = True
End If
If n = 2 Then
Set rs = Nothing
Call rec
While Not rs!itemid = gen
rs.MoveNext
Wend
Text1.Locked = True
Text2.Locked = True
Text3.Locked = True
Text4.Locked = True
Text5.Locked = True
Text6.Locked = True
End If
add.Visible = True
edt.Visible = True
dlt.Visible = True
search.Visible = True
first.Visible = True
last.Visible = True
prev.Visible = True
nxt.Visible = True
smnu.Visible = False
add.Enabled = True
edt.Enabled = True
End Sub
Private Sub cl_Click()
If save.Visible = True Then
  MsgBox "Please Save or Cancel Record Before Exiting", vbInformation + vbOKOnly, "Error"
Else
  menu.Enabled = True
  Unload Me
End If
End Sub
Private Sub Command3_Click()
rs.MoveFirst
While Not rs.EOF
  If rs!iname = Text7.Text Then
  Call byagcan_Click
    Exit Sub
  End If
  rs.MoveNext
Wend
Call byagcan_Click
End Sub
Private Sub dlt_Click()
rs.Delete
rs.MoveNext
If rs.E2OF Then
rs.MovePrevious
End If
End Sub
Private Sub edt_Click()
n = 2
Text1.Locked = False
Text2.Locked = False
Text3.Locked = False
Text4.Locked = False
Text5.Locked = False
Text6.Locked = False
gen = rs!itemid
Set rs = Nothing
add.Visible = False
dlt.Visible = False
search.Visible = False
first.Visible = False
last.Visible = False
prev.Visible = False
nxt.Visible = False
smnu.Visible = True
edt.Enabled = False
End Sub
Private Sub first_Click()
rs.MoveFirst
End Sub
Private Sub Form_Load()
cn.CursorLocation = adUseClient
cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data source=" & App.Path & "\appdata\NA2KB_FFC\image.accdb;Jet OLEDB:Database Password=19012019"
w = 1
'form apearance
Me.Height = Screen.Height - Screen.Height * 5 / 100 - 450
Me.Width = Screen.Width * 40 / 100
Me.Top = 450
Me.Left = Screen.Width * 20 / 100
Me.BackColor = &HFFC0C0
Me.Picture = LoadPicture(App.Path & "\appdata\images\back.jpg")
'file menus setting
'Call setfmnu
'title bar setting
Call settl
'textbox labe setting
'Call setfrm
'control menu setting
Call setfrm
'command button setting
Call setcmd
Call rec
End Sub
Private Sub rec()
sdata.ListItems.Clear
sql = "select * from item"
rs.Open sql, cn, adOpenDynamic, adLockOptimistic
Set Text1.DataSource = rs
Set Text2.DataSource = rs
Set Text3.DataSource = rs
Set Text4.DataSource = rs
Set Text5.DataSource = rs
Set Text6.DataSource = rs

Text1.DataField = rs.Fields(1).Name
Text2.DataField = rs.Fields(2).Name
Text3.DataField = rs.Fields(3).Name
Text4.DataField = rs.Fields(4).Name
Text5.DataField = rs.Fields(5).Name
Text6.DataField = rs.Fields(6).Name
While Not rs.EOF
   Set Item = sdata.ListItems.add(, , rs!iname)
   Item.SubItems(1) = rs!open_stock
   Item.SubItems(2) = rs!currt_stock
   Item.SubItems(3) = rs!unit
   Item.SubItems(4) = rs!sale_rate
   Item.SubItems(5) = rs!load_rate
   rs.MoveNext
Wend
rs.MoveFirst
End Sub
Private Sub setfrm()
Label1.Top = 1600
Label1.Left = 70
Text1.Left = Label1.Left + Label1.Width + 100
Text1.Top = Label1.Top
Text1.Width = Me.Width - 350 - Label1.Width
Label3.Top = Label1.Top + Label1.Height + 100
Label3.Left = 70
Text3.Top = Label3.Top
Text3.Left = 70 + Label3.Width + 100
Label4.Top = Label3.Top
Label4.Left = Me.Width - 350 - Label4.Width - Text4.Width
Text4.Top = Label3.Top
Text4.Left = Label4.Left + Label4.Width + 100
Label2.Left = 70
Label2.Top = Label3.Top + Label3.Height + 100
Text2.Top = Label2.Top
Text2.Left = 70 + Label2.Width + 100
Label6.Top = Label2.Top
Label6.Left = Label4.Left
Text6.Top = Label6.Top
Text6.Left = Label6.Left + Label6.Width + 100
Label5.Top = Label2.Top + Label2.Height + 100
Label5.Left = 70
Text5.Top = Label5.Top
Text5.Left = Label5.Left + Label5.Width + 100
sdata.Left = 0
sdata.Top = first.Top + first.Height
sdata.Width = Me.Width
sdata.Height = Me.Height - sdata.Top
sdata.ColumnHeaders(1).Width = sdata.Width * 20 / 100
sdata.ColumnHeaders(1).Width = sdata.Width * 15 / 100
sdata.ColumnHeaders(1).Width = sdata.Width * 15 / 100
sdata.ColumnHeaders(1).Width = sdata.Width * 15 / 100
sdata.ColumnHeaders(1).Width = sdata.Width * 20 / 100
sdata.ColumnHeaders(1).Width = sdata.Width * 15 / 100
sfrm.Left = 500
sfrm.Top = 2100
End Sub
Private Sub settl()
title.Width = Me.Width
title.Height = 450
title.Top = 0
title.Left = 0
cl.Top = 0
cl.Left = Me.Width - 495 - 50
End Sub
Private Sub setcmd()
 add.Top = 300 + title.Height
 add.Left = 300
 edt.Top = 300 + title.Height
 edt.Left = add.Left + add.Width
 dlt.Top = 300 + title.Height
 dlt.Left = edt.Left + edt.Width
 search.Top = 300 + title.Height
 search.Left = dlt.Left + dlt.Width
 first.Top = Label5.Top + Label5.Height + 300
 first.Left = 300
 prev.Top = first.Top
 prev.Left = first.Left + first.Width
 nxt.Top = first.Top
 nxt.Left = prev.Left + prev.Width
 last.Top = first.Top
 last.Left = nxt.Left + nxt.Width
 smnu.Top = first.Top
 smnu.Left = Text3.Left
 add.Picture = LoadPicture(App.Path & "\appdata\icons\cmd_add.jpg")
 edt.Picture = LoadPicture(App.Path & "\appdata\icons\cmd_edt.jpg")
 dlt.Picture = LoadPicture(App.Path & "\appdata\icons\cmd_dlt.jpg")
 search.Picture = LoadPicture(App.Path & "\appdata\icons\cmd_search.jpg")
 first.Picture = LoadPicture(App.Path & "\appdata\icons\cmd_first.jpg")
 nxt.Picture = LoadPicture(App.Path & "\appdata\icons\cmd_nxt.jpg")
 prev.Picture = LoadPicture(App.Path & "\appdata\icons\cmd_prev.jpg")
 last.Picture = LoadPicture(App.Path & "\appdata\icons\cmd_last.jpg")
 save.Picture = LoadPicture(App.Path & "\appdata\icons\cmd_save.jpg")
 cancel.Picture = LoadPicture(App.Path & "\appdata\icons\cmd_cancel.jpg")
End Sub
Private Sub Form_Unload(cancel As Integer)
rs.Close
cn.Close
End Sub
Private Sub last_Click()
rs.MoveLast
End Sub
Private Sub nxt_Click()
rs.MoveNext
If Text1.Text = "" Then
rs.MoveLast
End If
End Sub
Private Sub prev_Click()
rs.MovePrevious
If Text1.Text = "" Then
rs.MoveFirst
End If
End Sub
Private Sub save_Click()
If n = 1 Then
sql = "insert into item(itemid,iname,unit,open_stock,currt_stock,sale_rate,load_rate,adate) values(" & gen & ",'" _
& Text1.Text & "','" & Text2.Text & "'," & Val(Text3.Text) & "," & Val(Text4.Text) & "," & Val(Text6.Text) & "," _
& Val(Text5.Text) & ",'" & Date & "')"
rs.Open sql, cn
  Set rs = Nothing
   Call rec
   rs.MoveLast
   End If
If n = 2 Then
sql = "update item set iname= '" & Text1.Text & "',unit='" & Text2.Text & "',open_stock=" & Val(Text3.Text) & ",currt_stock=" & Val(Text4.Text) & ",sale_rate=" & Val(Text6.Text) & ",load_rate=" & Val(Text5.Text) & " where itemid=" & gen
rs.Open sql, cn
 Set rs = Nothing
    Call rec
    While Not rs!itemid = gen
    rs.MoveNext
    Wend
    End If
add.Visible = True
edt.Visible = True
dlt.Visible = True
search.Visible = True
first.Visible = True
last.Visible = True
prev.Visible = True
nxt.Visible = True
smnu.Visible = False
add.Enabled = True
edt.Enabled = True
End Sub
Private Sub search_Click()
sfrm.Visible = True
 sql1 = "select * from item"
     rsrch.Open sql1, cn
     inames.ListItems.Clear
     While Not rsrch.EOF
     Set Item = inames.ListItems.add(, , rsrch!iname)
     rsrch.MoveNext
    Wend
    Set rsrch = Nothing
End Sub
Private Sub Text7_Change()
Set rsrch = Nothing
If Text7.Locked = False Then
 sql1 = "select * from item where iname like '" & Text7.Text & "%'"
     rsrch.Open sql1, cn
     inames.ListItems.Clear
     While Not rsrch.EOF
     Set Item = inames.ListItems.add(, , rsrch!iname)
     rsrch.MoveNext
    Wend
    Set rsrch = Nothing
End If
End Sub
Private Sub text7_GotFocus()
     inames.Visible = True
     inames.Left = Text7.Left
End Sub
Private Sub text7_LostFocus()
inames.Visible = False
End Sub
Private Sub inames_Click()
Text7.Text = inames.SelectedItem
inames.Left = sfrm.Width
End Sub
''
''
''form shotcuts -->
''
''
Private Sub add_KeyDown(KeyCode As Integer, Shift As Integer)
If Text3.Locked = True Then
  If Shift = vbCtrlMask And KeyCode = vbKeyN Then
    Call nxt_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyL Then
    Call last_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyF Then
    Call first_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyP Then
    Call prev_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyA Then
    Call add_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyE Then
    Call edt_Click
  ElseIf KeyCode = vbKeyDelete Then
    Call dlt_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyH Then
    MsgBox "search button clicked"
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyX Then
    Call cl_Click
  Else
    'nothing
  End If
End If
If Text3.Locked = False Then
  If Shift = vbCtrlMask And KeyCode = vbKeyS Then
    Call save_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyC Then
    Call cancel_Click
Else
     'nothing
   End If
End If
End Sub
Private Sub edt_KeyDown(KeyCode As Integer, Shift As Integer)
If Text3.Locked = True Then
  If Shift = vbCtrlMask And KeyCode = vbKeyN Then
    Call nxt_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyL Then
    Call last_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyF Then
    Call first_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyP Then
    Call prev_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyA Then
    Call add_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyE Then
    Call edt_Click
  ElseIf KeyCode = vbKeyDelete Then
    Call dlt_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyH Then
    MsgBox "search button clicked"
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyX Then
    Call cl_Click
  Else
    'nothing
  End If
End If
If Text3.Locked = False Then
  If Shift = vbCtrlMask And KeyCode = vbKeyS Then
    Call save_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyC Then
    Call cancel_Click
Else
     'nothing
   End If
End If
End Sub
Private Sub dlt_KeyDown(KeyCode As Integer, Shift As Integer)
If Text3.Locked = True Then
  If Shift = vbCtrlMask And KeyCode = vbKeyN Then
    Call nxt_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyL Then
    Call last_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyF Then
    Call first_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyP Then
    Call prev_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyA Then
    Call add_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyE Then
    Call edt_Click
  ElseIf KeyCode = vbKeyDelete Then
    Call dlt_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyH Then
    MsgBox "search button clicked"
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyX Then
    Call cl_Click
  Else
    'nothing
  End If
End If
If Text3.Locked = False Then
  If Shift = vbCtrlMask And KeyCode = vbKeyS Then
    Call save_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyC Then
    Call cancel_Click
Else
     'nothing
   End If
End If
End Sub

Private Sub search_KeyDown(KeyCode As Integer, Shift As Integer)
If Text3.Locked = True Then
  If Shift = vbCtrlMask And KeyCode = vbKeyN Then
    Call nxt_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyL Then
    Call last_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyF Then
    Call first_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyP Then
    Call prev_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyA Then
    Call add_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyE Then
    Call edt_Click
  ElseIf KeyCode = vbKeyDelete Then
    Call dlt_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyH Then
    MsgBox "search button clicked"
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyX Then
    Call cl_Click
  Else
    'nothing
  End If
End If
If Text3.Locked = False Then
  If Shift = vbCtrlMask And KeyCode = vbKeyS Then
    Call save_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyC Then
    Call cancel_Click
Else
     'nothing
   End If
End If
End Sub
Private Sub first_KeyDown(KeyCode As Integer, Shift As Integer)
If Text3.Locked = True Then
  If Shift = vbCtrlMask And KeyCode = vbKeyN Then
    Call nxt_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyL Then
    Call last_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyF Then
    Call first_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyP Then
    Call prev_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyA Then
    Call add_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyE Then
    Call edt_Click
  ElseIf KeyCode = vbKeyDelete Then
    Call dlt_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyH Then
    MsgBox "search button clicked"
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyX Then
    Call cl_Click
  Else
    'nothing
  End If
End If
If Text3.Locked = False Then
  If Shift = vbCtrlMask And KeyCode = vbKeyS Then
    Call save_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyC Then
    Call cancel_Click
Else
     'nothing
   End If
End If
End Sub
Private Sub nxt_KeyDown(KeyCode As Integer, Shift As Integer)
If Text3.Locked = True Then
  If Shift = vbCtrlMask And KeyCode = vbKeyN Then
    Call nxt_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyL Then
    Call last_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyF Then
    Call first_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyP Then
    Call prev_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyA Then
    Call add_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyE Then
    Call edt_Click
  ElseIf KeyCode = vbKeyDelete Then
    Call dlt_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyH Then
    MsgBox "search button clicked"
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyX Then
    Call cl_Click
  Else
    'nothing
  End If
End If
If Text3.Locked = False Then
  If Shift = vbCtrlMask And KeyCode = vbKeyS Then
    Call save_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyC Then
    Call cancel_Click
Else
     'nothing
   End If
End If
End Sub
Private Sub prev_KeyDown(KeyCode As Integer, Shift As Integer)
If Text3.Locked = True Then
  If Shift = vbCtrlMask And KeyCode = vbKeyN Then
    Call nxt_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyL Then
    Call last_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyF Then
    Call first_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyP Then
    Call prev_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyA Then
    Call add_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyE Then
    Call edt_Click
  ElseIf KeyCode = vbKeyDelete Then
    Call dlt_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyH Then
    MsgBox "search button clicked"
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyX Then
    Call cl_Click
  Else
    'nothing
  End If
End If
If Text3.Locked = False Then
  If Shift = vbCtrlMask And KeyCode = vbKeyS Then
    Call save_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyC Then
    Call cancel_Click
Else
     'nothing
   End If
End If
End Sub
Private Sub last_KeyDown(KeyCode As Integer, Shift As Integer)
If Text3.Locked = True Then
  If Shift = vbCtrlMask And KeyCode = vbKeyN Then
    Call nxt_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyL Then
    Call last_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyF Then
    Call first_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyP Then
    Call prev_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyA Then
    Call add_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyE Then
    Call edt_Click
  ElseIf KeyCode = vbKeyDelete Then
    Call dlt_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyH Then
    MsgBox "search button clicked"
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyX Then
    Call cl_Click
  Else
    'nothing
  End If
End If
If Text3.Locked = False Then
  If Shift = vbCtrlMask And KeyCode = vbKeyS Then
    Call save_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyC Then
    Call cancel_Click
Else
     'nothing
   End If
End If
End Sub
Private Sub save_KeyDown(KeyCode As Integer, Shift As Integer)
If Text3.Locked = False Then
  If Shift = vbCtrlMask And KeyCode = vbKeyS Then
    Call save_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyC Then
    Call cancel_Click
Else
     'nothing
   End If
End If
End Sub
Private Sub cancel_KeyDown(KeyCode As Integer, Shift As Integer)
If Text3.Locked = False Then
  If Shift = vbCtrlMask And KeyCode = vbKeyS Then
    Call save_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyC Then
    Call cancel_Click
Else
     'nothing
   End If
End If
End Sub
Private Sub text1_KeyDown(KeyCode As Integer, Shift As Integer)
If Text3.Locked = True Then
  If Shift = vbCtrlMask And KeyCode = vbKeyN Then
    Call nxt_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyL Then
    Call last_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyF Then
    Call first_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyP Then
    Call prev_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyA Then
    Call add_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyE Then
    Call edt_Click
  ElseIf KeyCode = vbKeyDelete Then
    Call dlt_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyH Then
    MsgBox "search button clicked"
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyX Then
    Call cl_Click
  Else
    'nothing
  End If
End If
If Text3.Locked = False Then
  If Shift = vbCtrlMask And KeyCode = vbKeyS Then
    Call save_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyC Then
    Call cancel_Click
Else
     'nothing
   End If
End If
End Sub
Private Sub text2_KeyDown(KeyCode As Integer, Shift As Integer)
If Text3.Locked = True Then
  If Shift = vbCtrlMask And KeyCode = vbKeyN Then
    Call nxt_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyL Then
    Call last_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyF Then
    Call first_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyP Then
    Call prev_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyA Then
    Call add_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyE Then
    Call edt_Click
  ElseIf KeyCode = vbKeyDelete Then
    Call dlt_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyH Then
    MsgBox "search button clicked"
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyX Then
    Call cl_Click
  Else
    'nothing
  End If
End If
If Text3.Locked = False Then
  If Shift = vbCtrlMask And KeyCode = vbKeyS Then
    Call save_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyC Then
    Call cancel_Click
Else
     'nothing
   End If
End If
End Sub
Private Sub text3_KeyDown(KeyCode As Integer, Shift As Integer)
If Text3.Locked = True Then
  If Shift = vbCtrlMask And KeyCode = vbKeyN Then
    Call nxt_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyL Then
    Call last_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyF Then
    Call first_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyP Then
    Call prev_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyA Then
    Call add_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyE Then
    Call edt_Click
  ElseIf KeyCode = vbKeyDelete Then
    Call dlt_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyH Then
    MsgBox "search button clicked"
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyX Then
    Call cl_Click
  Else
    'nothing
  End If
End If
If Text3.Locked = False Then
  If Shift = vbCtrlMask And KeyCode = vbKeyS Then
    Call save_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyC Then
    Call cancel_Click
Else
     'nothing
   End If
End If
End Sub
Private Sub text4_KeyDown(KeyCode As Integer, Shift As Integer)
If Text3.Locked = True Then
  If Shift = vbCtrlMask And KeyCode = vbKeyN Then
    Call nxt_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyL Then
    Call last_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyF Then
    Call first_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyP Then
    Call prev_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyA Then
    Call add_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyE Then
    Call edt_Click
  ElseIf KeyCode = vbKeyDelete Then
    Call dlt_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyH Then
    MsgBox "search button clicked"
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyX Then
    Call cl_Click
  Else
    'nothing
  End If
End If
If Text3.Locked = False Then
  If Shift = vbCtrlMask And KeyCode = vbKeyS Then
    Call save_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyC Then
    Call cancel_Click
Else
     'nothing
   End If
End If
End Sub
Private Sub text5_KeyDown(KeyCode As Integer, Shift As Integer)
If Text3.Locked = True Then
  If Shift = vbCtrlMask And KeyCode = vbKeyN Then
    Call nxt_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyL Then
    Call last_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyF Then
    Call first_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyP Then
    Call prev_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyA Then
    Call add_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyE Then
    Call edt_Click
  ElseIf KeyCode = vbKeyDelete Then
    Call dlt_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyH Then
    MsgBox "search button clicked"
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyX Then
    Call cl_Click
  Else
    'nothing
  End If
End If
If Text3.Locked = False Then
  If Shift = vbCtrlMask And KeyCode = vbKeyS Then
    Call save_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyC Then
    Call cancel_Click
Else
     'nothing
   End If
End If
End Sub
Private Sub text6_KeyDown(KeyCode As Integer, Shift As Integer)
If Text3.Locked = True Then
  If Shift = vbCtrlMask And KeyCode = vbKeyN Then
    Call nxt_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyL Then
    Call last_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyF Then
    Call first_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyP Then
    Call prev_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyA Then
    Call add_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyE Then
    Call edt_Click
  ElseIf KeyCode = vbKeyDelete Then
    Call dlt_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyH Then
    MsgBox "search button clicked"
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyX Then
    Call cl_Click
  Else
    'nothing
  End If
End If
If Text3.Locked = False Then
  If Shift = vbCtrlMask And KeyCode = vbKeyS Then
    Call save_Click
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyC Then
    Call cancel_Click
Else
     'nothing
   End If
End If
End Sub
''
''
''form shotcuts -->(End)
''
''
