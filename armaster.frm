VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form armaster 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9735
   ClientLeft      =   30
   ClientTop       =   30
   ClientWidth     =   14895
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9735
   ScaleWidth      =   14895
   Begin MSComctlLib.ListView sdata 
      Height          =   2895
      Left            =   4320
      TabIndex        =   14
      Top             =   6480
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   5106
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Area Id"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Area Name "
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   3120
      Width           =   5415
   End
   Begin VB.CommandButton first 
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
      Height          =   540
      Left            =   1440
      MaskColor       =   &H00FFC0C0&
      Picture         =   "armaster.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4320
      UseMaskColor    =   -1  'True
      Width           =   1455
   End
   Begin VB.CommandButton last 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   550
      Left            =   5760
      Picture         =   "armaster.frx":330E
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4320
      Width           =   1455
   End
   Begin VB.CommandButton prev 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   550
      Left            =   2880
      Picture         =   "armaster.frx":8111
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4320
      Width           =   1455
   End
   Begin VB.CommandButton nxt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   550
      Left            =   4320
      Picture         =   "armaster.frx":CFC6
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4320
      Width           =   1455
   End
   Begin VB.Frame smnu 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   2880
      TabIndex        =   5
      Top             =   3840
      Visible         =   0   'False
      Width           =   2775
      Begin VB.CommandButton save 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   0
         Picture         =   "armaster.frx":104B0
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         Width           =   1335
      End
      Begin VB.CommandButton cancel 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Left            =   1320
         Picture         =   "armaster.frx":139A8
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   0
         Width           =   1455
      End
   End
   Begin VB.CommandButton edt 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   550
      Left            =   3600
      Picture         =   "armaster.frx":17159
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton dlt 
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   550
      Left            =   5040
      Picture         =   "armaster.frx":1A5D3
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton add 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   550
      Left            =   2160
      Picture         =   "armaster.frx":1DA9F
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "   Area"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   12
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label cl 
      BackColor       =   &H00808080&
      Caption         =   "  X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11880
      TabIndex        =   1
      Top             =   0
      Width           =   495
   End
   Begin VB.Label title 
      BackColor       =   &H00C0C0C0&
      Caption         =   "                                 AREA MASTER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12375
   End
End
Attribute VB_Name = "armaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim s, gen As Long
Dim sql As String
Private Sub add_Click()
Text1.Locked = False
s = 1
Set rs = Nothing
sql = "select * from area"
rs.Open sql, cn
Do While Not rs.EOF
gen = rs!areaid
rs.MoveNext
Loop
gen = gen + 1
Set rs = Nothing
Text1.Text = Clear
edt.Visible = False
dlt.Visible = False
search.Visible = False
first.Visible = False
last.Visible = False
prev.Visible = False
nxt.Visible = False
smnu.Visible = True
End Sub
Private Sub cancel_Click()
Text1.Locked = True
add.Visible = True
edt.Visible = True
dlt.Visible = True
search.Visible = True
first.Visible = True
last.Visible = True
prev.Visible = True
nxt.Visible = True
smnu.Visible = False
Call rec
End Sub
Private Sub dlt_Click()
rs.Delete
rs.MoveNext
If Text1.Text = "" Then
rs.MoveLast
End If
End Sub

Private Sub edt_Click()
s = 2
Text1.Locked = False
gen = rs!areaid
Set rs = Nothing
add.Visible = False
dlt.Visible = False
search.Visible = False
first.Visible = False
last.Visible = False
prev.Visible = False
nxt.Visible = False
smnu.Visible = True
End Sub
Private Sub first_Click()
rs.MoveFirst
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
If s = 1 Then
   sql = "insert into area(areaid,area_name,adate) values(" & gen & ",'" & Text1.Text & "','" & Date & "')"
   rs.Open sql, cn
   Set rs = Nothing
   Call rec
   rs.MoveLast
End If
If s = 2 Then
    sql = "update area set area_name = '" & Text1.Text & "' where areaid = " & gen
    rs.Open sql, cn
    Set rs = Nothing
    Call rec
    While Not rs!areaid = gen
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
Text1.Locked = True
End Sub
Private Sub cl_Click()
'smnu=frame
If smnu.Visible = True Then
MsgBox "Please Save or Cancel Current Record first", vbInformation + vbOKOnly, "Error"
Else
menu.Enabled = True
Unload Me
End If
End Sub
Private Sub Form_Load()
cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data source=" & App.Path & "\appdata\NA2KB_FFC\image.accdb;Jet OLEDB:Database Password=19012019"
'form apearance
Me.Height = Screen.Height - Screen.Height * 5 / 100 - 450
Me.Width = Screen.Width * 40 / 100
Me.Top = 450
Me.Left = Screen.Width * 20 / 100
Me.Picture = LoadPicture(App.Path & "\appdata\images\back.jpg")
 'command buttons setting
 Call setcmd
'title bar close settings
Call setfmnu
 ' label and textbox settings
 Call setlabel
Call rec
End Sub
Private Sub rec()
 sql = "select * from area"
 rs.Open sql, cn, adOpenDynamic, adLockOptimistic
 Set Text1.DataSource = rs
 Text1.DataField = rs.Fields(1).Name
 rs.MoveFirst
 sdata.ListItems.Clear
 While Not rs.EOF
 Set Item = sdata.ListItems.add(, , rs!areaid)
 Item.SubItems(1) = rs!area_name
 rs.MoveNext
 Wend
 rs.MoveFirst
End Sub
Private Sub setfmnu()
title.Height = 450
title.Width = Me.Width
title.Top = 0
title.Left = 0
cl.Top = 0
cl.Left = Me.Width - 600
cl.Height = title.Height - 50
smnu.Top = 4000
smnu.Left = prev.Left
End Sub
Private Sub setcmd()
add.Top = 700
edt.Top = 700
dlt.Top = 700
add.Left = 1800
edt.Left = add.Left + add.Width
dlt.Left = edt.Left + edt.Width
last.Top = 3200
nxt.Top = 3200
prev.Top = 3200
first.Top = 3200
first.Left = 1900 - add.Width / 2 'edt.Left - edt.Width
prev.Left = first.Left + first.Width 'dlt.Left - dlt.Width
nxt.Left = prev.Left + prev.Width 'search.Left - search.Width
last.Left = nxt.Left + nxt.Width 'Me.Width - 2400
End Sub
Private Sub setlabel()
Text1.Left = add.Left - 460
Text1.Top = 2200
Label1.Top = 1600
Label1.Width = edt.Width
Label1.Left = edt.Left
sdata.Left = 0
sdata.Top = first.Top + 600 + first.Width
sdata.Width = Me.Width
sdata.Height = Me.Height - sdata.Top
sdata.ColumnHeaders(1).Width = sdata.Width * 30 / 100
sdata.ColumnHeaders(2).Width = sdata.Width * 69 / 100
End Sub
''
''
''form shotcuts -->
''
''
Private Sub add_KeyDown(KeyCode As Integer, Shift As Integer)
If Text1.Locked = True Then
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
 
     
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyX Then
    Call cl_Click
  Else
    'nothing
  End If
End If
If Text1.Locked = False Then
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
If Text1.Locked = True Then
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
 
     
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyX Then
    Call cl_Click
  Else
    'nothing
  End If
End If
If Text1.Locked = False Then
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
If Text1.Locked = True Then
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
 
     
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyX Then
    Call cl_Click
  Else
    'nothing
  End If
End If
If Text1.Locked = False Then
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
If Text1.Locked = True Then
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
 
     
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyX Then
    Call cl_Click
  Else
    'nothing
  End If
End If
If Text1.Locked = False Then
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
If Text1.Locked = True Then
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
 
     
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyX Then
    Call cl_Click
  Else
    'nothing
  End If
End If
If Text1.Locked = False Then
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
If Text1.Locked = True Then
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
 
     
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyX Then
    Call cl_Click
  Else
    'nothing
  End If
End If
If Text1.Locked = False Then
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
If Text1.Locked = True Then
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
 
     
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyX Then
    Call cl_Click
  Else
    'nothing
  End If
End If
If Text1.Locked = False Then
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
If Text1.Locked = True Then
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
 
     
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyX Then
    Call cl_Click
  Else
    'nothing
  End If
End If
If Text1.Locked = False Then
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
If Text1.Locked = False Then
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
If Text1.Locked = False Then
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
If Text1.Locked = True Then
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
 
     
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyX Then
    Call cl_Click
  Else
    'nothing
  End If
End If
If Text1.Locked = False Then
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
''form shotcuts --> (End)
''
''
