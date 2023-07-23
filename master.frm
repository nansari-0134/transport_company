VERSION 5.00
Begin VB.Form acmaster 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   10965
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   12450
   ControlBox      =   0   'False
   FillColor       =   &H80000016&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   10965
   ScaleWidth      =   12450
   ShowInTaskbar   =   0   'False
   Begin VB.Frame cmnu 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   720
      TabIndex        =   32
      Top             =   9720
      Width           =   8175
      Begin VB.CommandButton last 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Caption         =   "&Last"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5280
         Picture         =   "master.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   1575
      End
      Begin VB.CommandButton nxt 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Caption         =   "&Next"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3720
         Picture         =   "master.frx":20D1
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   1575
      End
      Begin VB.CommandButton prev 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Caption         =   "&Previous"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2160
         Picture         =   "master.frx":5A1A
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   1575
      End
      Begin VB.CommandButton first 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Caption         =   "&First"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   600
         Picture         =   "master.frx":935E
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   1575
      End
   End
   Begin VB.TextBox Text15 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1020
      Left            =   1440
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   31
      Top             =   8400
      Width           =   8895
   End
   Begin VB.Frame fmnu 
      BackColor       =   &H00E0E0E0&
      Height          =   9495
      Left            =   10680
      TabIndex        =   30
      Top             =   480
      Width           =   1455
      Begin VB.CommandButton cancel 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   8160
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton save 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   7320
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton search 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Sea&rch"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   0
         Picture         =   "master.frx":B42E
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   2520
         UseMaskColor    =   -1  'True
         Width           =   1455
      End
      Begin VB.CommandButton dlt 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   0
         Picture         =   "master.frx":EDF6
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   1680
         UseMaskColor    =   -1  'True
         Width           =   1455
      End
      Begin VB.CommandButton edt 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "&Edit"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   0
         Picture         =   "master.frx":10E95
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   840
         UseMaskColor    =   -1  'True
         Width           =   1455
      End
      Begin VB.CommandButton add 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "&Add"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   0
         Picture         =   "master.frx":12F7A
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   1455
      End
   End
   Begin VB.TextBox Text14 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   5160
      TabIndex        =   29
      Top             =   7800
      Width           =   1335
   End
   Begin VB.TextBox Text13 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2520
      TabIndex        =   28
      Top             =   7800
      Width           =   2295
   End
   Begin VB.TextBox Text12 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   7800
      TabIndex        =   27
      Top             =   6960
      Width           =   2295
   End
   Begin VB.TextBox Text11 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1920
      TabIndex        =   26
      Top             =   6720
      Width           =   2295
   End
   Begin VB.TextBox Text10 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   7440
      TabIndex        =   25
      Top             =   6120
      Width           =   1815
   End
   Begin VB.TextBox Text9 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1560
      TabIndex        =   24
      Top             =   5880
      Width           =   1815
   End
   Begin VB.TextBox Text8 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1020
      Left            =   1560
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   17
      Top             =   4080
      Width           =   8895
   End
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   7680
      TabIndex        =   15
      Text            =   " "
      Top             =   3480
      Width           =   1815
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3840
      TabIndex        =   13
      Top             =   3480
      Width           =   1815
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3000
      TabIndex        =   11
      Top             =   2760
      Width           =   5175
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3000
      TabIndex        =   9
      Top             =   2160
      Width           =   5175
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1440
      TabIndex        =   7
      Top             =   1320
      Width           =   5175
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   8640
      TabIndex        =   5
      Top             =   480
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1560
      TabIndex        =   3
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Details :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   360
      TabIndex        =   23
      Top             =   8400
      Width           =   855
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Openig Balance :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   600
      TabIndex        =   22
      Top             =   7800
      Width           =   1695
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   6480
      TabIndex        =   21
      Top             =   7080
      Width           =   855
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Phone :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   600
      TabIndex        =   20
      Top             =   6840
      Width           =   855
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Area :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   6480
      TabIndex        =   19
      Top             =   6120
      Width           =   735
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "City :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   480
      TabIndex        =   18
      Top             =   5880
      Width           =   615
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Address :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   240
      TabIndex        =   16
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Valid To :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   6240
      TabIndex        =   14
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Valid From :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   2520
      TabIndex        =   12
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Driver Licence Number :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   240
      TabIndex        =   10
      Top             =   2760
      Width           =   2415
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Under Account Group :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   240
      TabIndex        =   8
      Top             =   2160
      Width           =   2295
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Name :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Account Type :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   6840
      TabIndex        =   4
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Account Id :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label cl 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   325
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
      Caption         =   "Account Master"
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
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12135
   End
End
Attribute VB_Name = "acmaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim h, n As Integer
Private Sub add_Click()
search.Visible = False
dlt.Visible = False
edt.Visible = False
cmnu.Visible = False
save.Visible = True
cancel.Visible = True
add.Enabled = False
End Sub
Private Sub cancel_Click()
cmnu.Visible = True
add.Visible = True
edt.Visible = True
dlt.Visible = True
search.Visible = True
save.Visible = False
cancel.Visible = False
add.Enabled = True
edt.Enabled = True
search.Enabled = True
End Sub
Private Sub cl_Click()
If save.Visible = True Then
  MsgBox "Please Save or Cancel Record Before Exiting", vbInformation + vbOKOnly, "Error"
Else
  menu.Enabled = True
  Unload Me
End If
End Sub

Private Sub edt_Click()
search.Visible = False
dlt.Visible = False
add.Visible = False
cmnu.Visible = False
save.Visible = True
cancel.Visible = True
edt.Enabled = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer) 'shortcut keys
   If Shift = vbCtrlMask And KeyCode = vbKeyX Then
     Call cl_Click
   End If
   If Shift = vbCtrlMask And KeyCode = vbKeyA Then
        'add key
   End If
   If Shift = vbCtrlMask And KeyCode = vbKeyE Then
        'edit key
   End If
   If Shift = vbCtrlMask And KeyCode = vbKeyD Then
        'delete key
   End If
   If Shift = vbCtrlMask And KeyCode = vbKeyR Then
        'search key
   End If
   If Shift = vbCtrlMask And KeyCode = vbKeyS Then
        'save key
   End If
   If Shift = vbCtrlMask And KeyCode = vbKeyC Then
        'cancel key
   End If
   If Shift = vbCtrlMask And KeyCode = vbKeyF Then
        'First key
   End If
   If Shift = vbCtrlMask And KeyCode = vbKeyP Then
        'previous key
   End If
   If Shift = vbCtrlMask And KeyCode = vbKeyN Then
        'exit key
   End If
   If Shift = vbCtrlMask And KeyCode = vbKeyL Then
        'last key
   End If
End Sub
Private Sub Form_Load()
w = 1
'form apearance
Me.Height = Screen.Height - Screen.Height * 5 / 100 - 450
Me.Width = Screen.Width * 40 / 100
Me.Top = 450
Me.Left = Screen.Width * 20 / 100
Me.BackColor = &HFFC0C0
'file menus setting
Call setfmnu
'title bar setting
Call settl
'textbox labe setting
Call setfrm
'control menu setting
Call setcmnu
'icons setting
Call setico
End Sub
Private Sub setico()

End Sub
Private Sub setcmnu()
cmnu.Width = Me.Width - fmnu.Width
cmnu.Height = Me.Height * 10 / 100
cmnu.Top = Me.Height - cmnu.Height
cmnu.Left = 0
n = cmnu.Width / 5
h = cmnu.Height / 2
first.Left = n / 5
prev.Left = n / 5 + first.Width
nxt.Left = n / 5 + prev.Width + first.Width
last.Left = prev.Width + n / 5 + first.Width + nxt.Width
End Sub
Private Sub settl()
title.Width = Me.Width
title.Height = 450
title.Top = 0
title.Left = 0
cl.Top = 0
cl.Height = title.Height - 50
cl.Left = Me.Width - 495
End Sub
Private Sub setfmnu()
fmnu.Top = title.Height + 80
fmnu.Left = Me.Width - fmnu.Width
fmnu.Height = Me.Height - title.Height
fmnu.Width = 1455
add.Top = 0
edt.Top = add.Height
dlt.Top = add.Height + edt.Height
search.Top = add.Height + edt.Height + dlt.Height
cancel.Top = Me.Height - cancel.Height - 545
save.Top = Me.Height - cancel.Height - save.Height - 545
End Sub
Private Sub setfrm()
'positioning by left side
Label1.Top = title.Height + 50
Label1.Left = 50
Text1.Top = title.Height + 50
Text1.Left = Label1.Width + 50
Label2.Top = title.Height + 50
Label2.Left = Me.Width - Text2.Width - 100 - Label2.Width - fmnu.Width
Text2.Top = title.Height + 50
Text2.Left = Me.Width - Text2.Width - 100 - fmnu.Width
Label3.Top = title.Height + 200 + Label2.Height
Label3.Left = 50
Text3.Top = title.Height + 200 + Label2.Height
Text3.Left = Label1.Width + 50
Text3.Width = Me.Width - Label3.Width - 250 - fmnu.Width
Label4.Top = title.Height + 350 + Label1.Height + Label3.Height
Label4.Left = 50
Text4.Top = title.Height + 350 + Label1.Height + Label3.Height
Text4.Left = Label4.Width + 50
Text4.Width = Me.Width - Label4.Width - 250 - fmnu.Width
Label5.Top = title.Height + Label1.Height + Label3.Height + Label4.Height + 500
Label5.Left = 50
Text5.Top = title.Height + Label1.Height + Label3.Height + Label4.Height + 500
Text5.Left = Label5.Width + 50
Text5.Width = Me.Width - Label5.Width - 250 - fmnu.Width
'positioning by right side
Text7.Top = title.Height + Label1.Height + Label3.Height + Label4.Height + 650 + Label5.Height
Text7.Left = Me.Width - 100 - Text7.Width - fmnu.Width
Label7.Left = Me.Width - 150 - Text7.Width - Label7.Width - fmnu.Width
Label7.Top = title.Height + Label1.Height + Label3.Height + Label4.Height + 650 + Label5.Height
Text6.Top = title.Height + Label1.Height + Label3.Height + Label4.Height + 650 + Label5.Height
Text6.Left = Me.Width - 350 - Text7.Width - Label7.Width - Text6.Width - fmnu.Width
Label6.Top = title.Height + Label1.Height + Label3.Height + Label4.Height + 650 + Label5.Height
Label6.Left = Me.Width - 450 - Text7.Width - Label7.Width - Text6.Width - Label6.Width - fmnu.Width
'positioning by left side
Label8.Top = title.Height + Label1.Height + Label3.Height + Label4.Height + 800 + Label5.Height + Label6.Height
Label8.Left = 50
Text8.Left = Label8.Width + 50
Text8.Top = title.Height + Label1.Height + Label3.Height + Label4.Height + 800 + Label5.Height + Label6.Height
Text8.Width = Me.Width - Label8.Width - 250 - fmnu.Width
Label9.Left = 50
Label9.Top = title.Height + Label1.Height + Label3.Height + Label4.Height + 950 + Label5.Height + Label6.Height + Text8.Height
Text9.Left = Label9.Width + 50
Text9.Top = title.Height + Label1.Height + Label3.Height + Label4.Height + 950 + Label5.Height + Label6.Height + Text8.Height
Label10.Left = Me.Width - fmnu.Width - Text10.Width - 50 - Label10.Width
Label10.Top = title.Height + Label1.Height + Label3.Height + Label4.Height + 950 + Label5.Height + Label6.Height + Text8.Height
Text10.Left = Me.Width - fmnu.Width - 50 - Text10.Width
Text10.Top = title.Height + Label1.Height + Label3.Height + Label4.Height + 950 + Label5.Height + Label6.Height + Text8.Height
Text11.Top = title.Height + Label1.Height + Label3.Height + Label4.Height + 1100 + Label5.Height + Label6.Height + Text8.Height + Label9.Height
Text11.Left = Label11.Width + 50
Label11.Top = title.Height + Label1.Height + Label3.Height + Label4.Height + 1100 + Label5.Height + Label6.Height + Text8.Height + Label9.Height
Label11.Left = 50
Text12.Top = title.Height + Label1.Height + Label3.Height + Label4.Height + 1100 + Label5.Height + Label6.Height + Text8.Height + Label9.Height
Text12.Left = Me.Width - fmnu.Width - 50 - Text12.Width
Label12.Top = title.Height + Label1.Height + Label3.Height + Label4.Height + 1100 + Label5.Height + Label6.Height + Text8.Height + Label9.Height
Label12.Left = Me.Width - fmnu.Width - 50 - Text12.Width - Label12.Width
Label13.Top = title.Height + Label1.Height + Label3.Height + Label4.Height + 1250 + Label5.Height + Label6.Height + Text8.Height + Label9.Height + Label11.Height
Label13.Left = 50
Text13.Top = title.Height + Label1.Height + Label3.Height + Label4.Height + 1250 + Label5.Height + Label6.Height + Text8.Height + Label9.Height + Label11.Height
Text13.Left = Label13.Width + 50
Text14.Top = title.Height + Label1.Height + Label3.Height + Label4.Height + 1250 + Label5.Height + Label6.Height + Text8.Height + Label9.Height + Label11.Height
Text14.Left = Label13.Width + 150 + Text13.Width
Label14.Top = title.Height + Label1.Height + Label3.Height + Label4.Height + 1400 + Label5.Height + Label6.Height + Text8.Height + Label9.Height + Label11.Height + Label13.Height
Label14.Left = 50
Text15.Top = title.Height + Label1.Height + Label3.Height + Label4.Height + 1400 + Label5.Height + Label6.Height + Text8.Height + Label9.Height + Label11.Height + Label13.Height
Text15.Left = Label14.Width + 50
Text15.Width = Me.Width - Label14.Width - fmnu.Width - 250
End Sub
Private Sub save_Click()
cmnu.Visible = True
add.Visible = True
edt.Visible = True
dlt.Visible = True
search.Visible = True
save.Visible = False
cancel.Visible = False
add.Enabled = True
edt.Enabled = True
search.Enabled = True
End Sub

