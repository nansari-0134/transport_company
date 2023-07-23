VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form tigen 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11025
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   17445
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   11025
   ScaleWidth      =   17445
   Begin VB.Frame sbypt 
      BackColor       =   &H8000000B&
      Height          =   6735
      Left            =   600
      TabIndex        =   57
      Top             =   840
      Visible         =   0   'False
      Width           =   6495
      Begin VB.CommandButton Command2 
         Caption         =   "Show"
         Height          =   495
         Left            =   5040
         TabIndex        =   60
         Top             =   6120
         Width           =   1215
      End
      Begin VB.TextBox Text22 
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
         Left            =   2520
         TabIndex        =   59
         Top             =   480
         Width           =   3735
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Search"
         Height          =   495
         Left            =   960
         TabIndex        =   58
         Top             =   1080
         Width           =   1215
      End
      Begin MSComctlLib.ListView inames 
         Height          =   4215
         Left            =   2520
         TabIndex        =   61
         Top             =   1080
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
      Begin MSComctlLib.ListView byptlv 
         Height          =   4215
         Left            =   240
         TabIndex        =   62
         Top             =   1680
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   7435
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Invoice No"
            Object.Width           =   2122
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Date"
            Object.Width           =   2122
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Party Name"
            Object.Width           =   2122
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Amount"
            Object.Width           =   2122
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Narratn"
            Object.Width           =   2122
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
         Left            =   6120
         TabIndex        =   65
         Top             =   0
         Width           =   375
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Party Name :"
         Height          =   735
         Left            =   240
         TabIndex        =   64
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label25 
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
         TabIndex        =   63
         Top             =   0
         Width           =   6495
      End
   End
   Begin VB.Frame sbyid 
      BackColor       =   &H8000000B&
      Height          =   2055
      Left            =   6120
      TabIndex        =   51
      Top             =   4920
      Visible         =   0   'False
      Width           =   6495
      Begin VB.TextBox sbyidtxt 
         BackColor       =   &H80000003&
         Height          =   495
         Left            =   2640
         TabIndex        =   53
         Top             =   720
         Width           =   3615
      End
      Begin VB.CommandButton byidsrh 
         Caption         =   "Search"
         Height          =   495
         Left            =   5040
         TabIndex        =   52
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label byidcan 
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
         Left            =   6120
         TabIndex        =   56
         Top             =   0
         Width           =   375
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Entry No. : "
         Height          =   495
         Left            =   240
         TabIndex        =   55
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label Label24 
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
         TabIndex        =   54
         Top             =   0
         Width           =   6495
      End
   End
   Begin MSComctlLib.ListView pname 
      Height          =   3015
      Left            =   15480
      TabIndex        =   49
      Top             =   1680
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   5318
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
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
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "name"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton refresh1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   12600
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   4200
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton remove 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   12600
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   3360
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton selectall 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   12600
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   2640
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton invoice 
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
      Left            =   360
      Picture         =   "tigen.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   8400
      Width           =   1455
   End
   Begin VB.Frame smnu 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   42
      Top             =   8280
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
         Picture         =   "tigen.frx":3B3B
         Style           =   1  'Graphical
         TabIndex        =   44
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
         Picture         =   "tigen.frx":7033
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   0
         Width           =   1455
      End
   End
   Begin VB.TextBox Text17 
      Height          =   450
      Left            =   9840
      Locked          =   -1  'True
      TabIndex        =   41
      Top             =   8400
      Width           =   2535
   End
   Begin VB.TextBox Text16 
      Height          =   450
      Left            =   9840
      Locked          =   -1  'True
      TabIndex        =   40
      Top             =   9000
      Width           =   2535
   End
   Begin VB.TextBox Text15 
      Height          =   450
      Left            =   12480
      Locked          =   -1  'True
      TabIndex        =   39
      Top             =   8400
      Width           =   3375
   End
   Begin VB.TextBox Text14 
      Height          =   450
      Left            =   12480
      Locked          =   -1  'True
      TabIndex        =   38
      Top             =   9000
      Width           =   3375
   End
   Begin VB.TextBox Text13 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12480
      Locked          =   -1  'True
      TabIndex        =   37
      Top             =   9600
      Width           =   3375
   End
   Begin VB.TextBox Text12 
      Height          =   450
      Left            =   9840
      Locked          =   -1  'True
      TabIndex        =   36
      Top             =   7800
      Width           =   2535
   End
   Begin VB.TextBox Text11 
      Height          =   450
      Left            =   12480
      Locked          =   -1  'True
      TabIndex        =   35
      Top             =   6000
      Width           =   3375
   End
   Begin VB.TextBox Text10 
      Height          =   450
      Left            =   12480
      Locked          =   -1  'True
      TabIndex        =   34
      Top             =   6600
      Width           =   3375
   End
   Begin VB.TextBox Text9 
      Height          =   450
      Left            =   12480
      Locked          =   -1  'True
      TabIndex        =   33
      Top             =   7200
      Width           =   3375
   End
   Begin VB.TextBox Text8 
      Height          =   450
      Left            =   12480
      Locked          =   -1  'True
      TabIndex        =   32
      Top             =   7800
      Width           =   3375
   End
   Begin VB.TextBox Text7 
      Height          =   450
      Left            =   12480
      Locked          =   -1  'True
      TabIndex        =   31
      Top             =   5400
      Width           =   3375
   End
   Begin VB.TextBox Text6 
      Height          =   450
      Left            =   9840
      Locked          =   -1  'True
      TabIndex        =   25
      Top             =   6600
      Width           =   2535
   End
   Begin VB.TextBox Text5 
      Height          =   450
      Left            =   11040
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   6000
      Width           =   1335
   End
   Begin VB.TextBox Text4 
      Height          =   1815
      Left            =   2040
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   20
      Top             =   6120
      Width           =   5055
   End
   Begin VB.TextBox Text3 
      Height          =   450
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   5400
      Width           =   1935
   End
   Begin MSComctlLib.ListView lv 
      Height          =   2775
      Left            =   240
      TabIndex        =   16
      Top             =   2520
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   4895
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Trip No"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Date"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Vehicle No."
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Slip No."
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Item name"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Site"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Qty"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Sale rate"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Total amt."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Driver name"
         Object.Width           =   3528
      EndProperty
   End
   Begin MSComCtl2.DTPicker dte 
      Height          =   450
      Left            =   6600
      TabIndex        =   15
      Top             =   1320
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   794
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
      Format          =   112984065
      CurrentDate     =   43581
   End
   Begin VB.TextBox Text2 
      Height          =   450
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   1920
      Width           =   6135
   End
   Begin VB.TextBox Text1 
      Height          =   450
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   1320
      Width           =   1935
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
      Height          =   525
      Left            =   7080
      MaskColor       =   &H00FFC0C0&
      Picture         =   "tigen.frx":A7E4
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   480
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
      Height          =   525
      Left            =   11400
      Picture         =   "tigen.frx":DAF2
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   480
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
      Height          =   525
      Left            =   8520
      Picture         =   "tigen.frx":128F5
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   480
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
      Height          =   525
      Left            =   9960
      Picture         =   "tigen.frx":177AA
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton search 
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
      Height          =   525
      Left            =   5160
      Picture         =   "tigen.frx":1AC94
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   480
      Width           =   1455
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
      Height          =   525
      Left            =   2280
      Picture         =   "tigen.frx":1E234
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   480
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
      Height          =   525
      Left            =   3720
      Picture         =   "tigen.frx":216AE
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   480
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
      Height          =   525
      Left            =   840
      Picture         =   "tigen.frx":24B7A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   480
      Width           =   1455
   End
   Begin MSComctlLib.ListView dlv 
      Height          =   3015
      Left            =   4320
      TabIndex        =   50
      Top             =   8280
      Visible         =   0   'False
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   5318
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
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
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "name"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "FINAL BALANCE"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10200
      TabIndex        =   30
      Top             =   9600
      Width           =   2175
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Receipt 2"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   7920
      TabIndex        =   29
      Top             =   9000
      Width           =   1695
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Receipt 1"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   7920
      TabIndex        =   28
      Top             =   8400
      Width           =   1695
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Old balance"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   7920
      TabIndex        =   27
      Top             =   7800
      Width           =   1695
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Invoice Amount"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   9840
      TabIndex        =   26
      Top             =   7200
      Width           =   2535
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Royalty % "
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   7920
      TabIndex        =   24
      Top             =   6600
      Width           =   1695
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Vat %"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   9840
      TabIndex        =   22
      Top             =   6000
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Gross Amount"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   9840
      TabIndex        =   21
      Top             =   5400
      Width           =   2535
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Narration"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   360
      TabIndex        =   19
      Top             =   6120
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Total Trips"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   360
      TabIndex        =   17
      Top             =   5400
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  Date"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   5160
      TabIndex        =   13
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Party name"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   240
      TabIndex        =   12
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Invoice  no."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   240
      TabIndex        =   10
      Top             =   1320
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
      ForeColor       =   &H00000000&
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
      Caption         =   "Trip Invoice Generation"
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
Attribute VB_Name = "tigen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim rsrch As New ADODB.Recordset
Dim sql As String
Dim ino As Integer
Dim n, gen As Integer
Dim msg As Integer
Dim a As Integer 'to store list count for edit option

Private Sub byidsrh_Click()
Dim res As String
Dim cnd As Boolean
Dim eno As Long
cnd = True
Set rsrch = Nothing
sql = "select * from trinv where invno = " & sbyidtxt.Text
rsrch.Open sql, cn
If rsrch.EOF = False And rsrch.BOF = False Then
 res = MsgBox("Invoice No : " & rsrch!invno & vbNewLine _
   & "Date : " & rsrch!adate & vbNewLine _
   & "Party Name : " & rsrch!partyname & vbNewLine _
   & "Net Amount  : " & rsrch!netamt & vbNewLine _
   & "Narration : " & rsrch!narratn & vbNewLine & vbNewLine _
   & "Do You Want to See Full Record ?", vbYesNo, "Search Result")
eno = rsrch!invno
 If res = vbYes Then
    rs.MoveFirst
    Call disp
    While Not rs!invno = eno And cnd = True
       Call nxt_Click
       If rs.EOF Then
          cnd = False
       End If
    Wend
    sbyid.Visible = False
 End If
Else
 MsgBox "No Records found ", vbOKOnly + vbInformation, "Search Result"
 sbyidtxt.Text = ""
End If
sbyidtxt.Text = ""
Set rsrch = Nothing
End Sub
Private Sub cl_Click()
If smnu.Visible = True Then
MsgBox "Please Save or Cancel Current Record first", vbInformation + vbOKOnly, "Error"
Else
menu.Enabled = True
Unload Me
End If
End Sub
Private Sub Command2_Click()
For i = 1 To byptlv.ListItems.Count
  If byptlv.ListItems.Item(i).Checked = True Then
  rs.MoveFirst
  Call disp
     While Not rs.EOF
       If rs!invno = byptlv.ListItems.Item(i) Then
          sbypt.Visible = False
          byptlv.ListItems.Clear
          Exit Sub
       End If
       Call nxt_Click
     Wend
  End If
Next
sbypt.Visible = False
byptlv.ListItems.Clear
End Sub
Private Sub Command3_Click()
Set rsrch = Nothing
byptlv.ListItems.Clear
If Text22.Text <> "" Then
 sql = "select * from trinv where partyname = '" & Text22.Text & "'"
 rsrch.Open sql, cn
 While Not rsrch.EOF
    Set Item = byptlv.ListItems.add(, , rsrch!invno)
    Item.SubItems(1) = rsrch!adate
    Item.SubItems(2) = rsrch!partyname
    Item.SubItems(3) = rsrch!netamt
    Item.SubItems(4) = rsrch!narratn
    rsrch.MoveNext
 Wend
End If
Set rsrch = Nothing
End Sub
Private Sub dlt_Click()
Dim rsrmv As New ADODB.Recordset
msg = MsgBox("Would you like to delete current record", vbOKCancel + vbExclamation, "Alert")
If msg = vbOK Then
  For a = lv.ListItems.Count To 1 Step -1
     sql = "update dtentry set invoice = 0 where eno = " & lv.ListItems(ind)
     rsrmv.Open sql, cn
     Set rsrmv = Nothing
  Next
  sql = "delete from trinv where invno = " & Text1.Text
  rsrmv.Open sql, cn
  Set rsrmv = Nothing
End If
Call nxt_Click
End Sub
Private Sub first_Click()
rs.MoveFirst
Call disp
End Sub
Private Sub Form_Load()
'form apearance
cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data source=" & App.Path & "\appdata\NA2KB_FFC\image.accdb;Jet OLEDB:Database Password=19012019"
Me.Height = Screen.Height - Screen.Height * 5 / 100 - 450
Me.Width = Screen.Width * 80 / 100
Me.Top = 450
Me.Left = Screen.Width * 20 / 100
Me.BackColor = vbWhite
'title bar setting
Call settl
Call setlabel
 'command buttons setting
 Call cmd
 Call rec
 'adding partyname in listview
 Call adac
Me.Picture = LoadPicture(App.Path & "\appdata\images\back.jpg")
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
Private Sub add_Click()
edt.Visible = False
dlt.Visible = False
search.Visible = False
first.Visible = False
last.Visible = False
prev.Visible = False
nxt.Visible = False
smnu.Visible = True
add.Enabled = False
invoice.Visible = False
gen = rs!invno
Set rs = Nothing
sql = "select *  from trinv"
rs1.Open sql, cn, 3, 3
rs1.MoveLast
Text1.Text = rs1!invno + 1
Set rs1 = Nothing
Call setadd
Call clrtext
Call entext
n = 1
'putting inital values
Text10.Text = 0
Text11.Text = 0
Text9.Text = 0
Text8.Text = 0
Text15.Text = 0
Text14.Text = 0
End Sub
Private Sub cancel_Click()
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
invoice.Visible = True
Call setnotadd
If n = 1 Then
            'add-cancel
            Set rs = Nothing
            Call rec
            While Not rs!invno = gen
            rs.MoveNext
            Wend
            Call dstext
ElseIf n = 2 Then
            'edit-cancel
            Set rs = Nothing
            Call rec
            While Not rs!invno = gen
            rs.MoveNext
            Wend
            Call dstext
End If
refresh1.Visible = False
dlv.ListItems.Clear
End Sub
Private Sub edt_Click()
add.Visible = False
dlt.Visible = False
search.Visible = False
first.Visible = False
last.Visible = False
prev.Visible = False
nxt.Visible = False
smnu.Visible = True
edt.Enabled = False
invoice.Visible = False
Call entext
Call setadd
n = 2
gen = rs!invno
Set rs = Nothing
Set rs2 = Nothing
sql = "select * from accm where acctype <> 3 and acctype <> 6"
 rs2.Open sql, cn, adOpenDynamic, adLockOptimistic
 While Not rs2.EOF
 Set Item = pname.ListItems.add(, , rs2!Name)
 rs2.MoveNext
 Wend
 Set rs1 = Nothing
 Set rs2 = Nothing
a = lv.ListItems.Count
dlv.ListItems.Clear
End Sub
Private Sub Form_Unload(cancel As Integer)
cn.Close
Set rs = Nothing
Set rs1 = Nothing
Set rs2 = Nothing
End Sub
Private Sub last_Click()
rs.MoveLast
Call disp
End Sub
Private Sub nxt_Click()
rs.MoveNext
If Text1.Text = "" Then
rs.MoveLast
End If
Call disp
End Sub
Private Sub prev_Click()
rs.MovePrevious
If Text1.Text = "" Then
rs.MoveFirst
End If
Call disp
End Sub
Private Sub refresh1_Click()
'Call disp
Call lvitem
End Sub
Private Sub remove_Click()
Dim ind As Integer
Dim rsrmv As New ADODB.Recordset
Dim gramt As Double
gramt = 0
msg = MsgBox("Would You Like to Remove the selected data from this Invoice", vbOKCancel + vbExclamation, "Info")
If msg = vbOK Then
For ind = lv.ListItems.Count To 1 Step -1
  If lv.ListItems.Item(ind) = True Then
    lm = dlv.ListItems.add(, , lv.ListItems.Item(ind))
    lv.ListItems.remove (ind)
  End If
Next
Text3.Text = lv.ListItems.Count
For msg = lv.ListItems.Count To 1 Step -1
  gramt = gramt + lv.ListItems(msg).SubItems(8)
Next
Text7.Text = gramt
End If
End Sub
Private Sub save_Click()
Dim ind As Integer
Dim rsrmv As New ADODB.Recordset
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
invoice.Visible = True
Call dstext
Call setnotadd
If n = 1 Then
 'add-save
   sql = "insert into trinv(invno,adate,partyname,grossamt,vat,royalty,invamt,oldblncdet,oldblnc,rec1det,rec1,rec2det,rec2,netamt,vatamt,ramt,narratn) " _
          & " values(" & Text1.Text & ",'" _
          & dte.Value & "','" _
          & Text2.Text & "'," _
          & Text7.Text & "," _
          & Text5.Text & "," _
          & Text6.Text & "," _
          & Text9.Text & ",'" _
          & Text12.Text & "'," _
          & Text8.Text & ",'" _
          & Text17.Text & "'," _
          & Text15.Text & ",'" _
          & Text16.Text & "'," _
          & Text14.Text & ",'" _
          & Text13.Text & "'," _
          & Text11.Text & "," _
          & Text10.Text & ",'" _
          & Text4.Text & "')"
   rs.Open sql, cn
   Set rs = Nothing
   For i = 1 To lv.ListItems.Count
       sql = "update dtentry set invoice = " & Text1.Text & " where eno = " & lv.ListItems(i)
       rs.Open sql, cn
       Set rs = Nothing
   Next
   Set rs = Nothing
     Call rec
     rs.MoveLast
     Call disp
ElseIf n = 2 Then
 'edit-save
   'update trinv
     sql = " update trinv set adate = '" & dte.Value & "', partyname = '" _
           & Text2.Text & "', grossamt = " _
           & Text7.Text & ", vat = " _
           & Text5.Text & ", royalty = " _
           & Text6.Text & ", invamt = " _
           & Text11.Text & ", oldblncdet = '" _
           & Text12.Text & "', oldblnc = " _
           & Text8.Text & ", rec1det = '" _
           & Text17.Text & "', rec1 = " _
           & Text15.Text & ", rec2det = '" _
           & Text16.Text & "', rec2 = " _
           & Text14.Text & ", netamt = " _
           & Text13.Text & ", vatamt = " _
           & Text11.Text & ", ramt = " _
           & Text10.Text & ", narratn = '" _
           & Text4.Text & "' where invno = " & Text1.Text
     rs.Open sql, cn
     Set rs = Nothing
     For i = dlv.ListItems.Count To 1 Step -1
        sql = "update dtentry set invoice = 0 where eno = " & dlv.ListItems(ind)
        rsrmv.Open sql, cn
        Set rsrmv = Nothing
     Next
     Call rec
     While Not rs!invno = gen
     rs.MoveNext
     Wend
     Call disp
End If
refresh1.Visible = False
dlv.ListItems.Clear
End Sub
Private Sub cmd()
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
invoice.Picture = LoadPicture(App.Path & "\appdata\icons\cmd_invoice.jpg")
selectall.Picture = LoadPicture(App.Path & "\appdata\icons\cmd_select.jpg")
remove.Picture = LoadPicture(App.Path & "\appdata\icons\cmd_remove.jpg")
refresh1.Picture = LoadPicture(App.Path & "\appdata\icons\cmd_refresh.jpg")
add.Top = 525
edt.Top = 525
dlt.Top = 525
search.Top = 525
add.Left = 2400
edt.Left = add.Left + add.Width
dlt.Left = edt.Left + edt.Width
search.Left = dlt.Left + dlt.Width
first.Top = add.Top
prev.Top = first.Top
nxt.Top = first.Top
last.Top = first.Top
first.Left = search.Left + search.Width + 200
prev.Left = first.Left + first.Width
nxt.Left = prev.Left + prev.Width
last.Left = nxt.Left + nxt.Width
invoice.Top = Me.Height - invoice.Height - 350
invoice.Left = 120
End Sub
Private Sub setlabel()
Dim a As Integer
a = 120
Label1.Left = a
Label1.Top = add.Top + add.Height + 300
Label2.Left = a
Label2.Top = Label1.Top + Label1.Height + 100
Text1.Top = Label1.Top
Text2.Top = Label2.Top
Text1.Left = Label1.Left + Label1.Width + 100
Text2.Left = Label2.Left + Label2.Width + 100
'''''''''
lv.Left = a
lv.Height = 3225
lv.Width = Me.Width - 500
lv.Top = Label2.Top + Label2.Height + 200
lv.ColumnHeaders(1).Width = lv.Width * 7 / 100
lv.ColumnHeaders(2).Width = lv.Width * 9 / 100
lv.ColumnHeaders(3).Width = lv.Width * 10 / 100
lv.ColumnHeaders(4).Width = lv.Width * 7 / 100
lv.ColumnHeaders(5).Width = lv.Width * 14 / 100
lv.ColumnHeaders(6).Width = lv.Width * 14 / 100
lv.ColumnHeaders(7).Width = lv.Width * 8 / 100
lv.ColumnHeaders(8).Width = lv.Width * 8 / 100
lv.ColumnHeaders(9).Width = lv.Width * 10 / 100
lv.ColumnHeaders(10).Width = lv.Width * 12 / 100
Label4.Left = a
Label4.Top = lv.Top + lv.Height + 100
Label5.Left = a
Label5.Top = Label4.Top + Label4.Height + 100
Text3.Top = lv.Top + lv.Height + 100
Text3.Left = Label4.Left + Label4.Width + 100
Text4.Top = Label4.Top + Label4.Height + 100
Text4.Left = Label5.Left + Label5.Width + 100
Text7.Top = lv.Top + lv.Height + 100
Text7.Left = lv.Left + lv.Width - Text7.Width
Label6.Top = lv.Top + lv.Height + 100
Label6.Left = Text7.Left - Label6.Width - 100
Text11.Left = Text7.Left
Text10.Left = Text7.Left
Text9.Left = Text7.Left
Text8.Left = Text7.Left
Text15.Left = Text7.Left
Text14.Left = Text7.Left
Text13.Left = Text7.Left
Label7.Left = Label6.Left
Label7.Top = Label6.Top + Label6.Height + 100
Text5.Left = Label7.Left + Label7.Width + 100
Text5.Top = Label7.Top
Text11.Top = Label7.Top
Text6.Left = Label6.Left
Text6.Top = Text5.Top + Text5.Height + 100
Text10.Top = Text6.Top
Label8.Top = Text6.Top
Label8.Left = Text6.Left - Label8.Width - 100
Label9.Left = Label6.Left
Label9.Top = Text6.Top + Text6.Height + 100
Text9.Top = Label9.Top
Text12.Left = Label6.Left
Text12.Top = Label9.Top + Label9.Height + 100
Text8.Top = Text12.Top
Label10.Top = Text12.Top
Text17.Left = Label6.Left
Label10.Left = Text12.Left - Label10.Width - 100
Text16.Left = Label6.Left
Text17.Top = Text12.Top + Text12.Height + 100
Text15.Top = Text17.Top
Label11.Top = Text17.Top
Label11.Left = Text17.Left - Label11.Width - 100
Text16.Top = Text17.Top + Text17.Height + 100
Text14.Top = Text17.Top + Text17.Height + 100
Label12.Top = Text14.Top
Label12.Left = Text16.Left - Label12.Width - 100
Text13.Top = Text14.Top + Text14.Height + 100
Label13.Top = Text13.Top
'''''''''
smnu.Left = Text4.Left
smnu.Top = Text4.Top + Text4.Height + 300
selectall.Top = lv.Top + 500
refresh1.Top = selectall.Top + selectall.Height + 150
remove.Top = refresh1.Top + refresh1.Height + 150
selectall.Left = Me.Width - selectall.Width - 500
refresh1.Left = selectall.Left
remove.Left = refresh1.Left
''
sbypt.Left = Me.Width / 2 - sbypt.Width / 2
sbypt.Top = Me.Height / 2 - sbypt.Height / 2
sbyid.Left = Me.Width / 2 - sbyid.Width / 2
sbyid.Top = Me.Height / 2 - sbyid.Height / 2
 End Sub
Private Sub rec()
sql = "select * from trinv"
rs.Open sql, cn, 3, 3
Set Text1.DataSource = rs
Set dte.DataSource = rs
Set Text2.DataSource = rs
Set Text7.DataSource = rs
Set Text5.DataSource = rs
Set Text11.DataSource = rs
Set Text6.DataSource = rs
Set Text10.DataSource = rs
Set Text9.DataSource = rs
Set Text12.DataSource = rs
Set Text8.DataSource = rs
Set Text17.DataSource = rs
Set Text15.DataSource = rs
Set Text16.DataSource = rs
Set Text14.DataSource = rs
Set Text13.DataSource = rs
Text1.DataField = "invno"
dte.DataField = "adate"
Text2.DataField = "partyname"
Text7.DataField = "grossamt"
Text5.DataField = "vat"
Text11.DataField = "vatamt"
Text6.DataField = "royalty"
Text10.DataField = "ramt"
Text9.DataField = "invamt"
Text12.DataField = "oldblncdet"
Text8.DataField = "oldblnc"
Text17.DataField = "rec1det"
Text15.DataField = "rec1"
Text16.DataField = "rec2det"
Text14.DataField = "rec2"
Text13.DataField = "netamt"
Call disp
End Sub
Private Sub disp()
lv.ListItems.Clear
If Text1.Text <> "" Then
On Error GoTo n:
sql = "select * from dtentry where invoice =" & Text1.Text
rs1.Open sql, cn, adOpenDynamic, adLockOptimistic
While Not rs1.EOF
Set Item = lv.ListItems.add(, , rs1!eno)
Item.SubItems(1) = rs1!edate
Item.SubItems(2) = rs1!vehno
Item.SubItems(3) = rs1!slipno
Item.SubItems(4) = rs1!Item
Item.SubItems(5) = rs1!site
Item.SubItems(6) = rs1!qty
Item.SubItems(7) = rs1!srate
Item.SubItems(8) = rs1!amt
Item.SubItems(9) = rs1!DriverName
rs1.MoveNext
Wend
Set rs1 = Nothing
Text3.Text = lv.ListItems.Count
End If
Exit Sub
n:
End Sub
Private Sub setadd()
lv.Width = Me.Width - 500 - (selectall.Width + 200)
selectall.Visible = True
refresh1.Visible = True
remove.Visible = True
End Sub
Private Sub setnotadd()
lv.Width = Me.Width - 500
selectall.Visible = False
If n <> 2 Then
refresh1.Visible = False
End If
remove.Visible = False
End Sub
Private Sub clrtext()
Text2.Text = Clear
dte.Value = Date
lv.ListItems.Clear
Text3.Text = Clear
Text4.Text = Clear
Text5.Text = Clear
Text6.Text = Clear
Text7.Text = Clear
Text8.Text = Clear
Text9.Text = Clear
Text10.Text = Clear
Text11.Text = Clear
Text12.Text = Clear
Text13.Text = Clear
Text14.Text = Clear
Text15.Text = Clear
Text16.Text = Clear
Text17.Text = Clear

End Sub
Private Sub dstext()
Text2.Locked = True
dte.Enabled = False
Text4.Locked = True
Text5.Locked = True
Text6.Locked = True
Text12.Locked = True
Text8.Locked = True
Text17.Locked = True
Text15.Locked = True
Text16.Locked = True
Text14.Locked = True
End Sub
Private Sub entext()
Text2.Locked = False
dte.Enabled = True
Text4.Locked = False
Text5.Locked = False
Text6.Locked = False
Text12.Locked = False
Text8.Locked = False
Text17.Locked = False
Text15.Locked = False
Text16.Locked = False
Text14.Locked = False
End Sub
Private Sub search_Click()
n = InputBox("Search With :" & vbNewLine & vbNewLine & "1. Invoice Number" & vbNewLine & "2. Party Name" _
    & vbNewLine & vbNewLine & vbNewLine & "Enter Your Choice:", "Search", 0)
        If n <> 0 And n <> "" Then
          If n = 1 Then
            ''Search by entry id
            sbyid.Visible = True
          ElseIf n = 2 Then
            '' Search By party Name
            sbypt.Visible = True
          Else
            MsgBox "Wrong Choice", vbInformation + vbOKOnly, "ERRROR"
            Call search_Click
          End If
        End If
Set rsrch = Nothing
 sql1 = "select * from accm where acctype <> 2 and acctype <> 6"
     rsrch.Open sql1, cn
     inames.ListItems.Clear
     While Not rsrch.EOF
     Set Item = inames.ListItems.add(, , rsrch!Name)
     rsrch.MoveNext
    Wend
    Set rsrch = Nothing
End Sub
Private Sub selectall_Click()
Dim ind As Integer
For ind = 1 To lv.ListItems.Count
  lv.ListItems.Item(ind).Checked = True
Next
End Sub
Private Sub Text14_Change()
If Text14.Text = "" Then
   Text14.Text = 0
End If
If Text2.Locked = False Then
  Text13.Text = CDbl(Text8.Text) + CDbl(Text9.Text) + CDbl(Text15.Text) + CDbl(Text14.Text)
End If
End Sub
Private Sub Text15_Change()
If Text15.Text = "" Then
   Text15.Text = 0
End If
If Text2.Locked = False Then
  Text13.Text = CDbl(Text8.Text) + CDbl(Text9.Text) + CDbl(Text15.Text) + CDbl(Text14.Text)
End If
End Sub
'Private Sub Text2_Change()
'Dim rs01 As New ADODB.Recordset
'Dim ls As ListItem
'If Text2.Locked = False Then
'  lv.ListItems.Clear
'  sql = "select * from dtentry where partyname = '" & Text2.Text & "' and invoice = 0"
'  rs01.Open sql, cn
 ' While Not rs01.EOF
'    ls = lv.ListItems.add(, , rs01!eno)
'    ls.SubItems(1) = rs01!edate
 '   ls.SubItems(2) = rs01!vehno
'    ls.SubItems(3) = rs01!slipno
 '   ls.SubItems(4) = rs01!Item
'    ls.SubItems(5) = rs01!site
'    ls.SubItems(6) = rs01!qty
'    ls.SubItems(7) = rs01!srate
'    ls.SubItems(8) = rs01!amt
'    ls.SubItems(9) = rs01!DriverName
'  Wend
'End If
'End Sub
Private Sub text5_Change()
If Text5.Text = "" Then
 Text5.Text = 0
End If
If Text2.Locked = False Then
  Text11.Text = (CDbl(Text7.Text) * CDbl(Text5.Text) / 100)
  Text9.Text = CDbl(Text10.Text) + CDbl(Text11.Text) + CDbl(Text7.Text)
End If
End Sub
Private Sub text6_Change()
If Text6.Text = "" Then
 Text6.Text = 0
End If
If Text2.Locked = False Then
  Text10.Text = CDbl(Text7.Text) * CDbl(Text6.Text) / 100
  Text9.Text = CDbl(Text10.Text) + CDbl(Text11.Text) + CDbl(Text7.Text)
End If
End Sub
Private Sub Text7_Change()
 If Text2.Locked = False Then
   'calculate
   Text9.Text = CDbl(Text10.Text) + CDbl(Text11.Text) + CDbl(Text7.Text)
 End If
End Sub
Private Sub Text8_Change()
If Text8.Text = "" Then
   Text8.Text = 0
End If
If Text2.Locked = False Then
  Text13.Text = CDbl(Text8.Text) + CDbl(Text9.Text) + CDbl(Text15.Text) + CDbl(Text14.Text)
End If
End Sub
Private Sub pname_Click()
Text2.Text = pname.SelectedItem
pname.Left = Me.Width
Call lvitem
End Sub
Private Sub lvitem()
Dim gramt As Double
gramt = 0
lv.ListItems.Clear
sql = "select * from dtentry where invoice = 0 and partyname = '" & Text2.Text & "' or invoice = " & Text1.Text
rs1.Open sql, cn, adOpenDynamic, adLockOptimistic
While Not rs1.EOF
Set Item = lv.ListItems.add(, , rs1!eno)
Item.SubItems(1) = rs1!edate
Item.SubItems(2) = rs1!vehno
Item.SubItems(3) = rs1!slipno
Item.SubItems(4) = rs1!Item
Item.SubItems(5) = rs1!site
Item.SubItems(6) = rs1!qty
Item.SubItems(7) = rs1!srate
Item.SubItems(8) = rs1!amt
Item.SubItems(9) = rs1!DriverName
rs1.MoveNext
Wend
Set rs1 = Nothing
Text3.Text = lv.ListItems.Count
For msg = lv.ListItems.Count To 1 Step -1
  gramt = gramt + lv.ListItems(msg).SubItems(8)
Next
Text7.Text = gramt
End Sub
Private Sub text2_Change()
If Text2.Locked = False Then
sql = "select * from accm where name like '" & Text2.Text & "%' and acctype <> 3 and acctype <> 6"
rs1.Open sql, cn, adOpenDynamic, adLockOptimistic
pname.ListItems.Clear
While Not rs1.EOF
Set Item = pname.ListItems.add(, , rs1!Name)
rs1.MoveNext
Wend
Set rs1 = Nothing
End If
End Sub
Private Sub text2_GotFocus()
If Text2.Locked = False Then
     pname.Visible = True
     pname.Left = Text2.Left
     pname.Top = Text2.Top + Text2.Height + 20
End If
End Sub
Private Sub text2_LostFocus()
pname.Visible = False
End Sub
Private Sub adac()
 pname.Left = Me.Width
 pname.ColumnHeaders(1).Width = pname.Width * 99 / 100
 sql = "select * from accm where acctype <> 3 and acctype <> 6"
 rs1.Open sql, cn, adOpenDynamic, adLockOptimistic
 While Not rs1.EOF
 Set Item = pname.ListItems.add(, , rs1!Name)
 rs1.MoveNext
 Wend
 Set rs1 = Nothing
 End Sub
Private Sub Text9_Change()
If Text2.Locked = False Then
 Text13.Text = CDbl(Text9.Text) + CDbl(Text8.Text) + CDbl(Text14.Text) + CDbl(Text15.Text)
End If
End Sub
Private Sub Text22_Change()
Set rsrch = Nothing
If Text22.Locked = False Then
 sql1 = "select * from accm where acctype <> 2 and acctype <> 6 and name like '" & Text22.Text & "%'"
     rsrch.Open sql1, cn
     inames.ListItems.Clear
     While Not rsrch.EOF
     Set Item = inames.ListItems.add(, , rsrch!Name)
     rsrch.MoveNext
    Wend
    Set rsrch = Nothing
End If
End Sub
Private Sub text22_GotFocus()
     inames.Visible = True
     inames.Left = Text22.Left
End Sub
Private Sub text22_LostFocus()
inames.Visible = False
End Sub
Private Sub inames_Click()
Text22.Text = inames.SelectedItem
inames.Left = sbypt.Width
End Sub

