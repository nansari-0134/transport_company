VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form acmaster 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11400
   ClientLeft      =   120
   ClientTop       =   120
   ClientWidth     =   17925
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "acmaster.frx":0000
   ScaleHeight     =   11400
   ScaleWidth      =   17925
   Begin VB.Frame sbytp 
      BackColor       =   &H8000000B&
      Height          =   6735
      Left            =   480
      TabIndex        =   53
      Top             =   3240
      Visible         =   0   'False
      Width           =   6495
      Begin VB.ComboBox sbytpcm 
         BackColor       =   &H80000003&
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
         ItemData        =   "acmaster.frx":B27F
         Left            =   3000
         List            =   "acmaster.frx":B298
         Style           =   2  'Dropdown List
         TabIndex        =   57
         Top             =   480
         Width           =   3375
      End
      Begin VB.CommandButton bytpsrh 
         Caption         =   "Search"
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
         Left            =   5160
         TabIndex        =   56
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
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
         Left            =   5040
         TabIndex        =   54
         Top             =   6120
         Width           =   1215
      End
      Begin MSComctlLib.ListView bytplv 
         Height          =   4215
         Left            =   240
         TabIndex        =   55
         Top             =   1920
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   7435
         View            =   3
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Ac. Id"
            Object.Width           =   2653
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Name"
            Object.Width           =   2653
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Mobile"
            Object.Width           =   2653
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Address"
            Object.Width           =   2651
         EndProperty
      End
      Begin VB.Label bytpcan 
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
         TabIndex        =   60
         Top             =   0
         Width           =   375
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Choose Account Type :"
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
         TabIndex        =   59
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label Label19 
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
         TabIndex        =   58
         Top             =   0
         Width           =   6495
      End
   End
   Begin VB.Frame sbyid 
      BackColor       =   &H8000000B&
      Height          =   2055
      Left            =   9720
      TabIndex        =   48
      Top             =   7680
      Visible         =   0   'False
      Width           =   6495
      Begin VB.TextBox sbyidtxt 
         BackColor       =   &H80000003&
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
         Left            =   2640
         TabIndex        =   50
         Top             =   720
         Width           =   3615
      End
      Begin VB.CommandButton byidsrh 
         Caption         =   "Search"
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
         Left            =   5040
         TabIndex        =   49
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
         TabIndex        =   62
         Top             =   0
         Width           =   375
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Account Id : "
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
         TabIndex        =   52
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label Label17 
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
         TabIndex        =   51
         Top             =   0
         Width           =   6495
      End
   End
   Begin VB.Frame sbyag 
      BackColor       =   &H8000000B&
      Height          =   6735
      Left            =   9720
      TabIndex        =   41
      Top             =   840
      Visible         =   0   'False
      Width           =   6495
      Begin VB.CommandButton Command2 
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
         Left            =   5040
         TabIndex        =   44
         Top             =   6120
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Search"
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
         Left            =   5160
         TabIndex        =   43
         Top             =   1080
         Width           =   1215
      End
      Begin VB.ComboBox sbyagcm 
         BackColor       =   &H80000003&
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
         ItemData        =   "acmaster.frx":B2E6
         Left            =   2640
         List            =   "acmaster.frx":B2FF
         Style           =   2  'Dropdown List
         TabIndex        =   42
         Top             =   480
         Width           =   3735
      End
      Begin MSComctlLib.ListView byaglv 
         Height          =   4215
         Left            =   240
         TabIndex        =   45
         Top             =   1800
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
            Text            =   "Ac. Id "
            Object.Width           =   2122
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Name"
            Object.Width           =   2122
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Mobile"
            Object.Width           =   2122
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Group"
            Object.Width           =   2122
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Address"
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
         TabIndex        =   61
         Top             =   0
         Width           =   375
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
         TabIndex        =   47
         Top             =   0
         Width           =   6495
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Choose Account Group (Heads) :"
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
         Left            =   240
         TabIndex        =   46
         Top             =   480
         Width           =   2175
      End
   End
   Begin MSComctlLib.ListView lv 
      Height          =   3255
      Left            =   9720
      TabIndex        =   40
      Top             =   4200
      Visible         =   0   'False
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   5741
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
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
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "select"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.ComboBox text2 
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
      Height          =   420
      ItemData        =   "acmaster.frx":B34D
      Left            =   7560
      List            =   "acmaster.frx":B366
      TabIndex        =   39
      Top             =   2160
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
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
      Height          =   420
      ItemData        =   "acmaster.frx":B3B4
      Left            =   5400
      List            =   "acmaster.frx":B3BE
      TabIndex        =   38
      Top             =   8280
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker text6 
      Height          =   495
      Left            =   4680
      TabIndex        =   36
      Top             =   4920
      Width           =   1455
      _ExtentX        =   2566
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
      CalendarBackColor=   -2147483644
      Format          =   112525313
      CurrentDate     =   43589
   End
   Begin VB.Frame smnu 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   8640
      TabIndex        =   33
      Top             =   10680
      Visible         =   0   'False
      Width           =   2775
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
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   0
         Width           =   1455
      End
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
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   0
         Width           =   1335
      End
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
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   10680
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
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   1200
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
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   10680
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
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   10680
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
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   1200
      Width           =   1455
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
      Left            =   2520
      MaskColor       =   &H00FFC0C0&
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   10680
      UseMaskColor    =   -1  'True
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
      Height          =   550
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   1200
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
      Height          =   550
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox Text13 
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   3480
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   24
      Top             =   9000
      Width           =   6255
   End
   Begin VB.TextBox Text12 
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   22
      Top             =   8280
      Width           =   1815
   End
   Begin VB.TextBox Text11 
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   20
      Top             =   7560
      Width           =   6255
   End
   Begin VB.TextBox Text10 
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7200
      TabIndex        =   18
      Top             =   6840
      Width           =   1695
   End
   Begin VB.TextBox Text9 
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   16
      Top             =   6840
      Width           =   1695
   End
   Begin VB.TextBox Text8 
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   14
      Top             =   5640
      Width           =   6255
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   10
      Top             =   4320
      Width           =   6255
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   8
      Top             =   3600
      Width           =   6255
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   7
      Top             =   2880
      Width           =   6255
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   2
      Top             =   2160
      Width           =   1695
   End
   Begin MSComCtl2.DTPicker text7 
      Height          =   495
      Left            =   7080
      TabIndex        =   37
      Top             =   4920
      Width           =   1455
      _ExtentX        =   2566
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
      CalendarBackColor=   -2147483644
      Format          =   112525313
      CurrentDate     =   43589
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Detail"
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
      Left            =   1080
      TabIndex        =   23
      Top             =   9000
      Width           =   2415
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Opening balance"
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
      Left            =   1080
      TabIndex        =   21
      Top             =   8280
      Width           =   2415
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Area"
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
      Left            =   1080
      TabIndex        =   19
      Top             =   7560
      Width           =   2415
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mobile No. "
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
      Left            =   5880
      TabIndex        =   17
      Top             =   6840
      Width           =   1335
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Phone No."
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
      Left            =   1080
      TabIndex        =   15
      Top             =   6840
      Width           =   2415
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Address "
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
      TabIndex        =   13
      Top             =   5640
      Width           =   2415
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " To"
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
      Left            =   6600
      TabIndex        =   12
      Top             =   4920
      Width           =   495
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Valid from"
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
      Left            =   3360
      TabIndex        =   11
      Top             =   4920
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Driver licence No."
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
      Top             =   4320
      Width           =   2415
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Under Account group "
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
      TabIndex        =   6
      Top             =   3600
      Width           =   2415
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Name "
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
      TabIndex        =   5
      Top             =   2880
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Account Type "
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
      Left            =   5520
      TabIndex        =   4
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Account ID "
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
      TabIndex        =   3
      Top             =   2160
      Width           =   1455
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
      Left            =   12360
      TabIndex        =   1
      Top             =   -120
      Width           =   495
   End
   Begin VB.Label title 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "ACCOUNT MASTER"
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
      Left            =   360
      TabIndex        =   0
      Top             =   0
      Width           =   12375
   End
End
Attribute VB_Name = "acmaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim s, gen As String 'to check it is edit or add button
Dim sql As String  'to store sql statement
Dim n, i As Integer ' to check account type
Dim a As String ' to store present record account id
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim srh As String
Dim rsrch As New ADODB.Recordset 'for searching record
Dim getgrp As New ADODB.Recordset  'for getting grp code in ac grp search
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
s = 2
a = rs!acid
rs.MoveLast
gen = Val(rs!acid) + 1
Set rs = Nothing
Call entext
Text1.Text = gen
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Text12.Text = ""
Text13.Text = ""
Text6.Value = Date
Text7.Value = Date
Combo1.Text = ""
End Sub
Private Sub entext()
Text2.Locked = False
Text3.Locked = False
Text4.Locked = False
Text5.Locked = False
Text6.Enabled = False
Text7.Enabled = False
Text8.Locked = False
Text9.Locked = False
Text10.Locked = False
Text11.Locked = False
Text12.Locked = False
Text13.Locked = False
Combo1.Locked = False
Text2.ForeColor = vbBlack
Text3.ForeColor = vbBlack
Text4.ForeColor = vbBlack
Text5.ForeColor = vbBlack
Text6.CalendarForeColor = vbBlack
Text7.CalendarForeColor = vbBlack
Text8.ForeColor = vbBlack
Text9.ForeColor = vbBlack
Text10.ForeColor = vbBlack
Text11.ForeColor = vbBlack
Text12.ForeColor = vbBlack
Text13.ForeColor = vbBlack
End Sub
Private Sub dstext()
Text1.Locked = True
Text2.Locked = True
Text3.Locked = True
Text4.Locked = True
Text5.Locked = True
Text6.Enabled = True
Text7.Enabled = True
Text8.Locked = True
Text9.Locked = True
Text10.Locked = True
Text11.Locked = True
Text12.Locked = True
Text13.Locked = True
Combo1.Locked = True
Text2.ForeColor = vbBlack
Text3.ForeColor = vbBlack
Text4.ForeColor = vbBlack
Text5.ForeColor = vbBlack
Text6.CalendarForeColor = vbBlack
Text7.CalendarForeColor = vbBlack
Text8.ForeColor = vbBlack
Text9.ForeColor = vbBlack
Text10.ForeColor = vbBlack
Text11.ForeColor = vbBlack
Text12.ForeColor = vbBlack
Text13.ForeColor = vbBlack
End Sub
Private Sub byidsrh_Click()
Dim res As String
Dim cnd As Boolean
cnd = True
Set rsrch = Nothing
sql = "select * from accm where acid = " & sbyidtxt.Text
rsrch.Open sql, cn
If rsrch.EOF = False And rsrch.BOF = False Then
 res = MsgBox("Account Id : " & rsrch!acid & vbNewLine _
   & "Account Holder Name : " & rsrch!Name & vbNewLine _
   & "Under Account Group : " & rsrch!grp & vbNewLine _
   & "Address : " & rsrch!address & vbNewLine _
   & "Mobile : " & rsrch!mobile & vbNewLine & vbNewLine _
   & "Do You Want to See Full Record ?", vbYesNo, "Search Result")
 If res = vbYes Then
    rs.MoveFirst
    While Not rs!acid = rsrch!acid And cnd = True
       rs.MoveNext
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
Set rsrch = Nothing
End Sub
Private Sub bytpsrh_Click()
Dim tp As Integer
If sbytpcm.Text = "Customer" Then
tp = 1
ElseIf sbytpcm.Text = "Khadan" Then
tp = 2
ElseIf sbytpcm.Text = "Driver" Then
tp = 3
ElseIf sbytpcm.Text = "Purchase party" Then
tp = 4
ElseIf sbytpcm.Text = "General" Then
tp = 5
ElseIf sbytpcm.Text = "Vehicle" Then
tp = 6
ElseIf sbytpcm.Text = "Stock Godown" Then
tp = 7
Else
 Exit Sub
End If
bytplv.ListItems.Clear
Set rsrch = Nothing
sql = "select * from accm where acctype = " & tp
rsrch.Open sql, cn
If rsrch.RecordCount >= -1 Then
  If rsrch.EOF = False And rsrch.BOF = False Then
    While Not rsrch.EOF
       Set Item = bytplv.ListItems.add(, , rsrch!acid)
       Item.SubItems(1) = CStr(rsrch!Name)
       Item.SubItems(2) = IIf(IsNull(rsrch!mobile), "", rsrch!mobile)
       Item.SubItems(3) = IIf(IsNull(rsrch!address), "", rsrch!address)
       rsrch.MoveNext
    Wend
  End If
Else
 MsgBox "No Records found ", vbOKOnly + vbInformation, "Search Result"
 sbytpcm.Text = ""
End If
Set rsrch = Nothing
End Sub
Private Sub cancel_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Text12.Text = ""
Text13.Text = ""
Text6.Value = Date
Text7.Value = Date
If s = 1 Then
    Set rs = Nothing
    Call rec
    While Not rs!acid = gen
     rs.MoveNext
    Wend
    Call dstext
End If
If s = 2 Then
    Set rs = Nothing
    Call rec
    While Not rs!acid = a
     rs.MoveNext
    Wend
    Call dstext
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
Call dstext
End Sub
Private Sub cl_Click()
If smnu.Visible = True Then
MsgBox "Please Save or Cancel Current Record first", vbInformation + vbOKOnly, "Error"
Else
menu.Enabled = True
Unload Me
End If
End Sub
Private Sub Command1_Click()
For i = 1 To bytplv.ListItems.Count
  If bytplv.ListItems.Item(i).Checked = True Then
  rs.MoveFirst
     While Not rs.EOF
       If rs!acid = bytplv.ListItems.Item(i) Then
          sbytp.Visible = False
          bytplv.ListItems.Clear
          Exit Sub
       End If
       rs.MoveNext
     Wend
  End If
Next
sbytp.Visible = False
bytplv.ListItems.Clear
End Sub
Private Sub Command2_Click()
'' search by groups ka show button
For i = 1 To byaglv.ListItems.Count
  If byaglv.ListItems.Item(i).Checked = True Then
  rs.MoveFirst
     While Not rs.EOF
       If rs!acid = byaglv.ListItems.Item(i) Then
          sbyag.Visible = False
          byaglv.ListItems.Clear
          Exit Sub
       End If
       rs.MoveNext
     Wend
  End If
Next
sbyag.Visible = False
byaglv.ListItems.Clear
End Sub
Private Sub Command3_Click()
''search button for search by accout groups
Dim res As String
Dim cnd As Boolean
cnd = True
byaglv.ListItems.Clear
Set getgrp = Nothing
res = "select grpcode from grp where grpname = '" & sbyagcm.Text & "'"
getgrp.Open res, cn
Set rsrch = Nothing
sql = "select * from accm where grp = " & getgrp!grpcode
rsrch.Open sql, cn
If rsrch.RecordCount >= -1 Then
  If rsrch.EOF = False And rsrch.BOF = False Then
    While Not rsrch.EOF
       Set getgrp = Nothing
       res = "select grpname from grp where grpcode = " & rsrch!grp
       getgrp.Open res, cn
       Set Item = byaglv.ListItems.add(, , rsrch!acid)
       Item.SubItems(1) = CStr(rsrch!Name)
       Item.SubItems(2) = IIf(IsNull(rsrch!mobile), "", rsrch!mobile)
       Item.SubItems(3) = CStr(getgrp!grpname)
       Item.SubItems(4) = IIf(IsNull(rsrch!address), "", rsrch!address)
       rsrch.MoveNext
       Set getgrp = Nothing
    Wend
  End If
Else
 MsgBox "No Records found ", vbOKOnly + vbInformation, "Search Result"
 sbyagcm.Text = ""
End If
Set rsrch = Nothing
End Sub
Private Sub dlt_Click()
s = MsgBox("Do You Want To Delete The Current Record", vbOKCancel + vbInformation, "Warning")
If s = vbOK Then
rs.Delete
If rs(0) = 1 Then
rs.MoveNext
If Text1.Text = "" Then
rs.MoveLast
End If
End If
End If
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
s = 1
gen = rs!acid
Set rs = Nothing
Call entext
End Sub
Private Sub first_Click()
On Error GoTo label:
rs.MoveFirst
  Set Text6.DataSource = rs
 Text6.DataField = rs.Fields(5).Name
  Set Text7.DataSource = rs
 Text7.DataField = rs.Fields(6).Name
  Set Text8.DataSource = rs
    If rs!acctype = 1 Then
       Text2.Text = "Customer"
  End If
  If rs!acctype = 2 Then
       Text2.Text = "Khadan"
  End If
    If rs!acctype = 3 Then
    Text2.Text = "Driver"
  End If
    If rs!acctype = 4 Then
     Text2.Text = "Purchase Party"
  End If
    If rs!acctype = 5 Then
     Text2.Text = "General"
  End If
    If rs!acctype = 6 Then
      Text2.Text = "Vehicle"
  End If
    If rs!acctype = 7 Then
     Text2.Text = "Stock Godown"
  End If
Exit Sub
label:
  Set Text6.DataSource = Nothing
 Text6.DataField = ""
  Set Text7.DataSource = Nothing
 Text7.DataField = ""
   If rs!acctype = 1 Then
       Text2.Text = "Customer"
  End If
  If rs!acctype = 2 Then
       Text2.Text = "Khadan"
  End If
    If rs!acctype = 3 Then
    Text2.Text = "Driver"
  End If
    If rs!acctype = 4 Then
     Text2.Text = "Purchase Party"
  End If
    If rs!acctype = 5 Then
     Text2.Text = "General"
  End If
    If rs!acctype = 6 Then
      Text2.Text = "Vehicle"
  End If
    If rs!acctype = 7 Then
     Text2.Text = "Stock Godown"
  End If
End Sub
Private Sub Form_Load()
'form apearance
cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data source=" & App.Path & "\appdata\NA2KB_FFC\image.accdb;Jet OLEDB:Database Password=19012019"
Me.Height = Screen.Height - Screen.Height * 5 / 100 - 450
Me.Width = Screen.Width * 40 / 100
Me.Top = 450
Me.Left = Screen.Width * 20 / 100
Me.BackColor = &HFFC0C0
Me.Picture = LoadPicture(App.Path & "\appdata\images\back.jpg")
'file menus setting
 Call setfmnu
 ' label and textbox settings
 Call setlabel
 'command buttons setting
 Call setcmd
 Call rec
 Call dstext
 Call adgrp
 End Sub
 Private Sub adgrp()
 lv.Top = Text4.Top + Text4.Height + 20
 lv.Left = Text4.Left
 lv.Width = Text4.Width
 lv.ColumnHeaders(1).Width = lv.Width * 99 / 100
 sql = "select * from grp "
 rs1.Open sql, cn, adOpenDynamic, adLockOptimistic
 While Not rs1.EOF
 Set Item = lv.ListItems.add(, , rs1!grpname)
 rs1.MoveNext
 Wend
 Set rs1 = Nothing
 End Sub
 Private Sub rec()
 sql = "select * from accm"
 rs.Open sql, cn, adOpenDynamic, adLockOptimistic
 Set Text1.DataSource = rs
 Text1.DataField = rs.Fields(0).Name
  If rs!acctype = 1 Then
       Text2.Text = "Customer"
  End If
  If rs!acctype = 2 Then
       Text2.Text = "Khadan"
  End If
    If rs!acctype = 3 Then
    Text2.Text = "Driver"
  End If
    If rs!acctype = 4 Then
     Text2.Text = "Purchase Party"
  End If
    If rs!acctype = 5 Then
     Text2.Text = "General"
  End If
    If rs!acctype = 6 Then
      Text2.Text = "Vehicle"
  End If
    If rs!acctype = 7 Then
     Text2.Text = "Stock Godown"
  End If
  Set Text3.DataSource = rs
 Text3.DataField = rs.Fields(2).Name
  Set Text4.DataSource = rs
 Text4.DataField = rs.Fields(3).Name
  Set Text5.DataSource = rs
 Text5.DataField = rs.Fields(4).Name
If Text2.Text = "khadan" Then
  Set Text6.DataSource = rs
 Text6.DataField = rs.Fields(5).Name
  Set Text7.DataSource = rs
 Text7.DataField = rs.Fields(6).Name
Else
     Set Text6.DataSource = Nothing
 Text6.DataField = ""
  Set Text7.DataSource = Nothing
 Text7.DataField = ""
End If
  Set Text8.DataSource = rs
 Text8.DataField = rs.Fields(7).Name
  Set Text9.DataSource = rs
 Text9.DataField = rs.Fields(8).Name
  Set Text10.DataSource = rs
 Text10.DataField = rs.Fields(9).Name
  Set Text11.DataSource = rs
 Text11.DataField = rs.Fields(10).Name
  Set Text12.DataSource = rs
 Text12.DataField = rs.Fields(11).Name
  Set Combo1.DataSource = rs
 Combo1.DataField = rs.Fields(12).Name
  Set Text13.DataSource = rs
 Text13.DataField = rs.Fields(13).Name
 rs.MoveFirst
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
 first.Top = Me.Height - first.Height - 350
 first.Left = 300
 prev.Top = Me.Height - first.Height - 350
 prev.Left = first.Left + first.Width
 nxt.Top = Me.Height - first.Height - 350
 nxt.Left = prev.Left + prev.Width
 last.Top = Me.Height - first.Height - 350
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
Private Sub setfmnu()
title.Width = Me.Width
title.Top = 0
title.Left = 0
cl.Top = 0
cl.Left = Me.Width - 600
cl.Height = title.Height - 50
End Sub
Private Sub setlabel()
Dim a, b As Integer
sbyid.Top = Me.Height / 2 - sbyid.Height / 2
sbyid.Left = Me.Width / 2 - sbyid.Width / 2
sbytp.Top = Me.Height / 2 - sbytp.Height / 2
sbytp.Left = Me.Width / 2 - sbytp.Width / 2
sbyag.Top = Me.Height / 2 - sbyag.Height / 2
sbyag.Left = Me.Width / 2 - sbyag.Width / 2
a = 550 + 1000
b = Label4.Width
Label1.Left = 70
Label1.Width = 1815
Label1.Top = a + 50
Text1.Top = a + 50
Text1.Left = Label1.Left + Label1.Width + 100
Label2.Left = Me.Width - Label2.Width - 350 - Text2.Width
Label2.Top = a + 50
Text2.Left = Label2.Left + Label2.Width + 100
Text2.Top = a + 50
Label3.Width = b
Label3.Left = 70
Label3.Top = a + 250 + Label1.Height + 100
Text3.Left = 70 + b + 100
Text3.Top = a + 250 + Label1.Height + 100
Text3.Width = Me.Width - 350 - Label4.Width
Label4.Left = 70
Label4.Top = a + 250 + Label1.Height + 100 + Label3.Height + 100
Text4.Left = 70 + b + 100
Text4.Top = a + 250 + Label1.Height + 100 + Label3.Height + 100
Text4.Width = Me.Width - 350 - Label4.Width
Label5.Left = 70
Label5.Top = a + 250 + Label1.Height + 100 + Label3.Height + 100 + Label4.Height + 100
Text5.Top = a + 250 + Label1.Height + 100 + Label3.Height + 100 + Label4.Height + 100
Text5.Left = 70 + b + 100
Text5.Width = Me.Width - 350 - Label4.Width
Label6.Top = a + 250 + Label1.Height + 100 + Label3.Height + 100 + Label4.Height + 100 + Label5.Height + 100
Label6.Left = 70 + b + 100
Text6.Top = a + 250 + Label1.Height + 100 + Label3.Height + 100 + Label4.Height + 100 + Label5.Height + 100
Text6.Left = Label6.Left + Label6.Width
Label7.Top = a + 250 + Label1.Height + 100 + Label3.Height + 100 + Label4.Height + 100 + Label5.Height + 100
Text7.Top = a + 250 + Label1.Height + 100 + Label3.Height + 100 + Label4.Height + 100 + Label5.Height + 100
Label7.Left = Me.Width - Label7.Width - 250 - Text7.Width
Text7.Left = Label7.Left + Label7.Width
Label8.Top = a + 250 + Label1.Height + 100 + Label3.Height + 100 + Label4.Height + 100 + Label5.Height + 100 + Label6.Height + 100
Label8.Left = 70
Text8.Top = a + 250 + Label1.Height + 100 + Label3.Height + 100 + Label4.Height + 100 + Label5.Height + 100 + Label6.Height + 100
Text8.Left = 70 + b + 100
Text8.Width = Me.Width - 350 - Label4.Width
Label9.Left = 70
Label9.Top = a + 250 + Label1.Height + 100 + Label3.Height + 100 + Label4.Height + 100 + Label5.Height + 100 + Label6.Height + 100 + Text8.Height + 100
Text9.Left = 70 + b + 100
Text9.Top = a + 250 + Label1.Height + 100 + Label3.Height + 100 + Label4.Height + 100 + Label5.Height + 100 + Label6.Height + 100 + Text8.Height + 100
Label10.Top = a + 250 + Label1.Height + 100 + Label3.Height + 100 + Label4.Height + 100 + Label5.Height + 100 + Label6.Height + 100 + Text8.Height + 100
Label10.Left = Me.Width - Label10.Width - 250 - Text10.Width
Text10.Top = a + 250 + Label1.Height + 100 + Label3.Height + 100 + Label4.Height + 100 + Label5.Height + 100 + Label6.Height + 100 + Text8.Height + 100
Text10.Left = Label10.Left + Label10.Width
Label11.Left = 70
Label11.Top = a + 250 + Label1.Height + 100 + Label3.Height + 100 + Label4.Height + 100 + Label5.Height + 100 + Label6.Height + 100 + Text8.Height + 100 + Label10.Height + 100
Text11.Top = a + 250 + Label1.Height + 100 + Label3.Height + 100 + Label4.Height + 100 + Label5.Height + 100 + Label6.Height + 100 + Text8.Height + 100 + Label10.Height + 100
Text11.Left = 70 + b + 100
Text11.Width = Me.Width - 350 - Label4.Width
Label12.Left = 70
Label12.Top = a + 250 + Label1.Height + 100 + Label3.Height + 100 + Label4.Height + 100 + Label5.Height + 100 + Label6.Height + 100 + Text8.Height + 100 + Label10.Height + 100 + Label11.Height + 100
Text12.Top = a + 250 + Label1.Height + 100 + Label3.Height + 100 + Label4.Height + 100 + Label5.Height + 100 + Label6.Height + 100 + Text8.Height + 100 + Label10.Height + 100 + Label10.Height + 100
Text12.Left = 70 + b + 100
Combo1.Top = Text12.Top
Combo1.Left = Text12.Left + Text12.Width + 150
Label13.Left = 70
Label13.Top = a + 250 + Label1.Height + 100 + Label3.Height + 100 + Label4.Height + 100 + Label5.Height + 100 + Label6.Height + 100 + Text8.Height + 100 + Label10.Height + 100 + Label11.Height + 100 + Label11.Height + 100
Text13.Top = a + 250 + Label1.Height + 100 + Label3.Height + 100 + Label4.Height + 100 + Label5.Height + 100 + Label6.Height + 100 + Text8.Height + 100 + Label10.Height + 100 + Label10.Height + 100 + Label11.Height + 100
Text13.Left = 70 + b + 100
Text13.Width = Me.Width - 350 - Label4.Width
End Sub
Private Sub byidcan_Click()
sbyid.Visible = False
End Sub
Private Sub bytpcan_Click()
sbytp.Visible = False
End Sub
Private Sub byagcan_Click()
sbyag.Visible = False
End Sub
Private Sub Form_Unload(cancel As Integer)
Set rs = Nothing
cn.Close
End Sub

Private Sub last_Click()
On Error GoTo label:
rs.MoveLast
  Set Text6.DataSource = rs
 Text6.DataField = rs.Fields(5).Name
  Set Text7.DataSource = rs
 Text7.DataField = rs.Fields(6).Name
  Set Text8.DataSource = rs
    If rs!acctype = 1 Then
       Text2.Text = "Customer"
  End If
  If rs!acctype = 2 Then
       Text2.Text = "Khadan"
  End If
    If rs!acctype = 3 Then
    Text2.Text = "Driver"
  End If
    If rs!acctype = 4 Then
     Text2.Text = "Purchase Party"
  End If
    If rs!acctype = 5 Then
     Text2.Text = "General"
  End If
    If rs!acctype = 6 Then
      Text2.Text = "Vehicle"
  End If
    If rs!acctype = 7 Then
     Text2.Text = "Stock Godown"
  End If
Exit Sub
label:
  Set Text6.DataSource = Nothing
 Text6.DataField = ""
  Set Text7.DataSource = Nothing
 Text7.DataField = ""
   If rs!acctype = 1 Then
       Text2.Text = "Customer"
  End If
  If rs!acctype = 2 Then
       Text2.Text = "Khadan"
  End If
    If rs!acctype = 3 Then
    Text2.Text = "Driver"
  End If
    If rs!acctype = 4 Then
     Text2.Text = "Purchase Party"
  End If
    If rs!acctype = 5 Then
     Text2.Text = "General"
  End If
    If rs!acctype = 6 Then
      Text2.Text = "Vehicle"
  End If
    If rs!acctype = 7 Then
     Text2.Text = "Stock Godown"
  End If
End Sub
Private Sub k()
If Text4.Text = "" Then
lv.Visible = True
Text4.Text = lv.SelectedItem
Else
lv.Visible = False
End If
End Sub
Private Sub lv_Click()
Text4.Text = lv.SelectedItem
lv.Left = Me.Width
End Sub
Private Sub nxt_Click()
On Error GoTo label:
rs.MoveNext
If Text1.Text = "" Then
rs.MoveLast
End If
  Set Text6.DataSource = rs
 Text6.DataField = rs.Fields(5).Name
  Set Text7.DataSource = rs
 Text7.DataField = rs.Fields(6).Name
  Set Text8.DataSource = rs
    If rs!acctype = 1 Then
       Text2.Text = "Customer"
  End If
  If rs!acctype = 2 Then
       Text2.Text = "Khadan"
  End If
    If rs!acctype = 3 Then
    Text2.Text = "Driver"
  End If
    If rs!acctype = 4 Then
     Text2.Text = "Purchase Party"
  End If
    If rs!acctype = 5 Then
     Text2.Text = "General"
  End If
    If rs!acctype = 6 Then
      Text2.Text = "Vehicle"
  End If
    If rs!acctype = 7 Then
     Text2.Text = "Stock Godown"
  End If
Exit Sub
label:
  Set Text6.DataSource = Nothing
 Text6.DataField = ""
  Set Text7.DataSource = Nothing
 Text7.DataField = ""
   If rs!acctype = 1 Then
       Text2.Text = "Customer"
  End If
  If rs!acctype = 2 Then
       Text2.Text = "Khadan"
  End If
    If rs!acctype = 3 Then
    Text2.Text = "Driver"
  End If
    If rs!acctype = 4 Then
     Text2.Text = "Purchase Party"
  End If
    If rs!acctype = 5 Then
     Text2.Text = "General"
  End If
    If rs!acctype = 6 Then
      Text2.Text = "Vehicle"
  End If
    If rs!acctype = 7 Then
     Text2.Text = "Stock Godown"
  End If
End Sub
Private Sub prev_Click()
On Error GoTo label:
rs.MovePrevious
If Text1.Text = "" Then
rs.MoveFirst
End If
  Set Text6.DataSource = rs
 Text6.DataField = rs.Fields(5).Name
  Set Text7.DataSource = rs
 Text7.DataField = rs.Fields(6).Name
  Set Text8.DataSource = rs
    If rs!acctype = 1 Then
       Text2.Text = "Customer"
  End If
  If rs!acctype = 2 Then
       Text2.Text = "Khadan"
  End If
    If rs!acctype = 3 Then
    Text2.Text = "Driver"
  End If
    If rs!acctype = 4 Then
     Text2.Text = "Purchase Party"
  End If
    If rs!acctype = 5 Then
     Text2.Text = "General"
  End If
    If rs!acctype = 6 Then
      Text2.Text = "Vehicle"
  End If
    If rs!acctype = 7 Then
     Text2.Text = "Stock Godown"
  End If
Exit Sub
label:
  Set Text6.DataSource = Nothing
 Text6.DataField = ""
  Set Text7.DataSource = Nothing
 Text7.DataField = ""
   If rs!acctype = 1 Then
       Text2.Text = "Customer"
  End If
  If rs!acctype = 2 Then
       Text2.Text = "Khadan"
  End If
    If rs!acctype = 3 Then
    Text2.Text = "Driver"
  End If
    If rs!acctype = 4 Then
     Text2.Text = "Purchase Party"
  End If
    If rs!acctype = 5 Then
     Text2.Text = "General"
  End If
    If rs!acctype = 6 Then
      Text2.Text = "Vehicle"
  End If
    If rs!acctype = 7 Then
     Text2.Text = "Stock Godown"
  End If
End Sub
Private Sub save_Click()
     If Text2.Text = "Customer" Then
        n = 1
     End If
     If Text2.Text = "Khadan" Then
        n = 2
     End If
     If Text2.Text = "Driver" Then
        n = 3
     End If
     If Text2.Text = "Purchase Party" Then
        n = 4
     End If
     If Text2.Text = "General" Then
        n = 5
     End If
     If Text2.Text = "Vehicle" Then
        n = 6
     End If
     If Text2.Text = "Stock Godown" Then
        n = 7
     End If
If s = 2 Then
     'add-save
     If Text2.Text = "Driver" Then
     sql = "insert into accm(acid,acctype,name,grp,license,valid_frm,valid_to,address,phone,mobile,area,opn_blnc,drcr,details,adate)" _
           & " values('" & Text1.Text & "'," _
           & n & "," _
           & "'" & Text3.Text & "'," _
           & "'" & Text4.Text & "'," _
           & "'" & "'," _
           & "'" & "'," _
           & "'" & "'," _
           & "'" & Text8.Text & "'," _
           & "'" & Text9.Text & "'," _
           & "'" & Text10.Text & "'," _
           & "'" & Text11.Text & "'," _
           & "'" & Text12.Text & "'," _
           & "'" & Combo1.Text & "'," _
           & "'" & Text13.Text & "'," _
           & "'" & Date & "')"
     Else
     sql = "insert into accm(acid,acctype,name,grp,license,valid_frm,valid_to,address,phone,mobile,area,opn_blnc,drcr,details,adate)" _
           & " values('" & Text1.Text & "'," _
           & n & "," _
           & "'" & Text3.Text & "'," _
           & "'" & Text4.Text & "'," _
           & "'" & Text5.Text & "'," _
           & "'" & Text6.Value & "'," _
           & "'" & Text7.Value & "'," _
           & "'" & Text8.Text & "'," _
           & "'" & Text9.Text & "'," _
           & "'" & Text10.Text & "'," _
           & "'" & Text11.Text & "'," _
           & "'" & Text12.Text & "'," _
           & "'" & Combo1.Text & "'," _
           & "'" & Text13.Text & "'," _
           & "'" & Date & "')"
    End If
     rs.Open sql, cn
     Set rs = Nothing
     Call rec
     rs.MoveLast
     Call dstext
End If
If s = 1 Then
     ' edit -update
     sql = "update accm set acctype = " & n & "," _
           & " name = '" & Text3.Text & "'," _
           & " grp = '" & Text4.Text & "'," _
           & " license = '" & Text5.Text & "'," _
           & " valid_frm = '" & Text6.Value & "'," _
           & " valid_to = '" & Text7.Value & "'," _
           & " address = '" & Text8.Text & "'," _
           & " phone = '" & Text9.Text & "'," _
           & " mobile = '" & Text10.Text & "'," _
           & " area = '" & Text11.Text & "'," _
           & " opn_blnc = '" & Text12.Text & "'," _
           & " drcr = '" & Combo1.Text & "'," _
           & " details = '" & Text13.Text & "'," _
           & " adate = '" & Date & "' where acid = '" & gen & "'"
    rs.Open sql, cn
    Set rs = Nothing
    Call rec
    While Not rs!acid = gen
     rs.MoveNext
    Wend
    Call dstext
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
Call dstext
End Sub
Private Sub sbyidtxt_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbEnter Then
   Call byidsrh_Click
End If
End Sub
Private Sub search_Click()
Dim rs01 As New ADODB.Recordset
n = InputBox("Search With :" & vbNewLine & vbNewLine & "1. Account Id" & vbNewLine & "2. Account Type" & vbNewLine _
    & "3. Account Group(Heads)" _
    & vbNewLine & vbNewLine & vbNewLine & "Enter Your Choice:", "Search", 0)
        If n <> 0 And n <> "" Then
          If n = 1 Then
            sbyid.Visible = True
            sbyidtxt.SetFocus
          ElseIf n = 2 Then
             sbytp.Visible = True
          ElseIf n = 3 Then
             sql = "select * from grp"
             rs01.Open sql, cn, adOpenDynamic, adLockOptimistic
             sbyagcm.Clear
             While Not rs01.EOF
                 sbyagcm.AddItem rs01!grpname
                 rs01.MoveNext
             Wend
             Set rs01 = Nothing
             sbyag.Visible = True
          Else
            MsgBox "Wrong Choice", vbInformation + vbOKOnly, "ERRROR"
            Call search_Click
          End If
        End If
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyDown Then
Text3.SetFocus
End If
End Sub
Private Sub Text2_Click()
If Text2.Text = "Driver" Then
  Text6.Enabled = True
  Text7.Enabled = True
  Text5.Enabled = True
Else
  Text6.Enabled = False
  Text7.Enabled = False
  Text5.Enabled = False
End If
End Sub
Private Sub Text4_Change()
sql = "select * from grp where grpname like '" & Text4.Text & "%'"
rs1.Open sql, cn, adOpenDynamic, adLockOptimistic
lv.ListItems.Clear
While Not rs1.EOF
Set Item = lv.ListItems.add(, , rs1!grpname)
rs1.MoveNext
Wend
Set rs1 = Nothing
End Sub
Private Sub Text4_GotFocus()
If Text4.Locked = False Then
     lv.Visible = True
     lv.Left = Text4.Left
End If
End Sub
Private Sub Text4_LostFocus()
lv.Visible = False
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
    Call search_Click
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
    Call search_Click
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
    Call search_Click
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
    Call search_Click
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
    Call search_Click
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
    Call search_Click
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
   Call search_Click
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
    Call search_Click
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
   Call search_Click
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
    Call search_Click
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
    Call search_Click
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
    Call search_Click
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
    Call search_Click
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
    Call search_Click
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
Private Sub text7_KeyDown(KeyCode As Integer, Shift As Integer)
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
    Call search_Click
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
Private Sub text8_KeyDown(KeyCode As Integer, Shift As Integer)
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
    Call search_Click
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
Private Sub text9_KeyDown(KeyCode As Integer, Shift As Integer)
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
   Call search_Click
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
Private Sub text10_KeyDown(KeyCode As Integer, Shift As Integer)
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
    Call search_Click
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
Private Sub text11_KeyDown(KeyCode As Integer, Shift As Integer)
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
    Call search_Click
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
Private Sub text12_KeyDown(KeyCode As Integer, Shift As Integer)
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
    Call search_Click
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
Private Sub text13_KeyDown(KeyCode As Integer, Shift As Integer)
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
    Call search_Click
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
Private Sub combo1_KeyDown(KeyCode As Integer, Shift As Integer)
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
    Call search_Click
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
''form shotcuts --> (End)
''
''
