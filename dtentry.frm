VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form dtentry 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11970
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   20235
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   12
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
   ScaleHeight     =   11970
   ScaleWidth      =   20235
   Begin VB.Frame sbypt 
      BackColor       =   &H8000000B&
      Height          =   6735
      Left            =   600
      TabIndex        =   72
      Top             =   2160
      Visible         =   0   'False
      Width           =   6495
      Begin VB.CommandButton Command3 
         Caption         =   "Search"
         Height          =   495
         Left            =   960
         TabIndex        =   78
         Top             =   1080
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
         TabIndex        =   77
         Top             =   480
         Width           =   3735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Show"
         Height          =   495
         Left            =   5040
         TabIndex        =   73
         Top             =   6120
         Width           =   1215
      End
      Begin MSComctlLib.ListView inames 
         Height          =   4215
         Left            =   2520
         TabIndex        =   79
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
         TabIndex        =   80
         Top             =   1680
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Entry Id"
            Object.Width           =   2122
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Date"
            Object.Width           =   2122
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Vehicle No."
            Object.Width           =   2122
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Driver Name"
            Object.Width           =   2122
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Payment"
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
         TabIndex        =   76
         Top             =   0
         Width           =   375
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
         TabIndex        =   75
         Top             =   0
         Width           =   6495
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Party Name :"
         Height          =   735
         Left            =   240
         TabIndex        =   74
         Top             =   480
         Width           =   2175
      End
   End
   Begin VB.Frame sbyid 
      BackColor       =   &H8000000B&
      Height          =   2055
      Left            =   7080
      TabIndex        =   66
      Top             =   2160
      Visible         =   0   'False
      Width           =   6495
      Begin VB.CommandButton byidsrh 
         Caption         =   "Search"
         Height          =   495
         Left            =   5040
         TabIndex        =   68
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox sbyidtxt 
         BackColor       =   &H80000003&
         Height          =   495
         Left            =   2640
         TabIndex        =   67
         Top             =   720
         Width           =   3615
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
         TabIndex        =   71
         Top             =   0
         Width           =   375
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
         TabIndex        =   70
         Top             =   0
         Width           =   6495
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Entry No. : "
         Height          =   495
         Left            =   240
         TabIndex        =   69
         Top             =   720
         Width           =   2175
      End
   End
   Begin MSComctlLib.ListView driver 
      Height          =   1935
      Left            =   17640
      TabIndex        =   65
      Top             =   4320
      Visible         =   0   'False
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   3413
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
         Size            =   11.25
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
   Begin MSComctlLib.ListView khadan 
      Height          =   1455
      Left            =   18240
      TabIndex        =   64
      Top             =   8640
      Visible         =   0   'False
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   2566
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
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
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "name"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView site 
      Height          =   1095
      Left            =   17520
      TabIndex        =   63
      Top             =   7080
      Visible         =   0   'False
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   1931
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
         Size            =   11.25
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
   Begin MSComctlLib.ListView mat 
      Height          =   1695
      Left            =   17520
      TabIndex        =   62
      Top             =   2880
      Visible         =   0   'False
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   2990
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
         Size            =   11.25
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
   Begin MSComctlLib.ListView vehno 
      Height          =   2655
      Left            =   17760
      TabIndex        =   61
      Top             =   1200
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   4683
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
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
         Size            =   11.25
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
   Begin MSComctlLib.ListView pname 
      Height          =   1695
      Left            =   16800
      TabIndex        =   60
      Top             =   360
      Visible         =   0   'False
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   2990
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
         Size            =   11.25
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
   Begin VB.ComboBox unit 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   7200
      TabIndex        =   59
      Text            =   "Units"
      Top             =   3000
      Width           =   1575
   End
   Begin VB.CheckBox ch 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   5880
      TabIndex        =   58
      Top             =   1200
      Width           =   210
   End
   Begin VB.TextBox Text21 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   385
      Left            =   8280
      TabIndex        =   57
      Top             =   7080
      Width           =   1695
   End
   Begin VB.TextBox Text20 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   385
      Left            =   8280
      TabIndex        =   56
      Top             =   6480
      Width           =   1695
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
      Left            =   4560
      MaskColor       =   &H00FFC0C0&
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   8160
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
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   8160
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
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   8160
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
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   8160
      Width           =   1455
   End
   Begin VB.Frame smnu 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   47
      Top             =   8160
      Visible         =   0   'False
      Width           =   2895
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
         Height          =   525
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   0
         Width           =   1455
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
         Height          =   525
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   0
         Width           =   1455
      End
   End
   Begin VB.TextBox Text19 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   385
      Left            =   7920
      TabIndex        =   46
      Top             =   1200
      Width           =   2600
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
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   0
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
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   0
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
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   0
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
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   0
      Width           =   1455
   End
   Begin VB.TextBox Text18 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   385
      Left            =   12360
      TabIndex        =   41
      Top             =   5160
      Width           =   1400
   End
   Begin VB.TextBox Text17 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   385
      Left            =   8400
      TabIndex        =   40
      Top             =   5160
      Width           =   1335
   End
   Begin VB.TextBox Text16 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   385
      Left            =   14640
      TabIndex        =   35
      Top             =   4560
      Width           =   1575
   End
   Begin VB.TextBox Text15 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   420
      Left            =   12120
      Locked          =   -1  'True
      TabIndex        =   34
      Top             =   4560
      Width           =   1215
   End
   Begin VB.TextBox Text14 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   385
      Left            =   9480
      TabIndex        =   33
      Top             =   4560
      Width           =   1400
   End
   Begin VB.TextBox Text13 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   385
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   32
      Top             =   4560
      Width           =   1335
   End
   Begin VB.TextBox Text12 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   27
      Top             =   7080
      Width           =   3615
   End
   Begin VB.TextBox Text11 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   385
      Left            =   2640
      TabIndex        =   26
      Top             =   6480
      Width           =   3615
   End
   Begin VB.TextBox Text10 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   385
      Left            =   2760
      TabIndex        =   25
      Top             =   5760
      Width           =   1695
   End
   Begin VB.TextBox Text9 
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
      Left            =   2760
      TabIndex        =   24
      Top             =   5160
      Width           =   1695
   End
   Begin VB.TextBox Text8 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   385
      Left            =   2760
      TabIndex        =   23
      Top             =   4560
      Width           =   1695
   End
   Begin VB.TextBox Text7 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   385
      Left            =   2760
      TabIndex        =   22
      Top             =   4080
      Width           =   4355
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   385
      Left            =   2640
      TabIndex        =   21
      Top             =   3480
      Width           =   4355
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   385
      Left            =   2760
      TabIndex        =   20
      Top             =   3000
      Width           =   4355
   End
   Begin VB.TextBox Text4 
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
      Left            =   2760
      TabIndex        =   19
      Top             =   2400
      Width           =   4355
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   385
      Left            =   2760
      TabIndex        =   18
      Top             =   1800
      Width           =   4335
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   385
      Left            =   2880
      TabIndex        =   7
      Top             =   1200
      Width           =   2600
   End
   Begin MSComCtl2.DTPicker dtdate 
      Height          =   390
      Left            =   7200
      TabIndex        =   6
      Top             =   600
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   688
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
      Format          =   112787457
      CurrentDate     =   43580
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
      ForeColor       =   &H00000000&
      Height          =   385
      Left            =   2760
      TabIndex        =   5
      Top             =   600
      Width           =   2600
   End
   Begin MSComctlLib.ListView lv 
      Height          =   2775
      Left            =   5040
      TabIndex        =   0
      Top             =   9000
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   4895
      View            =   3
      LabelWrap       =   0   'False
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
      NumItems        =   18
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Entry Id"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Vehicle no."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Party name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Slip no."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "item name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Site"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Khadan"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Loading charge"
         Object.Width           =   2893
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Quantity"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Sale Rate"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Sale amount"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "diesel ltr"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "Diesel amount"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "Km run"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "(Veh + Mtn)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "Expenses"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Text            =   "Date"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Payment"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   6720
      TabIndex        =   55
      Top             =   7080
      Width           =   1095
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Expenses"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   6720
      TabIndex        =   54
      Top             =   6480
      Width           =   1095
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Diesel Amt."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   10440
      TabIndex        =   39
      Top             =   5160
      Width           =   1350
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Diesel Rate"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   6000
      TabIndex        =   38
      Top             =   5160
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
      Height          =   420
      Left            =   11880
      TabIndex        =   37
      Top             =   0
      Width           =   495
   End
   Begin VB.Label title 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   " DAILY TRANSACTION ENTRY"
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
      Left            =   120
      TabIndex        =   36
      Top             =   0
      Width           =   12375
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Amount"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   13560
      TabIndex        =   31
      Top             =   4560
      Width           =   975
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Quantity"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   10920
      TabIndex        =   30
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Total Trips"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   8040
      TabIndex        =   29
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Sale Rate"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   5280
      TabIndex        =   28
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Driver name"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   120
      TabIndex        =   17
      Top             =   6960
      Width           =   2295
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Maintainance  Detail"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   120
      TabIndex        =   16
      Top             =   6480
      Width           =   2295
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " K.M.  Running"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   -120
      TabIndex        =   15
      Top             =   5760
      Width           =   2295
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Diesel litre"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   385
      Left            =   120
      TabIndex        =   14
      Top             =   5280
      Width           =   2295
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Loading Charges"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   385
      Left            =   120
      TabIndex        =   13
      Top             =   4680
      Width           =   2295
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Khadan Name"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   385
      Left            =   240
      TabIndex        =   12
      Top             =   4080
      Width           =   2295
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Site Address"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   385
      Left            =   240
      TabIndex        =   11
      Top             =   3480
      Width           =   2295
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Material"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   385
      Left            =   360
      TabIndex        =   10
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " slip no"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   385
      Left            =   360
      TabIndex        =   9
      Top             =   2400
      Width           =   2295
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Party Name"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   385
      Left            =   360
      TabIndex        =   8
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Not working"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6240
      TabIndex        =   4
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Vehicle No"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   385
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Date"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   2
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Entry No."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   385
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   2295
   End
End
Attribute VB_Name = "dtentry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim n, a, gen As Integer
Dim sql, sql1 As String
Dim rsrch As New ADODB.Recordset
Private Sub add_Click()
Dim rs2 As New ADODB.Recordset
edt.Visible = False
dlt.Visible = False
search.Visible = False
first.Visible = False
last.Visible = False
prev.Visible = False
nxt.Visible = False
smnu.Visible = True
add.Enabled = False
n = 1
a = rs!eno
rs.MoveLast
gen = rs!eno + 1
Set rs = Nothing
Call cltext
Text1.Text = gen
Call entext
Set rs2 = Nothing
sql = "select * from accm where acctype = 2"
rs2.Open sql, cn, adOpenDynamic, adLockOptimistic
While Not rs2.EOF
 Set tp = khadan.ListItems.add(, , rs2!Name)
 rs2.MoveNext
Wend
Set rs2 = Nothing
sql = "select * from accm where acctype = 3"
rs2.Open sql, cn, adOpenDynamic, adLockOptimistic
While Not rs2.EOF
 Set it = driver.ListItems.add(, , rs2!Name)
 rs2.MoveNext
Wend
Set rs2 = Nothing
sql = "select * from accm where acctype = 6"
rs2.Open sql, cn, adOpenDynamic, adLockOptimistic
While Not rs2.EOF
 Set Itp = vehno.ListItems.add(, , rs2!Name)
 rs2.MoveNext
Wend
Set rs2 = Nothing
sql = "select * from accm where acctype = 1 or acctype = 4 or acctype = 5"
rs2.Open sql, cn, adOpenDynamic, adLockOptimistic
While Not rs2.EOF
    Set Itp = pname.ListItems.add(, , rs2!Name)
 rs2.MoveNext
Wend
Set rs2 = Nothing
sql = "select iname from item"
rs2.Open sql, cn, adOpenDynamic, adLockOptimistic
While Not rs2.EOF
    Set Itp = mat.ListItems.add(, , rs2!iname)
 rs2.MoveNext
Wend
Set rs2 = Nothing
End Sub
Private Sub byagcan_Click()
sbypt.Visible = False
Text22.Text = ""
End Sub
Private Sub byidcan_Click()
sbyidtxt.Text = ""
sbyid.Visible = False
End Sub
Private Sub byidsrh_Click()
Dim res As String
Dim cnd As Boolean
Dim eno As Long
cnd = True
Set rsrch = Nothing
sql = "select * from dtentry where eno = " & sbyidtxt.Text
rsrch.Open sql, cn
If rsrch.EOF = False And rsrch.BOF = False Then
 res = MsgBox("Entry No : " & rsrch!eno & vbNewLine _
   & "Date : " & rsrch!edate & vbNewLine _
   & "Vehicle No : " & rsrch!vehno & vbNewLine _
   & "Driver Name  : " & rsrch!DriverName & vbNewLine _
   & "Payment : " & rsrch!payment & vbNewLine & vbNewLine _
   & "Do You Want to See Full Record ?", vbYesNo, "Search Result")
eno = rsrch!eno
 If res = vbYes Then
    rs.MoveFirst
    While Not rs!eno = eno And cnd = True
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
sbyidtxt.Text = ""
Set rsrch = Nothing
End Sub
Private Sub cancel_Click()
Call dstext
Call rec
rs.MoveLast
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
Private Sub ch_Click()
If ch.Value = 1 Then
Text3.Locked = True
Text5.Locked = True
Text6.Locked = True
Text7.Locked = True
Text8.Locked = True
Text9.Locked = True
Text10.Locked = True
Text13.Locked = True
Text14.Locked = True
Text15.Locked = True
Text16.Locked = True
Text17.Locked = True
Text18.Locked = True
End If
End Sub
Private Sub Command2_Click()
For i = 1 To byptlv.ListItems.Count
  If byptlv.ListItems.Item(i).Checked = True Then
  rs.MoveFirst
     While Not rs.EOF
       If rs!eno = byptlv.ListItems.Item(i) Then
          sbypt.Visible = False
          byptlv.ListItems.Clear
          Exit Sub
       End If
       rs.MoveNext
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
 sql = "select * from dtentry where partyname = '" & Text22.Text & "'"
 rsrch.Open sql, cn
 While Not rsrch.EOF
    Set Item = byptlv.ListItems.add(, , rsrch!eno)
    Item.SubItems(1) = rsrch!edate
    Item.SubItems(2) = rsrch!vehno
    Item.SubItems(3) = rsrch!DriverName
    Item.SubItems(4) = rsrch!payment
    rsrch.MoveNext
 Wend
End If
Set rsrch = Nothing
End Sub
Private Sub edt_Click()
n = 2
gen = rs!eno
Set rs = Nothing
Call entext
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
Private Sub Form_Click()
pname.Visible = False
vehno.Visible = False
mat.Visible = False
site.Visible = False
khadan.Visible = False
driver.Visible = False
End Sub
Private Sub Form_Unload(cancel As Integer)
Set rs = Nothing
cn.Close
End Sub
Private Sub save_Click()
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
 If Text20.Text = "" Then
  Text20.Text = 0
 End If
 If Text21.Text = "" Then
   Text21.Text = 0
 End If
 If Text8.Text = "" Then
   Text8.Text = 0
 End If
 If Text18.Text = "" Then
   Text18.Text = 0
 End If
If n = 1 Then
If ch.Value = 0 Then
sql = "insert into dtentry(eno,edate,vehno,notwrk,ntwrkdet,partyname,slipno,item,site,khadan,loadcharg,srate,trips,qty,unit,amt,diesel," _
    & "dieselrate,damt,kmrun,mtdet,mtexp,drivername,payment,invoice,expense) " _
    & "values(" & gen & ",'" _
    & dtdate.Value & "','" & Text2.Text & "','" _
    & ch.Value & "','" _
    & Text19.Text & "','" _
    & Text3.Text & "','" _
    & Text4.Text & "','" _
    & Text5.Text & "','" _
    & Text6.Text & "','" _
    & Text7.Text & "'," _
    & Text8.Text & "," _
    & Text13.Text & "," _
    & Text14.Text & "," _
    & Text15.Text & ",'" _
    & unit.Text & "'," _
    & Text16.Text & "," _
    & Text9.Text & "," _
    & Text17.Text & "," _
    & Text18.Text & "," _
    & Text10.Text & ",'" _
    & Text11.Text & "'," _
    & Text20.Text & ",'" _
    & Text12.Text & "'," _
    & Text21.Text & ",0," & CDbl(Text8.Text) + CDbl(Text18.Text) + CDbl(Text20.Text) + CDbl(Text21.Text) & ")"
Else
sql = "insert into dtentry(eno,edate,vehno,notwrk,ntwrkdet,partyname,slipno,item,site,khadan,loadcharg,srate,trips,qty,unit,amt,diesel," _
    & "dieselrate,damt,kmrun,mtdet,mtexp,drivername,payment,invoice,expense) " _
    & "values('" & gen & "','" _
    & dtdate.Value & "','" & Text2.Text & "','" _
    & ch.Value & "','" _
    & Text19.Text & "','" _
    & "','" _
    & Text4.Text & "','" _
    & "','" & "','" & "'," & "0," & "0," & "0," & "0,'" & "'," & "0," & "0," & "0," & "0," & "0,'" _
    & Text11.Text & "'," _
    & Text20.Text & ",'" _
    & Text12.Text & "'," & Text21.Text & ",0," & CDbl(Text8.Text) + CDbl(Text18.Text) + CDbl(Text20.Text) + CDbl(Text21.Text) & ");"
End If
rs.Open sql, cn, adOpenDynamic, adLockOptimistic
 Set rs = Nothing
    Call rec
    While Not rs!eno = a
    rs.MoveNext
    Wend

End If

If n = 2 Then
rs.Open sql, cn, adOpenDynamic, adLockOptimistic
  Set rs = Nothing
   Call rec
   rs.MoveLast
sql = "update dtentry set edate='" & dtdate.Value & "', vehno='" & Text2.Text & "', notwrk='" _
    & ch.Value & "', ntwrkdet='" _
    & Text19.Text & "', partyname='" _
    & Text3.Text & "', slipno='" _
    & Text4.Text & "', item='" _
    & Text5.Text & "', site='" _
    & Text6.Text & "', kahadan='" _
    & Text7.Text & "', loadcharg=" _
    & Text8.Text & ", srate=" _
    & Text13.Text & ", trips=" _
    & Text14.Text & ", qty=" & Text15.Text & ", unit='" & unit.Text & "', amt=" & Text16.Text & ", diesel=" & Text9.Text & ", dieselrate=" & Text17.Text & ", damt=" & Text18.Text & ", kmrun=" & Text10.Text & ", mtdet='" & Text11.Text & "', mtexp=" _
    & Text20.Text & ", drivername=" _
    & Text12.Text & "', payment=" _
    & Text21.Text & ", expense=" _
    & CDbl(Text8.Text) + CDbl(Text18.Text) + CDbl(Text20.Text) + CDbl(Text21.Text) & " where eno = " & gen
    Set rs = Nothing
rs.Open sql, cn
 Set rs = Nothing
    Call rec
    While Not rs!eno = gen
    rs.MoveNext
    Wend
End If
End Sub

Private Sub cl_Click()
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
Me.Width = Screen.Width * 80 / 100
Me.Top = 450
Me.Left = Screen.Width * 20 / 100
Me.Picture = LoadPicture(App.Path & "\appdata\images\back.jpg")

 'command box settings
Call cmd
'title bar setting
Call settl
' label and textbox settings
 Call setlabel
 Call dstext
 Call rec
 Call setlv
End Sub
Private Sub setlv()
pname.Width = Text3.Width
vehno.Width = Text2.Width
mat.Width = Text5.Width
site.Width = Text6.Width
khadan.Width = Text7.Width
driver.Width = Text12.Width
pname.Left = Text3.Left
vehno.Left = Text2.Left
mat.Left = Text5.Left
site.Left = Text6.Left
khadan.Left = Text7.Left
driver.Left = Text12.Left
pname.Top = Text3.Top + Text3.Height + 20
vehno.Top = Text2.Top + Text2.Height + 20
mat.Top = Text5.Top + Text5.Height + 20
site.Top = Text6.Top + Text6.Height + 20
khadan.Top = Text7.Top + Text7.Height + 20
driver.Top = Text12.Top + Text12.Height + 20
pname.ColumnHeaders(1).Width = pname.Width * 99 / 100
vehno.ColumnHeaders(1).Width = vehno.Width * 99 / 100
mat.ColumnHeaders(1).Width = mat.Width * 99 / 100
site.ColumnHeaders(1).Width = site.Width * 99 / 100
khadan.ColumnHeaders(1).Width = khadan.Width * 99 / 100
driver.ColumnHeaders(1).Width = driver.Width * 99 / 100
End Sub
Private Sub rec()
sql = " select * from dtentry"
rs.Open sql, cn, adOpenDynamic, adLockOptimistic
Set Text1.DataSource = rs
Set dtdate.DataSource = rs
Set Text2.DataSource = rs
Set Text3.DataSource = rs
Set Text4.DataSource = rs
Set Text5.DataSource = rs
Set Text6.DataSource = rs
Set Text7.DataSource = rs
Set Text8.DataSource = rs
Set Text9.DataSource = rs
Set Text10.DataSource = rs
Set Text11.DataSource = rs
Set Text12.DataSource = rs
Set Text13.DataSource = rs
Set Text14.DataSource = rs
Set Text15.DataSource = rs
Set Text16.DataSource = rs
Set Text17.DataSource = rs
Set Text18.DataSource = rs
Set Text19.DataSource = rs
Set Text20.DataSource = rs
Set Text21.DataSource = rs
Set unit.DataSource = rs
Text1.DataField = "eno"
dtdate.DataField = "edate"
Text2.DataField = "vehno"
Text19.DataField = "ntwrkdet"
Text3.DataField = "partyname"
Text4.DataField = "slipno"
Text5.DataField = "item"
Text6.DataField = "site"
Text7.DataField = "khadan"
Text8.DataField = "loadcharg"
Text13.DataField = "srate"
Text14.DataField = "trips"
Text15.DataField = "qty"
Text16.DataField = "amt"
Text9.DataField = "diesel"
Text17.DataField = "dieselrate"
Text18.DataField = "damt"
Text10.DataField = "kmrun"
Text11.DataField = "mtdet"
Text20.DataField = "mtexp"
Text12.DataField = "drivername"
Text21.DataField = "payment"
unit.DataField = "unit"
End Sub
Private Sub entext()
Text2.Locked = False
Text3.Locked = False
Text4.Locked = False
Text5.Locked = False
Text6.Locked = False
Text7.Locked = False
Text8.Locked = False
Text9.Locked = False
Text10.Locked = False
Text11.Locked = False
Text12.Locked = False
Text13.Locked = False
Text14.Locked = False
Text15.Locked = False
Text16.Locked = False
Text17.Locked = False
Text18.Locked = False
Text19.Locked = False
Text20.Locked = False
Text21.Locked = False
End Sub
Private Sub dstext()
Text1.Locked = True
Text2.Locked = True
Text3.Locked = True
Text4.Locked = True
Text5.Locked = True
Text6.Locked = True
Text7.Locked = True
Text8.Locked = True
Text9.Locked = True
Text10.Locked = True
Text11.Locked = True
Text12.Locked = True
Text13.Locked = True
Text14.Locked = True
Text15.Locked = True
Text16.Locked = True
Text17.Locked = True
Text18.Locked = True
Text19.Locked = True
Text20.Locked = True
Text21.Locked = True
End Sub
Private Sub cltext()
Text1.Text = Clear
Text2.Text = Clear
Text3.Text = Clear
Text4.Text = Clear
Text5.Text = Clear
Text6.Text = Clear
Text7.Text = Clear
Text8.Text = Clear
Text9.Text = Clear
Text10.Text = Clear
Text11.Text = Clear
Text11.Text = Clear
Text12.Text = Clear
Text13.Text = Clear
Text14.Text = Clear
Text15.Text = Clear
Text16.Text = Clear
Text17.Text = Clear
Text18.Text = Clear
Text19.Text = Clear
Text20.Text = Clear
Text21.Text = Clear
unit.Text = "Units"
End Sub
Private Sub settl()
title.Width = Me.Width
title.Height = 450
title.Top = 0
title.Left = 0
cl.Top = 0
cl.Left = Me.Width - 495 - 50
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = vbCtrlMask And KeyCode = vbKeyX Then
     Call cl_Click
   End If
End Sub
Private Sub cmd()
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
Private Sub setlabel()
Dim a, b As Integer
a = 120
Label1.Top = add.Height + 400 + add.Top
Label1.Left = a
Text1.Top = add.Height + 400 + add.Top
Label2.Top = Label1.Top

Label3.Left = a
Label3.Top = Label1.Top + Label1.Height + 100
Label5.Top = Label3.Top + Label3.Height + 100
Label5.Left = a
Label6.Left = a
Label6.Top = Label5.Top + Label5.Height + 100
Label7.Left = a
Label7.Top = Label6.Top + Label6.Height + 100
Label8.Left = a
Label8.Top = Label7.Top + Label7.Height + 100
Label9.Left = a
Label9.Top = Label8.Top + Label8.Height + 100
Label10.Left = a
Label10.Top = Label9.Top + Label9.Height + 100
Label11.Left = a
Label11.Top = Label10.Top + Label10.Height + 100
Label12.Left = a
Label12.Top = Label11.Top + Label11.Height + 100
Label13.Left = a
Label13.Top = Label12.Top + Label12.Height + 100
Label14.Left = a
Label14.Top = Label13.Top + Label13.Height + 100
Text1.Left = Label1.Left + Label1.Width + 150
Text2.Left = Text1.Left
Text3.Left = Text1.Left
Text4.Left = Text1.Left
Text5.Left = Text1.Left
Text6.Left = Text1.Left
Text7.Left = Text1.Left
Text8.Left = Text1.Left
Text9.Left = Text1.Left
Text10.Left = Text1.Left
Text11.Left = Text1.Left
Text12.Left = Text1.Left
Text2.Top = Label3.Top
Text3.Top = Label5.Top
Text4.Top = Label6.Top
Text5.Top = Label7.Top
Text6.Top = Label8.Top
Text7.Top = Label9.Top
Text8.Top = Label10.Top
Text9.Top = Label11.Top
Text10.Top = Label12.Top
Text11.Top = Label13.Top
Text12.Top = Label14.Top
Label15.Top = Text8.Top
Label16.Top = Text8.Top
Label17.Top = Text8.Top
Label18.Top = Text8.Top
Text13.Top = Label18.Top
Text14.Top = Label18.Top
Text15.Top = Label18.Top
Text16.Top = Label18.Top
Label19.Top = Text9.Top
Label20.Top = Text9.Top
Text17.Top = Label19.Top
Text18.Top = Label19.Top
Label15.Left = Text8.Left + Text8.Width + 200
Text13.Left = Label15.Left + Label15.Width + 100
Label16.Left = Text13.Left + Text13.Width + 200
Text14.Left = Label16.Left + Label16.Width + 100
Label17.Left = Text14.Left + Text14.Width + 200
Text15.Left = Label17.Left + Label17.Width + 100
Label18.Left = Text15.Left + Text15.Width + 200
Text16.Left = Label18.Left + Label18.Width + 100
Label19.Left = Text9.Left + Text9.Width + 200
Text17.Left = Label19.Left + Label19.Width + 100
Label20.Left = Text17.Left + Text17.Width + 200
Text18.Left = Label20.Left + Label20.Width + 100
ch.Left = Text3.Left + Text3.Width - 355
Label4.Top = Label3.Top
ch.Top = Label4.Top
Label4.Top = Label3.Top
Label4.Left = ch.Left + ch.Width + 150
Label2.Left = ch.Left + ch.Width + 150
dtdate.Left = Label2.Left + Label2.Width + 100
dtdate.Top = Label2.Top
Text19.Top = Label4.Top
Text19.Left = Label4.Left + Label4.Width + 100
Label21.Top = Text11.Top
Text20.Top = Text11.Top
Label21.Left = Text11.Left + Text11.Width + 300
Text20.Left = Label21.Left + Label21.Width + 100
Label22.Top = Text12.Top
Text21.Top = Text12.Top
Label22.Left = Label21.Left
Text21.Left = Text20.Left
smnu.Top = Text21.Top + Text21.Height + 100
unit.Left = Text5.Left + Text5.Width + 200
unit.Top = Text5.Top
'list view settings
lv.Width = Me.Width - 250
lv.Height = Me.Height - smnu.Top - smnu.Height - 150
lv.Top = Me.Height - lv.Height - 100
lv.Left = a
''''''
smnu.Left = lv.Left + lv.Width - smnu.Width
''
sbypt.Left = Me.Width / 2 - sbypt.Width / 2
sbypt.Top = Me.Height / 2 - sbypt.Height / 2
sbyid.Left = Me.Width / 2 - sbyid.Width / 2
sbyid.Top = Me.Height / 2 - sbyid.Height / 2
End Sub
Private Sub first_Click()
rs.MoveFirst
End Sub
Private Sub dlt_Click()
rs.Delete
rs.MoveNext
If rs.EOF Then
rs.MovePrevious
End If
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
Private Sub search_Click()
n = InputBox("Search With :" & vbNewLine & vbNewLine & "1. Entry Number" & vbNewLine & "2. Party Name" _
    & vbNewLine & vbNewLine & vbNewLine & "Enter Your Choice:", "Search", 0)
        If n <> 0 And n <> "" Then
          If n = 1 Then
            ''Search by entry id
            sbyid.Visible = True
          ElseIf n = 2 Then
            '' Search By party Name
            Set rsrch = Nothing
     sql1 = "select * from accm where acctype <> 2 and acctype <> 6"
     Text22.Text = ""
     rsrch.Open sql1, cn
     inames.ListItems.Clear
     While Not rsrch.EOF
     Set Item = inames.ListItems.add(, , rsrch!Name)
     rsrch.MoveNext
    Wend
    Set rsrch = Nothing
            sbypt.Visible = True
          Else
            MsgBox "Wrong Choice", vbInformation + vbOKOnly, "ERRROR"
            Call search_Click
          End If
        End If
End Sub
Private Sub Text13_Change()
If Text13.Locked = False Then
If Text15.Text = "" Then
   Text15.Text = 0
End If
If Text13.Text = "" Then
   Text16.Text = 0
   Exit Sub
End If
Text16.Text = CDbl(Text13.Text) * CDbl(Text15.Text)
End If
End Sub
Private Sub Text15_Change()
If Text15.Locked = False Then
If Text13.Text = "" Then
   Text13.Text = 0
End If
If Text15.Text = "" Then
   Text16.Text = 0
   Exit Sub
End If
Text16.Text = CDbl(Text13.Text) * CDbl(Text15.Text)
End If
End Sub
Private Sub Text17_Change()
If Text17.Locked = False Then
If Text9.Text = "" Then
   Text9.Text = 0
End If
If Text17.Text = "" Then
   Text18.Text = 0
   Exit Sub
End If
Text18.Text = CDbl(Text9.Text) * CDbl(Text17.Text)
End If
End Sub
Private Sub Text9_Change()
If Text9.Locked = False Then
If Text17.Text = "" Then
   Text17.Text = 0
End If
If Text9.Text = "" Then
   Text18.Text = 0
   Exit Sub
End If
Text18.Text = CDbl(Text9.Text) * CDbl(Text17.Text)
End If
End Sub
Private Sub text2_Change()
Set rsrch = Nothing
If Text2.Locked = False Then
 sql1 = "select * from accm where acctype = 6 and name like '" & Text2.Text & "%'"
     rsrch.Open sql1, cn, adOpenDynamic, adLockOptimistic
     vehno.ListItems.Clear
     While Not rsrch.EOF
     Set Item = vehno.ListItems.add(, , rsrch(2))
     rsrch.MoveNext
    Wend
    Set rsrch = Nothing
End If
End Sub
Private Sub text2_GotFocus()
If Text5.Locked = False Then
     vehno.Visible = True
     vehno.Left = Text2.Left
End If
End Sub
Private Sub text2_LostFocus()
vehno.Visible = False
End Sub
Private Sub vehno_Click()
Text2.Text = vehno.SelectedItem
vehno.Left = Me.Width
End Sub
Private Sub text3_Change()
Set rsrch = Nothing
If Text3.Locked = False Then
 sql = "select * from accm where acctype <> 2 and acctype <> 6 and name like '" & Text3.Text & "%'"
     rsrch.Open sql, cn, adOpenDynamic, adLockOptimistic
     pname.ListItems.Clear
     While Not rsrch.EOF
     Set Item = pname.ListItems.add(, , rsrch!Name)
     rsrch.MoveNext
    Wend
    Set rsrch = Nothing
End If
End Sub
Private Sub text3_GotFocus()
If Text3.Locked = False Then
     pname.Visible = True
     pname.Left = Text3.Left
End If
End Sub
Private Sub text3_LostFocus()
pname.Visible = False
End Sub
Private Sub pname_Click()
Dim rs2 As New ADODB.Recordset
Text3.Text = pname.SelectedItem
pname.Left = Me.Width
If Text5.Locked = False Then
 sql = "select * from dtentry where partyname = '" & Text3.Text & "'"
     rs2.Open sql, cn
     mat.ListItems.Clear
     While Not rs2.EOF
     Set Item = site.ListItems.add(, , rs2!site)
     rsrch.MoveNext
    Wend
    Set rs2 = Nothing
End If
End Sub
Private Sub text5_Change()
Set rsrch = Nothing
If Text5.Locked = False Then
 sql = "select * from item where iname like '" & Text5.Text & "%'"
     rsrch.Open sql, cn, adOpenDynamic, adLockOptimistic
     mat.ListItems.Clear
     While Not rsrch.EOF
     Set Item = mat.ListItems.add(, , rsrch!iname)
     rsrch.MoveNext
    Wend
    Set rsrch = Nothing
End If
End Sub
Private Sub text5_GotFocus()
If Text5.Locked = False Then
     mat.Visible = True
     mat.Left = Text5.Left
End If
End Sub
Private Sub text5_LostFocus()
mat.Visible = False
End Sub
Private Sub mat_Click()
Text5.Text = mat.SelectedItem
mat.Left = Me.Width
End Sub
Private Sub text6_Change()
Set rsrch = Nothing
If Text3.Text <> "" Then
If Text6.Locked = False Then
 sql = "select * from dtentry where partyname = '" & Text3.Text & "' and site like '" & Text6.Text & "%'"
     rsrch.Open sql, cn, adOpenDynamic, adLockOptimistic
     site.ListItems.Clear
     While Not rsrch.EOF
     Set it = site.ListItems.add(, , rsrch!Name)
     rsrch.MoveNext
    Wend
    Set rsrch = Nothing
End If
Else
End If
End Sub
Private Sub text6_GotFocus()
If Text6.Locked = False Then
     site.Visible = True
     site.Left = Text6.Left
End If
End Sub
Private Sub text6_LostFocus()
site.Visible = False
End Sub
Private Sub site_Click()
Text6.Text = site.SelectedItem
site.Left = Me.Width
End Sub
Private Sub text7_GotFocus()
If Text7.Locked = False Then
     khadan.Visible = True
     khadan.Left = Text7.Left
End If
End Sub
Private Sub text7_LostFocus()
khadan.Visible = False
End Sub
Private Sub khadan_Click()
Text7.Text = khadan.SelectedItem
khadan.Left = Me.Width
End Sub
Private Sub text12_GotFocus()
If Text12.Locked = False Then
     driver.Visible = True
     driver.Left = Text12.Left
End If
End Sub
Private Sub text12_LostFocus()
driver.Visible = False
End Sub
Private Sub driver_Click()
Text12.Text = driver.SelectedItem
driver.Left = Me.Width
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
