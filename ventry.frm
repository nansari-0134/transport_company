VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form ventry 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   10215
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   19050
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   10215
   ScaleWidth      =   19050
   ShowInTaskbar   =   0   'False
   Begin VB.Frame sbyid 
      BackColor       =   &H8000000B&
      Height          =   2055
      Left            =   0
      TabIndex        =   44
      Top             =   7440
      Visible         =   0   'False
      Width           =   6495
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
         TabIndex        =   46
         Top             =   1320
         Width           =   1215
      End
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
         TabIndex        =   45
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
         TabIndex        =   53
         Top             =   0
         Width           =   375
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
         TabIndex        =   48
         Top             =   0
         Width           =   6495
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Voucher No : "
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
         TabIndex        =   47
         Top             =   720
         Width           =   2175
      End
   End
   Begin VB.Frame sbydt 
      BackColor       =   &H8000000B&
      Height          =   6735
      Left            =   120
      TabIndex        =   40
      Top             =   840
      Visible         =   0   'False
      Width           =   6495
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
         Left            =   5040
         TabIndex        =   42
         Top             =   6120
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
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
         Height          =   445
         Left            =   5160
         TabIndex        =   41
         Top             =   600
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker fdte 
         Height          =   445
         Left            =   1440
         TabIndex        =   50
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   794
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
         CalendarForeColor=   -2147483647
         CalendarTitleBackColor=   -2147483645
         CalendarTrailingForeColor=   -2147483639
         Format          =   113180673
         CurrentDate     =   43580
      End
      Begin MSComCtl2.DTPicker tdte 
         Height          =   450
         Left            =   3600
         TabIndex        =   52
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   794
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
         CalendarForeColor=   -2147483647
         CalendarTitleBackColor=   -2147483645
         CalendarTrailingForeColor=   -2147483639
         Format          =   113180673
         CurrentDate     =   43580
      End
      Begin MSComctlLib.ListView bydtlv 
         Height          =   4215
         Left            =   240
         TabIndex        =   56
         Top             =   1560
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
            Text            =   "Vou. No"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Vou. Type"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "To  Ac."
            Object.Width           =   2651
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "By Ac."
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Amount"
            Object.Width           =   1764
         EndProperty
      End
      Begin VB.Label Label10 
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
         TabIndex        =   55
         Top             =   0
         Width           =   375
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "To :"
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
         Left            =   3120
         TabIndex        =   51
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "From Date :"
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
         Left            =   120
         TabIndex        =   49
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label12 
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
         TabIndex        =   43
         Top             =   0
         Width           =   6495
      End
   End
   Begin VB.Frame sbytp 
      BackColor       =   &H8000000B&
      Height          =   6735
      Left            =   7680
      TabIndex        =   33
      Top             =   1080
      Visible         =   0   'False
      Width           =   6495
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
         TabIndex        =   36
         Top             =   6120
         Width           =   1215
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
         TabIndex        =   35
         Top             =   1080
         Width           =   1215
      End
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
         ItemData        =   "ventry.frx":0000
         Left            =   3000
         List            =   "ventry.frx":0016
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   480
         Width           =   3375
      End
      Begin MSComctlLib.ListView bytplv 
         Height          =   4215
         Left            =   240
         TabIndex        =   37
         Top             =   1800
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
            Text            =   "Vou. No"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Vou. Type"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "To  Ac."
            Object.Width           =   2651
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "By Ac."
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Amount"
            Object.Width           =   1764
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
         TabIndex        =   54
         Top             =   0
         Width           =   375
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
         TabIndex        =   39
         Top             =   0
         Width           =   6495
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Choose Voucher Type :"
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
         TabIndex        =   38
         Top             =   480
         Width           =   2535
      End
   End
   Begin MSComctlLib.ListView blv 
      Height          =   2535
      Left            =   8640
      TabIndex        =   32
      Top             =   5040
      Visible         =   0   'False
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   4471
      View            =   3
      LabelEdit       =   1
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
         Object.Width           =   9948
      EndProperty
   End
   Begin MSComctlLib.ListView tlv 
      Height          =   2415
      Left            =   8640
      TabIndex        =   31
      Top             =   4320
      Visible         =   0   'False
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   4260
      View            =   3
      LabelEdit       =   1
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
         Object.Width           =   9948
      EndProperty
   End
   Begin VB.ComboBox text1 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      ItemData        =   "ventry.frx":0057
      Left            =   2160
      List            =   "ventry.frx":006D
      TabIndex        =   30
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Frame smnu 
      BorderStyle     =   0  'None
      Height          =   550
      Left            =   2400
      TabIndex        =   27
      Top             =   7320
      Visible         =   0   'False
      Width           =   2895
      Begin VB.CommandButton cancel 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   0
         Width           =   1450
      End
      Begin VB.CommandButton save 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   0
         Width           =   1450
      End
   End
   Begin VB.TextBox Text8 
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
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   2160
      TabIndex        =   25
      Top             =   6600
      Width           =   5775
   End
   Begin VB.TextBox Text7 
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
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   2160
      TabIndex        =   24
      Top             =   5880
      Width           =   1815
   End
   Begin VB.TextBox Text6 
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
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   2160
      TabIndex        =   22
      Top             =   5160
      Width           =   2175
   End
   Begin VB.TextBox Text5 
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
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   2160
      TabIndex        =   20
      Top             =   4440
      Width           =   5655
   End
   Begin VB.TextBox Text4 
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
      ForeColor       =   &H00404040&
      Height          =   615
      Left            =   2160
      TabIndex        =   18
      Top             =   3720
      Width           =   5655
   End
   Begin VB.TextBox Text3 
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
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   2160
      TabIndex        =   16
      Top             =   2880
      Width           =   1695
   End
   Begin VB.CommandButton last 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   550
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   7920
      Width           =   1450
   End
   Begin VB.CommandButton nxt 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   550
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7920
      Width           =   1450
   End
   Begin VB.CommandButton prev 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   550
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7920
      Width           =   1450
   End
   Begin VB.CommandButton first 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   550
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7920
      Width           =   1450
   End
   Begin VB.CommandButton search 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   550
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   720
      Width           =   1450
   End
   Begin VB.CommandButton dlt 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   550
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   720
      Width           =   1450
   End
   Begin VB.CommandButton edt 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   550
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   720
      Width           =   1450
   End
   Begin VB.CommandButton add 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   550
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   720
      Width           =   1450
   End
   Begin MSComCtl2.DTPicker dte 
      Height          =   495
      Left            =   2160
      TabIndex        =   6
      Top             =   2160
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
      CalendarForeColor=   -2147483647
      CalendarTitleBackColor=   -2147483645
      CalendarTrailingForeColor=   -2147483639
      Format          =   113180673
      CurrentDate     =   43580
   End
   Begin VB.TextBox Text2 
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
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   6000
      TabIndex        =   4
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Narration"
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
      Left            =   120
      TabIndex        =   26
      Top             =   6600
      Width           =   1935
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Transaction Ref."
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
      Left            =   120
      TabIndex        =   23
      Top             =   5880
      Width           =   1935
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Amount"
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
      Left            =   120
      TabIndex        =   21
      Top             =   5160
      Width           =   1935
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "By Account"
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
      Left            =   120
      TabIndex        =   19
      Top             =   4440
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "To Account"
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
      Left            =   120
      TabIndex        =   17
      Top             =   3720
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ref. No."
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
      Left            =   120
      TabIndex        =   15
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Date "
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
      Left            =   120
      TabIndex        =   5
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Voucher No."
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
      Left            =   4440
      TabIndex        =   3
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Voucher Type "
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
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   1935
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
      Left            =   10560
      TabIndex        =   1
      Top             =   0
      Width           =   495
   End
   Begin VB.Label title 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Voucher Entry"
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
Attribute VB_Name = "ventry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim rsrch As New ADODB.Recordset
Dim sql, a, gen As String
Dim s As Integer
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
s = 1
a = rs!vno
rs.MoveLast
gen = Val(rs!vno) + 1
Set rs = Nothing
Call entext
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
dte.Value = Date
Text2.Text = gen
End Sub
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
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyH Then
    MsgBox "search button clicked"
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

Private Sub byidcan_Click()
sbyid.Visible = False
sbyidtxt.Text = ""
End Sub
Private Sub byidsrh_Click()
Dim res As String
Dim cnd As Boolean
cnd = True
Set rsrch = Nothing
sql = "select * from ventry where vno = " & sbyidtxt.Text
rsrch.Open sql, cn
If rsrch.EOF = False And rsrch.BOF = False Then
 res = MsgBox("Voucher No : " & rsrch!vno & vbNewLine _
   & "Voucher Type : " & rsrch!vtype & vbNewLine _
   & "Date : " & rsrch!adate & vbNewLine _
   & "To Account : " & rsrch!to_accnt & vbNewLine _
   & "By Account : " & rsrch!by_accnt & vbNewLine & vbNewLine _
   & "Do You Want to See Full Record ?", vbYesNo, "Search Result")
 If res = vbYes Then
    rs.MoveFirst
    While Not rs!vno = rsrch!vno And cnd = True
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
Private Sub bytpcan_Click()
sbytp.Visible = False
sbytpcm.Text = ""
bytplv.ListItems.Clear
End Sub
Private Sub bytpsrh_Click()
bytplv.ListItems.Clear
Set rsrch = Nothing
sql = "select * from ventry where vtype = '" & sbytpcm.Text & "'"
rsrch.Open sql, cn
If rsrch.EOF = False And rsrch.BOF = False Then
  If rsrch.EOF = False And rsrch.BOF = False Then
    While Not rsrch.EOF
       Set Item = bytplv.ListItems.add(, , rsrch!vno)
       Item.SubItems(1) = CStr(rsrch!vtype)
       Item.SubItems(2) = IIf(IsNull(rsrch!to_accnt), "", rsrch!to_accnt)
       Item.SubItems(3) = IIf(IsNull(rsrch!by_accnt), "", rsrch!by_accnt)
       Item.SubItems(4) = IIf(IsNull(rsrch!by_accnt), "", rsrch!amt)
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
If s = 1 Then
    Set rs = Nothing
    Call rec
    While Not rs!vno = a
     rs.MoveNext
    Wend
    Call dstext
End If
If s = 2 Then
    Set rs = Nothing
    Call rec
    While Not rs!vno = gen
     rs.MoveNext
    Wend
    Call dstext
End If
End Sub
Private Sub cancel_KeyDown(KeyCode As Integer, Shift As Integer)
If Text1.Locked = False Then
  If Shift = vbCtrlMask And KeyCode = vbKeyS Then
    Call save_Click
  End If
  If Shift = vbCtrlMask And KeyCode = vbKeyC Then
    Call cancel_Click
  End If
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
Private Sub Command1_Click()
For i = 1 To bytplv.ListItems.Count
  If bytplv.ListItems.Item(i).Checked = True Then
  rs.MoveFirst
     While Not rs.EOF
       If rs!vno = bytplv.ListItems.Item(i) Then
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
bydtlv.ListItems.Clear
Set rsrch = Nothing
sql = "select * from ventry where adate between #" & fdte.Value & "# and #" & tdte.Value & "#"
rsrch.Open sql, cn
If rsrch.EOF = False And rsrch.BOF = False Then
  If rsrch.EOF = False And rsrch.BOF = False Then
    While Not rsrch.EOF
       Set Item = bydtlv.ListItems.add(, , rsrch!vno)
       Item.SubItems(1) = CStr(rsrch!vtype)
       Item.SubItems(2) = IIf(IsNull(rsrch!to_accnt), "", rsrch!to_accnt)
       Item.SubItems(3) = IIf(IsNull(rsrch!by_accnt), "", rsrch!by_accnt)
       Item.SubItems(4) = IIf(IsNull(rsrch!amt), "", rsrch!amt)
       rsrch.MoveNext
    Wend
  End If
Else
 MsgBox "No Records found ", vbOKOnly + vbInformation, "Search Result"
End If
Set rsrch = Nothing
End Sub
Private Sub Command3_Click()
For i = 1 To bydtlv.ListItems.Count
  If bydtlv.ListItems.Item(i).Checked = True Then
  rs.MoveFirst
     While Not rs.EOF
       If rs!vno = bydtlv.ListItems.Item(i) Then
          sbydt.Visible = False
          bydtlv.ListItems.Clear
          Exit Sub
       End If
       rs.MoveNext
     Wend
  End If
Next
sbydt.Visible = False
bydtlv.ListItems.Clear
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
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyH Then
    MsgBox "search button clicked"
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
Private Sub dte_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
If KeyCode = vbKeyUp Then
   Text2.SetFocus
ElseIf KeyCode = vbKeyDown Then
   Text3.SetFocus
Else
   'nothing
End If
If Text1.Locked = True Then
  If Shift = vbCtrlMask And KeyCode = vbKeyN Then
    Call nxt_Click
  End If
  If Shift = vbCtrlMask And KeyCode = vbKeyL Then
    Call last_Click
  End If
  If Shift = vbCtrlMask And KeyCode = vbKeyF Then
    Call first_Click
  End If
  If Shift = vbCtrlMask And KeyCode = vbKeyP Then
    Call prev_Click
  End If
  If Shift = vbCtrlMask And KeyCode = vbKeyA Then
    Call add_Click
  End If
  If Shift = vbCtrlMask And KeyCode = vbKeyE Then
    Call edt_Click
  End If
  If KeyCode = vbKeyDelete Then
    Call dlt_Click
  End If
  If Shift = vbCtrlMask And KeyCode = vbKeyH Then
    MsgBox "search button clicked"
  End If
End If
If Text1.Locked = False Then
  If Shift = vbCtrlMask And KeyCode = vbKeyS Then
    Call save_Click
  End If
  If Shift = vbCtrlMask And KeyCode = vbKeyC Then
    Call cancel_Click
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
s = 2
gen = rs!vno
Set rs = Nothing
Call entext
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
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyH Then
    MsgBox "search button clicked"
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
Private Sub first_Click()
rs.MoveFirst
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
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyH Then
    MsgBox "search button clicked"
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
Private Sub Form_Unload(cancel As Integer)
Set rs = Nothing
cn.Close
End Sub
Private Sub Label10_Click()
bydtlv.ListItems.Clear
sbydt.Visible = False
End Sub
Private Sub last_Click()
rs.MoveLast
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
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyH Then
    MsgBox "search button clicked"
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
Private Sub nxt_Click()
rs.MoveNext
If Text2.Text = "" Then
   rs.MovePrevious
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
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyH Then
    MsgBox "search button clicked"
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
Private Sub prev_Click()
rs.MovePrevious
If Text2.Text = "" Then
   rs.MoveNext
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
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyH Then
    MsgBox "search button clicked"
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
If Text3.Text = "" Then
   Text3.Text = 0
End If
If Text7.Text = "" Then
    Text7.Text = 0
End If
If s = 1 Then
sql = "insert into ventry(vno,vtype,adate,refno,to_accnt,by_accnt,amt,transc_ref,narratn) values(" & Text2.Text & ",'" _
       & Text1.Text & "','" _
       & dte.Value & "'," _
       & Text3.Text & ",'" _
       & Text4.Text & "','" _
       & Text5.Text & "'," _
       & Text6.Text & ",'" _
       & Text7.Text & "','" _
       & Text8.Text & "')"
       rs.Open sql, cn
       Set rs = Nothing
       Call rec
       rs.MoveLast
End If
If s = 2 Then
If Text3.Text = "" Then
  Text3.Text = 0
End If
 sql = "update ventry set vtype ='" & Text1.Text & "',adate ='" _
       & dte.Value & "',refno =" _
       & Text3.Text & ",to_accnt ='" _
       & Text4.Text & "',by_accnt ='" _
       & Text5.Text & "',amt =" _
       & Text6.Text & ",transc_ref ='" _
       & Text7.Text & "',narratn ='" _
       & Text8.Text & "' where vno =" & gen
       rs.Open sql, cn
       Set rs = Nothing
       Call rec
           While Not rs!vno = gen
       rs.MoveNext
    Wend
    Call dstext
End If
End Sub
Private Sub Form_Load()
cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data source=" & App.Path & "\appdata\NA2KB_FFC\image.accdb;Jet OLEDB:Database Password=19012019"
'form apearance
Me.Height = Screen.Height - Screen.Height * 5 / 100 - 450
Me.Width = Screen.Width * 40 / 100
Me.Top = 450
Me.Left = Screen.Width * 20 / 100
Me.BackColor = vbWhite
Me.Picture = LoadPicture(App.Path & "\appdata\images\back.jpg")
Call setcmd
'title bar setting
Call settl
'form setting
Call setfrm
'disable all text boxes
Call dstext
'connection of text boxes to data
Call rec
'adding names to listview
Call adac
End Sub
Private Sub rec()
sql = "select * from ventry"
rs.Open sql, cn, adOpenDynamic, adLockOptimistic
Set Text1.DataSource = rs
Set Text2.DataSource = rs
Set dte.DataSource = rs
Set Text3.DataSource = rs
Set Text4.DataSource = rs
Set Text5.DataSource = rs
Set Text6.DataSource = rs
Set Text7.DataSource = rs
Set Text8.DataSource = rs
Text1.DataField = "vtype"
Text2.DataField = "vno"
dte.DataField = "adate"
Text3.DataField = "refno"
Text4.DataField = "to_accnt"
Text5.DataField = "by_accnt"
Text6.DataField = "amt"
Text7.DataField = "transc_ref"
Text8.DataField = "narratn"
End Sub
Private Sub dstext()
Text1.Locked = True
Text2.Locked = True
dte.Enabled = False
Text3.Locked = True
Text4.Locked = True
Text5.Locked = True
Text6.Locked = True
Text7.Locked = True
Text8.Locked = True
End Sub
Private Sub entext()
Text1.Locked = False
dte.Enabled = True
Text3.Locked = False
Text4.Locked = False
Text5.Locked = False
Text6.Locked = False
Text7.Locked = False
Text8.Locked = False
End Sub
Private Sub settl()
title.Width = Me.Width
title.Height = 450
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
 first.Top = Label9.Top + Label9.Height + 600
 first.Left = 300
 prev.Top = first.Top
 prev.Left = first.Left + first.Width
 nxt.Top = first.Top
 nxt.Left = prev.Left + prev.Width
 last.Top = first.Top
 last.Left = nxt.Left + nxt.Width
 smnu.Top = first.Top
 smnu.Left = Text1.Left
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
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = vbCtrlMask And KeyCode = vbKeyX Then
     Call cl_Click
   End If
End Sub
Private Sub setfrm()
Label1.Top = 1600
Label1.Left = 50
Text1.Top = Label1.Top
Text1.Left = Label1.Left + Label1.Width + 100
Label2.Top = Label1.Top
Label2.Left = Me.Width - 250 - Label2.Width - Text2.Width
Text2.Top = Label1.Top
Text2.Left = Label2.Left + Label2.Width + 100
Label3.Top = Label1.Top + Label1.Height + 100
Label3.Left = 50
dte.Top = Label3.Top
dte.Left = Label3.Left + Label3.Width + 100
Label4.Top = Label3.Top + Label3.Height + 100
Label4.Left = 50
Text3.Top = Label4.Top
Text3.Left = Label4.Left + Label4.Width + 100
Label5.Top = Label4.Top + Label4.Height + 100
Label5.Left = 50
Text4.Top = Label5.Top
Text4.Left = Label5.Left + Label5.Width + 100
Text4.Width = Me.Width - Label5.Width - 350
Label6.Top = Label5.Top + Label5.Height + 100
Label6.Left = 50
Text5.Top = Label6.Top
Text5.Left = Label6.Left + Label6.Width + 100
Text5.Width = Me.Width - Label6.Width - 350
Label7.Top = Label6.Top + Label6.Height + 100
Label7.Left = 50
Text6.Top = Label7.Top
Text6.Left = Label7.Left + Label7.Width + 100
Label8.Top = Label7.Top + Label7.Height + 100
Label8.Left = 50
Text7.Top = Label8.Top
Text7.Left = Label8.Left + Label8.Width + 100
Label9.Top = Label8.Top + Label8.Height + 100
Label9.Left = 50
Text8.Top = Label9.Top
Text8.Left = Label9.Left + Label9.Width + 100
Text8.Width = Me.Width - Label9.Width - 350
sbyid.Left = Me.Width / 2 - sbyid.Width / 2
sbyid.Top = Me.Height / 2 - sbyid.Height / 2
sbydt.Left = Me.Width / 2 - sbydt.Width / 2
sbydt.Top = Me.Height / 2 - sbydt.Height / 2
sbytp.Left = Me.Width / 2 - sbytp.Width / 2
sbytp.Top = Me.Height / 2 - sbytp.Height / 2
End Sub
Private Sub save_KeyDown(KeyCode As Integer, Shift As Integer)
If Text1.Locked = False Then
  If Shift = vbCtrlMask And KeyCode = vbKeyS Then
    Call save_Click
  End If
  If Shift = vbCtrlMask And KeyCode = vbKeyC Then
    Call cancel_Click
  End If
End If
End Sub
Private Sub search_Click()
Dim rs01 As New ADODB.Recordset
n = InputBox("Search With :" & vbNewLine & vbNewLine & "1. Voucher No " & vbNewLine & "2. Voucher Type" & vbNewLine _
    & "3. Date" _
    & vbNewLine & vbNewLine & vbNewLine & "Enter Your Choice:", "Search", 0)
        If n <> 0 And n <> "" Then
          If n = 1 Then
            ''By Voucher Number
            sbyid.Visible = True
            sbyidtxt.SetFocus
          ElseIf n = 2 Then
            ''by voucher type
             sbytp.Visible = True
          ElseIf n = 3 Then
            ''by date
             sbydt.Visible = True
          Else
            MsgBox "Wrong Choice", vbInformation + vbOKOnly, "ERRROR"
            Call search_Click
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
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyH Then
    MsgBox "search button clicked"
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
Private Sub text1_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then
      KeyCode = vbKeyF4
   ElseIf KeyCode = vbKeyRight Then
       Text2.SetFocus
   ElseIf KeyCode = vbKeyDown Then
        dte.SetFocus
   Else
        'nothing
   End If
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
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyH Then
    MsgBox "search button clicked"
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
Private Sub text2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyLeft Then
   Text1.SetFocus
Else
   'nothing
End If
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
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyH Then
    MsgBox "search button clicked"
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
Private Sub text3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
   'dte.SetFocus
ElseIf KeyCode = vbKeyDown Then
   Text4.SetFocus
Else
   'nothing
End If
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
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyH Then
    MsgBox "search button clicked"
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
Private Sub text4_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
   Text3.SetFocus
ElseIf KeyCode = vbKeyDown Then
   Text5.SetFocus
Else
   'nothing
End If
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
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyH Then
    MsgBox "search button clicked"
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
Private Sub text5_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
   Text4.SetFocus
ElseIf KeyCode = vbKeyDown Then
   Text6.SetFocus
Else
   'nothing
End If
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
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyH Then
    MsgBox "search button clicked"
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
Private Sub text6_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
   Text5.SetFocus
ElseIf KeyCode = vbKeyDown Then
   Text7.SetFocus
Else
   'nothing
End If
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
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyH Then
    MsgBox "search button clicked"
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
Private Sub text7_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
   Text6.SetFocus
ElseIf KeyCode = vbKeyDown Then
   Text8.SetFocus
Else
   'nothing
End If
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
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyH Then
    MsgBox "search button clicked"
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
Private Sub text8_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
   Text7.SetFocus
End If
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
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyH Then
    MsgBox "search button clicked"
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
Private Sub tlv_Click()
Text4.Text = tlv.SelectedItem
tlv.Left = Me.Width
End Sub
Private Sub Text4_Change()
If Text4.Locked = False Then
sql = "select * from accm where name like '" & Text4.Text & "%' and acctype <> 3 and acctype <> 6"
rs1.Open sql, cn, adOpenDynamic, adLockOptimistic
tlv.ListItems.Clear
While Not rs1.EOF
Set Item = tlv.ListItems.add(, , rs1!Name)
rs1.MoveNext
Wend
Set rs1 = Nothing
End If
End Sub
Private Sub Text4_GotFocus()
If Text4.Locked = False Then
     tlv.Visible = True
     tlv.Left = Text4.Left
     tlv.Top = Text4.Top + Text4.Height + 20
End If
End Sub
Private Sub Text4_LostFocus()
tlv.Visible = False
End Sub
Private Sub blv_Click()
Text5.Text = blv.SelectedItem
blv.Left = Me.Width
End Sub
Private Sub text5_Change()
If Text5.Locked = False Then
sql = "select * from accm where name like '" & Text5.Text & "%' and acctype <> 3 and acctype <> 6"
rs1.Open sql, cn, adOpenDynamic, adLockOptimistic
blv.ListItems.Clear
While Not rs1.EOF
Set Item = blv.ListItems.add(, , rs1!Name)
rs1.MoveNext
Wend
Set rs1 = Nothing
End If
End Sub
Private Sub text5_GotFocus()
If Text5.Locked = False Then
     blv.Visible = True
     blv.Left = Text5.Left
     blv.Top = Text5.Top + Text5.Height + 20
End If
End Sub
Private Sub text5_LostFocus()
blv.Visible = False
End Sub
 Private Sub adac()
 blv.Left = Me.Width
 tlv.Left = Me.Width
 blv.ColumnHeaders(1).Width = blv.Width * 99 / 100
 tlv.ColumnHeaders(1).Width = blv.Width * 99 / 100
 sql = "select * from accm where acctype <> 3 and acctype <> 6"
 rs1.Open sql, cn, adOpenDynamic, adLockOptimistic
 While Not rs1.EOF
 Set Item = blv.ListItems.add(, , rs1!Name)
 rs1.MoveNext
 Wend
 Set rs1 = Nothing
 End Sub
