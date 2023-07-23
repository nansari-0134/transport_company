VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form centry 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11490
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   16290
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   11490
   ScaleWidth      =   16290
   ShowInTaskbar   =   0   'False
   Begin VB.Frame sbyid 
      BackColor       =   &H8000000B&
      Height          =   2055
      Left            =   6840
      TabIndex        =   63
      Top             =   1080
      Visible         =   0   'False
      Width           =   6495
      Begin VB.CommandButton byidsrh 
         Caption         =   "Search"
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
         Left            =   5040
         TabIndex        =   65
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox sbyidtxt 
         BackColor       =   &H80000003&
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
         Left            =   2640
         TabIndex        =   64
         Top             =   720
         Width           =   3615
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
         TabIndex        =   68
         Top             =   0
         Width           =   6495
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Entry No. : "
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
         TabIndex        =   67
         Top             =   720
         Width           =   2175
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
         TabIndex        =   66
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.Frame sbypt 
      BackColor       =   &H8000000B&
      Height          =   6735
      Left            =   1560
      TabIndex        =   54
      Top             =   3480
      Visible         =   0   'False
      Width           =   6495
      Begin VB.CommandButton Command3 
         Caption         =   "Search"
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
         Left            =   1200
         TabIndex        =   57
         Top             =   1320
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
         TabIndex        =   56
         Top             =   480
         Width           =   3735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Show"
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
         Left            =   5040
         TabIndex        =   55
         Top             =   6120
         Width           =   1215
      End
      Begin MSComctlLib.ListView inames 
         Height          =   4215
         Left            =   2520
         TabIndex        =   58
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
         Left            =   120
         TabIndex        =   59
         Top             =   1920
         Width           =   6255
         _ExtentX        =   11033
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
            Text            =   "ID"
            Object.Width           =   2122
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Date"
            Object.Width           =   2122
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Item"
            Object.Width           =   2122
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Net Amount"
            Object.Width           =   2122
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Narration"
            Object.Width           =   2122
         EndProperty
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
         TabIndex        =   62
         Top             =   0
         Width           =   6495
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Purchase Account:"
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
         TabIndex        =   61
         Top             =   480
         Width           =   2175
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
         TabIndex        =   60
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.CommandButton adcan 
      Caption         =   "Cancel"
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
      Left            =   7440
      TabIndex        =   53
      Top             =   2760
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton addmi 
      Caption         =   "Add More Items.."
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
      Left            =   9720
      TabIndex        =   52
      Top             =   2760
      Visible         =   0   'False
      Width           =   2055
   End
   Begin MSComctlLib.ListView palv 
      Height          =   1335
      Left            =   12120
      TabIndex        =   51
      Top             =   2520
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   2355
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
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
   Begin MSComctlLib.ListView plv 
      Height          =   1455
      Left            =   12000
      TabIndex        =   50
      Top             =   720
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   2566
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
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
   Begin MSComCtl2.DTPicker dte 
      Height          =   435
      Left            =   8400
      TabIndex        =   48
      Top             =   1320
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   767
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
      Format          =   112656385
      CurrentDate     =   43582
   End
   Begin VB.TextBox Text16 
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
      ForeColor       =   &H00404040&
      Height          =   435
      Left            =   9480
      TabIndex        =   44
      Top             =   9120
      Width           =   2055
   End
   Begin VB.TextBox Text15 
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
      ForeColor       =   &H00404040&
      Height          =   435
      Left            =   9480
      TabIndex        =   42
      Top             =   8520
      Width           =   2055
   End
   Begin VB.TextBox Text14 
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
      ForeColor       =   &H00404040&
      Height          =   435
      Left            =   7320
      TabIndex        =   41
      Top             =   8520
      Width           =   2055
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
      ForeColor       =   &H00404040&
      Height          =   435
      Left            =   9480
      TabIndex        =   39
      Top             =   7920
      Width           =   2055
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
      ForeColor       =   &H00404040&
      Height          =   435
      Left            =   7320
      TabIndex        =   38
      Top             =   7920
      Width           =   2055
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
      ForeColor       =   &H00404040&
      Height          =   435
      Left            =   9480
      TabIndex        =   36
      Top             =   7320
      Width           =   2055
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
      ForeColor       =   &H00404040&
      Height          =   435
      Left            =   1800
      TabIndex        =   34
      Top             =   7320
      Width           =   4935
   End
   Begin VB.Frame padd 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   120
      TabIndex        =   22
      Top             =   3240
      Visible         =   0   'False
      Width           =   11655
      Begin MSComctlLib.ListView ilv 
         Height          =   975
         Left            =   2280
         TabIndex        =   49
         Top             =   600
         Visible         =   0   'False
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   1720
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
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
      Begin VB.CommandButton pok 
         Caption         =   "Ok"
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
         Left            =   9840
         TabIndex        =   46
         Top             =   960
         Width           =   1695
      End
      Begin VB.CommandButton pmore 
         Caption         =   "More.."
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
         Left            =   9840
         TabIndex        =   45
         Top             =   120
         Width           =   1695
      End
      Begin VB.TextBox Text9 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   435
         Left            =   1920
         TabIndex        =   32
         Top             =   1080
         Width           =   7455
      End
      Begin VB.TextBox Text8 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   435
         Left            =   6960
         TabIndex        =   30
         Top             =   600
         Width           =   2415
      End
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   435
         Left            =   1920
         TabIndex        =   28
         Top             =   600
         Width           =   2415
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   435
         Left            =   6960
         TabIndex        =   26
         Top             =   120
         Width           =   2415
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   435
         Left            =   1920
         TabIndex        =   24
         Top             =   120
         Width           =   2415
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Description"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   120
         TabIndex        =   31
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Discount  %"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   5160
         TabIndex        =   29
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Rate"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   120
         TabIndex        =   27
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Qty"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   5160
         TabIndex        =   25
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Item name"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   120
         TabIndex        =   23
         Top             =   120
         Width           =   1575
      End
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
      ForeColor       =   &H00404040&
      Height          =   435
      Left            =   1800
      TabIndex        =   21
      Top             =   2640
      Width           =   2895
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
      ForeColor       =   &H00404040&
      Height          =   435
      Left            =   8160
      TabIndex        =   19
      Top             =   1920
      Width           =   2895
   End
   Begin VB.TextBox Text2 
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
      ForeColor       =   &H00404040&
      Height          =   435
      Left            =   1800
      TabIndex        =   17
      Top             =   1920
      Width           =   2895
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
      ForeColor       =   &H00404040&
      Height          =   435
      Left            =   1800
      TabIndex        =   14
      Top             =   1320
      Width           =   2415
   End
   Begin VB.Frame smnu 
      BorderStyle     =   0  'None
      Height          =   550
      Left            =   6120
      TabIndex        =   10
      Top             =   10800
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
         TabIndex        =   12
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
         TabIndex        =   11
         Top             =   0
         Width           =   1450
      End
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
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   10800
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
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   10800
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
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   10800
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
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   10800
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
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   600
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
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   600
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
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   600
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
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   600
      Width           =   1450
   End
   Begin MSComctlLib.ListView sdata 
      Height          =   2295
      Left            =   120
      TabIndex        =   47
      Top             =   4920
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   4048
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
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
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Sr. No."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Item Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Description"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Qty"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "UOM"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Rate"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Disc % "
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Disc Amt."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Amount"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Net Amount"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   7800
      TabIndex        =   43
      Top             =   9120
      Width           =   1575
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Less"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6240
      TabIndex        =   40
      Top             =   8520
      Width           =   975
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6240
      TabIndex        =   37
      Top             =   7920
      Width           =   975
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total Amount "
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   7800
      TabIndex        =   35
      Top             =   7320
      Width           =   1575
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Narration"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   120
      TabIndex        =   33
      Top             =   7320
      Width           =   1575
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ref. Number"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   120
      TabIndex        =   20
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Consumption Account"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5520
      TabIndex        =   18
      Top             =   1920
      Width           =   2415
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Item name "
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   120
      TabIndex        =   16
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6720
      TabIndex        =   15
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Entry Id"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   120
      TabIndex        =   13
      Top             =   1320
      Width           =   1575
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
      Caption         =   "Consumption Entry"
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
Attribute VB_Name = "centry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim h, ph, tmt As Double
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim rs3 As New ADODB.Recordset
Dim rsrch As New ADODB.Recordset
Dim sql, sql1 As String
Dim n, a, gen, i, sno As Integer
Private Sub adcan_Click()
Text5.Text = Clear
Text6.Text = Clear
Text7.Text = Clear
Text8.Text = Clear
Text9.Text = Clear
Call setnotadd
adcan.Visible = False
addmi.Visible = True
End Sub
Private Sub addmi_Click()
Call setadd
adcan.Visible = True
addmi.Visible = False
End Sub
Private Sub cl_Click()
If smnu.Visible = True Then
MsgBox "Please Save or Cancel Current Record first", vbInformation + vbOKOnly, "Error"
Else
menu.Enabled = True
Unload Me
End If
End Sub
Private Sub byagcan_Click()
sbypt.Visible = False
Text22.Text = ""
End Sub

Private Sub byidcan_Click()
sbyidtxt.Text = ""
sbyid.Visible = False
End Sub
Private Sub search_Click()
n = InputBox("Search With :" & vbNewLine & vbNewLine & "1. Entry Id " & vbNewLine & "2. Consumption Account" _
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
Private Sub Command3_Click()
Set rsrch = Nothing
byptlv.ListItems.Clear
If Text22.Text <> "" Then
 sql = "select * from centry where caccnt = '" & Text22.Text & "'"
 rsrch.Open sql, cn
 While Not rsrch.EOF
    Set Item = byptlv.ListItems.add(, , rsrch!entryid)
    Item.SubItems(1) = rsrch!adate
    Item.SubItems(2) = rsrch!party_name
    Item.SubItems(3) = rsrch!netamt
    Item.SubItems(4) = IIf(IsNull(rsrch!narratn), "", rsrch!narratn)
    rsrch.MoveNext
 Wend
End If
Set rsrch = Nothing
End Sub
Private Sub Command2_Click()
For i = 1 To byptlv.ListItems.Count
  If byptlv.ListItems.Item(i).Checked = True Then
  Call first_Click
     While Not rs.EOF
       If rs!entryid = byptlv.ListItems.Item(i) Then
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
Private Sub byidsrh_Click()
Dim res As String
Dim cnd As Boolean
Dim eno As Long
cnd = True
Set rsrch = Nothing
sql = "select * from centry where entryid = " & sbyidtxt.Text
rsrch.Open sql, cn
If rsrch.EOF = False And rsrch.BOF = False Then
 res = MsgBox("Entry Id : " & rsrch!entryid & vbNewLine _
   & "Date : " & rsrch!adate & vbNewLine _
   & "Item : " & rsrch!party_name & vbNewLine _
   & "Consumption Account  : " & rsrch!caccnt & vbNewLine _
   & "Net Amount : " & rsrch!netamt & vbNewLine & vbNewLine _
   & "Do You Want to See Full Record ?", vbYesNo, "Search Result")
eno = rsrch!entryid
 If res = vbYes Then
    Call first_Click
    While Not rs!entryid = eno And cnd = True
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
Private Sub dlt_Click()
n = 0
For i = 1 To sdata.ListItems.Count
If sdata.ListItems(i).Checked = True Then
   Set rs2 = Nothing
   sql = "delete from cdet where entryid = " & Text1.Text & " and srno = " & sdata.ListItems(i)
   rs2.Open sql, cn
   Set rs2 = Nothing
   Set rs3 = Nothing
   sql1 = "update item set currt_stock = currt_stock + " & sdata.ListItems(i).SubItems(3) & " where iname = '" _
         & sdata.ListItems(i).SubItems(1) & "'"
   rs3.Open sql1, cn
   Set rs3 = Nothing
   'sdata.ListItems.Remove (i)
   tmt = 0
    For j = 1 To sdata.ListItems.Count
    tmt = tmt + CDbl(sdata.ListItems(j).ListSubItems(8).Text)
    Next
    Text11.Text = tmt
    If Text13.Text = "" Then
        Text13.Text = 0
    End If
    If Text15.Text = "" Then
        Text15.Text = 0
    End If
    tmt = tmt + CDbl(Text13.Text) - CDbl(Text15.Text)
    sql1 = "update pentry set netamt = " & tmt & " where pid = " & Text1.Text
    rs3.Open sql1, cn
    Set rs3 = Nothing
    Text16.Text = tmt
   n = 5
End If
Next
If n = 5 Then
                    For i = sdata.ListItems.Count To 1 Step -1
                            If sdata.ListItems(i).Checked = True Then
                               sdata.ListItems.remove (i)
                            End If
                    Next
End If
If n <> 5 Then
   i = MsgBox("Do You Want To Delete the current record", vbInformation + vbOKCancel, "WARNING")
                    If i = vbOK Then
                      Set rs2 = Nothing
                      sql = "delete from centry where entryid = '" & Text1.Text & "'"
                      rs2.Open sql, cn, adOpenDynamic, adLockOptimistic
                      Set rs2 = Nothing
                      sql = "delete from cdet where entryid = '" & Text1.Text & "'"
                      rs2.Open sql, cn, adOpenDynamic, adLockOptimistic
                      Set rs2 = Nothing
                    End If
                    For i = 1 To sdata.ListItems.Count
                        Set rs3 = Nothing
                        sql1 = "update item set currt_stock = currt_stock + " & sdata.ListItems(i).SubItems(3) & " where iname = '" _
                              & sdata.ListItems(i).SubItems(1) & "'"
                        rs3.Open sql1, cn
                        Set rs3 = Nothing
                    Next
                    Call nxt_Click
End If
End Sub
Private Sub first_Click()
rs.MoveFirst
Call disp
End Sub
Private Sub Form_Load()
cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data source=" & App.Path & "\appdata\NA2KB_FFC\image.accdb;Jet OLEDB:Database Password=19012019"
h = sdata.Height
ph = padd.Height
'form apearance
Me.Height = Screen.Height - Screen.Height * 5 / 100 - 450
Me.Width = Screen.Width * 80 / 100
Me.Top = 450
Me.Left = Screen.Width * 20 / 100
Me.BackColor = vbWhite
Me.Picture = LoadPicture(App.Path & "\appdata\images\back.jpg")
tmt = 0
'title bar setting
Call settl
Call setcmd
Call setfrm
Call setnotadd
addmi.Left = Me.Width - addmi.Width - 80
addmi.Top = sdata.Top - addmi.Height - 20
adcan.Left = addmi.Left
adcan.Top = addmi.Top
Call setlv
Call rec
Call dstext
End Sub
Private Sub dstext()
Text1.Locked = True
dte.Enabled = False
Text2.Locked = True
Text3.Locked = True
Text4.Locked = True
Text10.Locked = True
Text11.Locked = True
Text12.Locked = True
Text13.Locked = True
Text14.Locked = True
Text15.Locked = True
Text16.Locked = True
End Sub
Private Sub entext()
dte.Enabled = True
Text2.Locked = False
Text3.Locked = False
Text4.Locked = False
Text10.Locked = False
Text11.Locked = False
Text12.Locked = False
Text13.Locked = False
Text14.Locked = False
Text15.Locked = False
Text16.Locked = False
End Sub
Private Sub rec()
Set rs = Nothing
sql = "select * from centry"
rs.Open sql, cn, adOpenDynamic, adLockOptimistic
Set Text1.DataSource = rs
Set dte.DataSource = rs
Set Text2.DataSource = rs
Set Text3.DataSource = rs
Set Text4.DataSource = rs
Set Text10.DataSource = rs
Set Text16.DataSource = rs
Set Text12.DataSource = rs
Set Text13.DataSource = rs
Set Text14.DataSource = rs
Set Text15.DataSource = rs
Text1.DataField = "entryid"
Text2.DataField = "party_name"
Text3.DataField = "caccnt"
Text4.DataField = "refno"
Text10.DataField = "narratn"
Text16.DataField = "netamt"
Text12.DataField = "add_desc"
Text13.DataField = "add_amt"
Text14.DataField = "less_desc"
Text15.DataField = "less_amt"
dte.DataField = "adate"
Call disp
End Sub
Private Sub disp()
sdata.ListItems.Clear
sql = "select * from cdet where entryid=" & Text1.Text
rs1.Open sql, cn, adOpenDynamic, adLockOptimistic
While Not rs1.EOF
Set Item = sdata.ListItems.add(, , rs1!srno)
Item.SubItems(1) = rs1!Item
Item.SubItems(2) = rs1!Desc & ""
Item.SubItems(3) = rs1!qty
Item.SubItems(4) = rs1!unit
Item.SubItems(5) = rs1!Rate
Item.SubItems(6) = rs1!disc
Item.SubItems(7) = rs1!discamt
Item.SubItems(8) = rs1!amt
rs1.MoveNext
Wend
tmt = 0
For i = 1 To sdata.ListItems.Count
tmt = tmt + CDbl(sdata.ListItems(i).ListSubItems(8).Text)
Next
Text11.Text = tmt
Set rs1 = Nothing
End Sub
Private Sub setlv()
sdata.ColumnHeaders(1).Width = sdata.Width * 7 / 100
sdata.ColumnHeaders(2).Width = sdata.Width * 21 / 100
sdata.ColumnHeaders(3).Width = sdata.Width * 19 / 100
sdata.ColumnHeaders(4).Width = sdata.Width * 8 / 100
sdata.ColumnHeaders(5).Width = sdata.Width * 6 / 100
sdata.ColumnHeaders(6).Width = sdata.Width * 6 / 100
sdata.ColumnHeaders(7).Width = sdata.Width * 7 / 100
sdata.ColumnHeaders(8).Width = sdata.Width * 10 / 100
sdata.ColumnHeaders(9).Width = sdata.Width * 15 / 100
ilv.ColumnHeaders(1).Width = Text5.Width
plv.ColumnHeaders(1).Width = Text5.Width
End Sub
Private Sub setfrm()
Label1.Top = add.Top + add.Height + 150
Label1.Left = 120
Text1.Top = Label1.Top
Text1.Left = Label1.Left + Label1.Width + 100
Label2.Top = Label1.Top
Label2.Left = Text2.Left + Text2.Width + 500
dte.Top = Label1.Top
dte.Left = Label2.Left + Label2.Width + 100
Label3.Top = Label1.Top + Label1.Height + 100
Label3.Left = 120
Text2.Top = Label3.Top
Text2.Left = Label3.Left + Label3.Width + 100
Label4.Top = Label3.Top
Label4.Left = Label2.Left
Text3.Top = Label3.Top
Text3.Left = Label4.Left + Label4.Width + 100
Label5.Top = Label3.Top + Label3.Height + 100
Label5.Left = 120
Text4.Top = Label5.Top
Text4.Left = Label5.Left + Label5.Width + 100
plv.Top = Text2.Top + Text2.Height + 20
palv.Top = Text3.Top + Text3.Height + 20
ilv.Top = Text5.Top + Text5.Height + 20
sbypt.Left = Me.Width / 2 - sbypt.Width / 2
sbypt.Top = Me.Height / 2 - sbypt.Height / 2
sbyid.Left = Me.Width / 2 - sbyid.Width / 2
sbyid.Top = Me.Height / 2 - sbyid.Height / 2
End Sub
Private Sub setadd()
padd.Visible = True
sdata.Height = h
sdata.Width = Me.Width - 200
padd.Top = Label5.Top + Label5.Height + 100
padd.Left = Me.Width / 2 - padd.Width / 2
sdata.Top = padd.Top + padd.Height + 150
sdata.Left = 120
sdata.Width = Me.Width - 120
Label11.Top = sdata.Top + sdata.Height + 150
Label11.Left = 120
Text10.Top = Label11.Top
Text10.Left = Label11.Left + Label11.Width + 100
Label12.Top = Label11.Top
Label12.Left = Me.Width - Label12.Width - Text11.Width - 350
Text11.Top = Label11.Top
Text11.Left = Label12.Left + Label12.Width + 100
Label13.Top = Label11.Top + Label11.Height + 300
Label13.Left = Me.Width - Label13.Width - Text12.Width - Text13.Width - 450
Text12.Top = Label13.Top
Text12.Left = Label13.Left + Label13.Width + 100
Text13.Top = Label13.Top
Text13.Left = Text12.Left + Text12.Width + 100
Label14.Top = Label13.Top + Label13.Height + 100
Label14.Left = Label13.Left
Text14.Top = Label14.Top
Text14.Left = Text12.Left
Text15.Top = Label14.Top
Text15.Left = Text13.Left
Text16.Top = Label14.Top + Label14.Height + 100
Text16.Left = Text15.Left
Label15.Top = Text16.Top
Label15.Left = Text16.Left - 100 - Label15.Width
End Sub
Private Sub setnotadd()
padd.Visible = False
sdata.Height = ph + h
sdata.Width = Me.Width - 200
sdata.Top = Label5.Top + Label5.Height + 150
sdata.Left = 120
Label11.Top = sdata.Top + sdata.Height + 150
Label11.Left = 120
Text10.Top = Label11.Top
Text10.Left = Label11.Left + Label11.Width + 100
Label12.Top = Label11.Top
Label12.Left = Me.Width - Label12.Width - Text11.Width - 350
Text11.Top = Label11.Top
Text11.Left = Label12.Left + Label12.Width + 100
Label13.Top = Label11.Top + Label11.Height + 300
Label13.Left = Me.Width - Label13.Width - Text12.Width - Text13.Width - 450
Text12.Top = Label13.Top
Text12.Left = Label13.Left + Label13.Width + 100
Text13.Top = Label13.Top
Text13.Left = Text12.Left + Text12.Width + 100
Label14.Top = Label13.Top + Label13.Height + 100
Label14.Left = Label13.Left
Text14.Top = Label14.Top
Text14.Left = Text12.Left
Text15.Top = Label14.Top
Text15.Left = Text13.Left
Text16.Top = Label14.Top + Label14.Height + 100
Text16.Left = Text15.Left
Label15.Top = Text16.Top
Label15.Left = Text16.Left - 100 - Label15.Width
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
padd.Visible = True
Call setadd
Call clrtxt
n = 1
sno = 0
gen = rs!entryid
sdata.ListItems.Clear
Set rs1 = Nothing
sql = "select * from centry"
rs1.Open sql, cn, adOpenDynamic, adLockOptimistic
rs1.MoveLast
Text1.Text = rs1!entryid + 1
Set rs1 = Nothing

Set rs2 = Nothing
sql = "select * from accm"
rs2.Open sql, cn, adOpenDynamic, adLockOptimistic
While Not rs2.EOF
 Set Item = plv.ListItems.add(, , rs2!Name)
 Set it = palv.ListItems.add(, , rs2!Name)
 rs2.MoveNext
Wend
Set rs2 = Nothing
sql = "select * from item"
rs2.Open sql, cn, adOpenDynamic, adLockOptimistic
While Not rs2.EOF
 Set Item = ilv.ListItems.add(, , rs2!iname)
 rs2.MoveNext
Wend
Set rs2 = Nothing
Set rs = Nothing
Call entext
End Sub
Private Sub clrtxt()
Text2.Text = Clear
Text3.Text = Clear
Text4.Text = Clear
Text10.Text = Clear
Text11.Text = Clear
Text12.Text = Clear
Text13.Text = Clear
Text14.Text = Clear
Text15.Text = Clear
Text16.Text = Clear
End Sub
Private Sub cancel_Click()
If n = 1 Then
'add-cancel
Set rs = Nothing
Call rec
While Not rs!entryid = gen
rs.MoveNext
Wend
Call disp
Call dstext
End If
If n = 2 Then
'edit-cancel
Set rs = Nothing
Call rec
While Not rs!entryid = gen
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
padd.Visible = False
Call setnotadd
addmi.Visible = False
adcan.Visible = False
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
Call entext
n = 2
gen = rs!entryid
Set rs = Nothing
Set rs2 = Nothing
sql = "select * from item"
rs2.Open sql, cn, adOpenDynamic, adLockOptimistic
While Not rs2.EOF
 Set Item = ilv.ListItems.add(, , rs2!iname)
 rs2.MoveNext
Wend
Set rs2 = Nothing
addmi.Visible = True
a = sdata.ListItems.Count
End Sub
Private Sub Form_Unload(cancel As Integer)
Set rs = Nothing
cn.Close
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
Private Sub pmore_Click()
If Text8.Text = "" Then
   Text8.Text = 0
End If
If n = 1 Then
 'add
        sno = sno + 1
        Set rs2 = Nothing
        sql = "select * from item where iname = '" & Text5.Text & "'"
        rs2.Open sql, cn
        ut = rs2!unit
        Set rs2 = Nothing
        Set Item = sdata.ListItems.add(, , sno)
        Item.SubItems(1) = Text5.Text
        Item.SubItems(2) = Text9.Text
        Item.SubItems(3) = Text6.Text
        Item.SubItems(4) = ut
        Item.SubItems(5) = Text7.Text
        Item.SubItems(6) = Text8.Text
        Item.SubItems(7) = (CDbl(Text6.Text) * CDbl(Text7.Text)) * CDbl(Text8.Text) / 100
        Item.SubItems(8) = ((CDbl(Text6.Text) * CDbl(Text7.Text)) - ((CDbl(Text6.Text) * CDbl(Text7.Text)) * CDbl(Text8.Text) / 100))
        Text5.Text = Clear
        Text6.Text = Clear
        Text7.Text = Clear
        Text8.Text = Clear
        Text9.Text = Clear
ElseIf n = 2 Then
 'edit
        sno = Val(sdata.ListItems(sdata.ListItems.Count)) + 1
        Set rs2 = Nothing
        sql = "select * from item where iname = '" & Text5.Text & "'"
        rs2.Open sql, cn
        ut = rs2!unit
        Set rs2 = Nothing
        Set Item = sdata.ListItems.add(, , sno)
        Item.SubItems(1) = Text5.Text
        Item.SubItems(2) = Text9.Text
        Item.SubItems(3) = Text6.Text
        Item.SubItems(4) = ut
        Item.SubItems(5) = Text7.Text
        Item.SubItems(6) = Text8.Text
        Item.SubItems(7) = (CDbl(Text6.Text) * CDbl(Text7.Text)) * CDbl(Text8.Text) / 100
        Item.SubItems(8) = ((CDbl(Text6.Text) * CDbl(Text7.Text)) - ((CDbl(Text6.Text) * CDbl(Text7.Text)) * CDbl(Text8.Text) / 100))
        Text5.Text = Clear
        Text6.Text = Clear
        Text7.Text = Clear
        Text8.Text = Clear
        Text9.Text = Clear
Else
   'nothing
End If
End Sub
Private Sub pok_Click()
If n = 1 Then
  'add
    addmi.Visible = True
    Call setnotadd
    If Text8.Text = "" Then
       Text8.Text = 0
    End If
    sno = sno + 1
    Set rs2 = Nothing
    sql = "select * from item where iname = '" & Text5.Text & "'"
    rs2.Open sql, cn
    ut = rs2!unit
    Set rs2 = Nothing
    Set Item = sdata.ListItems.add(, , sno)
    Item.SubItems(1) = Text5.Text
    Item.SubItems(2) = Text9.Text
    Item.SubItems(3) = Text6.Text
    Item.SubItems(4) = ut
    Item.SubItems(5) = Text7.Text
    Item.SubItems(6) = Text8.Text
    Item.SubItems(7) = (CDbl(Text6.Text) * CDbl(Text7.Text)) * CDbl(Text8.Text) / 100
    Item.SubItems(8) = ((CDbl(Text6.Text) * CDbl(Text7.Text)) - ((CDbl(Text6.Text) * CDbl(Text7.Text)) * CDbl(Text8.Text) / 100))
    Text5.Text = Clear
    Text6.Text = Clear
    Text7.Text = Clear
    Text8.Text = Clear
    Text9.Text = Clear
ElseIf n = 2 Then
  'edit
    addmi.Visible = True
    Call setnotadd
    If Text8.Text = "" Then
       Text8.Text = 0
    End If
    If sdata.ListItems.Count <> 0 Then
    sno = Val(sdata.ListItems(sdata.ListItems.Count)) + 1
    Else
    sno = 1
    End If
    Set rs2 = Nothing
    sql = "select * from item where iname = '" & Text5.Text & "'"
    rs2.Open sql, cn
    ut = rs2!unit
    Set rs2 = Nothing
    Set Item = sdata.ListItems.add(, , sno)
    Item.SubItems(1) = Text5.Text
    Item.SubItems(2) = Text9.Text
    Item.SubItems(3) = Text6.Text
    Item.SubItems(4) = ut
    Item.SubItems(5) = Text7.Text
    Item.SubItems(6) = Text8.Text
    Item.SubItems(7) = (CDbl(Text6.Text) * CDbl(Text7.Text)) * CDbl(Text8.Text) / 100
    Item.SubItems(8) = ((CDbl(Text6.Text) * CDbl(Text7.Text)) - ((CDbl(Text6.Text) * CDbl(Text7.Text)) * CDbl(Text8.Text) / 100))
    Text5.Text = Clear
    Text6.Text = Clear
    Text7.Text = Clear
    Text8.Text = Clear
    Text9.Text = Clear
Else
    'nothing
End If
adcan.Visible = False
tmt = 0
For i = 1 To sdata.ListItems.Count
tmt = tmt + CDbl(sdata.ListItems(i).ListSubItems(8).Text)
Next
Text11.Text = tmt
Text16.Text = tmt
Set rs1 = Nothing
End Sub
Private Sub prev_Click()
rs.MovePrevious
If Text1.Text = "" Then
rs.MoveFirst
End If
Call disp
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
padd.Visible = False
addmi.Visible = False
adcan.Visible = False
Call setnotadd
Call dstext
If Text4.Text = "" Then
     Text4.Text = 0
  End If
  If Text13.Text = "" Then
     Text13.Text = 0
  End If
  If Text15.Text = "" Then
     Text15.Text = 0
  End If
If n = 1 Then
  'add-save
    'inserting records in cdet
 For i = 1 To sdata.ListItems.Count
 If sdata.ListItems(i).SubItems(2) = "" Then
      sdata.ListItems(i).SubItems(2) = " "
 End If
  sql = "insert into cdet(entryid,srno,item,[desc],qty,unit,rate,disc,discamt,amt) " _
        & "values(" & Text1.Text & "," _
        & sdata.ListItems(i) & ",'" _
        & sdata.ListItems(i).SubItems(1) & "','" _
        & sdata.ListItems(i).SubItems(2) & "'," _
        & sdata.ListItems(i).SubItems(3) & ",'" _
        & sdata.ListItems(i).SubItems(4) & "'," _
        & sdata.ListItems(i).SubItems(5) & "," _
        & sdata.ListItems(i).SubItems(6) & "," _
        & sdata.ListItems(i).SubItems(7) & "," _
        & sdata.ListItems(i).SubItems(8) & ")"
  rs.Open sql, cn
  Set rs3 = Nothing
  sql1 = "update item set currt_stock = currt_stock - " & sdata.ListItems(i).SubItems(3) & " where iname = '" _
         & sdata.ListItems(i).SubItems(1) & "'"
  rs3.Open sql1, cn
  Set rs3 = Nothing
  Set rs = Nothing
 Next
  'inserting records in centry
  sql = "insert into centry(entryid,adate,party_name,caccnt,netamt,add_desc,add_amt,less_desc,less_amt,refno,narratn) " _
        & "values(" & Text1.Text & ",'" _
        & dte.Value & "','" _
        & Text2.Text & "','" _
        & Text3.Text & "'," _
        & Text16.Text & ",'" _
        & Text12.Text & "'," _
        & Text13.Text & ",'" _
        & Text14.Text & "'," _
        & Text15.Text & "," _
        & Text4.Text & ",'" _
        & Text10.Text & "')"
  rs.Open sql, cn
  Set rs = Nothing
    Call rec
    rs.MoveLast
    Call disp
ElseIf n = 2 Then
  'edit save
  If sdata.ListItems.Count <> a Then
     For i = a + 1 To sdata.ListItems.Count
      sql = "insert into cdet(entryid,srno,item,[desc],qty,unit,rate,disc,discamt,amt) " _
        & "values(" & Text1.Text & "," _
        & sdata.ListItems(i) & ",'" _
        & sdata.ListItems(i).SubItems(1) & "','" _
        & sdata.ListItems(i).SubItems(2) & "'," _
        & sdata.ListItems(i).SubItems(3) & ",'" _
        & sdata.ListItems(i).SubItems(4) & "'," _
        & sdata.ListItems(i).SubItems(5) & "," _
        & sdata.ListItems(i).SubItems(6) & "," _
        & sdata.ListItems(i).SubItems(7) & "," _
        & sdata.ListItems(i).SubItems(8) & ")"
  rs.Open sql, cn
  Set rs3 = Nothing
  sql1 = "update item set currt_stock = currt_stock - " & sdata.ListItems(i).SubItems(3) & " where iname = '" _
         & sdata.ListItems(i).SubItems(1) & "'"
  rs3.Open sql1, cn
  Set rs3 = Nothing
  Set rs = Nothing
 Next
 End If
     sql = "update centry set adate = '" & dte.Value & "',party_name = '" _
           & Text2.Text & "',caccnt = '" _
           & Text3.Text & "',netamt = " _
           & Text16.Text & ",add_desc = '" _
           & Text12.Text & "',add_amt = " _
           & Text13.Text & ",less_desc = '" _
           & Text14.Text & "',less_amt = " _
           & Text15.Text & ",refno = " _
           & Text4.Text & ",narratn = '" _
           & Text10.Text & "' where entryid = " & Text1.Text
      rs.Open sql, cn
      Set rs = Nothing
      Call rec
      rs.MoveLast
      Call disp
Else
  'nothing
End If
addmi.Visible = False
End Sub
Private Sub setcmd()
 add.Top = 200 + title.Height
 add.Left = 300
 edt.Top = add.Top
 edt.Left = add.Left + add.Width
 dlt.Top = add.Top
 dlt.Left = edt.Left + edt.Width
 search.Top = add.Top
 search.Left = dlt.Left + dlt.Width
 first.Top = add.Top
 first.Left = search.Left + search.Width + 600
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
Private Sub Text13_Change()
If Text13.Text <> "" Then
Text16.Text = tmt + CDbl(Text13.Text)
End If
End Sub
Private Sub Text13_LostFocus()
tmt = CDbl(Text16.Text)
End Sub
Private Sub Text15_LostFocus()
tmt = CDbl(Text16.Text)
End Sub
Private Sub Text15_Change()
If Text15.Text <> "" Then
Text16.Text = tmt - CDbl(Text15.Text)
End If
End Sub
Private Sub text2_Change()
If Text3.Locked = False Then
 sql = "select * from accm where name like '" & Text2.Text & "%'"
     rs2.Open sql, cn, adOpenDynamic, adLockOptimistic
     plv.ListItems.Clear
     While Not rs2.EOF
     Set Item = plv.ListItems.add(, , rs2!Name)
     rs2.MoveNext
    Wend
    Set rs2 = Nothing
End If
End Sub
Private Sub text2_GotFocus()
If Text2.Locked = False Then
     plv.Visible = True
     plv.Left = Text2.Left
End If
End Sub
Private Sub text2_LostFocus()
plv.Visible = False
End Sub
Private Sub plv_Click()
Text2.Text = plv.SelectedItem
plv.Left = Me.Width
End Sub
Private Sub text3_Change()
If Text3.Locked = False Then
 sql = "select * from accm where name like '" & Text3.Text & "%'"
     rs2.Open sql, cn, adOpenDynamic, adLockOptimistic
     palv.ListItems.Clear
     While Not rs2.EOF
     Set Item = palv.ListItems.add(, , rs2!Name)
     rs2.MoveNext
    Wend
    Set rs2 = Nothing
End If
End Sub
Private Sub text3_GotFocus()
If Text3.Locked = False Then
     palv.Visible = True
     palv.Left = Text3.Left
End If
End Sub
Private Sub text3_LostFocus()
palv.Visible = False
End Sub
Private Sub palv_Click()
Text3.Text = palv.SelectedItem
palv.Left = Me.Width
End Sub
Private Sub text5_Change()
If Text5.Locked = False Then
 sql = "select * from item where iname like '" & Text5.Text & "%'"
     rs2.Open sql, cn, adOpenDynamic, adLockOptimistic
     ilv.ListItems.Clear
     While Not rs2.EOF
     Set Item = ilv.ListItems.add(, , rs2!iname)
     rs2.MoveNext
    Wend
    Set rs2 = Nothing
End If
End Sub
Private Sub text5_GotFocus()
If Text5.Locked = False Then
     ilv.Visible = True
     ilv.Left = Text5.Left
End If
End Sub
Private Sub text5_LostFocus()
ilv.Visible = False
End Sub
Private Sub ilv_Click()
Text5.Text = ilv.SelectedItem
ilv.Left = Me.Width
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
Private Sub text14_KeyDown(KeyCode As Integer, Shift As Integer)
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
Private Sub text15_KeyDown(KeyCode As Integer, Shift As Integer)
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
Private Sub text16_KeyDown(KeyCode As Integer, Shift As Integer)
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
Private Sub dte_KeyDown(KeyCode As Integer, Shift As Integer)
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
''form shortcut -->(End)
''
''


