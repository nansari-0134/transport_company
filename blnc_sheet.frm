VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form blnc_sheet 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11325
   ClientLeft      =   15
   ClientTop       =   360
   ClientWidth     =   18420
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   11325
   ScaleWidth      =   18420
   Begin MSChart20Lib.MSChart ch 
      Height          =   3375
      Left            =   13800
      OleObjectBlob   =   "blnc_sheet.frx":0000
      TabIndex        =   104
      Top             =   1440
      Width           =   4935
   End
   Begin VB.ComboBox Combo2 
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
      ItemData        =   "blnc_sheet.frx":24B8
      Left            =   360
      List            =   "blnc_sheet.frx":24CE
      Style           =   2  'Dropdown List
      TabIndex        =   103
      Top             =   7200
      Width           =   2175
   End
   Begin VB.ComboBox Combo1 
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
      ItemData        =   "blnc_sheet.frx":252C
      Left            =   360
      List            =   "blnc_sheet.frx":253F
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label Label100 
      BackStyle       =   0  'Transparent
      Caption         =   "(Rs.)"
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
      Left            =   11160
      TabIndex        =   102
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Label99 
      BackStyle       =   0  'Transparent
      Caption         =   "(Rs.)"
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
      Left            =   9840
      TabIndex        =   101
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label98 
      BackStyle       =   0  'Transparent
      Caption         =   "(Rs.)"
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
      Left            =   8280
      TabIndex        =   100
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label97 
      BackStyle       =   0  'Transparent
      Caption         =   "(Rs.)"
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
      Left            =   5400
      TabIndex        =   99
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Label96 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "2018"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10800
      TabIndex        =   98
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label95 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "2018"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9000
      TabIndex        =   97
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label94 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "2018"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      TabIndex        =   96
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label93 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "2018"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   95
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label92 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash -in-hand"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   375
      Left            =   0
      TabIndex        =   94
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label91 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash -in-hand"
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
      Left            =   0
      TabIndex        =   93
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label90 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash -in-hand"
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
      Left            =   0
      TabIndex        =   92
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label89 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash -in-hand"
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
      Left            =   0
      TabIndex        =   91
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label88 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash -in-hand"
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
      Left            =   0
      TabIndex        =   90
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label87 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash -in-hand"
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
      Left            =   0
      TabIndex        =   89
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label86 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash -in-hand"
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
      Left            =   0
      TabIndex        =   88
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label85 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash -in-hand"
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
      Left            =   0
      TabIndex        =   87
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label84 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash -in-hand"
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
      Left            =   0
      TabIndex        =   86
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label83 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash -in-hand"
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
      Left            =   0
      TabIndex        =   85
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label82 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash -in-hand"
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
      Left            =   0
      TabIndex        =   84
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label81 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash -in-hand"
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
      Left            =   0
      TabIndex        =   83
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label80 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash -in-hand"
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
      Left            =   0
      TabIndex        =   82
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label79 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash -in-hand"
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
      Left            =   0
      TabIndex        =   81
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label78 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash -in-hand"
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
      Left            =   0
      TabIndex        =   80
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label77 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash -in-hand"
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
      Left            =   0
      TabIndex        =   79
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label76 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash -in-hand"
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
      Left            =   0
      TabIndex        =   78
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label75 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash -in-hand"
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
      Left            =   0
      TabIndex        =   77
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label74 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash -in-hand"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   375
      Left            =   0
      TabIndex        =   76
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label73 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash -in-hand"
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
      Height          =   375
      Left            =   0
      TabIndex        =   75
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label72 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash -in-hand"
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
      Height          =   375
      Left            =   0
      TabIndex        =   74
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label71 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash -in-hand"
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
      Height          =   375
      Left            =   0
      TabIndex        =   73
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label70 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash -in-hand"
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
      Height          =   375
      Left            =   0
      TabIndex        =   72
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label69 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash -in-hand"
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
      Height          =   375
      Left            =   0
      TabIndex        =   71
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label68 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash -in-hand"
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
      Height          =   375
      Left            =   0
      TabIndex        =   70
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label67 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash -in-hand"
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
      Height          =   375
      Left            =   0
      TabIndex        =   69
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label66 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash -in-hand"
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
      Height          =   375
      Left            =   0
      TabIndex        =   68
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label65 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash -in-hand"
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
      Height          =   375
      Left            =   0
      TabIndex        =   67
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label64 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash -in-hand"
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
      Height          =   375
      Left            =   0
      TabIndex        =   66
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label63 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash -in-hand"
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
      Height          =   375
      Left            =   0
      TabIndex        =   65
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label62 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash -in-hand"
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
      Height          =   375
      Left            =   0
      TabIndex        =   64
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label61 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash -in-hand"
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
      Height          =   375
      Left            =   0
      TabIndex        =   63
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label60 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash -in-hand"
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
      Height          =   375
      Left            =   0
      TabIndex        =   62
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label59 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash -in-hand"
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
      Height          =   375
      Left            =   0
      TabIndex        =   61
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label58 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash -in-hand"
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
      Height          =   375
      Left            =   0
      TabIndex        =   60
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash -in-hand"
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
      Height          =   375
      Left            =   0
      TabIndex        =   59
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label57 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash -in-hand"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   375
      Left            =   0
      TabIndex        =   58
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label56 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash -in-hand"
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
      Left            =   0
      TabIndex        =   57
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label55 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash -in-hand"
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
      Left            =   0
      TabIndex        =   56
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label54 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash -in-hand"
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
      Left            =   0
      TabIndex        =   55
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label53 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash -in-hand"
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
      Left            =   0
      TabIndex        =   54
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label52 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash -in-hand"
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
      Left            =   0
      TabIndex        =   53
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label51 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash -in-hand"
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
      Left            =   0
      TabIndex        =   52
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label50 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash -in-hand"
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
      Left            =   0
      TabIndex        =   51
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label49 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash -in-hand"
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
      Left            =   0
      TabIndex        =   50
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label48 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash -in-hand"
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
      Left            =   0
      TabIndex        =   49
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label47 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash -in-hand"
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
      Left            =   0
      TabIndex        =   48
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label46 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash -in-hand"
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
      Left            =   0
      TabIndex        =   47
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label45 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash -in-hand"
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
      Left            =   0
      TabIndex        =   46
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label44 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash -in-hand"
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
      Left            =   0
      TabIndex        =   45
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label43 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash -in-hand"
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
      Left            =   0
      TabIndex        =   44
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label42 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash -in-hand"
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
      Left            =   0
      TabIndex        =   43
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label41 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash -in-hand"
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
      Left            =   0
      TabIndex        =   42
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label40 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash -in-hand"
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
      Left            =   0
      TabIndex        =   41
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash -in-hand"
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
      Left            =   0
      TabIndex        =   40
      Top             =   0
      Width           =   1455
   End
   Begin VB.Line Line15 
      X1              =   11760
      X2              =   11760
      Y1              =   600
      Y2              =   8520
   End
   Begin VB.Line Line14 
      X1              =   10920
      X2              =   10920
      Y1              =   960
      Y2              =   8880
   End
   Begin VB.Shape Shape1 
      Height          =   2535
      Left            =   600
      Top             =   1680
      Width           =   6975
   End
   Begin VB.Label Label39 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash -in-hand"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   375
      Left            =   0
      TabIndex        =   39
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label38 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash -in-hand"
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
      Left            =   0
      TabIndex        =   38
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label37 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash -in-hand"
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
      Left            =   0
      TabIndex        =   37
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label36 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash -in-hand"
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
      Left            =   0
      TabIndex        =   36
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label35 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash -in-hand"
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
      Left            =   0
      TabIndex        =   35
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label34 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash -in-hand"
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
      Left            =   0
      TabIndex        =   34
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label33 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash -in-hand"
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
      Left            =   0
      TabIndex        =   33
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label32 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash -in-hand"
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
      Left            =   0
      TabIndex        =   32
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label31 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash -in-hand"
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
      Left            =   0
      TabIndex        =   31
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label30 
      BackStyle       =   0  'Transparent
      Caption         =   "sh"
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
      Left            =   0
      TabIndex        =   30
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label29 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash -in-hand"
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
      Left            =   0
      TabIndex        =   29
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label28 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash -in-hand"
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
      Left            =   0
      TabIndex        =   28
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label27 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash -in-hand"
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
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash -in-hand"
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
      Left            =   0
      TabIndex        =   26
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash -in-hand"
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
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash -in-hand"
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
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash -in-hand"
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
      Left            =   3840
      TabIndex        =   23
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Line Line12 
      X1              =   0
      X2              =   12480
      Y1              =   9360
      Y2              =   9360
   End
   Begin VB.Line Line11 
      X1              =   0
      X2              =   12480
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "RETAINED EARNINGS"
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
      Height          =   375
      Left            =   4440
      TabIndex        =   22
      Top             =   10320
      Width           =   3015
   End
   Begin VB.Line Line9 
      X1              =   0
      X2              =   12480
      Y1              =   10320
      Y2              =   10320
   End
   Begin VB.Line Line8 
      X1              =   0
      X2              =   12480
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Total liabillities :"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   21
      Top             =   6000
      Width           =   1935
   End
   Begin VB.Line Line5 
      X1              =   360
      X2              =   2040
      Y1              =   7080
      Y2              =   7080
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
      Left            =   11880
      TabIndex        =   20
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Assets"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   19
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Fixed assets :"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   18
      Top             =   4800
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Liabilities"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   17
      Top             =   6720
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Total assets :"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   16
      Top             =   6000
      Width           =   1815
   End
   Begin VB.Line Line1 
      X1              =   3360
      X2              =   3360
      Y1              =   2640
      Y2              =   10560
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   12480
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line3 
      X1              =   6120
      X2              =   6120
      Y1              =   480
      Y2              =   8280
   End
   Begin VB.Line Line4 
      X1              =   9480
      X2              =   9480
      Y1              =   480
      Y2              =   8400
   End
   Begin VB.Line Line6 
      X1              =   0
      X2              =   12480
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Line Line7 
      X1              =   240
      X2              =   12720
      Y1              =   5880
      Y2              =   5880
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Current assets :"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Reserves && surplus"
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
      Left            =   360
      TabIndex        =   14
      Top             =   3480
      Width           =   2055
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash -in-hand"
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
      Left            =   360
      TabIndex        =   13
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Stock -in - hand"
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
      Left            =   360
      TabIndex        =   12
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Sundry debitors"
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
      Left            =   360
      TabIndex        =   11
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Revenue"
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
      Left            =   360
      TabIndex        =   10
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Investment"
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
      Left            =   360
      TabIndex        =   9
      Top             =   4200
      Width           =   1815
   End
   Begin VB.Line Line10 
      X1              =   240
      X2              =   12720
      Y1              =   6600
      Y2              =   6600
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Income on Vehicles"
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
      Left            =   480
      TabIndex        =   8
      Top             =   5280
      Width           =   2415
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Duties && taxes"
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
      Left            =   840
      TabIndex        =   7
      Top             =   7800
      Width           =   2535
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Advance to Principles"
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
      Left            =   840
      TabIndex        =   6
      Top             =   8160
      Width           =   2535
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Expenses (Direct)"
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
      Left            =   840
      TabIndex        =   5
      Top             =   8400
      Width           =   2535
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Expenses(Indirect)"
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
      Left            =   360
      TabIndex        =   4
      Top             =   8880
      Width           =   2535
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Misc. expenses"
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
      Left            =   960
      TabIndex        =   3
      Top             =   9240
      Width           =   2535
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Sundry Creditors"
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
      Left            =   1440
      TabIndex        =   2
      Top             =   9600
      Width           =   2535
   End
   Begin VB.Line Line13 
      X1              =   240
      X2              =   12720
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Label title 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Balance sheet"
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
      TabIndex        =   1
      Top             =   120
      Width           =   12375
   End
End
Attribute VB_Name = "blnc_sheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim sql As String
Dim j As Integer
Dim t As Double
Dim i As Integer
Private Sub cl_Click()
menu.Enabled = True
Unload Me
End Sub
Private Sub Combo1_Click()
Set rs = Nothing
For i = Year(Now) To Year(Now) - 3 Step -1
If Combo1.Text = "Bank accounts" Then
   sql = "select sum(capitalacc),sum(salesacc),sum(deposits),sum(loanandadv) from asset where year(adate) = " & i
   rs.Open sql, cn
   If i = Year(Now) Then
   Label28.Caption = CDbl(rs.Fields(0)) + CDbl(rs.Fields(1)) + CDbl(rs.Fields(2)) + CDbl(rs.Fields(3))
   ElseIf i = Year(Now) - 1 Then
   Label46.Caption = CDbl(rs.Fields(0)) + CDbl(rs.Fields(1)) + CDbl(rs.Fields(2)) + CDbl(rs.Fields(3))
   ElseIf i = Year(Now) - 2 Then
   Label63.Caption = CDbl(rs.Fields(0)) + CDbl(rs.Fields(1)) + CDbl(rs.Fields(2)) + CDbl(rs.Fields(3))
   ElseIf i = Year(Now) - 3 Then
   Label81.Caption = CDbl(rs.Fields(0)) + CDbl(rs.Fields(1)) + CDbl(rs.Fields(2)) + CDbl(rs.Fields(3))
   Else
   End If
ElseIf Combo1.Text = "Capital account" Then
    sql = "select sum(capitalacc) from asset where year(adate) = " & i
    rs.Open sql, cn
   If i = Year(Now) Then
   Label28.Caption = CDbl(rs.Fields(0))
   ElseIf i = Year(Now) - 1 Then
   Label46.Caption = CDbl(rs.Fields(0))
   ElseIf i = Year(Now) - 2 Then
   Label63.Caption = CDbl(rs.Fields(0))
   ElseIf i = Year(Now) - 3 Then
   Label81.Caption = CDbl(rs.Fields(0))
   Else
   End If
ElseIf Combo1.Text = "Sales account" Then
    sql = "select sum(salesacc) from asset where year(adate) = " & i
    rs.Open sql, cn
   If i = Year(Now) Then
   Label28.Caption = CDbl(rs.Fields(0))
   ElseIf i = Year(Now) - 1 Then
   Label46.Caption = CDbl(rs.Fields(0))
   ElseIf i = Year(Now) - 2 Then
   Label63.Caption = CDbl(rs.Fields(0))
   ElseIf i = Year(Now) - 3 Then
   Label81.Caption = CDbl(rs.Fields(0))
   Else
   End If
ElseIf Combo1.Text = "Deposits" Then
    sql = "select sum(deposits) from asset where year(adate) = " & i
    rs.Open sql, cn
   If i = Year(Now) Then
   Label28.Caption = CDbl(rs.Fields(0))
   ElseIf i = Year(Now) - 1 Then
   Label46.Caption = CDbl(rs.Fields(0))
   ElseIf i = Year(Now) - 2 Then
   Label63.Caption = CDbl(rs.Fields(0))
   ElseIf i = Year(Now) - 3 Then
   Label81.Caption = CDbl(rs.Fields(0))
   Else
   End If
ElseIf Combo1.Text = "Loan & advances" Then
    sql = "select sum(loanandadv) from asset where year(adate) = " & i
    rs.Open sql, cn
   If i = Year(Now) Then
   Label28.Caption = CDbl(rs.Fields(0))
   ElseIf i = Year(Now) - 1 Then
   Label46.Caption = CDbl(rs.Fields(0))
   ElseIf i = Year(Now) - 2 Then
   Label63.Caption = CDbl(rs.Fields(0))
   ElseIf i = Year(Now) - 3 Then
   Label81.Caption = CDbl(rs.Fields(0))
   Else
   End If
Else
End If
Set rs = Nothing
Next
End Sub

Private Sub Combo2_Click()
Set rs = Nothing
For i = Year(Now) To Year(Now) - 3 Step -1
If Combo2.Text = "Amount payable" Then
   sql = "select sum(bankocc),sum(expacc),sum(loans),sum(secureloan),sum(unsecureloan) from liable where year(ldate) = " & i
   rs.Open sql, cn
   If i = Year(Now) Then
   Label31.Caption = CDbl(rs.Fields(0)) + CDbl(rs.Fields(1)) + CDbl(rs.Fields(2)) + CDbl(rs.Fields(3)) + CDbl(rs.Fields(4))
   ElseIf i = Year(Now) - 1 Then
   Label49.Caption = CDbl(rs.Fields(0)) + CDbl(rs.Fields(1)) + CDbl(rs.Fields(2)) + CDbl(rs.Fields(3)) + CDbl(rs.Fields(4))
   ElseIf i = Year(Now) - 2 Then
   Label66.Caption = CDbl(rs.Fields(0)) + CDbl(rs.Fields(1)) + CDbl(rs.Fields(2)) + CDbl(rs.Fields(3)) + CDbl(rs.Fields(4))
   ElseIf i = Year(Now) - 3 Then
   Label83.Caption = CDbl(rs.Fields(0)) + CDbl(rs.Fields(1)) + CDbl(rs.Fields(2)) + CDbl(rs.Fields(3)) + CDbl(rs.Fields(4))
   Else
   End If
ElseIf Combo2.Text = "Bank OCC A/C" Then
    sql = "select sum(bankocc) from liable where year(adate) = " & i
    rs.Open sql, cn
   If i = Year(Now) Then
   Label31.Caption = CDbl(rs.Fields(0))
   ElseIf i = Year(Now) - 1 Then
   Label49.Caption = CDbl(rs.Fields(0))
   ElseIf i = Year(Now) - 2 Then
   Label66.Caption = CDbl(rs.Fields(0))
   ElseIf i = Year(Now) - 3 Then
   Label83.Caption = CDbl(rs.Fields(0))
   Else
   End If
ElseIf Combo2.Text = "Expenditure account" Then
    sql = "select sum(expacc) from liable where year(adate) = " & i
    rs.Open sql, cn
   If i = Year(Now) Then
   Label31.Caption = CDbl(rs.Fields(0))
   ElseIf i = Year(Now) - 1 Then
   Label49.Caption = CDbl(rs.Fields(0))
   ElseIf i = Year(Now) - 2 Then
   Label66.Caption = CDbl(rs.Fields(0))
   ElseIf i = Year(Now) - 3 Then
   Label83.Caption = CDbl(rs.Fields(0))
   Else
   End If
ElseIf Combo2.Text = "Loans" Then
    sql = "select sum(loans) from liable where year(adate) = " & i
    rs.Open sql, cn
   If i = Year(Now) Then
   Label31.Caption = CDbl(rs.Fields(0))
   ElseIf i = Year(Now) - 1 Then
   Label49.Caption = CDbl(rs.Fields(0))
   ElseIf i = Year(Now) - 2 Then
   Label66.Caption = CDbl(rs.Fields(0))
   ElseIf i = Year(Now) - 3 Then
   Label83.Caption = CDbl(rs.Fields(0))
   Else
   End If
ElseIf Combo2.Text = "Secured loans" Then
    sql = "select sum(secureloan) from liable where year(adate) = " & i
    rs.Open sql, cn
   If i = Year(Now) Then
   Label31.Caption = CDbl(rs.Fields(0))
   ElseIf i = Year(Now) - 1 Then
   Label49.Caption = CDbl(rs.Fields(0))
   ElseIf i = Year(Now) - 2 Then
   Label66.Caption = CDbl(rs.Fields(0))
   ElseIf i = Year(Now) - 3 Then
   Label83.Caption = CDbl(rs.Fields(0))
   Else
   End If
ElseIf Combo2.Text = "unsecured loans" Then
    sql = "select sum(unsecureloan) from liable where year(adate) = " & i
    rs.Open sql, cn
   If i = Year(Now) Then
   Label31.Caption = CDbl(rs.Fields(0))
   ElseIf i = Year(Now) - 1 Then
   Label49.Caption = CDbl(rs.Fields(0))
   ElseIf i = Year(Now) - 2 Then
   Label66.Caption = CDbl(rs.Fields(0))
   ElseIf i = Year(Now) - 3 Then
   Label83.Caption = CDbl(rs.Fields(0))
   Else
   End If
Else
End If
Set rs = Nothing
Next

End Sub

Private Sub Form_Load()
cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data source=" & App.Path & "\appdata\NA2KB_FFC\image.accdb;Jet OLEDB:Database Password=19012019"
'form apearance
Me.Height = Screen.Height - Screen.Height * 5 / 100 - 430
Me.Width = Screen.Width * 80 / 100
Me.Top = 430
Me.Left = Screen.Width * 20 / 100
Me.Picture = LoadPicture(App.Path & "\appdata\images\back.jpg")
Call setfmnu
Call setctrl
Call setlbl
Combo1.Text = "Bank accounts"
Combo2.Text = "Amount payable"
Call setval
Call ret
Call setic
Call setchrt
End Sub
Private Sub setchrt()
With ch
.ShowLegend = True
.ColumnCount = 4
.RowCount = 1
.RowLabel = "Years"
End With
With ch
.Column = 1
.Row = 1
.Data = CDbl(Label39.Caption)
.ColumnLabel = Year(Now)

.Column = 2
.Row = 1
.Data = CDbl(Label57.Caption)
.ColumnLabel = Year(Now) - 1

.Column = 3
.Row = 1
.Data = CDbl(Label74.Caption)
.ColumnLabel = Year(Now) - 2

.Column = 4
.Row = 1
.Data = CDbl(Label92.Caption)
.ColumnLabel = Year(Now) - 3
End With
End Sub
Private Sub setic()
Label97.Top = 530
Label98.Top = 530
Label99.Top = 530
Label100.Top = 530
Label97.Left = Label93.Left + Label93.Width + 350
Label98.Left = Label94.Left + Label93.Width + 350
Label99.Left = Label95.Left + Label93.Width + 350
Label100.Left = Label96.Left + Label93.Width + 350
End Sub
Private Sub setfmnu()
Shape1.Left = 0
Shape1.Top = 430
Shape1.Width = Me.Width
Shape1.Height = Me.Height
title.Height = 430
title.Width = Me.Width
title.Top = 0
title.Left = 0
cl.Top = 0
cl.Left = Me.Width - 600
cl.Height = title.Height - 30
End Sub
Private Sub setctrl()
Line6.X1 = Shape1.Left
Line7.X1 = Shape1.Left
Line10.X1 = Shape1.Left
Line13.X1 = Shape1.Left
Line2.X1 = Shape1.Left
Line1.Y1 = Shape1.Top
Line3.Y1 = Shape1.Top
Line4.Y1 = Shape1.Top
Line1.Y2 = Shape1.Height
Line3.Y2 = Shape1.Height + 430
Line4.Y2 = Shape1.Height + 430
Label1.Left = Shape1.Left + Label1.Width
Label1.Top = title.Height + 70
Label6.Left = Shape1.Left + 70
Label6.Top = Label1.Top + Label1.Height + 70
Label9.Left = 420
Label9.Top = Label6.Top + Label6.Height + 80
Label8.Left = 420
Label8.Top = Label9.Top + Label9.Height + 30
Label7.Left = 420
Label7.Top = Label8.Top + Label8.Height + 30
Label10.Left = 420
Label10.Top = Label7.Top + Label7.Height + 30
Label11.Left = 420
Label11.Top = Label10.Top + Label10.Height + 30
Label12.Left = 420
Label12.Top = Label11.Top + Label11.Height + 30
Combo1.Top = Label12.Top + Label12.Height + 30
Combo1.Left = 420
Label2.Top = Combo1.Top + Combo1.Height + 80
Line6.Y1 = Combo1.Top + Combo1.Height + 30
Line6.Y2 = Combo1.Top + Combo1.Height + 30
Label2.Left = Label6.Left
Label13.Top = Label2.Top + Label2.Height + 30
Label13.Left = 420
Line7.Y1 = Label13.Top + Label13.Height + 30
Line7.Y2 = Label13.Top + Label13.Height + 30
Label5.Left = Label2.Left
Label5.Top = Line7.Y2 + 30
Line10.Y1 = Label5.Top + Label5.Height + 130
Line10.Y2 = Label5.Top + Label5.Height + 130
Line13.Y1 = Line10.Y2 + 20
Line13.Y2 = Line10.Y2 + 20
Label4.Top = Line13.Y2 + 90
Label4.Left = Label1.Left - 200
Line5.Y1 = Label4.Top + Label4.Height + 30
Line5.Y2 = Label4.Top + Label4.Height + 30
Combo2.Left = 420
Combo2.Top = Line5.Y2 + 130
Label14.Left = 420
Label14.Top = Combo2.Top + Combo2.Height + 30
Label15.Left = 420
Label15.Top = Label14.Top + Label14.Height + 30
Label16.Left = 420
Label16.Top = Label15.Top + Label15.Height + 30
Label17.Left = 420
Label17.Top = Label16.Top + Label16.Height + 30
Label18.Left = 420
Label18.Top = Label17.Top + Label17.Height + 30
Label19.Left = 420
Label19.Top = Label18.Top + Label18.Height + 30
Line9.X1 = 0
Line9.Y1 = Label19.Top + Label19.Height + 30
Line9.Y2 = Line9.Y1
Line8.X1 = 0
Line8.Y1 = Label19.Top + Label19.Height + 50
Line8.Y2 = Line8.Y1
Label20.Left = Label5.Left
Label20.Top = Label19.Top + Label19.Height + 150
Line11.X1 = 0
Line11.Y1 = Label20.Top + Label20.Height + 80
Line11.Y2 = Line11.Y1
Line12.X1 = 0
Line12.Y1 = Line11.Y1 + 20
Line12.Y2 = Line12.Y1
Label21.Left = Label5.Left
Label21.Top = Label20.Top + Label20.Height + 230
Line1.X1 = Label21.Left + Label21.Width + 10
Line1.X2 = Line1.X1
Line5.X1 = 0
Line5.X2 = Line1.X1
Line3.X1 = Line1.X1 + Me.Width * 15 / 100
Line3.X2 = Line3.X1
Line4.X1 = Line3.X1 + Me.Width * 15 / 100
Line4.X2 = Line4.X1
Line14.X1 = Line4.X1 + Me.Width * 15 / 100
Line14.X2 = Line14.X1
Line14.Y1 = Shape1.Top
Line14.Y2 = Shape1.Height
Line15.X1 = Line14.X1 + Me.Width * 15 / 100
Line15.X2 = Line15.X1
Line15.Y1 = Shape1.Top
Line15.Y2 = Shape1.Height
Line2.X2 = Line15.X1
Line6.X2 = Line15.X1
Line7.X2 = Line15.X1
Line13.X2 = Line15.X1
Line9.X2 = Line15.X1
Line10.X2 = Line15.X1
Line8.X2 = Line15.X1
Line11.X2 = Line15.X1
Line12.X2 = Line15.X1

End Sub
Private Sub setlbl()
Label3.Caption = ""
Label22.Caption = ""
Label23.Caption = ""
Label24.Caption = ""
Label25.Caption = ""
Label26.Caption = ""
Label27.Caption = ""
Label28.Caption = ""
Label29.Caption = ""
Label30.Caption = ""
Label31.Caption = ""
Label32.Caption = ""
Label33.Caption = ""
Label34.Caption = ""
Label35.Caption = ""
Label36.Caption = ""
Label37.Caption = ""
Label38.Caption = ""
Label39.Caption = ""
Label40.Caption = ""
Label41.Caption = ""
Label42.Caption = ""
Label43.Caption = ""
Label44.Caption = ""
Label45.Caption = ""
Label46.Caption = ""
Label47.Caption = ""
Label48.Caption = ""
Label49.Caption = ""
Label50.Caption = ""
Label51.Caption = ""
Label52.Caption = ""
Label53.Caption = ""
Label54.Caption = ""
Label55.Caption = ""
Label56.Caption = ""
Label57.Caption = ""
Label58.Caption = ""
Label59.Caption = ""
Label60.Caption = ""
Label61.Caption = ""
Label62.Caption = ""
Label63.Caption = ""
Label64.Caption = ""
Label65.Caption = ""
Label66.Caption = ""
Label67.Caption = ""
Label68.Caption = ""
Label69.Caption = ""
Label70.Caption = ""
Label71.Caption = ""
Label72.Caption = ""
Label73.Caption = ""
Label74.Caption = ""
Label75.Caption = ""
Label76.Caption = ""
Label77.Caption = ""
Label78.Caption = ""
Label79.Caption = ""
Label80.Caption = ""
Label81.Caption = ""
Label82.Caption = ""
Label83.Caption = ""
Label84.Caption = ""
Label85.Caption = ""
Label86.Caption = ""
Label87.Caption = ""
Label88.Caption = ""
Label89.Caption = ""
Label90.Caption = ""
Label91.Caption = ""
Label92.Caption = ""

Label22.Top = Label9.Top
Label22.Left = Line1.X1 + 100
Label22.Width = Me.Width * 14 / 100
Label22.Alignment = 1

Label23.Top = Label8.Top
Label23.Left = Label22.Left
Label23.Width = Label22.Width
Label23.Alignment = 1
''['''
Label24.Top = Label7.Top
Label24.Left = Label22.Left
Label24.Width = Label22.Width
Label24.Alignment = 1

Label25.Top = Label10.Top
Label25.Left = Label22.Left
Label25.Width = Label22.Width
Label25.Alignment = 1

Label26.Top = Label11.Top
Label26.Left = Label22.Left
Label26.Width = Label22.Width
Label26.Alignment = 1

Label27.Top = Label12.Top
Label27.Left = Label22.Left
Label27.Width = Label22.Width
Label27.Alignment = 1

Label28.Top = Combo1.Top
Label28.Left = Label22.Left
Label28.Width = Label22.Width
Label28.Alignment = 1

Label29.Top = Label13.Top
Label29.Left = Label22.Left
Label29.Width = Label22.Width
Label29.Alignment = 1

Label30.Top = Label5.Top
Label30.Left = Label22.Left
Label30.Width = Label22.Width
Label30.Alignment = 1

Label31.Top = Combo2.Top
Label31.Left = Label22.Left
Label31.Width = Label22.Width
Label31.Alignment = 1

Label32.Top = Label14.Top
Label32.Left = Label22.Left
Label32.Width = Label22.Width
Label32.Alignment = 1

Label33.Top = Label15.Top
Label33.Left = Label22.Left
Label33.Width = Label22.Width
Label33.Alignment = 1

Label34.Top = Label16.Top
Label34.Left = Label22.Left
Label34.Width = Label22.Width
Label34.Alignment = 1

Label35.Top = Label17.Top
Label35.Left = Label22.Left
Label35.Width = Label22.Width
Label35.Alignment = 1

Label36.Top = Label18.Top
Label36.Left = Label22.Left
Label36.Width = Label22.Width
Label36.Alignment = 1

Label37.Top = Label19.Top
Label37.Left = Label22.Left
Label37.Width = Label22.Width
Label37.Alignment = 1

Label38.Top = Label20.Top
Label38.Left = Label22.Left
Label38.Width = Label22.Width
Label38.Alignment = 1

Label39.Top = Label21.Top
Label39.Left = Label22.Left
Label39.Width = Label22.Width
Label39.Alignment = 1

Label40.Top = Label22.Top
Label40.Left = Line3.X1 + 100
Label40.Width = Label22.Width
Label40.Alignment = 1

Label41.Top = Label8.Top
Label41.Left = Label40.Left
Label41.Width = Label22.Width
Label41.Alignment = 1

Label42.Top = Label7.Top
Label42.Left = Label40.Left
Label42.Width = Label22.Width
Label42.Alignment = 1

Label43.Top = Label10.Top
Label43.Left = Label40.Left
Label43.Width = Label22.Width
Label43.Alignment = 1

Label44.Top = Label11.Top
Label44.Left = Label40.Left
Label44.Width = Label22.Width
Label44.Alignment = 1

Label45.Top = Label12.Top
Label45.Left = Label40.Left
Label45.Width = Label22.Width
Label45.Alignment = 1

Label46.Top = Label28.Top
Label46.Left = Label40.Left
Label46.Width = Label22.Width
Label46.Alignment = 1

Label47.Top = Label29.Top
Label47.Left = Label40.Left
Label47.Width = Label22.Width
Label47.Alignment = 1

Label48.Top = Label30.Top
Label48.Left = Label40.Left
Label48.Width = Label22.Width
Label48.Alignment = 1

Label49.Top = Label31.Top
Label49.Left = Label40.Left
Label49.Width = Label22.Width
Label49.Alignment = 1

Label50.Top = Label32.Top
Label50.Left = Label40.Left
Label50.Width = Label22.Width
Label50.Alignment = 1

Label51.Top = Label33.Top
Label51.Left = Label40.Left
Label51.Width = Label22.Width
Label51.Alignment = 1

Label52.Top = Label34.Top
Label52.Left = Label40.Left
Label52.Width = Label22.Width
Label52.Alignment = 1

Label53.Top = Label35.Top
Label53.Left = Label40.Left
Label53.Width = Label22.Width
Label53.Alignment = 1

Label54.Top = Label36.Top
Label54.Left = Label40.Left
Label54.Width = Label22.Width
Label54.Alignment = 1

Label55.Top = Label37.Top
Label55.Left = Label40.Left
Label55.Width = Label22.Width
Label55.Alignment = 1

Label56.Top = Label38.Top
Label56.Left = Label40.Left
Label56.Width = Label22.Width
Label56.Alignment = 1

Label57.Top = Label39.Top
Label57.Left = Label40.Left
Label57.Width = Label22.Width
Label57.Alignment = 1

Label3.Top = Label22.Top
Label3.Left = Line4.X1 + 100
Label3.Width = Label22.Width
Label3.Alignment = 1

Label58.Top = Label23.Top
Label58.Left = Label3.Left
Label58.Width = Label22.Width
Label58.Alignment = 1

Label59.Top = Label24.Top
Label59.Left = Label3.Left
Label59.Width = Label22.Width
Label59.Alignment = 1

Label60.Top = Label25.Top
Label60.Left = Label3.Left
Label60.Width = Label22.Width
Label60.Alignment = 1

Label61.Top = Label26.Top
Label61.Left = Label3.Left
Label61.Width = Label22.Width
Label61.Alignment = 1

Label62.Top = Label27.Top
Label62.Left = Label3.Left
Label62.Width = Label22.Width
Label62.Alignment = 1

Label63.Top = Label28.Top
Label63.Left = Label3.Left
Label63.Width = Label22.Width
Label63.Alignment = 1

Label64.Top = Label29.Top
Label64.Left = Label3.Left
Label64.Width = Label22.Width
Label64.Alignment = 1

Label65.Top = Label30.Top
Label65.Left = Label3.Left
Label65.Width = Label22.Width
Label65.Alignment = 1

Label66.Top = Label31.Top
Label66.Left = Label3.Left
Label66.Width = Label22.Width
Label66.Alignment = 1

Label67.Top = Label32.Top
Label67.Left = Label3.Left
Label67.Width = Label22.Width
Label67.Alignment = 1

Label68.Top = Label33.Top
Label68.Left = Label3.Left
Label68.Width = Label22.Width
Label68.Alignment = 1

Label69.Top = Label34.Top
Label69.Left = Label3.Left
Label69.Width = Label22.Width
Label69.Alignment = 1

Label70.Top = Label35.Top
Label70.Left = Label3.Left
Label70.Width = Label22.Width
Label70.Alignment = 1

Label71.Top = Label36.Top
Label71.Left = Label3.Left
Label71.Width = Label22.Width
Label71.Alignment = 1

Label72.Top = Label37.Top
Label72.Left = Label3.Left
Label72.Width = Label22.Width
Label72.Alignment = 1

Label73.Top = Label38.Top
Label73.Left = Label3.Left
Label73.Width = Label22.Width
Label73.Alignment = 1

Label74.Top = Label39.Top
Label74.Left = Label3.Left
Label74.Width = Label22.Width
Label74.Alignment = 1

Label75.Top = Label22.Top
Label75.Left = Line14.X1 + 100
Label75.Width = Label22.Width
Label75.Alignment = 1

Label76.Top = Label23.Top
Label76.Left = Label75.Left
Label76.Width = Label22.Width
Label76.Alignment = 1

Label77.Top = Label24.Top
Label77.Left = Label75.Left
Label77.Width = Label22.Width
Label77.Alignment = 1

Label78.Top = Label25.Top
Label78.Left = Label75.Left
Label78.Width = Label22.Width
Label78.Alignment = 1

Label79.Top = Label26.Top
Label79.Left = Label75.Left
Label79.Width = Label22.Width
Label79.Alignment = 1

Label80.Top = Label27.Top
Label80.Left = Label75.Left
Label80.Width = Label22.Width
Label80.Alignment = 1

Label81.Top = Label28.Top
Label81.Left = Label75.Left
Label81.Width = Label22.Width
Label81.Alignment = 1

Label82.Top = Label29.Top
Label82.Left = Label75.Left
Label82.Width = Label22.Width
Label82.Alignment = 1

Label83.Top = Label30.Top
Label83.Left = Label75.Left
Label83.Width = Label22.Width
Label83.Alignment = 1

Label84.Top = Label31.Top
Label84.Left = Label75.Left
Label84.Width = Label22.Width
Label84.Alignment = 1

Label85.Top = Label32.Top
Label85.Left = Label75.Left
Label85.Width = Label22.Width
Label85.Alignment = 1

Label86.Top = Label33.Top
Label86.Left = Label75.Left
Label86.Width = Label22.Width
Label86.Alignment = 1

Label87.Top = Label34.Top
Label87.Left = Label75.Left
Label87.Width = Label22.Width
Label87.Alignment = 1

Label88.Top = Label35.Top
Label88.Left = Label75.Left
Label88.Width = Label22.Width
Label88.Alignment = 1

Label89.Top = Label36.Top
Label89.Left = Label75.Left
Label89.Width = Label22.Width
Label89.Alignment = 1

Label90.Top = Label37.Top
Label90.Left = Label75.Left
Label90.Width = Label22.Width
Label90.Alignment = 1

Label91.Top = Label38.Top
Label91.Left = Label75.Left
Label91.Width = Label22.Width
Label91.Alignment = 1

Label92.Top = Label39.Top
Label92.Left = Label75.Left
Label92.Width = Label22.Width
Label92.Alignment = 1

Label94.Top = Label1.Top
Label95.Top = Label1.Top
Label96.Top = Label1.Top
Label93.Top = Label1.Top
Label93.Left = Label22.Left
Label94.Left = Label40.Left
Label95.Left = Label3.Left
Label96.Left = Label75.Left
Label93.Caption = Year(Now)
Label94.Caption = Year(Now) - 1
Label95.Caption = Year(Now) - 2
Label96.Caption = Year(Now) - 3
End Sub
Private Sub setval()
For i = Year(Now) To Year(Now) - 3 Step -1
Set rs = Nothing
  'sql = "select sum(cashinhand,stockinhand,sundrydebt,revandsur,revenue,investment,capitalacc,salesacc,deposits,loanandadv,vehicles) from asset where year(adate) = " & i
  sql = "select sum(cashinhand),sum(stockinhand),sum(sundrydebt),sum(revandsur),sum(revenue),sum(investment),sum(capitalacc),sum(salesacc),sum(deposits),sum(loanandadv),sum(vehicles) from asset where year(adate) = " & i
  rs.Open sql, cn
  If i = Year(Now) Then
        Label22.Caption = rs.Fields(0)
        Label23.Caption = rs.Fields(1)
        Label24.Caption = rs.Fields(2)
        Label25.Caption = rs.Fields(3)
        Label26.Caption = rs.Fields(4)
        Label27.Caption = rs.Fields(5)
        Label28.Caption = rs.Fields(6) + rs.Fields(7) + rs.Fields(8) + rs.Fields(9)
        Label29.Caption = rs.Fields(10)
  ElseIf i = Year(Now) - 1 Then
      Label40.Caption = rs.Fields(0)
      Label41.Caption = rs.Fields(1)
      Label42.Caption = rs.Fields(2)
      Label43.Caption = rs.Fields(3)
      Label44.Caption = rs.Fields(4)
      Label45.Caption = rs.Fields(5)
      Label46.Caption = rs.Fields(6) + rs.Fields(7) + rs.Fields(8) + rs.Fields(9)
      Label47.Caption = rs.Fields(10)

  ElseIf i = Year(Now) - 2 Then
      Label3.Caption = rs.Fields(0)
      Label58.Caption = rs.Fields(1)
      Label59.Caption = rs.Fields(2)
      Label60.Caption = rs.Fields(3)
      Label61.Caption = rs.Fields(4)
      Label62.Caption = rs.Fields(5)
     Label63.Caption = rs.Fields(6) + rs.Fields(7) + rs.Fields(8) + rs.Fields(9)
     Label64.Caption = rs.Fields(10)
  ElseIf i = Year(Now) - 3 Then
      Label75.Caption = rs.Fields(0)
      Label76.Caption = rs.Fields(1)
      Label77.Caption = rs.Fields(2)
      Label78.Caption = rs.Fields(3)
      Label79.Caption = rs.Fields(4)
      Label80.Caption = rs.Fields(5)
      Label81.Caption = rs.Fields(6) + rs.Fields(7) + rs.Fields(8) + rs.Fields(9)
      Label82.Caption = rs.Fields(10)
  Else
  End If
Next
Label30.Caption = CDbl(Label22.Caption) + CDbl(Label23.Caption) + CDbl(Label24.Caption) + CDbl(Label25.Caption) + CDbl(Label26.Caption) + CDbl(Label27.Caption) + CDbl(Label28.Caption) + CDbl(Label29.Caption)
Label48.Caption = CDbl(Label40.Caption) + CDbl(Label41.Caption) + CDbl(Label42.Caption) + CDbl(Label43.Caption) + CDbl(Label44.Caption) + CDbl(Label45.Caption) + CDbl(Label46.Caption) + CDbl(Label47.Caption)
Label65.Caption = CDbl(Label3.Caption) + CDbl(Label58.Caption) + CDbl(Label59.Caption) + CDbl(Label60.Caption) + CDbl(Label61.Caption) + CDbl(Label62.Caption) + CDbl(Label63.Caption) + CDbl(Label64.Caption)
Label83.Caption = CDbl(Label75.Caption) + CDbl(Label76.Caption) + CDbl(Label77.Caption) + CDbl(Label78.Caption) + CDbl(Label79.Caption) + CDbl(Label80.Caption) + CDbl(Label81.Caption) + CDbl(Label82.Caption)
Set rs = Nothing

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

For i = Year(Now) To Year(Now) - 3 Step -1
Set rs = Nothing
   sql = "select sum(tax),sum(advance),sum(dexpe),sum(iexpe),sum(miscexpe),sum(sundrycreditor), sum(bankocc),sum(expacc),sum(loans),sum(secureloan),sum(unsecureloan) from liable where year(ldate) = " & i
   rs.Open sql, cn
 If i = Year(Now) Then
  Label31.Caption = rs.Fields(6) + rs.Fields(7) + rs.Fields(8) + rs.Fields(9) + rs.Fields(10)
  Label32.Caption = rs.Fields(0)
  Label33.Caption = rs.Fields(1)
  Label34.Caption = rs.Fields(2)
  Label35.Caption = rs.Fields(3)
  Label36.Caption = rs.Fields(4)
  Label37.Caption = rs.Fields(5)
ElseIf i = Year(Now) - 1 Then
  Label49.Caption = rs.Fields(6) + rs.Fields(7) + rs.Fields(8) + rs.Fields(9) + rs.Fields(10)
  Label50.Caption = rs.Fields(0)
  Label51.Caption = rs.Fields(1)
  Label52.Caption = rs.Fields(2)
  Label53.Caption = rs.Fields(3)
  Label54.Caption = rs.Fields(4)
  Label55.Caption = rs.Fields(5)
ElseIf i = Year(Now) - 2 Then
 Label66.Caption = rs.Fields(6) + rs.Fields(7) + rs.Fields(8) + rs.Fields(9) + rs.Fields(10)
  Label67.Caption = rs.Fields(0)
  Label68.Caption = rs.Fields(1)
  Label69.Caption = rs.Fields(2)
  Label70.Caption = rs.Fields(3)
  Label71.Caption = rs.Fields(4)
  Label72.Caption = rs.Fields(5)
ElseIf i = Year(Now) - 3 Then
  Label84.Caption = rs.Fields(6) + rs.Fields(7) + rs.Fields(8) + rs.Fields(9) + rs.Fields(10)
  Label85.Caption = rs.Fields(0)
  Label86.Caption = rs.Fields(1)
  Label87.Caption = rs.Fields(2)
  Label88.Caption = rs.Fields(3)
  Label89.Caption = rs.Fields(4)
  Label90.Caption = rs.Fields(5)
   Else
  End If
Next
Label38.Caption = CDbl(Label31.Caption) + CDbl(Label32.Caption) + CDbl(Label33.Caption) + CDbl(Label34.Caption) + CDbl(Label35.Caption) + CDbl(Label36.Caption) + CDbl(Label37.Caption)
Label56.Caption = CDbl(Label49.Caption) + CDbl(Label50.Caption) + CDbl(Label51.Caption) + CDbl(Label52.Caption) + CDbl(Label53.Caption) + CDbl(Label54.Caption) + CDbl(Label55.Caption)
Label73.Caption = CDbl(Label66.Caption) + CDbl(Label67.Caption) + CDbl(Label68.Caption) + CDbl(Label69.Caption) + CDbl(Label70.Caption) + CDbl(Label71.Caption) + CDbl(Label72.Caption)
Label91.Caption = CDbl(Label84.Caption) + CDbl(Label85.Caption) + CDbl(Label86.Caption) + CDbl(Label87.Caption) + CDbl(Label88.Caption) + CDbl(Label89.Caption) + CDbl(Label90.Caption)
Set rs = Nothing
End Sub
Private Sub ret()
Label39.Caption = CDbl(Label30.Caption) - CDbl(Label38.Caption)
Label57.Caption = CDbl(Label48.Caption) - CDbl(Label56.Caption)
Label74.Caption = CDbl(Label65.Caption) - CDbl(Label73.Caption)
Label92.Caption = CDbl(Label83.Caption) - CDbl(Label91.Caption)
End Sub

