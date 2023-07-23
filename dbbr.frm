VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form dbbr 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   10155
   ClientLeft      =   15
   ClientTop       =   435
   ClientWidth     =   14745
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10155
   ScaleWidth      =   14745
   Begin MSComDlg.CommonDialog cd1 
      Left            =   11400
      Top             =   4680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      Height          =   4455
      Left            =   1680
      TabIndex        =   6
      Top             =   4440
      Width           =   8655
      Begin VB.CommandButton process1 
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
         TabIndex        =   19
         Top             =   3720
         Width           =   1455
      End
      Begin VB.TextBox Text5 
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
         Left            =   3240
         TabIndex        =   17
         Top             =   1680
         Width           =   3855
      End
      Begin VB.TextBox Text4 
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
         Left            =   3240
         TabIndex        =   16
         Top             =   2280
         Width           =   3855
      End
      Begin VB.TextBox Text3 
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
         Left            =   3240
         TabIndex        =   15
         Top             =   2880
         Width           =   3855
      End
      Begin VB.CommandButton canc1 
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
         Picture         =   "dbbr.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   3720
         Width           =   1455
      End
      Begin VB.TextBox Text2 
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
         Left            =   3240
         TabIndex        =   7
         Top             =   840
         Width           =   3855
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   " file Date  :"
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
         Left            =   1320
         TabIndex        =   14
         Top             =   2880
         Width           =   2055
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   " file Size   :"
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
         Left            =   1320
         TabIndex        =   13
         Top             =   2280
         Width           =   2055
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   " file Name :"
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
         Left            =   1320
         TabIndex        =   12
         Top             =   1680
         Width           =   2055
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Select backup file :"
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
         Left            =   240
         TabIndex        =   9
         Top             =   840
         Width           =   2895
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Data Restore"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   8
         Top             =   90
         Width           =   2535
      End
      Begin VB.Shape Shape2 
         Height          =   4455
         Left            =   0
         Top             =   0
         Width           =   8655
      End
      Begin VB.Line Line2 
         X1              =   0
         X2              =   8640
         Y1              =   480
         Y2              =   480
      End
   End
   Begin VB.Frame party 
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      Height          =   2655
      Left            =   1680
      TabIndex        =   2
      Top             =   1560
      Width           =   8655
      Begin VB.CommandButton process 
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
         TabIndex        =   18
         Top             =   1920
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
         Height          =   550
         Left            =   5400
         Picture         =   "dbbr.frx":37B1
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1920
         Width           =   1455
      End
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
         Left            =   3240
         TabIndex        =   3
         Top             =   840
         Width           =   3855
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   8640
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Shape Shape1 
         Height          =   2655
         Left            =   0
         Top             =   0
         Width           =   8655
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Data Backup"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   5
         Top             =   90
         Width           =   2535
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Select backup folder:"
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
         Left            =   240
         TabIndex        =   4
         Top             =   840
         Width           =   2895
      End
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
      Left            =   11640
      TabIndex        =   1
      Top             =   0
      Width           =   495
   End
   Begin VB.Label title 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "         Data Backup And Restore        "
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
Attribute VB_Name = "DBBR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
'form apearance
Me.Height = Screen.Height - Screen.Height * 5 / 100 - 450
Me.Width = Screen.Width * 40 / 100
Me.Top = 450
Me.Left = Screen.Width * 20 / 100
Me.Picture = LoadPicture(App.Path & "\appdata\images\back.jpg")
 cancel.Picture = LoadPicture(App.Path & "\appdata\icons\cmd_cancel.jpg")
 canc1.Picture = LoadPicture(App.Path & "\appdata\icons\cmd_cancel.jpg")
 process.Picture = LoadPicture(App.Path & "\appdata\icons\cmd_process.jpg")
 process1.Picture = LoadPicture(App.Path & "\appdata\icons\cmd_process.jpg")

'title bar and close button setting
 Call setfmnu
 'seting contents
 Call cont
 End Sub

Private Sub setfmnu()
title.Height = 450
title.Width = Me.Width
title.Top = 0
title.Left = 0
cl.Top = 0
cl.Left = Me.Width - 600
cl.Height = title.Height - 50
End Sub
Private Sub cont()
party.Top = 500 + title.Height
party.Left = 200
party.Width = Me.Width - 500
Shape1.Width = party.Width
Line1.X2 = party.Width
Shape1.Height = party.Height
Frame1.Top = party.Top + party.Height + 500
Frame1.Left = party.Left
Frame1.Width = Me.Width - 500
Shape2.Width = Frame1.Width
Line2.X2 = Frame1.Width

End Sub

Private Sub cl_Click()
menu.Enabled = True
Unload Me
End Sub

Private Sub Text1_Click()
cd1.ShowOpen
Text1.Text = cd1.FileName
End Sub
