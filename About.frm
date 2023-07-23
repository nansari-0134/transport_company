VERSION 5.00
Begin VB.Form About 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9735
   ClientLeft      =   -45
   ClientTop       =   375
   ClientWidth     =   14745
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "About.frx":0000
   ScaleHeight     =   9735
   ScaleWidth      =   14745
   Begin VB.Frame Frame1 
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   7095
      Left            =   3600
      TabIndex        =   0
      Top             =   1800
      Width           =   8055
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   " Email     :   nansari.0134@gmail.com"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   6
         Top             =   6120
         Width           =   5415
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   " Contact :   9399420889"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   5
         Top             =   5640
         Width           =   5415
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   " Name    :   Nizam Ansari   and  T. Shilpa"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   4
         Top             =   5160
         Width           =   5415
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "DEVELOPERS :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   3
         Top             =   4440
         Width           =   2175
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "This software is owned by  M S K Yadu Transportation and is entitled only to Yadu Transport Company."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   360
         TabIndex        =   2
         Top             =   1440
         Width           =   7095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "PROPRIETOR:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   1
         Top             =   600
         Width           =   2175
      End
      Begin VB.Shape Shape1 
         Height          =   7095
         Left            =   0
         Top             =   0
         Width           =   8055
      End
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
      TabIndex        =   8
      Top             =   0
      Width           =   495
   End
   Begin VB.Label title 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "About"
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
      TabIndex        =   7
      Top             =   0
      Width           =   12135
   End
End
Attribute VB_Name = "ABOUT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cl_Click()
menu.Enabled = True
Unload Me
End Sub
Private Sub Form_Load()
0 'form apearance
Call seta
'title setting
Call settl
'frame setting
Call setfr
End Sub
Private Sub seta()
Me.Height = Screen.Height - Screen.Height * 5 / 100 - 450
Me.Width = Screen.Width * 80 / 100
Me.Top = title.Height
Me.Left = Screen.Width * 20 / 100
Me.Picture = LoadPicture(App.Path & "\appdata\images\back.jpg")
End Sub
Private Sub setfr()
Frame1.Left = Me.Width / 2 - Frame1.Width / 2
Frame1.Top = Me.Height / 2 - Frame1.Height / 2
End Sub
Private Sub settl()
title.Height = 450
title.Width = Me.Width
title.Top = 0
title.Left = 0
cl.Top = 0
cl.Left = Me.Width - cl.Width - 50
cl.Height = title.Height - 50
End Sub
