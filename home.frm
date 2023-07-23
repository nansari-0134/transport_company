VERSION 5.00
Begin VB.Form menu 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11055
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   20460
   ControlBox      =   0   'False
   Icon            =   "home.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   11055
   ScaleWidth      =   20460
   Begin VB.Frame ramnu 
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   8280
      TabIndex        =   36
      Top             =   1080
      Visible         =   0   'False
      Width           =   2655
      Begin VB.Label rabs 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Balance Sheet               "
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   38
         Top             =   0
         Width           =   2655
      End
      Begin VB.Label rais 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Income Statement        "
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   37
         Top             =   480
         Width           =   2655
      End
   End
   Begin VB.Frame mmnu 
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   4320
      TabIndex        =   30
      Top             =   2040
      Visible         =   0   'False
      Width           =   3255
      Begin VB.Label mim 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Item Master                    "
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   33
         Top             =   960
         Width           =   3255
      End
      Begin VB.Label marm 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Area Master                    "
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   32
         Top             =   480
         Width           =   3255
      End
      Begin VB.Label macm 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Account Master              "
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   31
         Top             =   0
         Width           =   3255
      End
   End
   Begin VB.Frame tmnu 
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      Height          =   2415
      Left            =   5640
      TabIndex        =   24
      Top             =   5880
      Visible         =   0   'False
      Width           =   3255
      Begin VB.Label tce 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Consumption Entry        "
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   29
         Top             =   1920
         Width           =   3255
      End
      Begin VB.Label tdte 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Daily Transaction Entry  "
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   28
         Top             =   0
         Width           =   3255
      End
      Begin VB.Label ttig 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Trip Invoice Genegration"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   27
         Top             =   480
         Width           =   3255
      End
      Begin VB.Label tve 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Voucher Entry                "
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   26
         Top             =   960
         Width           =   3255
      End
      Begin VB.Label tpe 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Purchase Entry              "
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   25
         Top             =   1440
         Width           =   3255
      End
   End
   Begin VB.Frame rmnu 
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      Height          =   5295
      Left            =   9240
      TabIndex        =   12
      Top             =   3120
      Visible         =   0   'False
      Width           =   3615
      Begin VB.Label rebao 
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "    Entry Based Accounting Outstanding"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   23
         Top             =   4800
         Width           =   3615
      End
      Begin VB.Label rvl 
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "    Vehicle Ledger"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   22
         Top             =   4320
         Width           =   3615
      End
      Begin VB.Label rss 
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "    Stock Summary"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   21
         Top             =   3840
         Width           =   3615
      End
      Begin VB.Label ral 
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "    Account Ledger"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   20
         Top             =   3360
         Width           =   3615
      End
      Begin VB.Label rvts 
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "    Vehicle Trip Summary"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   19
         Top             =   2880
         Width           =   3615
      End
      Begin VB.Label rpwiwsr 
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "    Party Wise Item Wise sales Report"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   18
         Top             =   2400
         Width           =   3615
      End
      Begin VB.Label rtds 
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "    Trip Driver Summary"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   17
         Top             =   1920
         Width           =   3615
      End
      Begin VB.Label rtr 
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "    Trip Record "
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   16
         Top             =   0
         Width           =   3615
      End
      Begin VB.Label rtrtir 
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "    Trip Pending To Invoice Register       "
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   15
         Top             =   480
         Width           =   3615
      End
      Begin VB.Label rir 
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "    Invoice Register"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   14
         Top             =   960
         Width           =   3615
      End
      Begin VB.Label rvr 
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "    Voucher Register"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   13
         Top             =   1440
         Width           =   3615
      End
   End
   Begin VB.Frame umnu 
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   13080
      TabIndex        =   7
      Top             =   7680
      Visible         =   0   'False
      Width           =   2655
      Begin VB.Label uad 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Help                            "
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   11
         Top             =   1440
         Width           =   2655
      End
      Begin VB.Label udr 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "About                           "
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   10
         Top             =   960
         Width           =   2655
      End
      Begin VB.Label udb 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "   Data Backup And Restore     "
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   9
         Top             =   480
         Width           =   2655
      End
      Begin VB.Label ucp 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Change Password          "
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   2655
      End
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "                Revenue Analysis        "
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   0
      TabIndex        =   35
      Top             =   7800
      Width           =   5055
   End
   Begin VB.Label ourname 
      BackStyle       =   0  'Transparent
      Height          =   1215
      Left            =   0
      TabIndex        =   34
      Top             =   9720
      Width           =   4335
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   19320
      TabIndex        =   6
      Top             =   0
      Width           =   495
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   19920
      TabIndex        =   5
      Top             =   0
      Width           =   495
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   " M.S.K. YADU TRANSPORT"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   20295
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Utilities          "
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   0
      TabIndex        =   3
      Top             =   6840
      Width           =   5055
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "       Transactions       "
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   0
      TabIndex        =   2
      Top             =   5760
      Width           =   5055
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Reports        "
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   0
      TabIndex        =   1
      Top             =   4680
      Width           =   5055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Master           "
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   3600
      Width           =   5055
   End
   Begin VB.Image mnucntr 
      Height          =   11295
      Left            =   -600
      Top             =   -240
      Width           =   5055
   End
End
Attribute VB_Name = "menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
w = 0
   If Shift = vbCtrlMask And KeyCode = vbKeyA Then
     acmaster.Show
     Me.Enabled = False
   ElseIf Shift = vbCtrlMask + vbShiftMask And KeyCode = vbKeyA Then
     armaster.Show
     Me.Enabled = False
   ElseIf Shift = vbCtrlMask And KeyCode = vbKeyI Then
     itmaster.Show
     Me.Enabled = False
   ElseIf Shift = vbCtrlMask And KeyCode = vbKeyT Then
    trrecord.Show
   ElseIf Shift = vbCtrlMask And KeyCode = vbKeyY Then
     trtig.Show
     Me.Enabled = False
   ElseIf Shift = vbCtrlMask + vbShiftMask And KeyCode = vbKeyI Then
    inrgstr.Show
    Me.Enabled = False
   ElseIf Shift = vbCtrlMask And KeyCode = vbKeyV Then
    vrgstr.Show
    Me.Enabled = False
   ElseIf Shift = vbCtrlMask And KeyCode = vbKeyD Then
     tdsum.Show
     Me.Enabled = False
   ElseIf Shift = vbCtrlMask And KeyCode = vbKeyW Then
     pwiwsrprt.Show
     Me.Enabled = False
   ElseIf Shift = vbCtrlMask And KeyCode = vbKeyS Then
     vtsum.Show
     Me.Enabled = False
   ElseIf Shift = vbCtrlMask And KeyCode = vbKeyL Then
     acldgr.Show
     Me.Enabled = False
   ElseIf Shift = vbCtrlMask And KeyCode = vbKeyK Then
     stsum.Show
     Me.Enabled = False
   ElseIf Shift = vbCtrlMask And KeyCode = vbKeyG Then
     vldgr.Show
     Me.Enabled = False
   ElseIf Shift = vbCtrlMask And KeyCode = vbKeyE Then
     ebao.Show
     Me.Enabled = False
   ElseIf Shift = vbCtrlMask + vbShiftMask And KeyCode = vbKeyD Then
     dtentry.Show
     Me.Enabled = False
   ElseIf Shift = vbCtrlMask And KeyCode = vbKeyM Then
     tigen.Show
     Me.Enabled = False
   ElseIf Shift = vbCtrlMask + vbShiftMask And KeyCode = vbKeyV Then
     ventry.Show
     Me.Enabled = False
   ElseIf Shift = vbCtrlMask And KeyCode = vbKeyP Then
     pentry.Show
     Me.Enabled = False
   ElseIf Shift = vbCtrlMask And KeyCode = vbKeyC Then
     centry.Show
     Me.Enabled = False
   ElseIf Shift = vbCtrlMask And KeyCode = vbKeyB Then
     DBBR.Show
     Me.Enabled = False
   ElseIf Shift = vbCtrlMask And KeyCode = vbKeyH Then
     help.Show
     Me.Enabled = False
   ElseIf Shift = vbCtrlMask + vbShiftMask And KeyCode = vbKeyP Then
     chpass.Show
     Me.Enabled = False
   ElseIf Shift = vbCtrlMask And KeyCode = vbKeyQ Then
     i = MsgBox("Do You Want To Quit Apllication", vbOKCancel + vbCritical, "Exit")
     If i = vbOK Then
      End
     End If
   Else
        'nothing
 End If
End Sub
Private Sub Form_Load()
w1 = 0
w2 = 0
ourname.Left = 0
ourname.Top = Screen.Height - Screen.Height * 12 / 100
ourname.Caption = "Devlopers:- " & vbNewLine & "Nizam Ansari     9399420889 " & vbNewLine & "T. Shilpa"
'screen size
Me.Height = Screen.Height - Screen.Height * 5 / 100
Me.Width = Screen.Width
Me.Left = 0
Me.Top = 0
'label5 title bar
Label5.Width = Screen.Width
Label5.Height = 450
Call setmnu
Call setsubmnu
Call setcntrlbx
End Sub
Private Sub setmnu()
mnucntr.Picture = LoadPicture(App.Path & "\appdata\images\mnu.jpeg")
mnucntr.Height = Screen.Height
mnucntr.Stretch = True
mnucntr.Left = 0
mnucntr.Top = 0
mnucntr.Width = (Screen.Width * (20 / 100))
Label1.BackStyle = 0
Label2.BackStyle = 0
Label3.BackStyle = 0
Label4.BackStyle = 0
Label8.BackStyle = 0
Label1.Width = mnucntr.Width
Label2.Width = mnucntr.Width
Label3.Width = mnucntr.Width
Label4.Width = mnucntr.Width
Label8.Width = mnucntr.Width
Label1.Top = 4675
Label2.Top = 4675 + Label1.Height
Label3.Top = 4675 + Label1.Height + Label2.Height
Label4.Top = 4675 + Label1.Height + Label2.Height + Label3.Height
Label8.Top = 4675 + Label1.Height + Label2.Height + Label3.Height + Label4.Height
Label1.BackColor = &H8000000F
Label2.BackColor = &H8000000F
Label3.BackColor = &H8000000F
Label4.BackColor = &H8000000F
Label8.BackColor = &H8000000F
End Sub
Private Sub setsubmnu()
mmnu.Top = 4675
mmnu.Left = Label1.Width - (Label1.Width * 10 / 100)
rmnu.Top = 4275 + Label1.Height
rmnu.Left = Label1.Width - (Label1.Width * 10 / 100)
tmnu.Top = 4675 + Label1.Height + Label2.Height
tmnu.Left = Label1.Width - (Label1.Width * 10 / 100)
umnu.Top = 4675 + Label1.Height + Label2.Height + Label3.Height
umnu.Left = Label1.Width - (Label1.Width * 10 / 100)
ramnu.Top = 4675 + Label1.Height + Label2.Height + Label3.Height + Label4.Height
ramnu.Left = Label1.Width - (Label1.Width * 10 / 100)
End Sub
Private Sub setcntrlbx()
Label6.Left = Screen.Width - 550
Label6.Top = 0
Label7.Left = Screen.Width - 550 - 545
Label7.Top = 0
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call mnudsgn
Label6.BackColor = &HFF&
Label7.BackColor = &HC0C0C0
Call sbmnudsgn
End Sub
Private Sub mnudsgn()
Label1.BackStyle = 0
Label2.BackStyle = 0
Label3.BackStyle = 0
Label4.BackStyle = 0
Label8.BackStyle = 0
End Sub
Private Sub sbmnudsgn()
'changing backstyle of submenus when cursor move to form
marm.BackStyle = 0
macm.BackStyle = 0
mim.BackStyle = 0
rtr.BackStyle = 0
rtrtir.BackStyle = 0
rir.BackStyle = 0
rvr.BackStyle = 0
rtds.BackStyle = 0
rpwiwsr.BackStyle = 0
rvts.BackStyle = 0
ral.BackStyle = 0
rss.BackStyle = 0
rvl.BackStyle = 0
rebao.BackStyle = 0
tdte.BackStyle = 0
ttig.BackStyle = 0
tve.BackStyle = 0
tpe.BackStyle = 0
tce.BackStyle = 0
ucp.BackStyle = 0
udb.BackStyle = 0
udr.BackStyle = 0
uad.BackStyle = 0
rabs.BackStyle = 0
rais.BackStyle = 0
'frames of sub menus
umnu.Visible = False
rmnu.Visible = False
tmnu.Visible = False
mmnu.Visible = False
ramnu.Visible = False
End Sub
Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.BackStyle = 1
Label2.BackStyle = 0
rtr.BackStyle = 0
rmnu.Visible = False
macm.BackStyle = 0
mmnu.Visible = True
End Sub
Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.BackStyle = 1
Label1.BackStyle = 0
Label3.BackStyle = 0
mim.BackStyle = 0
marm.BackStyle = 0
rtr.BackStyle = 0
rir.BackStyle = 0
rtrtir.BackStyle = 0
tdte.BackStyle = 0
tmnu.Visible = False
mmnu.Visible = False
ramnu.Visible = False
rmnu.Visible = True
End Sub
Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.BackStyle = 1
Label2.BackStyle = 0
Label4.BackStyle = 0
Label8.BackStyle = 0
mim.BackStyle = 0
rir.BackStyle = 0
rvr.BackStyle = 0
rtds.BackStyle = 0
tdte.BackStyle = 0
ttig.BackStyle = 0
ucp.BackStyle = 0
umnu.Visible = False
rmnu.Visible = False
mmnu.Visible = False
ramnu.Visible = False
tmnu.Visible = True
End Sub
Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.BackStyle = 1
Label3.BackStyle = 0
Label8.BackStyle = 0
rtds.BackStyle = 0
rpwiwsr.BackStyle = 0
ttig.BackStyle = 0
tve.BackStyle = 0
tpe.BackStyle = 0
ucp.BackStyle = 0
udb.BackStyle = 0
rabs.BackStyle = 0
rais.BackStyle = 0
rmnu.Visible = False
tmnu.Visible = False
mmnu.Visible = False
ramnu.Visible = False
umnu.Visible = True
End Sub
Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label6.BackColor = &HFF&
Label7.BackColor = &HC0C0C0
End Sub
Private Sub Label6_Click()
End
End Sub
Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label6.BackColor = &H8080FF
Label7.BackColor = &HC0C0C0
End Sub
Private Sub Label7_Click()
MDIForm1.WindowState = vbMinimized
End Sub
Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label7.BackColor = &HE0E0E0
Label6.BackColor = &HFF&
End Sub
Private Sub Label8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label8.BackStyle = 1
Label4.BackStyle = 0
Label3.BackStyle = 0
udb.BackStyle = 0
udr.BackStyle = 0
uad.BackStyle = 0
rmnu.Visible = False
tmnu.Visible = False
mmnu.Visible = False
ramnu.Visible = True
umnu.Visible = False
End Sub
Private Sub macm_Click()
acmaster.Show
Label1.BackStyle = 0
macm.BackStyle = 0
mmnu.Visible = False
Me.Enabled = False
End Sub
Private Sub macm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
marm.BackStyle = 0
macm.BackStyle = 1
End Sub
Private Sub marm_Click()
armaster.Show
Label1.BackStyle = 0
marm.BackStyle = 0
mmnu.Visible = False
Me.Enabled = False
End Sub
Private Sub marm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
macm.BackStyle = 0
mim.BackStyle = 0
marm.BackStyle = 1
End Sub
Private Sub mim_Click()
itmaster.Show
Label1.BackStyle = 0
mim.BackStyle = 0
mmnu.Visible = False
Me.Enabled = False
End Sub
Private Sub mim_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
marm.BackStyle = 0
mim.BackStyle = 1
End Sub
Private Sub mnucntr_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.BackStyle = 0
Label3.BackStyle = 0
Label2.BackStyle = 0
Label1.BackStyle = 0
rpwiwsr.BackStyle = 0
rvts.BackStyle = 0
ral.BackStyle = 0
rss.BackStyle = 0
rvl.BackStyle = 0
rebao.BackStyle = 0
tve.BackStyle = 0
tpe.BackStyle = 0
tce.BackStyle = 0
udb.BackStyle = 0
udr.BackStyle = 0
uad.BackStyle = 0
rais.BackStyle = 0
rabs.BackStyle = 0
Label8.BackStyle = 0
umnu.Visible = False
rmnu.Visible = False
tmnu.Visible = False
mmnu.Visible = False
ramnu.Visible = False
macm.BackStyle = 0
End Sub
Private Sub rabs_Click()
blnc_sheet.Show
Label8.BackStyle = 0
ramnu.Visible = False
Me.Enabled = False
End Sub
Private Sub rabs_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
rais.BackStyle = 0
rabs.BackStyle = 1
End Sub
Private Sub rais_Click()
income.Show
Label8.BackStyle = 0
ramnu.Visible = False
Me.Enabled = False
End Sub
Private Sub rais_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
rabs.BackStyle = 0
rais.BackStyle = 1
End Sub
Private Sub ral_Click()
acldgr.Show
Label2.BackStyle = 0
ral.BackStyle = 0
rmnu.Visible = False
Me.Enabled = False
End Sub
Private Sub rebao_Click()
ebao.Show
Label2.BackStyle = 0
rebao.BackStyle = 0
rmnu.Visible = False
Me.Enabled = False
End Sub
Private Sub rir_Click()
inrgstr.Show
Label2.BackStyle = 0
rir.BackStyle = 0
rmnu.Visible = False
Me.Enabled = False
End Sub
Private Sub rpwiwsr_Click()
pwiwsrprt.Show
Label2.BackStyle = 0
rpwiwsr.BackStyle = 0
rmnu.Visible = False
Me.Enabled = False
End Sub
Private Sub rss_Click()
stsum.Show
Label2.BackStyle = 0
rss.BackStyle = 0
rmnu.Visible = False
Me.Enabled = False
End Sub
Private Sub rtds_Click()
tdsum.Show
Label2.BackStyle = 0
rtds.BackStyle = 0
rmnu.Visible = False
Me.Enabled = False
End Sub
Private Sub rtr_Click()
trrecord.Show
Label2.BackStyle = 0
rtr.BackStyle = 0
rmnu.Visible = False
Me.Enabled = False
End Sub
Private Sub rtrtir_Click()
trtig.Show
Label2.BackStyle = 0
rtrtir.BackStyle = 0
rmnu.Visible = False
Me.Enabled = False
End Sub
Private Sub rtrtir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
rtrtir.BackStyle = 1
rtr.BackStyle = 0
rir.BackStyle = 0
End Sub
Private Sub rtr_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
rtr.BackStyle = 1
rtrtir.BackStyle = 0
End Sub
Private Sub rir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
rir.BackStyle = 1
rtrtir.BackStyle = 0
rvr.BackStyle = 0
End Sub
Private Sub rvl_Click()
vhldgr.Show
Label2.BackStyle = 0
rvl.BackStyle = 0
rmnu.Visible = False
Me.Enabled = False
End Sub
Private Sub rvr_Click()
vrgstr.Show
Label2.BackStyle = 0
rvr.BackStyle = 0
rmnu.Visible = False
Me.Enabled = False
End Sub
Private Sub rvr_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
rvr.BackStyle = 1
rir.BackStyle = 0
rtds.BackStyle = 0
End Sub
Private Sub rtds_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
rtds.BackStyle = 1
rvr.BackStyle = 0
rpwiwsr.BackStyle = 0
End Sub
Private Sub rpwiwsr_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
rpwiwsr.BackStyle = 1
rtds.BackStyle = 0
rvts.BackStyle = 0
End Sub
Private Sub rvts_Click()
vtsum.Show
Label2.BackStyle = 0
rvts.BackStyle = 0
rmnu.Visible = False
Me.Enabled = False
End Sub
Private Sub rvts_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
rvts.BackStyle = 1
rpwiwsr.BackStyle = 0
ral.BackStyle = 0
End Sub
Private Sub ral_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ral.BackStyle = 1
rvts.BackStyle = 0
rss.BackStyle = 0
End Sub
Private Sub rss_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
rss.BackStyle = 1
ral.BackStyle = 0
rvl.BackStyle = 0
End Sub
Private Sub rvl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
rvl.BackStyle = 1
rss.BackStyle = 0
rebao.BackStyle = 0
End Sub
Private Sub rebao_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
rebao.BackStyle = 1
rvl.BackStyle = 0
End Sub
Private Sub tce_Click()
centry.Show
Label3.BackStyle = 0
tce.BackStyle = 0
tmnu.Visible = False
Me.Enabled = False
End Sub
Private Sub tdte_Click()
dtentry.Show
Label3.BackStyle = 0
tdte.BackStyle = 0
tmnu.Visible = False
Me.Enabled = False
End Sub
Private Sub tdte_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
tdte.BackStyle = 1
ttig.BackStyle = 0
End Sub
Private Sub tpe_Click()
pentry.Show
Label3.BackStyle = 0
tpe.BackStyle = 0
tmnu.Visible = False
Me.Enabled = False
End Sub
Private Sub ttig_Click()
tigen.Show
Label3.BackStyle = 0
ttig.BackStyle = 0
tmnu.Visible = False
Me.Enabled = False
End Sub
Private Sub ttig_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ttig.BackStyle = 1
tdte.BackStyle = 0
tve.BackStyle = 0
End Sub
Private Sub tve_Click()
ventry.Show
Label3.BackStyle = 0
tve.BackStyle = 0
tmnu.Visible = False
Me.Enabled = False
End Sub
Private Sub tve_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
tve.BackStyle = 1
ttig.BackStyle = 0
tpe.BackStyle = 0
End Sub
Private Sub tpe_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
tpe.BackStyle = 1
tve.BackStyle = 0
tce.BackStyle = 0
End Sub
Private Sub tce_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
tce.BackStyle = 1
tpe.BackStyle = 0
End Sub

Private Sub uad_Click()
help.Show
Label4.BackStyle = 0
uad.BackStyle = 0
umnu.Visible = False
Me.Enabled = False
End Sub
Private Sub ucp_Click()
chpass.Show
Label4.BackStyle = 0
ucp.BackStyle = 0
umnu.Visible = False
Me.Enabled = False
End Sub
Private Sub ucp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ucp.BackStyle = 1
udb.BackStyle = 0
End Sub

Private Sub udb_Click()
DBBR.Show
Label4.BackStyle = 0
udb.BackStyle = 0
umnu.Visible = False
Me.Enabled = False
End Sub

Private Sub udb_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
udb.BackStyle = 1
ucp.BackStyle = 0
udr.BackStyle = 0
End Sub
Private Sub udr_Click()
ABOUT.Show
Label4.BackStyle = 0
udr.BackStyle = 0
umnu.Visible = False
Me.Enabled = False
End Sub
Private Sub udr_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
udr.BackStyle = 1
udb.BackStyle = 0
uad.BackStyle = 0
End Sub
Private Sub uad_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
uad.BackStyle = 1
udr.BackStyle = 0
End Sub
