VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form income 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   10605
   ClientLeft      =   75
   ClientTop       =   75
   ClientWidth     =   17415
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10605
   ScaleWidth      =   17415
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5895
      Left            =   480
      TabIndex        =   40
      Top             =   6360
      Width           =   16695
      Begin MSChart20Lib.MSChart ch 
         Height          =   5655
         Left            =   -240
         OleObjectBlob   =   "income_statement.frx":0000
         TabIndex        =   41
         Top             =   -240
         Width           =   17175
      End
   End
   Begin VB.CommandButton Command1 
      Height          =   550
      Left            =   0
      Picture         =   "income_statement.frx":21DC
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   480
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3135
      Left            =   16080
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   9495
      Begin VB.CommandButton Command2 
         Height          =   575
         Left            =   7800
         Picture         =   "income_statement.frx":56A3
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   2280
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   6720
         TabIndex        =   20
         Top             =   840
         Width           =   2175
      End
      Begin VB.OptionButton Option6 
         BackColor       =   &H80000003&
         Caption         =   "Option1"
         Height          =   195
         Left            =   5160
         TabIndex        =   17
         Top             =   2760
         Width           =   175
      End
      Begin VB.OptionButton Option7 
         BackColor       =   &H80000003&
         Caption         =   "Option1"
         Height          =   195
         Left            =   2640
         TabIndex        =   15
         Top             =   2760
         Width           =   175
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H80000003&
         Caption         =   "Option1"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   2760
         Width           =   175
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H80000003&
         Caption         =   "Option1"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   2160
         Width           =   175
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H80000003&
         Caption         =   "Option1"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   1560
         Width           =   175
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H80000003&
         Caption         =   "00"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   175
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H80000003&
         Caption         =   "Option1"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   175
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   375
         Left            =   2760
         TabIndex        =   7
         Top             =   1440
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
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
         CalendarBackColor=   14737632
         CalendarTitleBackColor=   14737632
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   123076611
         CurrentDate     =   43625
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   1440
         TabIndex        =   11
         Top             =   2040
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
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
         CalendarBackColor=   14737632
         CalendarTitleBackColor=   14737632
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   58523651
         CurrentDate     =   43625
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   4200
         TabIndex        =   12
         Top             =   2040
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
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
         CalendarBackColor=   14737632
         CalendarTitleBackColor=   14737632
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   122355715
         CurrentDate     =   43625
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Vehicle Number"
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
         Left            =   4200
         TabIndex        =   19
         Top             =   840
         Width           =   2415
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "This Year"
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
         Left            =   5640
         TabIndex        =   18
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "This Month"
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
         Left            =   3000
         TabIndex        =   16
         Top             =   2640
         Width           =   1335
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "This week"
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
         Left            =   480
         TabIndex        =   14
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "To"
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
         Left            =   3720
         TabIndex        =   10
         Top             =   2040
         Width           =   375
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   " From "
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
         Left            =   480
         TabIndex        =   9
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "On Specific Date"
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
         Left            =   480
         TabIndex        =   6
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Select a particular vehicle"
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
         Left            =   480
         TabIndex        =   3
         Top             =   840
         Width           =   3015
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "All Vehicles"
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
         Left            =   480
         TabIndex        =   1
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Label Label26 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
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
      Left            =   3120
      TabIndex        =   43
      Top             =   4680
      Width           =   2175
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "Retained Earnings :"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   42
      Top             =   4680
      Width           =   2655
   End
   Begin VB.Label Label24 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "231323215413"
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
      Left            =   2520
      TabIndex        =   39
      Top             =   2400
      Width           =   2175
   End
   Begin VB.Label Label23 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "212313"
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
      Left            =   8400
      TabIndex        =   38
      Top             =   3840
      Width           =   2775
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "1412354137"
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
      Left            =   8400
      TabIndex        =   37
      Top             =   3360
      Width           =   2775
   End
   Begin VB.Label Label21 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "312231231"
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
      Left            =   8400
      TabIndex        =   36
      Top             =   2880
      Width           =   2775
   End
   Begin VB.Label Label20 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "1232313232"
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
      Left            =   8520
      TabIndex        =   35
      Top             =   2400
      Width           =   2775
   End
   Begin VB.Line Line4 
      X1              =   120
      X2              =   11520
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line3 
      X1              =   8040
      X2              =   8040
      Y1              =   1320
      Y2              =   3840
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Loading Charges"
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
      Left            =   5160
      TabIndex        =   34
      Top             =   3840
      Width           =   1815
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Diesel Amount"
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
      Left            =   5160
      TabIndex        =   33
      Top             =   3360
      Width           =   2415
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Payment"
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
      Left            =   5160
      TabIndex        =   32
      Top             =   2880
      Width           =   2175
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Maintainance Charges"
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
      Left            =   5160
      TabIndex        =   31
      Top             =   2400
      Width           =   2415
   End
   Begin VB.Line Line2 
      X1              =   4920
      X2              =   4920
      Y1              =   1800
      Y2              =   4320
   End
   Begin VB.Line Line1 
      X1              =   2160
      X2              =   2160
      Y1              =   1680
      Y2              =   4320
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Sale rate"
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
      Left            =   360
      TabIndex        =   30
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "(Rs.)"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   29
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "(Rs.)"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9000
      TabIndex        =   28
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Expenses"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      TabIndex        =   27
      Top             =   1800
      Width           =   2655
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Income"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   26
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      Height          =   2655
      Left            =   120
      Top             =   1680
      Width           =   11295
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
      TabIndex        =   25
      Top             =   0
      Width           =   495
   End
   Begin VB.Label title 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Income Statement"
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
      TabIndex        =   24
      Top             =   0
      Width           =   12375
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Report :"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   22
      Top             =   840
      Width           =   975
   End
End
Attribute VB_Name = "income"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Frame1.Visible = True
Command1.Visible = False
Call setlbl
End Sub
Private Sub Command2_Click()
Frame1.Visible = False
Command1.Visible = True
Call setlabel
End Sub
Private Sub cl_Click()
menu.Enabled = True
Unload Me
End Sub
Private Sub Form_Load()
 'form apearance
Me.Height = Screen.Height - Screen.Height * 5 / 100 - 430
Me.Width = Screen.Width * 80 / 100
Me.Top = 430
Me.Left = Screen.Width * 20 / 100
Me.Picture = LoadPicture(App.Path & "\appdata\images\back.jpg")
Call setfmnu
Call setlabel
Call setval
End Sub
Private Sub setval()

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
Command1.Top = title.Height + 50
Command1.Left = 70
Label11.Left = 70
Label11.Top = Command1.Height + Command1.Top + 100
'''''
Shape1.Left = 70
Shape1.Top = Label11.Top + Label11.Height + 100
Label9.Top = Shape1.Top
Label9.Left = Shape1.Left + 100
Line4.Y1 = Label9.Height + Label9.Top - 100
Line4.Y2 = Line4.Y1
Line4.X1 = Shape1.Left
Line4.X2 = Shape1.Width + 70
Line1.X1 = Label9.Left + Label9.Width + 300
Line1.X2 = Line1.X1
Line1.Y1 = Shape1.Top
Line1.Y2 = Shape1.Height + Shape1.Top
Label13.Top = Label9.Top
Label13.Left = Line1.X1 + 150
Line2.X1 = Label13.Left + Label13.Width + 500
Line2.X2 = Line2.X1
Line2.Y1 = Shape1.Top
Line2.Y2 = Shape1.Height + Shape1.Top
Label12.Top = Label13.Top
Label12.Left = Line2.X1 + 150
Line3.X1 = Label12.Left + Label12.Width + 300
Line3.X2 = Line3.X1
Line3.Y1 = Shape1.Top
Line3.Y2 = Shape1.Height + Shape1.Top
Label14.Top = Label12.Top
Label14.Left = Line3.X1 + 150
Label15.Top = Line4.Y2 + 200
Label15.Left = Label9.Left
Label24.Top = Label15.Top
Label24.Left = Line2.X2 - 100 - Label24.Width
Label16.Top = Label15.Top
Label16.Left = Label12.Left
Label17.Top = Label16.Top + 100 + Label16.Height
Label17.Left = Label12.Left
Label18.Top = Label17.Top + 100 + Label17.Height
Label18.Left = Label12.Left
Label19.Top = Label18.Top + 100 + Label18.Height
Label19.Left = Label12.Left
Label20.Top = Label15.Top
Label20.Left = Label14.Left
Label21.Top = Label17.Top
Label21.Left = Label14.Left
Label22.Top = Label18.Top
Label22.Left = Label14.Left
Label23.Top = Label19.Top
Label23.Left = Label14.Left
Label25.Top = Label23.Top + Label23.Height + 250
Label25.Left = 70
Label26.Top = Label25.Top
Label26.Left = Label25.Left + Label25.Width + 100
Frame2.Height = Me.Height - (Label25.Top + Label25.Height + 1000)
Frame2.Width = Me.Width - 1000
Frame2.Left = 500
Frame2.Top = Label25.Top + Label25.Height + 500
ch.Height = Frame2.Height + 200
ch.Width = Frame2.Width + 200
ch.Top = -100
ch.Left = -100
End Sub
Private Sub setlbl()
Frame1.Visible = True
Frame1.Top = Command1.Top
Frame1.Left = Command1.Left
Command1.Visible = False
Label11.Top = Frame1.Top + Frame1.Height + 100
'''''
Shape1.Left = 70
Shape1.Top = Label11.Top + Label11.Height + 100
Label9.Top = Shape1.Top
Label9.Left = Shape1.Left + 100
Line4.Y1 = Label9.Height + Label9.Top - 100
Line4.Y2 = Line4.Y1
Line4.X1 = Shape1.Left
Line4.X2 = Shape1.Width + 70
Line1.X1 = Label9.Left + Label9.Width + 300
Line1.X2 = Line1.X1
Line1.Y1 = Shape1.Top
Line1.Y2 = Shape1.Height + Shape1.Top
Label13.Top = Label9.Top
Label13.Left = Line1.X1 + 150
Line2.X1 = Label13.Left + Label13.Width + 500
Line2.X2 = Line2.X1
Line2.Y1 = Shape1.Top
Line2.Y2 = Shape1.Height + Shape1.Top
Label12.Top = Label13.Top
Label12.Left = Line2.X1 + 150
Line3.X1 = Label12.Left + Label12.Width + 300
Line3.X2 = Line3.X1
Line3.Y1 = Shape1.Top
Line3.Y2 = Shape1.Height + Shape1.Top
Label14.Top = Label12.Top
Label14.Left = Line3.X1 + 150
Label15.Top = Line4.Y2 + 200
Label15.Left = Label9.Left
Label24.Top = Label15.Top
Label24.Left = Line2.X2 - 100 - Label24.Width
Label16.Top = Label15.Top
Label16.Left = Label12.Left
Label17.Top = Label16.Top + 100 + Label16.Height
Label17.Left = Label12.Left
Label18.Top = Label17.Top + 100 + Label17.Height
Label18.Left = Label12.Left
Label19.Top = Label18.Top + 100 + Label18.Height
Label19.Left = Label12.Left
Label20.Top = Label15.Top
Label20.Left = Label14.Left
Label21.Top = Label17.Top
Label21.Left = Label14.Left
Label22.Top = Label18.Top
Label22.Left = Label14.Left
Label23.Top = Label19.Top
Label23.Left = Label14.Left
Label25.Top = Label23.Top + Label23.Height + 250
Label25.Left = 70
Label26.Top = Label25.Top
Label26.Left = Label25.Left + Label25.Width + 100
Frame2.Height = Me.Height - (Label25.Top + Label25.Height + 1000)
Frame2.Width = Me.Width - 1000
Frame2.Left = 500
Frame2.Top = Label25.Top + Label25.Height + 500
ch.Height = Frame2.Height + 200
ch.Width = Frame2.Width + 200
ch.Top = -100
ch.Left = -100
End Sub

