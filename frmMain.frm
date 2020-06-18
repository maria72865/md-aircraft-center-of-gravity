VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Weight & Balance and Fuel Calculations"
   ClientHeight    =   7875
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   11850
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7875
   ScaleWidth      =   11850
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtSeats 
      Height          =   390
      Left            =   120
      TabIndex        =   80
      Top             =   7200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   7320
      TabIndex        =   63
      Top             =   4200
      Width           =   4215
      Begin VB.OptionButton Option8 
         Caption         =   "500"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   67
         Top             =   45
         Width           =   855
      End
      Begin VB.OptionButton Option7 
         Caption         =   "1000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1080
         TabIndex        =   66
         Top             =   45
         Width           =   855
      End
      Begin VB.OptionButton Option6 
         Caption         =   "1500"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2160
         TabIndex        =   65
         Top             =   45
         Width           =   855
      End
      Begin VB.OptionButton Option5 
         Caption         =   "2000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3240
         TabIndex        =   64
         Top             =   45
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   1680
      TabIndex        =   58
      Top             =   4200
      Width           =   4215
      Begin VB.OptionButton Option1 
         Caption         =   "1000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   62
         Top             =   45
         Width           =   855
      End
      Begin VB.OptionButton Option2 
         Caption         =   "1500"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1200
         TabIndex        =   61
         Top             =   45
         Width           =   855
      End
      Begin VB.OptionButton Option3 
         Caption         =   "2000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2280
         TabIndex        =   60
         Top             =   45
         Width           =   855
      End
      Begin VB.OptionButton Option4 
         Caption         =   "2224"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3360
         TabIndex        =   59
         Top             =   45
         Width           =   855
      End
   End
   Begin VB.CommandButton cmd2 
      Caption         =   "Calc"
      Height          =   375
      Left            =   7320
      TabIndex        =   11
      Top             =   3720
      Width           =   975
   End
   Begin VB.CommandButton cmd1 
      Caption         =   "Calc"
      Height          =   375
      Left            =   1680
      TabIndex        =   9
      Top             =   3720
      Width           =   975
   End
   Begin VB.TextBox txtTOWt 
      Height          =   390
      Left            =   6480
      TabIndex        =   51
      Top             =   7200
      Width           =   1575
   End
   Begin VB.TextBox txtFuel 
      Height          =   390
      Left            =   8520
      TabIndex        =   12
      Text            =   "Enter in Pounds"
      Top             =   3720
      Width           =   3015
   End
   Begin VB.TextBox txtSIC 
      Height          =   390
      Left            =   7320
      TabIndex        =   5
      Top             =   2520
      Width           =   4215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   10200
      TabIndex        =   20
      Top             =   7200
      Width           =   1215
   End
   Begin VB.TextBox txtPIC 
      Height          =   390
      Left            =   1680
      TabIndex        =   4
      Top             =   2520
      Width           =   4215
   End
   Begin VB.TextBox txtSN 
      Height          =   390
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1920
      Width           =   1455
   End
   Begin VB.TextBox txtDep 
      Height          =   405
      Left            =   7440
      TabIndex        =   2
      Top             =   1920
      Width           =   1455
   End
   Begin VB.TextBox txtFltNo 
      Height          =   405
      Left            =   10440
      TabIndex        =   3
      Text            =   "PI "
      Top             =   1920
      Width           =   1095
   End
   Begin VB.TextBox txtTotal 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   405
      Left            =   9360
      TabIndex        =   32
      Top             =   6540
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "CARGO in Kilograms"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2535
      Left            =   240
      TabIndex        =   25
      Top             =   4560
      Width           =   11295
      Begin VB.TextBox txtSeat11 
         ForeColor       =   &H00FF0000&
         Height          =   390
         Left            =   7170
         TabIndex        =   82
         Top             =   1320
         Visible         =   0   'False
         Width           =   1075
      End
      Begin VB.TextBox txtZoneD 
         ForeColor       =   &H000000FF&
         Height          =   390
         Left            =   4365
         TabIndex        =   79
         Text            =   "0.00"
         Top             =   2040
         Visible         =   0   'False
         Width           =   1075
      End
      Begin VB.TextBox txtZoneC 
         ForeColor       =   &H000000FF&
         Height          =   390
         Left            =   1485
         TabIndex        =   78
         Text            =   "0.00"
         Top             =   2040
         Visible         =   0   'False
         Width           =   1075
      End
      Begin VB.TextBox txtZoneB 
         ForeColor       =   &H000000FF&
         Height          =   390
         Left            =   4365
         TabIndex        =   77
         Text            =   "0.00"
         Top             =   1560
         Visible         =   0   'False
         Width           =   1075
      End
      Begin VB.TextBox txtZoneA 
         ForeColor       =   &H000000FF&
         Height          =   390
         Left            =   1485
         TabIndex        =   72
         Text            =   "0.00"
         Top             =   1560
         Visible         =   0   'False
         Width           =   1075
      End
      Begin VB.CommandButton cmdKilo 
         BackColor       =   &H000000FF&
         Caption         =   "Kilograms"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   425
         Left            =   9120
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   1080
         Width           =   1560
      End
      Begin VB.CommandButton cmdLbs 
         BackColor       =   &H00FF0000&
         Caption         =   "Pounds"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   425
         Left            =   9120
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   480
         Width           =   1560
      End
      Begin VB.TextBox txtZone1 
         ForeColor       =   &H000000FF&
         Height          =   390
         Left            =   1485
         TabIndex        =   13
         Text            =   "0.00"
         Top             =   360
         Width           =   1075
      End
      Begin VB.TextBox txtZone2 
         ForeColor       =   &H000000FF&
         Height          =   390
         Left            =   4365
         TabIndex        =   14
         Text            =   "0.00"
         Top             =   360
         Width           =   1075
      End
      Begin VB.TextBox txtZone3 
         ForeColor       =   &H000000FF&
         Height          =   390
         Left            =   7170
         TabIndex        =   15
         Text            =   "0.00"
         Top             =   360
         Width           =   1075
      End
      Begin VB.TextBox txtZone4 
         ForeColor       =   &H000000FF&
         Height          =   390
         Left            =   1485
         TabIndex        =   16
         Text            =   "0.00"
         Top             =   840
         Width           =   1075
      End
      Begin VB.TextBox txtZone5 
         ForeColor       =   &H000000FF&
         Height          =   390
         Left            =   4365
         TabIndex        =   17
         Text            =   "0.00"
         Top             =   840
         Width           =   1075
      End
      Begin VB.TextBox txtZone6 
         ForeColor       =   &H000000FF&
         Height          =   390
         Left            =   7170
         TabIndex        =   18
         Text            =   "0.00"
         Top             =   840
         Width           =   1075
      End
      Begin VB.Label Label40 
         Caption         =   "kgs."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   8280
         TabIndex        =   83
         Top             =   1440
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label39 
         Alignment       =   1  'Right Justify
         Caption         =   "Seat 11"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   6100
         TabIndex        =   81
         Top             =   1485
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label38 
         Caption         =   "kgs."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   5520
         TabIndex        =   76
         Top             =   2210
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label37 
         Caption         =   "kgs."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2595
         TabIndex        =   75
         Top             =   2210
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label36 
         Caption         =   "kgs."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   5520
         TabIndex        =   74
         Top             =   1730
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label35 
         Caption         =   "kgs."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2595
         TabIndex        =   73
         Top             =   1730
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label34 
         Alignment       =   1  'Right Justify
         Caption         =   "Pod D"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   3240
         TabIndex        =   71
         Top             =   2160
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label33 
         Alignment       =   1  'Right Justify
         Caption         =   "Pod C"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   550
         TabIndex        =   70
         Top             =   2160
         Visible         =   0   'False
         Width           =   870
      End
      Begin VB.Label Label32 
         Alignment       =   1  'Right Justify
         Caption         =   "Pod B"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   3240
         TabIndex        =   69
         Top             =   1680
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label31 
         Alignment       =   1  'Right Justify
         Caption         =   "Pod A"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   550
         TabIndex        =   68
         Top             =   1680
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         Caption         =   "Total Cargo in Kgs (not to exceed 1542):"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   510
         Left            =   5880
         TabIndex        =   49
         Top             =   1920
         Width           =   3135
      End
      Begin VB.Label Label22 
         Caption         =   "kgs."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Left            =   8280
         TabIndex        =   47
         Top             =   1010
         Width           =   375
      End
      Begin VB.Label Label21 
         Caption         =   "kgs."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Left            =   8280
         TabIndex        =   46
         Top             =   530
         Width           =   375
      End
      Begin VB.Label Label20 
         Caption         =   "kgs."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Left            =   5490
         TabIndex        =   45
         Top             =   1010
         Width           =   375
      End
      Begin VB.Label Label19 
         Caption         =   "kgs."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Left            =   5490
         TabIndex        =   44
         Top             =   530
         Width           =   375
      End
      Begin VB.Label Label18 
         Caption         =   "kgs."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Left            =   2595
         TabIndex        =   43
         Top             =   1010
         Width           =   375
      End
      Begin VB.Label Label17 
         Caption         =   "kgs."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Left            =   2595
         TabIndex        =   42
         Top             =   530
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Hold 1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   95
         TabIndex        =   31
         Top             =   480
         Width           =   1320
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Hold 2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   3010
         TabIndex        =   30
         Top             =   480
         Width           =   1320
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Hold 3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   5805
         TabIndex        =   29
         Top             =   480
         Width           =   1290
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Hold 4"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   95
         TabIndex        =   28
         Top             =   960
         Width           =   1320
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Hold 5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   3000
         TabIndex        =   27
         Top             =   960
         Width           =   1320
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Hold 6"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   5805
         TabIndex        =   26
         Top             =   960
         Width           =   1290
      End
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next"
      Height          =   375
      Left            =   8640
      TabIndex        =   19
      Top             =   7200
      Width           =   1215
   End
   Begin VB.TextBox txtFOB 
      Height          =   390
      Left            =   2880
      TabIndex        =   10
      Text            =   "Enter in Pounds"
      Top             =   3720
      Width           =   3015
   End
   Begin VB.ComboBox cmbAC 
      Height          =   390
      ItemData        =   "frmMain.frx":0000
      Left            =   1680
      List            =   "frmMain.frx":0002
      OLEDragMode     =   1  'Automatic
      TabIndex        =   0
      Top             =   1920
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   390
      Left            =   2280
      TabIndex        =   52
      Text            =   "Text1"
      Top             =   1920
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.ComboBox cmbOrig 
      Height          =   390
      Left            =   1680
      TabIndex        =   6
      Top             =   3120
      Width           =   4215
   End
   Begin VB.TextBox txtOrig 
      Height          =   405
      Left            =   1680
      TabIndex        =   21
      Top             =   3120
      Width           =   4215
   End
   Begin VB.ComboBox cmbDest 
      Height          =   390
      Left            =   7320
      TabIndex        =   7
      Top             =   3120
      Width           =   4215
   End
   Begin VB.TextBox txtDest 
      Height          =   405
      Left            =   7320
      TabIndex        =   8
      Top             =   3120
      Width           =   4215
   End
   Begin VB.Label Label30 
      Alignment       =   2  'Center
      Caption         =   "Anytime"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   57
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label29 
      Alignment       =   2  'Center
      Caption         =   "Anywhere"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   56
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Caption         =   "Any Aircraft"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   55
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label27 
      Alignment       =   2  'Center
      Caption         =   "We Deliver"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   54
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Label26 
      Alignment       =   2  'Center
      Caption         =   "PILOT INTERNATIONAL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   53
      Top             =   120
      Width           =   4935
   End
   Begin VB.Label Label24 
      Alignment       =   2  'Center
      Caption         =   "Take Off Weight (not to exceed 8750 lbs):"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   50
      Top             =   7320
      Width           =   4695
   End
   Begin VB.Label Label23 
      Alignment       =   1  'Right Justify
      Caption         =   "Co-Pilot:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   48
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   1095
      Left            =   3240
      Picture         =   "frmMain.frx":0004
      Stretch         =   -1  'True
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      Caption         =   "Captain:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   41
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label12 
      Caption         =   "SN:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   40
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      Caption         =   "Dep. Date:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6480
      TabIndex        =   37
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      Caption         =   "Orig:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   36
      Top             =   3240
      Width           =   615
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      Caption         =   "Dest:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      TabIndex        =   35
      Top             =   3240
      Width           =   735
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "Flight No:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9120
      TabIndex        =   34
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "CESSNA CARAVAN C208B WEIGHT and BALANCE / LOADPLAN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6360
      TabIndex        =   33
      Top             =   360
      Width           =   5055
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "Expected Fuel Use:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5880
      TabIndex        =   24
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "Fuel on Board:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      TabIndex        =   23
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Aircraft:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   22
      Top             =   2040
      Width           =   975
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
      Begin VB.Menu mnuAboutPI 
         Caption         =   "&Pilot International"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim intTotal2 As Double, intMouseCheck As Integer

Private Sub cmbAC_Click()

    If cmbAC = "208 Cargo Master" Then
            
        txtSN = "208B0936"
        intFOB = Val(txtFOB)
        intBurn = Val(txtFuel)
        intBEW = 4816
        txtTOWt = intBEW
        txtTOWt = Format(txtTOWt.Text, "Fixed")
        intArm1 = 180.23
        intMom1 = 867987.68
        Label31.Visible = True
        Label32.Visible = True
        Label33.Visible = True
        Label34.Visible = True
        Label35.Visible = True
        Label36.Visible = True
        Label37.Visible = True
        Label38.Visible = True
        Label2 = "Hold 1"
        Label4 = "Hold 2"
        Label5 = "Hold 3"
        Label6 = "Hold 4"
        Label7 = "Hold 5"
        Label8 = "Hold 6"
        Label8.Visible = True
        Label22.Visible = True
        txtZone6.Visible = True
        txtZoneA.Visible = True
        txtZoneB.Visible = True
        txtZoneC.Visible = True
        txtZoneD.Visible = True
        txtZoneA.TabIndex = 19
        txtZoneB.TabIndex = 20
        txtZoneC.TabIndex = 21
        txtZoneD.TabIndex = 22
        txtSeats = "208BCM"
        
    ElseIf cmbAC = "208 Grand Caravan" Then
    
        txtSN = "208B0937"
        intFOB = Val(txtFOB)
        intBurn = Val(txtFuel)
        intBEW = 4839
        txtTOWt = intBEW
        txtTOWt = Format(txtTOWt.Text, "Fixed")
        intArm1 = 180.93
        intMom1 = 875520.27
        Label31.Visible = True
        Label32.Visible = True
        Label33.Visible = True
        Label34.Visible = True
        Label35.Visible = True
        Label36.Visible = True
        Label37.Visible = True
        Label38.Visible = True
        txtZoneA.Visible = True
        txtZoneB.Visible = True
        txtZoneC.Visible = True
        txtZoneD.Visible = True
        txtZoneA.TabIndex = 19
        txtZoneB.TabIndex = 20
        txtZoneC.TabIndex = 21
        txtZoneD.TabIndex = 22
        frm208B.Show
        frmMain.Hide
        
    ElseIf cmbAC = "208 Amphib" Then
    
        txtSN = "208B0938"
        intFOB = Val(txtFOB)
        intBurn = Val(txtFuel)
        intBEW = 4839
        txtTOWt = intBEW
        txtTOWt = Format(txtTOWt.Text, "Fixed")
        intArm1 = 180.93
        intMom1 = 875520.27
        Label31.Visible = False
        Label32.Visible = False
        Label33.Visible = False
        Label34.Visible = False
        Label35.Visible = False
        Label36.Visible = False
        Label37.Visible = False
        Label38.Visible = False
        txtZoneA.Visible = False
        txtZoneB.Visible = False
        txtZoneC.Visible = False
        txtZoneD.Visible = False
        frm208A.Show
        frmMain.Hide
        frm208A.Label3 = "CESSNA 208 AMPHIB PASSENGER SEATING ARRANGEMENT"
        
    ElseIf cmbAC = "208 Caravan" Then
    
        txtSN = "208A0939"
        intFOB = Val(txtFOB)
        intBurn = Val(txtFuel)
        intBEW = 4839
        txtTOWt = intBEW
        txtTOWt = Format(txtTOWt.Text, "Fixed")
        intArm1 = 180.93
        intMom1 = 875520.27
        Label31.Visible = False
        Label32.Visible = False
        Label33.Visible = False
        Label34.Visible = False
        Label35.Visible = False
        Label36.Visible = False
        Label37.Visible = False
        Label38.Visible = False
        txtZoneA.Visible = False
        txtZoneB.Visible = False
        txtZoneC.Visible = False
        txtZoneD.Visible = False
        frm208A.Show
        frmMain.Hide
    End If
                                                                                            Text1 = int0
    'If frmMain.Text1 <> "4" Then
    '    Call MsgBox("Nice try Pal!", vbOKOnly, "No Changes Allowed!")
    '    End
    'Else
    '    frmMain.txtFltNo.SetFocus
    'End If
                                            
End Sub

Private Sub cmbDest_Click()
    If cmbDest.Text = "Other" Then
        txtDest.SetFocus
        cmbDest.Visible = False
        txtDest.Text = ""
        Exit Sub
    Else
        txtDest = cmbDest
    End If
End Sub

Private Sub cmbOrig_Click()
    If cmbOrig.Text = "Other" Then
        cmbOrig.Visible = False
        txtOrig.Text = ""
        txtOrig.SetFocus
    Else
        txtOrig = cmbOrig
    End If

End Sub

Private Sub cmd2_Click()
    frmBurnCalc.Show
End Sub

Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdKilo_Click()

    If Frame1.Caption <> "CARGO in Kilograms" Then
    
        Frame1.Caption = "CARGO in Kilograms"
        Frame1.ForeColor = &HFF&
        txtZone1.ForeColor = &HFF&
        txtZone2.ForeColor = &HFF&
        txtZone3.ForeColor = &HFF&
        txtZone4.ForeColor = &HFF&
        txtZone5.ForeColor = &HFF&
        txtZone6.ForeColor = &HFF&
        txtZoneA.ForeColor = &HFF&
        txtZoneB.ForeColor = &HFF&
        txtZoneC.ForeColor = &HFF&
        txtZoneD.ForeColor = &HFF&
        txtSeat11.ForeColor = &HFF&
        Label2.ForeColor = &HFF&
        Label3.ForeColor = &HFF&
        Label4.ForeColor = &HFF&
        Label5.ForeColor = &HFF&
        Label6.ForeColor = &HFF&
        Label7.ForeColor = &HFF&
        Label8.ForeColor = &HFF&
        Label39.ForeColor = &HFF&
        Label40.ForeColor = &HFF&
        Label17.ForeColor = &HFF&
        Label17.Caption = "kgs."
        Label18.ForeColor = &HFF&
        Label18.Caption = "kgs."
        Label19.ForeColor = &HFF&
        Label19.Caption = "kgs."
        Label20.ForeColor = &HFF&
        Label20.Caption = "kgs."
        Label21.ForeColor = &HFF&
        Label21.Caption = "kgs."
        Label22.ForeColor = &HFF&
        Label22.Caption = "kgs."
        Label25.ForeColor = &HFF&
        Label25.Caption = "Total Cargo in Kgs (not to exceed 1542):"
        txtTotal.ForeColor = &HFF&
        Label31.ForeColor = &HFF&
        Label32.ForeColor = &HFF&
        Label33.ForeColor = &HFF&
        Label34.ForeColor = &HFF&
        Label35.ForeColor = &HFF&
        Label35.Caption = "kgs."
        Label36.ForeColor = &HFF&
        Label36.Caption = "kgs."
        Label37.ForeColor = &HFF&
        Label37.Caption = "kgs."
        Label38.ForeColor = &HFF&
        Label38.Caption = "kgs."
        
        If txtZone1 <> "" Then
            txtZone1 = Val(txtZone1) / 2.20458553791887
            txtZone1 = Format(txtZone1.Text, "Fixed")
            int1 = Val(txtZone1)
        End If
        
        If txtZone2 <> "" Then
            txtZone2 = Val(txtZone2) / 2.20458553791887
            txtZone2 = Format(txtZone2.Text, "Fixed")
            int2 = Val(txtZone2)
        End If
        
        If txtZone3 <> "" Then
            txtZone3 = Val(txtZone3) / 2.20458553791887
            txtZone3 = Format(txtZone3.Text, "Fixed")
            int3 = Val(txtZone3)
        End If
        
        If txtZone4 <> "" Then
            txtZone4 = Val(txtZone4) / 2.20458553791887
            txtZone4 = Format(txtZone4.Text, "Fixed")
            int4 = Val(txtZone4)
        End If
        
        If txtZone5 <> "" Then
            txtZone5 = Val(txtZone5) / 2.20458553791887
            txtZone5 = Format(txtZone5.Text, "Fixed")
            int5 = Val(txtZone5)
        End If
        
        If txtZone6 <> "" Then
            txtZone6 = Val(txtZone6) / 2.20458553791887
            txtZone6 = Format(txtZone6.Text, "Fixed")
            int6 = Val(txtZone6)
        End If
        
        If txtZoneA <> "" Then
            txtZoneA = Val(txtZoneA) / 2.20458553791887
            txtZoneA = Format(txtZoneA.Text, "Fixed")
            int7 = Val(txtZoneA)
        End If
        
        If txtZoneB <> "" Then
            txtZoneB = Val(txtZoneB) / 2.20458553791887
            txtZoneB = Format(txtZoneB.Text, "Fixed")
            int8 = Val(txtZoneB)
        End If
        
        If txtZoneC <> "" Then
            txtZoneC = Val(txtZoneC) / 2.20458553791887
            txtZoneC = Format(txtZoneC.Text, "Fixed")
            int9 = Val(txtZoneC)
        End If
        
        If txtZoneD <> "" Then
            txtZoneD = Val(txtZoneD) / 2.20458553791887
            txtZoneD = Format(txtZoneD.Text, "Fixed")
            int10 = Val(txtZoneD)
        End If
        
        If txtSeat11 <> "" Then
            txtSeat11 = Val(txtSeat11) / 2.20458553791887
            txtSeat11 = Format(txtSeat11.Text, "Fixed")
            int11 = Val(txtSeat11)
        End If
    
    End If
    
    'intCargoK = int1 + int2 + int3 + int4 + int5 + int6
    txtTotal = Val(txtZone1) + Val(txtZone2) + Val(txtZone3) + Val(txtZone4) + Val(txtZone5) + Val(txtZone6) + Val(txtZoneA) + Val(txtZoneB) + Val(txtZoneC) + Val(txtZoneD) + Val(txtSeat11)
    'txtTotal = intCargoK
    txtTotal = Format(txtTotal.Text, "Fixed")
    
End Sub

Private Sub cmdLbs_Click()

    If Frame1.Caption <> "CARGO in Pounds" Then
    
        Frame1.Caption = "CARGO in Pounds"
        Frame1.ForeColor = 12582912
        txtZone1.ForeColor = 12582912
        txtZone2.ForeColor = 12582912
        txtZone3.ForeColor = 12582912
        txtZone4.ForeColor = 12582912
        txtZone5.ForeColor = 12582912
        txtZone6.ForeColor = 12582912
        txtZoneA.ForeColor = 12582912
        txtZoneB.ForeColor = 12582912
        txtZoneC.ForeColor = 12582912
        txtZoneD.ForeColor = 12582912
        txtSeat11.ForeColor = 12582912
        Label2.ForeColor = 12582912
        Label3.ForeColor = 12582912
        Label4.ForeColor = 12582912
        Label5.ForeColor = 12582912
        Label6.ForeColor = 12582912
        Label7.ForeColor = 12582912
        Label8.ForeColor = 12582912
        Label39.ForeColor = 12582912
        Label40.ForeColor = 12582912
        Label17.ForeColor = 12582912
        Label17.Caption = "lbs."
        Label18.ForeColor = 12582912
        Label18.Caption = "lbs."
        Label19.ForeColor = 12582912
        Label19.Caption = "lbs."
        Label20.ForeColor = 12582912
        Label20.Caption = "lbs."
        Label21.ForeColor = 12582912
        Label21.Caption = "lbs."
        Label22.ForeColor = 12582912
        Label22.Caption = "lbs."
        Label25.ForeColor = 12582912
        Label25.Caption = "Total Cargo in Lbs (not to exceed 3400):"
        txtTotal.ForeColor = 12582912
        Label31.ForeColor = 12582912
        Label32.ForeColor = 12582912
        Label33.ForeColor = 12582912
        Label34.ForeColor = 12582912
        Label35.ForeColor = 12582912
        Label35.Caption = "lbs."
        Label36.ForeColor = 12582912
        Label36.Caption = "lbs."
        Label37.ForeColor = 12582912
        Label37.Caption = "lbs."
        Label38.ForeColor = 12582912
        Label38.Caption = "lbs."
    
        If txtZone1 <> "" Then
            txtZone1 = Val(txtZone1) * 2.20458553791887
            txtZone1 = Format(txtZone1.Text, "Fixed")
            int1 = Val(txtZone1)
        End If
        
        If txtZone2 <> "" Then
            txtZone2 = Val(txtZone2) * 2.20458553791887
            txtZone2 = Format(txtZone2.Text, "Fixed")
            int2 = Val(txtZone2)
        End If
        
        If txtZone3 <> "" Then
            txtZone3 = Val(txtZone3) * 2.20458553791887
            txtZone3 = Format(txtZone3.Text, "Fixed")
            int3 = Val(txtZone3)
        End If
        
        If txtZone4 <> "" Then
            txtZone4 = Val(txtZone4) * 2.20458553791887
            txtZone4 = Format(txtZone4.Text, "Fixed")
            int4 = Val(txtZone4)
        End If
        
        If txtZone5 <> "" Then
            txtZone5 = Val(txtZone5) * 2.20458553791887
            txtZone5 = Format(txtZone5.Text, "Fixed")
            int5 = Val(txtZone5)
        End If
        
        If txtZone6 <> "" Then
            txtZone6 = Val(txtZone6) * 2.20458553791887
            txtZone6 = Format(txtZone6.Text, "Fixed")
            int6 = Val(txtZone6)
        End If
        
        If txtZoneA <> "" Then
            txtZoneA = Val(txtZoneA) * 2.20458553791887
            txtZoneA = Format(txtZoneA.Text, "Fixed")
            int7 = Val(txtZoneA)
        End If
        
        If txtZoneB <> "" Then
            txtZoneB = Val(txtZoneB) * 2.20458553791887
            txtZoneB = Format(txtZoneB.Text, "Fixed")
            int8 = Val(txtZoneB)
        End If
        
        If txtZoneC <> "" Then
            txtZoneC = Val(txtZoneC) * 2.20458553791887
            txtZoneC = Format(txtZoneC.Text, "Fixed")
            int9 = Val(txtZoneC)
        End If
        
        If txtZoneD <> "" Then
            txtZoneD = Val(txtZoneD) * 2.20458553791887
            txtZoneD = Format(txtZoneD.Text, "Fixed")
            int10 = Val(txtZoneD)
        End If
        
        If txtSeat11 <> "" Then
            txtSeat11 = Val(txtSeat11) * 2.20458553791887
            txtSeat11 = Format(txtSeat11.Text, "Fixed")
            int11 = Val(txtSeat11)
        End If
    
    End If
        
    intCargoP = int1 + int2 + int3 + int4 + int5 + int6 + int7 + int8 + int9 + int10 + int11
    txtTotal = intCargoP
    txtTotal = Format(txtTotal.Text, "Fixed")

End Sub

Private Sub cmdNext_Click()

    If cmbAC = "208 Cargo Master" Then
            
        intFOB = Val(txtFOB)
        intBurn = Val(txtFuel)
        intBEW = 4816
        intArm1 = 180.23
        intMom1 = 867987.68
        
    ElseIf cmbAC = "208 Grand Caravan" Then
        
        intFOB = Val(txtFOB)
        intBurn = Val(txtFuel)
        intBEW = 4839
        intArm1 = 180.93
        intMom1 = 875520.27
        
    ElseIf cmbAC = "208 Amphib" Then
        
        intFOB = Val(txtFOB)
        intBurn = Val(txtFuel)
        intBEW = 4839
        intArm1 = 180.93
        intMom1 = 875520.27
        
    ElseIf cmbAC = "208 Caravan" Then
        
        intFOB = Val(txtFOB)
        intBurn = Val(txtFuel)
        intBEW = 4839
        intArm1 = 180.93
        intMom1 = 875520.27
        
    End If
                
    intCargoP = intP1 + intP2 + intP3 + intP4 + intP5 + intP6 + intP7 + intP8 + intP9 + intP10 + intP11
    intTOWt = intBEW + intFOB + intCargoP + intCrew - 35
    txtTOWt = intTOWt
    txtTOWt = Format(txtTOWt.Text, "Fixed")
                                                                                                                                    'If txtSN <> "208B0936" And txtSN <> "208B0937" And txtSN <> "208B0938" And txtSN <> "208A0939" Then
                                                                                                                                    '    Call MsgBox("Please enter a valid aircraft", vbOKOnly, "Invalid Aircraft!")
                                                                                                                                    '    cmbAC.SetFocus
                                                                                                                                    '    Exit Sub
                                                                                                                                    'End If
    If intCargoP > 3400 Then
        Call Zone1
    Else
        intCargoK = intK1 + intK2 + intK3 + intK4 + intK5 + intK6 + intK7 + intK8 + intK9 + intK10 + intK11
        intRampFuel = intFOB
        intTaxi = 35
        If intTOWt > 8750 And intTOWt < 9090 Then
            Call TOWt
            If intRetVal = 6 Then
                frm2.Show
                Exit Sub
            Else
                txtZone1.SetFocus
                Exit Sub
            End If
        Else
            If intTOWt > 9090 Then
                Call MsgBox("Take Off Weight Limit exceeded, even with zero crew weight.  Please re-enter acceptable values", vbOKOnly, "Take Off Weight Limit Exceeded!")
                txtZone1.SetFocus
                Exit Sub
            End If
        End If
        
    End If
        frm2.Show
        'frmMain.Hide
        
End Sub

Private Sub cmd1_Click()
    'txtFOB = "35"
    frmFuelCalc.Show
End Sub

Private Sub Form_Load()

    cmbAC.AddItem "208 Cargo Master"
    cmbAC.AddItem "208 Grand Caravan"
    cmbAC.AddItem "208 Amphib"
    cmbAC.AddItem "208 Caravan"
    int0 = cmbAC.ListCount
    
    cmbOrig.AddItem "LEBL"
    cmbOrig.AddItem "LEMA"
    cmbOrig.AddItem "LEPA"
    cmbOrig.AddItem "Other"
    
    cmbDest.AddItem "LEBL"
    cmbDest.AddItem "LEMA"
    cmbDest.AddItem "LEPA"
    cmbDest.AddItem "Other"
    
    Call cmdLbs_Click
        
    txtDep.Text = Date
    intCrew = 340
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub mnuAboutPI_Click()
    frmAbout.Show
End Sub

Private Sub Option1_Click()
    txtFOB = "1000"
    txtFOB.SetFocus
End Sub

Private Sub Option2_Click()
    txtFOB = "1500"
    txtFOB.SetFocus
End Sub

Private Sub Option3_Click()
    txtFOB = "2000"
    txtFOB.SetFocus
End Sub

Private Sub Option4_Click()
    txtFOB = "2224"
    txtFOB.SetFocus
End Sub
Private Sub Option8_Click()
    txtFuel = "500"
    txtFuel.SetFocus
End Sub

Private Sub Option7_Click()
    txtFuel = "1000"
    txtFuel.SetFocus
End Sub

Private Sub Option6_Click()
    txtFuel = "1500"
    txtFuel.SetFocus
End Sub

Private Sub Option5_Click()
    txtFuel = "2000"
    txtFuel.SetFocus
End Sub
Private Sub txtDest_DblClick()
    txtDest.Text = ""
    cmbDest.Text = ""
    cmbDest.Visible = True
End Sub

Private Sub txtDest_GotFocus()
    txtDest.TabIndex = 8
End Sub

Private Sub txtFltNo_GotFocus()
    txtFltNo.SelStart = 2
End Sub


Private Sub txtFOB_GotFocus()
    
    If txtFOB = "Enter in Pounds" Then
        txtFOB = ""
    End If
    txtFOB.SelStart = 0
    txtFOB.SelLength = Len(txtFOB.Text)

End Sub

Private Sub txtFOB_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 48 And KeyAscii <= 57 Then
        'OK
    Else
        If KeyAscii = 46 Then
            'Decimal OK
        Else
            If KeyAscii = 8 Then
                'Backspace OK
            Else
                KeyAscii = 0
                Beep
            End If
        End If
        
    End If
End Sub

Private Sub txtFOB_LostFocus()

    If txtFOB = "" Then
        Call MsgBox("Please enter valid Fuel Wieght.", vbOKOnly, "Fuel Weight Limit Exceeded")
        txtFOB.SetFocus
        Exit Sub
    Else
        If Val(txtFOB) > 2224 Then
            txtFOB = ""
            Call MsgBox("Fuel Wieght must not exceed 2224 pounds.  Please enter a value within limits.", vbOKOnly, "Fuel Weight Limit Exceeded")
            txtFOB.SetFocus
        ElseIf Val(txtFOB) < 35 Then
            txtFOB = ""
            Call MsgBox("Fuel Wieght must be at least 35 pounds.  Please enter a value within limits.", vbOKOnly, "Fuel Weight Limit Exceeded")
            txtFOB.SetFocus
        End If
    End If
    
    intFOB = Val(txtFOB)
    
    intCargoP = intP1 + intP2 + intP3 + intP4 + intP5 + intP6 + intP7 + intP8 + intP9 + intP10 + intP11
    intTOWt = intBEW + intFOB + intCargoP + intCrew - 35
    txtTOWt = intTOWt
    txtTOWt = Format(txtTOWt.Text, "Fixed")
    If intTOWt > 8750 Then
        Call TOWt
        txtFOB.SetFocus
        Exit Sub
    End If

End Sub

'Private Sub txtFOB_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    'If Button = vbRightButton Then
        'frmMain.PopupMenu mnuPopupLabel
        'intMouseCheck = 2
   ' End If
        
'End Sub

Private Sub txtFuel_GotFocus()
    
    If txtFuel = "Enter in Pounds" Then
        txtFuel = ""
    End If
    txtFuel.SelStart = 0
    txtFuel.SelLength = Len(txtFuel.Text)
    
End Sub

Private Sub txtFuel_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 48 And KeyAscii <= 57 Then
        'OK
    Else
        If KeyAscii = 46 Then
            'Decimal OK
        Else
            If KeyAscii = 8 Then
                'Backspace OK
            Else
                KeyAscii = 0
                Beep
            End If
        End If
        
    End If
End Sub

Private Sub txtFuel_LostFocus()
    
    If txtFOB = "" Then
        txtFOB.SetFocus
        Exit Sub
    ElseIf Val(txtFuel) > (Val(txtFOB) - 35) Then
        Call MsgBox("Trip Fuel can not exceed Fuel On Board less Taxi Fuel.  Please re-enter a valid value.", vbOKOnly)
        txtFuel = "Enter in Pounds"
        txtFuel.SetFocus
    End If
        
End Sub

Private Sub txtOrig_DblClick()
    txtOrig.Text = ""
    cmbOrig.Text = ""
    cmbOrig.Visible = True
End Sub

Private Sub txtOrig_GotFocus()
    txtOrig.TabIndex = 7
End Sub


Private Sub txtZone1_GotFocus()
    txtZone1.SelStart = 0
    txtZone1.SelLength = Len(txtZone1.Text)
End Sub

Private Sub txtZone1_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 48 And KeyAscii <= 57 Then
        'OK
    Else
        If KeyAscii = 46 Then
            'Decimal OK
        Else
            If KeyAscii = 8 Then
                'Backspace OK
            Else
                KeyAscii = 0
                Beep
            End If
        End If
        
    End If
End Sub

Private Sub txtZone1_LostFocus()
    
        If Frame1.Caption = "CARGO in Kilograms" Then
            If txtZone1.Text <> "" Then
                If IsNumeric(txtZone1.Text) Then
                    If txtZone1.Text > 807 Then
                        Call MsgBox("Please enter a number between 0 - 807", vbOKOnly, "Invalid Value")
                        txtZone1.Text = ""
                        txtZone1.SetFocus
                    End If
                Else
                    Call MsgBox("Please enter a number between 0 - 807", vbOKOnly, "Invalid Value")
                    txtZone1.Text = ""
                    txtZone1.SetFocus
                End If
                    If txtZone1.Text <> "" Then
                        txtZone1 = Format(txtZone1.Text, "Fixed")
                        int1 = Val(txtZone1)
                        intK1 = int1
                        intP1 = int1 * 2.20458553791887
                    End If
            Else
                int1 = 0
                intK1 = 0
                intP1 = 0
            End If
        Else
            If txtZone1.Text <> "" Then
                If IsNumeric(txtZone1.Text) Then
                    If txtZone1.Text > 1780 Then
                        Call MsgBox("Please enter a number between 0 - 1780", vbOKOnly, "Invalid Value")
                        txtZone1.Text = ""
                        txtZone1.SetFocus
                    End If
                Else
                    Call MsgBox("Please enter a number between 0 - 1780", vbOKOnly, "Invalid Value")
                    txtZone1.Text = ""
                    txtZone1.SetFocus
                End If
                    If txtZone1.Text <> "" Then
                        txtZone1 = Format(txtZone1.Text, "Fixed")
                        int1 = Val(txtZone1)
                        intK1 = int1 / 2.20458553791887
                        intP1 = int1
        
                    End If
            Else
                int1 = 0
                intK1 = 0
                intP1 = 0
            End If
        End If

    intTotal = int1 + int2 + int3 + int4 + int5 + int6 + int7 + int8 + int9 + int10 + int11
    intTotal1 = intP1 + intP2 + intP3 + intP4 + intP5 + intP6 + intP7 + intP8 + intP9 + intP10 + intP11
    intCargoP = intP1 + intP2 + intP3 + intP4 + intP5 + intP6 + intP7 + intP8 + intP9 + intP10 + intP11
    txtTotal = intTotal
    txtTotal = Format(txtTotal.Text, "Fixed")
    intTOWt = intBEW + intFOB + intCargoP + intCrew - 35
    txtTOWt = intTOWt
    txtTOWt = Format(txtTOWt.Text, "Fixed")
    Call Zone1

End Sub

Private Sub Zone1()
  
    If Frame1.Caption = "CARGO in Kilograms" Then
  
        If Val(txtTotal) > 1542 Then
            txtTotal.Text = ""
            txtZone1.Text = ""
            int1 = 0
            intK1 = 0
            intP1 = 0
            Call OverLimit
            txtZone1.SetFocus
        End If
        
    Else
        
        If Val(txtTotal) > 3400 Then
            txtTotal.Text = ""
            txtZone1.Text = ""
            int1 = 0
            intK1 = 0
            intP1 = 0
            Call OverLimit
            txtZone1.SetFocus
        End If
        
    End If
    
End Sub

Private Sub txtZone2_GotFocus()
    txtZone2.SelStart = 0
    txtZone2.SelLength = Len(txtZone2.Text)
End Sub

Private Sub txtZone2_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 48 And KeyAscii <= 57 Then
        'OK
    Else
        If KeyAscii = 46 Then
            'Decimal OK
        Else
            If KeyAscii = 8 Then
                'Backspace OK
            Else
                KeyAscii = 0
                Beep
            End If
        End If
        
    End If
End Sub

Private Sub txtZone2_LostFocus()
    
        If Frame1.Caption = "CARGO in Kilograms" Then
            If txtZone2.Text <> "" Then
                If IsNumeric(txtZone2.Text) Then
                    If txtZone2.Text > 1406 Then
                        Call MsgBox("Please enter a number between 0 - 1406", vbOKOnly, "Invalid Value")
                        txtZone2.Text = ""
                        txtZone2.SetFocus
                    End If
                Else
                    Call MsgBox("Please enter a number between 0 - 1406", vbOKOnly, "Invalid Value")
                    txtZone2.Text = ""
                    txtZone2.SetFocus
                End If
                If txtZone2.Text <> "" Then
                    txtZone2 = Format(txtZone2.Text, "Fixed")
                    int2 = Val(txtZone2)
                    intK2 = int2
                    intP2 = int2 * 2.20458553791887
                End If
            Else
                int2 = 0
                intK2 = 0
                intP2 = 0
            End If
        Else
            If txtZone2.Text <> "" Then
                If IsNumeric(txtZone2.Text) Then
                    If txtZone2.Text > 3100 Then
                        Call MsgBox("Please enter a number between 0 - 3100", vbOKOnly, "Invalid Value")
                        txtZone2.Text = ""
                        txtZone2.SetFocus
                    End If
                Else
                    Call MsgBox("Please enter a number between 0 - 3100", vbOKOnly, "Invalid Value")
                    txtZone2.Text = ""
                    txtZone2.SetFocus
                End If
                If txtZone2.Text <> "" Then
                    txtZone2 = Format(txtZone2.Text, "Fixed")
                    int2 = Val(txtZone2)
                    intK2 = int2 / 2.20458553791887
                    intP2 = int2
                End If
            Else
                int2 = 0
                intK2 = 0
                intP2 = 0
            End If
        End If
        
    intTotal = int1 + int2 + int3 + int4 + int5 + int6 + int7 + int8 + int9 + int10 + int11
    intTotal1 = intP1 + intP2 + intP3 + intP4 + intP5 + intP6 + intP7 + intP8 + intP9 + intP10 + intP11
    intCargoP = intP1 + intP2 + intP3 + intP4 + intP5 + intP6 + intP7 + intP8 + intP9 + intP10 + intP11
    txtTotal = intTotal
    txtTotal = Format(txtTotal.Text, "Fixed")
    intTOWt = intBEW + intFOB + intCargoP + intCrew - 35
    txtTOWt = intTOWt
    txtTOWt = Format(txtTOWt.Text, "Fixed")

    Call Zone2

End Sub

Private Sub Zone2()

    If Frame1.Caption = "CARGO in Kilograms" Then
  
        If Val(txtTotal) > 1542 Then
            txtTotal.Text = ""
            txtZone2.Text = ""
            int2 = 0
            intK2 = 0
            intP2 = 0
            Call OverLimit
            txtZone2.SetFocus
        End If
        
    Else
        
        If Val(txtTotal) > 3400 Then
            txtTotal.Text = ""
            txtZone2.Text = ""
            int2 = 0
            intK2 = 0
            intP2 = 0
            Call OverLimit
            txtZone2.SetFocus
        End If
        
    End If

End Sub

Private Sub txtZone3_GotFocus()
    txtZone3.SelStart = 0
    txtZone3.SelLength = Len(txtZone3.Text)
End Sub

Private Sub txtZone3_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 48 And KeyAscii <= 57 Then
        'OK
    Else
        If KeyAscii = 46 Then
            'Decimal OK
        Else
            If KeyAscii = 8 Then
                'Backspace OK
            Else
                KeyAscii = 0
                Beep
            End If
        End If
        
    End If
End Sub

Private Sub txtZone3_LostFocus()
    
        If Frame1.Caption = "CARGO in Kilograms" Then
            If txtZone3.Text <> "" Then
                If IsNumeric(txtZone3.Text) Then
                    If txtZone3.Text > 862 Then
                        Call MsgBox("Please enter a number between 0 - 862", vbOKOnly, "Invalid Value")
                        txtZone3.Text = ""
                        txtZone3.SetFocus
                    End If
                Else
                    Call MsgBox("Please enter a number between 0 - 862", vbOKOnly, "Invalid Value")
                    txtZone3.Text = ""
                    txtZone3.SetFocus
                End If
                If txtZone3.Text <> "" Then
                    txtZone3 = Format(txtZone3.Text, "Fixed")
                    int3 = Val(txtZone3)
                    intK3 = int3
                    intP3 = int3 * 2.20458553791887
                End If
            Else
                int3 = 0
                intK3 = 0
                intP3 = 0
            End If
        Else
            If txtZone3.Text <> "" Then
                If IsNumeric(txtZone3.Text) Then
                    If txtZone3.Text > 1900 Then
                        Call MsgBox("Please enter a number between 0 - 1900", vbOKOnly, "Invalid Value")
                        txtZone3.Text = ""
                        txtZone3.SetFocus
                    End If
                Else
                    Call MsgBox("Please enter a number between 0 - 1900", vbOKOnly, "Invalid Value")
                    txtZone3.Text = ""
                    txtZone3.SetFocus
                End If
                If txtZone3.Text <> "" Then
                    txtZone3 = Format(txtZone3.Text, "Fixed")
                    int3 = Val(txtZone3)
                    intK3 = int3 / 2.20458553791887
                    intP3 = int3
                End If
            Else
                int3 = 0
                intK3 = 0
                intP3 = 0
            End If
        End If

    intTotal = int1 + int2 + int3 + int4 + int5 + int6 + int7 + int8 + int9 + int10 + int11
    intTotal1 = intP1 + intP2 + intP3 + intP4 + intP5 + intP6 + intP7 + intP8 + intP9 + intP10 + intP11
    intCargoP = intP1 + intP2 + intP3 + intP4 + intP5 + intP6 + intP7 + intP8 + intP9 + intP10 + intP11
    txtTotal = intTotal
    txtTotal = Format(txtTotal.Text, "Fixed")
    intTOWt = intBEW + intFOB + intCargoP + intCrew - 35
    txtTOWt = intTOWt
    txtTOWt = Format(txtTOWt.Text, "Fixed")

    Call Zone3

End Sub

Private Sub Zone3()

    If Frame1.Caption = "CARGO in Kilograms" Then
  
        If Val(txtTotal) > 1542 Then
            txtTotal.Text = ""
            txtZone3.Text = ""
            int3 = 0
            intK3 = 0
            intP3 = 0
            Call OverLimit
            txtZone3.SetFocus
        End If
        
    Else
        
        If Val(txtTotal) > 3400 Then
            txtTotal.Text = ""
            txtZone3.Text = ""
            int3 = 0
            intK3 = 0
            intP3 = 0
            Call OverLimit
            txtZone3.SetFocus
        End If
        
    End If

End Sub

Private Sub txtZone4_GotFocus()
    txtZone4.SelStart = 0
    txtZone4.SelLength = Len(txtZone4.Text)
End Sub

Private Sub txtZone4_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 48 And KeyAscii <= 57 Then
        'OK
    Else
        If KeyAscii = 46 Then
            'Decimal OK
        Else
            If KeyAscii = 8 Then
                'Backspace OK
            Else
                KeyAscii = 0
                Beep
            End If
        End If
        
    End If
End Sub

Private Sub txtZone4_LostFocus()
    
        If Frame1.Caption = "CARGO in Kilograms" Then
            If txtZone4.Text <> "" Then
                If IsNumeric(txtZone4.Text) Then
                    If txtZone4.Text > 626 Then
                        Call MsgBox("Please enter a number between 0 - 626", vbOKOnly, "Invalid Value")
                        txtZone4.Text = ""
                        txtZone4.SetFocus
                    End If
                Else
                    Call MsgBox("Please enter a number between 0 - 626", vbOKOnly, "Invalid Value")
                    txtZone4.Text = ""
                    txtZone4.SetFocus
                End If
                If txtZone4.Text <> "" Then
                    txtZone4 = Format(txtZone4.Text, "Fixed")
                    int4 = Val(txtZone4)
                    intK4 = int4
                    intP4 = int4 * 2.20458553791887
                End If
            Else
                int4 = 0
                intK4 = 0
                intP4 = 0
            End If
        Else
            If txtZone4.Text <> "" Then
                If IsNumeric(txtZone4.Text) Then
                    If txtZone4.Text > 1380 Then
                        Call MsgBox("Please enter a number between 0 - 1380", vbOKOnly, "Invalid Value")
                        txtZone4.Text = ""
                        txtZone4.SetFocus
                    End If
                Else
                    Call MsgBox("Please enter a number between 0 - 1380", vbOKOnly, "Invalid Value")
                    txtZone4.Text = ""
                    txtZone4.SetFocus
                End If
                If txtZone4.Text <> "" Then
                    txtZone4 = Format(txtZone4.Text, "Fixed")
                    int4 = Val(txtZone4)
                    intK4 = int4 / 2.20458553791887
                    intP4 = int4
                End If
            Else
                int4 = 0
                intK4 = 0
                intP4 = 0
            End If
        End If

    intTotal = int1 + int2 + int3 + int4 + int5 + int6 + int7 + int8 + int9 + int10 + int11
    intTotal1 = intP1 + intP2 + intP3 + intP4 + intP5 + intP6 + intP7 + intP8 + intP9 + intP10 + intP11
    intCargoP = intP1 + intP2 + intP3 + intP4 + intP5 + intP6 + intP7 + intP8 + intP9 + intP10 + intP11
    txtTotal = intTotal
    txtTotal = Format(txtTotal.Text, "Fixed")
    intTOWt = intBEW + intFOB + intCargoP + intCrew - 35
    txtTOWt = intTOWt
    txtTOWt = Format(txtTOWt.Text, "Fixed")

    Call Zone4

End Sub

Private Sub Zone4()

    If Frame1.Caption = "CARGO in Kilograms" Then
  
        If Val(txtTotal) > 1542 Then
            txtTotal.Text = ""
            txtZone4.Text = ""
            int4 = 0
            intK4 = 0
            intP4 = 0
            Call OverLimit
            txtZone4.SetFocus
        End If
        
    Else
        
        If Val(txtTotal) > 3400 Then
            txtTotal.Text = ""
            txtZone4.Text = ""
            int4 = 0
            intK4 = 0
            intP4 = 0
            Call OverLimit
            txtZone4.SetFocus
        End If
        
    End If

End Sub

Private Sub txtZone5_GotFocus()
    txtZone5.SelStart = 0
    txtZone5.SelLength = Len(txtZone5.Text)
End Sub

Private Sub txtZone5_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 48 And KeyAscii <= 57 Then
        'OK
    Else
        If KeyAscii = 46 Then
            'Decimal OK
        Else
            If KeyAscii = 8 Then
                'Backspace OK
            Else
                KeyAscii = 0
                Beep
            End If
        End If
        
    End If
End Sub

Private Sub txtZone5_LostFocus()

        If Frame1.Caption = "CARGO in Kilograms" Then
            If txtZone5.Text <> "" Then
                If IsNumeric(txtZone5.Text) Then
                    If txtZone5.Text > 575 Then
                        Call MsgBox("Please enter a number between 0 - 575", vbOKOnly, "Invalid Value")
                        txtZone5.Text = ""
                        txtZone5.SetFocus
                    End If
                Else
                    Call MsgBox("Please enter a number between 0 - 575", vbOKOnly, "Invalid Value")
                    txtZone5.Text = ""
                    txtZone5.SetFocus
                End If
                If txtZone5.Text <> "" Then
                    txtZone5 = Format(txtZone5.Text, "Fixed")
                    int5 = Val(txtZone5)
                    intK5 = int5
                    intP5 = int5 * 2.20458553791887
                End If
            Else
                int5 = 0
                intK5 = 0
                intP5 = 0
            End If
        Else
            If txtZone5.Text <> "" Then
                If IsNumeric(txtZone5.Text) Then
                    If txtZone5.Text > 1270 Then
                        Call MsgBox("Please enter a number between 0 - 1270", vbOKOnly, "Invalid Value")
                        txtZone5.Text = ""
                        txtZone5.SetFocus
                    End If
                Else
                    Call MsgBox("Please enter a number between 0 - 1270", vbOKOnly, "Invalid Value")
                    txtZone5.Text = ""
                    txtZone5.SetFocus
                End If
                If txtZone5.Text <> "" Then
                    txtZone5 = Format(txtZone5.Text, "Fixed")
                    int5 = Val(txtZone5)
                    intK5 = int5 / 2.20458553791887
                    intP5 = int5
                End If
            Else
                int5 = 0
                intK5 = 0
                intP5 = 0
            End If
        End If
    
    intTotal = int1 + int2 + int3 + int4 + int5 + int6 + int7 + int8 + int9 + int10 + int11
    intTotal1 = intP1 + intP2 + intP3 + intP4 + intP5 + intP6 + intP7 + intP8 + intP9 + intP10 + intP11
    intCargoP = intP1 + intP2 + intP3 + intP4 + intP5 + intP6 + intP7 + intP8 + intP9 + intP10 + intP11
    txtTotal = intTotal
    txtTotal = Format(txtTotal.Text, "Fixed")
    intTOWt = intBEW + intFOB + intCargoP + intCrew - 35
    txtTOWt = intTOWt
    txtTOWt = Format(txtTOWt.Text, "Fixed")

    Call Zone5

End Sub

Private Sub Zone5()

    If Frame1.Caption = "CARGO in Kilograms" Then
  
        If Val(txtTotal) > 1542 Then
            txtTotal.Text = ""
            txtZone5.Text = ""
            int5 = 0
            intK5 = 0
            intP5 = 0
            Call OverLimit
            txtZone5.SetFocus
        End If
        
    Else
        
        If Val(txtTotal) > 3400 Then
            txtTotal.Text = ""
            txtZone5.Text = ""
            int5 = 0
            intK5 = 0
            intP5 = 0
            Call OverLimit
            txtZone5.SetFocus
        End If
        
    End If

End Sub

Private Sub txtZone6_GotFocus()
    txtZone6.SelStart = 0
    txtZone6.SelLength = Len(txtZone6.Text)
End Sub

Private Sub txtZone6_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 48 And KeyAscii <= 57 Then
        'OK
    Else
        If KeyAscii = 46 Then
            'Decimal OK
        Else
            If KeyAscii = 8 Then
                'Backspace OK
            Else
                KeyAscii = 0
                Beep
            End If
        End If
        
    End If
End Sub

Private Sub txtZone6_LostFocus()

        If Frame1.Caption = "CARGO in Kilograms" Then
            If txtZone6.Text <> "" Then
                If IsNumeric(txtZone6.Text) Then
                    If txtZone6.Text > 145 Then
                        Call MsgBox("Please enter a number between 0 - 145", vbOKOnly, "Invalid Value")
                        txtZone6.Text = ""
                        txtZone6.SetFocus
                    End If
                Else
                    Call MsgBox("Please enter a number between 0 - 145", vbOKOnly, "Invalid Value")
                    txtZone6.Text = ""
                    txtZone6.SetFocus
                End If
                If txtZone6.Text <> "" Then
                    txtZone6 = Format(txtZone6.Text, "Fixed")
                    int6 = Val(txtZone6)
                    intK6 = int6
                    intP6 = int6 * 2.20458553791887
                End If
            Else
                int6 = 0
                intK6 = 0
                intP6 = 0
            End If
        Else
            If txtZone6.Text <> "" Then
                If IsNumeric(txtZone6.Text) Then
                    If txtZone6.Text > 320 Then
                        Call MsgBox("Please enter a number between 0 - 320", vbOKOnly, "Invalid Value")
                        txtZone6.Text = ""
                        txtZone6.SetFocus
                    End If
                Else
                    Call MsgBox("Please enter a number between 0 - 320", vbOKOnly, "Invalid Value")
                    txtZone6.Text = ""
                    txtZone6.SetFocus
                End If
                If txtZone6.Text <> "" Then
                    txtZone6 = Format(txtZone6.Text, "Fixed")
                    int6 = Val(txtZone6)
                    intK6 = int6 / 2.20458553791887
                    intP6 = int6
                End If
            Else
                int6 = 0
                intK6 = 0
                intP6 = 0
            End If
        End If

    intTotal = int1 + int2 + int3 + int4 + int5 + int6 + int7 + int8 + int9 + int10 + int11
    intTotal1 = intP1 + intP2 + intP3 + intP4 + intP5 + intP6 + intP7 + intP8 + intP9 + intP10 + intP11
    intCargoP = intP1 + intP2 + intP3 + intP4 + intP5 + intP6 + intP7 + intP8 + intP9 + intP10 + intP11
    txtTotal = intTotal
    txtTotal = Format(txtTotal.Text, "Fixed")
    intTOWt = intBEW + intFOB + intCargoP + intCrew - 35
    txtTOWt = intTOWt
    txtTOWt = Format(txtTOWt.Text, "Fixed")

    Call Zone6

End Sub

Private Sub Zone6()

    If Frame1.Caption = "CARGO in Kilograms" Then
  
        If Val(txtTotal) > 1542 Then
            txtTotal.Text = ""
            txtZone6.Text = ""
            int6 = 0
            intK6 = 0
            intP6 = 0
            Call OverLimit
            txtZone6.SetFocus
        End If
        
    Else
        
        If Val(txtTotal) > 3400 Then
            txtTotal.Text = ""
            txtZone6.Text = ""
            int6 = 0
            intK6 = 0
            intP6 = 0
            Call OverLimit
            txtZone6.SetFocus
        End If
        
    End If

End Sub

Private Sub OverLimit()
    Call MsgBox("Cargo Limit exceeded. Please re-enter values within limits", vbOKOnly, "Cargo Limit Exceeded!")
End Sub

Private Sub TOWt()

    intRetVal = MsgBox("Take Off Weight Limit exceeded with default Crew Weight." & vbNewLine & "Click 'Yes' to change Crew Weight or 'No' to re-enter values within limits", vbYesNo, "Take Off Weight Limit Exceeded!")

End Sub

Private Sub txtZoneA_GotFocus()
    txtZoneA.SelStart = 0
    txtZoneA.SelLength = Len(txtZoneA.Text)
End Sub

Private Sub txtZoneA_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 48 And KeyAscii <= 57 Then
        'OK
    Else
        If KeyAscii = 46 Then
            'Decimal OK
        Else
            If KeyAscii = 8 Then
                'Backspace OK
            Else
                KeyAscii = 0
                Beep
            End If
        End If
        
    End If
End Sub

Private Sub txtZoneA_LostFocus()

If Frame1.Caption = "CARGO in Kilograms" Then
    If txtZoneA.Text <> "" Then
        If IsNumeric(txtZoneA.Text) Then
            If txtZoneA.Text > 104 Then
                Call MsgBox("Please enter a number between 0 - 104", vbOKOnly, "Invalid Value")
                txtZoneA.Text = ""
                txtZoneA.SetFocus
            End If
        Else
            Call MsgBox("Please enter a number between 0 - 104", vbOKOnly, "Invalid Value")
            txtZoneA.Text = ""
            txtZoneA.SetFocus
        End If
        If txtZoneA.Text <> "" Then
            txtZoneA = Format(txtZoneA.Text, "Fixed")
            int7 = Val(txtZoneA)
            intK7 = int7
            intP7 = int7 * 2.20458553791887
        End If
    Else
        int7 = 0
        intK7 = 0
        intP7 = 0
    End If
Else
    If txtZoneA.Text <> "" Then
        If IsNumeric(txtZoneA.Text) Then
            If txtZoneA.Text > 230 Then
                Call MsgBox("Please enter a number between 0 - 230", vbOKOnly, "Invalid Value")
                txtZoneA.Text = ""
                txtZoneA.SetFocus
            End If
        Else
            Call MsgBox("Please enter a number between 0 - 230", vbOKOnly, "Invalid Value")
            txtZoneA.Text = ""
            txtZoneA.SetFocus
        End If
        If txtZoneA.Text <> "" Then
            txtZoneA = Format(txtZoneA.Text, "Fixed")
            int7 = Val(txtZoneA)
            intK7 = int7 / 2.20458553791887
            intP7 = int7
        End If
    Else
        int7 = 0
        intK7 = 0
        intP7 = 0
    End If
End If

    intTotal = int1 + int2 + int3 + int4 + int5 + int6 + int7 + int8 + int9 + int10 + int11
    intTotal1 = intP1 + intP2 + intP3 + intP4 + intP5 + intP6 + intP7 + intP8 + intP9 + intP10 + intP11
    intCargoP = intP1 + intP2 + intP3 + intP4 + intP5 + intP6 + intP7 + intP8 + intP9 + intP10 + intP11
    txtTotal = intTotal
    txtTotal = Format(txtTotal.Text, "Fixed")
    intTOWt = intBEW + intFOB + intCargoP + intCrew - 35
    txtTOWt = intTOWt
    txtTOWt = Format(txtTOWt.Text, "Fixed")

    Call ZoneA

End Sub

Private Sub ZoneA()

    If Frame1.Caption = "CARGO in Kilograms" Then
  
        If Val(txtTotal) > 1542 Then
            txtTotal.Text = ""
            txtZoneA.Text = ""
            int7 = 0
            intK7 = 0
            intP7 = 0
            Call OverLimit
            txtZoneA.SetFocus
        End If
        
    Else
        
        If Val(txtTotal) > 3400 Then
            txtTotal.Text = ""
            txtZoneA.Text = ""
            int7 = 0
            intK7 = 0
            intP7 = 0
            Call OverLimit
            txtZoneA.SetFocus
        End If
        
    End If

End Sub

Private Sub txtZoneB_GotFocus()
    txtZoneB.SelStart = 0
    txtZoneB.SelLength = Len(txtZoneB.Text)
End Sub

Private Sub txtZoneB_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 48 And KeyAscii <= 57 Then
        'OK
    Else
        If KeyAscii = 46 Then
            'Decimal OK
        Else
            If KeyAscii = 8 Then
                'Backspace OK
            Else
                KeyAscii = 0
                Beep
            End If
        End If
        
    End If
End Sub

Private Sub txtZoneB_LostFocus()

If Frame1.Caption = "CARGO in Kilograms" Then
    If txtZoneB.Text <> "" Then
        If IsNumeric(txtZoneB.Text) Then
            If txtZoneB.Text > 140 Then
                Call MsgBox("Please enter a number between 0 - 140", vbOKOnly, "Invalid Value")
                txtZoneB.Text = ""
                txtZoneB.SetFocus
            End If
        Else
            Call MsgBox("Please enter a number between 0 - 140", vbOKOnly, "Invalid Value")
            txtZoneB.Text = ""
            txtZoneB.SetFocus
        End If
        If txtZoneB.Text <> "" Then
            txtZoneB = Format(txtZoneB.Text, "Fixed")
            int8 = Val(txtZoneB)
            intK8 = int8
            intP8 = int8 * 2.20458553791887
        End If
    Else
        int8 = 0
        intK8 = 0
        intP8 = 0
    End If
Else
    If txtZoneB.Text <> "" Then
        If IsNumeric(txtZoneB.Text) Then
            If txtZoneB.Text > 310 Then
                Call MsgBox("Please enter a number between 0 - 310", vbOKOnly, "Invalid Value")
                txtZoneB.Text = ""
                txtZoneB.SetFocus
            End If
        Else
            Call MsgBox("Please enter a number between 0 - 310", vbOKOnly, "Invalid Value")
            txtZoneB.Text = ""
            txtZoneB.SetFocus
        End If
        If txtZoneB.Text <> "" Then
            txtZoneB = Format(txtZoneB.Text, "Fixed")
            int8 = Val(txtZoneB)
            intK8 = int8 / 2.20458553791887
            intP8 = int8
        End If
    Else
        int8 = 0
        intK8 = 0
        intP8 = 0
    End If
End If

    intTotal = int1 + int2 + int3 + int4 + int5 + int6 + int7 + int8 + int9 + int10 + int11
    intTotal1 = intP1 + intP2 + intP3 + intP4 + intP5 + intP6 + intP7 + intP8 + intP9 + intP10 + intP11
    intCargoP = intP1 + intP2 + intP3 + intP4 + intP5 + intP6 + intP7 + intP8 + intP9 + intP10 + intP11
    txtTotal = intTotal
    txtTotal = Format(txtTotal.Text, "Fixed")
    intTOWt = intBEW + intFOB + intCargoP + intCrew - 35
    txtTOWt = intTOWt
    txtTOWt = Format(txtTOWt.Text, "Fixed")

    Call ZoneB

End Sub

Private Sub ZoneB()

    If Frame1.Caption = "CARGO in Kilograms" Then
  
        If Val(txtTotal) > 1542 Then
            txtTotal.Text = ""
            txtZoneB.Text = ""
            int8 = 0
            intK8 = 0
            intP8 = 0
            Call OverLimit
            txtZoneB.SetFocus
        End If
        
    Else
        
        If Val(txtTotal) > 3400 Then
            txtTotal.Text = ""
            txtZoneB.Text = ""
            int8 = 0
            intK8 = 0
            intP8 = 0
            Call OverLimit
            txtZoneB.SetFocus
        End If
        
    End If

End Sub


Private Sub txtZoneC_GotFocus()
    txtZoneC.SelStart = 0
    txtZoneC.SelLength = Len(txtZoneC.Text)
End Sub

Private Sub txtZoneC_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 48 And KeyAscii <= 57 Then
        'OK
    Else
        If KeyAscii = 46 Then
            'Decimal OK
        Else
            If KeyAscii = 8 Then
                'Backspace OK
            Else
                KeyAscii = 0
                Beep
            End If
        End If
        
    End If
End Sub

Private Sub txtZoneC_LostFocus()

If Frame1.Caption = "CARGO in Kilograms" Then
    If txtZoneC.Text <> "" Then
        If IsNumeric(txtZoneC.Text) Then
            If txtZoneC.Text > 122 Then
                Call MsgBox("Please enter a number between 0 - 122", vbOKOnly, "Invalid Value")
                txtZoneC.Text = ""
                txtZoneC.SetFocus
            End If
        Else
            Call MsgBox("Please enter a number between 0 - 122", vbOKOnly, "Invalid Value")
            txtZoneC.Text = ""
            txtZoneC.SetFocus
        End If
        If txtZoneC.Text <> "" Then
            txtZoneC = Format(txtZoneC.Text, "Fixed")
            int9 = Val(txtZoneC)
            intK9 = int9
            intP9 = int9 * 2.20458553791887
        End If
    Else
        int9 = 0
        intK9 = 0
        intP9 = 0
    End If
Else
    If txtZoneC.Text <> "" Then
        If IsNumeric(txtZoneC.Text) Then
            If txtZoneC.Text > 270 Then
                Call MsgBox("Please enter a number between 0 - 270", vbOKOnly, "Invalid Value")
                txtZoneC.Text = ""
                txtZoneC.SetFocus
            End If
        Else
            Call MsgBox("Please enter a number between 0 - 270", vbOKOnly, "Invalid Value")
            txtZoneC.Text = ""
            txtZoneC.SetFocus
        End If
        If txtZoneC.Text <> "" Then
            txtZoneC = Format(txtZoneC.Text, "Fixed")
            int9 = Val(txtZoneC)
            intK9 = int9 / 2.20458553791887
            intP9 = int9
        End If
    Else
        int9 = 0
        intK9 = 0
        intP9 = 0
    End If
End If

    intTotal = int1 + int2 + int3 + int4 + int5 + int6 + int7 + int8 + int9 + int10 + int11
    intTotal1 = intP1 + intP2 + intP3 + intP4 + intP5 + intP6 + intP7 + intP8 + intP9 + intP10 + intP11
    intCargoP = intP1 + intP2 + intP3 + intP4 + intP5 + intP6 + intP7 + intP8 + intP9 + intP10 + intP11
    txtTotal = intTotal
    txtTotal = Format(txtTotal.Text, "Fixed")
    intTOWt = intBEW + intFOB + intCargoP + intCrew - 35
    txtTOWt = intTOWt
    txtTOWt = Format(txtTOWt.Text, "Fixed")

    Call ZoneC

End Sub

Private Sub ZoneC()

    If Frame1.Caption = "CARGO in Kilograms" Then
  
        If Val(txtTotal) > 1542 Then
            txtTotal.Text = ""
            txtZoneC.Text = ""
            int9 = 0
            intK9 = 0
            intP9 = 0
            Call OverLimit
            txtZoneC.SetFocus
        End If
        
    Else
        
        If Val(txtTotal) > 3400 Then
            txtTotal.Text = ""
            txtZoneC.Text = ""
            int9 = 0
            intK9 = 0
            intP9 = 0
            Call OverLimit
            txtZoneC.SetFocus
        End If
        
    End If

End Sub

Private Sub txtZoneD_GotFocus()
    txtZoneD.SelStart = 0
    txtZoneD.SelLength = Len(txtZoneD.Text)
End Sub

Private Sub txtZoneD_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 48 And KeyAscii <= 57 Then
        'OK
    Else
        If KeyAscii = 46 Then
            'Decimal OK
        Else
            If KeyAscii = 8 Then
                'Backspace OK
            Else
                KeyAscii = 0
                Beep
            End If
        End If
        
    End If
End Sub

Private Sub txtZoneD_LostFocus()

If Frame1.Caption = "CARGO in Kilograms" Then
    If txtZoneD.Text <> "" Then
        If IsNumeric(txtZoneD.Text) Then
            If txtZoneD.Text > 127 Then
                Call MsgBox("Please enter a number between 0 - 127", vbOKOnly, "Invalid Value")
                txtZoneD.Text = ""
                txtZoneD.SetFocus
            End If
        Else
            Call MsgBox("Please enter a number between 0 - 127", vbOKOnly, "Invalid Value")
            txtZoneD.Text = ""
            txtZoneD.SetFocus
        End If
        If txtZoneD.Text <> "" Then
            txtZoneD = Format(txtZoneD.Text, "Fixed")
            int10 = Val(txtZoneD)
            intK10 = int10
            intP10 = int10 * 2.20458553791887
        End If
    Else
        int10 = 0
        intK10 = 0
        intP10 = 0
    End If
Else
    If txtZoneD.Text <> "" Then
        If IsNumeric(txtZoneD.Text) Then
            If txtZoneD.Text > 280 Then
                Call MsgBox("Please enter a number between 0 - 280", vbOKOnly, "Invalid Value")
                txtZoneD.Text = ""
                txtZoneD.SetFocus
            End If
        Else
            Call MsgBox("Please enter a number between 0 - 280", vbOKOnly, "Invalid Value")
            txtZoneD.Text = ""
            txtZoneD.SetFocus
        End If
        If txtZoneD.Text <> "" Then
            txtZoneD = Format(txtZoneD.Text, "Fixed")
            int10 = Val(txtZoneD)
            intK10 = int10 / 2.20458553791887
            intP10 = int10
        End If
    Else
        int10 = 0
        intK10 = 0
        intP10 = 0
    End If
End If

    intTotal = int1 + int2 + int3 + int4 + int5 + int6 + int7 + int8 + int9 + int10 + int11
    intTotal1 = intP1 + intP2 + intP3 + intP4 + intP5 + intP6 + intP7 + intP8 + intP9 + intP10 + intP11
    intCargoP = intP1 + intP2 + intP3 + intP4 + intP5 + intP6 + intP7 + intP8 + intP9 + intP10 + intP11
    txtTotal = intTotal
    txtTotal = Format(txtTotal.Text, "Fixed")
    intTOWt = intBEW + intFOB + intCargoP + intCrew - 35
    txtTOWt = intTOWt
    txtTOWt = Format(txtTOWt.Text, "Fixed")

    Call ZoneD

End Sub

Private Sub ZoneD()

    If Frame1.Caption = "CARGO in Kilograms" Then
  
        If Val(txtTotal) > 1542 Then
            txtTotal.Text = ""
            txtZoneD.Text = ""
            int10 = 0
            intK10 = 0
            intP10 = 0
            Call OverLimit
            txtZoneD.SetFocus
        End If
        
    Else
        
        If Val(txtTotal) > 3400 Then
            txtTotal.Text = ""
            txtZoneD.Text = ""
            int10 = 0
            intK10 = 0
            intP10 = 0
            Call OverLimit
            txtZoneD.SetFocus
        End If
        
    End If

End Sub

Private Sub txtSeat11_GotFocus()
    txtSeat11.SelStart = 0
    txtSeat11.SelLength = Len(txtSeat11.Text)
End Sub

Private Sub txtSeat11_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 48 And KeyAscii <= 57 Then
        'OK
    Else
        If KeyAscii = 46 Then
            'Decimal OK
        Else
            If KeyAscii = 8 Then
                'Backspace OK
            Else
                KeyAscii = 0
                Beep
            End If
        End If
        
    End If
End Sub

Private Sub txtSeat11_LostFocus()

If Frame1.Caption = "CARGO in Kilograms" Then
    If txtSeat11.Text <> "" Then
        If IsNumeric(txtSeat11.Text) Then
            If txtSeat11.Text > 127 Then
                Call MsgBox("Please enter a number between 0 - 127", vbOKOnly, "Invalid Value")
                txtSeat11.Text = ""
                txtSeat11.SetFocus
            End If
        Else
            Call MsgBox("Please enter a number between 0 - 127", vbOKOnly, "Invalid Value")
            txtSeat11.Text = ""
            txtSeat11.SetFocus
        End If
        If txtSeat11.Text <> "" Then
            txtSeat11 = Format(txtSeat11.Text, "Fixed")
            int11 = Val(txtSeat11)
            intK11 = int11
            intP11 = int11 * 2.20458553791887
        End If
    Else
        int11 = 0
        intK11 = 0
        intP11 = 0
    End If
Else
    If txtSeat11.Text <> "" Then
        If IsNumeric(txtSeat11.Text) Then
            If txtSeat11.Text > 320 Then
                Call MsgBox("Please enter a number between 0 - 320", vbOKOnly, "Invalid Value")
                txtSeat11.Text = ""
                txtSeat11.SetFocus
            End If
        Else
            Call MsgBox("Please enter a number between 0 - 320", vbOKOnly, "Invalid Value")
            txtSeat11.Text = ""
            txtSeat11.SetFocus
        End If
        If txtSeat11.Text <> "" Then
            txtSeat11 = Format(txtSeat11.Text, "Fixed")
            int11 = Val(txtSeat11)
            intK11 = int11 / 2.20458553791887
            intP11 = int11
        End If
    Else
        int11 = 0
        intK11 = 0
        intP11 = 0
    End If
End If

    intTotal = int1 + int2 + int3 + int4 + int5 + int6 + int7 + int8 + int9 + int10 + int11
    intTotal1 = intP1 + intP2 + intP3 + intP4 + intP5 + intP6 + intP7 + intP8 + intP9 + intP10 + intP11
    intCargoP = intP1 + intP2 + intP3 + intP4 + intP5 + intP6 + intP7 + intP8 + intP9 + intP10 + intP11
    txtTotal = intTotal
    txtTotal = Format(txtTotal.Text, "Fixed")
    intTOWt = intBEW + intFOB + intCargoP + intCrew - 35
    txtTOWt = intTOWt
    txtTOWt = Format(txtTOWt.Text, "Fixed")

    Call Zone11

End Sub

Private Sub Zone11()

    If Frame1.Caption = "CARGO in Kilograms" Then
  
        If Val(txtTotal) > 1542 Then
            txtTotal.Text = ""
            txtSeat11.Text = ""
            int11 = 0
            intK11 = 0
            intP11 = 0
            Call OverLimit
            txtSeat11.SetFocus
        End If
        
    Else
        
        If Val(txtTotal) > 3400 Then
            txtTotal.Text = ""
            txtSeat11.Text = ""
            int11 = 0
            intK11 = 0
            intP11 = 0
            Call OverLimit
            txtSeat11.SetFocus
        End If
        
    End If

End Sub

