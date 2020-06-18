VERSION 5.00
Begin VB.Form frmMainBU 
   Caption         =   "Weight & Balance and Fuel Calculations"
   ClientHeight    =   7875
   ClientLeft      =   60
   ClientTop       =   450
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
   Begin VB.TextBox txtTOWt 
      Height          =   390
      Left            =   6480
      TabIndex        =   50
      Top             =   7200
      Width           =   1575
   End
   Begin VB.TextBox txtFuel 
      Height          =   390
      Left            =   7560
      TabIndex        =   9
      Text            =   "Enter in Pounds"
      Top             =   4080
      Width           =   3975
   End
   Begin VB.TextBox txtSIC 
      Height          =   390
      Left            =   7560
      TabIndex        =   5
      Top             =   2640
      Width           =   3975
   End
   Begin VB.TextBox txtTaxiFuel 
      Height          =   390
      Left            =   10800
      TabIndex        =   10
      Text            =   "35"
      Top             =   4080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   10200
      TabIndex        =   18
      Top             =   7200
      Width           =   1215
   End
   Begin VB.ComboBox cmbDest 
      Height          =   390
      Left            =   7560
      TabIndex        =   7
      Top             =   3360
      Width           =   3975
   End
   Begin VB.ComboBox cmbOrig 
      Height          =   390
      Left            =   1680
      TabIndex        =   6
      Top             =   3360
      Width           =   3975
   End
   Begin VB.TextBox txtPIC 
      Height          =   390
      Left            =   1680
      TabIndex        =   4
      Top             =   2640
      Width           =   3975
   End
   Begin VB.TextBox txtSN 
      Height          =   390
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1920
      Width           =   1455
   End
   Begin VB.TextBox txtDep 
      Height          =   405
      Left            =   7200
      TabIndex        =   2
      Top             =   1920
      Width           =   1575
   End
   Begin VB.TextBox txtOrig 
      Height          =   405
      Left            =   1680
      TabIndex        =   19
      Top             =   3360
      Width           =   3975
   End
   Begin VB.TextBox txtDest 
      Height          =   405
      Left            =   7560
      TabIndex        =   20
      Top             =   3360
      Width           =   3975
   End
   Begin VB.TextBox txtFltNo 
      Height          =   405
      Left            =   10320
      TabIndex        =   3
      Text            =   "SWT"
      Top             =   1920
      Width           =   1215
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
      Height          =   390
      Left            =   9600
      TabIndex        =   31
      Top             =   6480
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "CARGO in Kilograms"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2415
      Left            =   600
      TabIndex        =   24
      Top             =   4680
      Width           =   10935
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
         Left            =   9000
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   1200
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
         Left            =   9000
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   480
         Width           =   1560
      End
      Begin VB.TextBox txtZone1 
         ForeColor       =   &H000000FF&
         Height          =   390
         Left            =   1125
         TabIndex        =   11
         Text            =   "0.00"
         Top             =   480
         Width           =   1075
      End
      Begin VB.TextBox txtZone2 
         ForeColor       =   &H000000FF&
         Height          =   390
         Left            =   4005
         TabIndex        =   12
         Text            =   "0.00"
         Top             =   480
         Width           =   1075
      End
      Begin VB.TextBox txtZone3 
         ForeColor       =   &H000000FF&
         Height          =   390
         Left            =   6930
         TabIndex        =   13
         Text            =   "0.00"
         Top             =   480
         Width           =   1075
      End
      Begin VB.TextBox txtZone4 
         ForeColor       =   &H000000FF&
         Height          =   390
         Left            =   1125
         TabIndex        =   14
         Text            =   "0.00"
         Top             =   1200
         Width           =   1075
      End
      Begin VB.TextBox txtZone5 
         ForeColor       =   &H000000FF&
         Height          =   390
         Left            =   4005
         TabIndex        =   15
         Text            =   "0.00"
         Top             =   1200
         Width           =   1075
      End
      Begin VB.TextBox txtZone6 
         ForeColor       =   &H000000FF&
         Height          =   390
         Left            =   6930
         TabIndex        =   16
         Text            =   "0.00"
         Top             =   1200
         Width           =   1075
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         Caption         =   "Total Cargo in Kgs (not to exceed 1542):"
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
         Height          =   270
         Left            =   3960
         TabIndex        =   48
         Top             =   1920
         Width           =   4815
      End
      Begin VB.Label Label22 
         Caption         =   "kgs."
         ForeColor       =   &H000000FF&
         Height          =   270
         Left            =   8160
         TabIndex        =   46
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label21 
         Caption         =   "kgs."
         ForeColor       =   &H000000FF&
         Height          =   270
         Left            =   8160
         TabIndex        =   45
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label20 
         Caption         =   "kgs."
         ForeColor       =   &H000000FF&
         Height          =   270
         Left            =   5160
         TabIndex        =   44
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label19 
         Caption         =   "kgs."
         ForeColor       =   &H000000FF&
         Height          =   270
         Left            =   5160
         TabIndex        =   43
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label18 
         Caption         =   "kgs."
         ForeColor       =   &H000000FF&
         Height          =   270
         Left            =   2280
         TabIndex        =   42
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label Label17 
         Caption         =   "kgs."
         ForeColor       =   &H000000FF&
         Height          =   270
         Left            =   2280
         TabIndex        =   41
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Hold 1:"
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
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Hold 2:"
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
         Height          =   255
         Left            =   2805
         TabIndex        =   29
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Hold 3:"
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
         Height          =   255
         Left            =   5685
         TabIndex        =   28
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Hold 4:"
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
         Height          =   255
         Left            =   165
         TabIndex        =   27
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Hold 5:"
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
         Height          =   255
         Left            =   2805
         TabIndex        =   26
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Hold 6:"
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
         Height          =   255
         Left            =   5730
         TabIndex        =   25
         Top             =   1320
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next"
      Height          =   375
      Left            =   8640
      TabIndex        =   17
      Top             =   7200
      Width           =   1215
   End
   Begin VB.TextBox txtFOB 
      Height          =   390
      Left            =   1680
      TabIndex        =   8
      Text            =   "Enter in Pounds"
      Top             =   4080
      Width           =   3975
   End
   Begin VB.ComboBox cmbAC 
      Height          =   390
      ItemData        =   "frmMainBU.frx":0000
      Left            =   1680
      List            =   "frmMainBU.frx":0002
      OLEDragMode     =   1  'Automatic
      TabIndex        =   0
      Top             =   1920
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   390
      Left            =   2280
      TabIndex        =   51
      Text            =   "Text1"
      Top             =   1920
      Visible         =   0   'False
      Width           =   210
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
      TabIndex        =   49
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
      Left            =   6360
      TabIndex        =   47
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   1335
      Left            =   1680
      Picture         =   "frmMainBU.frx":0004
      Stretch         =   -1  'True
      Top             =   240
      Width           =   3015
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
      TabIndex        =   40
      Top             =   2760
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
      Left            =   3960
      TabIndex        =   39
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
      Left            =   6240
      TabIndex        =   36
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
      TabIndex        =   35
      Top             =   3480
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
      Left            =   6720
      TabIndex        =   34
      Top             =   3480
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
      Left            =   9000
      TabIndex        =   33
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
      Left            =   5760
      TabIndex        =   32
      Top             =   480
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
      Left            =   6000
      TabIndex        =   23
      Top             =   3960
      Width           =   1455
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
      Left            =   480
      TabIndex        =   22
      Top             =   3960
      Width           =   1095
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
      TabIndex        =   21
      Top             =   2040
      Width           =   975
   End
   Begin VB.Menu mnuPopupLabel 
      Caption         =   "&mnuPopupLabel"
      Visible         =   0   'False
      Begin VB.Menu mnuPopupLabel1 
         Caption         =   "&35 - 999"
         Begin VB.Menu mnuPopupLabel199 
            Caption         =   "&35 - 99"
         End
         Begin VB.Menu mnuPopupLabel1199 
            Caption         =   "&100 - 199"
         End
         Begin VB.Menu mnuPopupLabel1299 
            Caption         =   "&200 - 299"
         End
         Begin VB.Menu mnuPopupLabel1399 
            Caption         =   "&300 - 399"
         End
         Begin VB.Menu mnuPopupLabel1499 
            Caption         =   "&400 - 499"
         End
         Begin VB.Menu mnuPopupLabel1599 
            Caption         =   "&500 - 599"
         End
      End
   End
   Begin VB.Menu mnuPopup1 
      Caption         =   "&mnu1"
      Visible         =   0   'False
      Begin VB.Menu mnuPopupLabel150 
         Caption         =   "&35 - 50"
         Begin VB.Menu mnuPopupLabel1aa35 
            Caption         =   "&35"
         End
         Begin VB.Menu mnuPopupLabel1aa36 
            Caption         =   "&36"
         End
         Begin VB.Menu mnuPopupLabel1aa37 
            Caption         =   "&37"
         End
         Begin VB.Menu mnuPopupLabel1aa38 
            Caption         =   "&38"
            Begin VB.Menu mnuPopupLabel1aa39 
               Caption         =   "&39"
            End
            Begin VB.Menu mnuPopupLabel1aa40 
               Caption         =   "&40"
            End
            Begin VB.Menu mnuPopupLabel1aa41 
               Caption         =   "&41"
            End
            Begin VB.Menu mnuPopupLabel1aa42 
               Caption         =   "&42"
            End
            Begin VB.Menu mnuPopupLabel1aa43 
               Caption         =   "&43"
            End
            Begin VB.Menu mnuPopupLabel1aa44 
               Caption         =   "&44"
            End
            Begin VB.Menu mnuPopupLabel1aa45 
               Caption         =   "&45"
            End
            Begin VB.Menu mnuPopupLabel1aa46 
               Caption         =   "&46"
            End
            Begin VB.Menu mnuPopupLabel1aa47 
               Caption         =   "&47"
            End
            Begin VB.Menu mnuPopupLabel1aa48 
               Caption         =   "&48"
            End
            Begin VB.Menu mnuPopupLabel1aa49 
               Caption         =   "&49"
            End
            Begin VB.Menu mnuPopupLabel1aa50 
               Caption         =   "&50"
            End
         End
         Begin VB.Menu mnuPopupLabel175 
            Caption         =   "&51 - 75"
            Begin VB.Menu mnuPopupLabel1aba 
               Caption         =   "&51 - 65"
               Begin VB.Menu mnuPopupLabel1ab51 
                  Caption         =   "&51"
               End
               Begin VB.Menu mnuPopupLabel1ab52 
                  Caption         =   "&52"
               End
               Begin VB.Menu mnuPopupLabel1ab53 
                  Caption         =   "&53"
               End
               Begin VB.Menu mnuPopupLabel1ab54 
                  Caption         =   "&54"
               End
               Begin VB.Menu mnuPopupLabel1ab55 
                  Caption         =   "&55"
               End
               Begin VB.Menu mnuPopupLabel1ab56 
                  Caption         =   "&56"
               End
               Begin VB.Menu mnuPopupLabel1ab57 
                  Caption         =   "&57"
               End
               Begin VB.Menu mnuPopupLabel1ab58 
                  Caption         =   "&58"
               End
               Begin VB.Menu mnuPopupLabel1ab59 
                  Caption         =   "&59"
               End
               Begin VB.Menu mnuPopupLabel1ab60 
                  Caption         =   "&60"
               End
               Begin VB.Menu mnuPopupLabel1ab61 
                  Caption         =   "&61"
               End
               Begin VB.Menu mnuPopupLabel1ab62 
                  Caption         =   "&62"
               End
               Begin VB.Menu mnuPopupLabel1ab63 
                  Caption         =   "&63"
               End
               Begin VB.Menu mnuPopupLabel1ab64 
                  Caption         =   "&64"
               End
               Begin VB.Menu mnuPopupLabel1ab65 
                  Caption         =   "&65"
               End
            End
            Begin VB.Menu mnuPopupLabel1abb 
               Caption         =   "&66 - 75"
               Begin VB.Menu mnuPopupLabel1abb66 
                  Caption         =   "&66"
               End
               Begin VB.Menu mnuPopupLabel1abb67 
                  Caption         =   "&67"
               End
               Begin VB.Menu mnuPopupLabel1abb68 
                  Caption         =   "&68"
               End
               Begin VB.Menu mnuPopupLabel1abb69 
                  Caption         =   "&69"
               End
               Begin VB.Menu mnuPopupLabel1abb70 
                  Caption         =   "&70"
               End
               Begin VB.Menu mnuPopupLabel1abb71 
                  Caption         =   "&71"
               End
               Begin VB.Menu mnuPopupLabel1abb72 
                  Caption         =   "&72"
               End
               Begin VB.Menu mnuPopupLabel1abb73 
                  Caption         =   "&73"
               End
               Begin VB.Menu mnuPopupLabel1abb74 
                  Caption         =   "&74"
               End
               Begin VB.Menu mnuPopupLabel1abb75 
                  Caption         =   "&75"
               End
            End
         End
         Begin VB.Menu mnuPopupLabel1ac 
            Caption         =   "&76 - 99"
            Begin VB.Menu mnuPopupLabel1aca 
               Caption         =   "&76 - 85"
               Begin VB.Menu mnuPopupLabel1aca76 
                  Caption         =   "&76"
               End
               Begin VB.Menu mnuPopupLabel1aca77 
                  Caption         =   "&77"
               End
               Begin VB.Menu mnuPopupLabel1aca78 
                  Caption         =   "&78"
               End
               Begin VB.Menu mnuPopupLabel1aca79 
                  Caption         =   "&79"
               End
               Begin VB.Menu mnuPopupLabel1aca80 
                  Caption         =   "&80"
               End
               Begin VB.Menu mnuPopupLabel1aca81 
                  Caption         =   "&81"
               End
               Begin VB.Menu mnuPopupLabel1aca82 
                  Caption         =   "&82"
               End
               Begin VB.Menu mnuPopupLabel1aca83 
                  Caption         =   "&83"
               End
               Begin VB.Menu mnuPopupLabel1aca84 
                  Caption         =   "&84"
               End
               Begin VB.Menu mnuPopupLabel1aca85 
                  Caption         =   "&85"
               End
            End
            Begin VB.Menu mnuPopupLabel1acb 
               Caption         =   "&86 - 99"
               Begin VB.Menu mnuPopupLabel1acb86 
                  Caption         =   "&86"
               End
               Begin VB.Menu mnuPopupLabel1acb87 
                  Caption         =   "&87"
               End
               Begin VB.Menu mnuPopupLabel1acb88 
                  Caption         =   "&88"
               End
               Begin VB.Menu mnuPopupLabel1acb89 
                  Caption         =   "&89"
               End
               Begin VB.Menu mnuPopupLabel1acb90 
                  Caption         =   "&90"
               End
               Begin VB.Menu mnuPopupLabel1acb91 
                  Caption         =   "&91"
               End
               Begin VB.Menu mnuPopupLabel1acb92 
                  Caption         =   "&92"
               End
               Begin VB.Menu mnuPopupLabel1acb93 
                  Caption         =   "&93"
               End
               Begin VB.Menu mnuPopupLabel1acb94 
                  Caption         =   "&94"
               End
               Begin VB.Menu mnuPopupLabel1acb95 
                  Caption         =   "&95"
               End
               Begin VB.Menu mnuPopupLabel1acb96 
                  Caption         =   "&96"
               End
               Begin VB.Menu mnuPopupLabel1acb97 
                  Caption         =   "&97"
               End
               Begin VB.Menu mnuPopupLabel1acb98 
                  Caption         =   "&98"
               End
               Begin VB.Menu mnuPopupLabel1acb99 
                  Caption         =   "&99"
               End
            End
            Begin VB.Menu mnuPopupLabel1ba 
               Caption         =   "&100 - 109"
               Begin VB.Menu mnuPopupLabel1ba100 
                  Caption         =   "&100"
               End
               Begin VB.Menu mnuPopupLabel1ba101 
                  Caption         =   "&101"
               End
               Begin VB.Menu mnuPopupLabel1ba102 
                  Caption         =   "&102"
               End
               Begin VB.Menu mnuPopupLabel1ba103 
                  Caption         =   "&103"
               End
               Begin VB.Menu mnuPopupLabel1ba104 
                  Caption         =   "&104"
               End
               Begin VB.Menu mnuPopupLabel1ba105 
                  Caption         =   "&105"
               End
               Begin VB.Menu mnuPopupLabel1ba106 
                  Caption         =   "&106"
               End
               Begin VB.Menu mnuPopupLabel1ba107 
                  Caption         =   "&107"
               End
               Begin VB.Menu mnuPopupLabel1ba108 
                  Caption         =   "&108"
               End
               Begin VB.Menu mnuPopupLabel1ba109 
                  Caption         =   "&109"
               End
            End
            Begin VB.Menu mnuPopupLabel1bb 
               Caption         =   "&110 - 119"
               Begin VB.Menu mnuPopupLabel1bb110 
                  Caption         =   "&110"
               End
               Begin VB.Menu mnuPopupLabel1bb111 
                  Caption         =   "&111"
               End
               Begin VB.Menu mnuPopupLabel1bb112 
                  Caption         =   "&112"
               End
               Begin VB.Menu mnuPopupLabel1bb113 
                  Caption         =   "&113"
               End
               Begin VB.Menu mnuPopupLabel1bb114 
                  Caption         =   "&114"
               End
               Begin VB.Menu mnuPopupLabel1bb115 
                  Caption         =   "&115"
               End
               Begin VB.Menu mnuPopupLabel1bb116 
                  Caption         =   "&116"
               End
               Begin VB.Menu mnuPopupLabel1bb117 
                  Caption         =   "&117"
               End
               Begin VB.Menu mnuPopupLabel1bb118 
                  Caption         =   "&118"
               End
               Begin VB.Menu mnuPopupLabel1bb119 
                  Caption         =   "&119"
               End
            End
            Begin VB.Menu mnuPopupLabel1bc 
               Caption         =   "&120 - 129"
               Begin VB.Menu mnuPopupLabel1bc120 
                  Caption         =   "&120"
               End
               Begin VB.Menu mnuPopupLabel1bc121 
                  Caption         =   "&121"
               End
               Begin VB.Menu mnuPopupLabel1bc122 
                  Caption         =   "&122"
               End
               Begin VB.Menu mnuPopupLabel1bc123 
                  Caption         =   "&123"
               End
               Begin VB.Menu mnuPopupLabel1bc124 
                  Caption         =   "&124"
               End
               Begin VB.Menu mnuPopupLabel1bc125 
                  Caption         =   "&125"
               End
               Begin VB.Menu mnuPopupLabel1bc126 
                  Caption         =   "&126"
               End
               Begin VB.Menu mnuPopupLabel1bc127 
                  Caption         =   "&127"
               End
               Begin VB.Menu mnuPopupLabel1bc128 
                  Caption         =   "&128"
               End
               Begin VB.Menu mnuPopupLabel1bc129 
                  Caption         =   "&129"
               End
            End
            Begin VB.Menu mnuPopupLabel1bd 
               Caption         =   "&130 - 139"
               Begin VB.Menu mnuPopupLabel1bd130 
                  Caption         =   "&130"
               End
               Begin VB.Menu mnuPopupLabel1bd131 
                  Caption         =   "&131"
               End
               Begin VB.Menu mnuPopupLabel1bd132 
                  Caption         =   "&132"
               End
               Begin VB.Menu mnuPopupLabel1bd133 
                  Caption         =   "&133"
               End
               Begin VB.Menu mnuPopupLabel1bd134 
                  Caption         =   "&134"
               End
               Begin VB.Menu mnuPopupLabel1bd135 
                  Caption         =   "&135"
               End
               Begin VB.Menu mnuPopupLabel1bd136 
                  Caption         =   "&136"
               End
               Begin VB.Menu mnuPopupLabel1bd137 
                  Caption         =   "&137"
               End
               Begin VB.Menu mnuPopupLabel1bd138 
                  Caption         =   "&138"
               End
               Begin VB.Menu mnuPopupLabel1bd139 
                  Caption         =   "&139"
               End
            End
            Begin VB.Menu mnuPopupLabel1be 
               Caption         =   "&140 - 149"
               Begin VB.Menu mnuPopupLabel1be140 
                  Caption         =   "&140"
               End
               Begin VB.Menu mnuPopupLabel1be141 
                  Caption         =   "&141"
               End
               Begin VB.Menu mnuPopupLabel1be142 
                  Caption         =   "&142"
               End
               Begin VB.Menu mnuPopupLabel1be143 
                  Caption         =   "&143"
               End
               Begin VB.Menu mnuPopupLabel1be144 
                  Caption         =   "&144"
               End
               Begin VB.Menu mnuPopupLabel1be145 
                  Caption         =   "&145"
               End
               Begin VB.Menu mnuPopupLabel1be146 
                  Caption         =   "&146"
               End
               Begin VB.Menu mnuPopupLabel1be147 
                  Caption         =   "&147"
               End
               Begin VB.Menu mnuPopupLabel1be148 
                  Caption         =   "&148"
               End
               Begin VB.Menu mnuPopupLabel1be149 
                  Caption         =   "&149"
               End
            End
            Begin VB.Menu mnuPopupLabel1bf 
               Caption         =   "&150 - 159"
               Begin VB.Menu mnuPopupLabel1bf150 
                  Caption         =   "&150"
               End
               Begin VB.Menu mnuPopupLabel1bf151 
                  Caption         =   "&151"
               End
               Begin VB.Menu mnuPopupLabel1bf152 
                  Caption         =   "&152"
               End
               Begin VB.Menu mnuPopupLabel1bf153 
                  Caption         =   "&153"
               End
               Begin VB.Menu mnuPopupLabel1bf154 
                  Caption         =   "&154"
               End
               Begin VB.Menu mnuPopupLabel1bf155 
                  Caption         =   "&155"
               End
               Begin VB.Menu mnuPopupLabel1bf156 
                  Caption         =   "&156"
               End
               Begin VB.Menu mnuPopupLabel1bf157 
                  Caption         =   "&157"
               End
               Begin VB.Menu mnuPopupLabel1bf158 
                  Caption         =   "&158"
               End
               Begin VB.Menu mnuPopupLabel1bf159 
                  Caption         =   "&159"
               End
            End
            Begin VB.Menu mnuPopupLabel1bg 
               Caption         =   "&160 - 169"
               Begin VB.Menu mnuPopupLabel1bg160 
                  Caption         =   "&160"
               End
               Begin VB.Menu mnuPopupLabel1bg161 
                  Caption         =   "&161"
               End
               Begin VB.Menu mnuPopupLabel1bg162 
                  Caption         =   "&162"
               End
               Begin VB.Menu mnuPopupLabel1bg163 
                  Caption         =   "&163"
               End
               Begin VB.Menu mnuPopupLabel1bg164 
                  Caption         =   "&164"
               End
               Begin VB.Menu mnuPopupLabel1bg165 
                  Caption         =   "&165"
               End
               Begin VB.Menu mnuPopupLabel1bg166 
                  Caption         =   "&166"
               End
               Begin VB.Menu mnuPopupLabel1bg167 
                  Caption         =   "&167"
               End
               Begin VB.Menu mnuPopupLabel1bg168 
                  Caption         =   "&168"
               End
               Begin VB.Menu mnuPopupLabel1bg169 
                  Caption         =   "&169"
               End
            End
            Begin VB.Menu mnuPopupLabel1bh 
               Caption         =   "&170 - 179"
               Begin VB.Menu mnuPopupLabel1bh170 
                  Caption         =   "&170"
               End
               Begin VB.Menu mnuPopupLabel1bh171 
                  Caption         =   "&171"
               End
               Begin VB.Menu mnuPopupLabel1bh172 
                  Caption         =   "&172"
               End
               Begin VB.Menu mnuPopupLabel1bh173 
                  Caption         =   "&173"
               End
               Begin VB.Menu mnuPopupLabel1bh174 
                  Caption         =   "&174"
               End
               Begin VB.Menu mnuPopupLabel1bh175 
                  Caption         =   "&175"
               End
               Begin VB.Menu mnuPopupLabel1bh176 
                  Caption         =   "&176"
               End
               Begin VB.Menu mnuPopupLabel1bh177 
                  Caption         =   "&177"
               End
               Begin VB.Menu mnuPopupLabel1bh178 
                  Caption         =   "&178"
               End
               Begin VB.Menu mnuPopupLabel1bh179 
                  Caption         =   "&179"
               End
            End
            Begin VB.Menu mnuPopupLabel1bi 
               Caption         =   "&180 - 189"
               Begin VB.Menu mnuPopupLabel1bi180 
                  Caption         =   "&180"
               End
               Begin VB.Menu mnuPopupLabel1bi181 
                  Caption         =   "&181"
               End
               Begin VB.Menu mnuPopupLabel1bi182 
                  Caption         =   "&182"
               End
               Begin VB.Menu mnuPopupLabel1bi183 
                  Caption         =   "&183"
               End
               Begin VB.Menu mnuPopupLabel1bi184 
                  Caption         =   "&184"
               End
               Begin VB.Menu mnuPopupLabel1bi185 
                  Caption         =   "&185"
               End
               Begin VB.Menu mnuPopupLabel1bi186 
                  Caption         =   "&186"
               End
               Begin VB.Menu mnuPopupLabel1bi187 
                  Caption         =   "&187"
               End
               Begin VB.Menu mnuPopupLabel1bi188 
                  Caption         =   "&188"
               End
               Begin VB.Menu mnuPopupLabel1bi189 
                  Caption         =   "&189"
               End
            End
            Begin VB.Menu mnuPopupLabel1bj 
               Caption         =   "&190 - 199"
               Begin VB.Menu mnuPopupLabel1bj190 
                  Caption         =   "&190"
               End
               Begin VB.Menu mnuPopupLabel1bj191 
                  Caption         =   "&191"
               End
               Begin VB.Menu mnuPopupLabel1bj192 
                  Caption         =   "&192"
               End
               Begin VB.Menu mnuPopupLabel1bj193 
                  Caption         =   "&193"
               End
               Begin VB.Menu mnuPopupLabel1bj194 
                  Caption         =   "&194"
               End
               Begin VB.Menu mnuPopupLabel1bj195 
                  Caption         =   "&195"
               End
               Begin VB.Menu mnuPopupLabel1bj196 
                  Caption         =   "&196"
               End
               Begin VB.Menu mnuPopupLabel1bj197 
                  Caption         =   "&197"
               End
               Begin VB.Menu mnuPopupLabel1bj198 
                  Caption         =   "&198"
               End
               Begin VB.Menu mnuPopupLabel1bj199 
                  Caption         =   "&199"
               End
            End
            Begin VB.Menu mnuPopupLabel1ca 
               Caption         =   "&200 - 209"
            End
            Begin VB.Menu mnuPopupLabel1cb 
               Caption         =   "&210 - 219"
            End
            Begin VB.Menu mnuPopupLabel1cc 
               Caption         =   "&220 - 229"
            End
            Begin VB.Menu mnuPopupLabel1cd 
               Caption         =   "&230 - 239"
            End
            Begin VB.Menu mnuPopupLabel1ce 
               Caption         =   "&240 - 249"
            End
            Begin VB.Menu mnuPopupLabel1cf 
               Caption         =   "&250 - 259"
            End
            Begin VB.Menu mnuPopupLabel1cg 
               Caption         =   "&260 - 269"
            End
            Begin VB.Menu mnuPopupLabel1ch 
               Caption         =   "&270 - 279"
            End
            Begin VB.Menu mnuPopupLabel1ci 
               Caption         =   "&280 - 289"
            End
            Begin VB.Menu mnuPopupLabel1cj 
               Caption         =   "&290 - 299"
            End
         End
      End
   End
End
Attribute VB_Name = "frmMainBU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim intTotal2 As Double, intMouseCheck As Integer

Private Sub cmbAC_Click()

    If cmbAC = "EC-IEV" Then
            
        txtSN = "208B0936"
        intFOB = Val(txtFOB)
        intBurn = Val(txtFuel)
        intBEW = 4816
        intArm1 = 180.23
        intMom1 = 867987.68
        
    Else
    
        txtSN = "208B0937"
        intFOB = Val(txtFOB)
        intBurn = Val(txtFuel)
        intBEW = 4839
        intArm1 = 180.93
        intMom1 = 875520.27
        
    End If
                                                                                            Text1 = int0
            If Text1 <> "2" Then
                Call MsgBox("Nice try Pal!", vbOKOnly, "No Changes Allowed!")
                End
            Else
                txtFltNo.SetFocus
            End If
                                            
End Sub

Private Sub cmbDest_Click()
    If cmbDest.Text = "Other" Then
        cmbDest.Visible = False
        txtDest.Text = ""
        txtDest.SetFocus
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
        Label2.ForeColor = &HFF&
        Label3.ForeColor = &HFF&
        Label4.ForeColor = &HFF&
        Label5.ForeColor = &HFF&
        Label6.ForeColor = &HFF&
        Label7.ForeColor = &HFF&
        Label8.ForeColor = &HFF&
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
    
    End If
    
    'intCargoK = int1 + int2 + int3 + int4 + int5 + int6
    txtTotal = Val(txtZone1) + Val(txtZone2) + Val(txtZone3) + Val(txtZone4) + Val(txtZone5) + Val(txtZone6)
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
        Label2.ForeColor = 12582912
        Label3.ForeColor = 12582912
        Label4.ForeColor = 12582912
        Label5.ForeColor = 12582912
        Label6.ForeColor = 12582912
        Label7.ForeColor = 12582912
        Label8.ForeColor = 12582912
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
    
    End If
        
    intCargoP = int1 + int2 + int3 + int4 + int5 + int6
    txtTotal = intCargoP
    txtTotal = Format(txtTotal.Text, "Fixed")

End Sub

Private Sub cmdNext_Click()

    If cmbAC = "EC-IEV" Then
            
        intFOB = Val(txtFOB)
        intBurn = Val(txtFuel)
        intBEW = 4816
        intArm1 = 180.23
        intMom1 = 867987.68
        
    Else
        
        intFOB = Val(txtFOB)
        intBurn = Val(txtFuel)
        intBEW = 4839
        intArm1 = 180.93
        intMom1 = 875520.27
    End If
                
    intCargoP = intP1 + intP2 + intP3 + intP4 + intP5 + intP6
    intTOWt = intBEW + intFOB + intCargoP + intCrew - 35
    txtTOWt = intTOWt
    txtTOWt = Format(txtTOWt.Text, "Fixed")
                                                                                                                                    If txtSN <> "208B0936" And txtSN <> "208B0937" Then
                                                                                                                                        Call MsgBox("Please enter a valid aircraft", vbOKOnly, "Invalid Aircraft!")
                                                                                                                                        cmbAC.SetFocus
                                                                                                                                        Exit Sub
                                                                                                                                    End If
    If intCargoP > 3400 Then
        Call Zone1
    Else
        intCargoK = intK1 + intK2 + intK3 + intK4 + intK5 + intK6
        intRampFuel = intFOB
        intTaxi = Val(txtTaxiFuel)
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

Private Sub Form_Load()

    cmbAC.AddItem "EC-IEV"
    cmbAC.AddItem "EC-IEX"
                                                                                       int0 = cmbAC.ListCount
    cmbOrig.AddItem "LEBL"
    cmbOrig.AddItem "LEMA"
    cmbOrig.AddItem "LEPA"
    cmbOrig.AddItem "Other"
    
    cmbDest.AddItem "LEBL"
    cmbDest.AddItem "LEMA"
    cmbDest.AddItem "LEPA"
    cmbDest.AddItem "Other"
    
    txtDep.Text = Date
    intCrew = 340
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub mnuPopupLabel1aa35_Click()
    txtFOB = 35
End Sub

Private Sub mnuPopupLabel1aa36_Click()
    txtFOB = 36
End Sub

Private Sub mnuPopupLabel1aa37_Click()
    txtFOB = 37
End Sub

Private Sub mnuPopupLabel1aa38_Click()
    txtFOB = 38
End Sub

Private Sub mnuPopupLabel1aa39_Click()
    txtFOB = 39
End Sub

Private Sub mnuPopupLabel1aa40_Click()
    txtFOB = 40
End Sub

Private Sub mnuPopupLabel1aa41_Click()
    txtFOB = 41
End Sub

Private Sub mnuPopupLabel1aa42_Click()
    txtFOB = 42
End Sub

Private Sub mnuPopupLabel1aa43_Click()
    txtFOB = 43
End Sub

Private Sub mnuPopupLabel1aa44_Click()
    txtFOB = 44
End Sub

Private Sub mnuPopupLabel1aa45_Click()
    txtFOB = 45
End Sub

Private Sub mnuPopupLabel1aa46_Click()
    txtFOB = 46
End Sub

Private Sub mnuPopupLabel1aa47_Click()
    txtFOB = 47
End Sub

Private Sub mnuPopupLabel1aa48_Click()
    txtFOB = 48
End Sub

Private Sub mnuPopupLabel1aa49_Click()
    txtFOB = 49
End Sub

Private Sub mnuPopupLabel1aa50_Click()
    txtFOB = 50
End Sub

Private Sub mnuPopupLabel1ab51_Click()
    txtFOB = 51
End Sub

Private Sub mnuPopupLabel1ab52_Click()
    txtFOB = 52
End Sub

Private Sub mnuPopupLabel1ab53_Click()
    txtFOB = 53
End Sub

Private Sub mnuPopupLabel1ab54_Click()
    txtFOB = 54
End Sub

Private Sub mnuPopupLabel1ab55_Click()
    txtFOB = 55
End Sub

Private Sub mnuPopupLabel1ab56_Click()
    txtFOB = 56
End Sub

Private Sub mnuPopupLabel1ab57_Click()
    txtFOB = 57
End Sub

Private Sub mnuPopupLabel1ab58_Click()
    txtFOB = 58
End Sub

Private Sub mnuPopupLabel1ab59_Click()
    txtFOB = 59
End Sub

Private Sub mnuPopupLabel1ab60_Click()
    txtFOB = 60
End Sub

Private Sub mnuPopupLabel1ab61_Click()
    txtFOB = 61
End Sub

Private Sub mnuPopupLabel1ab62_Click()
    txtFOB = 62
End Sub

Private Sub mnuPopupLabel1ab63_Click()
    txtFOB = 63
End Sub

Private Sub mnuPopupLabel1ab64_Click()
    txtFOB = 64
End Sub

Private Sub mnuPopupLabel1ab65_Click()
    txtFOB = 65
End Sub

Private Sub mnuPopupLabel1abb66_Click()
    txtFOB = 66
End Sub

Private Sub mnuPopupLabel1abb67_Click()
    txtFOB = 67
End Sub

Private Sub mnuPopupLabel1abb68_Click()
    txtFOB = 68
End Sub

Private Sub mnuPopupLabel1abb69_Click()
    txtFOB = 69
End Sub

Private Sub mnuPopupLabel1abb70_Click()
    txtFOB = 70
End Sub

Private Sub mnuPopupLabel1abb71_Click()
    txtFOB = 71
End Sub

Private Sub mnuPopupLabel1abb72_Click()
    txtFOB = 72
End Sub

Private Sub mnuPopupLabel1abb73_Click()
    txtFOB = 73
End Sub

Private Sub mnuPopupLabel1abb74_Click()
    txtFOB = 74
End Sub

Private Sub mnuPopupLabel1abb75_Click()
    txtFOB = 75
End Sub

Private Sub mnuPopupLabel1aca76_Click()
    txtFOB = 76
End Sub

Private Sub mnuPopupLabel1aca77_Click()
    txtFOB = 77
End Sub

Private Sub mnuPopupLabel1aca78_Click()
    txtFOB = 78
End Sub

Private Sub mnuPopupLabel1aca79_Click()
    txtFOB = 79
End Sub

Private Sub mnuPopupLabel1aca80_Click()
    txtFOB = 80
End Sub

Private Sub mnuPopupLabel1aca81_Click()
    txtFOB = 81
End Sub

Private Sub mnuPopupLabel1aca82_Click()
    txtFOB = 82
End Sub

Private Sub mnuPopupLabel1aca83_Click()
    txtFOB = 83
End Sub

Private Sub mnuPopupLabel1aca84_Click()
    txtFOB = 84
End Sub

Private Sub mnuPopupLabel1aca85_Click()
    txtFOB = 85
End Sub

Private Sub mnuPopupLabel1acb86_Click()
    txtFOB = 86
End Sub

Private Sub mnuPopupLabel1acb87_Click()
    txtFOB = 87
End Sub

Private Sub mnuPopupLabel1acb88_Click()
    txtFOB = 88
End Sub

Private Sub mnuPopupLabel1acb89_Click()
    txtFOB = 89
End Sub

Private Sub mnuPopupLabel1acb90_Click()
    txtFOB = 90
End Sub

Private Sub mnuPopupLabel1acb91_Click()
    txtFOB = 91
End Sub

Private Sub mnuPopupLabel1acb92_Click()
    txtFOB = 92
End Sub

Private Sub mnuPopupLabel1acb93_Click()
    txtFOB = 93
End Sub

Private Sub mnuPopupLabel1acb94_Click()
    txtFOB = 94
End Sub

Private Sub mnuPopupLabel1acb95_Click()
    txtFOB = 95
End Sub

Private Sub mnuPopupLabel1acb96_Click()
    txtFOB = 96
End Sub

Private Sub mnuPopupLabel1acb97_Click()
    txtFOB = 97
End Sub

Private Sub mnuPopupLabel1acb98_Click()
    txtFOB = 98
End Sub

Private Sub mnuPopupLabel1acb99_Click()
    txtFOB = 99
End Sub

Private Sub mnuPopupLabel1ba100_Click()
    txtFOB = 100
End Sub

Private Sub mnuPopupLabel1ba101_Click()
    txtFOB = 101
End Sub

Private Sub mnuPopupLabel1ba102_Click()
    txtFOB = 102
End Sub

Private Sub mnuPopupLabel1ba103_Click()
    txtFOB = 103
End Sub

Private Sub mnuPopupLabel1ba104_Click()
    txtFOB = 104
End Sub

Private Sub mnuPopupLabel1ba105_Click()
    txtFOB = 105
End Sub

Private Sub mnuPopupLabel1ba106_Click()
    txtFOB = 106
End Sub

Private Sub mnuPopupLabel1ba107_Click()
    txtFOB = 107
End Sub

Private Sub mnuPopupLabel1ba108_Click()
    txtFOB = 108
End Sub

Private Sub mnuPopupLabel1ba109_Click()
    txtFOB = 109
End Sub

Private Sub mnuPopupLabel1bb110_Click()
    txtFOB = 110
End Sub

Private Sub mnuPopupLabel1bb111_Click()
    txtFOB = 111
End Sub

Private Sub mnuPopupLabel1bb112_Click()
    txtFOB = 112
End Sub

Private Sub mnuPopupLabel1bb113_Click()
    txtFOB = 113
End Sub

Private Sub mnuPopupLabel1bb114_Click()
    txtFOB = 114
End Sub

Private Sub mnuPopupLabel1bb115_Click()
    txtFOB = 115
End Sub

Private Sub mnuPopupLabel1bb116_Click()
    txtFOB = 116
End Sub

Private Sub mnuPopupLabel1bb117_Click()
    txtFOB = 117
End Sub

Private Sub mnuPopupLabel1bb118_Click()
    txtFOB = 118
End Sub

Private Sub mnuPopupLabel1bb119_Click()
    txtFOB = 119
End Sub

Private Sub mnuPopupLabel1bc120_Click()
    txtFOB = 120
End Sub

Private Sub mnuPopupLabel1bc121_Click()
    txtFOB = 121
End Sub

Private Sub mnuPopupLabel1bc122_Click()
    txtFOB = 122
End Sub

Private Sub mnuPopupLabel1bc123_Click()
    txtFOB = 123
End Sub

Private Sub mnuPopupLabel1bc124_Click()
    txtFOB = 124
End Sub

Private Sub mnuPopupLabel1bc125_Click()
    txtFOB = 125
End Sub

Private Sub mnuPopupLabel1bc126_Click()
    txtFOB = 126
End Sub

Private Sub mnuPopupLabel1bc127_Click()
    txtFOB = 127
End Sub

Private Sub mnuPopupLabel1bc128_Click()
    txtFOB = 128
End Sub

Private Sub mnuPopupLabel1bc129_Click()
    txtFOB = 129
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
    txtFltNo.SelStart = 3
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
    If Val(txtFOB) > 2224 Then
        txtFOB = ""
        Call MsgBox("Fuel Wieght must not exceed 2224 pounds.  Please enter a value within limits.", vbOKOnly, "Fuel Weight Limit Exceeded")
        txtFOB.SetFocus
    ElseIf Val(txtFOB) < 35 Then
        txtFOB = ""
        Call MsgBox("Fuel Wieght must be at least 35 pounds.  Please enter a value within limits.", vbOKOnly, "Fuel Weight Limit Exceeded")
        txtFOB.SetFocus
    End If
    intFOB = Val(txtFOB)
    
    intCargoP = intP1 + intP2 + intP3 + intP4 + intP5 + intP6
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
    If Val(txtFuel) > (Val(txtFOB) - 35) Then
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

Private Sub txtTaxiFuel_GotFocus()
    txtTaxiFuel.SelStart = 0
    txtTaxiFuel.SelLength = Len(txtTaxiFuel.Text)
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

    intTotal = int1 + int2 + int3 + int4 + int5 + int6
    intTotal1 = intP1 + intP2 + intP3 + intP4 + intP5 + intP6
    intCargoP = intP1 + intP2 + intP3 + intP4 + intP5 + intP6
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

    intTotal = int1 + int2 + int3 + int4 + int5 + int6
    intTotal1 = intP1 + intP2 + intP3 + intP4 + intP5 + intP6
    intCargoP = intP1 + intP2 + intP3 + intP4 + intP5 + intP6
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

    intTotal = int1 + int2 + int3 + int4 + int5 + int6
    intTotal1 = intP1 + intP2 + intP3 + intP4 + intP5 + intP6
    intCargoP = intP1 + intP2 + intP3 + intP4 + intP5 + intP6
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

    intTotal = int1 + int2 + int3 + int4 + int5 + int6
    intTotal1 = intP1 + intP2 + intP3 + intP4 + intP5 + intP6
    intCargoP = intP1 + intP2 + intP3 + intP4 + intP5 + intP6
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

    intTotal = int1 + int2 + int3 + int4 + int5 + int6
    intTotal1 = intP1 + intP2 + intP3 + intP4 + intP5 + intP6
    intCargoP = intP1 + intP2 + intP3 + intP4 + intP5 + intP6
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

    intTotal = int1 + int2 + int3 + int4 + int5 + int6
    intTotal1 = intP1 + intP2 + intP3 + intP4 + intP5 + intP6
    intCargoP = intP1 + intP2 + intP3 + intP4 + intP5 + intP6
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

