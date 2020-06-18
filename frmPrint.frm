VERSION 5.00
Begin VB.Form frmPrint 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Weight and Balance"
   ClientHeight    =   7095
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11805
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7095
   ScaleWidth      =   11805
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   12480
      Left            =   -360
      ScaleHeight     =   12420
      ScaleWidth      =   15900
      TabIndex        =   0
      Top             =   -5760
      Width           =   15960
      Begin VB.TextBox Text3 
         Height          =   330
         Left            =   9480
         TabIndex        =   177
         Top             =   600
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Height          =   330
         Left            =   9480
         TabIndex        =   176
         Text            =   "Text2"
         Top             =   6960
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   330
         Left            =   9480
         TabIndex        =   175
         Text            =   "Text1"
         Top             =   6600
         Width           =   975
      End
      Begin VB.TextBox txtSIC 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   12000
         TabIndex        =   174
         Text            =   " "
         Top             =   7755
         Width           =   1335
      End
      Begin VB.TextBox Text58 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   340
         TabIndex        =   172
         Text            =   "Underload before LMC"
         Top             =   6900
         Width           =   1930
      End
      Begin VB.TextBox txtPrep 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   650
         Left            =   12000
         TabIndex        =   162
         Top             =   8400
         Width           =   1335
      End
      Begin VB.TextBox txtPIC1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   12000
         TabIndex        =   158
         Top             =   7440
         Width           =   1335
      End
      Begin VB.TextBox Text56 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   10680
         TabIndex        =   156
         Text            =   "Captain"
         Top             =   7440
         Width           =   1335
      End
      Begin VB.TextBox Text73 
         Appearance      =   0  'Flat
         Height          =   650
         Left            =   13320
         TabIndex        =   171
         Text            =   " "
         Top             =   10200
         Width           =   1335
      End
      Begin VB.TextBox Text72 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   650
         Left            =   12000
         TabIndex        =   170
         Text            =   " "
         Top             =   10200
         Width           =   1335
      End
      Begin VB.TextBox Text71 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   10680
         TabIndex        =   169
         Text            =   "Prepared by"
         Top             =   10515
         Width           =   1335
      End
      Begin VB.TextBox Text70 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   10680
         TabIndex        =   168
         Text            =   "(*) Load"
         Top             =   10200
         Width           =   1335
      End
      Begin VB.TextBox Text69 
         Appearance      =   0  'Flat
         Height          =   650
         Left            =   13320
         TabIndex        =   167
         Text            =   " "
         Top             =   9360
         Width           =   1335
      End
      Begin VB.TextBox Text68 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   650
         Left            =   12000
         TabIndex        =   166
         Text            =   " "
         Top             =   9360
         Width           =   1335
      End
      Begin VB.TextBox Text67 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   10680
         TabIndex        =   165
         Text            =   "Checked by"
         Top             =   9675
         Width           =   1335
      End
      Begin VB.TextBox Text66 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   10680
         TabIndex        =   164
         Text            =   "Load"
         Top             =   9360
         Width           =   1335
      End
      Begin VB.TextBox Text65 
         Appearance      =   0  'Flat
         Height          =   650
         Left            =   13320
         TabIndex        =   163
         Text            =   " "
         Top             =   8400
         Width           =   1335
      End
      Begin VB.TextBox Text63 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   10680
         TabIndex        =   161
         Text            =   "Prepared by"
         Top             =   8715
         Width           =   1335
      End
      Begin VB.TextBox Text62 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   10680
         TabIndex        =   160
         Text            =   "W & B"
         Top             =   8400
         Width           =   1335
      End
      Begin VB.TextBox Text61 
         Appearance      =   0  'Flat
         Height          =   650
         Left            =   13320
         TabIndex        =   159
         Text            =   " "
         Top             =   7440
         Width           =   1335
      End
      Begin VB.TextBox txtUnderLoad 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   7080
         TabIndex        =   155
         Text            =   " "
         Top             =   8370
         Width           =   1095
      End
      Begin VB.TextBox Text57 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   4920
         TabIndex        =   154
         Text            =   "Underload Before LMC"
         Top             =   8370
         Width           =   2175
      End
      Begin VB.TextBox txtTotalPay 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   7080
         TabIndex        =   153
         Text            =   " "
         Top             =   8055
         Width           =   1095
      End
      Begin VB.TextBox Text55 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   4920
         TabIndex        =   152
         Text            =   "Total Payload"
         Top             =   8055
         Width           =   2175
      End
      Begin VB.TextBox Text54 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   7080
         TabIndex        =   151
         Text            =   "1542.00"
         Top             =   7740
         Width           =   1095
      End
      Begin VB.TextBox Text53 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   4920
         TabIndex        =   150
         Text            =   "Allowed Payload"
         Top             =   7740
         Width           =   2175
      End
      Begin VB.TextBox Text52 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   4920
         TabIndex        =   149
         Text            =   "All these weights in Kilograms"
         Top             =   7425
         Width           =   3255
      End
      Begin VB.TextBox Text51 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   4920
         MultiLine       =   -1  'True
         TabIndex        =   148
         Top             =   10680
         Width           =   5415
      End
      Begin VB.TextBox Text50 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7800
         TabIndex        =   147
         Text            =   "Weight"
         Top             =   6320
         Width           =   1575
      End
      Begin VB.TextBox Text49 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6360
         TabIndex        =   146
         Text            =   "+/-"
         Top             =   6320
         Width           =   1455
      End
      Begin VB.TextBox Text48 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4920
         TabIndex        =   145
         Text            =   "Comp"
         Top             =   6320
         Width           =   1455
      End
      Begin VB.TextBox Text47 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3480
         TabIndex        =   144
         Text            =   "Destination"
         Top             =   6320
         Width           =   1455
      End
      Begin VB.TextBox Text46 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3480
         TabIndex        =   143
         Text            =   "Last Minute Change"
         Top             =   6015
         Width           =   5895
      End
      Begin VB.TextBox Text45 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3670
         TabIndex        =   142
         Text            =   "Center of Gravity Calculations"
         Top             =   1455
         Width           =   5675
      End
      Begin VB.TextBox Text44 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3680
         TabIndex        =   141
         Text            =   "TO Wt"
         Top             =   5550
         Width           =   1005
      End
      Begin VB.TextBox Text43 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3680
         TabIndex        =   140
         Text            =   "TO Fuel"
         Top             =   5235
         Width           =   1005
      End
      Begin VB.TextBox Text42 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3680
         TabIndex        =   139
         Text            =   "ZFW"
         Top             =   4920
         Width           =   1005
      End
      Begin VB.TextBox Text41 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3680
         TabIndex        =   138
         Text            =   "Hold 6"
         Top             =   4605
         Width           =   1005
      End
      Begin VB.TextBox Text40 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3680
         TabIndex        =   137
         Text            =   "Hold 5"
         Top             =   4290
         Width           =   1005
      End
      Begin VB.TextBox Text36 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3680
         TabIndex        =   136
         Text            =   "Hold 4"
         Top             =   3975
         Width           =   1005
      End
      Begin VB.TextBox Text35 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3680
         TabIndex        =   135
         Text            =   "Hold 3"
         Top             =   3660
         Width           =   1005
      End
      Begin VB.TextBox Text34 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3680
         TabIndex        =   134
         Text            =   "Hold 2"
         Top             =   3345
         Width           =   1005
      End
      Begin VB.TextBox Text33 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3680
         TabIndex        =   133
         Text            =   "Hold 1"
         Top             =   3030
         Width           =   1005
      End
      Begin VB.TextBox Text32 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3680
         TabIndex        =   132
         Text            =   "Extra Crew"
         Top             =   2715
         Width           =   1005
      End
      Begin VB.TextBox Text31 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3680
         TabIndex        =   131
         Text            =   "Crew"
         Top             =   2400
         Width           =   1005
      End
      Begin VB.TextBox Text30 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3680
         TabIndex        =   130
         Text            =   "BEW"
         Top             =   2085
         Width           =   1005
      End
      Begin VB.TextBox Text29 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   8160
         TabIndex        =   129
         Text            =   "Max. Load"
         Top             =   1770
         Width           =   1185
      End
      Begin VB.TextBox Text28 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   6990
         TabIndex        =   128
         Text            =   "Moment"
         Top             =   1770
         Width           =   1185
      End
      Begin VB.TextBox Text27 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   5820
         TabIndex        =   127
         Text            =   "ARM"
         Top             =   1770
         Width           =   1185
      End
      Begin VB.TextBox Text26 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   4650
         TabIndex        =   126
         Text            =   "Weight"
         Top             =   1770
         Width           =   1185
      End
      Begin VB.TextBox Text25 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   3680
         TabIndex        =   125
         Top             =   1770
         Width           =   1005
      End
      Begin VB.TextBox Text24 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2220
         TabIndex        =   124
         Text            =   "Max Wt lbs"
         Top             =   2820
         Width           =   950
      End
      Begin VB.TextBox Text23 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1340
         TabIndex        =   123
         Text            =   "Weight"
         Top             =   2820
         Width           =   950
      End
      Begin VB.TextBox Text22 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Height          =   330
         Left            =   340
         TabIndex        =   122
         Text            =   " "
         Top             =   2820
         Width           =   1010
      End
      Begin VB.TextBox Text21 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   340
         TabIndex        =   121
         Text            =   "Gross Weight Computation"
         Top             =   2520
         Width           =   2815
      End
      Begin VB.TextBox Text20 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   340
         TabIndex        =   120
         Text            =   "LW"
         Top             =   6585
         Width           =   1010
      End
      Begin VB.TextBox Text19 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   340
         TabIndex        =   119
         Text            =   "Trip Fuel"
         Top             =   6270
         Width           =   1010
      End
      Begin VB.TextBox Text18 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   340
         TabIndex        =   118
         Text            =   "TOW"
         Top             =   5955
         Width           =   1010
      End
      Begin VB.TextBox Text17 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   340
         TabIndex        =   117
         Text            =   "Taxi Fuel"
         Top             =   5640
         Width           =   1010
      End
      Begin VB.TextBox Text16 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   340
         TabIndex        =   116
         Text            =   "Ramp Wt"
         Top             =   5325
         Width           =   1010
      End
      Begin VB.TextBox Text15 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   340
         TabIndex        =   115
         Text            =   "Ramp Fuel"
         Top             =   5010
         Width           =   1010
      End
      Begin VB.TextBox Text14 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   340
         TabIndex        =   114
         Text            =   "ZFW"
         Top             =   4695
         Width           =   1010
      End
      Begin VB.TextBox Text13 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   340
         TabIndex        =   113
         Text            =   "Payload"
         Top             =   4380
         Width           =   1010
      End
      Begin VB.TextBox Text12 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   340
         TabIndex        =   112
         Text            =   "BOW"
         Top             =   4065
         Width           =   1010
      End
      Begin VB.TextBox Text11 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   340
         TabIndex        =   111
         Text            =   "Extra Crew"
         Top             =   3750
         Width           =   1010
      End
      Begin VB.TextBox Text10 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   340
         TabIndex        =   110
         Text            =   "Crew"
         Top             =   3435
         Width           =   1010
      End
      Begin VB.TextBox Text9 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   340
         TabIndex        =   109
         Text            =   "BEW"
         Top             =   3120
         Width           =   1010
      End
      Begin VB.TextBox txtLMCWt2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7800
         TabIndex        =   107
         Text            =   " "
         Top             =   6920
         Width           =   1575
      End
      Begin VB.TextBox txtAS2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6360
         TabIndex        =   106
         Text            =   " "
         Top             =   6920
         Width           =   1455
      End
      Begin VB.TextBox txtHold2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4920
         TabIndex        =   105
         Text            =   " "
         Top             =   6920
         Width           =   1455
      End
      Begin VB.TextBox txtDest2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3480
         TabIndex        =   104
         Text            =   " "
         Top             =   6920
         Width           =   1455
      End
      Begin VB.TextBox txtLMCWt1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7800
         TabIndex        =   103
         Text            =   " "
         Top             =   6615
         Width           =   1575
      End
      Begin VB.TextBox txtAS1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6360
         TabIndex        =   102
         Text            =   " "
         Top             =   6615
         Width           =   1455
      End
      Begin VB.TextBox txtHold1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4920
         TabIndex        =   101
         Text            =   " "
         Top             =   6615
         Width           =   1455
      End
      Begin VB.TextBox txtDest1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3480
         TabIndex        =   100
         Text            =   " "
         Top             =   6615
         Width           =   1455
      End
      Begin VB.TextBox txtMax6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   8160
         TabIndex        =   99
         Text            =   "145 kgs"
         Top             =   4605
         Width           =   1185
      End
      Begin VB.TextBox txtMax9 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   8160
         TabIndex        =   98
         Text            =   "8750 lbs"
         Top             =   5550
         Width           =   1185
      End
      Begin VB.TextBox txtMax8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   8160
         TabIndex        =   97
         Text            =   "2224 lbs"
         Top             =   5235
         Width           =   1185
      End
      Begin VB.TextBox txtMax7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Height          =   330
         Left            =   8160
         TabIndex        =   96
         Text            =   " "
         Top             =   4920
         Width           =   1185
      End
      Begin VB.TextBox txtMax5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   8160
         TabIndex        =   95
         Text            =   "575 kgs"
         Top             =   4290
         Width           =   1185
      End
      Begin VB.TextBox txtMax4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   8160
         TabIndex        =   94
         Text            =   "626 kgs"
         Top             =   3975
         Width           =   1185
      End
      Begin VB.TextBox txtMax3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   8160
         TabIndex        =   93
         Text            =   "862 kgs"
         Top             =   3660
         Width           =   1185
      End
      Begin VB.TextBox txtMax2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   8160
         TabIndex        =   92
         Text            =   "1406 kgs"
         Top             =   3345
         Width           =   1185
      End
      Begin VB.TextBox txtMax1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   8160
         TabIndex        =   91
         Text            =   "807 kgs"
         Top             =   3030
         Width           =   1185
      End
      Begin VB.TextBox Text39 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Height          =   330
         Left            =   8160
         TabIndex        =   90
         Text            =   " "
         Top             =   2715
         Width           =   1185
      End
      Begin VB.TextBox Text38 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Height          =   330
         Left            =   8160
         TabIndex        =   89
         Text            =   " "
         Top             =   2400
         Width           =   1185
      End
      Begin VB.TextBox Text37 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Height          =   330
         Left            =   8160
         TabIndex        =   88
         Text            =   " "
         Top             =   2085
         Width           =   1185
      End
      Begin VB.TextBox txtTOWM1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   6990
         TabIndex        =   87
         Text            =   " "
         Top             =   5550
         Width           =   1185
      End
      Begin VB.TextBox txtFuelM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   6990
         TabIndex        =   86
         Text            =   " "
         Top             =   5235
         Width           =   1185
      End
      Begin VB.TextBox txtZEWM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   6990
         TabIndex        =   85
         Text            =   " "
         Top             =   4920
         Width           =   1185
      End
      Begin VB.TextBox txtM6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   6990
         TabIndex        =   84
         Text            =   " "
         Top             =   4605
         Width           =   1185
      End
      Begin VB.TextBox txtM5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   6990
         TabIndex        =   83
         Text            =   " "
         Top             =   4290
         Width           =   1185
      End
      Begin VB.TextBox txtM4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   6990
         TabIndex        =   82
         Text            =   " "
         Top             =   3975
         Width           =   1185
      End
      Begin VB.TextBox txtM3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   6990
         TabIndex        =   81
         Text            =   " "
         Top             =   3660
         Width           =   1185
      End
      Begin VB.TextBox txtM2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   6990
         TabIndex        =   80
         Text            =   " "
         Top             =   3345
         Width           =   1185
      End
      Begin VB.TextBox txtM1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   6990
         TabIndex        =   79
         Text            =   " "
         Top             =   3030
         Width           =   1185
      End
      Begin VB.TextBox txtXtraM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   6990
         TabIndex        =   78
         Text            =   " "
         Top             =   2715
         Width           =   1185
      End
      Begin VB.TextBox txtCrewM1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   6990
         TabIndex        =   77
         Text            =   " "
         Top             =   2400
         Width           =   1185
      End
      Begin VB.TextBox txtBEWM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   6990
         TabIndex        =   76
         Text            =   " "
         Top             =   2085
         Width           =   1185
      End
      Begin VB.TextBox txtTOWA 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   5820
         TabIndex        =   75
         Text            =   " "
         Top             =   5550
         Width           =   1185
      End
      Begin VB.TextBox txtFuelA 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   5820
         TabIndex        =   74
         Text            =   " "
         Top             =   5235
         Width           =   1185
      End
      Begin VB.TextBox txtZEWA 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   5820
         TabIndex        =   73
         Text            =   " "
         Top             =   4920
         Width           =   1185
      End
      Begin VB.TextBox txtA6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   5820
         TabIndex        =   72
         Text            =   " "
         Top             =   4605
         Width           =   1185
      End
      Begin VB.TextBox txtA5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   5820
         TabIndex        =   71
         Text            =   " "
         Top             =   4290
         Width           =   1185
      End
      Begin VB.TextBox txtA4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   5820
         TabIndex        =   70
         Text            =   " "
         Top             =   3975
         Width           =   1185
      End
      Begin VB.TextBox txtA3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   5820
         TabIndex        =   69
         Text            =   " "
         Top             =   3660
         Width           =   1185
      End
      Begin VB.TextBox txtA2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   5820
         TabIndex        =   68
         Text            =   " "
         Top             =   3345
         Width           =   1185
      End
      Begin VB.TextBox txtA1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   5820
         TabIndex        =   67
         Text            =   " "
         Top             =   3030
         Width           =   1185
      End
      Begin VB.TextBox txtXtraA 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   5820
         TabIndex        =   66
         Text            =   " "
         Top             =   2715
         Width           =   1185
      End
      Begin VB.TextBox txtCrewA 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   5820
         TabIndex        =   65
         Text            =   " "
         Top             =   2400
         Width           =   1185
      End
      Begin VB.TextBox txtBEWA 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   5820
         TabIndex        =   64
         Text            =   " "
         Top             =   2085
         Width           =   1185
      End
      Begin VB.TextBox txtTOWW1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   4650
         TabIndex        =   63
         Text            =   " "
         Top             =   5550
         Width           =   1185
      End
      Begin VB.TextBox txtFuelW 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   4650
         TabIndex        =   62
         Text            =   " "
         Top             =   5235
         Width           =   1185
      End
      Begin VB.TextBox txtZEWW1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   4650
         TabIndex        =   61
         Text            =   " "
         Top             =   4920
         Width           =   1185
      End
      Begin VB.TextBox txtW6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   4650
         TabIndex        =   60
         Text            =   " "
         Top             =   4605
         Width           =   1185
      End
      Begin VB.TextBox txtW5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   4650
         TabIndex        =   59
         Text            =   " "
         Top             =   4290
         Width           =   1185
      End
      Begin VB.TextBox txtW4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   4650
         TabIndex        =   58
         Text            =   " "
         Top             =   3975
         Width           =   1185
      End
      Begin VB.TextBox txtW3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   4650
         TabIndex        =   57
         Text            =   " "
         Top             =   3660
         Width           =   1185
      End
      Begin VB.TextBox txtW2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   4650
         TabIndex        =   56
         Text            =   " "
         Top             =   3345
         Width           =   1185
      End
      Begin VB.TextBox txtW1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   4650
         TabIndex        =   55
         Text            =   " "
         Top             =   3030
         Width           =   1185
      End
      Begin VB.TextBox txtXtraW 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   4650
         TabIndex        =   54
         Top             =   2715
         Width           =   1185
      End
      Begin VB.TextBox txtCrewW1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   4650
         TabIndex        =   53
         Top             =   2400
         Width           =   1185
      End
      Begin VB.TextBox txtBEWW 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   4650
         TabIndex        =   52
         Text            =   " "
         Top             =   2085
         Width           =   1185
      End
      Begin VB.TextBox txtLWM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2220
         TabIndex        =   51
         Text            =   "8500"
         Top             =   6585
         Width           =   950
      End
      Begin VB.TextBox txtTripM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   2220
         TabIndex        =   50
         Top             =   6270
         Width           =   950
      End
      Begin VB.TextBox txtTOWM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2220
         TabIndex        =   49
         Text            =   "8750"
         Top             =   5955
         Width           =   950
      End
      Begin VB.TextBox txtTaxiM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   2220
         TabIndex        =   48
         Top             =   5640
         Width           =   950
      End
      Begin VB.TextBox txtRampWM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2220
         TabIndex        =   47
         Text            =   "8785"
         Top             =   5325
         Width           =   950
      End
      Begin VB.TextBox txtRampFuelM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2220
         TabIndex        =   46
         Text            =   "2224"
         Top             =   5010
         Width           =   950
      End
      Begin VB.TextBox txtZFWM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2220
         TabIndex        =   45
         Top             =   4695
         Width           =   950
      End
      Begin VB.TextBox txtPayloadM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2220
         TabIndex        =   44
         Text            =   "3129"
         Top             =   4380
         Width           =   950
      End
      Begin VB.TextBox txtBOWM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   2220
         TabIndex        =   43
         Top             =   4065
         Width           =   950
      End
      Begin VB.TextBox txtXCrewM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2220
         TabIndex        =   42
         Top             =   3750
         Width           =   950
      End
      Begin VB.TextBox txtCrewM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2220
         TabIndex        =   41
         Top             =   3435
         Width           =   950
      End
      Begin VB.TextBox txtBWM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2220
         TabIndex        =   40
         Top             =   3120
         Width           =   950
      End
      Begin VB.TextBox txtTOWW 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1340
         TabIndex        =   39
         Top             =   5955
         Width           =   950
      End
      Begin VB.TextBox txtLWW 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1340
         TabIndex        =   38
         Top             =   6585
         Width           =   950
      End
      Begin VB.TextBox txtTripW 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   1340
         TabIndex        =   37
         Top             =   6270
         Width           =   950
      End
      Begin VB.TextBox txtTaxiW 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   1340
         TabIndex        =   36
         Top             =   5640
         Width           =   950
      End
      Begin VB.TextBox txtRampWW 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1340
         TabIndex        =   35
         Top             =   5325
         Width           =   950
      End
      Begin VB.TextBox txtRampFuelW 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   1340
         TabIndex        =   34
         Top             =   5010
         Width           =   950
      End
      Begin VB.TextBox txtZFWW 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1340
         TabIndex        =   33
         Top             =   4695
         Width           =   950
      End
      Begin VB.TextBox txtPayloadW 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1340
         TabIndex        =   32
         Top             =   4380
         Width           =   950
      End
      Begin VB.TextBox txtBOWW 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1340
         TabIndex        =   31
         Top             =   4065
         Width           =   950
      End
      Begin VB.TextBox txtXCrewW 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   1340
         TabIndex        =   30
         Top             =   3750
         Width           =   950
      End
      Begin VB.TextBox txtCrewW 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1340
         TabIndex        =   29
         Top             =   3435
         Width           =   950
      End
      Begin VB.TextBox txtBWW 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1340
         TabIndex        =   28
         Top             =   3120
         Width           =   950
      End
      Begin VB.TextBox txtReg1 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   1320
         TabIndex        =   27
         Top             =   1920
         Width           =   1815
      End
      Begin VB.TextBox txtDate 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1320
         TabIndex        =   26
         Top             =   1440
         Width           =   1815
      End
      Begin VB.TextBox txtDest 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   6600
         TabIndex        =   23
         Top             =   840
         Width           =   2800
      End
      Begin VB.TextBox txtFltNo 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   3480
         TabIndex        =   22
         Top             =   840
         Width           =   2800
      End
      Begin VB.TextBox txtOrig 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   360
         TabIndex        =   21
         Top             =   840
         Width           =   2800
      End
      Begin VB.TextBox txtReg 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7560
         TabIndex        =   16
         Top             =   9615
         Width           =   1095
      End
      Begin VB.TextBox txtZone6 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6705
         TabIndex        =   15
         Top             =   9990
         Width           =   400
      End
      Begin VB.TextBox txtZone5 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6230
         TabIndex        =   14
         Top             =   9990
         Width           =   415
      End
      Begin VB.TextBox txtZone4 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5760
         TabIndex        =   13
         Text            =   "22.55"
         Top             =   9990
         Width           =   400
      End
      Begin VB.TextBox txtZone3 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5190
         TabIndex        =   12
         Top             =   9990
         Width           =   415
      End
      Begin VB.TextBox txtZone2 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4320
         TabIndex        =   11
         Top             =   9990
         Width           =   415
      End
      Begin VB.TextBox txtZone1 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3450
         TabIndex        =   10
         Top             =   9990
         Width           =   415
      End
      Begin VB.TextBox txtPIC 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1800
         TabIndex        =   3
         Top             =   9840
         Width           =   1335
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   13680
         TabIndex        =   2
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox Text59 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   10680
         TabIndex        =   157
         Text            =   "Co-Pilot"
         Top             =   7755
         Width           =   1335
      End
      Begin VB.TextBox txtUnderLoad1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2220
         TabIndex        =   173
         Top             =   6900
         Width           =   950
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   14680
         X2              =   11880
         Y1              =   3840
         Y2              =   5160
      End
      Begin VB.Label Label52 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "*Enter the Holds Freight in Kgs. Rest of the datums in lbs."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   108
         Top             =   7560
         Width           =   4695
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFFFFF&
         Caption         =   "AC Reg"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   25
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   24
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Destination"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6600
         TabIndex        =   20
         Top             =   600
         Width           =   2805
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Flight Number"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3480
         TabIndex        =   19
         Top             =   600
         Width           =   2805
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Origin"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   18
         Top             =   600
         Width           =   2805
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CESSNA CARAVAN C208B WEIGHT and BALANCE / LOADPLAN"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   17
         Top             =   0
         Width           =   14775
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "seats 9/10"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   6765
         TabIndex        =   9
         Top             =   9405
         Width           =   315
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "E5"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   6300
         TabIndex        =   8
         Top             =   9405
         Width           =   315
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "E4"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   5835
         TabIndex        =   7
         Top             =   9405
         Width           =   300
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "E3"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   5310
         TabIndex        =   6
         Top             =   9405
         Width           =   195
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "E2"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   4200
         TabIndex        =   5
         Top             =   9405
         Width           =   675
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Seats 1/2"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   3405
         TabIndex        =   4
         Top             =   9405
         Width           =   525
      End
      Begin VB.Image Image2 
         Height          =   4845
         Left            =   360
         Picture         =   "frmPrint.frx":0000
         Stretch         =   -1  'True
         Top             =   7320
         Width           =   10080
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   5775
         Left            =   9840
         Picture         =   "frmPrint.frx":9EC9
         Top             =   1080
         Width           =   5640
      End
   End
   Begin VB.PictureBox Picture2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      ScaleHeight     =   315
      ScaleWidth      =   675
      TabIndex        =   1
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim intAP As Integer, intUnder As Integer
 Private Const twipFactor = 1440
    Private Const WM_PAINT = &HF
    Private Const WM_PRINT = &H317
    Private Const PRF_CLIENT = &H4&    ' Draw the window's client area.
    Private Const PRF_CHILDREN = &H10& ' Draw all visible child windows.
    Private Const PRF_OWNED = &H20&    ' Draw all owned windows.

    Private Declare Function SendMessage Lib "user32" Alias _
        "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
        ByVal wParam As Long, ByVal lParam As Long) As Long

Private Sub cmdPrint_Click()

    Dim sWide As Single, sTall As Single
    Dim rv As Long
    Me.ScaleMode = vbTwips   ' default
    sWide = 11
    sTall = 8.5
    Me.Width = twipFactor * sWide
    Me.Height = twipFactor * sTall
    cmdPrint.Visible = False
    
    With Picture1
        .Top = 0
        .Left = 0
        .Width = twipFactor * sWide
        .Height = twipFactor * sTall
    End With
    
    With Picture2
        .Top = 0
        .Left = 0
        .Width = twipFactor * sWide
        .Height = twipFactor * sTall
    End With
   
        Me.Visible = True
        DoEvents
        Picture2.Visible = False
        Picture1.SetFocus
        Picture2.AutoRedraw = True
        rv = SendMessage(Picture1.hwnd, WM_PAINT, Picture2.hDC, 0)
        rv = SendMessage(Picture1.hwnd, WM_PRINT, Picture2.hDC, _
        PRF_CHILDREN + PRF_CLIENT + PRF_OWNED)
        Picture2.Picture = Picture2.Image
        Picture2.AutoRedraw = False
        
        'Printer.PrintQuality 300
        Printer.Orientation = vbPRORLandscape
        Printer.Print ""
        Printer.PaintPicture Picture2.Picture, 0, 0
   
        Printer.EndDoc
        frmMain.Show
    'End
    
End Sub

Private Sub Form_Activate()
    txtDate = Date
    txtPIC = frmMain.txtPIC
    txtSIC = frmMain.txtSIC
    txtPIC1 = frmMain.txtPIC
    Text33 = frm2.Text33
    Text34 = frm2.Text34
    Text35 = frm2.Text35
    Text36 = frm2.Text36
    Text40 = frm2.Text40
    Text41 = frm2.Text41
    Label1 = frm2.Text33
    Label2 = frm2.Text34
    Label3 = frm2.Text35
    Label4 = frm2.Text36
    Label5 = frm2.Text40
    Label6 = frm2.Text41
    txtZone1 = frm2.txtZone1
    txtZone1 = Format(txtZone1.Text, "Fixed")
    txtZone2 = frm2.txtZone2
    txtZone2 = Format(txtZone2.Text, "Fixed")
    txtZone3 = frm2.txtZone3
    txtZone3 = Format(txtZone3.Text, "Fixed")
    txtZone4 = frm2.txtZone4
    txtZone4 = Format(txtZone4.Text, "Fixed")
    txtZone5 = frm2.txtZone5
    txtZone5 = Format(txtZone5.Text, "Fixed")
    txtZone6 = frm2.txtZone6
    txtZone6 = Format(txtZone6.Text, "Fixed")
    txtW1 = frm2.txtW1
    txtW1 = Format(txtW1.Text, "Fixed")
    txtW2 = frm2.txtW2
    txtW2 = Format(txtW2.Text, "Fixed")
    txtW3 = frm2.txtW3
    txtW3 = Format(txtW3.Text, "Fixed")
    txtW4 = frm2.txtW4
    txtW4 = Format(txtW4.Text, "Fixed")
    txtW5 = frm2.txtW5
    txtW5 = Format(txtW5.Text, "Fixed")
    txtW6 = frm2.txtW6
    txtW6 = Format(txtW6.Text, "Fixed")
    txtReg = frmMain.cmbAC.Text
    txtReg1 = frmMain.cmbAC.Text
    txtOrig = frmMain.txtOrig.Text
    txtDest = frmMain.txtDest.Text
    txtFltNo = frmMain.txtFltNo.Text
    txtPayloadW = frm2.txtPayloadW
    txtPayloadW = Format(txtPayloadW.Text, "Fixed")
    txtTotalPay = frm2.txtTotalPay
    txtTotalPay = Format(txtTotalPay.Text, "Fixed")
    txtFuelW = frm2.txtFuelW
    txtFuelW = Format(txtFuelW.Text, "Fixed")
    txtTripW = frm2.txtTripW
    txtTripW = Format(txtTripW.Text, "Fixed")
    txtBWW = frm2.txtBWW
    txtBWW = Format(txtBWW.Text, "Fixed")
    txtBEWW = frm2.txtBEWW
    txtBEWW = Format(txtBEWW.Text, "Fixed")
    txtBOWW = frm2.txtBOWW
    txtBOWW = Format(txtBOWW.Text, "Fixed")
    txtZFWW = frm2.txtZFWW
    txtZFWW = Format(txtZFWW.Text, "Fixed")
    txtRampFuelW = frm2.txtRampFuelW
    txtRampFuelW = Format(txtRampFuelW.Text, "Fixed")
    txtZEWW1 = frm2.txtZEWW1
    txtZEWW1 = Format(txtZEWW1.Text, "Fixed")
    txtRampWW = frm2.txtRampWW
    txtRampWW = Format(txtRampWW.Text, "Fixed")
    txtTaxiW = frm2.txtTaxiW
    txtTaxiW = Format(txtTaxiW.Text, "Fixed")
    txtTOWW = frm2.txtTOWW
    txtTOWW = Format(txtTOWW.Text, "Fixed")
    txtTOWW1 = frm2.txtTOWW1
    txtTOWW1 = Format(txtTOWW1.Text, "Fixed")
    txtTripW = frm2.txtTripW
    txtTripW = Format(txtTripW.Text, "Fixed")
    txtLWW = frm2.txtLWW
    txtLWW = Format(txtLWW.Text, "Fixed")
    txtUnderLoad = frm2.txtUnderLoad
    txtUnderLoad1 = frm2.txtUnderLoad1
    txtBEWA = frm2.txtBEWA
    txtBEWA = Format(txtBEWA.Text, "Fixed")
    txtCrewA = frm2.txtCrewA
    txtCrewA = Format(txtCrewA.Text, "Fixed")
    txtXtraA = frm2.txtXtraA
    txtXtraA = Format(txtXtraA.Text, "Fixed")
    txtA1 = frm2.txtA1
    txtA1 = Format(txtA1.Text, "Fixed")
    txtA2 = frm2.txtA2
    txtA2 = Format(txtA2.Text, "Fixed")
    txtA3 = frm2.txtA3
    txtA3 = Format(txtA3.Text, "Fixed")
    txtA4 = frm2.txtA4
    txtA4 = Format(txtA4.Text, "Fixed")
    txtA5 = frm2.txtA5
    txtA5 = Format(txtA5.Text, "Fixed")
    txtA6 = frm2.txtA6
    txtA6 = Format(txtA6.Text, "Fixed")
    txtZEWA = frm2.txtZEWA
    txtZEWA = Format(txtZEWA.Text, "Fixed")
    txtFuelA = frm2.txtFuelA
    txtFuelA = Format(txtFuelA.Text, "Fixed")
    txtTOWA = frm2.txtTOWA
    txtTOWA = Format(txtTOWA.Text, "Fixed")
    txtBEWM = frm2.txtBEWM
    txtBEWM = Format(txtBEWM.Text, "Fixed")
    txtCrewM1 = frm2.txtCrewM1
    txtCrewM1 = Format(txtCrewM1.Text, "Fixed")
    txtXtraM = frm2.txtXtraM
    txtXtraM = Format(txtXtraM.Text, "Fixed")
    txtM1 = frm2.txtM1
    txtM1 = Format(txtM1.Text, "Fixed")
    txtM2 = frm2.txtM2
    txtM2 = Format(txtM2.Text, "Fixed")
    txtM3 = frm2.txtM3
    txtM3 = Format(txtM3.Text, "Fixed")
    txtM4 = frm2.txtM4
    txtM4 = Format(txtM4.Text, "Fixed")
    txtM5 = frm2.txtM5
    txtM5 = Format(txtM5.Text, "Fixed")
    txtM6 = frm2.txtM6
    txtM6 = Format(txtM6.Text, "Fixed")
    txtZEWM = frm2.txtZEWM
    txtZEWM = Format(txtZEWM.Text, "Fixed")
    txtFuelM = frm2.txtFuelM
    txtFuelM = Format(txtFuelM.Text, "Fixed")
    txtTOWM1 = frm2.txtTOWM1
    txtTOWM1 = Format(txtTOWM1.Text, "Fixed")
    txtDest1 = frm2.txtDest
    txtHold1 = frm2.txtComp1
    txtAS1 = frm2.txtAS1
    txtLMCWt1 = frm2.txtLMCWt1
    txtLMCWt2 = frm2.txtLMCWt2
    txtDest2 = frm2.txtDest
    txtHold2 = frm2.txtComp2
    txtAS2 = frm2.txtAS2
    txtLMCWt2 = frm2.txtLMCWt2
    txtCrewW = intWt2
    txtCrewW1 = intWt2
    txtXCrewW = intWt3
    txtXtraW = intWt3
    txtMax1 = frm2.txtMax1
    txtMax2 = frm2.txtMax2
    txtMax3 = frm2.txtMax3
    txtMax4 = frm2.txtMax4
    txtMax5 = frm2.txtMax5
    txtMax6 = frm2.txtMax6
    Text52 = frm2.Text52
    Text54 = frm2.Text54
    Text50 = frm2.Text50
    
    Text1 = intArm12
    Text2 = intLandArm
    'Text1 = Format(Text1.Text, "###")
    'intArm12 = Val(Text1)
    'Text1 = intArm12
    
    If frmMain.cmbAC = "208 Cargo Master" Then
        'Image2.Visible = False
        'Image3.Visible = True
    End If
    
    Text3 = intTOWt
    Text3 = Format(Text3, "####")
    intTOWt = Val(Text3)
    Select Case intTOWt
    
        Case 4800 To 4825
            Line1.Y1 = 5624
        Case 4826 To 4849
            Line1.Y1 = 5601
        Case 4850 To 4875
            Line1.Y1 = 5578
        Case 4876 To 4899
            Line1.Y1 = 5555
        Case 4900 To 4925
            Line1.Y1 = 5532
        Case 4926 To 4949
            Line1.Y1 = 5509
        Case 4950 To 4975
            Line1.Y1 = 5486
        Case 4976 To 4999
            Line1.Y1 = 5463
        Case 5000 To 5025
            Line1.Y1 = 5440
        Case 5026 To 5049
            Line1.Y1 = 5417
        Case 5050 To 5075
            Line1.Y1 = 5394
        Case 5076 To 5099
            Line1.Y1 = 5371
        Case 5100 To 5125
            Line1.Y1 = 5348
        Case 5126 To 5149
            Line1.Y1 = 5325
        Case 5150 To 5175
            Line1.Y1 = 5302
        Case 5176 To 5199
            Line1.Y1 = 5279
        Case 5200 To 5225
            Line1.Y1 = 5256
        Case 5226 To 5249
            Line1.Y1 = 5233
        Case 5250 To 5275
            Line1.Y1 = 5210
        Case 5276 To 5299
            Line1.Y1 = 5187
        Case 5300 To 5325
            Line1.Y1 = 5164
        Case 5326 To 5349
            Line1.Y1 = 5141
        Case 5350 To 5375
            Line1.Y1 = 5118
        Case 5376 To 5399
            Line1.Y1 = 5095
        Case 5400 To 5425
            Line1.Y1 = 5072
        Case 5426 To 5449
            Line1.Y1 = 5049
        Case 5450 To 5475
            Line1.Y1 = 5026
        Case 5476 To 5499
            Line1.Y1 = 5003
        Case 5500 To 5525
            Line1.Y1 = 4980
        Case 5526 To 5549
            Line1.Y1 = 4957
        Case 5550 To 5575
            Line1.Y1 = 4934
        Case 5576 To 5599
            Line1.Y1 = 4911
        Case 5600 To 5625
            Line1.Y1 = 4888
        Case 5626 To 5649
            Line1.Y1 = 4865
        Case 5650 To 5675
            Line1.Y1 = 4842
        Case 5676 To 5699
            Line1.Y1 = 4819
        Case 5700 To 5725
            Line1.Y1 = 4796
        Case 5726 To 5749
            Line1.Y1 = 4773
        Case 5750 To 5775
            Line1.Y1 = 4750
        Case 5776 To 5799
            Line1.Y1 = 4727
        Case 5800 To 5825
            Line1.Y1 = 4704
        Case 5826 To 5849
            Line1.Y1 = 4681
        Case 5850 To 5875
            Line1.Y1 = 4658
        Case 5876 To 5899
            Line1.Y1 = 4635
        Case 5900 To 5925
            Line1.Y1 = 4612
        Case 5926 To 5949
            Line1.Y1 = 4589
        Case 5950 To 5975
            Line1.Y1 = 4566
        Case 5976 To 5999
            Line1.Y1 = 4543
        Case 6000 To 6025
            Line1.Y1 = 4520
        Case 6026 To 6049
            Line1.Y1 = 4497
        Case 6050 To 6075
            Line1.Y1 = 4474
        Case 6076 To 6099
            Line1.Y1 = 4451
        Case 6100 To 6125
            Line1.Y1 = 4428
        Case 6126 To 6149
            Line1.Y1 = 4405
        Case 6150 To 6175
            Line1.Y1 = 4382
        Case 6176 To 6199
            Line1.Y1 = 4359
        Case 6200 To 6225
            Line1.Y1 = 4336
        Case 6226 To 6249
            Line1.Y1 = 4313
        Case 6250 To 6275
            Line1.Y1 = 4290
        Case 6276 To 6299
            Line1.Y1 = 4267
        Case 6300 To 6325
            Line1.Y1 = 4244
        Case 6326 To 6349
            Line1.Y1 = 4221
        Case 6350 To 6375
            Line1.Y1 = 4198
        Case 6376 To 6399
            Line1.Y1 = 4175
        Case 6400 To 6425
            Line1.Y1 = 4152
        Case 6426 To 6449
            Line1.Y1 = 4129
        Case 6450 To 6475
            Line1.Y1 = 4106
        Case 6476 To 6499
            Line1.Y1 = 4083
        Case 6500 To 6525
            Line1.Y1 = 4060
        Case 6526 To 6549
            Line1.Y1 = 4037
        Case 6550 To 6575
            Line1.Y1 = 4014
        Case 6576 To 6599
            Line1.Y1 = 3991
        Case 6600 To 6625
            Line1.Y1 = 3968
        Case 6626 To 6649
            Line1.Y1 = 3945
        Case 6650 To 6675
            Line1.Y1 = 3922
        Case 6676 To 6699
            Line1.Y1 = 3899
        Case 6700 To 6725
            Line1.Y1 = 3876
        Case 6726 To 6749
            Line1.Y1 = 3853
        Case 6750 To 6775
            Line1.Y1 = 3830
        Case 6776 To 6799
            Line1.Y1 = 3807
        Case 6800 To 6849
            Line1.Y1 = 3784
        Case 6800 To 6849
            Line1.Y1 = 3761
        Case 6850 To 6875
            Line1.Y1 = 3738
        Case 6876 To 6899
            Line1.Y1 = 3715
        Case 6900 To 6925
            Line1.Y1 = 3692
        Case 6926 To 6949
            Line1.Y1 = 3669
        Case 6950 To 6975
            Line1.Y1 = 3646
        Case 6976 To 6999
            Line1.Y1 = 3623
        Case 7000 To 7025
            Line1.Y1 = 3600
        Case 7026 To 7049
            Line1.Y1 = 3577
        Case 7050 To 7099
            Line1.Y1 = 3554
        Case 7050 To 7099
            Line1.Y1 = 3531
        Case 7100 To 7125
            Line1.Y1 = 3508
        Case 7126 To 7149
            Line1.Y1 = 3485
        Case 7150 To 7175
            Line1.Y1 = 3462
        Case 7176 To 7199
            Line1.Y1 = 3439
        Case 7200 To 7225
            Line1.Y1 = 3416
        Case 7226 To 7249
            Line1.Y1 = 3393
        Case 7250 To 7275
            Line1.Y1 = 3370
        Case 7276 To 7299
            Line1.Y1 = 3347
        Case 7300 To 7325
            Line1.Y1 = 3324
        Case 7326 To 7349
            Line1.Y1 = 3301
        Case 7350 To 7375
            Line1.Y1 = 3278
        Case 7376 To 7399
            Line1.Y1 = 3255
        Case 7400 To 7425
            Line1.Y1 = 3232
        Case 7426 To 7449
            Line1.Y1 = 3209
        Case 7450 To 7475
            Line1.Y1 = 3186
        Case 7476 To 7499
            Line1.Y1 = 3163
        Case 7500 To 7525
            Line1.Y1 = 3140
        Case 7526 To 7549
            Line1.Y1 = 3117
        Case 7550 To 7575
            Line1.Y1 = 3094
        Case 7576 To 7599
            Line1.Y1 = 3071
        Case 7600 To 7625
            Line1.Y1 = 3048
        Case 7626 To 7649
            Line1.Y1 = 3025
        Case 7650 To 7699
            Line1.Y1 = 3002
        Case 7650 To 7699
            Line1.Y1 = 2979
        Case 7700 To 7725
            Line1.Y1 = 2956
        Case 7726 To 7749
            Line1.Y1 = 2933
        Case 7750 To 7775
            Line1.Y1 = 2910
        Case 7776 To 7799
            Line1.Y1 = 2887
        Case 7800 To 7825
            Line1.Y1 = 2864
        Case 7826 To 7849
            Line1.Y1 = 2841
        Case 7850 To 7875
            Line1.Y1 = 2818
        Case 7876 To 7899
            Line1.Y1 = 2795
        Case 7900 To 7925
            Line1.Y1 = 2772
        Case 7926 To 7949
            Line1.Y1 = 2749
        Case 7950 To 7975
            Line1.Y1 = 2726
        Case 7976 To 7999
            Line1.Y1 = 2703
        Case 8000 To 8025
            Line1.Y1 = 2680
        Case 8026 To 8049
            Line1.Y1 = 2657
        Case 8050 To 8075
            Line1.Y1 = 2634
        Case 8076 To 8099
            Line1.Y1 = 2611
        Case 8100 To 8125
            Line1.Y1 = 2588
        Case 8126 To 8149
            Line1.Y1 = 2565
        Case 8150 To 8175
            Line1.Y1 = 2542
        Case 8176 To 8199
            Line1.Y1 = 2519
        Case 8200 To 8225
            Line1.Y1 = 2496
        Case 8226 To 8249
            Line1.Y1 = 2473
        Case 8250 To 8275
            Line1.Y1 = 2450
        Case 8276 To 8299
            Line1.Y1 = 2427
        Case 8300 To 8325
            Line1.Y1 = 2404
        Case 8326 To 8349
            Line1.Y1 = 2381
        Case 8350 To 8375
            Line1.Y1 = 2358
        Case 8376 To 8399
            Line1.Y1 = 2335
        Case 8400 To 8425
            Line1.Y1 = 2312
        Case 8426 To 8449
            Line1.Y1 = 2289
        Case 8450 To 8475
            Line1.Y1 = 2266
        Case 8476 To 8499
            Line1.Y1 = 2243
        Case 8500 To 8525
            Line1.Y1 = 2220
        Case 8526 To 8549
            Line1.Y1 = 2197
        Case 8550 To 8575
            Line1.Y1 = 2174
        Case 8576 To 8599
            Line1.Y1 = 2151
        Case 8600 To 8625
            Line1.Y1 = 2128
        Case 8626 To 8649
            Line1.Y1 = 2105
        Case 8650 To 8675
            Line1.Y1 = 2082
        Case 8676 To 8699
            Line1.Y1 = 2059
        Case 8700 To 8725
            Line1.Y1 = 2036
        Case 8726 To 8749
            Line1.Y1 = 2013
        Case 8750 To 8775
            Line1.Y1 = 1990
        Case 8776 To 8799
            Line1.Y1 = 1967
            
     End Select
     
     Select Case intArm12
     
        Case 175# To 175.25
            Line1.X1 = 10600
        Case 175.26 To 175.5
            Line1.X1 = 10634
        Case 175.51 To 175.75
            Line1.X1 = 10668
        Case 175.76 To 175.99
            Line1.X1 = 10702
        Case 176 To 176.25
            Line1.X1 = 10736
        Case 176.26 To 176.5
            Line1.X1 = 10770
        Case 176.51 To 176.75
            Line1.X1 = 10804
        Case 176.76 To 176.99
            Line1.X1 = 10738
        Case 177 To 177.25
            Line1.X1 = 10872
        Case 177.26 To 177.5
            Line1.X1 = 10906
        Case 177.51 To 177.75
            Line1.X1 = 10940
        Case 177.76 To 177.99
            Line1.X1 = 10974
        Case 178 To 178.25
            Line1.X1 = 11008
        Case 178.26 To 178.5
            Line1.X1 = 11042
        Case 178.51 To 178.75
            Line1.X1 = 11076
        Case 178.76 To 178.99
            Line1.X1 = 11110
        Case 179 To 179.25
            Line1.X1 = 11144
        Case 179.26 To 179.5
            Line1.X1 = 11178
        Case 179.51 To 179.75
            Line1.X1 = 11212
        Case 179.76 To 179.99
            Line1.X1 = 11246
        Case 180 To 180.25
            Line1.X1 = 11280
        Case 180.26 To 180.5
            Line1.X1 = 11314
        Case 180.51 To 180.75
            Line1.X1 = 11348
        Case 180.76 To 180.99
            Line1.X1 = 11382
        Case 181 To 181.25
            Line1.X1 = 11416
        Case 181.26 To 181.5
            Line1.X1 = 11450
        Case 181.51 To 181.75
            Line1.X1 = 11484
        Case 181.76 To 181.99
            Line1.X1 = 11518
        Case 182 To 182.25
            Line1.X1 = 11552
        Case 182.26 To 182.5
            Line1.X1 = 11586
        Case 182.51 To 182.75
            Line1.X1 = 11620
        Case 182.76 To 182.99
            Line1.X1 = 11654
        Case 183 To 183.25
            Line1.X1 = 11688
        Case 183.26 To 183.5
            Line1.X1 = 11722
        Case 183.51 To 183.75
            Line1.X1 = 11756
        Case 183.76 To 183.99
            Line1.X1 = 11790
        Case 184 To 184.25
            Line1.X1 = 11824
        Case 184.26 To 184.5
            Line1.X1 = 11858
        Case 184.51 To 184.75
            Line1.X1 = 11892
        Case 184.76 To 184.99
            Line1.X1 = 11926
        Case 185 To 185.25
            Line1.X1 = 11960
        Case 185.26 To 185.5
            Line1.X1 = 11994
        Case 185.51 To 185.75
            Line1.X1 = 12028
        Case 185.76 To 185.99
            Line1.X1 = 12062
        Case 186 To 186.25
            Line1.X1 = 12096
        Case 186.26 To 186.5
            Line1.X1 = 12130
        Case 186.51 To 186.75
            Line1.X1 = 12164
        Case 186.76 To 186.99
            Line1.X1 = 12198
        Case 187 To 187.25
            Line1.X1 = 12232
        Case 187.26 To 187.5
            Line1.X1 = 12266
        Case 187.51 To 187.75
            Line1.X1 = 12300
        Case 187.76 To 187.99
            Line1.X1 = 12334
        Case 188 To 188.25
            Line1.X1 = 12368
        Case 188.26 To 188.5
            Line1.X1 = 12402
        Case 188.51 To 188.75
            Line1.X1 = 12436
        Case 188.76 To 188.99
            Line1.X1 = 12470
        Case 189 To 189.25
            Line1.X1 = 12504
        Case 189.26 To 189.5
            Line1.X1 = 12538
        Case 189.51 To 189.75
            Line1.X1 = 12572
        Case 189.76 To 189.99
            Line1.X1 = 12606
        Case 190 To 190.25
            Line1.X1 = 12640
        Case 190.26 To 190.5
            Line1.X1 = 12674
        Case 190.51 To 190.75
            Line1.X1 = 12708
        Case 190.76 To 190.99
            Line1.X1 = 12742
        Case 191 To 191.25
            Line1.X1 = 12776
        Case 191.26 To 191.5
            Line1.X1 = 12810
        Case 191.51 To 191.75
            Line1.X1 = 12844
        Case 191.76 To 191.99
            Line1.X1 = 12878
        Case 192 To 192.25
            Line1.X1 = 12912
        Case 192.26 To 192.5
            Line1.X1 = 12946
        Case 192.51 To 192.75
            Line1.X1 = 12980
        Case 192.76 To 192.99
            Line1.X1 = 13014
        Case 193 To 193.25
            Line1.X1 = 13048
        Case 193.26 To 193.5
            Line1.X1 = 13082
        Case 193.51 To 193.75
            Line1.X1 = 13116
        Case 193.76 To 193.99
            Line1.X1 = 13150
        Case 194 To 194.25
            Line1.X1 = 13184
        Case 194.26 To 194.5
            Line1.X1 = 13218
        Case 194.51 To 194.75
            Line1.X1 = 13252
        Case 194.76 To 194.99
            Line1.X1 = 13286
        Case 195 To 195.25
            Line1.X1 = 13320
        Case 195.26 To 195.5
            Line1.X1 = 13354
        Case 195.51 To 195.75
            Line1.X1 = 13388
        Case 195.76 To 195.99
            Line1.X1 = 13422
        Case 196 To 196.25
            Line1.X1 = 13456
        Case 196.26 To 196.5
            Line1.X1 = 13490
        Case 196.51 To 196.75
            Line1.X1 = 13524
        Case 196.76 To 196.99
            Line1.X1 = 13558
        Case 197 To 197.25
            Line1.X1 = 13592
        Case 197.26 To 197.5
            Line1.X1 = 13626
        Case 197.51 To 197.75
            Line1.X1 = 13660
        Case 197.76 To 197.99
            Line1.X1 = 13694
        Case 198 To 198.25
            Line1.X1 = 13728
        Case 198.26 To 198.5
            Line1.X1 = 13762
        Case 198.51 To 198.75
            Line1.X1 = 13796
        Case 198.76 To 198.99
            Line1.X1 = 13830
        Case 199 To 199.25
            Line1.X1 = 13864
        Case 199.26 To 199.5
            Line1.X1 = 13898
        Case 199.51 To 199.75
            Line1.X1 = 13932
        Case 199.76 To 199.99
            Line1.X1 = 13966
        Case 200 To 200.25
            Line1.X1 = 14000
        Case 200.26 To 200.5
            Line1.X1 = 14034
        Case 200.51 To 200.75
            Line1.X1 = 14068
        Case 200.76 To 200.99
            Line1.X1 = 14102
        Case 201 To 201.25
            Line1.X1 = 14136
        Case 201.26 To 201.5
            Line1.X1 = 14170
        Case 201.51 To 201.75
            Line1.X1 = 14204
        Case 201.76 To 201.99
            Line1.X1 = 14238
        Case 202 To 202.25
            Line1.X1 = 14272
        Case 202.26 To 202.5
            Line1.X1 = 14306
        Case 202.51 To 202.75
            Line1.X1 = 14340
        Case 202.76 To 202.99
            Line1.X1 = 14374
        Case 203 To 203.25
            Line1.X1 = 14408
        Case 203.26 To 203.5
            Line1.X1 = 14442
        Case 203.51 To 203.75
            Line1.X1 = 14476
        Case 203.76 To 203.99
            Line1.X1 = 14510
        Case 204 To 204.25
            Line1.X1 = 14544
        Case 204.26 To 204.5
            Line1.X1 = 14578
        Case 204.51 To 204.75
            Line1.X1 = 14612
        Case 204.76 To 204.99
            Line1.X1 = 14646
        Case 205 To 205.25
            Line1.X1 = 14680
        Case 205.26 To 204.5
            Line1.X1 = 14714
        Case 205.51 To 205.75
            Line1.X1 = 14748
        Case 205.76 To 205.99
            Line1.X1 = 14782
            
    End Select
    
    Text3 = intLandWt
    Text3 = Format(Text3, "####")
    intLandWt = Val(Text3)
    Select Case intLandWt
    
        Case 4800 To 4825
            Line1.Y2 = 5624
        Case 4826 To 4849
            Line1.Y2 = 5601
        Case 4850 To 4875
            Line1.Y2 = 5578
        Case 4876 To 4899
            Line1.Y2 = 5555
        Case 4900 To 4925
            Line1.Y2 = 5532
        Case 4926 To 4949
            Line1.Y2 = 5509
        Case 4950 To 4975
            Line1.Y2 = 5486
        Case 4976 To 4999
            Line1.Y2 = 5463
        Case 5000 To 5025
            Line1.Y2 = 5440
        Case 5026 To 5049
            Line1.Y2 = 5417
        Case 5050 To 5075
            Line1.Y2 = 5394
        Case 5076 To 5099
            Line1.Y2 = 5371
        Case 5100 To 5125
            Line1.Y2 = 5348
        Case 5126 To 5149
            Line1.Y2 = 5325
        Case 5150 To 5175
            Line1.Y2 = 5302
        Case 5176 To 5199
            Line1.Y2 = 5279
        Case 5200 To 5225
            Line1.Y2 = 5256
        Case 5226 To 5249
            Line1.Y2 = 5233
        Case 5250 To 5275
            Line1.Y2 = 5210
        Case 5276 To 5299
            Line1.Y2 = 5187
        Case 5300 To 5325
            Line1.Y2 = 5164
        Case 5326 To 5349
            Line1.Y2 = 5141
        Case 5350 To 5375
            Line1.Y2 = 5118
        Case 5376 To 5399
            Line1.Y2 = 5095
        Case 5400 To 5425
            Line1.Y2 = 5072
        Case 5426 To 5449
            Line1.Y2 = 5049
        Case 5450 To 5475
            Line1.Y2 = 5026
        Case 5476 To 5499
            Line1.Y2 = 5003
        Case 5500 To 5525
            Line1.Y2 = 4980
        Case 5526 To 5549
            Line1.Y2 = 4957
        Case 5550 To 5575
            Line1.Y2 = 4934
        Case 5576 To 5599
            Line1.Y2 = 4911
        Case 5600 To 5625
            Line1.Y2 = 4888
        Case 5626 To 5649
            Line1.Y2 = 4865
        Case 5650 To 5675
            Line1.Y2 = 4842
        Case 5676 To 5699
            Line1.Y2 = 4819
        Case 5700 To 5725
            Line1.Y2 = 4796
        Case 5726 To 5749
            Line1.Y2 = 4773
        Case 5750 To 5775
            Line1.Y2 = 4750
        Case 5776 To 5799
            Line1.Y2 = 4727
        Case 5800 To 5825
            Line1.Y2 = 4704
        Case 5826 To 5849
            Line1.Y2 = 4681
        Case 5850 To 5875
            Line1.Y2 = 4658
        Case 5876 To 5899
            Line1.Y2 = 4635
        Case 5900 To 5925
            Line1.Y2 = 4612
        Case 5926 To 5949
            Line1.Y2 = 4589
        Case 5950 To 5975
            Line1.Y2 = 4566
        Case 5976 To 5999
            Line1.Y2 = 4543
        Case 6000 To 6025
            Line1.Y2 = 4520
        Case 6026 To 6049
            Line1.Y2 = 4497
        Case 6050 To 6075
            Line1.Y2 = 4474
        Case 6076 To 6099
            Line1.Y2 = 4451
        Case 6100 To 6125
            Line1.Y2 = 4428
        Case 6126 To 6149
            Line1.Y2 = 4405
        Case 6150 To 6175
            Line1.Y2 = 4382
        Case 6176 To 6199
            Line1.Y2 = 4359
        Case 6200 To 6225
            Line1.Y2 = 4336
        Case 6226 To 6249
            Line1.Y2 = 4313
        Case 6250 To 6275
            Line1.Y2 = 4290
        Case 6276 To 6299
            Line1.Y2 = 4267
        Case 6300 To 6325
            Line1.Y2 = 4244
        Case 6326 To 6349
            Line1.Y2 = 4221
        Case 6350 To 6375
            Line1.Y2 = 4198
        Case 6376 To 6399
            Line1.Y2 = 4175
        Case 6400 To 6425
            Line1.Y2 = 4152
        Case 6426 To 6449
            Line1.Y2 = 4129
        Case 6450 To 6475
            Line1.Y2 = 4106
        Case 6476 To 6499
            Line1.Y2 = 4083
        Case 6500 To 6525
            Line1.Y2 = 4060
        Case 6526 To 6549
            Line1.Y2 = 4037
        Case 6550 To 6575
            Line1.Y2 = 4014
        Case 6576 To 6599
            Line1.Y2 = 3991
        Case 6600 To 6625
            Line1.Y2 = 3968
        Case 6626 To 6649
            Line1.Y2 = 3945
        Case 6650 To 6675
            Line1.Y2 = 3922
        Case 6676 To 6699
            Line1.Y2 = 3899
        Case 6700 To 6725
            Line1.Y2 = 3876
        Case 6726 To 6749
            Line1.Y2 = 3853
        Case 6750 To 6775
            Line1.Y2 = 3830
        Case 6776 To 6799
            Line1.Y2 = 3807
        Case 6800 To 6849
            Line1.Y2 = 3784
        Case 6800 To 6849
            Line1.Y2 = 3761
        Case 6850 To 6875
            Line1.Y2 = 3738
        Case 6876 To 6899
            Line1.Y2 = 3715
        Case 6900 To 6925
            Line1.Y2 = 3692
        Case 6926 To 6949
            Line1.Y2 = 3669
        Case 6950 To 6975
            Line1.Y2 = 3646
        Case 6976 To 6999
            Line1.Y2 = 3623
        Case 7000 To 7025
            Line1.Y2 = 3600
        Case 7026 To 7049
            Line1.Y2 = 3577
        Case 7050 To 7099
            Line1.Y2 = 3554
        Case 7050 To 7099
            Line1.Y2 = 3531
        Case 7100 To 7125
            Line1.Y2 = 3508
        Case 7126 To 7149
            Line1.Y2 = 3485
        Case 7150 To 7175
            Line1.Y2 = 3462
        Case 7176 To 7199
            Line1.Y2 = 3439
        Case 7200 To 7225
            Line1.Y2 = 3416
        Case 7226 To 7249
            Line1.Y2 = 3393
        Case 7250 To 7275
            Line1.Y2 = 3370
        Case 7276 To 7299
            Line1.Y2 = 3347
        Case 7300 To 7325
            Line1.Y2 = 3324
        Case 7326 To 7349
            Line1.Y2 = 3301
        Case 7350 To 7375
            Line1.Y2 = 3278
        Case 7376 To 7399
            Line1.Y2 = 3255
        Case 7400 To 7425
            Line1.Y2 = 3232
        Case 7426 To 7449
            Line1.Y2 = 3209
        Case 7450 To 7475
            Line1.Y2 = 3186
        Case 7476 To 7499
            Line1.Y2 = 3163
        Case 7500 To 7525
            Line1.Y2 = 3140
        Case 7526 To 7549
            Line1.Y2 = 3117
        Case 7550 To 7575
            Line1.Y2 = 3094
        Case 7576 To 7599
            Line1.Y2 = 3071
        Case 7600 To 7625
            Line1.Y2 = 3048
        Case 7626 To 7649
            Line1.Y2 = 3025
        Case 7650 To 7699
            Line1.Y2 = 3002
        Case 7650 To 7699
            Line1.Y2 = 2979
        Case 7700 To 7725
            Line1.Y2 = 2956
        Case 7726 To 7749
            Line1.Y2 = 2933
        Case 7750 To 7775
            Line1.Y2 = 2910
        Case 7776 To 7799
            Line1.Y2 = 2887
        Case 7800 To 7825
            Line1.Y2 = 2864
        Case 7826 To 7849
            Line1.Y2 = 2841
        Case 7850 To 7875
            Line1.Y2 = 2818
        Case 7876 To 7899
            Line1.Y2 = 2795
        Case 7900 To 7925
            Line1.Y2 = 2772
        Case 7926 To 7949
            Line1.Y2 = 2749
        Case 7950 To 7975
            Line1.Y2 = 2726
        Case 7976 To 7999
            Line1.Y2 = 2703
        Case 8000 To 8025
            Line1.Y2 = 2680
        Case 8026 To 8049
            Line1.Y2 = 2657
        Case 8050 To 8075
            Line1.Y2 = 2634
        Case 8076 To 8099
            Line1.Y2 = 2611
        Case 8100 To 8125
            Line1.Y2 = 2588
        Case 8126 To 8149
            Line1.Y2 = 2565
        Case 8150 To 8175
            Line1.Y2 = 2542
        Case 8176 To 8199
            Line1.Y2 = 2519
        Case 8200 To 8225
            Line1.Y2 = 2496
        Case 8226 To 8249
            Line1.Y2 = 2473
        Case 8250 To 8275
            Line1.Y2 = 2450
        Case 8276 To 8299
            Line1.Y2 = 2427
        Case 8300 To 8325
            Line1.Y2 = 2404
        Case 8326 To 8349
            Line1.Y2 = 2381
        Case 8350 To 8375
            Line1.Y2 = 2358
        Case 8376 To 8399
            Line1.Y2 = 2335
        Case 8400 To 8425
            Line1.Y2 = 2312
        Case 8426 To 8449
            Line1.Y2 = 2289
        Case 8450 To 8475
            Line1.Y2 = 2266
        Case 8476 To 8499
            Line1.Y2 = 2243
        Case 8500 To 8525
            Line1.Y2 = 2220
        Case 8526 To 8549
            Line1.Y2 = 2197
        Case 8550 To 8575
            Line1.Y2 = 2174
        Case 8576 To 8599
            Line1.Y2 = 2151
        Case 8600 To 8625
            Line1.Y2 = 2128
        Case 8626 To 8649
            Line1.Y2 = 2105
        Case 8650 To 8675
            Line1.Y2 = 2082
        Case 8676 To 8699
            Line1.Y2 = 2059
        Case 8700 To 8725
            Line1.Y2 = 2036
        Case 8726 To 8749
            Line1.Y2 = 2013
        Case 8750 To 8775
            Line1.Y2 = 1990
        Case 8776 To 8799
            Line1.Y2 = 1967
            
     End Select
     
     Select Case intLandArm
     
        Case 175# To 175.25
            Line1.X2 = 10600
        Case 175.26 To 175.5
            Line1.X2 = 10634
        Case 175.51 To 175.75
            Line1.X2 = 10668
        Case 175.76 To 175.99
            Line1.X2 = 10702
        Case 176# To 176.25
            Line1.X2 = 10736
        Case 176.26 To 176.5
            Line1.X2 = 10770
        Case 176.51 To 176.75
            Line1.X2 = 10804
        Case 176.76 To 176.99
            Line1.X2 = 10738
        Case 177 To 177.25
            Line1.X2 = 10872
        Case 177.26 To 177.5
            Line1.X2 = 10906
        Case 177.51 To 177.75
            Line1.X2 = 10940
        Case 177.76 To 177.99
            Line1.X2 = 10974
        Case 178 To 178.25
            Line1.X2 = 11008
        Case 178.26 To 178.5
            Line1.X2 = 11042
        Case 178.51 To 178.75
            Line1.X2 = 11076
        Case 178.76 To 178.99
            Line1.X2 = 11110
        Case 179 To 179.25
            Line1.X2 = 11144
        Case 179.26 To 179.5
            Line1.X2 = 11178
        Case 179.51 To 179.75
            Line1.X2 = 11212
        Case 179.76 To 179.99
            Line1.X2 = 11246
        Case 180 To 180.25
            Line1.X2 = 11280
        Case 180.26 To 180.5
            Line1.X2 = 11314
        Case 180.51 To 180.75
            Line1.X2 = 11348
        Case 180.76 To 180.99
            Line1.X2 = 11382
        Case 181 To 181.25
            Line1.X2 = 11416
        Case 181.26 To 181.5
            Line1.X2 = 11450
        Case 181.51 To 181.75
            Line1.X2 = 11484
        Case 181.76 To 181.99
            Line1.X2 = 11518
        Case 182 To 182.25
            Line1.X2 = 11552
        Case 182.26 To 182.5
            Line1.X2 = 11586
        Case 182.51 To 182.75
            Line1.X2 = 11620
        Case 182.76 To 182.99
            Line1.X2 = 11654
        Case 183 To 183.25
            Line1.X2 = 11688
        Case 183.26 To 183.5
            Line1.X2 = 11722
        Case 183.51 To 183.75
            Line1.X2 = 11756
        Case 183.76 To 183.99
            Line1.X2 = 11790
        Case 184 To 184.25
            Line1.X2 = 11824
        Case 184.26 To 184.5
            Line1.X2 = 11858
        Case 184.51 To 184.75
            Line1.X2 = 11892
        Case 184.76 To 184.99
            Line1.X2 = 11926
        Case 185 To 185.25
            Line1.X2 = 11960
        Case 185.26 To 185.5
            Line1.X2 = 11994
        Case 185.51 To 185.75
            Line1.X2 = 12028
        Case 185.76 To 185.99
            Line1.X2 = 12062
        Case 186 To 186.25
            Line1.X2 = 12096
        Case 186.26 To 186.5
            Line1.X2 = 12130
        Case 186.51 To 186.75
            Line1.X2 = 12164
        Case 186.76 To 186.99
            Line1.X2 = 12198
        Case 187 To 187.25
            Line1.X2 = 12232
        Case 187.26 To 187.5
            Line1.X2 = 12266
        Case 187.51 To 187.75
            Line1.X2 = 12300
        Case 187.76 To 187.99
            Line1.X2 = 12334
        Case 188 To 188.25
            Line1.X2 = 12368
        Case 188.26 To 188.5
            Line1.X2 = 12402
        Case 188.51 To 188.75
            Line1.X2 = 12436
        Case 188.76 To 188.99
            Line1.X2 = 12470
        Case 189 To 189.25
            Line1.X2 = 12504
        Case 189.26 To 189.5
            Line1.X2 = 12538
        Case 189.51 To 189.75
            Line1.X2 = 12572
        Case 189.76 To 189.99
            Line1.X2 = 12606
        Case 190 To 190.25
            Line1.X2 = 12640
        Case 190.26 To 190.5
            Line1.X2 = 12674
        Case 190.51 To 190.75
            Line1.X2 = 12708
        Case 190.76 To 190.99
            Line1.X2 = 12742
        Case 191 To 191.25
            Line1.X2 = 12776
        Case 191.26 To 191.5
            Line1.X2 = 12810
        Case 191.51 To 191.75
            Line1.X2 = 12844
        Case 191.76 To 191.99
            Line1.X2 = 12878
        Case 192 To 192.25
            Line1.X2 = 12912
        Case 192.26 To 192.5
            Line1.X2 = 12946
        Case 192.51 To 192.75
            Line1.X2 = 12980
        Case 192.76 To 192.99
            Line1.X2 = 13014
        Case 193 To 193.25
            Line1.X2 = 13048
        Case 193.26 To 193.5
            Line1.X2 = 13082
        Case 193.51 To 193.75
            Line1.X2 = 13116
        Case 193.76 To 193.99
            Line1.X2 = 13150
        Case 194 To 194.25
            Line1.X2 = 13184
        Case 194.26 To 194.5
            Line1.X2 = 13218
        Case 194.51 To 194.75
            Line1.X2 = 13252
        Case 194.76 To 194.99
            Line1.X2 = 13286
        Case 195 To 195.25
            Line1.X2 = 13320
        Case 195.26 To 195.5
            Line1.X2 = 13354
        Case 195.51 To 195.75
            Line1.X2 = 13388
        Case 195.76 To 195.99
            Line1.X2 = 13422
        Case 196 To 196.25
            Line1.X2 = 13456
        Case 196.26 To 196.5
            Line1.X2 = 13490
        Case 196.51 To 196.75
            Line1.X2 = 13524
        Case 196.76 To 196.99
            Line1.X2 = 13558
        Case 197 To 197.25
            Line1.X2 = 13592
        Case 197.26 To 197.5
            Line1.X2 = 13626
        Case 197.51 To 197.75
            Line1.X2 = 13660
        Case 197.76 To 197.99
            Line1.X2 = 13694
        Case 198 To 198.25
            Line1.X2 = 13728
        Case 198.26 To 198.5
            Line1.X2 = 13762
        Case 198.51 To 198.75
            Line1.X2 = 13796
        Case 198.76 To 198.99
            Line1.X2 = 13830
        Case 199 To 199.25
            Line1.X2 = 13864
        Case 199.26 To 199.5
            Line1.X2 = 13898
        Case 199.51 To 199.75
            Line1.X2 = 13932
        Case 199.76 To 199.99
            Line1.X2 = 13966
        Case 200 To 200.25
            Line1.X2 = 14000
        Case 200.26 To 200.5
            Line1.X2 = 14034
        Case 200.51 To 200.75
            Line1.X2 = 14068
        Case 200.76 To 200.99
            Line1.X2 = 14102
        Case 201 To 201.25
            Line1.X2 = 14136
        Case 201.26 To 201.5
            Line1.X2 = 14170
        Case 201.51 To 201.75
            Line1.X2 = 14204
        Case 201.76 To 201.99
            Line1.X2 = 14238
        Case 202 To 202.25
            Line1.X2 = 14272
        Case 202.26 To 202.5
            Line1.X2 = 14306
        Case 202.51 To 202.75
            Line1.X2 = 14340
        Case 202.76 To 202.99
            Line1.X2 = 14374
        Case 203 To 203.25
            Line1.X2 = 14408
        Case 203.26 To 203.5
            Line1.X2 = 14442
        Case 203.51 To 203.75
            Line1.X2 = 14476
        Case 203.76 To 203.99
            Line1.X2 = 14510
        Case 204 To 204.25
            Line1.X2 = 14544
        Case 204.26 To 204.5
            Line1.X2 = 14578
        Case 204.51 To 204.75
            Line1.X2 = 14612
        Case 204.76 To 204.99
            Line1.X2 = 14646
        Case 205 To 205.25
            Line1.X2 = 14680
        Case 205.26 To 204.5
            Line1.X2 = 14714
        Case 205.51 To 205.75
            Line1.X2 = 14748
        Case 205.76 To 205.99
            Line1.X2 = 14782
            
    End Select
    
    Call cmdPrint_Click
    frmPrint.Hide
    frmMain.Show
    
End Sub

Private Sub Form_Load()

    Text51 = "(*) I CERTIFY THAT THE LOADING OF THIS AC IS IN ACCORDANCE WITH CURRENT" & vbNewLine & _
    "LOAD INSTRUCTIONS AS PROVIDED BY SIGNED LOADPLAN AND I CERTIFY THAT THE" & vbNewLine & _
    "CARGO ON THE ABOVE FLIGHT HAS BEEN HANDLED IN ACCORDANCE WITH THE" & vbNewLine & _
    "SECURITY REQUIREMENTS APPLICABLE TO ECAC POLICY (DOC 30)"
    
End Sub
