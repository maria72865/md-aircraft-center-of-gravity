VERSION 5.00
Begin VB.Form frm2 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Weight and Balance"
   ClientHeight    =   8175
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11685
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
   ScaleHeight     =   8175
   ScaleWidth      =   11685
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
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
      Height          =   12600
      Left            =   0
      ScaleHeight     =   12540
      ScaleWidth      =   15900
      TabIndex        =   15
      Top             =   0
      Width           =   15960
      Begin VB.TextBox Text75 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
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
         Left            =   3680
         TabIndex        =   215
         Top             =   3800
         Width           =   1005
      End
      Begin VB.TextBox txtW11 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
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
         Left            =   4650
         TabIndex        =   219
         Top             =   3800
         Width           =   1185
      End
      Begin VB.TextBox Text79 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
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
         Left            =   8160
         TabIndex        =   218
         Top             =   3800
         Width           =   1185
      End
      Begin VB.TextBox txtM11 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
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
         Left            =   6990
         TabIndex        =   217
         Top             =   3800
         Width           =   1185
      End
      Begin VB.TextBox txtA11 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
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
         Left            =   5820
         TabIndex        =   216
         Top             =   3800
         Width           =   1185
      End
      Begin VB.TextBox txtMax13 
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
         Left            =   8160
         TabIndex        =   214
         Top             =   4980
         Width           =   1185
      End
      Begin VB.TextBox txtMax12 
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
         Left            =   8160
         TabIndex        =   213
         Top             =   4680
         Width           =   1185
      End
      Begin VB.TextBox txtMax11 
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
         Left            =   8160
         TabIndex        =   212
         Top             =   4380
         Width           =   1185
      End
      Begin VB.TextBox txtMax10 
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
         Left            =   8160
         TabIndex        =   211
         Top             =   4080
         Width           =   1185
      End
      Begin VB.TextBox txtM10 
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
         Left            =   6990
         TabIndex        =   210
         Top             =   4980
         Width           =   1185
      End
      Begin VB.TextBox txtM9 
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
         Left            =   6990
         TabIndex        =   209
         Top             =   4680
         Width           =   1185
      End
      Begin VB.TextBox txtM8 
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
         Left            =   6990
         TabIndex        =   208
         Top             =   4380
         Width           =   1185
      End
      Begin VB.TextBox txtM7 
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
         Left            =   6990
         TabIndex        =   207
         Top             =   4080
         Width           =   1185
      End
      Begin VB.TextBox txtA10 
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
         Left            =   5820
         TabIndex        =   206
         Top             =   4980
         Width           =   1185
      End
      Begin VB.TextBox txtA9 
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
         Left            =   5820
         TabIndex        =   205
         Top             =   4680
         Width           =   1185
      End
      Begin VB.TextBox txtA8 
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
         Left            =   5820
         TabIndex        =   204
         Top             =   4380
         Width           =   1185
      End
      Begin VB.TextBox txtA7 
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
         Left            =   5820
         TabIndex        =   203
         Top             =   4080
         Width           =   1185
      End
      Begin VB.TextBox Text74 
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
         Left            =   3680
         TabIndex        =   198
         Text            =   "Pod D"
         Top             =   4980
         Width           =   1005
      End
      Begin VB.TextBox Text64 
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
         Left            =   3680
         TabIndex        =   197
         Text            =   "Pod C"
         Top             =   4680
         Width           =   1005
      End
      Begin VB.TextBox Text60 
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
         Left            =   3680
         TabIndex        =   196
         Text            =   "Pod B"
         Top             =   4380
         Width           =   1005
      End
      Begin VB.TextBox Text8 
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
         Left            =   3680
         TabIndex        =   195
         Text            =   "Pod A"
         Top             =   4080
         Width           =   1005
      End
      Begin VB.TextBox txtZoneD 
         Height          =   375
         Left            =   5520
         TabIndex        =   194
         Top             =   10440
         Width           =   615
      End
      Begin VB.TextBox txtZoneC 
         Height          =   375
         Left            =   4320
         TabIndex        =   193
         Top             =   10440
         Width           =   735
      End
      Begin VB.TextBox txtZoneB 
         Height          =   375
         Left            =   3480
         TabIndex        =   192
         Top             =   10440
         Width           =   615
      End
      Begin VB.TextBox txtZoneA 
         Height          =   330
         Left            =   2040
         TabIndex        =   191
         Top             =   10440
         Width           =   975
      End
      Begin VB.TextBox Text7 
         Height          =   330
         Left            =   10440
         TabIndex        =   184
         Text            =   "Text7"
         Top             =   3480
         Visible         =   0   'False
         Width           =   1150
      End
      Begin VB.TextBox Text6 
         Height          =   330
         Left            =   10440
         TabIndex        =   182
         Text            =   "Text6"
         Top             =   3000
         Visible         =   0   'False
         Width           =   1150
      End
      Begin VB.TextBox Text5 
         Height          =   330
         Left            =   10440
         TabIndex        =   179
         Text            =   "Text5"
         Top             =   3840
         Visible         =   0   'False
         Width           =   1150
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   10440
         TabIndex        =   177
         Text            =   "Text4"
         Top             =   4200
         Visible         =   0   'False
         Width           =   1150
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   1080
         TabIndex        =   176
         Text            =   "Text3"
         Top             =   8520
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   1200
         TabIndex        =   175
         Text            =   "Text2"
         Top             =   8520
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   330
         Left            =   1320
         TabIndex        =   174
         Text            =   "Text1"
         Top             =   8520
         Visible         =   0   'False
         Width           =   855
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
         Locked          =   -1  'True
         TabIndex        =   171
         TabStop         =   0   'False
         Top             =   7500
         Width           =   950
      End
      Begin VB.CommandButton cmdClear 
         BackColor       =   &H000000FF&
         Caption         =   "C&lear Changes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   9600
         Style           =   1  'Graphical
         TabIndex        =   173
         Top             =   6690
         Width           =   1215
      End
      Begin VB.CommandButton cmdLMC 
         BackColor       =   &H000000FF&
         Caption         =   "&Calculate Changes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   9600
         Style           =   1  'Graphical
         TabIndex        =   172
         Top             =   5930
         Width           =   1215
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "&Exit"
         Height          =   375
         Left            =   9840
         TabIndex        =   10
         Top             =   1920
         Width           =   1095
      End
      Begin VB.CommandButton cmdBack 
         Caption         =   "&Back"
         Height          =   375
         Left            =   9840
         TabIndex        =   9
         Top             =   1320
         Width           =   1095
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
         TabIndex        =   170
         TabStop         =   0   'False
         Text            =   "Underload before LMC"
         Top             =   7500
         Width           =   1920
      End
      Begin VB.TextBox txtPrep 
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
         Height          =   650
         Left            =   12000
         TabIndex        =   160
         TabStop         =   0   'False
         Text            =   "L. Pamos"
         Top             =   8400
         Visible         =   0   'False
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
         Height          =   650
         Left            =   12000
         TabIndex        =   156
         TabStop         =   0   'False
         Text            =   " A. Haynes"
         Top             =   7440
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox Text56 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   10680
         TabIndex        =   154
         TabStop         =   0   'False
         Text            =   "Captain"
         Top             =   7440
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox Text73 
         Appearance      =   0  'Flat
         Height          =   650
         Left            =   13320
         TabIndex        =   169
         TabStop         =   0   'False
         Text            =   " "
         Top             =   10200
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox Text72 
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
         Height          =   650
         Left            =   12000
         TabIndex        =   168
         TabStop         =   0   'False
         Text            =   " "
         Top             =   10200
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox Text71 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   10680
         TabIndex        =   167
         TabStop         =   0   'False
         Text            =   "Prepared by"
         Top             =   10515
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox Text70 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   10680
         TabIndex        =   166
         TabStop         =   0   'False
         Text            =   "(*) Load"
         Top             =   10200
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox Text69 
         Appearance      =   0  'Flat
         Height          =   650
         Left            =   13320
         TabIndex        =   165
         TabStop         =   0   'False
         Text            =   " "
         Top             =   9360
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox Text68 
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
         Height          =   650
         Left            =   12000
         TabIndex        =   164
         TabStop         =   0   'False
         Text            =   " "
         Top             =   9360
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox Text67 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   10680
         TabIndex        =   163
         TabStop         =   0   'False
         Text            =   "Checked by"
         Top             =   9675
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox Text66 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   10680
         TabIndex        =   162
         TabStop         =   0   'False
         Text            =   "Load"
         Top             =   9360
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox Text65 
         Appearance      =   0  'Flat
         Height          =   650
         Left            =   13320
         TabIndex        =   161
         TabStop         =   0   'False
         Text            =   " "
         Top             =   8400
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox Text63 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   10680
         TabIndex        =   159
         TabStop         =   0   'False
         Text            =   "Prepared by"
         Top             =   8715
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox Text62 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   10680
         TabIndex        =   158
         TabStop         =   0   'False
         Text            =   "W & B"
         Top             =   8400
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox Text61 
         Appearance      =   0  'Flat
         Height          =   650
         Left            =   13320
         TabIndex        =   157
         TabStop         =   0   'False
         Text            =   " "
         Top             =   7440
         Visible         =   0   'False
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
         TabIndex        =   153
         TabStop         =   0   'False
         Text            =   " "
         Top             =   9090
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox Text57 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   4920
         TabIndex        =   152
         TabStop         =   0   'False
         Text            =   "Underload Before LMC"
         Top             =   9090
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.TextBox txtTotalPay 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   7080
         TabIndex        =   151
         TabStop         =   0   'False
         Text            =   " "
         Top             =   8775
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox Text55 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   4920
         TabIndex        =   150
         TabStop         =   0   'False
         Text            =   "Total Payload"
         Top             =   8775
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.TextBox Text54 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   7080
         TabIndex        =   149
         TabStop         =   0   'False
         Text            =   "1542.00"
         Top             =   8460
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox Text53 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   4920
         TabIndex        =   148
         TabStop         =   0   'False
         Text            =   "Allowed Payload"
         Top             =   8460
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.TextBox Text52 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   4920
         TabIndex        =   147
         TabStop         =   0   'False
         Text            =   "All these weights in Kilograms"
         Top             =   8145
         Visible         =   0   'False
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
         Height          =   855
         Left            =   4920
         MultiLine       =   -1  'True
         TabIndex        =   146
         Top             =   10800
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
         TabIndex        =   145
         Text            =   "Weight in Kgs"
         Top             =   6915
         Width           =   1575
      End
      Begin VB.TextBox Text49 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6360
         TabIndex        =   144
         TabStop         =   0   'False
         Text            =   "+/-"
         Top             =   6915
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
         TabIndex        =   143
         TabStop         =   0   'False
         Text            =   "Comp"
         Top             =   6915
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
         TabIndex        =   142
         TabStop         =   0   'False
         Text            =   "Destination"
         Top             =   6915
         Width           =   1455
      End
      Begin VB.TextBox Text46 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3480
         TabIndex        =   141
         TabStop         =   0   'False
         Text            =   "Last Minute Change"
         Top             =   6615
         Width           =   5895
      End
      Begin VB.TextBox Text45 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3670
         TabIndex        =   140
         TabStop         =   0   'False
         Text            =   "Center of Gravity Calculations"
         Top             =   495
         Width           =   5675
      End
      Begin VB.TextBox Text44 
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
         Left            =   3680
         TabIndex        =   139
         TabStop         =   0   'False
         Text            =   "TO Wt"
         Top             =   5880
         Width           =   1005
      End
      Begin VB.TextBox Text43 
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
         Left            =   3680
         TabIndex        =   138
         TabStop         =   0   'False
         Text            =   "TO Fuel"
         Top             =   5580
         Width           =   1005
      End
      Begin VB.TextBox Text42 
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
         Left            =   3680
         TabIndex        =   137
         TabStop         =   0   'False
         Text            =   "ZFW"
         Top             =   5280
         Width           =   1005
      End
      Begin VB.TextBox Text41 
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
         Left            =   3680
         TabIndex        =   136
         TabStop         =   0   'False
         Text            =   "Hold 6"
         Top             =   3500
         Width           =   1005
      End
      Begin VB.TextBox Text40 
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
         Left            =   3680
         TabIndex        =   135
         TabStop         =   0   'False
         Text            =   "Hold 5"
         Top             =   3200
         Width           =   1005
      End
      Begin VB.TextBox Text36 
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
         Left            =   3680
         TabIndex        =   134
         TabStop         =   0   'False
         Text            =   "Hold 4"
         Top             =   2900
         Width           =   1005
      End
      Begin VB.TextBox Text35 
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
         Left            =   3680
         TabIndex        =   133
         TabStop         =   0   'False
         Text            =   "Hold 3"
         Top             =   2600
         Width           =   1005
      End
      Begin VB.TextBox Text34 
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
         Left            =   3680
         TabIndex        =   132
         TabStop         =   0   'False
         Text            =   "Hold 2"
         Top             =   2300
         Width           =   1005
      End
      Begin VB.TextBox Text33 
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
         Left            =   3680
         TabIndex        =   131
         TabStop         =   0   'False
         Text            =   "Hold 1"
         Top             =   2000
         Width           =   1005
      End
      Begin VB.TextBox Text32 
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
         Left            =   3680
         TabIndex        =   130
         TabStop         =   0   'False
         Text            =   "Extra Crew"
         Top             =   1700
         Width           =   1005
      End
      Begin VB.TextBox Text31 
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
         Left            =   3680
         TabIndex        =   129
         TabStop         =   0   'False
         Text            =   "Crew"
         Top             =   1400
         Width           =   1005
      End
      Begin VB.TextBox Text30 
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
         Left            =   3680
         TabIndex        =   128
         TabStop         =   0   'False
         Text            =   "BEW"
         Top             =   1100
         Width           =   1005
      End
      Begin VB.TextBox Text29 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   8160
         TabIndex        =   127
         TabStop         =   0   'False
         Text            =   "Max. Load"
         Top             =   790
         Width           =   1185
      End
      Begin VB.TextBox Text28 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6990
         TabIndex        =   126
         TabStop         =   0   'False
         Text            =   "Moment"
         Top             =   790
         Width           =   1185
      End
      Begin VB.TextBox Text27 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5820
         TabIndex        =   125
         TabStop         =   0   'False
         Text            =   "ARM"
         Top             =   790
         Width           =   1185
      End
      Begin VB.TextBox Text25 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   3680
         TabIndex        =   123
         TabStop         =   0   'False
         Top             =   790
         Width           =   1005
      End
      Begin VB.TextBox Text24 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2220
         TabIndex        =   122
         TabStop         =   0   'False
         Text            =   "Max Wt lbs"
         Top             =   3420
         Width           =   950
      End
      Begin VB.TextBox Text23 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1340
         TabIndex        =   121
         TabStop         =   0   'False
         Text            =   "Weight"
         Top             =   3420
         Width           =   950
      End
      Begin VB.TextBox Text22 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Height          =   330
         Left            =   340
         TabIndex        =   120
         TabStop         =   0   'False
         Text            =   " "
         Top             =   3420
         Width           =   1010
      End
      Begin VB.TextBox Text21 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   340
         TabIndex        =   119
         TabStop         =   0   'False
         Text            =   "Gross Weight Computation"
         Top             =   3120
         Width           =   2815
      End
      Begin VB.TextBox Text20 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   340
         TabIndex        =   118
         TabStop         =   0   'False
         Text            =   "LW"
         Top             =   7185
         Width           =   1010
      End
      Begin VB.TextBox Text19 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   340
         TabIndex        =   117
         TabStop         =   0   'False
         Text            =   "Trip Fuel"
         Top             =   6870
         Width           =   1010
      End
      Begin VB.TextBox Text18 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   340
         TabIndex        =   116
         TabStop         =   0   'False
         Text            =   "TOW"
         Top             =   6555
         Width           =   1010
      End
      Begin VB.TextBox Text17 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   340
         TabIndex        =   115
         TabStop         =   0   'False
         Text            =   "Taxi Fuel"
         Top             =   6240
         Width           =   1010
      End
      Begin VB.TextBox Text16 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   340
         TabIndex        =   114
         TabStop         =   0   'False
         Text            =   "Ramp Wt"
         Top             =   5925
         Width           =   1010
      End
      Begin VB.TextBox Text15 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   340
         TabIndex        =   113
         TabStop         =   0   'False
         Text            =   "Ramp Fuel"
         Top             =   5610
         Width           =   1010
      End
      Begin VB.TextBox Text14 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   340
         TabIndex        =   112
         TabStop         =   0   'False
         Text            =   "ZFW"
         Top             =   5295
         Width           =   1010
      End
      Begin VB.TextBox Text13 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   340
         TabIndex        =   111
         TabStop         =   0   'False
         Text            =   "Cargo"
         Top             =   4980
         Width           =   1010
      End
      Begin VB.TextBox Text12 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   340
         TabIndex        =   110
         TabStop         =   0   'False
         Text            =   "BOW"
         Top             =   4665
         Width           =   1010
      End
      Begin VB.TextBox Text11 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   340
         TabIndex        =   109
         TabStop         =   0   'False
         Text            =   "Extra Crew"
         Top             =   4350
         Width           =   1010
      End
      Begin VB.TextBox Text10 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   340
         TabIndex        =   108
         TabStop         =   0   'False
         Text            =   "Crew"
         Top             =   4035
         Width           =   1010
      End
      Begin VB.TextBox Text9 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   340
         TabIndex        =   107
         TabStop         =   0   'False
         Text            =   "BEW"
         Top             =   3720
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
         Height          =   330
         Left            =   7800
         TabIndex        =   7
         Text            =   " "
         Top             =   7515
         Width           =   1575
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
         Height          =   330
         Left            =   3480
         TabIndex        =   5
         Text            =   " "
         Top             =   7515
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
         Height          =   320
         Left            =   7800
         TabIndex        =   4
         Text            =   " "
         Top             =   7215
         Width           =   1575
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
         Height          =   320
         Left            =   3480
         TabIndex        =   2
         Text            =   " "
         Top             =   7215
         Width           =   1455
      End
      Begin VB.TextBox txtMax6 
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
         Left            =   8160
         Locked          =   -1  'True
         TabIndex        =   105
         TabStop         =   0   'False
         Top             =   3500
         Width           =   1185
      End
      Begin VB.TextBox txtMax9 
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
         Left            =   8160
         Locked          =   -1  'True
         TabIndex        =   104
         TabStop         =   0   'False
         Text            =   "8750 lbs"
         Top             =   5880
         Width           =   1185
      End
      Begin VB.TextBox txtMax8 
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
         Left            =   8160
         Locked          =   -1  'True
         TabIndex        =   103
         TabStop         =   0   'False
         Text            =   "2224 lbs"
         Top             =   5580
         Width           =   1185
      End
      Begin VB.TextBox txtMax7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
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
         Left            =   8160
         Locked          =   -1  'True
         TabIndex        =   102
         TabStop         =   0   'False
         Top             =   5280
         Width           =   1185
      End
      Begin VB.TextBox txtMax5 
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
         Left            =   8160
         Locked          =   -1  'True
         TabIndex        =   101
         TabStop         =   0   'False
         Top             =   3200
         Width           =   1185
      End
      Begin VB.TextBox txtMax4 
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
         Left            =   8160
         Locked          =   -1  'True
         TabIndex        =   100
         TabStop         =   0   'False
         Top             =   2900
         Width           =   1185
      End
      Begin VB.TextBox txtMax3 
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
         Left            =   8160
         Locked          =   -1  'True
         TabIndex        =   99
         TabStop         =   0   'False
         Top             =   2600
         Width           =   1185
      End
      Begin VB.TextBox txtMax2 
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
         Left            =   8160
         Locked          =   -1  'True
         TabIndex        =   98
         TabStop         =   0   'False
         Top             =   2300
         Width           =   1185
      End
      Begin VB.TextBox txtMax1 
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
         Left            =   8160
         Locked          =   -1  'True
         TabIndex        =   97
         TabStop         =   0   'False
         Top             =   2000
         Width           =   1185
      End
      Begin VB.TextBox Text39 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
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
         Left            =   8160
         TabIndex        =   96
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1700
         Width           =   1185
      End
      Begin VB.TextBox Text38 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
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
         Left            =   8160
         TabIndex        =   95
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1400
         Width           =   1185
      End
      Begin VB.TextBox Text37 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
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
         Left            =   8160
         TabIndex        =   94
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1100
         Width           =   1185
      End
      Begin VB.TextBox txtTOWM1 
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
         Left            =   6990
         Locked          =   -1  'True
         TabIndex        =   93
         TabStop         =   0   'False
         Text            =   " "
         Top             =   5880
         Width           =   1185
      End
      Begin VB.TextBox txtFuelM 
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
         Left            =   6990
         Locked          =   -1  'True
         TabIndex        =   92
         TabStop         =   0   'False
         Text            =   " "
         Top             =   5580
         Width           =   1185
      End
      Begin VB.TextBox txtZEWM 
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
         Left            =   6990
         Locked          =   -1  'True
         TabIndex        =   91
         TabStop         =   0   'False
         Text            =   " "
         Top             =   5280
         Width           =   1185
      End
      Begin VB.TextBox txtM6 
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
         Left            =   6990
         Locked          =   -1  'True
         TabIndex        =   90
         TabStop         =   0   'False
         Text            =   " "
         Top             =   3500
         Width           =   1185
      End
      Begin VB.TextBox txtM5 
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
         Left            =   6990
         Locked          =   -1  'True
         TabIndex        =   89
         TabStop         =   0   'False
         Text            =   " "
         Top             =   3200
         Width           =   1185
      End
      Begin VB.TextBox txtM4 
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
         Left            =   6990
         Locked          =   -1  'True
         TabIndex        =   88
         TabStop         =   0   'False
         Text            =   " "
         Top             =   2900
         Width           =   1185
      End
      Begin VB.TextBox txtM3 
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
         Left            =   6990
         Locked          =   -1  'True
         TabIndex        =   87
         TabStop         =   0   'False
         Text            =   " "
         Top             =   2600
         Width           =   1185
      End
      Begin VB.TextBox txtM2 
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
         Left            =   6990
         Locked          =   -1  'True
         TabIndex        =   86
         TabStop         =   0   'False
         Text            =   " "
         Top             =   2300
         Width           =   1185
      End
      Begin VB.TextBox txtM1 
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
         Left            =   6990
         Locked          =   -1  'True
         TabIndex        =   85
         TabStop         =   0   'False
         Text            =   " "
         Top             =   2000
         Width           =   1185
      End
      Begin VB.TextBox txtXtraM 
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
         Left            =   6990
         Locked          =   -1  'True
         TabIndex        =   84
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1700
         Width           =   1185
      End
      Begin VB.TextBox txtCrewM1 
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
         Left            =   6990
         Locked          =   -1  'True
         TabIndex        =   83
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1400
         Width           =   1185
      End
      Begin VB.TextBox txtBEWM 
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
         Left            =   6990
         Locked          =   -1  'True
         TabIndex        =   82
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1100
         Width           =   1185
      End
      Begin VB.TextBox txtTOWA 
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
         Left            =   5820
         Locked          =   -1  'True
         TabIndex        =   81
         TabStop         =   0   'False
         Text            =   " "
         Top             =   5880
         Width           =   1185
      End
      Begin VB.TextBox txtFuelA 
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
         Left            =   5820
         Locked          =   -1  'True
         TabIndex        =   80
         TabStop         =   0   'False
         Text            =   " "
         Top             =   5580
         Width           =   1185
      End
      Begin VB.TextBox txtZEWA 
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
         Left            =   5820
         Locked          =   -1  'True
         TabIndex        =   79
         TabStop         =   0   'False
         Text            =   " "
         Top             =   5280
         Width           =   1185
      End
      Begin VB.TextBox txtA6 
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
         Left            =   5820
         Locked          =   -1  'True
         TabIndex        =   78
         TabStop         =   0   'False
         Text            =   " "
         Top             =   3500
         Width           =   1185
      End
      Begin VB.TextBox txtA5 
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
         Left            =   5820
         Locked          =   -1  'True
         TabIndex        =   77
         TabStop         =   0   'False
         Text            =   " "
         Top             =   3200
         Width           =   1185
      End
      Begin VB.TextBox txtA4 
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
         Left            =   5820
         Locked          =   -1  'True
         TabIndex        =   76
         TabStop         =   0   'False
         Text            =   " "
         Top             =   2900
         Width           =   1185
      End
      Begin VB.TextBox txtA3 
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
         Left            =   5820
         Locked          =   -1  'True
         TabIndex        =   75
         TabStop         =   0   'False
         Text            =   " "
         Top             =   2600
         Width           =   1185
      End
      Begin VB.TextBox txtA2 
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
         Left            =   5820
         Locked          =   -1  'True
         TabIndex        =   74
         TabStop         =   0   'False
         Text            =   " "
         Top             =   2300
         Width           =   1185
      End
      Begin VB.TextBox txtA1 
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
         Left            =   5820
         Locked          =   -1  'True
         TabIndex        =   73
         TabStop         =   0   'False
         Text            =   " "
         Top             =   2000
         Width           =   1185
      End
      Begin VB.TextBox txtXtraA 
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
         Left            =   5820
         Locked          =   -1  'True
         TabIndex        =   72
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1700
         Width           =   1185
      End
      Begin VB.TextBox txtCrewA 
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
         Left            =   5820
         Locked          =   -1  'True
         TabIndex        =   71
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1400
         Width           =   1185
      End
      Begin VB.TextBox txtBEWA 
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
         Left            =   5820
         Locked          =   -1  'True
         TabIndex        =   70
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1100
         Width           =   1185
      End
      Begin VB.TextBox txtTOWW1 
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
         Left            =   4650
         Locked          =   -1  'True
         TabIndex        =   69
         TabStop         =   0   'False
         Text            =   " "
         Top             =   5880
         Width           =   1185
      End
      Begin VB.TextBox txtFuelW 
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
         Left            =   4650
         Locked          =   -1  'True
         TabIndex        =   68
         TabStop         =   0   'False
         Text            =   " "
         Top             =   5580
         Width           =   1185
      End
      Begin VB.TextBox txtZEWW1 
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
         Left            =   4650
         Locked          =   -1  'True
         TabIndex        =   67
         TabStop         =   0   'False
         Text            =   " "
         Top             =   5280
         Width           =   1185
      End
      Begin VB.TextBox txtW6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   4650
         Locked          =   -1  'True
         TabIndex        =   66
         TabStop         =   0   'False
         Text            =   " "
         Top             =   3500
         Width           =   1185
      End
      Begin VB.TextBox txtW5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   4650
         Locked          =   -1  'True
         TabIndex        =   65
         TabStop         =   0   'False
         Text            =   " "
         Top             =   3200
         Width           =   1185
      End
      Begin VB.TextBox txtW4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   4650
         Locked          =   -1  'True
         TabIndex        =   64
         TabStop         =   0   'False
         Text            =   " "
         Top             =   2900
         Width           =   1185
      End
      Begin VB.TextBox txtW3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   4650
         Locked          =   -1  'True
         TabIndex        =   63
         TabStop         =   0   'False
         Text            =   " "
         Top             =   2600
         Width           =   1185
      End
      Begin VB.TextBox txtW2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   4650
         Locked          =   -1  'True
         TabIndex        =   62
         TabStop         =   0   'False
         Text            =   " "
         Top             =   2300
         Width           =   1185
      End
      Begin VB.TextBox txtW1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   61
         TabStop         =   0   'False
         Text            =   " "
         Top             =   2000
         Width           =   1185
      End
      Begin VB.TextBox txtXtraW 
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
         Left            =   4650
         TabIndex        =   60
         TabStop         =   0   'False
         Text            =   "0.00"
         Top             =   1700
         Width           =   1185
      End
      Begin VB.TextBox txtCrewW1 
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
         Left            =   4650
         Locked          =   -1  'True
         TabIndex        =   59
         TabStop         =   0   'False
         Text            =   "340.00"
         Top             =   1400
         Width           =   1185
      End
      Begin VB.TextBox txtBEWW 
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
         Left            =   4650
         Locked          =   -1  'True
         TabIndex        =   58
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1100
         Width           =   1185
      End
      Begin VB.TextBox txtLWM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2220
         Locked          =   -1  'True
         TabIndex        =   57
         TabStop         =   0   'False
         Text            =   "8500"
         Top             =   7185
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
         Locked          =   -1  'True
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   6870
         Width           =   950
      End
      Begin VB.TextBox txtTOWM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2220
         Locked          =   -1  'True
         TabIndex        =   55
         TabStop         =   0   'False
         Text            =   "8750"
         Top             =   6555
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
         Locked          =   -1  'True
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   6240
         Width           =   950
      End
      Begin VB.TextBox txtRampWM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2220
         Locked          =   -1  'True
         TabIndex        =   53
         TabStop         =   0   'False
         Text            =   "8785"
         Top             =   5925
         Width           =   950
      End
      Begin VB.TextBox txtRampFuelM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2220
         Locked          =   -1  'True
         TabIndex        =   52
         TabStop         =   0   'False
         Text            =   "2224"
         Top             =   5610
         Width           =   950
      End
      Begin VB.TextBox txtZFWM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Height          =   330
         Left            =   2220
         Locked          =   -1  'True
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   5295
         Width           =   950
      End
      Begin VB.TextBox txtPayloadM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2220
         Locked          =   -1  'True
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   4980
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
         Locked          =   -1  'True
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   4665
         Width           =   950
      End
      Begin VB.TextBox txtXCrewM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Locked          =   -1  'True
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   4350
         Width           =   950
      End
      Begin VB.TextBox txtCrewM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2220
         Locked          =   -1  'True
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   4035
         Width           =   950
      End
      Begin VB.TextBox txtBWM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2220
         Locked          =   -1  'True
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   3720
         Width           =   950
      End
      Begin VB.TextBox txtTOWW 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1340
         Locked          =   -1  'True
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   6555
         Width           =   950
      End
      Begin VB.TextBox txtLWW 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   7185
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
         Locked          =   -1  'True
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   6870
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
         Locked          =   -1  'True
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   6240
         Width           =   950
      End
      Begin VB.TextBox txtRampWW 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1340
         Locked          =   -1  'True
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   5925
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
         Locked          =   -1  'True
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   5610
         Width           =   950
      End
      Begin VB.TextBox txtZFWW 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1340
         Locked          =   -1  'True
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   5295
         Width           =   950
      End
      Begin VB.TextBox txtPayloadW 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1340
         Locked          =   -1  'True
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   4980
         Width           =   950
      End
      Begin VB.TextBox txtBOWW 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1340
         Locked          =   -1  'True
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   4665
         Width           =   950
      End
      Begin VB.TextBox txtXCrewW 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
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
         TabIndex        =   1
         Text            =   "0.00"
         Top             =   4350
         Width           =   950
      End
      Begin VB.TextBox txtCrewW 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
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
         TabIndex        =   0
         Text            =   "340.00"
         Top             =   4035
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
         Locked          =   -1  'True
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   3720
         Width           =   950
      End
      Begin VB.TextBox txtReg1 
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
         Height          =   360
         Left            =   960
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   2520
         Width           =   2445
      End
      Begin VB.TextBox txtDate 
         Alignment       =   2  'Center
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
         Left            =   960
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   1080
         Width           =   2445
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
         Left            =   960
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   2040
         Width           =   2445
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
         Left            =   960
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   600
         Width           =   2445
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
         Left            =   960
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   1560
         Width           =   2445
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
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   9615
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtZone6 
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
         Left            =   6710
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   9840
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtZone5 
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
         Left            =   6250
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   9840
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtZone4 
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
         Left            =   5770
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   9840
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtZone3 
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
         Left            =   5160
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   9840
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtZone2 
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
         Left            =   4200
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   9840
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtZone1 
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
         Left            =   3420
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   9840
         Visible         =   0   'False
         Width           =   495
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
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   9840
         Visible         =   0   'False
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
         Left            =   9840
         TabIndex        =   8
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox Text59 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   10680
         TabIndex        =   155
         TabStop         =   0   'False
         Text            =   "Accepted by"
         Top             =   7755
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   6360
         TabIndex        =   185
         Top             =   7080
         Width           =   1575
         Begin VB.OptionButton optS1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   840
            TabIndex        =   187
            Top             =   120
            Width           =   495
         End
         Begin VB.OptionButton optA1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "+"
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
            Left            =   120
            TabIndex        =   186
            Top             =   120
            Width           =   615
         End
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
         TabIndex        =   12
         Text            =   " "
         Top             =   6615
         Width           =   1455
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
         TabIndex        =   14
         Text            =   " "
         Top             =   6920
         Width           =   1455
      End
      Begin VB.ComboBox cmbHold2 
         Height          =   345
         Left            =   4910
         TabIndex        =   6
         Top             =   7515
         Width           =   1485
      End
      Begin VB.ComboBox cmbHold1 
         Height          =   345
         Left            =   4910
         TabIndex        =   3
         Top             =   7200
         Width           =   1485
      End
      Begin VB.TextBox txtComp2 
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
         Height          =   320
         Left            =   4920
         TabIndex        =   13
         Text            =   " "
         Top             =   6922
         Width           =   1455
      End
      Begin VB.TextBox txtComp1 
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
         TabIndex        =   11
         Text            =   " "
         Top             =   6615
         Width           =   1455
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   6360
         TabIndex        =   188
         Top             =   7410
         Width           =   1455
         Begin VB.OptionButton optA2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            TabIndex        =   190
            Top             =   120
            Width           =   495
         End
         Begin VB.OptionButton optS2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   840
            TabIndex        =   189
            Top             =   120
            Width           =   495
         End
      End
      Begin VB.TextBox txtW10 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   4650
         TabIndex        =   202
         Top             =   4980
         Width           =   1185
      End
      Begin VB.TextBox txtW9 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   4650
         TabIndex        =   201
         Top             =   4680
         Width           =   1185
      End
      Begin VB.TextBox txtW8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   4650
         TabIndex        =   200
         Top             =   4380
         Width           =   1185
      End
      Begin VB.TextBox txtW7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   4650
         TabIndex        =   199
         Top             =   4080
         Width           =   1185
      End
      Begin VB.TextBox Text26 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4650
         TabIndex        =   124
         TabStop         =   0   'False
         Text            =   "Weight"
         Top             =   790
         Width           =   1185
      End
      Begin VB.Label Label16 
         Caption         =   "Land FM"
         Height          =   255
         Left            =   9360
         TabIndex        =   183
         Top             =   3480
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label15 
         Caption         =   "Land Fuel"
         Height          =   255
         Left            =   9360
         TabIndex        =   181
         Top             =   3000
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label14 
         Caption         =   "Land MOM"
         Height          =   375
         Left            =   9360
         TabIndex        =   180
         Top             =   3840
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label13 
         Caption         =   "Land ARM"
         Height          =   375
         Left            =   9360
         TabIndex        =   178
         Top             =   4200
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label52 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "*Enter the Holds freight in Kgs. Rest of the datums in lbs."
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
         Left            =   120
         TabIndex        =   106
         Top             =   8040
         Visible         =   0   'False
         Width           =   4695
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "AC Reg"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   33
         Top             =   2640
         Width           =   855
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   32
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Dest"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   28
         Top             =   2160
         Width           =   765
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Flt No"
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
         Left            =   0
         TabIndex        =   27
         Top             =   720
         Width           =   885
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
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
         Height          =   255
         Left            =   0
         TabIndex        =   26
         Top             =   1680
         Width           =   885
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CESSNA CARAVAN WEIGHT and BALANCE / LOADPLAN"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   385
         Left            =   360
         TabIndex        =   25
         Top             =   0
         Width           =   10545
      End
      Begin VB.Image Image2 
         Height          =   3885
         Left            =   360
         Picture         =   "frm2.frx":0000
         Stretch         =   -1  'True
         Top             =   8280
         Visible         =   0   'False
         Width           =   10080
      End
      Begin VB.Image Image1 
         Height          =   3165
         Left            =   360
         Picture         =   "frm2.frx":9EC9
         Stretch         =   -1  'True
         Top             =   9000
         Visible         =   0   'False
         Width           =   10080
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
      TabIndex        =   16
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "frm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dteEndDate As Date, strOrig As String, strDest As String, strVia As String
Dim intAP As Double, intUnder As Double, intLMCWt1 As Double, intLMCWt2 As Double
Dim intA As Double, intB As Double, intC As Double, intD As Double, intE As Double, intF As Double, intBOW As Double
Dim intCheck As Integer, intCheck1 As Integer

'Private Sub cmbAS1_Click()
'    txtAS1 = cmbAS1
'    txtLMCWt1.SetFocus
'End Sub

'Private Sub cmbAS2_Click()
'    txtAS2 = cmbAS2
'    txtLMCWt2.SetFocus
'End Sub

Private Sub cmbHold1_Click()
    txtDest1 = txtDest
    txtComp1 = cmbHold1
    'cmbAS1.SetFocus
End Sub

Private Sub cmbHold2_Click()
    txtDest2 = txtDest
    txtComp2 = cmbHold2
    'cmbAS2.SetFocus
End Sub

Private Sub cmdBack_Click()

    frmMain.Show
    frmMain.txtTOWt = ""
    frmMain.Show
    frmMain.txtZone1.SetFocus
    frmMain.txtZone2.SetFocus
    frmMain.txtZone3.SetFocus
    frmMain.txtZone4.SetFocus
    frmMain.txtZone5.SetFocus
        
    If frmMain.txtZone6.Visible = True Then
        frmMain.txtZone6.SetFocus
    End If
    
    If frmMain.txtZoneA.Visible = True Then
        frmMain.txtZoneA.SetFocus
    End If
    
    If frmMain.txtZoneB.Visible = True Then
        frmMain.txtZoneB.SetFocus
    End If
    
    If frmMain.txtZoneC.Visible = True Then
        frmMain.txtZoneC.SetFocus
    End If
    
    If frmMain.txtZoneD.Visible = True Then
        frmMain.txtZoneD.SetFocus
    End If
    
    If frmMain.txtSeat11.Visible = True Then
        frmMain.txtSeat11.SetFocus
        frmMain.txtSeat11.TabIndex = 19
    End If
    
    frmMain.txtZone1.SetFocus
    frm2.Hide
    
End Sub

Private Sub cmdClear_Click()

    If frmMain.Frame1.Caption = "CARGO in Kilograms" Then
        txtW1 = frmMain.txtZone1
        intK1 = Val(txtW1)
        
        txtW2 = frmMain.txtZone2
        intK2 = Val(txtW2)
        
        txtW3 = frmMain.txtZone3
        intK3 = Val(txtW3)
        
        txtW4 = frmMain.txtZone4
        intK4 = Val(txtW4)
        
        txtW5 = frmMain.txtZone5
        intK5 = Val(txtW5)
        
        txtW6 = frmMain.txtZone6
        intK6 = Val(txtW6)
        
        txtW7 = frmMain.txtZoneA
        intK7 = Val(txtW7)
        
        txtW8 = frmMain.txtZoneB
        intK8 = Val(txtW8)
        
        txtW9 = frmMain.txtZoneC
        intK9 = Val(txtW9)
        
        txtW10 = frmMain.txtZoneD
        intK10 = Val(txtW10)
        
        txtW11 = frmMain.txtSeat11
        intK11 = Val(txtW11)
        
    Else
        txtW1 = frmMain.txtZone1
        intP1 = Val(txtW1)
        intK1 = intP1 / 2.20458553791887
        txtW1 = intK1
        
        txtW2 = frmMain.txtZone2
        intP2 = Val(txtW2)
        intK2 = intP2 / 2.20458553791887
        txtW2 = intK2
        
        txtW3 = frmMain.txtZone3
        intP3 = Val(txtW3)
        intK3 = intP3 / 2.20458553791887
        txtW3 = intK3
        
        txtW4 = frmMain.txtZone4
        intP4 = Val(txtW4)
        intK4 = intP4 / 2.20458553791887
        txtW4 = intK4
        
        txtW5 = frmMain.txtZone5
        intP5 = Val(txtW5)
        intK5 = intP5 / 2.20458553791887
        txtW5 = intK5
        
        txtW6 = frmMain.txtZone6
        intP6 = Val(txtW6)
        intK6 = intP6 / 2.20458553791887
        txtW6 = intK6
        
        txtW7 = frmMain.txtZoneA
        intP7 = Val(txtW7)
        intK7 = intP7 / 2.20458553791887
        txtW7 = intK7
        
        txtW8 = frmMain.txtZoneB
        intP8 = Val(txtW8)
        intK8 = intP8 / 2.20458553791887
        txtW8 = intK8
        
        txtW9 = frmMain.txtZoneC
        intP9 = Val(txtW9)
        intK9 = intP9 / 2.20458553791887
        txtW9 = intK9
        
        txtW10 = frmMain.txtZoneD
        intP10 = Val(txtW10)
        intK10 = intP10 / 2.20458553791887
        txtW10 = intK10
        
        txtW11 = frmMain.txtSeat11
        intP11 = Val(txtW11)
        intK11 = intP11 / 2.20458553791887
        txtW11 = intK11
        
    End If
    
    'txtCrewW = intCrew
    'txtCrewW = Format(txtCrewW.Text, "Fixed")
    'txtXCrewW = "0.00"
    txtDest1 = ""
    cmbHold1 = ""
    txtLMCWt1 = ""
    txtDest2 = ""
    cmbHold2 = ""
    txtLMCWt2 = ""
    intTOWt = Val(frmMain.txtTOWt)
    Call Calc
    cmbHold1.SetFocus
    
End Sub

Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdPrint_Click()
                                                                                                        'If txtReg1 <> "208 Cargo Master" And txtReg1 <> "208 Grand Caravan" And txtReg1 <> "208 Amphib" And txtReg1 <> "208 Caravan" Then
                                                                                                        '    End
                                                                                                        'Else
    If frmMain.txtSeats = "208BCM" Then
        frmPrint2.Show
    ElseIf frmMain.txtSeats = "208B10C" Then
        frmPrint3.Show
    ElseIf frmMain.txtSeats = "208B10U" Then
        frmPrint3.Show
    ElseIf frmMain.txtSeats = "208B11C" Then
            frmPrint6.Show
    ElseIf frmMain.txtSeats = "208B12" Then
        frmPrint3.Show
    ElseIf frmMain.txtSeats = "208B13" Then
        frmPrint3.Show
    ElseIf frmMain.txtSeats = "208B14" Then
        frmPrint3.Show
    ElseIf frmMain.txtSeats = "208Amphib" Then
        frmPrint4.Show
    ElseIf frmMain.txtSeats = "208AN" Then
        frmPrint5.Show
    ElseIf frmMain.txtSeats = "208A10U" Then
        frmPrint5.Show
    ElseIf frmMain.txtSeats = "208A10C" Then
        frmPrint5.Show
    End If
    
    frm2.Hide
                                                                                                        'End If
End Sub

Private Sub Form_Activate()

    txtDate = Date
    txtPIC = frmMain.txtPIC
    txtPIC1 = frmMain.txtPIC
    intA = -23.61702
    intB = 205.65724
    intC = -0.00084699097
    
    Text33 = frmMain.Label2
    Text34 = frmMain.Label4
    Text35 = frmMain.Label5
    Text36 = frmMain.Label6
    Text40 = frmMain.Label7
    Text41 = frmMain.Label8
    
    If frmMain.txtSeats = "208B14" Then
        
        Text33.FontSize = 8
        Text34.FontSize = 8
        Text35.FontSize = 8
        Text36.FontSize = 7.5
        Text40.FontSize = 8
        Text41.FontSize = 8
        Text41 = ""
        Text33.Font = "Arial Narrow"
        Text34.Font = "Arial Narrow"
        Text35.Font = "Arial Narrow"
        Text36.Font = "Arial Narrow"
        Text40.Font = "Arial Narrow"
        Text41.Font = "Arial Narrow"
        txtA5 = 344#
        txtA6 = ""
        
    End If
    
    cmbHold1.Clear
    cmbHold2.Clear
    
    If frmMain.txtSeats = "208CM" Then

        cmbHold1.AddItem "Hold 1"
        cmbHold1.AddItem "Hold 2"
        cmbHold1.AddItem "Hold 3"
        cmbHold1.AddItem "Hold 4"
        cmbHold1.AddItem "Hold 5"
        cmbHold1.AddItem "Hold 6"
        cmbHold1.AddItem "Pod A"
        cmbHold1.AddItem "Pod B"
        cmbHold1.AddItem "Pod C"
        cmbHold1.AddItem "Pod D"
        cmbHold2.AddItem "Hold 1"
        cmbHold2.AddItem "Hold 2"
        cmbHold2.AddItem "Hold 3"
        cmbHold2.AddItem "Hold 4"
        cmbHold2.AddItem "Hold 5"
        cmbHold2.AddItem "Hold 6"
        cmbHold2.AddItem "Pod A"
        cmbHold2.AddItem "Pod B"
        cmbHold2.AddItem "Pod C"
        cmbHold2.AddItem "Pod D"
         
    ElseIf frmMain.txtSeats = "208B11C" Then

        cmbHold1.AddItem "Seat 3"
        cmbHold1.AddItem "Seats 4/5"
        cmbHold1.AddItem "Seat 6"
        cmbHold1.AddItem "Seats 7/8"
        cmbHold1.AddItem "Seats 9/10"
        cmbHold1.AddItem "Seat 11"
        cmbHold1.AddItem "Baggage"
        cmbHold1.AddItem "Pod A"
        cmbHold1.AddItem "Pod B"
        cmbHold1.AddItem "Pod C"
        cmbHold1.AddItem "Pod D"

        cmbHold2.AddItem "Seat 3"
        cmbHold2.AddItem "Seats 4/5"
        cmbHold2.AddItem "Seat 6"
        cmbHold2.AddItem "Seats 7/8"
        cmbHold2.AddItem "Seats 9/10"
        cmbHold2.AddItem "Seat 11"
        cmbHold2.AddItem "Baggage"
        cmbHold2.AddItem "Pod A"
        cmbHold2.AddItem "Pod B"
        cmbHold2.AddItem "Pod C"
        cmbHold2.AddItem "Pod D"

    ElseIf frmMain.txtSeats = "208B10C" Then
    
        cmbHold1.AddItem "Seats 3/4"
        cmbHold1.AddItem "Seats 5/6"
        cmbHold1.AddItem "Seats 7/8"
        cmbHold1.AddItem "Seats 9/10"
        cmbHold1.AddItem "Baggage"
        cmbHold1.AddItem "Pod A"
        cmbHold1.AddItem "Pod B"
        cmbHold1.AddItem "Pod C"
        cmbHold1.AddItem "Pod D"

        cmbHold2.AddItem "Seats 3/4"
        cmbHold2.AddItem "Seats 5/6"
        cmbHold2.AddItem "Seats 7/8"
        cmbHold2.AddItem "Seats 9/10"
        cmbHold2.AddItem "Baggage"
        cmbHold2.AddItem "Pod A"
        cmbHold2.AddItem "Pod B"
        cmbHold2.AddItem "Pod C"
        cmbHold2.AddItem "Pod D"

    ElseIf frmMain.txtSeats = "208B10U" Then
    
        cmbHold1.AddItem "Seats 3/4"
        cmbHold1.AddItem "Seats 5/6"
        cmbHold1.AddItem "Seats 7/8"
        cmbHold1.AddItem "Seats 9/10"
        cmbHold1.AddItem "Baggage"
        cmbHold1.AddItem "Pod A"
        cmbHold1.AddItem "Pod B"
        cmbHold1.AddItem "Pod C"
        cmbHold1.AddItem "Pod D"

        cmbHold2.AddItem "Seats 3/4"
        cmbHold2.AddItem "Seats 5/6"
        cmbHold2.AddItem "Seats 7/8"
        cmbHold2.AddItem "Seats 9/10"
        cmbHold2.AddItem "Baggage"
        cmbHold2.AddItem "Pod A"
        cmbHold2.AddItem "Pod B"
        cmbHold2.AddItem "Pod C"
        cmbHold2.AddItem "Pod D"

    ElseIf frmMain.txtSeats = "208B12" Then
    
        cmbHold1.AddItem "Seats 3/4"
        cmbHold1.AddItem "Seats 5/6"
        cmbHold1.AddItem "Seats 7/8"
        cmbHold1.AddItem "Seats 9/10"
        cmbHold1.AddItem "Seats 11/12"
        cmbHold1.AddItem "Baggage"
        cmbHold1.AddItem "Pod A"
        cmbHold1.AddItem "Pod B"
        cmbHold1.AddItem "Pod C"
        cmbHold1.AddItem "Pod D"

        cmbHold2.AddItem "Seats 3/4"
        cmbHold2.AddItem "Seats 5/6"
        cmbHold2.AddItem "Seats 7/8"
        cmbHold2.AddItem "Seats 9/10"
        cmbHold2.AddItem "Seats 11/12"
        cmbHold2.AddItem "Baggage"
        cmbHold2.AddItem "Pod A"
        cmbHold2.AddItem "Pod B"
        cmbHold2.AddItem "Pod C"
        cmbHold2.AddItem "Pod D"
                
    ElseIf frmMain.txtSeats = "208B13" Then
    
        cmbHold1.AddItem "Seats 3/4"
        cmbHold1.AddItem "Seats 5/6"
        cmbHold1.AddItem "Seats 7/8"
        cmbHold1.AddItem "Seats 9/10"
        cmbHold1.AddItem "Seats 11/12/13"
        cmbHold1.AddItem "Pod A"
        cmbHold1.AddItem "Pod B"
        cmbHold1.AddItem "Pod C"
        cmbHold1.AddItem "Pod D"

        cmbHold2.AddItem "Seats 3/4"
        cmbHold2.AddItem "Seats 5/6"
        cmbHold2.AddItem "Seats 7/8"
        cmbHold2.AddItem "Seats 9/10"
        cmbHold2.AddItem "Seats 11/12/13"
        cmbHold2.AddItem "Pod A"
        cmbHold2.AddItem "Pod B"
        cmbHold2.AddItem "Pod C"
        cmbHold2.AddItem "Pod D"
                
    ElseIf frmMain.txtSeats = "208B14" Then
    
        cmbHold1.AddItem "Seats 3/4/5"
        cmbHold1.AddItem "Seats 6/7/8"
        cmbHold1.AddItem "Seats 9/10/11"
        cmbHold1.AddItem "Seats 12/13/14"
        cmbHold1.AddItem "Baggage"
        cmbHold1.AddItem "Pod A"
        cmbHold1.AddItem "Pod B"
        cmbHold1.AddItem "Pod C"
        cmbHold1.AddItem "Pod D"

        cmbHold2.AddItem "Seats 3/4/5"
        cmbHold2.AddItem "Seats 6/7/8"
        cmbHold2.AddItem "Seats 9/10/11"
        cmbHold2.AddItem "Seats 12/13/14"
        cmbHold2.AddItem "Baggage"
        cmbHold2.AddItem "Pod A"
        cmbHold2.AddItem "Pod B"
        cmbHold2.AddItem "Pod C"
        cmbHold2.AddItem "Pod D"
       
    End If
    
    If frmMain.txtSeats = "208CM" Then
        Label7.Caption = "CESSNA 208 Cargo Master WEIGHT and BALANCE/LOADPLAN"
    ElseIf frmMain.txtSeats = "208B11C" Then
        Label7.Caption = "CESSNA 208 Grand Caravan WEIGHT and BALANCE/LOADPLAN"
    ElseIf frmMain.txtSeats = "208B10C" Then
        Label7.Caption = "CESSNA 208 Grand Caravan WEIGHT and BALANCE/LOADPLAN"
    ElseIf frmMain.txtSeats = "208B10U" Then
        Label7.Caption = "CESSNA 208 Grand Caravan WEIGHT and BALANCE/LOADPLAN"
    ElseIf frmMain.txtSeats = "208B12" Then
        Label7.Caption = "CESSNA 208 Grand Caravan WEIGHT and BALANCE/LOADPLAN"
    ElseIf frmMain.txtSeats = "208B13" Then
        Label7.Caption = "CESSNA 208 Grand Caravan WEIGHT and BALANCE/LOADPLAN"
    ElseIf frmMain.txtSeats = "208B14" Then
        Label7.Caption = "CESSNA 208 Grand Caravan WEIGHT and BALANCE/LOADPLAN"
    ElseIf frmMain.txtSeats = "208Amphib" Then
        Label7.Caption = "CESSNA 208 AMPHIB WEIGHT and BALANCE/LOADPLAN"
    ElseIf frmMain.txtSeats = "208AN" Then
        Label7.Caption = "CESSNA 208A Caravan WEIGHT and BALANCE/LOADPLAN"
    ElseIf frmMain.txtSeats = "208A10U" Then
        Label7.Caption = "CESSNA 208A Caravan WEIGHT and BALANCE/LOADPLAN"
    ElseIf frmMain.txtSeats = "20810C" Then
        Label7.Caption = "CESSNA 208A Caravan WEIGHT and BALANCE/LOADPLAN"
    End If
    
    If frmMain.Frame1.Caption = "CARGO in Kilograms" Then
        
        txtZone1 = intK1
        txtZone1 = Format(txtZone1.Text, "Fixed")
        txtZone2 = intK2
        txtZone2 = Format(txtZone2.Text, "Fixed")
        txtZone3 = intK3
        txtZone3 = Format(txtZone3.Text, "Fixed")
        txtZone4 = intK4
        txtZone4 = Format(txtZone4.Text, "Fixed")
        txtZone5 = intK5
        txtZone5 = Format(txtZone5.Text, "Fixed")
        txtZone6 = intK6
        txtZone6 = Format(txtZone6.Text, "Fixed")
        txtZoneA = intK7
        txtZoneA = Format(txtZoneA.Text, "Fixed")
        txtZoneB = intK8
        txtZoneB = Format(txtZoneB.Text, "Fixed")
        txtZoneC = intK9
        txtZoneC = Format(txtZoneC.Text, "Fixed")
        txtZoneD = intK10
        txtZoneD = Format(txtZoneD.Text, "Fixed")
        txtW1 = intK1
        txtW1 = Format(txtW1.Text, "Fixed")
        txtW2 = intK2
        txtW2 = Format(txtW2.Text, "Fixed")
        txtW3 = intK3
        txtW3 = Format(txtW3.Text, "Fixed")
        txtW4 = intK4
        txtW4 = Format(txtW4.Text, "Fixed")
        txtW5 = intK5
        txtW5 = Format(txtW5.Text, "Fixed")
        txtW6 = intK6
        txtW6 = Format(txtW6.Text, "Fixed")
        txtW7 = intK7
        txtW7 = Format(txtW7.Text, "Fixed")
        txtW8 = intK8
        txtW8 = Format(txtW8.Text, "Fixed")
        txtW9 = intK9
        txtW9 = Format(txtW9.Text, "Fixed")
        txtW10 = intK10
        txtW10 = Format(txtW10.Text, "Fixed")
        Text50 = "Weight in Kilograms"
        Text52 = "All these weights in Kilograms"
        Text54 = "1542.00"
        
        If frmMain.cmbAC = "208 Cargo Master" Or frmMain.cmbAC = "208 Grand Caravan" Then
            txtMax1 = "188 kgs"
            txtMax2 = "390 kgs"
            txtMax3 = "224 kgs"
            txtMax4 = "154 kgs"
            txtMax5 = "142 kgs"
            txtMax6 = "111 kgs"
            txtMax10 = "104 kgs"
            txtMax11 = "140 kgs"
            txtMax12 = "122 kgs"
            txtMax13 = "127 kgs"
        Else
            txtMax1 = "807 kgs"
            txtMax2 = "1406 kgs"
            txtMax3 = "862 kgs"
            txtMax4 = "626 kgs"
            txtMax5 = "575 kgs"
            txtMax6 = "145 kgs"
            txtMax10 = ""
            txtMax11 = ""
            txtMax12 = ""
            txtMax13 = ""
        End If
        
        Label52.Visible = True
        
    Else
    
        txtZone1 = intP1
        txtZone1 = Format(txtZone1.Text, "Fixed")
        txtZone2 = intP2
        txtZone2 = Format(txtZone2.Text, "Fixed")
        txtZone3 = intP3
        txtZone3 = Format(txtZone3.Text, "Fixed")
        txtZone4 = intP4
        txtZone4 = Format(txtZone4.Text, "Fixed")
        txtZone5 = intP5
        txtZone5 = Format(txtZone5.Text, "Fixed")
        txtZone6 = intP6
        txtZone6 = Format(txtZone6.Text, "Fixed")
        txtZoneA = intP7
        txtZoneA = Format(txtZoneA.Text, "Fixed")
        txtZoneB = intP8
        txtZoneB = Format(txtZoneB.Text, "Fixed")
        txtZoneC = intP9
        txtZoneC = Format(txtZoneC.Text, "Fixed")
        txtZoneD = intP10
        txtZoneD = Format(txtZoneD.Text, "Fixed")
        txtW1 = intP1
        txtW1 = Format(txtW1.Text, "Fixed")
        txtW2 = intP2
        txtW2 = Format(txtW2.Text, "Fixed")
        txtW3 = intP3
        txtW3 = Format(txtW3.Text, "Fixed")
        txtW4 = intP4
        txtW4 = Format(txtW4.Text, "Fixed")
        txtW5 = intP5
        txtW5 = Format(txtW5.Text, "Fixed")
        txtW6 = intP6
        txtW6 = Format(txtW6.Text, "Fixed")
        txtW7 = intP7
        txtW7 = Format(txtW7.Text, "Fixed")
        txtW8 = intP8
        txtW8 = Format(txtW8.Text, "Fixed")
        txtW9 = intP9
        txtW9 = Format(txtW9.Text, "Fixed")
        txtW10 = intP10
        txtW10 = Format(txtW10.Text, "Fixed")
        Text52 = "All these weights in Pounds"
        Text54 = "3400.00"
        
        If frmMain.cmbAC = "208 Cargo Master" Or frmMain.cmbAC = "208 Grand Caravan" Then
            txtMax1 = "415 lbs"
            txtMax2 = "860 lbs"
            txtMax3 = "495 lbs"
            txtMax4 = "340 lbs"
            txtMax5 = "315 lbs"
            txtMax6 = "245 lbs"
            txtMax10 = "230 lbs"
            txtMax11 = "310 lbs"
            txtMax12 = "270 lbs"
            txtMax13 = "280 lbs"
        Else
            txtMax1 = "1780 lbs"
            txtMax2 = "3100 lbs"
            txtMax3 = "1400 lbs"
            txtMax4 = "1380 lbs"
            txtMax5 = "1270 lbs"
            txtMax6 = "320 lbs"
            txtMax10 = ""
            txtMax11 = ""
            txtMax12 = ""
            txtMax13 = ""
            txtW7.BackColor = &HC0C0C0
            txtW8.BackColor = &HC0C0C0
            txtW9.BackColor = &HC0C0C0
            txtW10.BackColor = &HC0C0C0
            txtA7.BackColor = &HC0C0C0
            txtA8.BackColor = &HC0C0C0
            txtA9.BackColor = &HC0C0C0
            txtA10.BackColor = &HC0C0C0
            txtM7.BackColor = &HC0C0C0
            txtM8.BackColor = &HC0C0C0
            txtM9.BackColor = &HC0C0C0
            txtM10.BackColor = &HC0C0C0
            txtMax10.BackColor = &HC0C0C0
            txtMax11.BackColor = &HC0C0C0
            txtMax12.BackColor = &HC0C0C0
            txtMax13.BackColor = &HC0C0C0
            txtW7.ForeColor = &HFFFFFF
            txtW8.ForeColor = &HFFFFFF
            txtW9.ForeColor = &HFFFFFF
            txtW10.ForeColor = &HFFFFFF
            txtA7.ForeColor = &HFFFFFF
            txtA8.ForeColor = &HFFFFFF
            txtA9.ForeColor = &HFFFFFF
            txtA10.ForeColor = &HFFFFFF
            txtM7.ForeColor = &HFFFFFF
            txtM8.ForeColor = &HFFFFFF
            txtM9.ForeColor = &HFFFFFF
            txtM10.ForeColor = &HFFFFFF
            txtMax10.ForeColor = &HFFFFFF
            txtMax11.ForeColor = &HFFFFFF
            txtMax12.ForeColor = &HFFFFFF
            txtMax13.ForeColor = &HFFFFFF

            txtW7.Text = ""
            txtW8.Text = ""
            txtW9.Text = ""
            txtW10.Text = ""
            txtA7.Text = ""
            txtA8.Text = ""
            txtA9.Text = ""
            txtA10.Text = ""
            txtM7.Text = ""
            txtM8.Text = ""
            txtM9.Text = ""
            txtM10.Text = ""
            txtMax10.Text = ""
            txtMax11.Text = ""
            txtMax12.Text = ""
            txtMax13.Text = ""
            intWt13 = 0
            intWt14 = 0
            intWt15 = 0
            intWt16 = 0
            intArm13 = 0
            intArm14 = 0
            intArm15 = 0
            intArm16 = 0
            intMom13 = 0
            intMom14 = 0
            intMom15 = 0
            intMom16 = 0
            
        End If
        
        Label52.Visible = False
        Text50 = "Weight in Pounds"
        
    End If
        
    txtReg = frmMain.cmbAC.Text
    txtReg1 = frmMain.cmbAC.Text
                                                                                                        'If txtReg1 <> "208 Cargo Master" And txtReg1 <> "208 Grand Caravan" And txtReg1 <> "208 Amphib" And txtReg1 <> "208 Caravan" Then
                                                                                                        '    End
                                                                                                        'End If
    txtOrig = frmMain.txtOrig.Text
    txtDest = frmMain.txtDest.Text
    txtDest1 = frmMain.txtDest.Text
    txtDest2 = frmMain.txtDest.Text
    txtFltNo = frmMain.txtFltNo.Text
    txtPayloadW = intCargoP
    txtPayloadW = Format(txtPayloadW.Text, "Fixed")
    If intRetVal = 6 Then
        intCrew = 0
        intRetVal = 0
    End If
    txtCrewW = intCrew
    txtCrewW = Format(txtCrewW.Text, "Fixed")
    
    If frmMain.Frame1.Caption = "CARGO in Kilograms" Then
        txtTotalPay = intCargoK
    Else
        txtTotalPay = intCargoP
    End If
    txtTotalPay = Format(txtTotalPay.Text, "Fixed")
    txtFuelW = intFOB - 35
    txtFuelW = Format(txtFuelW.Text, "Fixed")
    txtTripW = intBurn
    txtTripW = Format(txtTripW.Text, "Fixed")
    txtBWW = intBEW
    txtBWW = Format(txtBWW.Text, "Fixed")
    txtBEWW = intBEW
    txtBEWW = Format(txtBEWW.Text, "Fixed")
    txtBOWW = intBEW + Val(txtCrewW) + Val(txtXCrewW)
    intBOW = Val(txtBOWW)
    txtPayloadM = (8750 - intBOW)
    If Val(txtPayloadM) > 3400 Then
        txtPayloadM = "3400"
    End If
    'txtPayloadM = Format(txtPayloadM.Text, "Fixed")
    txtBOWW = Format(txtBOWW.Text, "Fixed")
    txtZFWW = intCargoP + Val(txtBOWW)
    txtZFWW = Format(txtZFWW.Text, "Fixed")
    intRampFuel = intFOB
    txtRampFuelW = intRampFuel
    txtRampFuelW = Format(txtRampFuelW.Text, "Fixed")
    intZFW = Val(txtZFWW)
    txtZEWW1 = intZFW
    txtZEWW1 = Format(txtZEWW1.Text, "Fixed")
    txtRampWW = intZFW + intFOB
    txtRampWW = Format(txtRampWW.Text, "Fixed")
    intRampWt = Val(txtRampWW)
    txtTaxiW = intTaxi
    txtTaxiW = Format(txtTaxiW.Text, "Fixed")
    intTOWt = intRampWt - intTaxi
    txtTOWW = intTOWt
    txtTOWW = Format(txtTOWW.Text, "Fixed")
    txtTOWW1 = intTOWt
    txtTOWW1 = Format(txtTOWW1.Text, "Fixed")
    txtTripW = intBurn
    txtTripW = Format(txtTripW.Text, "Fixed")
    intLandWt = intTOWt - intBurn
    txtLWW = intLandWt
    txtLWW = Format(txtLWW.Text, "Fixed")
    
    If frmMain.Frame1.Caption = "CARGO in Kilograms" Then
        intAP = 1542
        intUnder = intAP - intCargoK
        txtUnderLoad.Text = intUnder & " kgs"
    Else
        intAP = 3400
        intUnder = intAP - intCargoP
        txtUnderLoad.Text = intUnder & " lbs"
    End If
    txtUnderLoad1 = txtUnderLoad
    
    intWt1 = intBEW
    intWt2 = Val(txtCrewW1)
    intWt3 = Val(txtXtraW)
    intWt4 = intP1
    intWt5 = intP2
    intWt6 = intP3
    intWt7 = intP4
    intWt8 = intP5
    intWt9 = intP6
    intWt13 = intP7
    intWt14 = intP8
    intWt15 = intP9
    intWt16 = intP10
    intWt10 = intZFW
    intWt11 = intFOB
    intWt12 = intTOWt
    'intArm1 = 183.64
    
    If frmMain.cmbAC = "208 Cargo Master" Or frmMain.cmbAC = "208 Grand Caravan" Then

        intArm2 = 135.5
        intArm3 = 172
        intArm4 = 172
        intArm5 = 217.8
        intArm6 = 264.4
        intArm7 = 294.5
        intArm8 = 319.5
        intArm9 = 344
        'intArm11 = 196.54
        intArm13 = 132.4
        intArm14 = 182.1
        intArm15 = 233.4
        intArm16 = 287.6
        
    Else
    
        intArm2 = 135.5
        intArm3 = 168.4
        intArm4 = 168.4
        intArm5 = 194.8
        intArm6 = 221
        intArm7 = 246.5
        intArm8 = 271.5
        intArm9 = 296
        'intArm11 = 196.54
        intArm13 = 0
        intArm14 = 0
        intArm15 = 0
        intArm16 = 0
        
    End If

    'intMom1 = CLng(intWt1) * intArm1
    intMom2 = intWt2 * intArm2
    intMom3 = intWt3 * intArm3
    intMom4 = intWt4 * intArm4
    intMom5 = intWt5 * intArm5
    intMom6 = intWt6 * intArm6
    intMom7 = intWt7 * intArm7
    intMom8 = intWt8 * intArm8
    intMom9 = intWt9 * intArm9
    intMom13 = intWt13 * intArm13
    intMom14 = intWt14 * intArm14
    intMom15 = intWt15 * intArm15
    intMom16 = intWt16 * intArm16
    intTotal = intMom1 + intMom2 + intMom3 + intMom4 + intMom5 + intMom6 + intMom7 + intMom8 + intMom9 + intMom13 + intMom14 + intMom15 + intMom16
    intMom10 = intTotal
    intArm10 = intTotal / intWt10
    intX2 = intFOB * intFOB
    intMom11 = intA + (intB * intFOB) + (intC * intX2)
    intMom12 = intTotal + intMom11
    intArm12 = intMom12 / intWt12

    txtBEWA = intArm1
    txtBEWA = Format(txtBEWA.Text, "Fixed")
    txtCrewA = intArm2
    txtCrewA = Format(txtCrewA.Text, "Fixed")
    txtXtraA = intArm3
    txtXtraA = Format(txtXtraA.Text, "Fixed")
    txtA1 = intArm4
    txtA1 = Format(txtA1.Text, "Fixed")
    txtA2 = intArm5
    txtA2 = Format(txtA2.Text, "Fixed")
    txtA3 = intArm6
    txtA3 = Format(txtA3.Text, "Fixed")
    txtA4 = intArm7
    txtA4 = Format(txtA4.Text, "Fixed")
    txtA5 = intArm8
    txtA5 = Format(txtA5.Text, "Fixed")
    txtA6 = intArm9
    txtA6 = Format(txtA6.Text, "Fixed")
    txtA7 = intArm13
    txtA7 = Format(txtA7.Text, "Fixed")
    txtA8 = intArm14
    txtA8 = Format(txtA8.Text, "Fixed")
    txtA9 = intArm15
    txtA9 = Format(txtA9.Text, "Fixed")
    txtA10 = intArm16
    txtA10 = Format(txtA10.Text, "Fixed")
    txtZEWA = intArm10
    txtZEWA = Format(txtZEWA.Text, "Fixed")
    'txtFuelA = intArm11
    'txtFuelA = Format(txtFuelA.Text, "Fixed")
    txtTOWA = intArm12
    txtTOWA = Format(txtTOWA.Text, "Fixed")
    
    txtBEWM = intMom1
    txtBEWM = Format(txtBEWM.Text, "Fixed")
    txtCrewM1 = intMom2
    txtCrewM1 = Format(txtCrewM1.Text, "Fixed")
    txtXtraM = intMom3
    txtXtraM = Format(txtXtraM.Text, "Fixed")
    txtM1 = intMom4
    txtM1 = Format(txtM1.Text, "Fixed")
    txtM2 = intMom5
    txtM2 = Format(txtM2.Text, "Fixed")
    txtM3 = intMom6
    txtM3 = Format(txtM3.Text, "Fixed")
    txtM4 = intMom7
    txtM4 = Format(txtM4.Text, "Fixed")
    txtM5 = intMom8
    txtM5 = Format(txtM5.Text, "Fixed")
    txtM6 = intMom9
    txtM6 = Format(txtM6.Text, "Fixed")
    txtM7 = intMom13
    txtM7 = Format(txtM7.Text, "Fixed")
    txtM8 = intMom14
    txtM8 = Format(txtM8.Text, "Fixed")
    txtM9 = intMom15
    txtM9 = Format(txtM9.Text, "Fixed")
    txtM10 = intMom16
    txtM10 = Format(txtM10.Text, "Fixed")
    txtZEWM = intMom10
    txtZEWM = Format(txtZEWM.Text, "Fixed")
    txtFuelM = intMom11
    txtFuelM = Format(txtFuelM.Text, "Fixed")
    txtTOWM1 = intMom12
    txtTOWM1 = Format(txtTOWM1.Text, "Fixed")
    Call Calc
    
    If intCheck = 1 Then
            frmMain.Show
            Unload Me
    Else
        Call MsgBox("Please enter Crew Weight, Extra Crew Weight, and any Last Minute Changes.", vbOKOnly, "Crew Weight and LMC")
        txtCrewW.SetFocus
        txtCrewW.SelStart = 0
        txtCrewW.SelLength = Len(txtCrewW.Text)
    End If
    
    intCheck1 = 0

End Sub

Private Sub Calc()

    If frmMain.Frame1.Caption = "CARGO in Kilograms" Then
        
        txtZone1 = intK1
        txtZone1 = Format(txtZone1.Text, "Fixed")
        txtZone2 = intK2
        txtZone2 = Format(txtZone2.Text, "Fixed")
        txtZone3 = intK3
        txtZone3 = Format(txtZone3.Text, "Fixed")
        txtZone4 = intK4
        txtZone4 = Format(txtZone4.Text, "Fixed")
        txtZone5 = intK5
        txtZone5 = Format(txtZone5.Text, "Fixed")
        txtZone6 = intK6
        txtZone6 = Format(txtZone6.Text, "Fixed")
        txtZoneA = intK7
        txtZoneA = Format(txtZoneA.Text, "Fixed")
        txtZoneB = intK8
        txtZoneB = Format(txtZoneB.Text, "Fixed")
        txtZoneC = intK9
        txtZoneC = Format(txtZoneC.Text, "Fixed")
        txtZoneD = intK10
        txtZoneD = Format(txtZoneD.Text, "Fixed")
        txtW1 = intK1
        txtW1 = Format(txtW1.Text, "Fixed")
        txtW2 = intK2
        txtW2 = Format(txtW2.Text, "Fixed")
        txtW3 = intK3
        txtW3 = Format(txtW3.Text, "Fixed")
        txtW4 = intK4
        txtW4 = Format(txtW4.Text, "Fixed")
        txtW5 = intK5
        txtW5 = Format(txtW5.Text, "Fixed")
        txtW6 = intK6
        txtW6 = Format(txtW6.Text, "Fixed")
        txtW7 = intK7
        txtW7 = Format(txtW7.Text, "Fixed")
        txtW8 = intK8
        txtW8 = Format(txtW8.Text, "Fixed")
        txtW9 = intK9
        txtW9 = Format(txtW9.Text, "Fixed")
        txtW10 = intK10
        txtW10 = Format(txtW10.Text, "Fixed")
        intCargoK = intK1 + intK2 + intK3 + intK4 + intK5 + intK6 + intK7 + intK8 + intK9 + intK10
        intCargoP = (intK1 + intK2 + intK3 + intK4 + intK5 + intK6 + intK7 + intK8 + intK9 + intK10) * 2.20458553791887
        
    Else
    
        txtZone1 = intP1
        txtZone1 = Format(txtZone1.Text, "Fixed")
        txtZone2 = intP2
        txtZone2 = Format(txtZone2.Text, "Fixed")
        txtZone3 = intP3
        txtZone3 = Format(txtZone3.Text, "Fixed")
        txtZone4 = intP4
        txtZone4 = Format(txtZone4.Text, "Fixed")
        txtZone5 = intP5
        txtZone5 = Format(txtZone5.Text, "Fixed")
        txtZone6 = intP6
        txtZone6 = Format(txtZone6.Text, "Fixed")
        txtZoneA = intP7
        txtZoneA = Format(txtZoneA.Text, "Fixed")
        txtZoneB = intP8
        txtZoneB = Format(txtZoneB.Text, "Fixed")
        txtZoneC = intP9
        txtZoneC = Format(txtZoneC.Text, "Fixed")
        txtZoneD = intP10
        txtZoneD = Format(txtZoneD.Text, "Fixed")
        txtW1 = intP1
        txtW1 = Format(txtW1.Text, "Fixed")
        txtW2 = intP2
        txtW2 = Format(txtW2.Text, "Fixed")
        txtW3 = intP3
        txtW3 = Format(txtW3.Text, "Fixed")
        txtW4 = intP4
        txtW4 = Format(txtW4.Text, "Fixed")
        txtW5 = intP5
        txtW5 = Format(txtW5.Text, "Fixed")
        txtW6 = intP6
        txtW6 = Format(txtW6.Text, "Fixed")
        txtW7 = intP7
        txtW7 = Format(txtW7.Text, "Fixed")
        txtW8 = intP8
        txtW8 = Format(txtW8.Text, "Fixed")
        txtW9 = intP9
        txtW9 = Format(txtW9.Text, "Fixed")
        txtW10 = intP10
        txtW10 = Format(txtW10.Text, "Fixed")
        intCargoK = (intP1 + intP2 + intP3 + intP4 + intP5 + intP6 + intP7 + intP8 + intP9 + intP10 + intP11) / 2.20458553791887
        intCargoP = intP1 + intP2 + intP3 + intP4 + intP5 + intP6 + intP7 + intP8 + intP9 + intP10 + intP11
        
    End If
    
    
    If intCargoP > 3400 Then
        intCargoP = 0
        Call MsgBox("Weight Limits Exceeded.  Please re-enter acceptable values.", vbOKOnly, "Weight Limits Exceeded")
        txtLMCWt1 = ""
        txtLMCWt2 = ""
        txtLMCWt1.SetFocus
        'Call cmdClear_Click
        Exit Sub
    Else
        txtPayloadW = intCargoP
        txtPayloadW = Format(txtPayloadW.Text, "Fixed")
    
        If frmMain.Frame1.Caption = "CARGO in Kilograms" Then
            txtTotalPay = intCargoK
        Else
            txtTotalPay = intCargoP
        End If
        txtTotalPay = Format(txtTotalPay.Text, "Fixed")
        
        txtFuelW = intFOB - 35
        txtFuelW = Format(txtFuelW.Text, "Fixed")
        txtTripW = intBurn
        txtTripW = Format(txtTripW.Text, "Fixed")
        txtBWW = intBEW
        txtBWW = Format(txtBWW.Text, "Fixed")
        txtBEWW = intBEW
        txtBEWW = Format(txtBEWW.Text, "Fixed")
        txtBOWW = intBEW + Val(txtCrewW) + Val(txtXCrewW)
        txtBOWW = Format(txtBOWW.Text, "Fixed")
        intBOW = Val(txtBOWW)
        txtPayloadM = (8750 - intBOW)
        If Val(txtPayloadM) > 3400 Then
            txtPayloadM = "3400"
        End If
        'txtPayloadM = Format(txtPayloadM.Text, "Fixed")
        txtZFWW = intCargoP + Val(txtBOWW)
        txtZFWW = Format(txtZFWW.Text, "Fixed")
        intRampFuel = intFOB
        txtRampFuelW = intRampFuel
        txtRampFuelW = Format(txtRampFuelW.Text, "Fixed")
        intZFW = Val(txtZFWW)
        txtZEWW1 = intZFW
        txtZEWW1 = Format(txtZEWW1.Text, "Fixed")
        txtRampWW = intZFW + intFOB
        txtRampWW = Format(txtRampWW.Text, "Fixed")
        intRampWt = Val(txtRampWW)
        txtTaxiW = intTaxi
        txtTaxiW = Format(txtTaxiW.Text, "Fixed")
        intTOWt = intRampWt - intTaxi
        If intTOWt > 8750 Then
            'If intRetVal <> 6 Then
                'intRetVal = 0
                'intTOWt = 0
                Call MsgBox("Take Off Weight Limit Exceeded.  Please re-enter acceptable values.", vbOKOnly, "Weight Limits Exceeded")
                'txtTOWW = ""
                'txtTOWW = Format(txtTOWW.Text, "Fixed")
                Exit Sub
            'Else
                'intRetVal = 0
            'End If
        Else
            txtTOWW = intTOWt
            txtTOWW = Format(txtTOWW.Text, "Fixed")
            txtTOWW1 = intTOWt
            txtTOWW1 = Format(txtTOWW1.Text, "Fixed")
        End If
        
        txtTripW = intBurn
        txtTripW = Format(txtTripW.Text, "Fixed")
        intLandWt = intTOWt - intBurn
        txtLWW = intLandWt
        txtLWW = Format(txtLWW.Text, "Fixed")
        
        If Val(txtLWW) > 8500 Then
            intLandWt = 0
            Call MsgBox("Landing Weight Limit Exceeded.  Please re-enter acceptable values.", vbOKOnly, "Weight Limits Exceeded")
            txtLWW = ""
            Exit Sub
        End If
        
        intWt1 = intBEW
        intWt2 = Val(txtCrewW1)
        intWt3 = Val(txtXtraW)
        
        If frmMain.Frame1.Caption = "CARGO in Kilograms" Then
            intP1 = intK1 * 2.20458553791887
            intP2 = intK2 * 2.20458553791887
            intP3 = intK3 * 2.20458553791887
            intP4 = intK4 * 2.20458553791887
            intP5 = intK5 * 2.20458553791887
            intP6 = intK6 * 2.20458553791887
            intP7 = intK7 * 2.20458553791887
            intP8 = intK8 * 2.20458553791887
            intP9 = intK9 * 2.20458553791887
            intP10 = intK10 * 2.20458553791887
        End If
        
        intWt4 = intP1
        intWt5 = intP2
        intWt6 = intP3
        intWt7 = intP4
        intWt8 = intP5
        intWt9 = intP6
        intWt13 = intP7
        intWt14 = intP8
        intWt15 = intP9
        intWt16 = intP10
        intWt10 = intZFW
        intWt11 = intFOB
        intWt12 = intTOWt
        'intArm1 = 183.64
        
    If frmMain.txtSeats = "208BCM" Then

        intArm2 = 135.5
        intArm3 = 172
        intArm4 = 172
        intArm5 = 217.8
        intArm6 = 264.4
        intArm7 = 294.5
        intArm8 = 319.5
        intArm9 = 344
        'intArm11 = 196.54
        intArm13 = 132.4
        intArm14 = 182.1
        intArm15 = 233.4
        intArm16 = 287.6
        
    ElseIf frmMain.txtSeats = "208B11C" Then

        intArm2 = 135.5
        intArm3 = 0
        intArm4 = 189.9
        intArm5 = 173.9
        intArm6 = 225.9
        intArm7 = 209.9
        intArm8 = 245.9
        intArm17 = 344
        intArm9 = 261.9
        'intArm11 = 196.54
        intArm13 = 132.4
        intArm14 = 182.1
        intArm15 = 233.4
        intArm16 = 287.6
        Text75.BackColor = &H80000005
        Text75 = "Baggage"
        Text79.BackColor = &H80000005
        txtW11.BackColor = &H0&
        txtW11.ForeColor = &HFFFFFF
        txtW11.FontBold = True
        txtW11 = Val(frmMain.txtSeat11)
        txtW11 = Format(txtW11.Text, "Fixed")
        txtA11.BackColor = &H80000005
        txtM11.BackColor = &H80000005
        intWt17 = Val(txtW11)
        txtA11 = intArm17
        txtA11 = Format(txtA11.Text, "Fixed")
        intMom17 = intWt17 * intArm17
        txtM11 = intMom17
        txtM11 = Format(txtM11.Text, "Fixed")
        
        If frmMain.Frame1.Caption = "Cargo in Kilograms" Then
            intK11 = intWt17
            intP11 = intWt17 * 2.20458553791887
        Else
            intK11 = intWt17 / 2.20458553791887
            intP11 = intWt17
        End If
        
    ElseIf frmMain.txtSeats = "208B10C" Then

        intArm2 = 135.5
        intArm3 = 0
        intArm4 = 173.9
        intArm5 = 209.9
        intArm6 = 245.9
        intArm7 = 281.9
        intArm8 = 344
        intArm17 = 0
        intArm9 = 0
        'intArm11 = 196.54
        intArm13 = 132.4
        intArm14 = 182.1
        intArm15 = 233.4
        intArm16 = 287.6
        Text41 = ""
        Text41.BackColor = &HC0C0C0
        txtW6.BackColor = &HC0C0C0
        txtA6.BackColor = &HC0C0C0
        txtM6.BackColor = &HC0C0C0
        txtMax6.BackColor = &HC0C0C0
        txtMax6 = ""
        txtW11 = "0.00"
        txtW6.ForeColor = &H80000012
        txtW6.FontBold = False
        txtA11 = "0.00"
        txtM11 = "0.00"
        
    ElseIf frmMain.txtSeats = "208B10U" Then

        intArm2 = 135.5
        intArm3 = 0
        intArm4 = 170.5
        intArm5 = 206.5
        intArm6 = 242.5
        intArm7 = 278.5
        intArm8 = 344
        intArm17 = 0
        intArm9 = 0
        'intArm11 = 196.54
        intArm13 = 132.4
        intArm14 = 182.1
        intArm15 = 233.4
        intArm16 = 287.6
        Text41 = ""
        Text41.BackColor = &HC0C0C0
        txtW6.BackColor = &HC0C0C0
        txtA6.BackColor = &HC0C0C0
        txtM6.BackColor = &HC0C0C0
        txtMax6.BackColor = &HC0C0C0
        txtMax6 = ""
        txtW11 = "0.00"
        txtW6.ForeColor = &H80000012
        txtW6.FontBold = False
        txtA11 = "0.00"
        txtM11 = "0.00"
        
    ElseIf frmMain.txtSeats = "208B12" Then

        intArm2 = 135.5
        intArm3 = 0
        intArm4 = 170.5
        intArm5 = 200.5
        intArm6 = 230.5
        intArm7 = 260.5
        intArm8 = 290.5
        intArm17 = 0
        intArm9 = 344
        'intArm11 = 196.54
        intArm13 = 132.4
        intArm14 = 182.1
        intArm15 = 233.4
        intArm16 = 287.6
        
    ElseIf frmMain.txtSeats = "208B13" Then

        intArm2 = 135.5
        intArm3 = 0
        intArm4 = 173.9
        intArm5 = 209.9
        intArm6 = 245.9
        intArm7 = 281.9
        intArm8 = 342.4
        intArm17 = 0
        intArm9 = 0
        'intArm11 = 196.54
        intArm13 = 132.4
        intArm14 = 182.1
        intArm15 = 233.4
        intArm16 = 287.6
        
    ElseIf frmMain.txtSeats = "208B14" Then

        intArm2 = 135.5
        intArm3 = 0
        intArm4 = 173.9
        intArm5 = 209.9
        intArm6 = 245.9
        intArm7 = 281.9
        intArm8 = 344
        intArm17 = 0
        intArm9 = 0
        'intArm11 = 196.54
        intArm13 = 132.4
        intArm14 = 182.1
        intArm15 = 233.4
        intArm16 = 287.6
        Text41 = ""
        Text41.BackColor = &HC0C0C0
        txtW6.BackColor = &HC0C0C0
        txtA6.BackColor = &HC0C0C0
        txtM6.BackColor = &HC0C0C0
        txtMax6.BackColor = &HC0C0C0
        txtMax6 = ""
        txtW11 = "0.00"
        txtW6.ForeColor = &H80000012
        txtW6.FontBold = False
        txtA11 = "0.00"
        txtM11 = "0.00"
        
    ElseIf frmMain.txtSeats = "208AN" Then
    
        intArm2 = 135.5
        intArm3 = 168.4
        intArm4 = 168.4
        intArm5 = 194.8
        intArm6 = 221
        intArm7 = 246.5
        intArm8 = 271.5
        intArm9 = 296
        'intArm11 = 196.54
        intArm13 = 0
        intArm14 = 0
        intArm15 = 0
        intArm16 = 0
        
    ElseIf frmMain.txtSeats = "208A10U" Then
    
        intArm2 = 135.5
        intArm3 = 0
        intArm4 = 166.5
        intArm5 = 193.5
        intArm6 = 220.5
        intArm7 = 248.5
        intArm8 = 245.5
        intArm9 = 296
        'intArm11 = 196.54
        intArm13 = 0
        intArm14 = 0
        intArm15 = 0
        intArm16 = 0
        
    ElseIf frmMain.txtSeats = "208A10C" Then
    
        intArm2 = 135.5
        intArm3 = 0
        intArm4 = 185.9
        intArm5 = 169.9
        intArm6 = 217.9
        intArm7 = 201.9
        intArm8 = 233.9
        intArm9 = 296
        'intArm11 = 196.54
        intArm13 = 0
        intArm14 = 0
        intArm15 = 0
        intArm16 = 0
        
    End If

        intMom2 = intWt2 * intArm2
        intMom3 = intWt3 * intArm3
        intMom4 = intWt4 * intArm4
        intMom5 = intWt5 * intArm5
        intMom6 = intWt6 * intArm6
        intMom7 = intWt7 * intArm7
        intMom8 = intWt8 * intArm8
        intMom9 = intWt9 * intArm9
        intMom13 = intWt13 * intArm13
        intMom14 = intWt14 * intArm14
        intMom15 = intWt15 * intArm15
        intMom16 = intWt16 * intArm16
        intTotal = intMom1 + intMom2 + intMom3 + intMom4 + intMom5 + intMom6 + intMom7 + intMom8 + intMom9 + intMom13 + intMom14 + intMom15 + intMom16 + intMom17
        intMom10 = intTotal
        intArm10 = intTotal / intWt10
        intX2 = intLandFuel * intLandFuel
        intLandFuelMom = intA + (intB * intLandFuel) + (intC * intX2)
        intMom12 = intTotal + intMom11
        intArm12 = intMom12 / intWt12

        txtBEWA = intArm1
        txtBEWA = Format(txtBEWA.Text, "Fixed")
        txtCrewA = intArm2
        txtCrewA = Format(txtCrewA.Text, "Fixed")
        txtXtraA = intArm3
        txtXtraA = Format(txtXtraA.Text, "Fixed")
        txtA1 = intArm4
        txtA1 = Format(txtA1.Text, "Fixed")
        txtA2 = intArm5
        txtA2 = Format(txtA2.Text, "Fixed")
        txtA3 = intArm6
        txtA3 = Format(txtA3.Text, "Fixed")
        txtA4 = intArm7
        txtA4 = Format(txtA4.Text, "Fixed")
        txtA5 = intArm8
        txtA5 = Format(txtA5.Text, "Fixed")
        txtA6 = intArm9
        txtA6 = Format(txtA6.Text, "Fixed")
        txtA7 = intArm13
        txtA7 = Format(txtA7.Text, "Fixed")
        txtA8 = intArm14
        txtA8 = Format(txtA8.Text, "Fixed")
        txtA9 = intArm15
        txtA9 = Format(txtA9.Text, "Fixed")
        txtA10 = intArm16
        txtA10 = Format(txtA10.Text, "Fixed")
        
        txtZEWA = intArm10
        txtZEWA = Format(txtZEWA.Text, "Fixed")
        'txtFuelA = intArm11
        'txtFuelA = Format(txtFuelA.Text, "Fixed")
                
        If intTOWt <= 5500 Then
            
            If intArm12 < 179.6 Or intArm12 > 204.35 Then
                intCheck = 1
                Call CG
                Exit Sub
            Else
                txtTOWA = intArm12
                txtTOWA = Format(txtTOWA.Text, "Fixed")
                intCheck = 0
            End If
            
        ElseIf intTOWt > 5500 And intTOWt <= 8000 Then
        
            If intArm12 < 149.306 + (intTOWt * 0.005508) Or intArm12 > 204.35 Then
                intCheck = 1
                Call CG
                Exit Sub
            Else
                txtTOWA = intArm12
                txtTOWA = Format(txtTOWA.Text, "Fixed")
                intCheck = 0
            End If
            
        ElseIf intTOWt > 8000 And intTOWt <= 8750 Then
        
            If intArm12 < 131.71667 + intTOWt * 0.0077066667 Or intArm12 > 204.35 Then
                intCheck = 1
                Call CG
                Exit Sub
            Else
                txtTOWA = intArm12
                txtTOWA = Format(txtTOWA.Text, "Fixed")
                intCheck = 0
            End If
            
        'ElseIf intTOWt = 8750 Then
        
            'If intArm12 > 204.35 Then
                'intCheck = 1
                'Call CG
                'Exit Sub
            'Else
                'txtTOWA = intArm12
                'txtTOWA = Format(txtTOWA.Text, "Fixed")
                'intCheck = 0
            'End If
    
        End If
        
            txtBEWM = intMom1
            txtBEWM = Format(txtBEWM.Text, "Fixed")
            txtCrewM1 = intMom2
            txtCrewM1 = Format(txtCrewM1.Text, "Fixed")
            txtXtraM = intMom3
            txtXtraM = Format(txtXtraM.Text, "Fixed")
            txtM1 = intMom4
            txtM1 = Format(txtM1.Text, "Fixed")
            txtM2 = intMom5
            txtM2 = Format(txtM2.Text, "Fixed")
            txtM3 = intMom6
            txtM3 = Format(txtM3.Text, "Fixed")
            txtM4 = intMom7
            txtM4 = Format(txtM4.Text, "Fixed")
            txtM5 = intMom8
            txtM5 = Format(txtM5.Text, "Fixed")
            txtM6 = intMom9
            txtM6 = Format(txtM6.Text, "Fixed")
            txtM7 = intMom13
            txtM7 = Format(txtM7.Text, "Fixed")
            txtM8 = intMom14
            txtM8 = Format(txtM8.Text, "Fixed")
            txtM9 = intMom15
            txtM9 = Format(txtM9.Text, "Fixed")
            txtM10 = intMom16
            txtM10 = Format(txtM10.Text, "Fixed")
            txtZEWM = intMom10
            txtZEWM = Format(txtZEWM.Text, "Fixed")
            txtFuelM = intMom11
            txtFuelM = Format(txtFuelM.Text, "Fixed")
            txtTOWM1 = intMom12
            txtTOWM1 = Format(txtTOWM1.Text, "Fixed")
        End If
        intD = 8750 - intTOWt
        intE = 3400 - intCargoP
        intF = 8785 - intRampWt
        
        If intD <= intE And intD <= intE Then
            intUnder = intD
        ElseIf intE <= intD And intE <= intE Then
            intUnder = intB
        ElseIf intE <= intD And intE <= intE Then
            intUnder = intE
        End If
        
        'intUnder = intUnder - 2
                    
        If frmMain.Frame1.Caption = "CARGO in Kilograms" Then
            intUnder = intUnder / 2.20458553791887
            If intUnder < 0 Then
                intUnder = 0
            End If
            txtUnderLoad.Text = intUnder
            txtUnderLoad = Format(txtUnderLoad.Text, "Fixed")
            txtUnderLoad.Text = txtUnderLoad & " kgs"
        Else
            If intUnder < 0 Then
                intUnder = 0
            End If
            txtUnderLoad.Text = intUnder
            txtUnderLoad = Format(txtUnderLoad.Text, "Fixed")
            txtUnderLoad.Text = txtUnderLoad & " lbs"
        End If
            txtUnderLoad1.Text = txtUnderLoad
    
    intLandFuel = intFOB - intBurn - 35
    intX2 = intLandFuel * intLandFuel
    intLandFuelMom = intA + (intB * intLandFuel) + (intC * intX2)
    intLandMom = intMom10 + intLandFuelMom
    intLandArm = intLandMom / intLandWt
    Text4 = intLandArm
    Text4 = Format(Text4.Text, "Fixed")
    Text5 = intLandMom
    Text5 = Format(Text5.Text, "Fixed")
    Text6 = intLandFuel
    Text6 = Format(Text6.Text, "Fixed")
    Text7 = intLandFuelMom
    Text7 = Format(Text7.Text, "Fixed")
    
    'Text4 = Format(Text4, "###")
    'intLandArm = Val(Text4)
    'Text4 = intLandArm
    
    Select Case intLandWt
        
            Case 0 To 5500
            
                If intLandArm < 179.6 Or intLandArm > 204.35 Then
                    intCheck = 1
                    Call CG1
                    Exit Sub
                Else
                    intCheck = 0
                End If
                
            Case 5500.001 To 8000
            
                If intLandArm < 149.306 + (intLandWt * 0.005508) Or intLandArm > 204.35 Then
                    intCheck = 1
                    Call CG1
                    Exit Sub
                Else
                    intCheck = 0
                End If
                
            Case 8000.001 To 8749.999
            
                If intLandArm < 131.71667 + intLandWt * 0.0077066667 Or intLandArm > 204.35 Then
                    intCheck = 1
                    Call CG1
                    Exit Sub
                Else
                    intCheck = 0
                End If
                
            Case 8750
            
                If intLandArm > 204.35 Then
                    intCheck = 1
                    Call CG1
                    Exit Sub
                Else
                    intCheck = 0
                End If
        
            End Select
        If frmMain.cmbAC = "208 Cargo Master" Or frmMain.cmbAC = "208 Grand Caravan" Then
            txtMax1 = "415 lbs"
            txtMax2 = "860 lbs"
            txtMax3 = "495 lbs"
            txtMax4 = "340 lbs"
            txtMax5 = "315 lbs"
            txtMax6 = "245 lbs"
            txtMax10 = "230 lbs"
            txtMax11 = "310 lbs"
            txtMax12 = "270 lbs"
            txtMax13 = "280 lbs"
            Text8.BackColor = &HFFFFFF
            Text8 = "Pod A"
            Text60.BackColor = &HFFFFFF
            Text60 = "Pod B"
            Text64.BackColor = &HFFFFFF
            Text64 = "Pod C"
            Text74.BackColor = &HFFFFFF
            Text74 = "Pod D"
        Else
            txtMax1 = "1780 lbs"
            txtMax2 = "3100 lbs"
            txtMax3 = "1400 lbs"
            txtMax4 = "1380 lbs"
            txtMax5 = "1270 lbs"
            txtMax6 = "320 lbs"
            txtMax10 = ""
            txtMax11 = ""
            txtMax12 = ""
            txtMax13 = ""
            Text8.BackColor = &HC0C0C0
            Text8 = ""
            Text60.BackColor = &HC0C0C0
            Text60 = ""
            Text64.BackColor = &HC0C0C0
            Text64 = ""
            Text74.BackColor = &HC0C0C0
            Text74 = ""
            txtW7.BackColor = &HC0C0C0
            txtW7 = ""
            txtW8.BackColor = &HC0C0C0
            txtW8 = ""
            txtW9.BackColor = &HC0C0C0
            txtW9 = ""
            txtW10.BackColor = &HC0C0C0
            txtW10 = ""
            txtA7.BackColor = &HC0C0C0
            txtA8.BackColor = &HC0C0C0
            txtA9.BackColor = &HC0C0C0
            txtA10.BackColor = &HC0C0C0
            txtM7.BackColor = &HC0C0C0
            txtM8.BackColor = &HC0C0C0
            txtM9.BackColor = &HC0C0C0
            txtM10.BackColor = &HC0C0C0
            txtMax10.BackColor = &HC0C0C0
            txtMax11.BackColor = &HC0C0C0
            txtMax12.BackColor = &HC0C0C0
            txtMax13.BackColor = &HC0C0C0
            txtW7.ForeColor = &HFFFFFF
            txtW8.ForeColor = &HFFFFFF
            txtW9.ForeColor = &HFFFFFF
            txtW10.ForeColor = &HFFFFFF
            txtA7.ForeColor = &HFFFFFF
            txtA8.ForeColor = &HFFFFFF
            txtA9.ForeColor = &HFFFFFF
            txtA10.ForeColor = &HFFFFFF
            txtM7.ForeColor = &HFFFFFF
            txtM8.ForeColor = &HFFFFFF
            txtM9.ForeColor = &HFFFFFF
            txtM10.ForeColor = &HFFFFFF
            txtMax10.ForeColor = &HFFFFFF
            txtMax11.ForeColor = &HFFFFFF
            txtMax12.ForeColor = &HFFFFFF
            txtMax13.ForeColor = &HFFFFFF

            txtW7.Text = ""
            txtW8.Text = ""
            txtW9.Text = ""
            txtW10.Text = ""
            txtA7.Text = ""
            txtA8.Text = ""
            txtA9.Text = ""
            txtA10.Text = ""
            txtM7.Text = ""
            txtM8.Text = ""
            txtM9.Text = ""
            txtM10.Text = ""
            txtMax10.Text = ""
            txtMax11.Text = ""
            txtMax12.Text = ""
            txtMax13.Text = ""

            txtW7.Locked = True
            txtW8.Locked = True
            txtW9.Locked = True
            txtW10.Locked = True
            txtA7.Locked = True
            txtA8.Locked = True
            txtA9.Locked = True
            txtA10.Locked = True
            txtM7.Locked = True
            txtM8.Locked = True
            txtM9.Locked = True
            txtM10.Locked = True
            txtMax10.Locked = True
            txtMax11.Locked = True
            txtMax12.Locked = True
            txtMax13.Locked = True
            
            intWt13 = 0
            intWt14 = 0
            intWt15 = 0
            intWt16 = 0
            intArm13 = 0
            intArm14 = 0
            intArm15 = 0
            intArm16 = 0
            intMom13 = 0
            intMom14 = 0
            intMom15 = 0
            intMom16 = 0
            
        End If
    
End Sub

Private Sub txtCrewW_GotFocus()
    txtCrewW.SelStart = 0
    txtCrewW.SelLength = Len(txtCrewW.Text)
End Sub

Private Sub txtCrewW_KeyPress(KeyAscii As Integer)
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

Private Sub txtCrewW_LostFocus()
    'If Val(txtCrewW) > 340 Then
     '   Call MsgBox("Maximum Crew Weight exceeded. Please enter 340 lbs. or less", vbOKOnly, "Maximum Weight Exceeded")
    '    txtCrewW1 = ""
    '    txtCrewW = ""
    '    txtCrewW.SetFocus
    'Else
        txtCrewW1 = txtCrewW
        Call Calc
        txtCrewW = Format(txtCrewW.Text, "Fixed")
        txtCrewW1 = Format(txtCrewW1.Text, "Fixed")
    'End If
End Sub

Private Sub txtDest1_GotFocus()
    txtDest1 = txtDest
End Sub

Private Sub txtDest2_GotFocus()
    txtDest2 = txtDest
End Sub

Private Sub cmdLMC_Click()

intCheck1 = 1

If frmMain.txtSeats = "208BCM" Then

    If frmMain.Frame1.Caption = "CARGO in Kilograms" Then
    
        If cmbHold1 = "Hold 1" Then
            intK1a = Val(txtLMCWt1)
            If optA1 = False Then
                intK1 = intK1 - intK1a
            Else
                intK1 = intK1 + intK1a
            End If
            intP1 = intK1 * 2.20458553791887
            If intK1 < 0 Then
                Call MsgBox("Hold Weight can not be less than zero. Please re-enter a value within limits", vbOKOnly, "Compartment Limit Exceeded!")
                intK1 = Val(txtW1)
                txtLMCWt1 = ""
                txtLMCWt1.SetFocus
                Exit Sub
            End If
            If intK1 > 807 Then
                Call MsgBox("Compartment Limit exceeded. Please re-enter a value within limits", vbOKOnly, "Compartment Limit Exceeded!")
                intK1 = Val(txtW1)
                txtLMCWt1 = ""
                txtLMCWt1.SetFocus
                Exit Sub
            End If
        ElseIf cmbHold1 = "Hold 2" Then
            intK2a = Val(txtLMCWt1)
            If optA1 = False Then
                intK2 = intK2 - intK2a
            Else
                intK2 = intK2 + intK2a
            End If
            intP2 = intK2 * 2.20458553791887
            If intK2 < 0 Then
                Call MsgBox("Hold Weight can not be less than zero. Please re-enter a value within limits", vbOKOnly, "Compartment Limit Exceeded!")
                intK2 = Val(txtW2)
                txtLMCWt1 = ""
                txtLMCWt1.SetFocus
                Exit Sub
            End If
            If intK2 > 1406 Then
                Call MsgBox("Compartment Limit exceeded. Please re-enter a value within limits", vbOKOnly, "Compartment Limit Exceeded!")
                intK2 = Val(txtW2)
                txtLMCWt1 = ""
                txtLMCWt1.SetFocus
                Exit Sub
            End If
        ElseIf cmbHold1 = "Hold 3" Then
            intK3a = Val(txtLMCWt1)
            If optA1 = False Then
                intK3 = intK3 - intK3a
            Else
                intK3 = intK3 + intK3a
            End If
            intP3 = intK3 * 2.20458553791887
            If intK3 < 0 Then
                Call MsgBox("Hold Weight can not be less than zero. Please re-enter a value within limits", vbOKOnly, "Compartment Limit Exceeded!")
                intK3 = Val(txtW3)
                txtLMCWt1 = ""
                txtLMCWt1.SetFocus
                Exit Sub
            End If
            If intK3 > 862 Then
                Call MsgBox("Compartment Limit exceeded. Please re-enter a value within limits", vbOKOnly, "Compartment Limit Exceeded!")
                intK3 = Val(txtW3)
                txtLMCWt1 = ""
                txtLMCWt1.SetFocus
                Exit Sub
            End If
        ElseIf cmbHold1 = "Hold 4" Then
            intK4a = Val(txtLMCWt1)
            If optA1 = False Then
                intK4 = intK4 - intK4a
            Else
                intK4 = intK4 + intK4a
            End If
            intP4 = intK4 * 2.20458553791887
            If intK4 < 0 Then
                Call MsgBox("Hold Weight can not be less than zero. Please re-enter a value within limits", vbOKOnly, "Compartment Limit Exceeded!")
                intK4 = Val(txtW4)
                txtLMCWt1 = ""
                txtLMCWt1.SetFocus
                Exit Sub
            End If
            If intK4 > 626 Then
                Call MsgBox("Compartment Limit exceeded. Please re-enter a value within limits", vbOKOnly, "Compartment Limit Exceeded!")
                intK4 = Val(txtW4)
                txtLMCWt1 = ""
                txtLMCWt1.SetFocus
                Exit Sub
            End If
        ElseIf cmbHold1 = "Hold 5" Then
            intK5a = Val(txtLMCWt1)
            If optA1 = False Then
                intK5 = intK5 - intK5a
            Else
                intK5 = intK5 + intK5a
            End If
            intP5 = intK5 * 2.20458553791887
            If intK5 < 0 Then
                Call MsgBox("Hold Weight can not be less than zero. Please re-enter a value within limits", vbOKOnly, "Compartment Limit Exceeded!")
                intK5 = Val(txtW5)
                txtLMCWt1 = ""
                txtLMCWt1.SetFocus
                Exit Sub
            End If
            If intK5 > 575 Then
                Call MsgBox("Compartment Limit exceeded. Please re-enter a value within limits", vbOKOnly, "Compartment Limit Exceeded!")
                intK5 = Val(txtW5)
                txtLMCWt1 = ""
                txtLMCWt1.SetFocus
                Exit Sub
            End If
        ElseIf cmbHold1 = "Hold 6" Then
            intK6a = Val(txtLMCWt1)
            If optA1 = False Then
                intK6 = intK6 - intK6a
            Else
                intK6 = intK6 + intK6a
            End If
            intP6 = intK6 * 2.20458553791887
            If intK6 < 0 Then
                Call MsgBox("Hold Weight can not be less than zero. Please re-enter a value within limits", vbOKOnly, "Compartment Limit Exceeded!")
                intK6 = Val(txtW6)
                txtLMCWt1 = ""
                txtLMCWt1.SetFocus
                Exit Sub
            End If
            If intK6 > 145 Then
                Call MsgBox("Compartment Limit exceeded. Please re-enter a value within limits", vbOKOnly, "Compartment Limit Exceeded!")
                intK6 = Val(txtW6)
                txtLMCWt1 = ""
                txtLMCWt1.SetFocus
                Exit Sub
            End If
        End If
        'Call Calc
    
    
        If cmbHold2 = "Hold 1" Then
            intK1a = Val(txtLMCWt2)
            If optA2 = False Then
                intK1 = intK1 - intK1a
            Else
                intK1 = intK1 + intK1a
            End If
            intP1 = intK1 * 2.20458553791887
            If intK1 < 0 Then
                Call MsgBox("Hold Weight can not be less than zero. Please re-enter a value within limits", vbOKOnly, "Compartment Limit Exceeded!")
                intK1 = Val(txtW1)
                txtLMCWt2 = ""
                txtLMCWt2.SetFocus
                Exit Sub
            End If
            If intK1 > 807 Then
                Call MsgBox("Compartment Limit exceeded. Please re-enter a value within limits", vbOKOnly, "Compartment Limit Exceeded!")
                intK1 = Val(txtW1)
                txtLMCWt2 = ""
                txtLMCWt2.SetFocus
                Exit Sub
            End If
        ElseIf cmbHold2 = "Hold 2" Then
            intK2a = Val(txtLMCWt2)
            If optA2 = False Then
                intK2 = intK2 - intK2a
            Else
                intK2 = intK2 + intK2a
            End If
            intP2 = intK2 * 2.20458553791887
            If intK2 < 0 Then
                Call MsgBox("Hold Weight can not be less than zero. Please re-enter a value within limits", vbOKOnly, "Compartment Limit Exceeded!")
                intK2 = Val(txtW2)
                txtLMCWt2 = ""
                txtLMCWt2.SetFocus
                Exit Sub
            End If
            If intK2 > 1406 Then
                Call MsgBox("Compartment Limit exceeded. Please re-enter a value within limits", vbOKOnly, "Compartment Limit Exceeded!")
                intK2 = Val(txtW2)
                txtLMCWt2 = ""
                txtLMCWt2.SetFocus
                Exit Sub
            End If
        ElseIf cmbHold2 = "Hold 3" Then
            intK3a = Val(txtLMCWt2)
            If optA2 = False Then
                intK3 = intK3 - intK3a
            Else
                intK3 = intK3 + intK3a
            End If
            intP3 = intK3 * 2.20458553791887
            If intK3 < 0 Then
                Call MsgBox("Hold Weight can not be less than zero. Please re-enter a value within limits", vbOKOnly, "Compartment Limit Exceeded!")
                intK3 = Val(txtW3)
                txtLMCWt2 = ""
                txtLMCWt2.SetFocus
                Exit Sub
            End If
            If intK3 > 862 Then
                Call MsgBox("Compartment Limit exceeded. Please re-enter a value within limits", vbOKOnly, "Compartment Limit Exceeded!")
                intK3 = Val(txtW3)
                txtLMCWt2 = ""
                txtLMCWt2.SetFocus
                Exit Sub
            End If
        ElseIf cmbHold2 = "Hold 4" Then
            intK4a = Val(txtLMCWt2)
            If optA2 = False Then
                intK4 = intK4 - intK4a
            Else
                intK4 = intK4 + intK4a
            End If
            intP4 = intK4 * 2.20458553791887
            If intK4 < 0 Then
                Call MsgBox("Hold Weight can not be less than zero. Please re-enter a value within limits", vbOKOnly, "Compartment Limit Exceeded!")
                intK4 = Val(txtW4)
                txtLMCWt2 = ""
                txtLMCWt2.SetFocus
                Exit Sub
            End If
            If intK4 > 626 Then
                Call MsgBox("Compartment Limit exceeded. Please re-enter a value within limits", vbOKOnly, "Compartment Limit Exceeded!")
                intK4 = Val(txtW4)
                txtLMCWt2 = ""
                txtLMCWt2.SetFocus
                Exit Sub
            End If
        ElseIf cmbHold2 = "Hold 5" Then
            intK5a = Val(txtLMCWt2)
            If optA2 = False Then
                intK5 = intK5 - intK5a
            Else
                intK5 = intK5 + intK5a
            End If
            intP5 = intK5 * 2.20458553791887
            If intK5 < 0 Then
                Call MsgBox("Hold Weight can not be less than zero. Please re-enter a value within limits", vbOKOnly, "Compartment Limit Exceeded!")
                intK5 = Val(txtW5)
                txtLMCWt2 = ""
                txtLMCWt2.SetFocus
                Exit Sub
            End If
            If intK5 > 575 Then
                Call MsgBox("Compartment Limit exceeded. Please re-enter a value within limits", vbOKOnly, "Compartment Limit Exceeded!")
                intK5 = Val(txtW5)
                txtLMCWt2 = ""
                txtLMCWt2.SetFocus
                Exit Sub
            End If
        ElseIf cmbHold2 = "Hold 6" Then
            intK6a = Val(txtLMCWt2)
            If optA2 = False Then
                intK6 = intK6 - intK6a
            Else
                intK6 = intK6 + intK6a
            End If
            intP6 = intK6 * 2.20458553791887
            If intK6 < 0 Then
                Call MsgBox("Hold Weight can not be less than zero. Please re-enter a value within limits", vbOKOnly, "Compartment Limit Exceeded!")
                intK6 = Val(txtW6)
                txtLMCWt2 = ""
                txtLMCWt2.SetFocus
                Exit Sub
            End If
            If intK6 > 145 Then
                Call MsgBox("Compartment Limit exceeded. Please re-enter a value within limits", vbOKOnly, "Compartment Limit Exceeded!")
                intK6 = Val(txtW6)
                txtLMCWt2 = ""
                txtLMCWt2.SetFocus
                Exit Sub
            End If
        End If
        'Call Calc
        
    Else
    
        If cmbHold1 = "Hold 1" Then
            intP1a = Val(txtLMCWt1)
            If optA1 = False Then
                intP1 = intP1 - intP1a
            Else
                intP1 = intP1 + intP1a
            End If
            If intP1 < 0 Then
                Call MsgBox("Hold Weight can not be less than zero. Please re-enter a value within limits", vbOKOnly, "Compartment Limit Exceeded!")
                intP1 = Val(txtW1)
                txtLMCWt1 = ""
                txtLMCWt1.SetFocus
                Exit Sub
            End If
            If intP1 > 1780 Then
                'intP1 = intP1a
                Call MsgBox("Compartment Limit exceeded. Please re-enter a value within limits", vbOKOnly, "Compartment Limit Exceeded!")
                intP1 = Val(txtW1)
                txtLMCWt1 = ""
                txtLMCWt1.SetFocus
                Exit Sub
            End If
        ElseIf cmbHold1 = "Hold 2" Then
            intP2a = Val(txtLMCWt1)
            If optA1 = False Then
                intP2 = intP2 - intP2a
            Else
                intP2 = intP2 + intP2a
            End If
            If intP2 < 0 Then
                Call MsgBox("Hold Weight can not be less than zero. Please re-enter a value within limits", vbOKOnly, "Compartment Limit Exceeded!")
                intP2 = Val(txtW2)
                txtLMCWt1 = ""
                txtLMCWt1.SetFocus
                Exit Sub
            End If
            If intP2 > 3100 Then
                Call MsgBox("Compartment Limit exceeded. Please re-enter a value within limits", vbOKOnly, "Compartment Limit Exceeded!")
                intP2 = Val(txtW2)
                txtLMCWt1 = ""
                txtLMCWt1.SetFocus
                Exit Sub
            End If
        ElseIf cmbHold1 = "Hold 3" Then
            intP3a = Val(txtLMCWt1)
            If optA1 = False Then
                intP3 = intP3 - intP3a
            Else
                intP3 = intP3 + intP3a
            End If
            If intP3 < 0 Then
                Call MsgBox("Hold Weight can not be less than zero. Please re-enter a value within limits", vbOKOnly, "Compartment Limit Exceeded!")
                intP3 = Val(txtW3)
                txtLMCWt1 = ""
                txtLMCWt1.SetFocus
                Exit Sub
            End If
            If intP3 > 1400 Then
                Call MsgBox("Compartment Limit exceeded. Please re-enter a value within limits", vbOKOnly, "Compartment Limit Exceeded!")
                intP3 = Val(txtW3)
                txtLMCWt1 = ""
                txtLMCWt1.SetFocus
                Exit Sub
            End If
        ElseIf cmbHold1 = "Hold 4" Then
            intP4a = Val(txtLMCWt1)
            If optA1 = False Then
                intP4 = intP4 - intP4a
            Else
                intP4 = intP4 + intP4a
            End If
            If intP4 < 0 Then
                Call MsgBox("Hold Weight can not be less than zero. Please re-enter a value within limits", vbOKOnly, "Compartment Limit Exceeded!")
                intP4 = Val(txtW4)
                txtLMCWt1 = ""
                txtLMCWt1.SetFocus
                Exit Sub
            End If
            If intP4 > 1380 Then
                Call MsgBox("Compartment Limit exceeded. Please re-enter a value within limits", vbOKOnly, "Compartment Limit Exceeded!")
                intP4 = Val(txtW4)
                txtLMCWt1 = ""
                txtLMCWt1.SetFocus
                Exit Sub
            End If
        ElseIf cmbHold1 = "Hold 5" Then
            intP5a = Val(txtLMCWt1)
            If optA1 = False Then
                intP5 = intP5 - intP5a
            Else
                intP5 = intP5 + intP5a
            End If
            If intP5 < 0 Then
                Call MsgBox("Hold Weight can not be less than zero. Please re-enter a value within limits", vbOKOnly, "Compartment Limit Exceeded!")
                intP5 = Val(txtW5)
                txtLMCWt1 = ""
                txtLMCWt1.SetFocus
                Exit Sub
            End If
            If intP5 > 1270 Then
                Call MsgBox("Compartment Limit exceeded. Please re-enter a value within limits", vbOKOnly, "Compartment Limit Exceeded!")
                intP5 = Val(txtW5)
                txtLMCWt1 = ""
                txtLMCWt1.SetFocus
                Exit Sub
            End If
        ElseIf cmbHold1 = "Hold 6" Then
            intP6a = Val(txtLMCWt1)
            If optA1 = False Then
                intP6 = intP6 - intP6a
            Else
                intP6 = intP6 + intP6a
            End If
            If intP6 < 0 Then
                Call MsgBox("Hold Weight can not be less than zero. Please re-enter a value within limits", vbOKOnly, "Compartment Limit Exceeded!")
                intP6 = Val(txtW6)
                txtLMCWt1 = ""
                txtLMCWt1.SetFocus
                Exit Sub
            End If
            If intP6 > 320 Then
                Call MsgBox("Compartment Limit exceeded. Please re-enter a value within limits", vbOKOnly, "Compartment Limit Exceeded!")
                intP6 = Val(txtW6)
                txtLMCWt1 = ""
                txtLMCWt1.SetFocus
                Exit Sub
            End If
        End If
        'Call Calc
    
    
        If cmbHold2 = "Hold 1" Then
            intP1a = Val(txtLMCWt2)
            If optA2 = False Then
                intP1 = intP1 - intP1a
            Else
                intP1 = intP1 + intP1a
            End If
            If intP1 < 0 Then
                Call MsgBox("Hold Weight can not be less than zero. Please re-enter a value within limits", vbOKOnly, "Compartment Limit Exceeded!")
                intP1 = Val(txtW1)
                txtLMCWt2 = ""
                txtLMCWt2.SetFocus
                Exit Sub
            End If
            If intP1 > 1780 Then
                Call MsgBox("Compartment Limit exceeded. Please re-enter a value within limits", vbOKOnly, "Compartment Limit Exceeded!")
                intP1 = Val(txtW1)
                txtLMCWt2 = ""
                txtLMCWt2.SetFocus
                Exit Sub
            End If
        ElseIf cmbHold2 = "Hold 2" Then
            intP2a = Val(txtLMCWt2)
            If optA2 = False Then
                intP2 = intP2 - intP2a
            Else
                intP2 = intP2 + intP2a
            End If
            If intP2 < 0 Then
                Call MsgBox("Hold Weight can not be less than zero. Please re-enter a value within limits", vbOKOnly, "Compartment Limit Exceeded!")
                intP2 = Val(txtW2)
                txtLMCWt2 = ""
                txtLMCWt2.SetFocus
                Exit Sub
            End If
            If intP2 > 3100 Then
                Call MsgBox("Compartment Limit exceeded. Please re-enter a value within limits", vbOKOnly, "Compartment Limit Exceeded!")
                intP2 = Val(txtW2)
                txtLMCWt2 = ""
                txtLMCWt2.SetFocus
                Exit Sub
            End If
        ElseIf cmbHold2 = "Hold 3" Then
            intP3a = Val(txtLMCWt2)
            If optA2 = False Then
                intP3 = intP3 - intP3a
            Else
                intP3 = intP3 + intP3a
            End If
            If intP3 < 0 Then
                Call MsgBox("Hold Weight can not be less than zero. Please re-enter a value within limits", vbOKOnly, "Compartment Limit Exceeded!")
                intP3 = Val(txtW3)
                txtLMCWt2 = ""
                txtLMCWt2.SetFocus
                Exit Sub
            End If
            If intP3 > 1400 Then
                Call MsgBox("Compartment Limit exceeded. Please re-enter a value within limits", vbOKOnly, "Compartment Limit Exceeded!")
                intP3 = Val(txtW3)
                txtLMCWt2 = ""
                txtLMCWt2.SetFocus
                Exit Sub
            End If
        ElseIf cmbHold2 = "Hold 4" Then
            intP4a = Val(txtLMCWt2)
            If optA2 = False Then
                intP4 = intP4 - intP4a
            Else
                intP4 = intP4 + intP4a
            End If
            If intP4 < 0 Then
                Call MsgBox("Hold Weight can not be less than zero. Please re-enter a value within limits", vbOKOnly, "Compartment Limit Exceeded!")
                intP4 = Val(txtW4)
                txtLMCWt2 = ""
                txtLMCWt2.SetFocus
                Exit Sub
            End If
            If intP4 > 1380 Then
                Call MsgBox("Compartment Limit exceeded. Please re-enter a value within limits", vbOKOnly, "Compartment Limit Exceeded!")
                intP4 = Val(txtW4)
                txtLMCWt2 = ""
                txtLMCWt2.SetFocus
                Exit Sub
            End If
        ElseIf cmbHold2 = "Hold 5" Then
            intP5a = Val(txtLMCWt2)
            If optA2 = False Then
                intP5 = intP5 - intP5a
            Else
                intP5 = intP5 + intP5a
            End If
            If intP5 < 0 Then
                Call MsgBox("Hold Weight can not be less than zero. Please re-enter a value within limits", vbOKOnly, "Compartment Limit Exceeded!")
                intP5 = Val(txtW5)
                txtLMCWt2 = ""
                txtLMCWt2.SetFocus
                Exit Sub
            End If
            If intP5 > 1270 Then
                Call MsgBox("Compartment Limit exceeded. Please re-enter a value within limits", vbOKOnly, "Compartment Limit Exceeded!")
                intP5 = Val(txtW5)
                txtLMCWt2 = ""
                txtLMCWt2.SetFocus
                Exit Sub
            End If
        ElseIf cmbHold2 = "Hold 6" Then
            intP6a = Val(txtLMCWt2)
            If optA2 = False Then
                intP6 = intP6 - intP6a
            Else
                intP6 = intP6 + intP6a
            End If
            If intP6 < 0 Then
                Call MsgBox("Hold Weight can not be less than zero. Please re-enter a value within limits", vbOKOnly, "Compartment Limit Exceeded!")
                intP6 = Val(txtW6)
                txtLMCWt2 = ""
                txtLMCWt2.SetFocus
                Exit Sub
            End If
            If intP6 > 320 Then
                Call MsgBox("Compartment Limit exceeded. Please re-enter a value within limits", vbOKOnly, "Compartment Limit Exceeded!")
                intP6 = Val(txtW6)
                txtLMCWt2 = ""
                txtLMCWt2.SetFocus
                Exit Sub
            End If
        End If
    
    End If
    
End If

'    txtDest1 = ""
 '   txtDest2 = ""
    Call Calc
    txtComp1 = ""
    cmbHold1 = ""
    txtComp2 = ""
    cmbHold2 = ""
    txtAS1 = ""
    'cmbAS1 = ""
    txtAS2 = ""
    'cmbAS2 = ""
    txtLMCWt1 = ""
    txtLMCWt2 = ""
    
End Sub

Private Sub txtLMCWt1_KeyPress(KeyAscii As Integer)
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

Private Sub txtLMCWt2_KeyPress(KeyAscii As Integer)
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

Private Sub txtXCrewW_GotFocus()
    txtXCrewW.SelStart = 0
    txtXCrewW.SelLength = Len(txtXCrewW.Text)
End Sub

Private Sub txtXCrewW_KeyPress(KeyAscii As Integer)
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

Private Sub txtXCrewW_LostFocus()
    'If Val(txtXCrewW) > 170 Then
    '    Call MsgBox("Maximum Extra Crew Weight exceeded. Please enter 170 lbs. or less", vbOKOnly, "Maximum Weight Exceeded")
    '    txtXCrewW = ""
    '    txtXCrewW.SetFocus
   ' Else
        txtXtraW = txtXCrewW
        Call Calc
        txtXCrewW = Format(txtXCrewW.Text, "Fixed")
        txtXtraW = Format(txtXtraW.Text, "Fixed")
    'End If
End Sub

Private Sub Recall()

If frmMain.Frame1.Caption = "CARGO in Kilograms" Then

    If cmbHold1 = "Hold 1" Then
        If optA1 = False Then
            intK1 = intK1 + intK1a
        Else
            intK1 = intK1 - intK1a
        End If
        txtW1 = intK1
        txtW1 = Format(txtW1.Text, "Fixed")
        
    ElseIf cmbHold1 = "Hold 2" Then
        If optA1 = False Then
            intK2 = intK2 + intK2a
        Else
            intK2 = intK2 - intK2a
        End If
        txtW2 = intK2
        txtW2 = Format(txtW2.Text, "Fixed")
        
    ElseIf cmbHold1 = "Hold 3" Then
        If optA1 = False Then
            intK3 = intK3 + intK3a
        Else
            intK3 = intK3 - intK3a
        End If
        txtW3 = intK3
        txtW3 = Format(txtW3.Text, "Fixed")
        
    ElseIf cmbHold1 = "Hold 4" Then
        If optA1 = False Then
            intK4 = intK4 + intK4a
        Else
            intK4 = intK4 - intK4a
        End If
        txtW4 = intK4
        txtW4 = Format(txtW4.Text, "Fixed")
        
    ElseIf cmbHold1 = "Hold 5" Then
        If optA1 = False Then
            intK5 = intK5 + intK5a
        Else
            intK5 = intK5 - intK5a
        End If
        txtW5 = intK5
        txtW5 = Format(txtW5.Text, "Fixed")
        
    ElseIf cmbHold1 = "Hold 6" Then
        If optA1 = False Then
            intK6 = intK6 + intK6a
        Else
            intK6 = intK6 - intK6a
        End If
        txtW6 = intK6
        txtW6 = Format(txtW6.Text, "Fixed")
        
    ElseIf cmbHold1 = "Pod A" Then
        If optA1 = False Then
            intK7 = intK7 + intK7a
        Else
            intK7 = intK7 - intK7a
        End If
        txtW7 = intK7
        txtW7 = Format(txtW7.Text, "Fixed")
        
    ElseIf cmbHold1 = "Pod B" Then
        If optA1 = False Then
            intK8 = intK8 + intK8a
        Else
            intK8 = intK8 - intK8a
        End If
        txtW8 = intK8
        txtW8 = Format(txtW8.Text, "Fixed")
        
    ElseIf cmbHold1 = "Pod C" Then
        If optA1 = False Then
            intK8 = intK8 + intK8a
        Else
            intK8 = intK8 - intK8a
        End If
        txtW8 = intK8
        txtW8 = Format(txtW8.Text, "Fixed")
    
    ElseIf cmbHold1 = "Pod D" Then
        If optA1 = False Then
            intK9 = intK9 + intK9a
        Else
            intK9 = intK9 - intK9a
        End If
        txtW9 = intK9
        txtW9 = Format(txtW9.Text, "Fixed")
        
    End If

    If cmbHold2 = "Hold 1" Then
        If optA2 = False Then
            intK1 = intK1 + intK1a
        Else
            intK1 = intK1 - intK1a
        End If
        txtW1 = intK1
        txtW1 = Format(txtW1.Text, "Fixed")
        
    ElseIf cmbHold2 = "Hold 2" Then
        If optA2 = False Then
            intK2 = intK2 + intK2a
        Else
            intK2 = intK2 - intK2a
        End If
        txtW2 = intK2
        txtW2 = Format(txtW2.Text, "Fixed")
        
    ElseIf cmbHold2 = "Hold 3" Then
        If optA2 = False Then
            intK3 = intK3 + intK3a
        Else
            intK3 = intK3 - intK3a
        End If
        txtW3 = intK3
        txtW3 = Format(txtW3.Text, "Fixed")
        
    ElseIf cmbHold2 = "Hold 4" Then
        If optA2 = False Then
            intK4 = intK4 + intK4a
        Else
            intK4 = intK4 - intK4a
        End If
        txtW4 = intK4
        txtW4 = Format(txtW4.Text, "Fixed")
        
    ElseIf cmbHold2 = "Hold 5" Then
        If optA2 = False Then
            intK5 = intK5 + intK5a
        Else
            intK5 = intK5 - intK5a
        End If
        txtW5 = intK5
        txtW5 = Format(txtW5.Text, "Fixed")
        
    ElseIf cmbHold2 = "Hold 6" Then
        If optA2 = False Then
            intK6 = intK6 + intK6a
        Else
            intK6 = intK6 - intK6a
        End If
        txtW6 = intK6
        txtW6 = Format(txtW6.Text, "Fixed")
        
    ElseIf cmbHold2 = "Pod A" Then
        If optA2 = False Then
            intK7 = intK7 + intK7a
        Else
            intK7 = intK7 - intK7a
        End If
        txtW7 = intK7
        txtW7 = Format(txtW7.Text, "Fixed")
        
    ElseIf cmbHold2 = "Pod B" Then
        If optA2 = False Then
            intK8 = intK8 + intK8a
        Else
            intK8 = intK8 - intK8a
        End If
        txtW8 = intK8
        txtW8 = Format(txtW8.Text, "Fixed")
        
    ElseIf cmbHold2 = "Pod C" Then
        If optA2 = False Then
            intK9 = intK9 + intK9a
        Else
            intK9 = intK9 - intK9a
        End If
        txtW9 = intK9
        txtW9 = Format(txtW9.Text, "Fixed")
        
    ElseIf cmbHold2 = "Pod D" Then
        If optA2 = False Then
            intK10 = intK10 + intK10a
        Else
            intK10 = intK10 - intK10a
        End If
        txtW10 = intK10
        txtW10 = Format(txtW10.Text, "Fixed")
        
    End If
    
Else

    If cmbHold1 = "Hold 1" Then
        If optA1 = False Then
            intP1 = intP1 + intP1a
        Else
            intP1 = intP1 - intP1a
        End If
        txtW1 = intP1
        txtW1 = Format(txtW1.Text, "Fixed")
        
    ElseIf cmbHold1 = "Hold 2" Then
        If optA1 = False Then
            intP2 = intP2 + intP2a
        Else
            intP2 = intP2 - intP2a
        End If
        txtW2 = intP2
        txtW2 = Format(txtW2.Text, "Fixed")
        
    ElseIf cmbHold1 = "Hold 3" Then
        If optA1 = False Then
            intP3 = intP3 + intP3a
        Else
            intP3 = intP3 - intP3a
        End If
        txtW3 = intP3
        txtW3 = Format(txtW3.Text, "Fixed")
        
    ElseIf cmbHold1 = "Hold 4" Then
        If optA1 = False Then
            intP4 = intP4 + intP4a
        Else
            intP4 = intP4 - intP4a
        End If
        txtW4 = intP4
        txtW4 = Format(txtW4.Text, "Fixed")
        
    ElseIf cmbHold1 = "Hold 5" Then
        If optA1 = False Then
            intP5 = intP5 + intP5a
        Else
            intP5 = intP5 - intP5a
        End If
        txtW5 = intP5
        txtW5 = Format(txtW5.Text, "Fixed")
        
    ElseIf cmbHold1 = "Hold 6" Then
        If optA1 = False Then
            intP6 = intP6 + intP6a
        Else
            intP6 = intP6 - intP6a
        End If
        txtW6 = intP6
        txtW6 = Format(txtW6.Text, "Fixed")
    
    ElseIf cmbHold1 = "Pod A" Then
        If optA1 = False Then
            intP7 = intP7 + intP7a
        Else
            intP7 = intP7 - intP7a
        End If
        txtW7 = intP7
        txtW7 = Format(txtW7.Text, "Fixed")
        
    ElseIf cmbHold1 = "Pod B" Then
        If optA1 = False Then
            intP8 = intP8 + intP8a
        Else
            intP8 = intP8 - intP8a
        End If
        txtW8 = intP8
        txtW8 = Format(txtW8.Text, "Fixed")
        
    ElseIf cmbHold1 = "Pod C" Then
        If optA1 = False Then
            intP9 = intP9 + intP9a
        Else
            intP9 = intP9 - intP9a
        End If
        txtW9 = intP9
        txtW9 = Format(txtW9.Text, "Fixed")
        
    ElseIf cmbHold1 = "Pod D" Then
        If optA1 = False Then
            intP10 = intP10 + intP10a
        Else
            intP10 = intP10 - intP10a
        End If
        txtW10 = intP10
        txtW10 = Format(txtW10.Text, "Fixed")
    
    End If

    If cmbHold2 = "Hold 1" Then
        If optA2 = False Then
            intP1 = intP1 + intP1a
        Else
            intP1 = intP1 - intP1a
        End If
        txtW1 = intP1
        txtW1 = Format(txtW1.Text, "Fixed")
        
    ElseIf cmbHold2 = "Hold 2" Then
        If optA2 = False Then
            intP2 = intP2 + intP2a
        Else
            intP2 = intP2 - intP2a
        End If
        txtW2 = intP2
        txtW2 = Format(txtW2.Text, "Fixed")
        
    ElseIf cmbHold2 = "Hold 3" Then
        If optA2 = False Then
            intP3 = intP3 + intP3a
        Else
            intP3 = intP3 - intP3a
        End If
        txtW3 = intP3
        txtW3 = Format(txtW3.Text, "Fixed")
        
    ElseIf cmbHold2 = "Hold 4" Then
        If optA2 = False Then
            intP4 = intP4 + intP4a
        Else
            intP4 = intP4 - intP4a
        End If
        txtW4 = intP4
        txtW4 = Format(txtW4.Text, "Fixed")
        
    ElseIf cmbHold2 = "Hold 5" Then
        If optA2 = False Then
            intP5 = intP5 + intP5a
        Else
            intP5 = intP5 - intP5a
        End If
        txtW5 = intP5
        txtW5 = Format(txtW5.Text, "Fixed")
        
    ElseIf cmbHold2 = "Hold 6" Then
        If optA2 = False Then
            intP6 = intP6 + intP6a
        Else
            intP6 = intP6 - intP6a
        End If
        txtW6 = intP6
        txtW6 = Format(txtW6.Text, "Fixed")
        
    ElseIf cmbHold2 = "Pod A" Then
        If optA2 = False Then
            intP7 = intP7 + intP7a
        Else
            intP7 = intP7 - intP7a
        End If
        txtW7 = intP7
        txtW7 = Format(txtW7.Text, "Fixed")
        
    ElseIf cmbHold2 = "Pod B" Then
        If optA2 = False Then
            intP8 = intP8 + intP8a
        Else
            intP8 = intP8 - intP8a
        End If
        txtW8 = intP8
        txtW8 = Format(txtW8.Text, "Fixed")
        
    ElseIf cmbHold2 = "Pod C" Then
        If optA2 = False Then
            intP9 = intP9 + intP9a
        Else
            intP9 = intP9 - intP9a
        End If
        txtW9 = intP9
        txtW9 = Format(txtW9.Text, "Fixed")
        
    ElseIf cmbHold2 = "Pod D" Then
        If optA2 = False Then
            intP10 = intP10 + intP10a
        Else
            intP10 = intP10 - intP10a
        End If
        txtW10 = intP10
        txtW10 = Format(txtW10.Text, "Fixed")
        
    End If
    
End If
        
        txtBEWM = intMom1
        txtBEWM = Format(txtBEWM.Text, "Fixed")
        txtCrewM1 = intMom2
        txtCrewM1 = Format(txtCrewM1.Text, "Fixed")
        txtXtraM = intMom3
        txtXtraM = Format(txtXtraM.Text, "Fixed")
        txtM1 = intMom4
        txtM1 = Format(txtM1.Text, "Fixed")
        txtM2 = intMom5
        txtM2 = Format(txtM2.Text, "Fixed")
        txtM3 = intMom6
        txtM3 = Format(txtM3.Text, "Fixed")
        txtM4 = intMom7
        txtM4 = Format(txtM4.Text, "Fixed")
        txtM5 = intMom8
        txtM5 = Format(txtM5.Text, "Fixed")
        txtM6 = intMom9
        txtM6 = Format(txtM6.Text, "Fixed")
        txtM7 = intMom13
        txtM7 = Format(txtM7.Text, "Fixed")
        txtM8 = intMom14
        txtM8 = Format(txtM8.Text, "Fixed")
        txtM9 = intMom15
        txtM9 = Format(txtM9.Text, "Fixed")
        txtM10 = intMom16
        txtM10 = Format(txtM10.Text, "Fixed")
        txtZEWM = intMom10
        txtZEWM = Format(txtZEWM.Text, "Fixed")
        txtFuelM = intMom11
        txtFuelM = Format(txtFuelM.Text, "Fixed")
        txtTOWM1 = intMom12
        txtTOWM1 = Format(txtTOWM1.Text, "Fixed")
        
        intD = 8750 - intTOWt
        intE = 3400 - intCargoP
        intF = 8785 - intRampWt
        
        If intD <= intE And intD <= intE Then
            intUnder = intD
        ElseIf intE <= intD And intE <= intE Then
            intUnder = intB
        ElseIf intE <= intD And intE <= intE Then
            intUnder = intE
        End If
        
        'intUnder = intUnder - 2
                    
        If frmMain.Frame1.Caption = "CARGO in Kilograms" Then
            intUnder = intUnder / 2.20458553791887
            If intUnder < 0 Then
                intUnder = 0
            End If
            txtUnderLoad.Text = intUnder
            txtUnderLoad = Format(txtUnderLoad.Text, "Fixed")
            txtUnderLoad.Text = txtUnderLoad & " kgs"
        Else
            If intUnder < 0 Then
                intUnder = 0
            End If
            txtUnderLoad.Text = intUnder
            txtUnderLoad = Format(txtUnderLoad.Text, "Fixed")
            txtUnderLoad.Text = txtUnderLoad & " lbs"
        End If
            txtUnderLoad1.Text = txtUnderLoad
    
    intLandFuel = intFOB - intBurn - 35
    intX2 = intLandFuel * intLandFuel
    intLandFuelMom = intA + (intB * intLandFuel) + (intC * intX2)
    intLandMom = intMom10 + intLandFuelMom
    intLandArm = intLandMom / intLandWt
    Text4 = intLandArm
    Text4 = Format(Text4.Text, "Fixed")
    'Text4 = Format(Text4, "###")
    'intLandArm = Val(Text4)
    'Text4 = intLandArm
    Call Calc
    
End Sub

Private Sub CG()

    If intArm12 > 204.35 Then
        Call MsgBox("Out of Aft Center of Gravity Limit with Take Off configuration." & vbNewLine & "Please re-enter acceptable values.", vbOKOnly, "Out of Aft Center of Gravity Limit")
    Else
        Call MsgBox("Out of Forward Center of Gravity Limit with Take Off configuration." & vbNewLine & "Please re-enter acceptable values.", vbOKOnly, "Out of Forward Center of Gravity Limit")
    End If
    
    intTOWt = Val(frmMain.txtTOWt)
    intCrew = 340
    txtTOWW1 = intTOWt
    txtTOWW = intTOWt
    If intCheck1 = 0 Then
        'frmMain.txtZone1 = ""
        'frmMain.txtZone2 = ""
        'frmMain.txtZone3 = ""
        'frmMain.txtZone4 = ""
        'frmMain.txtZone5 = ""
        'frmMain.txtZone6 = ""
        intTOWt = 0
        intTotal = 0
        intK1 = 0
        intK2 = 0
        intK3 = 0
        intK4 = 0
        intK5 = 0
        intK6 = 0
        intK7 = 0
        intK8 = 0
        intK9 = 0
        intK10 = 0
        int1 = 0
        int2 = 0
        int3 = 0
        int4 = 0
        int5 = 0
        int6 = 0
        int7 = 0
        int8 = 0
        int9 = 0
        int10 = 0
        intP1 = 0
        intP2 = 0
        intP3 = 0
        intP4 = 0
        intP5 = 0
        intP6 = 0
        intP7 = 0
        intP8 = 0
        intP9 = 0
        intP10 = 0
        frmMain.txtTOWt = ""
        frmMain.Show
        frmMain.txtZone1.SetFocus
        frmMain.txtZone2.SetFocus
        frmMain.txtZone3.SetFocus
        frmMain.txtZone4.SetFocus
        frmMain.txtZone5.SetFocus
'        frmMain.txtZone6.SetFocus
        frmMain.txtZone1.SetFocus
        Unload Me
    Else
        Call Recall
        cmbHold1.SetFocus
        'Exit Sub
    End If

End Sub

Private Sub CG1()


    If intLandArm > 204.35 Then
        Call MsgBox("Out of Aft Center of Gravity Limit with Landing configuration." & vbNewLine & "Please re-enter acceptable values.", vbOKOnly, "Out of Aft Center of Gravity Limit")
    Else
        Call MsgBox("Out of Forward Center of Gravity Limit with Landing configuration." & vbNewLine & "Please re-enter acceptable values.", vbOKOnly, "Out of Forward Center of Gravity Limit")
    End If
    
    intTOWt = Val(frmMain.txtTOWt)
    intCrew = 340
    txtTOWW1 = intTOWt
    txtTOWW = intTOWt
    If intCheck1 = 0 Then
        'frmMain.txtZone1 = ""
        'frmMain.txtZone2 = ""
        'frmMain.txtZone3 = ""
        'frmMain.txtZone4 = ""
        'frmMain.txtZone5 = ""
        'frmMain.txtZone6 = ""
        intTOWt = 0
        intTotal = 0
        intK1 = 0
        intK2 = 0
        intK3 = 0
        intK4 = 0
        intK5 = 0
        intK6 = 0
        intK7 = 0
        intK8 = 0
        intK9 = 0
        intK10 = 0
        int1 = 0
        int2 = 0
        int3 = 0
        int4 = 0
        int5 = 0
        int6 = 0
        int7 = 0
        int8 = 0
        int9 = 0
        int10 = 0
        intP1 = 0
        intP2 = 0
        intP3 = 0
        intP4 = 0
        intP5 = 0
        intP6 = 0
        intP7 = 0
        intP8 = 0
        intP9 = 0
        intP10 = 0
        frmMain.txtTOWt = ""
        frmMain.Show
        frmMain.txtZone1.SetFocus
        frmMain.txtZone2.SetFocus
        frmMain.txtZone3.SetFocus
        frmMain.txtZone4.SetFocus
        frmMain.txtZone5.SetFocus
        'frmMain.txtZone6.SetFocus
        frmMain.txtZone1.SetFocus
        Unload Me
    Else
        Call Recall
        cmbHold1.SetFocus
        'Exit Sub
    End If

End Sub

