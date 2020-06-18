VERSION 5.00
Begin VB.Form frmChart 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Chart Data"
   ClientHeight    =   8490
   ClientLeft      =   -3705
   ClientTop       =   -1035
   ClientWidth     =   11880
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
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtTotal 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   3240
      TabIndex        =   27
      Text            =   " "
      Top             =   4800
      Width           =   1175
   End
   Begin VB.TextBox txt6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   3240
      TabIndex        =   26
      Text            =   " "
      Top             =   4320
      Width           =   1175
   End
   Begin VB.TextBox txt5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   3240
      TabIndex        =   25
      Text            =   " "
      Top             =   3840
      Width           =   1175
   End
   Begin VB.TextBox txt4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   3240
      TabIndex        =   24
      Text            =   " "
      Top             =   3360
      Width           =   1175
   End
   Begin VB.TextBox txt3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   3240
      TabIndex        =   23
      Text            =   " "
      Top             =   2880
      Width           =   1175
   End
   Begin VB.TextBox txt2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   3240
      TabIndex        =   22
      Text            =   " "
      Top             =   2400
      Width           =   1175
   End
   Begin VB.TextBox txt1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   3240
      TabIndex        =   21
      Text            =   " "
      Top             =   1920
      Width           =   1175
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "&Back"
      Height          =   495
      Left            =   6960
      TabIndex        =   20
      Top             =   6240
      Width           =   975
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8400
      TabIndex        =   1
      Top             =   6240
      Width           =   975
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      TabIndex        =   0
      Top             =   6240
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   5475
      Left            =   4680
      Picture         =   "frmChart.frx":0000
      Top             =   240
      Width           =   5820
   End
   Begin VB.Label Label29 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "3400"
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
      Left            =   1800
      TabIndex        =   19
      Top             =   4800
      Width           =   1170
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Total"
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
      Left            =   960
      TabIndex        =   18
      Top             =   4800
      Width           =   735
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "320"
      Height          =   255
      Left            =   1800
      TabIndex        =   17
      Top             =   4320
      Width           =   1170
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "1270"
      Height          =   255
      Left            =   1800
      TabIndex        =   16
      Top             =   3840
      Width           =   1170
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "1380"
      Height          =   255
      Left            =   1800
      TabIndex        =   15
      Top             =   3360
      Width           =   1170
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "1900"
      Height          =   255
      Left            =   1800
      TabIndex        =   14
      Top             =   2880
      Width           =   1170
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "3100"
      Height          =   255
      Left            =   1800
      TabIndex        =   13
      Top             =   2400
      Width           =   1170
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "1780"
      Height          =   255
      Left            =   1800
      TabIndex        =   12
      Top             =   1920
      Width           =   1170
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "6"
      Height          =   255
      Left            =   1200
      TabIndex        =   11
      Top             =   4320
      Width           =   255
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "5"
      Height          =   255
      Left            =   1200
      TabIndex        =   10
      Top             =   3840
      Width           =   255
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "4"
      Height          =   255
      Left            =   1200
      TabIndex        =   9
      Top             =   3360
      Width           =   255
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "3"
      Height          =   255
      Left            =   1200
      TabIndex        =   8
      Top             =   2880
      Width           =   255
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "2"
      Height          =   255
      Left            =   1200
      TabIndex        =   7
      Top             =   2400
      Width           =   255
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "1"
      Height          =   255
      Left            =   1200
      TabIndex        =   6
      Top             =   1920
      Width           =   255
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Item Weight"
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
      Left            =   3240
      TabIndex        =   5
      Top             =   1080
      Width           =   1170
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Max Load in Kgs"
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
      Left            =   1800
      TabIndex        =   4
      Top             =   1080
      Width           =   1170
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Zone"
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
      Left            =   1080
      TabIndex        =   3
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "           C  A   B   I   N       C  A   R  G  O"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   480
      TabIndex        =   2
      Top             =   1080
      Width           =   375
   End
End
Attribute VB_Name = "frmChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBack_Click()
    frmMain.Show
    Unload Me
End Sub

Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdPrint_Click()

     frmPrint.Show
'    Printer.Orientation = "Landscape"
'    cmdPrint.Visible = False
'    cmdExit.Visible = False
'    cmdBack.Visible = False
'    PrintForm
'    cmdPrint.Visible = True
'    cmdExit.Visible = True
'    cmdBack.Visible = True
End Sub

Private Sub Form_Load()

    txt1 = intK1
    txt2 = intK2
    txt3 = intK3
    txt4 = intK4
    txt5 = intK5
    txt6 = intK6
    txtTotal = intCargoK
    
End Sub

