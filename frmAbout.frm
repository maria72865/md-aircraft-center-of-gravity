VERSION 5.00
Begin VB.Form frmAbout 
   Caption         =   "Info"
   ClientHeight    =   5715
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8220
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5715
   ScaleWidth      =   8220
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdTS 
      Caption         =   "&Contact Tech Support"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3480
      TabIndex        =   10
      Top             =   4800
      Width           =   1335
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "www.pilot-international.com"
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
      TabIndex        =   9
      Top             =   4200
      Width           =   7935
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Valley Center, Kansas 67147"
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
      TabIndex        =   8
      Top             =   2880
      Width           =   7935
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "techsupport@pilot-international.com"
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
      TabIndex        =   7
      Top             =   3840
      Width           =   7935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "P.O. Box 384 "
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
      TabIndex        =   6
      Top             =   2520
      Width           =   7935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Ph - (316) 755-0134  Fax - (316) 755-0136"
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
      TabIndex        =   5
      Top             =   3360
      Width           =   7935
   End
   Begin VB.Image Image1 
      Height          =   1095
      Left            =   4200
      Picture         =   "frmAbout.frx":0000
      Stretch         =   -1  'True
      Top             =   1080
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
      Left            =   1680
      TabIndex        =   4
      Top             =   600
      Width           =   4935
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
      Left            =   2040
      TabIndex        =   3
      Top             =   1080
      Width           =   1815
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
      Left            =   2280
      TabIndex        =   1
      Top             =   1680
      Width           =   1335
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
      Left            =   2280
      TabIndex        =   0
      Top             =   1920
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
      Left            =   2040
      TabIndex        =   2
      Top             =   1440
      Width           =   1695
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdTS_Click()
    frmTS.Show
    Unload Me
End Sub

Private Sub Form_Click()
    frmMain.Show
    Unload Me
End Sub

Private Sub Form_LostFocus()
    frmMain.Show
    Unload Me
End Sub

Private Sub Image1_Click()
    frmMain.Show
    Unload Me
End Sub

Private Sub Label1_Click()
    frmMain.Show
    Unload Me
End Sub

Private Sub Label2_Click()
    frmMain.Show
    Unload Me
End Sub

Private Sub Label26_Click()
    frmMain.Show
    Unload Me
End Sub

Private Sub Label27_Click()
    frmMain.Show
    Unload Me
End Sub

Private Sub Label28_Click()
    frmMain.Show
    Unload Me
End Sub

Private Sub Label29_Click()
    frmMain.Show
    Unload Me
End Sub

Private Sub Label3_Click()
    frmMain.Show
    Unload Me
End Sub

Private Sub Label30_Click()
    frmMain.Show
    Unload Me
End Sub

Private Sub Label4_Click()
    frmMain.Show
    Unload Me
End Sub

Private Sub Label5_Click()
    frmMain.Show
    Unload Me
End Sub
