VERSION 5.00
Begin VB.Form frm208BInt 
   Caption         =   "208B International Seating"
   ClientHeight    =   8070
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11475
   LinkTopic       =   "Form1"
   ScaleHeight     =   8070
   ScaleWidth      =   11475
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chk9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6315
      TabIndex        =   36
      Top             =   5040
      Visible         =   0   'False
      Width           =   385
   End
   Begin VB.CheckBox chk7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   35
      Top             =   5040
      Visible         =   0   'False
      Width           =   385
   End
   Begin VB.CheckBox chk5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   34
      Top             =   5040
      Visible         =   0   'False
      Width           =   385
   End
   Begin VB.CheckBox chk3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   33
      Top             =   5040
      Visible         =   0   'False
      Width           =   385
   End
   Begin VB.CheckBox chk10 
      BackColor       =   &H00FFFFFF&
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6255
      TabIndex        =   32
      Top             =   3480
      Visible         =   0   'False
      Width           =   535
   End
   Begin VB.CheckBox chk8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5265
      TabIndex        =   31
      Top             =   3480
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.CheckBox chk6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4215
      TabIndex        =   30
      Top             =   3480
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.CheckBox chk4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   29
      Top             =   3480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CheckBox chkS11 
      BackColor       =   &H00FFFFFF&
      Caption         =   "11"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5760
      TabIndex        =   28
      Top             =   3480
      Visible         =   0   'False
      Width           =   565
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   495
      Left            =   5280
      TabIndex        =   27
      Top             =   7200
      Width           =   1095
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   7080
      TabIndex        =   18
      Top             =   7200
      Width           =   1095
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "&Done"
      Height          =   495
      Left            =   3480
      TabIndex        =   17
      Top             =   7200
      Width           =   1095
   End
   Begin VB.CheckBox chkD6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4560
      TabIndex        =   16
      Top             =   3480
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.CheckBox chkD3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   15
      Top             =   5160
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.CheckBox chkD9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6960
      TabIndex        =   14
      Top             =   5160
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.CheckBox chkD7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5760
      TabIndex        =   12
      Top             =   5160
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.CheckBox chkD8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5760
      TabIndex        =   11
      Top             =   3480
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.CheckBox chkD4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   10
      Top             =   3480
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.CheckBox chkD5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4560
      TabIndex        =   9
      Top             =   5160
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.CheckBox chkS4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3390
      TabIndex        =   8
      Top             =   4080
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.CheckBox chkS6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5205
      TabIndex        =   7
      Top             =   5160
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.CheckBox chkS8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4560
      TabIndex        =   6
      Top             =   3480
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.CheckBox chkS10 
      BackColor       =   &H00FFFFFF&
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5760
      TabIndex        =   5
      Top             =   4080
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.CheckBox chkS3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4020
      TabIndex        =   4
      Top             =   5160
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.CheckBox chkS5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3390
      TabIndex        =   3
      Top             =   3480
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.CheckBox chkS7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4560
      TabIndex        =   2
      Top             =   4080
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.CheckBox chkS9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6405
      TabIndex        =   1
      Top             =   5160
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.ComboBox cmbAC 
      Height          =   315
      Left            =   7680
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CheckBox chkD10 
      BackColor       =   &H00FFFFFF&
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6960
      TabIndex        =   13
      Top             =   3480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image7 
      Height          =   1875
      Left            =   3120
      Picture         =   "frm208BInt.frx":0000
      Top             =   4560
      Width           =   5250
   End
   Begin VB.Image Image5 
      Height          =   1710
      Left            =   5880
      Picture         =   "frm208BInt.frx":733F
      Top             =   2760
      Width           =   4755
   End
   Begin VB.Image Image4 
      Height          =   1710
      Left            =   360
      Picture         =   "frm208BInt.frx":D8BF
      Top             =   2760
      Width           =   5250
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Where will passengers be seated?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   26
      Top             =   2280
      Visible         =   0   'False
      Width           =   11415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Choose Aircraft Seating Option:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   25
      Top             =   2280
      Width           =   11415
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "CESSNA CARAVAN 208B PASSENGER SEATING ARRANGEMENT"
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
      Left            =   6480
      TabIndex        =   24
      Top             =   480
      Width           =   4095
   End
   Begin VB.Image Image3 
      Height          =   1095
      Left            =   3360
      Picture         =   "frm208BInt.frx":1461E
      Stretch         =   -1  'True
      Top             =   720
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
      Left            =   840
      TabIndex        =   23
      Top             =   240
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
      Left            =   1200
      TabIndex        =   22
      Top             =   720
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
      Left            =   1440
      TabIndex        =   20
      Top             =   1320
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
      Left            =   1440
      TabIndex        =   19
      Top             =   1560
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
      Left            =   1200
      TabIndex        =   21
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Image Image2 
      Height          =   2865
      Left            =   960
      Picture         =   "frm208BInt.frx":1AD91
      Top             =   3000
      Visible         =   0   'False
      Width           =   9405
   End
   Begin VB.Image Image8 
      Height          =   2880
      Left            =   960
      Picture         =   "frm208BInt.frx":24C06
      Top             =   3000
      Visible         =   0   'False
      Width           =   9405
   End
   Begin VB.Image Image1 
      Height          =   2550
      Left            =   960
      Picture         =   "frm208BInt.frx":2F70F
      Top             =   3000
      Visible         =   0   'False
      Width           =   9450
   End
End
Attribute VB_Name = "frm208BInt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbAC_Click()

    If cmbAC = "208A, Singles" Then
        Image1.Visible = True
        chkS3.Visible = True
        chkS4.Visible = True
        chkS5.Visible = True
        chkS6.Visible = True
        chkS7.Visible = True
        chkS8.Visible = True
        chkS9.Visible = True
        chkS10.Visible = True
        chkS11.Visible = True
        Image2.Visible = False
        chkD3.Visible = False
        chkD4.Visible = False
        chkD5.Visible = False
        chkD6.Visible = False
        chkD7.Visible = False
        chkD8.Visible = False
        chkD9.Visible = False
        chkD10.Visible = False
    ElseIf cmbAC = "208A, Singles and Doubles" Then
        Image2.Visible = True
        chkD3.Visible = True
        chkD4.Visible = True
        chkD5.Visible = True
        chkD6.Visible = True
        chkD7.Visible = True
        chkD8.Visible = True
        chkD9.Visible = True
        chkD10.Visible = True
        Image1.Visible = False
        chkS3.Visible = False
        chkS4.Visible = False
        chkS5.Visible = False
        chkS6.Visible = False
        chkS7.Visible = False
        chkS8.Visible = False
        chkS9.Visible = False
        chkS10.Visible = False
        chkS11.Visible = False
    End If
    Label2.Visible = True

End Sub

Private Sub cmdClear_Click()
    
    Image1.Visible = False
    Image4.Visible = True
    Image5.Visible = True
    'Image6.Visible = True
    Image7.Visible = True
    Label1.Visible = True
    Label2.Visible = False
    chkS3.Visible = False
    chkS4.Visible = False
    chkS5.Visible = False
    chkS6.Visible = False
    chkS7.Visible = False
    chkS8.Visible = False
    chkS9.Visible = False
    chkS10.Visible = False
    chkS11.Visible = False
    Image2.Visible = False
    chkD3.Visible = False
    chkD4.Visible = False
    chkD5.Visible = False
    chkD6.Visible = False
    chkD7.Visible = False
    chkD8.Visible = False
    chkD9.Visible = False
    chkD10.Visible = False
    Image8.Visible = False
    chk3.Visible = False
    chk4.Visible = False
    chk5.Visible = False
    chk6.Visible = False
    chk7.Visible = False
    chk8.Visible = False
    chk9.Visible = False
    chk10.Visible = False
    chkS3.Value = 0
    chkS4.Value = 0
    chkS5.Value = 0
    chkS6.Value = 0
    chkS7.Value = 0
    chkS8.Value = 0
    chkS9.Value = 0
    chkS10.Value = 0
    chkS11.Value = 0
    chkD3.Value = 0
    chkD4.Value = 0
    chkD5.Value = 0
    chkD6.Value = 0
    chkD7.Value = 0
    chkD8.Value = 0
    chkD9.Value = 0
    chkD10.Value = 0
    chk3.Value = 0
    chk4.Value = 0
    chk5.Value = 0
    chk6.Value = 0
    chk7.Value = 0
    chk8.Value = 0
    chk9.Value = 0
    chk10.Value = 0

End Sub

Private Sub cmdDone_Click()
    frmMain.Show
    If frmMain.Text1 <> "4" Then
        Call MsgBox("Nice try Pal!", vbOKOnly, "No Changes Allowed!")
        End
    Else
        frmMain.txtFltNo.SetFocus
    End If
    Unload Me
End Sub

Private Sub cmdExit_Click()
    End
End Sub

Private Sub Form_Load()

    cmbAC.AddItem "208A, Singles"
    cmbAC.AddItem "208A, Singles and Doubles"
    
End Sub

Private Sub Image4_Click()

    Image1.Visible = True
    Image4.Visible = False
    Image5.Visible = False
    'Image6.Visible = False
    Image7.Visible = False
    Label1.Visible = False
    Label2.Visible = True
    chkS3.Visible = True
    chkS4.Visible = True
    chkS5.Visible = True
    chkS6.Visible = True
    chkS7.Visible = True
    chkS8.Visible = True
    chkS9.Visible = True
    chkS10.Visible = True
    chkS11.Visible = True
    Image2.Visible = False
    chkD3.Visible = False
    chkD4.Visible = False
    chkD5.Visible = False
    chkD6.Visible = False
    chkD7.Visible = False
    chkD8.Visible = False
    chkD9.Visible = False
    chkD10.Visible = False
    Image8.Visible = False
    chk3.Visible = False
    chk4.Visible = False
    chk5.Visible = False
    chk6.Visible = False
    chk7.Visible = False
    chk8.Visible = False
    chk9.Visible = False
    chk10.Visible = False
    
End Sub

Private Sub Image5_Click()

    Image2.Visible = True
    Image4.Visible = False
    Image5.Visible = False
    'Image6.Visible = False
    Image7.Visible = False
    Label1.Visible = False
    Label2.Visible = True
    chkD3.Visible = True
    chkD4.Visible = True
    chkD5.Visible = True
    chkD6.Visible = True
    chkD7.Visible = True
    chkD8.Visible = True
    chkD9.Visible = True
    chkD10.Visible = True
    Image1.Visible = False
    chkS3.Visible = False
    chkS4.Visible = False
    chkS5.Visible = False
    chkS6.Visible = False
    chkS7.Visible = False
    chkS8.Visible = False
    chkS9.Visible = False
    chkS10.Visible = False
    chkS11.Visible = False
    Image8.Visible = False
    chk3.Visible = False
    chk4.Visible = False
    chk5.Visible = False
    chk6.Visible = False
    chk7.Visible = False
    chk8.Visible = False
    chk9.Visible = False
    chk10.Visible = False

End Sub



Private Sub Image7_Click()
    Image8.Visible = True
    chk3.Visible = True
    chk4.Visible = True
    chk5.Visible = True
    chk6.Visible = True
    chk7.Visible = True
    chk8.Visible = True
    chk9.Visible = True
    chk10.Visible = True
    Image2.Visible = False
    Image4.Visible = False
    Image5.Visible = False
    'Image6.Visible = False
    Image7.Visible = False
    Label1.Visible = False
    Label2.Visible = True
    chkD3.Visible = False
    chkD4.Visible = False
    chkD5.Visible = False
    chkD6.Visible = False
    chkD7.Visible = False
    chkD8.Visible = False
    chkD9.Visible = False
    chkD10.Visible = False
    Image1.Visible = False
    chkS3.Visible = False
    chkS4.Visible = False
    chkS5.Visible = False
    chkS6.Visible = False
    chkS7.Visible = False
    chkS8.Visible = False
    chkS9.Visible = False
    chkS10.Visible = False
    chkS11.Visible = False
   
End Sub

