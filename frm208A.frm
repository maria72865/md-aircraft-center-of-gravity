VERSION 5.00
Begin VB.Form frm208A 
   Caption         =   "208A Seating"
   ClientHeight    =   7725
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10590
   LinkTopic       =   "Form1"
   ScaleHeight     =   7725
   ScaleWidth      =   10590
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   28
      Text            =   "Text1"
      Top             =   1920
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   495
      Left            =   4800
      TabIndex        =   27
      Top             =   7080
      Width           =   1095
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   6600
      TabIndex        =   18
      Top             =   7080
      Width           =   1095
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "&Done"
      Height          =   495
      Left            =   3000
      TabIndex        =   17
      Top             =   7080
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
      Left            =   5760
      TabIndex        =   16
      Top             =   5400
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
      Left            =   4320
      TabIndex        =   15
      Top             =   5400
      Visible         =   0   'False
      Width           =   525
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
      Left            =   6360
      TabIndex        =   14
      Top             =   4080
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.CheckBox chkD10 
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
      Left            =   6360
      TabIndex        =   13
      Top             =   3480
      Visible         =   0   'False
      Width           =   570
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
      Left            =   4920
      TabIndex        =   12
      Top             =   4080
      Visible         =   0   'False
      Width           =   570
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
      Left            =   4920
      TabIndex        =   11
      Top             =   3480
      Visible         =   0   'False
      Width           =   570
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
      Left            =   3600
      TabIndex        =   10
      Top             =   4080
      Visible         =   0   'False
      Width           =   570
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
      Left            =   3600
      TabIndex        =   9
      Top             =   3360
      Visible         =   0   'False
      Width           =   570
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
      Left            =   3480
      TabIndex        =   8
      Top             =   3360
      Visible         =   0   'False
      Width           =   570
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
      Left            =   4590
      TabIndex        =   7
      Top             =   3360
      Visible         =   0   'False
      Width           =   570
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
      Left            =   5760
      TabIndex        =   6
      Top             =   3360
      Visible         =   0   'False
      Width           =   570
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
      Left            =   6840
      TabIndex        =   5
      Top             =   3360
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
      Left            =   3480
      TabIndex        =   4
      Top             =   5160
      Visible         =   0   'False
      Width           =   570
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
      Left            =   4590
      TabIndex        =   3
      Top             =   5160
      Visible         =   0   'False
      Width           =   570
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
      Left            =   5760
      TabIndex        =   2
      Top             =   5160
      Visible         =   0   'False
      Width           =   570
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
      Left            =   6960
      TabIndex        =   1
      Top             =   5160
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.ComboBox cmbAC 
      Height          =   315
      Left            =   7680
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "All Passenger Weights are 175 pounds by default.  This can be changed on Main Form after clicking ""Done""."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      TabIndex        =   29
      Top             =   6360
      Visible         =   0   'False
      Width           =   10455
   End
   Begin VB.Image Image6 
      Height          =   1830
      Left            =   3120
      Picture         =   "frm208A.frx":0000
      Top             =   4800
      Width           =   4500
   End
   Begin VB.Image Image5 
      Height          =   1830
      Left            =   5640
      Picture         =   "frm208A.frx":527C
      Top             =   2760
      Width           =   4500
   End
   Begin VB.Image Image4 
      Height          =   1830
      Left            =   600
      Picture         =   "frm208A.frx":BF8A
      Top             =   2760
      Width           =   4500
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
      Width           =   10575
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
      Width           =   10455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "CESSNA CARAVAN 208A PASSENGER SEATING ARRANGEMENT"
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
      Left            =   6000
      TabIndex        =   24
      Top             =   480
      Width           =   4095
   End
   Begin VB.Image Image3 
      Height          =   1095
      Left            =   3360
      Picture         =   "frm208A.frx":12F0D
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
      Height          =   3510
      Left            =   480
      Picture         =   "frm208A.frx":19680
      Top             =   2760
      Visible         =   0   'False
      Width           =   9750
   End
   Begin VB.Image Image1 
      Height          =   3255
      Left            =   480
      Picture         =   "frm208A.frx":24838
      Top             =   2760
      Visible         =   0   'False
      Width           =   9750
   End
End
Attribute VB_Name = "frm208A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdClear_Click()
    
    Image1.Visible = False
    Image2.Visible = False
    Image4.Visible = True
    Image5.Visible = True
    Image6.Visible = True
    Label1.Visible = True
    Label2.Visible = False
    Label4.Visible = False
    chkS3.Visible = False
    chkS4.Visible = False
    chkS5.Visible = False
    chkS6.Visible = False
    chkS7.Visible = False
    chkS8.Visible = False
    chkS9.Visible = False
    chkS10.Visible = False
    chkD3.Visible = False
    chkD4.Visible = False
    chkD5.Visible = False
    chkD6.Visible = False
    chkD7.Visible = False
    chkD8.Visible = False
    chkD9.Visible = False
    chkD10.Visible = False
    chkS3.Value = 0
    chkS4.Value = 0
    chkS5.Value = 0
    chkS6.Value = 0
    chkS7.Value = 0
    chkS8.Value = 0
    chkS9.Value = 0
    chkS10.Value = 0
    chkD3.Value = 0
    chkD4.Value = 0
    chkD5.Value = 0
    chkD6.Value = 0
    chkD7.Value = 0
    chkD8.Value = 0
    chkD9.Value = 0
    chkD10.Value = 0

End Sub

Private Sub cmdDone_Click()

If Image2.Visible = True Then

    If chkD3.Value = 1 Then
        intPass3 = 175
    Else
        intPass3 = 0
    End If
    
    If chkD4.Value = 1 Then
        intPass4 = 175
    Else
        intPass4 = 0
    End If
    
    If chkD5.Value = 1 Then
        intPass5 = 175
    Else
        intPass5 = 0
    End If
    
    If chkD6.Value = 1 Then
        intPass6 = 175
    Else
        intPass6 = 0
    End If
    
    If chkD7.Value = 1 Then
        intPass7 = 175
    Else
        intPass7 = 0
    End If
    
    If chkD8.Value = 1 Then
        intPass8 = 175
    Else
        intPass8 = 0
    End If
    
    If chkD9.Value = 1 Then
        intPass9 = 175
    Else
        intPass9 = 0
    End If
    
    If chkD10.Value = 1 Then
        intPass10 = 175
    Else
        intPass10 = 0
    End If
    
ElseIf Image1.Visible = True Then
    
    If chkS3.Value = 1 Then
        intPass3 = 175
    Else
        intPass3 = 0
    End If

    If chkS4.Value = 1 Then
        intPass4 = 175
    Else
        intPass4 = 0
    End If
    
    If chkS5.Value = 1 Then
        intPass5 = 175
    Else
        intPass5 = 0
    End If

    If chkS6.Value = 1 Then
        intPass6 = 175
    Else
        intPass6 = 0
    End If
    
    If chkS7.Value = 1 Then
        intPass7 = 175
    Else
        intPass7 = 0
    End If

    If chkS8.Value = 1 Then
        intPass8 = 175
    Else
        intPass8 = 0
    End If
    
    If chkS9.Value = 1 Then
        intPass9 = 175
    Else
        intPass9 = 0
    End If
    
    If chkS10.Value = 1 Then
        intPass10 = 175
    Else
        intPass10 = 0
    End If
    
End If

    frmMain.Show
    
    If Image1.Visible = True Then
    
        frmMain.txtSeats = "208A10U"
        
        frmMain.Label2.Caption = "Seats 3/4"
        frmMain.Label2.FontSize = 9
        frmMain.txtZone1 = intPass3 + intPass4
        frmMain.txtZone1 = Format(frmMain.txtZone1, "Fixed")
        
        frmMain.Label4.Caption = "Seats 5/6"
        frmMain.Label4.FontSize = 9
        frmMain.txtZone2 = intPass5 + intPass6
        frmMain.txtZone2 = Format(frmMain.txtZone2, "Fixed")
        
        frmMain.Label5.Caption = "Seats 7/8"
        frmMain.Label5.FontSize = 9
        frmMain.txtZone3 = intPass7 + intPass8
        frmMain.txtZone3 = Format(frmMain.txtZone3, "Fixed")
        
        frmMain.Label6.Caption = "Seat 9"
        frmMain.Label6.FontSize = 9
        frmMain.txtZone4 = intPass9
        frmMain.txtZone4 = Format(frmMain.txtZone4, "Fixed")
        
        frmMain.Label7.Caption = "Seat 10"
        frmMain.Label7.FontSize = 9
        frmMain.txtZone5 = intPass10
        frmMain.txtZone5 = Format(frmMain.txtZone5, "Fixed")
        
        frmMain.Label8.Caption = "Baggage"
        frmMain.Label8.FontSize = 9
        frmMain.txtZone6 = 0
        frmMain.txtZone6 = Format(frmMain.txtZone6, "Fixed")
        
    ElseIf Image2.Visible = True Then
    
        frmMain.txtSeats = "208A10C"
    
        frmMain.Label2.Caption = "Seat 3"
        frmMain.Label2.FontSize = 9
        frmMain.txtZone1 = intPass3
        frmMain.txtZone1 = Format(frmMain.txtZone1, "Fixed")
        
        frmMain.Label4.Caption = "Seats 4/5"
        frmMain.Label4.FontSize = 9
        frmMain.txtZone2 = intPass5 + intPass4
        frmMain.txtZone2 = Format(frmMain.txtZone2, "Fixed")
        
        frmMain.Label5.Caption = "Seat 6"
        frmMain.Label5.FontSize = 9
        frmMain.txtZone3 = intPass6
        frmMain.txtZone3 = Format(frmMain.txtZone3, "Fixed")
        
        frmMain.Label6.Caption = "Seats 7/8"
        frmMain.Label6.FontSize = 9
        frmMain.txtZone4 = intPass7 + intPass8
        frmMain.txtZone4 = Format(frmMain.txtZone4, "Fixed")
        
        frmMain.Label7.Caption = "Seats 9/10"
        frmMain.Label7.FontSize = 9
        frmMain.txtZone5 = intPass9 + intPass10
        frmMain.txtZone5 = Format(frmMain.txtZone5, "Fixed")
        
        frmMain.Label8.Caption = "Baggage"
        frmMain.Label8.FontSize = 9
        frmMain.txtZone6 = 0
        frmMain.txtZone6 = Format(frmMain.txtZone6, "Fixed")
        
    End If
    
    'If frmMain.Text1 <> "4" Then
    '    Call MsgBox("Nice try Pal!", vbOKOnly, "No Changes Allowed!")
    '    End
    'Else
        
        If Label3 = "CESSNA 208 AMPHIB PASSENGER SEATING ARRANGEMENT" Then
            frmMain.txtSeats = "208Amphib"
        End If
        
        frmMain.txtZone1.SetFocus
        frmMain.txtZone2.SetFocus
        frmMain.txtZone3.SetFocus
        frmMain.txtZone4.SetFocus
        frmMain.txtZone5.SetFocus
        'frmMain.txtZone6.SetFocus
        frmMain.txtFltNo.SetFocus
    'End If
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
    Image6.Visible = False
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
    Image2.Visible = False
    chkD3.Visible = False
    chkD4.Visible = False
    chkD5.Visible = False
    chkD6.Visible = False
    chkD7.Visible = False
    chkD8.Visible = False
    chkD9.Visible = False
    chkD10.Visible = False
    Label4.Visible = True
    
End Sub

Private Sub Image5_Click()

    Image2.Visible = True
    Image4.Visible = False
    Image5.Visible = False
    Image6.Visible = False
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
    Label4.Visible = True

End Sub

Private Sub Image6_Click()

    frmMain.Show
    frmMain.Label2 = "Hold 1"
    frmMain.Label4 = "Hold 2"
    frmMain.Label5 = "Hold 3"
    frmMain.Label6 = "Hold 4"
    frmMain.Label7 = "Hold 5"
    frmMain.Label8 = "Hold 6"
    frmMain.Label8.Visible = True
    frmMain.txtZone6.Visible = True
    frmMain.Label22.Visible = True
    frmMain.txtZone1 = "0.00"
    frmMain.txtZone2 = "0.00"
    frmMain.txtZone3 = "0.00"
    frmMain.txtZone4 = "0.00"
    frmMain.txtZone5 = "0.00"
    frmMain.txtZone6 = "0.00"
    frmMain.txtSeats = "208AN"
    frmMain.txtFltNo.SetFocus
    Unload Me

End Sub
