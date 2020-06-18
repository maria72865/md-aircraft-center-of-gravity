VERSION 5.00
Begin VB.Form frmBurnCalc 
   Caption         =   "Fuel"
   ClientHeight    =   4335
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4485
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4335
   ScaleWidth      =   4485
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command10 
      Caption         =   "0"
      Height          =   495
      Left            =   1680
      TabIndex        =   13
      Top             =   3600
      Width           =   495
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "&Enter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   12
      Top             =   1680
      Width           =   855
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   11
      Top             =   840
      Width           =   855
   End
   Begin VB.CommandButton Command9 
      Caption         =   "9"
      Height          =   495
      Left            =   2400
      TabIndex        =   10
      Top             =   2880
      Width           =   495
   End
   Begin VB.CommandButton Command8 
      Caption         =   "8"
      Height          =   495
      Left            =   1680
      TabIndex        =   9
      Top             =   2880
      Width           =   495
   End
   Begin VB.CommandButton Command7 
      Caption         =   "7"
      Height          =   495
      Left            =   960
      TabIndex        =   8
      Top             =   2880
      Width           =   495
   End
   Begin VB.CommandButton Command6 
      Caption         =   "6"
      Height          =   495
      Left            =   2400
      TabIndex        =   7
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton Command5 
      Caption         =   "5"
      Height          =   495
      Left            =   1680
      TabIndex        =   6
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      Caption         =   "4"
      Height          =   495
      Left            =   960
      TabIndex        =   5
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "3"
      Height          =   495
      Left            =   2400
      TabIndex        =   4
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "2"
      Height          =   495
      Left            =   1680
      TabIndex        =   3
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   495
      Left            =   960
      TabIndex        =   2
      Top             =   1440
      Width           =   495
   End
   Begin VB.TextBox txtFOB 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Expected Trip Fuel Use in Pounds:"
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
      TabIndex        =   1
      Top             =   240
      Width           =   4215
   End
End
Attribute VB_Name = "frmBurnCalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClear_Click()
    txtFOB = ""
End Sub

Private Sub cmdDone_Click()
    
    If Val(txtFOB) < 35 Then
        Call MsgBox("Fuel Weight must be at least 35 pounds.", vbOKOnly, "Invalid Weight")
        txtFOB = ""
        Exit Sub
    ElseIf Val(txtFOB) > 2224 Then
        Call MsgBox("Fuel Weight must not exceed 2224 pounds.", vbOKOnly, "Invalid Weight")
        txtFOB = ""
        Exit Sub
    Else
        frmMain.txtFuel = txtFOB
        frmMain.txtFuel.SetFocus
        frmMain.txtZone1.SetFocus
        frmMain.Show
        Unload Me
    End If
    
End Sub

Private Sub Command1_Click()
    txtFOB = txtFOB & "1"
End Sub

Private Sub Command10_Click()
    txtFOB = txtFOB & "0"
End Sub

Private Sub Command2_Click()
    txtFOB = txtFOB & "2"
End Sub

Private Sub Command3_Click()
    txtFOB = txtFOB & "3"
End Sub

Private Sub Command4_Click()
    txtFOB = txtFOB & "4"
End Sub

Private Sub Command5_Click()
    txtFOB = txtFOB & "5"
End Sub

Private Sub Command6_Click()
    txtFOB = txtFOB & "6"
End Sub

Private Sub Command7_Click()
    txtFOB = txtFOB & "7"
End Sub

Private Sub Command8_Click()
    txtFOB = txtFOB & "8"
End Sub

Private Sub Command9_Click()
    txtFOB = txtFOB & "9"
End Sub


