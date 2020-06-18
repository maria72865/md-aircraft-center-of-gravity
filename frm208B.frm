VERSION 5.00
Begin VB.Form frm208B 
   Caption         =   "208B Seating"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11025
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   11025
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chk1412 
      BackColor       =   &H00FFFFFF&
      Caption         =   "12"
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
      Left            =   6290
      TabIndex        =   69
      Top             =   5160
      Visible         =   0   'False
      Width           =   475
   End
   Begin VB.CheckBox chk1413 
      BackColor       =   &H00FFFFFF&
      Caption         =   "13"
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
      Left            =   6300
      TabIndex        =   68
      Top             =   4080
      Visible         =   0   'False
      Width           =   475
   End
   Begin VB.CheckBox chk1414 
      BackColor       =   &H00FFFFFF&
      Caption         =   "14"
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
      Left            =   6300
      TabIndex        =   67
      Top             =   3600
      Visible         =   0   'False
      Width           =   475
   End
   Begin VB.CheckBox chk149 
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
      Left            =   5080
      TabIndex        =   66
      Top             =   5160
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.CheckBox chk1410 
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
      Left            =   5080
      TabIndex        =   65
      Top             =   4080
      Visible         =   0   'False
      Width           =   475
   End
   Begin VB.CheckBox chk1411 
      BackColor       =   &H00FFFFFF&
      Caption         =   "11"
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
      Left            =   5080
      TabIndex        =   64
      Top             =   3600
      Visible         =   0   'False
      Width           =   475
   End
   Begin VB.CheckBox chk148 
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
      Left            =   3920
      TabIndex        =   63
      Top             =   3600
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.CheckBox chk147 
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
      Left            =   3920
      TabIndex        =   62
      Top             =   4080
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.CheckBox chk146 
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
      Left            =   3930
      TabIndex        =   61
      Top             =   5160
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.CheckBox chk145 
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
      Left            =   2760
      TabIndex        =   60
      Top             =   3600
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.CheckBox chk144 
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
      Left            =   2760
      TabIndex        =   59
      Top             =   4080
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.CheckBox chk143 
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
      Left            =   2760
      TabIndex        =   58
      Top             =   5160
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.CheckBox chk1311 
      BackColor       =   &H00FFFFFF&
      Caption         =   "11"
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
      Left            =   8640
      TabIndex        =   57
      Top             =   4920
      Visible         =   0   'False
      Width           =   475
   End
   Begin VB.CheckBox chk1312 
      BackColor       =   &H00FFFFFF&
      Caption         =   "12"
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
      Left            =   8640
      TabIndex        =   56
      Top             =   4320
      Visible         =   0   'False
      Width           =   475
   End
   Begin VB.CheckBox chk1313 
      BackColor       =   &H00FFFFFF&
      Caption         =   "13"
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
      Left            =   8640
      TabIndex        =   55
      Top             =   3720
      Visible         =   0   'False
      Width           =   475
   End
   Begin VB.CheckBox chk139 
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
      Left            =   6480
      TabIndex        =   54
      Top             =   5160
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.CheckBox chk1310 
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
      Left            =   6480
      TabIndex        =   53
      Top             =   3480
      Visible         =   0   'False
      Width           =   475
   End
   Begin VB.CheckBox chk137 
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
      Left            =   5290
      TabIndex        =   52
      Top             =   5160
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.CheckBox chk138 
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
      Left            =   5290
      TabIndex        =   51
      Top             =   3480
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.CheckBox chk135 
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
      Left            =   4080
      TabIndex        =   50
      Top             =   5160
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.CheckBox chk136 
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
      Left            =   4080
      TabIndex        =   49
      Top             =   3480
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.CheckBox chk134 
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
      Left            =   2880
      TabIndex        =   48
      Top             =   3480
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.CheckBox chk133 
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
      Left            =   2880
      TabIndex        =   47
      Top             =   5160
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.CheckBox chk1212 
      BackColor       =   &H00FFFFFF&
      Caption         =   "12"
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
      Left            =   7290
      TabIndex        =   46
      Top             =   5025
      Visible         =   0   'False
      Width           =   475
   End
   Begin VB.CheckBox chk1211 
      BackColor       =   &H00FFFFFF&
      Caption         =   "11"
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
      Left            =   7300
      TabIndex        =   45
      Top             =   3480
      Visible         =   0   'False
      Width           =   475
   End
   Begin VB.CheckBox chk1210 
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
      Left            =   6240
      TabIndex        =   43
      Top             =   3480
      Visible         =   0   'False
      Width           =   475
   End
   Begin VB.CheckBox chk127 
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
      Left            =   5170
      TabIndex        =   42
      Top             =   5025
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.CheckBox chk128 
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
      Left            =   5190
      TabIndex        =   41
      Top             =   3480
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.CheckBox chk126 
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
      Left            =   4120
      TabIndex        =   39
      Top             =   3480
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.CheckBox chk123 
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
      Left            =   3030
      TabIndex        =   38
      Top             =   5025
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.CheckBox chk124 
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
      Left            =   3030
      TabIndex        =   37
      Top             =   3480
      Visible         =   0   'False
      Width           =   450
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
      Left            =   5760
      TabIndex        =   28
      Top             =   4080
      Visible         =   0   'False
      Width           =   550
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   495
      Left            =   4920
      TabIndex        =   27
      Top             =   7440
      Width           =   1095
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   6480
      TabIndex        =   18
      Top             =   7440
      Width           =   1095
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "&Done"
      Height          =   495
      Left            =   3360
      TabIndex        =   17
      Top             =   7440
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
         Name            =   "Arial Narrow"
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
      Top             =   3480
      Visible         =   0   'False
      Width           =   555
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
   Begin VB.CheckBox chkS11 
      BackColor       =   &H00FFFFFF&
      Caption         =   "11"
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
      Left            =   6395
      TabIndex        =   1
      Top             =   5160
      Visible         =   0   'False
      Width           =   515
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
   Begin VB.CheckBox chk125 
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
      Left            =   4120
      TabIndex        =   40
      Top             =   5025
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.CheckBox chk129 
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
      Left            =   6255
      TabIndex        =   44
      Top             =   5025
      Visible         =   0   'False
      Width           =   450
   End
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
      Top             =   2040
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
      Picture         =   "frm208B.frx":0000
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
   Begin VB.Image Image11 
      Height          =   4965
      Left            =   9480
      Picture         =   "frm208B.frx":6773
      Top             =   2280
      Width           =   1350
   End
   Begin VB.Image Image10 
      Height          =   4965
      Left            =   7920
      Picture         =   "frm208B.frx":C7BA
      Top             =   2280
      Width           =   1350
   End
   Begin VB.Image Image6 
      Height          =   4965
      Left            =   120
      Picture         =   "frm208B.frx":12590
      Top             =   2280
      Width           =   1350
   End
   Begin VB.Image Image9 
      Height          =   4965
      Left            =   6360
      Picture         =   "frm208B.frx":170BE
      Top             =   2280
      Width           =   1350
   End
   Begin VB.Image Image7 
      Height          =   4965
      Left            =   3240
      Picture         =   "frm208B.frx":1CC7D
      Top             =   2280
      Width           =   1350
   End
   Begin VB.Image Image5 
      Height          =   4965
      Left            =   1680
      Picture         =   "frm208B.frx":22435
      Top             =   2280
      Width           =   1350
   End
   Begin VB.Image Image4 
      Height          =   4965
      Left            =   4800
      Picture         =   "frm208B.frx":2820E
      Top             =   2280
      Width           =   1350
   End
   Begin VB.Image Image2 
      Height          =   2865
      Left            =   960
      Picture         =   "frm208B.frx":2DA97
      Top             =   3000
      Visible         =   0   'False
      Width           =   9405
   End
   Begin VB.Image Image8 
      Height          =   2880
      Left            =   960
      Picture         =   "frm208B.frx":3776B
      Top             =   3000
      Visible         =   0   'False
      Width           =   9405
   End
   Begin VB.Image Image1 
      Height          =   2880
      Left            =   960
      Picture         =   "frm208B.frx":416FD
      Top             =   3000
      Visible         =   0   'False
      Width           =   9405
   End
   Begin VB.Image Image14 
      Height          =   2880
      Left            =   360
      Picture         =   "frm208B.frx":4B59F
      Top             =   3000
      Visible         =   0   'False
      Width           =   9405
   End
   Begin VB.Image Image13 
      Height          =   2865
      Left            =   480
      Picture         =   "frm208B.frx":560A8
      Top             =   3000
      Visible         =   0   'False
      Width           =   9405
   End
   Begin VB.Image Image12 
      Height          =   2550
      Left            =   600
      Picture         =   "frm208B.frx":5FF1D
      Top             =   3120
      Visible         =   0   'False
      Width           =   9450
   End
End
Attribute VB_Name = "frm208B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdClear_Click()
    
    Image1.Visible = False
    Image4.Visible = True
    Image5.Visible = True
    Image6.Visible = True
    Image7.Visible = True
    Image9.Visible = True
    Image10.Visible = True
    Image11.Visible = True
    Image12.Visible = False
    Image13.Visible = False
    Image14.Visible = False
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
    chk123.Visible = False
    chk124.Visible = False
    chk125.Visible = False
    chk126.Visible = False
    chk127.Visible = False
    chk128.Visible = False
    chk129.Visible = False
    chk1210.Visible = False
    chk1211.Visible = False
    chk1212.Visible = False
    chk133.Visible = False
    chk134.Visible = False
    chk135.Visible = False
    chk136.Visible = False
    chk137.Visible = False
    chk138.Visible = False
    chk139.Visible = False
    chk1310.Visible = False
    chk1311.Visible = False
    chk1312.Visible = False
    chk1313.Visible = False
    chk143.Visible = False
    chk144.Visible = False
    chk145.Visible = False
    chk146.Visible = False
    chk147.Visible = False
    chk148.Visible = False
    chk149.Visible = False
    chk1410.Visible = False
    chk1411.Visible = False
    chk1412.Visible = False
    chk1413.Visible = False
    chk1414.Visible = False
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
    chk123.Value = 0
    chk124.Value = 0
    chk125.Value = 0
    chk126.Value = 0
    chk127.Value = 0
    chk128.Value = 0
    chk129.Value = 0
    chk1210.Value = 0
    chk1211.Value = 0
    chk1212.Value = 0
    chk133.Value = 0
    chk134.Value = 0
    chk135.Value = 0
    chk136.Value = 0
    chk137.Value = 0
    chk138.Value = 0
    chk139.Value = 0
    chk1310.Value = 0
    chk1311.Value = 0
    chk1312.Value = 0
    chk1313.Value = 0
    chk143.Value = 0
    chk144.Value = 0
    chk145.Value = 0
    chk146.Value = 0
    chk147.Value = 0
    chk148.Value = 0
    chk149.Value = 0
    chk1410.Value = 0
    chk1411.Value = 0
    chk1412.Value = 0
    chk1413.Value = 0
    chk1414.Value = 0
    cmdDone.Top = 7440
    cmdClear.Top = 7440
    cmdExit.Top = 7440
    
End Sub

Private Sub cmdDone_Click()

If Image1.Visible = True Then
    
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
    
    If chkS11.Value = 1 Then
        intPass11 = 175
    Else
        intPass11 = 0
    End If
    
    intPass12 = 0
    intPass13 = 0
    intPass14 = 0

ElseIf Image2.Visible = True Then

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
    
    intPass11 = 0
    intPass12 = 0
    intPass13 = 0
    intPass14 = 0
    
ElseIf Image8.Visible = True Then

    If chk3.Value = 1 Then
        intPass3 = 175
    Else
        intPass3 = 0
    End If
    
    If chk4.Value = 1 Then
        intPass4 = 175
    Else
        intPass4 = 0
    End If
    
    If chk5.Value = 1 Then
        intPass5 = 175
    Else
        intPass5 = 0
    End If
    
    If chk6.Value = 1 Then
        intPass6 = 175
    Else
        intPass6 = 0
    End If
    
    If chk7.Value = 1 Then
        intPass7 = 175
    Else
        intPass7 = 0
    End If
    
    If chk8.Value = 1 Then
        intPass8 = 175
    Else
        intPass8 = 0
    End If
    
    If chk9.Value = 1 Then
        intPass9 = 175
    Else
        intPass9 = 0
    End If
    
    If chk10.Value = 1 Then
        intPass10 = 175
    Else
        intPass10 = 0
    End If
    
    intPass11 = 0
    intPass12 = 0
    intPass13 = 0
    intPass14 = 0

ElseIf Image12.Visible = True Then

    If chk123.Value = 1 Then
        intPass3 = 175
    Else
        intPass3 = 0
    End If
    
    If chk124.Value = 1 Then
        intPass4 = 175
    Else
        intPass4 = 0
    End If
    
    If chk125.Value = 1 Then
        intPass5 = 175
    Else
        intPass5 = 0
    End If
    
    If chk126.Value = 1 Then
        intPass6 = 175
    Else
        intPass6 = 0
    End If
    
    If chk127.Value = 1 Then
        intPass7 = 175
    Else
        intPass7 = 0
    End If
    
    If chk128.Value = 1 Then
        intPass8 = 175
    Else
        intPass8 = 0
    End If
    
    If chk129.Value = 1 Then
        intPass9 = 175
    Else
        intPass9 = 0
    End If
    
    If chk1210.Value = 1 Then
        intPass10 = 175
    Else
        intPass10 = 0
    End If
    
    If chk1211.Value = 1 Then
        intPass11 = 175
    Else
        intPass11 = 0
    End If
    
    If chk1212.Value = 1 Then
        intPass12 = 175
    Else
        intPass12 = 0
    End If
    
    intPass13 = 0
    intPass14 = 0
    
ElseIf Image13.Visible = True Then

    If chk133.Value = 1 Then
        intPass3 = 175
    Else
        intPass3 = 0
    End If
    
    If chk134.Value = 1 Then
        intPass4 = 175
    Else
        intPass4 = 0
    End If
    
    If chk135.Value = 1 Then
        intPass5 = 175
    Else
        intPass5 = 0
    End If
    
    If chk136.Value = 1 Then
        intPass6 = 175
    Else
        intPass6 = 0
    End If
    
    If chk137.Value = 1 Then
        intPass7 = 175
    Else
        intPass7 = 0
    End If
    
    If chk138.Value = 1 Then
        intPass8 = 175
    Else
        intPass8 = 0
    End If
    
    If chk139.Value = 1 Then
        intPass9 = 175
    Else
        intPass9 = 0
    End If
    
    If chk1310.Value = 1 Then
        intPass10 = 175
    Else
        intPass10 = 0
    End If
    
    If chk1311.Value = 1 Then
        intPass11 = 106
    Else
        intPass11 = 0
    End If
    
    If chk1312.Value = 1 Then
        intPass12 = 106
    Else
        intPass12 = 0
    End If
    
    If chk1313.Value = 1 Then
        intPass13 = 106
    Else
        intPass13 = 0
    End If
    
    intPass14 = 0
    
ElseIf Image14.Visible = True Then

    If chk143.Value = 1 Then
        intPass3 = 175
    Else
        intPass3 = 0
    End If
    
    If chk144.Value = 1 Then
        intPass4 = 175
    Else
        intPass4 = 0
    End If
    
    If chk145.Value = 1 Then
        intPass5 = 175
    Else
        intPass5 = 0
    End If
    
    If chk146.Value = 1 Then
        intPass6 = 175
    Else
        intPass6 = 0
    End If
    
    If chk147.Value = 1 Then
        intPass7 = 175
    Else
        intPass7 = 0
    End If
    
    If chk148.Value = 1 Then
        intPass8 = 175
    Else
        intPass8 = 0
    End If
    
    If chk149.Value = 1 Then
        intPass9 = 175
    Else
        intPass9 = 0
    End If
    
    If chk1410.Value = 1 Then
        intPass10 = 175
    Else
        intPass10 = 0
    End If
    
    If chk1411.Value = 1 Then
        intPass11 = 175
    Else
        intPass11 = 0
    End If
    
    If chk1412.Value = 1 Then
        intPass12 = 175
    Else
        intPass12 = 0
    End If
    
    If chk1413.Value = 1 Then
        intPass13 = 175
    Else
        intPass13 = 0
    End If
    
    If chk1414.Value = 1 Then
        intPass14 = 175
    Else
        intPass14 = 0
    End If
    
End If

    frmMain.Show
    
    If Image1.Visible = True Then
    
        frmMain.txtSeats = "208B11C"
        
        frmMain.Label2.Caption = "Seat 3"
        frmMain.Label2.FontSize = 9
        frmMain.txtZone1 = intPass3
        frmMain.txtZone1 = Format(frmMain.txtZone1, "Fixed")
        
        frmMain.Label4.Caption = "Seats 4/5"
        frmMain.Label4.FontSize = 9
        frmMain.txtZone2 = intPass4 + intPass5
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
        
        frmMain.Label8.Caption = "Seat 11"
        frmMain.Label8.FontSize = 9
        frmMain.txtZone6 = intPass11
        frmMain.txtZone6 = Format(frmMain.txtZone6, "Fixed")
        
        frmMain.Label39.Visible = True
        frmMain.Label40.Visible = True
        frmMain.txtSeat11.Visible = True
        frmMain.Label39.Caption = "Baggage"
        frmMain.Label8.FontSize = 9
        frmMain.txtSeat11 = 0
        frmMain.txtSeat11 = Format(frmMain.txtSeat11, "Fixed")
        
    ElseIf Image2.Visible = True Then
    
        frmMain.txtSeats = "208B10C"
    
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
        
        frmMain.Label6.Caption = "Seats 9/10"
        frmMain.Label6.FontSize = 9
        frmMain.txtZone4 = intPass9 + intPass10
        frmMain.txtZone4 = Format(frmMain.txtZone4, "Fixed")
        
        frmMain.Label7.Caption = "Baggage"
        frmMain.Label7.FontSize = 9
        frmMain.txtZone5 = 0
        frmMain.txtZone5 = Format(frmMain.txtZone5, "Fixed")
        
        frmMain.Label8.Visible = False
        frmMain.Label8.FontSize = 9
        frmMain.txtZone6.Visible = False
        frmMain.txtZone6 = Format(frmMain.txtZone6, "Fixed")
        frmMain.Label22.Visible = False
        
        frmMain.Label39.Visible = False
        frmMain.Label40.Visible = False
        frmMain.txtSeat11.Visible = False
        
    ElseIf Image8.Visible = True Then
    
        frmMain.txtSeats = "208B10U"
    
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
        
        frmMain.Label6.Caption = "Seats 9/10"
        frmMain.Label6.FontSize = 9
        frmMain.txtZone4 = intPass9 + intPass10
        frmMain.txtZone4 = Format(frmMain.txtZone4, "Fixed")
        
        frmMain.Label7.Caption = "Baggage"
        frmMain.Label7.FontSize = 9
        frmMain.txtZone5 = 0
        frmMain.txtZone5 = Format(frmMain.txtZone5, "Fixed")
        
        frmMain.Label8.Visible = False
        frmMain.Label8.FontSize = 9
        frmMain.txtZone6.Visible = False
        frmMain.txtZone6 = Format(frmMain.txtZone6, "Fixed")
        frmMain.Label22.Visible = False
        
        frmMain.Label39.Visible = False
        frmMain.Label40.Visible = False
        frmMain.txtSeat11.Visible = False
        
    ElseIf Image12.Visible = True Then
    
        frmMain.txtSeats = "208B12"
    
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
        
        frmMain.Label6.Caption = "Seats 9/10"
        frmMain.Label6.FontSize = 9
        frmMain.txtZone4 = intPass9 + intPass10
        frmMain.txtZone4 = Format(frmMain.txtZone4, "Fixed")
        
        frmMain.Label7.Caption = "Seats 11/12"
        frmMain.Label7.FontSize = 9
        frmMain.txtZone5 = intPass11 + intPass12
        frmMain.txtZone5 = Format(frmMain.txtZone5, "Fixed")
        
        frmMain.Label8.Caption = "Baggage"
        frmMain.Label8.FontSize = 9
        frmMain.txtZone6 = 0
        frmMain.txtZone6 = Format(frmMain.txtZone6, "Fixed")
        
        frmMain.Label39.Visible = False
        frmMain.Label40.Visible = False
        frmMain.txtSeat11.Visible = False
        
    ElseIf Image13.Visible = True Then
    
        frmMain.txtSeats = "208B13"
    
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
        
        frmMain.Label6.Caption = "Seats 9/10"
        frmMain.Label6.FontSize = 9
        frmMain.txtZone4 = intPass9 + intPass10
        frmMain.txtZone4 = Format(frmMain.txtZone4, "Fixed")
        
        frmMain.Label7.Caption = "Seats 11/12/13"
        frmMain.Label7.FontSize = 9
        frmMain.txtZone5 = intPass11 + intPass12 + intPass13
        frmMain.txtZone5 = Format(frmMain.txtZone5, "Fixed")
        
        frmMain.Label8.Visible = False
        frmMain.Label8.FontSize = 9
        frmMain.txtZone6.Visible = False
        frmMain.txtZone6 = Format(frmMain.txtZone6, "Fixed")
        frmMain.Label22.Visible = False
        
        frmMain.Label39.Visible = False
        frmMain.Label40.Visible = False
        frmMain.txtSeat11.Visible = False
        
    ElseIf Image14.Visible = True Then
    
        frmMain.txtSeats = "208B14"
    
        frmMain.Label2.Caption = "Seats 3/4/5"
        frmMain.Label2.FontSize = 9
        frmMain.txtZone1 = intPass3 + intPass4 + intPass5
        frmMain.txtZone1 = Format(frmMain.txtZone1, "Fixed")
        
        frmMain.Label4.Caption = "Seats 6/7/8"
        frmMain.Label4.FontSize = 9
        frmMain.txtZone2 = intPass6 + intPass7 + intPass8
        frmMain.txtZone2 = Format(frmMain.txtZone2, "Fixed")
        
        frmMain.Label5.Caption = "Seats 9/10/11"
        frmMain.Label5.FontSize = 9
        frmMain.txtZone3 = intPass9 + intPass10 + intPass11
        frmMain.txtZone3 = Format(frmMain.txtZone3, "Fixed")
        
        frmMain.Label6.Caption = "Seats 12/13/14"
        frmMain.Label6.FontSize = 9
        frmMain.txtZone4 = intPass12 + intPass13 + intPass14
        frmMain.txtZone4 = Format(frmMain.txtZone4, "Fixed")
        
        frmMain.Label7.Caption = "Baggage"
        frmMain.Label7.FontSize = 9
        frmMain.txtZone5 = 0
        frmMain.txtZone5 = Format(frmMain.txtZone5, "Fixed")
        
        frmMain.Label8.Visible = False
        frmMain.Label8.FontSize = 9
        frmMain.txtZone6.Visible = False
        frmMain.txtZone6 = Format(frmMain.txtZone6, "Fixed")
        frmMain.Label22.Visible = False
        
        frmMain.Label39.Visible = False
        frmMain.Label40.Visible = False
        frmMain.txtSeat11.Visible = False
    End If
    
    'If frmMain.Text1 <> "4" Then
    '    Call MsgBox("Nice try Pal!", vbOKOnly, "No Changes Allowed!")
    '    End
    'Else
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

Private Sub Image10_Click()

    Image1.Visible = False
    Image4.Visible = False
    Image5.Visible = False
    Image6.Visible = False
    Image7.Visible = False
    Image9.Visible = False
    Image10.Visible = False
    Image11.Visible = False
    Image12.Visible = False
    Image13.Visible = True
    Image14.Visible = False
    Label1.Visible = False
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
    chk123.Visible = False
    chk124.Visible = False
    chk125.Visible = False
    chk126.Visible = False
    chk127.Visible = False
    chk128.Visible = False
    chk129.Visible = False
    chk1210.Visible = False
    chk1211.Visible = False
    chk1212.Visible = False
    chk133.Visible = True
    chk134.Visible = True
    chk135.Visible = True
    chk136.Visible = True
    chk137.Visible = True
    chk138.Visible = True
    chk139.Visible = True
    chk1310.Visible = True
    chk1311.Visible = True
    chk1312.Visible = True
    chk1313.Visible = True
    chk143.Visible = False
    chk144.Visible = False
    chk145.Visible = False
    chk146.Visible = False
    chk147.Visible = False
    chk148.Visible = False
    chk149.Visible = False
    chk1410.Visible = False
    chk1411.Visible = False
    chk1412.Visible = False
    chk1413.Visible = False
    chk1414.Visible = False
    cmdDone.Top = 6720
    cmdClear.Top = 6720
    cmdExit.Top = 6720
    
End Sub

Private Sub Image11_Click()

    Image1.Visible = False
    Image4.Visible = False
    Image5.Visible = False
    Image6.Visible = False
    Image7.Visible = False
    Image9.Visible = False
    Image10.Visible = False
    Image11.Visible = False
    Image12.Visible = False
    Image13.Visible = False
    Image14.Visible = True
    Label1.Visible = False
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
    chk123.Visible = False
    chk124.Visible = False
    chk125.Visible = False
    chk126.Visible = False
    chk127.Visible = False
    chk128.Visible = False
    chk129.Visible = False
    chk1210.Visible = False
    chk1211.Visible = False
    chk1212.Visible = False
    chk133.Visible = False
    chk134.Visible = False
    chk135.Visible = False
    chk136.Visible = False
    chk137.Visible = False
    chk138.Visible = False
    chk139.Visible = False
    chk1310.Visible = False
    chk1311.Visible = False
    chk1312.Visible = False
    chk1313.Visible = False
    chk143.Visible = True
    chk144.Visible = True
    chk145.Visible = True
    chk146.Visible = True
    chk147.Visible = True
    chk148.Visible = True
    chk149.Visible = True
    chk1410.Visible = True
    chk1411.Visible = True
    chk1412.Visible = True
    chk1413.Visible = True
    chk1414.Visible = True
    cmdDone.Top = 6720
    cmdClear.Top = 6720
    cmdExit.Top = 6720
    
End Sub

Private Sub Image4_Click()

    Image1.Visible = True
    Image4.Visible = False
    Image5.Visible = False
    Image6.Visible = False
    Image7.Visible = False
    Image9.Visible = False
    Image10.Visible = False
    Image11.Visible = False
    Image12.Visible = False
    Image13.Visible = False
    Image14.Visible = False
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
    chk123.Visible = False
    chk124.Visible = False
    chk125.Visible = False
    chk126.Visible = False
    chk127.Visible = False
    chk128.Visible = False
    chk129.Visible = False
    chk1210.Visible = False
    chk1211.Visible = False
    chk1212.Visible = False
    chk133.Visible = False
    chk134.Visible = False
    chk135.Visible = False
    chk136.Visible = False
    chk137.Visible = False
    chk138.Visible = False
    chk139.Visible = False
    chk1310.Visible = False
    chk1311.Visible = False
    chk1312.Visible = False
    chk1313.Visible = False
    chk143.Visible = False
    chk144.Visible = False
    chk145.Visible = False
    chk146.Visible = False
    chk147.Visible = False
    chk148.Visible = False
    chk149.Visible = False
    chk1410.Visible = False
    chk1411.Visible = False
    chk1412.Visible = False
    chk1413.Visible = False
    chk1414.Visible = False
    cmdDone.Top = 6720
    cmdClear.Top = 6720
    cmdExit.Top = 6720
End Sub

Private Sub Image5_Click()

    Image2.Visible = True
    Image4.Visible = False
    Image5.Visible = False
    Image6.Visible = False
    Image7.Visible = False
    Image9.Visible = False
    Image10.Visible = False
    Image11.Visible = False
    Image12.Visible = False
    Image13.Visible = False
    Image14.Visible = False
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
    chk123.Visible = False
    chk124.Visible = False
    chk125.Visible = False
    chk126.Visible = False
    chk127.Visible = False
    chk128.Visible = False
    chk129.Visible = False
    chk1210.Visible = False
    chk1211.Visible = False
    chk1212.Visible = False
    chk133.Visible = False
    chk134.Visible = False
    chk135.Visible = False
    chk136.Visible = False
    chk137.Visible = False
    chk138.Visible = False
    chk139.Visible = False
    chk1310.Visible = False
    chk1311.Visible = False
    chk1312.Visible = False
    chk1313.Visible = False
    chk143.Visible = False
    chk144.Visible = False
    chk145.Visible = False
    chk146.Visible = False
    chk147.Visible = False
    chk148.Visible = False
    chk149.Visible = False
    chk1410.Visible = False
    chk1411.Visible = False
    chk1412.Visible = False
    chk1413.Visible = False
    chk1414.Visible = False
    cmdDone.Top = 6720
    cmdClear.Top = 6720
    cmdExit.Top = 6720
    
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
    frmMain.txtSeats = "208BCM"
    frmMain.txtFltNo.SetFocus
    Unload Me
    
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
    Image6.Visible = False
    Image7.Visible = False
    Image9.Visible = False
    Image10.Visible = False
    Image11.Visible = False
    Image12.Visible = False
    Image13.Visible = False
    Image14.Visible = False
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
    chk123.Visible = False
    chk124.Visible = False
    chk125.Visible = False
    chk126.Visible = False
    chk127.Visible = False
    chk128.Visible = False
    chk129.Visible = False
    chk1210.Visible = False
    chk1211.Visible = False
    chk1212.Visible = False
    chk133.Visible = False
    chk134.Visible = False
    chk135.Visible = False
    chk136.Visible = False
    chk137.Visible = False
    chk138.Visible = False
    chk139.Visible = False
    chk1310.Visible = False
    chk1311.Visible = False
    chk1312.Visible = False
    chk1313.Visible = False
    chk143.Visible = False
    chk144.Visible = False
    chk145.Visible = False
    chk146.Visible = False
    chk147.Visible = False
    chk148.Visible = False
    chk149.Visible = False
    chk1410.Visible = False
    chk1411.Visible = False
    chk1412.Visible = False
    chk1413.Visible = False
    chk1414.Visible = False
    cmdDone.Top = 6720
    cmdClear.Top = 6720
    cmdExit.Top = 6720
    
End Sub

Private Sub Image9_Click()

    Image1.Visible = False
    Image4.Visible = False
    Image5.Visible = False
    Image6.Visible = False
    Image7.Visible = False
    Image9.Visible = False
    Image10.Visible = False
    Image11.Visible = False
    Image12.Visible = True
    Image13.Visible = False
    Image14.Visible = False
    Label1.Visible = False
    Label2.Visible = True
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
    chk123.Visible = True
    chk124.Visible = True
    chk125.Visible = True
    chk126.Visible = True
    chk127.Visible = True
    chk128.Visible = True
    chk129.Visible = True
    chk1210.Visible = True
    chk1211.Visible = True
    chk1212.Visible = True
    chk133.Visible = False
    chk134.Visible = False
    chk135.Visible = False
    chk136.Visible = False
    chk137.Visible = False
    chk138.Visible = False
    chk139.Visible = False
    chk1310.Visible = False
    chk1311.Visible = False
    chk1312.Visible = False
    chk1313.Visible = False
    chk143.Visible = False
    chk144.Visible = False
    chk145.Visible = False
    chk146.Visible = False
    chk147.Visible = False
    chk148.Visible = False
    chk149.Visible = False
    chk1410.Visible = False
    chk1411.Visible = False
    chk1412.Visible = False
    chk1413.Visible = False
    chk1414.Visible = False
    cmdDone.Top = 6720
    cmdClear.Top = 6720
    cmdExit.Top = 6720
    
End Sub
