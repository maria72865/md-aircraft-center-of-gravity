VERSION 5.00
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form frmTS 
   Caption         =   "Contact Tech Support"
   ClientHeight    =   6840
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9930
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
   ScaleHeight     =   6840
   ScaleWidth      =   9930
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5400
      TabIndex        =   13
      Top             =   5760
      Width           =   1095
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3360
      TabIndex        =   12
      Top             =   5760
      Width           =   1095
   End
   Begin VB.TextBox txtMsg 
      Height          =   1335
      Left            =   240
      TabIndex        =   11
      Top             =   3960
      Width           =   9495
   End
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin MSMAPI.MAPISession MAPISession1 
      Left            =   840
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin VB.Label Label6 
      Caption         =   "Please type your question below: "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   3600
      Width           =   9495
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
      Left            =   3120
      TabIndex        =   7
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
      Left            =   2880
      TabIndex        =   9
      Top             =   960
      Width           =   1695
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
      Left            =   3120
      TabIndex        =   8
      Top             =   1440
      Width           =   1335
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
      Left            =   2880
      TabIndex        =   6
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
      Left            =   2520
      TabIndex        =   5
      Top             =   120
      Width           =   4935
   End
   Begin VB.Image Image1 
      Height          =   1095
      Left            =   5040
      Picture         =   "frmTS.frx":0000
      Stretch         =   -1  'True
      Top             =   600
      Width           =   1815
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
      TabIndex        =   4
      Top             =   2520
      Width           =   9615
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
      TabIndex        =   3
      Top             =   1920
      Width           =   9615
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
      TabIndex        =   2
      Top             =   2880
      Width           =   9615
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
      TabIndex        =   1
      Top             =   2160
      Width           =   9615
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
      TabIndex        =   0
      Top             =   3120
      Width           =   9615
   End
End
Attribute VB_Name = "frmTS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strMsg As String

Private Sub cmdCancel_Click()
    frmMain.Show
    Unload Me
End Sub

Sub cmdSend_Click()

    strMsg = txtMsg

      'MAPI constants:
      Const SESSION_SIGNON = 1
      Const MESSAGE_COMPOSE = 6
      Const ATTACHTYPE_DATA = 0
      Const RECIPTYPE_TO = 1
      Const RECIPTYPE_CC = 2
      Const MESSAGE_RESOLVENAME = 13
      Const MESSAGE_SEND = 3
      Const SESSION_SIGNOFF = 2
    
      'Open up a MAPI session:
      MAPISession1.Action = SESSION_SIGNON
      'Point the MAPI messages control to the open MAPI session:
      MAPIMessages1.SessionID = frmTS.MAPISession1.SessionID
    
      MAPIMessages1.Action = MESSAGE_COMPOSE   'Start a new message
    
      'Set the subject of the message:
      MAPIMessages1.MsgSubject = "Tech Support"
      
      'Set the message content:
      MAPIMessages1.MsgNoteText = strMsg
    
      'Set the recipients
      'MAPIMessages1.RecipIndex = 0                    'First recipient
      MAPIMessages1.RecipType = RECIPTYPE_TO          'Recipient in TO line
      MAPIMessages1.RecipDisplayName = "techsupport@pilot-international.com"   'e-mail name
    
      
      'MESSAGE_RESOLVENAME checks to ensure the recipient is valid and puts
      'the recipient address in MapiMessages1.RecipAddress
      'If the E-Mail name is not valid, a trappable error will occur.
      'MAPIMessages1.Action = MESSAGE_RESOLVENAME
      'Send the message:
      MAPIMessages1.Action = MESSAGE_SEND
    
      'Close MAPI mail session:
      MAPISession1.Action = SESSION_SIGNOFF
      
    Call MsgBox("Your message has been sent. You will be contacted within the next two business days.", vbOKOnly, "Message Sent")
    frmMain.Show
    Unload Me
        

End Sub
