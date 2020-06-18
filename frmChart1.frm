VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form frmChart1 
   Caption         =   "Form1"
   ClientHeight    =   6945
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9255
   LinkTopic       =   "Form1"
   ScaleHeight     =   6945
   ScaleWidth      =   9255
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBack 
      Caption         =   "&Back"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8040
      TabIndex        =   4
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8040
      TabIndex        =   3
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&View Grid"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8040
      TabIndex        =   1
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   7920
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   5775
      Left            =   600
      OleObjectBlob   =   "frmChart1.frx":0000
      TabIndex        =   2
      Top             =   480
      Width           =   7215
   End
End
Attribute VB_Name = "frmChart1"
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

Private Sub Command2_Click()
         ' Set chart type to 2d bar
         frmChart1.MSChart1.chartType = VtChChartType2dLine

         ' Use manual scale to display y axis (value axis)
         With frmChart1.MSChart1.Plot.Axis(VtChAxisIdY).ValueScale
            .Auto = False
            .Minimum = 4000
            .Maximum = 9000
         End With
'         frmChart1.MSChart1.Plot.Axis(VtChAxisIdY).AxisGrid = True
End Sub
Private Sub Command1_Click()

 With frmChart1.MSChart1
        '.ChartData = arrMPGandMiles
        '.Title = "Costs"
        .ColumnCount = 7
        .ColumnLabelCount = 7
        .Column = 1
        .ColumnLabel = "175"
        .Column = 2
        .ColumnLabel = "180"
        .Column = 3
        .ColumnLabel = "185"
        .Refresh
    End With
    
    Dim x(1 To 7, 1 To 2) As Variant
    Dim y(1 To 11, 1 To 2) As Variant
 
      x(1, 1) = "179"
      x(1, 2) = 4000
      x(2, 1) = "179"
      x(2, 2) = 4000
      x(3, 1) = "185"
      x(3, 2) = 5000
      x(4, 1) = "190"
      x(4, 2) = 5500
      x(5, 1) = "195"
      x(5, 2) = 6000
      x(6, 1) = "200"
      x(6, 2) = 6500
      x(7, 1) = "205"
      x(7, 2) = 7000
      
      y(1, 1) = "4000"
      y(2, 1) = "4500"
      y(3, 1) = "5000"
      y(4, 1) = "5500"
      y(5, 1) = "6000"
      y(6, 1) = "6500"
      y(7, 1) = "7000"
      y(8, 1) = "7500"
      y(9, 1) = "8000"
      y(10, 1) = "8500"
      y(11, 1) = "9000"
      
      MSChart1.Plot.UniformAxis = False
      MSChart1 = x

    ' Note that MSChart20Lib is for Visual Basic 6.0
    ' it should be MSChartLib in Visual Basic 5.0
    Dim currentaxis As MSChart20Lib.Axis
    Dim currentlabel As MSChart20Lib.Label
    ' Get a reference to the x axis
    Set currentaxis = MSChart1.Plot.Axis(VtChAxisIdX)
    ' Loop though and set the font of each label
    For Each currentlabel In currentaxis.Labels
        currentlabel.VtFont.Name = "Courier"
        currentlabel.VtFont.Size = 16
    Next currentlabel
    ' get a reference to the y axis
    Set currentaxis = MSChart1.Plot.Axis(VtChAxisIdY)
    ' loop through and set the font of each label
    For Each currentlabel In currentaxis.Labels
        currentlabel.VtFont.Name = "Courier"
        currentlabel.VtFont.Size = 16
    Next currentlabel
End Sub
         
         'frmChart.MSChart1.chartType = VtChChartType2dLine

          'Use manual scale to display y axis (value axis)
         'With frmChart.MSChart1.Plot.Axis(VtChAxisIdY).ValueScale
            '.Auto = False
           ' .Minimum = 4000
          '  .Maximum = 9000
         'End With
      
Private Sub Form_Load()
    
End Sub
