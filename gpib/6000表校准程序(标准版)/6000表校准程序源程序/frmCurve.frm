VERSION 5.00
Begin VB.Form frmCurve 
   Caption         =   "Form1"
   ClientHeight    =   3615
   ClientLeft      =   900
   ClientTop       =   6195
   ClientWidth     =   4215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   4215
   ShowInTaskbar   =   0   'False
   Begin DDM6000_Calitor.Curve Curve1 
      Height          =   3615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   6376
      GridRange       =   "20,20"
      GridColor       =   8421376
      GridVisible     =   -1  'True
      BorderColor     =   255
      BorderSize      =   0
      CurveColor      =   65280
      BackColor       =   0
      MaxValue        =   0
      MinValue        =   0
      CurvCount       =   100
      ForeColor       =   15066597
      ShowMidLine     =   -1  'True
      MidLineColor    =   65280
      ShowScale       =   -1  'True
      MidValue        =   0
      ValueRange      =   100
      AutoRange       =   -1  'True
      CurveStyle      =   0
   End
End
Attribute VB_Name = "frmCurve"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Caption = IIf(Lang, "Real-time curve", "数据实时曲线")
End Sub
