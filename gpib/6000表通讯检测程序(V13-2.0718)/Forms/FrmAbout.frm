VERSION 5.00
Begin VB.Form FrmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "关于"
   ClientHeight    =   4950
   ClientLeft      =   7125
   ClientTop       =   4005
   ClientWidth     =   5385
   Icon            =   "FrmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   5385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton Command1 
      Caption         =   "确 定"
      Default         =   -1  'True
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   4440
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   2655
      Left            =   0
      Picture         =   "FrmAbout.frx":0442
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5400
   End
   Begin VB.Label LabelTop 
      AutoSize        =   -1  'True
      Caption         =   "邮箱：BJ83391064@163.com"
      Height          =   180
      Index           =   4
      Left            =   840
      TabIndex        =   4
      Top             =   3720
      Width           =   2160
   End
   Begin VB.Label LabelTop 
      AutoSize        =   -1  'True
      Caption         =   "网址：www.BJHHFA.com"
      Height          =   180
      Index           =   3
      Left            =   840
      TabIndex        =   3
      Top             =   3360
      Width           =   1800
   End
   Begin VB.Label LabelTop 
      Caption         =   "本软件用于6000系列精密数字测量仪串口检测"
      Height          =   180
      Index           =   0
      Left            =   840
      TabIndex        =   2
      Top             =   3000
      Width           =   4035
   End
   Begin VB.Label LabelTop 
      AutoSize        =   -1  'True
      Caption         =   "电话：010-83391064   83391320"
      Height          =   180
      Index           =   2
      Left            =   840
      TabIndex        =   1
      Top             =   4080
      Width           =   2610
   End
End
Attribute VB_Name = "FrmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   Unload Me
End Sub

Private Sub Form_Activate()
SetFormToAlpha Me.hWnd, 220
End Sub

