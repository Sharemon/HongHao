VERSION 5.00
Begin VB.Form FrmSupport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "技术支持"
   ClientHeight    =   3930
   ClientLeft      =   7875
   ClientTop       =   4560
   ClientWidth     =   3855
   Icon            =   "FrmSupport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   3855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton Command1 
      Caption         =   "确 定"
      Default         =   -1  'True
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Label LabelTop 
      AutoSize        =   -1  'True
      Caption         =   "北京弘豪福安仪器有限公司"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   216
      Index           =   0
      Left            =   720
      TabIndex        =   6
      Top             =   240
      Width           =   2592
   End
   Begin VB.Label LabelTop 
      AutoSize        =   -1  'True
      Caption         =   "电话：010－83391064,13801004676"
      Height          =   180
      Index           =   3
      Left            =   480
      TabIndex        =   5
      Top             =   1440
      Width           =   2790
   End
   Begin VB.Label LabelTop 
      AutoSize        =   -1  'True
      Caption         =   "传真：010－84279078,84278704"
      Height          =   180
      Index           =   4
      Left            =   480
      TabIndex        =   4
      Top             =   1800
      Width           =   2520
   End
   Begin VB.Label LabelTop 
      AutoSize        =   -1  'True
      Caption         =   "地址：北京市丰台区云冈"
      Height          =   180
      Index           =   1
      Left            =   480
      TabIndex        =   3
      Top             =   720
      Width           =   1980
   End
   Begin VB.Label LabelTop 
      AutoSize        =   -1  'True
      Caption         =   "邮编：100074"
      Height          =   180
      Index           =   2
      Left            =   480
      TabIndex        =   2
      Top             =   2160
      Width           =   1116
   End
   Begin VB.Label LabelTop 
      Caption         =   "E-mail： 83391064@163.com"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   5
      Left            =   480
      TabIndex        =   1
      Top             =   2520
      Width           =   2892
   End
   Begin VB.Label LabelTop 
      AutoSize        =   -1  'True
      Caption         =   "网址：www.bjhhfa.com"
      Height          =   180
      Index           =   6
      Left            =   480
      TabIndex        =   7
      Top             =   1080
      Width           =   1884
   End
End
Attribute VB_Name = "FrmSupport"
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

