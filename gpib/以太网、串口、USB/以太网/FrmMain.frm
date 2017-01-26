VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form LAN 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "以太网通讯"
   ClientHeight    =   7245
   ClientLeft      =   4035
   ClientTop       =   4410
   ClientWidth     =   6630
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7245
   ScaleWidth      =   6630
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame3 
      Caption         =   "连接"
      Height          =   1815
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   6375
      Begin VB.CommandButton Cmdlink 
         Caption         =   "连接"
         Height          =   735
         Left            =   1440
         TabIndex        =   9
         Top             =   960
         Width           =   1572
      End
      Begin VB.TextBox txtip 
         Height          =   372
         Left            =   840
         TabIndex        =   8
         Text            =   "169.254.115.210"
         Top             =   360
         Width           =   2172
      End
      Begin VB.TextBox txtport 
         Height          =   372
         Left            =   4200
         TabIndex        =   7
         Text            =   "3490"
         Top             =   360
         Width           =   1452
      End
      Begin VB.Shape islink 
         BorderColor     =   &H80000016&
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   240
         Shape           =   3  'Circle
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "IP地址："
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "端口号："
         Height          =   375
         Left            =   3360
         TabIndex        =   10
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "发送区"
      Height          =   1095
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   6375
      Begin VB.ComboBox combosend 
         Height          =   300
         ItemData        =   "FrmMain.frx":08CA
         Left            =   240
         List            =   "FrmMain.frx":08E0
         TabIndex        =   5
         Top             =   480
         Width           =   6015
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "接收区"
      Height          =   3015
      Left            =   120
      TabIndex        =   2
      Top             =   3240
      Width           =   6375
      Begin VB.TextBox txtreceive 
         Height          =   2535
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   360
         Width           =   6135
      End
   End
   Begin VB.CommandButton cmdclc 
      Caption         =   "清空"
      Height          =   735
      Left            =   4440
      TabIndex        =   1
      Top             =   6360
      Width           =   1695
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   360
      Top             =   0
   End
   Begin VB.CommandButton cmdsend 
      Caption         =   "发送"
      Height          =   735
      Left            =   360
      TabIndex        =   0
      Top             =   6360
      Width           =   1695
   End
   Begin MSWinsockLib.Winsock Winsock 
      Left            =   240
      Top             =   3960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   3490
   End
End
Attribute VB_Name = "LAN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' 安捷伦地址、端口号可见于机身上

Option Explicit
Dim LinkCount As Integer

Private Sub cmdclc_Click()               '清除接受信息
    txtreceive = ""
End Sub

Private Sub cmdlink_Click()
    If Winsock.State <> 7 Then          '判断是否连接上，若已连接则不执行
        LinkCount = LinkCount + 1       '点两次连接，若还未连接，则判断为线缆问题
        If LinkCount > 1 Then
            MsgBox "请检查线缆!", vbExclamation, "Error"
            LinkCount = 0
        End If
        Winsock.Close
        Winsock.Connect Trim(txtip.Text), Val(txtport.Text)
        Timer1.Enabled = True
    End If
End Sub

Private Sub cmdsend_Click()
    If Winsock.State = 7 Then
        Winsock.SendData combosend.Text & vbCrLf & vbCrLf       '通过LAN发送加两个CrLf才能被安捷伦识别
    Else
        MsgBox "请先连接!", vbExclamation, "Error"
    End If
End Sub

Private Sub Form_Load()
    Timer1.Interval = 500
    Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()                 '定时检测连接状态
    If Winsock.State = 7 Then
        islink.FillColor = vbGreen
        LinkCount = 0
    Else
        islink.FillColor = vbRed
    End If
End Sub

Private Sub txtreceive_Change()
txtreceive.SelStart = Len(txtreceive)
End Sub

Private Sub Winsock_DataArrival(ByVal bytesTotal As Long)       '有数据输入
    Dim strdata As String
    Winsock.GetData strdata
    txtreceive = txtreceive & strdata & vbCrLf
End Sub

