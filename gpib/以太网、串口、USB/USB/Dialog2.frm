VERSION 5.00
Object = "{1C98F15C-068A-11D4-98C2-00108301CB39}#2.0#0"; "agt3494A.ocx"
Begin VB.Form USB 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "USB连接器"
   ClientHeight    =   4044
   ClientLeft      =   4032
   ClientTop       =   4416
   ClientWidth     =   6228
   Icon            =   "Dialog2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4044
   ScaleWidth      =   6228
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdclc 
      Caption         =   "清空"
      Height          =   372
      Left            =   5040
      TabIndex        =   9
      Top             =   3480
      Width           =   972
   End
   Begin Agt3494ALib.Agt3494A Agt3494A1 
      Left            =   360
      Top             =   2040
      _ExtentX        =   762
      _ExtentY        =   677
   End
   Begin VB.CommandButton cmdreceive 
      Caption         =   "接收"
      Height          =   372
      Left            =   3720
      TabIndex        =   8
      Top             =   3480
      Width           =   1092
   End
   Begin VB.CommandButton cmdsend 
      Caption         =   "发送"
      Height          =   372
      Left            =   2400
      TabIndex        =   7
      Top             =   3480
      Width           =   1092
   End
   Begin VB.CommandButton cmdlink 
      Caption         =   "连接"
      Height          =   372
      Left            =   1080
      TabIndex        =   6
      Top             =   3480
      Width           =   1092
   End
   Begin VB.TextBox txtreceive 
      Height          =   1812
      Left            =   1080
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   1440
      Width           =   4932
   End
   Begin VB.ComboBox Combosend 
      Height          =   288
      ItemData        =   "Dialog2.frx":16C02
      Left            =   1080
      List            =   "Dialog2.frx":16C18
      TabIndex        =   3
      Top             =   960
      Width           =   4932
   End
   Begin VB.TextBox txtaddr 
      Height          =   372
      Left            =   1080
      TabIndex        =   1
      Text            =   "USB0::2391::1543::my47020574::0::INSTR"
      Top             =   360
      Width           =   4932
   End
   Begin VB.Shape islink 
      BorderColor     =   &H80000016&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   372
      Left            =   360
      Shape           =   3  'Circle
      Top             =   3480
      Width           =   372
   End
   Begin VB.Label Label3 
      Caption         =   "接收："
      Height          =   372
      Left            =   360
      TabIndex        =   4
      Top             =   1440
      Width           =   492
   End
   Begin VB.Label Label2 
      Caption         =   "发送："
      Height          =   252
      Left            =   360
      TabIndex        =   2
      Top             =   960
      Width           =   612
   End
   Begin VB.Label Label1 
      Caption         =   "设备地址："
      Height          =   252
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   732
   End
End
Attribute VB_Name = "USB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'使用Agt3494A1控件，需安装安捷伦IO Library，USB地址可用Agilent Connection Expert获取

Option Explicit
Dim Receivable As Boolean

Private Sub cmdclc_Click()
txtreceive = ""
End Sub

Private Sub cmdlink_Click()
On Error GoTo HadErr
    Agt3494A1.Disconnect
    Agt3494A1.Address = txtaddr
    Agt3494A1.Connect
    If Agt3494A1.IsConnected Then islink.FillColor = vbGreen
    If Not Agt3494A1.IsConnected Then islink.FillColor = vbRed
HadErr: Exit Sub
End Sub

Private Sub cmdreceive_Click()
    If Agt3494A1.IsConnected Then
        If Receivable Then
            If Right(Combosend.Text, 1) = "?" Then      '判断发送的是否为疑问句，若不是则不调用接受函数
                Dim strdata As String
                Agt3494A1.Enter strdata
                txtreceive = txtreceive & strdata & vbCrLf
            Else
                txtreceive = txtreceive & "NULL" & vbCrLf
            End If
            Receivable = False
        End If
    Else
        MsgBox "请先连接！", vbExclamation, "Error"
    End If
End Sub

Private Sub cmdsend_Click()
Combosend.Text = Trim(Combosend.Text)
    If Agt3494A1.IsConnected Then
        If Combosend.Text <> "" Then
            Agt3494A1.Output Combosend.Text
            Receivable = True
        End If
    Else
        MsgBox "请先连接！", vbExclamation, "Error"
    End If
End Sub

Private Sub Form_Load()
    Receivable = False
End Sub

Private Sub txtreceive_Change()
txtreceive.SelStart = Len(txtreceive)
End Sub
