VERSION 5.00
Begin VB.Form FrmMain 
   Caption         =   "Gpib通讯（Agilent USB）"
   ClientHeight    =   6435
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10455
   LinkTopic       =   "Form1"
   ScaleHeight     =   6435
   ScaleWidth      =   10455
   StartUpPosition =   3  '窗口缺省
   Begin VB.ComboBox ComboTim 
      Height          =   300
      ItemData        =   "FrmMain.frx":0000
      Left            =   5040
      List            =   "FrmMain.frx":002B
      TabIndex        =   9
      Text            =   "200"
      Top             =   3720
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   3720
      Top             =   3360
   End
   Begin VB.TextBox TxtLocal 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   66
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   1440
      Left            =   0
      TabIndex        =   8
      Text            =   "98765432.1"
      Top             =   1560
      Width           =   10455
   End
   Begin VB.CommandButton CmdRES 
      Caption         =   "Ω"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4320
      TabIndex        =   7
      Top             =   4680
      Width           =   1455
   End
   Begin VB.TextBox TxtGpibAddress 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   5
      Text            =   "16"
      Top             =   3720
      Width           =   975
   End
   Begin VB.TextBox TxtIDE 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   66
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   1440
      Left            =   0
      TabIndex        =   3
      Text            =   "98765432.1"
      Top             =   0
      Width           =   10455
   End
   Begin VB.CommandButton CmdDCI 
      Caption         =   "DCI"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2280
      TabIndex        =   2
      Top             =   4680
      Width           =   1455
   End
   Begin VB.CommandButton CmdDCV 
      Caption         =   "DCV"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   1
      Top             =   4680
      Width           =   1455
   End
   Begin VB.CommandButton CmdConnect 
      Caption         =   "连接"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "频率"
      Height          =   255
      Index           =   1
      Left            =   4560
      TabIndex        =   10
      Top             =   3840
      Width           =   855
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      DrawMode        =   9  'Not Mask Pen
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   3720
      Shape           =   2  'Oval
      Top             =   3720
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "GPIB地址"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   6
      Top             =   3840
      Width           =   855
   End
   Begin VB.Label LabInfo 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      TabIndex        =   4
      Top             =   5880
      Width           =   10455
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rm As VisaComLib.ResourceManager
 Dim session As VisaComLib.IMessage
 Dim status As Long
 Dim idn As String
 Dim ReadString As String

Private Sub CreateResource()
    On Error GoTo errorHandler
    Set rm = New VisaComLib.ResourceManager
    Set session = rm.Open("GPIB0::" & Trim(Me.TxtGpibAddress.Text))
    session.WriteString "*IDN?" & vbLf
    idn = session.ReadString(1000)
    Me.LabInfo.Caption = "连接成功：The IDN String is: " & idn
    Me.CmdConnect.BackColor = vbGreen
    Exit Sub
errorHandler:
    Me.LabInfo.Caption = "连接失败：" & Err.Description
    Me.CmdConnect.BackColor = &H8000000F
End Sub

Private Sub CmdConnect_Click()
    CreateResource
End Sub

Private Sub CmdDCI_Click()
    On Error Resume Next
    session.WriteString ":MEASure:VOLTage:AC?" & vbLf
    ReadString = session.ReadString(1000)
'    If Trim(ReadString) <> "" Then
        Me.LabInfo.Caption = ReadString
        Timer1.Enabled = True
'    End If
End Sub

Private Sub CmdDCV_Click()
    On Error Resume Next
    session.WriteString ":MEASure:VOLTage:DC?" & vbLf
    ReadString = session.ReadString(1000)
    If Trim(ReadString) <> "" Then
        Me.LabInfo.Caption = ReadString
        Timer1.Enabled = True
    End If
End Sub

Private Sub ComboTim_Click()
    Me.Timer1.Interval = Val(Me.ComboTim.Text)
End Sub

Private Sub CmdRES_Click()
    On Error Resume Next
    session.WriteString ":MEASure:RESistance?" & vbLf
    ReadString = session.ReadString(1000)
    If Trim(ReadString) <> "" Then
        Me.LabInfo.Caption = ReadString
        Timer1.Enabled = True
    End If
End Sub

Private Sub Timer1_Timer()
    On Error Resume Next
    If Me.Shape1.FillColor = vbGreen Then
        Me.Shape1.FillColor = &H8000&
    Else
        Me.Shape1.FillColor = vbGreen
    End If
    session.WriteString ":READ?" & vbLf
    ReadString = session.ReadString(1000)
    Me.TxtIDE.SelStart = 0
    Me.TxtIDE.SelLength = Len(Me.TxtIDE.Text)
    Me.TxtIDE.SelText = ReadString
End Sub
