VERSION 5.00
Begin VB.Form FrmMain 
   Caption         =   "Gpib通讯（82357A） 2015.8.31"
   ClientHeight    =   6432
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   10452
   LinkTopic       =   "Form1"
   ScaleHeight     =   6432
   ScaleWidth      =   10452
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton CmdRec 
      Caption         =   "接收"
      Height          =   732
      Left            =   6240
      TabIndex        =   13
      Top             =   4800
      Width           =   1212
   End
   Begin VB.TextBox TxtSend 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1212
      Left            =   6120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Text            =   "FrmMain.frx":0000
      Top             =   3120
      Width           =   4212
   End
   Begin VB.CommandButton CmdSend 
      Caption         =   "发送"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.4
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7680
      TabIndex        =   11
      Top             =   4560
      Width           =   1455
   End
   Begin VB.ComboBox ComboTim 
      Height          =   276
      ItemData        =   "FrmMain.frx":0011
      Left            =   3600
      List            =   "FrmMain.frx":003C
      TabIndex        =   9
      Text            =   "200"
      Top             =   4080
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   3720
      Top             =   3480
   End
   Begin VB.TextBox TxtLocal 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   60
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   1320
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
         Size            =   14.4
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4320
      TabIndex        =   7
      Top             =   4800
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
      Left            =   960
      TabIndex        =   5
      Text            =   "GPIB1::16::INSTR"
      Top             =   3240
      Width           =   3132
   End
   Begin VB.TextBox TxtIDE 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   60
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   1320
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
         Size            =   14.4
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2280
      TabIndex        =   2
      Top             =   4800
      Width           =   1455
   End
   Begin VB.CommandButton CmdDCV 
      Caption         =   "DCV"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.4
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   1
      Top             =   4800
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
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808000&
      DrawMode        =   9  'Not Mask Pen
      FillColor       =   &H00808000&
      FillStyle       =   0  'Solid
      Height          =   372
      Left            =   2520
      Shape           =   2  'Oval
      Top             =   4080
      Width           =   372
   End
   Begin VB.Label Label1 
      Caption         =   "频率"
      Height          =   252
      Index           =   1
      Left            =   3120
      TabIndex        =   10
      Top             =   4200
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "GPIB地址"
      Height          =   252
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   3360
      Width           =   852
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
Dim GPIBaddress As String
Dim status As Long
Dim defrm As Long
Dim vi As Long
Dim strRes As String * 200
Dim actual As Long
Dim writeStr As String
Dim readStr As String

Private Sub CmdConnect_Click()
    On Error GoTo ErrorHandler
    status = viOpenDefaultRM(defrm)
    If (status < VI_SUCCESS) Then GoTo VisaErrorHandler
    
    status = viOpen(defrm, Trim(Me.TxtGpibAddress.Text), 0, 0, vi)
    If (status < VI_SUCCESS) Then GoTo VisaErrorHandler
    
    writeStr = "*IDN?" & Chr$(10)
    status = viVPrintf(vi, writeStr, 0)
    If (status < VI_SUCCESS) Then GoTo VisaErrorHandler
    
    status = viVScanf(vi, "%t", strRes)
    If (status < VI_SUCCESS) Then GoTo VisaErrorHandler
    Me.LabInfo.Caption = "连接成功：The IDN String is: " & strRes
    
    Me.CmdConnect.BackColor = vbGreen
    Exit Sub
ErrorHandler:
    Me.CmdConnect.BackColor = &H8000000F
    Me.LabInfo.Caption = Error$
    Exit Sub
VisaErrorHandler:
    Me.CmdConnect.BackColor = &H8000000F
    Dim strVisaErr As String * 200
    Call viStatusDesc(defrm, status, strVisaErr)
    Me.LabInfo.Caption = strVisaErr
    Exit Sub
End Sub

Private Sub CmdDCI_Click()
    writeStr = ":MEASure:VOLTage:AC?" + Chr$(10)
    status = viVPrintf(vi, writeStr, 0)
    If (status < VI_SUCCESS) Then GoTo VisaErrorHandler
    Timer1.Enabled = True
    Exit Sub
VisaErrorHandler:
    Dim strVisaErr As String * 200
    Call viStatusDesc(defrm, status, strVisaErr)
    Me.LabInfo.Caption = strVisaErr
    Exit Sub
End Sub

Private Sub CmdDCV_Click()
    writeStr = ":MEASure:VOLTage:DC?" + Chr$(10)
    status = viVPrintf(vi, writeStr, 0)
    If (status < VI_SUCCESS) Then GoTo VisaErrorHandler
    Timer1.Enabled = True
    Exit Sub
VisaErrorHandler:
    Dim strVisaErr As String * 200
    Call viStatusDesc(defrm, status, strVisaErr)
    Me.LabInfo.Caption = strVisaErr
    Exit Sub
End Sub

Private Sub CmdRec_Click()
    If Me.Timer1 = False Then
        
    Else
    
    End If
End Sub

Private Sub CmdSend_Click()
    writeStr = Trim(Me.TxtSend.Text) + Chr$(10)
    status = viVPrintf(vi, writeStr, 0)
    If (status < VI_SUCCESS) Then GoTo VisaErrorHandler
    status = viVScanf(vi, "%t", strRes)
    If (status < VI_SUCCESS) Then GoTo VisaErrorHandler
    Me.LabInfo.Caption = strRes
    Exit Sub
VisaErrorHandler:
    Dim strVisaErr As String * 200
    Call viStatusDesc(defrm, status, strVisaErr)
    Me.LabInfo.Caption = strVisaErr
    Exit Sub
End Sub

Private Sub ComboTim_Click()
    Me.Timer1.Interval = Val(Me.ComboTim.Text)
End Sub

Private Sub CmdRES_Click()
    writeStr = ":MEASure:RESistance?" + Chr$(10)
    status = viVPrintf(vi, writeStr, 0)
    If (status < VI_SUCCESS) Then GoTo VisaErrorHandler
    Timer1.Enabled = True
    Exit Sub
VisaErrorHandler:
    Dim strVisaErr As String * 200
    Call viStatusDesc(defrm, status, strVisaErr)
    Me.LabInfo.Caption = strVisaErr
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call viClose(vi)
    Call viClose(defrm)
End Sub

Private Sub Timer1_Timer()
    If Me.Shape1.FillColor = &HFFFF00 Then
        Me.Shape1.FillColor = &H808000
    Else
        Me.Shape1.FillColor = &HFFFF00
    End If
    writeStr = ":READ?" + Chr$(10)
    status = viWrite(vi, writeStr, Len(writeStr), actual)
    If (status < VI_SUCCESS) Then GoTo VisaErrorHandler
    status = viVScanf(vi, "%t", strRes)
    If (status < VI_SUCCESS) Then GoTo VisaErrorHandler
    Me.TxtIDE.Text = strRes
    Exit Sub
VisaErrorHandler:
    Dim strVisaErr As String * 200
    Call viStatusDesc(defrm, status, strVisaErr)
    Me.LabInfo.Caption = strVisaErr
    Exit Sub
End Sub
