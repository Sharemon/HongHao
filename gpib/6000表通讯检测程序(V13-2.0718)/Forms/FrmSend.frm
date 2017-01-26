VERSION 5.00
Begin VB.Form FrmSend 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "发送指令(用于指令调试)"
   ClientHeight    =   3300
   ClientLeft      =   7875
   ClientTop       =   4935
   ClientWidth     =   3855
   Icon            =   "FrmSend.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   3855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton CmdExit 
      Caption         =   "退出"
      Height          =   372
      Left            =   2640
      TabIndex        =   9
      Top             =   2760
      Width           =   1092
   End
   Begin VB.Timer TmrAutoSend 
      Left            =   0
      Top             =   0
   End
   Begin VB.CheckBox ChkSent 
      Caption         =   "自动发送"
      Height          =   252
      Left            =   360
      TabIndex        =   6
      Top             =   1800
      Width           =   3252
   End
   Begin VB.TextBox TxtTime 
      Height          =   264
      Left            =   1680
      TabIndex        =   5
      Text            =   "1000"
      Top             =   2280
      Width           =   660
   End
   Begin VB.CommandButton CmdEmpty 
      Caption         =   "清空重填"
      Height          =   372
      Left            =   120
      TabIndex        =   4
      Top             =   2760
      Width           =   1092
   End
   Begin VB.CommandButton CmdManual 
      Caption         =   "手动发送"
      Height          =   372
      Left            =   1320
      TabIndex        =   3
      Top             =   2760
      Width           =   1212
   End
   Begin VB.OptionButton OptSendHex 
      Caption         =   "十六进制指令"
      Height          =   180
      Left            =   360
      TabIndex        =   2
      Top             =   1440
      Width           =   3252
   End
   Begin VB.OptionButton OptSendASC 
      Caption         =   "ASCII指令"
      Height          =   180
      Left            =   360
      TabIndex        =   1
      Top             =   1080
      Value           =   -1  'True
      Width           =   3252
   End
   Begin VB.TextBox TxtSend 
      Height          =   612
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   3372
   End
   Begin VB.Label LabelOnTop 
      AutoSize        =   -1  'True
      Caption         =   "自动发送周期"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   204
      Index           =   20
      Left            =   360
      TabIndex        =   8
      Top             =   2280
      Width           =   1152
   End
   Begin VB.Label LabelOnTop 
      AutoSize        =   -1  'True
      Caption         =   "毫秒/次"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   204
      Index           =   21
      Left            =   2400
      TabIndex        =   7
      Top             =   2280
      Width           =   624
   End
End
Attribute VB_Name = "FrmSend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdEmpty_Click()
If ChkSent.Value = 1 Then
   ChkSent.Value = 0
End If
TxtSend.Text = ""
End Sub

Private Sub CmdExit_Click()
FrmMain.FrmSendShow = False
Unload Me
Set FrmSend = Nothing
End Sub

Private Sub CmdManual_Click()
On Error GoTo ErrHndl
If FrmMain.MSComPort.PortOpen = True Then
   If TxtSend.Text = "" Then
      MsgBox "请输入要发送的语句或命令", 16, "2000标准负荷测量仪"
   Else
      If OptSendHex.Value = True Then
         FrmMain.MSComPort.InputMode = comInputModeBinary
         Call hexSend
      Else
         FrmMain.MSComPort.Output = Trim(TxtSend.Text)
      End If
   End If
   Automatic = False
End If
ErrHndl:
    Debug.Print "Error No. " & Err.Number & ": Error Description: " & Err.Description
End Sub

Private Sub ChkSent_Click()
On Error GoTo ErrHndl
ErrHndl:

If ChkSent.Value = 1 Then
        If FrmMain.MSComPort.PortOpen = True Then
            TmrAutoSend.Interval = Val(TxtTime.Text)
            TmrAutoSend.Enabled = True
        Else
            ChkSent.Value = 0
            MsgBox "串口还没有打开，请先打开串口", 48, "2000标准负荷测量仪"
        End If
ElseIf ChkSent.Value = 0 Then
        TmrAutoSend.Enabled = False
End If

Debug.Print "Error No. " & Err.Number & ": Error Description: " & Err.Description
End Sub

Private Sub hexSend()

On Error Resume Next
    Dim outputLen As Integer
    Dim outData As String
    Dim SendArr() As Byte
    Dim TemporarySave As String
    Dim dataCount As Integer
    Dim i As Integer
    
    outData = UCase(Replace(TxtSend.Text, Space(1), Space(0)))
    outData = UCase(outData)
    outputLen = Len(outData)
    
    For i = 0 To outputLen
        TemporarySave = Mid(outData, i + 1, 1)
        If (Asc(TemporarySave) >= 48 And Asc(TemporarySave) <= 57) Or (Asc(TemporarySave) >= 65 And Asc(TemporarySave) <= 70) Then
            dataCount = dataCount + 1
        Else
            Exit For
            Exit Sub
        End If
    Next
    
    If dataCount Mod 2 <> 0 Then
        dataCount = dataCount - 1
    End If
    
    outData = Left(outData, dataCount)
    
    ReDim SendArr(dataCount / 2 - 1)
    For i = 0 To dataCount / 2 - 1
        SendArr(i) = Val("&H" + Mid(outData, i * 2 + 1, 2))
    Next
    
    sendcount = sendcount + (dataCount / 2)
         
    FrmMain.MSComPort.Output = SendArr
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 FrmMain.LabInfo.Caption = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
Set FrmSend = Nothing
End Sub

Private Sub TmrAutoSend_Timer()

On Error GoTo Err
    If TxtSend.Text = "" Then
        ChkSent.Value = 0
        MsgBox "发送数据不能为空,请输入要发送的语句或命令", 16, "2000标准负荷测量仪"
    Else
        If ChkSent.Value = 1 Then
            If OptSendHex.Value = True Then
                FrmMain.MSComPort.InputMode = comInputModeBinary
                Call hexSend
            ElseIf OptSendASC.Value = True Then
                FrmMain.MSComPort.Output = Trim(TxtSend.Text)
                OutputSignal = TxtSend.Text
                sendcount = sendcount + LenB(StrConv(OutputSignal, vbFromUnicode))
            End If
        End If
    End If
Err:
End Sub


Private Sub CmdEmpty_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Lan = 0 Then
   FrmMain.LabInfo.Caption = "清空发送区，重新填写"
ElseIf Lan = 1 Then
   FrmMain.LabInfo.Caption = "Send area to fill empty"
End If
End Sub

Private Sub CmdManual_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Lan = 0 Then
   FrmMain.LabInfo.Caption = "手动发送发送区中的指令"
ElseIf Lan = 1 Then
   FrmMain.LabInfo.Caption = "Manually send commands in the send area"
End If
End Sub

Private Sub ChkSent_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Lan = 0 Then
   FrmMain.LabInfo.Caption = "自动发送发送区中的指令"
ElseIf Lan = 1 Then
   FrmMain.LabInfo.Caption = "Automatically send commands in the send area"
End If
End Sub

