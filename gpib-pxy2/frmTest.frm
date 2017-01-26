VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmTest 
   Caption         =   "≤‚ ‘ƒ£øÈ"
   ClientHeight    =   6620
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6120
   LinkTopic       =   "Form1"
   ScaleHeight     =   6620
   ScaleWidth      =   6120
   StartUpPosition =   3  '¥∞ø⁄»± °
   Begin VB.OptionButton IO 
      Caption         =   "LAN"
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   360
      TabIndex        =   22
      Top             =   1080
      Width           =   1335
   End
   Begin MSWinsockLib.Winsock Winsock 
      Left            =   1920
      Top             =   960
      _ExtentX        =   494
      _ExtentY        =   494
      _Version        =   393216
   End
   Begin VB.ComboBox sendTxt 
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      ItemData        =   "frmTest.frx":0000
      Left            =   480
      List            =   "frmTest.frx":0019
      TabIndex        =   21
      Text            =   "*IDN?"
      Top             =   3720
      Width           =   5295
   End
   Begin VB.CommandButton cancel 
      Caption         =   "»°œ˚"
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      TabIndex        =   20
      Top             =   5760
      Width           =   1400
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   1920
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.CommandButton sendRead 
      Caption         =   "∑¢ÀÕ≤¢Ω” ‹"
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   18
      Top             =   5760
      Width           =   1400
   End
   Begin VB.CommandButton send 
      Caption         =   "∑¢   ÀÕ"
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   17
      Top             =   5760
      Width           =   1400
   End
   Begin VB.Frame Frame3 
      Caption         =   "LAN…Ë÷√"
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   2280
      TabIndex        =   7
      Top             =   120
      Width           =   3495
      Begin VB.TextBox IPport 
         BeginProperty Font 
            Name            =   "Œ¢»Ì—≈∫⁄"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   24
         Text            =   "5025"
         Top             =   840
         Width           =   2895
      End
      Begin VB.TextBox IPaddr 
         BeginProperty Font 
            Name            =   "Œ¢»Ì—≈∫⁄"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   23
         Text            =   "169.254.4.10"
         Top             =   360
         Width           =   2895
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "232…Ë÷√"
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   2280
      TabIndex        =   6
      Top             =   1680
      Width           =   3495
      Begin VB.ComboBox cr 
         BeginProperty Font 
            Name            =   "Œ¢»Ì—≈∫⁄"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         TabIndex        =   16
         Text            =   "NONE"
         Top             =   840
         Width           =   975
      End
      Begin VB.ComboBox stopbit 
         BeginProperty Font 
            Name            =   "Œ¢»Ì—≈∫⁄"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Text            =   "1"
         Top             =   840
         Width           =   975
      End
      Begin VB.ComboBox databit 
         BeginProperty Font 
            Name            =   "Œ¢»Ì—≈∫⁄"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   14
         Text            =   "8"
         Top             =   360
         Width           =   975
      End
      Begin VB.ComboBox comport 
         BeginProperty Font 
            Name            =   "Œ¢»Ì—≈∫⁄"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Text            =   "COM1"
         Top             =   360
         Width           =   975
      End
      Begin VB.ComboBox baudrate 
         BeginProperty Font 
            Name            =   "Œ¢»Ì—≈∫⁄"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         TabIndex        =   12
         Text            =   "9600"
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "GPIB…Ë÷√"
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   360
      TabIndex        =   5
      Top             =   1680
      Width           =   1695
      Begin VB.TextBox GPIBnum 
         BeginProperty Font 
            Name            =   "Œ¢»Ì—≈∫⁄"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   11
         Text            =   "0"
         Top             =   920
         Width           =   615
      End
      Begin VB.TextBox GPIBaddr 
         BeginProperty Font 
            Name            =   "Œ¢»Ì—≈∫⁄"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   8
         Text            =   "22"
         Top             =   430
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "–Ú∫≈£∫"
         BeginProperty Font 
            Name            =   "Œ¢»Ì—≈∫⁄"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "µÿ÷∑£∫"
         BeginProperty Font 
            Name            =   "Œ¢»Ì—≈∫⁄"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.TextBox recTxt 
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   480
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   4680
      Width           =   5295
   End
   Begin VB.OptionButton IO 
      Caption         =   "232"
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   360
      TabIndex        =   1
      Top             =   720
      Width           =   1335
   End
   Begin VB.OptionButton IO 
      Caption         =   "GPIB"
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label LabInfo 
      Caption         =   "–≈œ¢Ã· æ"
      BeginProperty Font 
         Name            =   "ÀŒÃÂ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   19
      Top             =   6360
      Width           =   6255
   End
   Begin VB.Label Label2 
      Caption         =   "Ω” ‹Œƒ±æ£∫"
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   4320
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "∑¢ÀÕŒƒ±æ£∫"
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   3240
      Width           =   1335
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private IOChoice As Integer
Dim status As Long
Dim defrm As Long
Dim vi As Long
Dim strRes As String * 200
Dim recflag As Boolean

Private Sub cancel_Click()
    Unload Me
End Sub

Private Sub IO_Click(Index As Integer)
    IOChoice = Index
End Sub

Private Sub closeDev()
On Error GoTo ErrorHandler
    Select Case dataDim.IOSetting
    Case 0
        status = viClose(vi)
        If (status < VI_SUCCESS) Then GoTo VisaErrorHandler
    Case 1
        Me.MSComm1.PortOpen = False
    Case 2
        Me.Winsock.Close
    End Select
    Exit Sub
ErrorHandler:
    Me.LabInfo.Caption = Error$
    Exit Sub
VisaErrorHandler:
    Dim strVisaErr As String * 200
    Call viStatusDesc(defrm, status, strVisaErr)
    Me.LabInfo.Caption = strVisaErr
    Exit Sub
End Sub

Private Sub openDev()
    On Error GoTo ErrorHandler
    Select Case IOChoice
    Case 0
        status = viOpenDefaultRM(defrm)
        If (status < VI_SUCCESS) Then GoTo VisaErrorHandler
        status = viOpen(defrm, "GPIB" + Trim(Me.GPIBnum.Text) + "::" + Trim(Me.GPIBaddr.Text) + "::INSTR", 0, 100, vi)
        If (status < VI_SUCCESS) Then GoTo VisaErrorHandler
    Case 1
        Me.MSComm1.CommPort = Right(Me.comport.Text, 1)
        Me.MSComm1.Settings = Me.baudrate.Text + "," + Left(Me.cr.Text, 1) + "," + Trim(str(Me.databit.Text)) + "," + Trim(str(Me.stopbit.Text))
        Me.MSComm1.InputMode = comInputModeText
        Me.MSComm1.RThreshold = 1
        Me.MSComm1.PortOpen = True
    Case 2
        If (Me.Winsock.State <> 7) Then
            Me.Winsock.Close
            Call Me.Winsock.connect(Trim(Me.IPaddr.Text), Val(Me.IPport.Text))
            Dim start As Long
            start = timeGetTime()
            While ((timeGetTime() - start) < 500)
                DoEvents
            Wend
            If (Me.Winsock.State = 7) Then
                Me.LabInfo = "connect success"
            Else
                Me.LabInfo = "connect fail"
            End If
        End If
    End Select
    Exit Sub
ErrorHandler:
    Me.LabInfo.Caption = Error$
    Exit Sub
VisaErrorHandler:
    Dim strVisaErr As String * 200
    Call viStatusDesc(defrm, status, strVisaErr)
    Me.LabInfo.Caption = strVisaErr
    Exit Sub
End Sub

Private Sub wrIO(content As String, wrIOFlag As Boolean)
On Error GoTo ErrorHandler
    Select Case IOChoice
    Case 0
        status = viVPrintf(vi, content + Chr$(10), 0)
        If (status < VI_SUCCESS) Then Exit Sub
        If (wrIOFlag) Then
            status = viVScanf(vi, "%t", strRes)
            Me.recTxt.Text = Now & vbcrfl & strRes
            If (status < VI_SUCCESS) Then GoTo VisaErrorHandler
        End If
    Case 1
        Me.MSComm1.OutBufferCount = 0
        Me.MSComm1.InBufferCount = 0
        Me.MSComm1.Output = content + vbCrLf
        If (wrIOFlag) Then
            start = timeGetTime()
            While ((Me.MSComm1.InBufferCount = 0) And (timeGetTime() - start < 2000))
                DoEvents
            Wend
            Sleep (100)
            strRes = Me.MSComm1.Input
            Me.recTxt.Text = Now & vbcrfl & strRes
            Me.MSComm1.InBufferCount = 0
        End If
    Case 2
        recflag = False
        Winsock.SendData content & vbCrLf
        If (wrIOFlag) Then
            start = timeGetTime()
            While ((Not recflag) And (timeGetTime() - start < 2000))
                DoEvents
            Wend
            recflag = False
            Me.recTxt.Text = Now & vbCrLf & strRes
        End If
    End Select
    Exit Sub
ErrorHandler:
    Me.LabInfo.Caption = Error$
    Exit Sub
VisaErrorHandler:
    Dim strVisaErr As String * 200
    Call viStatusDesc(defrm, status, strVisaErr)
    Me.LabInfo.Caption = strVisaErr
    Exit Sub
End Sub

Private Sub send_Click()
    Me.send.Enabled = False
    Me.recTxt.Text = ""
    Call openDev
    Call wrIO(Me.sendTxt.Text, False)
    Call closeDev
    Me.sendRead.Enabled = True
    Me.send.Enabled = True
End Sub

Private Sub sendRead_Click()
    Me.sendRead.Enabled = False
    Me.recTxt.Text = ""
    Call openDev
    Call wrIO(Me.sendTxt.Text, True)
    Call closeDev
    Me.sendRead.Enabled = True
End Sub

Private Sub Winsock_DataArrival(ByVal bytesTotal As Long)
    Winsock.GetData strRes
    recflag = True
End Sub

