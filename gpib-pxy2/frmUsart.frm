VERSION 5.00
Begin VB.Form frmUsart 
   Caption         =   "串口设置"
   ClientHeight    =   4935
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7380
   LinkTopic       =   "Form1"
   ScaleHeight     =   4935
   ScaleWidth      =   7380
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton cancel 
      Caption         =   "取 消"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   5280
      TabIndex        =   24
      Top             =   4200
      Width           =   1400
   End
   Begin VB.CheckBox outEn 
      Caption         =   "输出使能"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5640
      TabIndex        =   16
      Top             =   240
      Width           =   1575
   End
   Begin VB.ComboBox crOut 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   5160
      TabIndex        =   15
      Text            =   "NONE"
      Top             =   3360
      Width           =   1695
   End
   Begin VB.ComboBox stopBitOut 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   5160
      TabIndex        =   14
      Text            =   "1"
      Top             =   2760
      Width           =   1695
   End
   Begin VB.ComboBox dataBitOut 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   5160
      TabIndex        =   13
      Text            =   "8"
      Top             =   2160
      Width           =   1695
   End
   Begin VB.ComboBox baudRateOut 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      ItemData        =   "frmUsart.frx":0000
      Left            =   5160
      List            =   "frmUsart.frx":0007
      TabIndex        =   12
      Text            =   "9600"
      Top             =   1560
      Width           =   1695
   End
   Begin VB.ComboBox comPortOut 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      ItemData        =   "frmUsart.frx":0011
      Left            =   5160
      List            =   "frmUsart.frx":0013
      TabIndex        =   11
      Top             =   960
      Width           =   1695
   End
   Begin VB.ComboBox comPort 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      ItemData        =   "frmUsart.frx":0015
      Left            =   1440
      List            =   "frmUsart.frx":0017
      TabIndex        =   9
      Top             =   960
      Width           =   1695
   End
   Begin VB.CommandButton ok 
      Caption         =   "确 定"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   3240
      TabIndex        =   8
      Top             =   4200
      Width           =   1400
   End
   Begin VB.ComboBox cr 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1440
      TabIndex        =   7
      Text            =   "NONE"
      Top             =   3360
      Width           =   1695
   End
   Begin VB.ComboBox stopBit 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1440
      TabIndex        =   6
      Text            =   "1"
      Top             =   2760
      Width           =   1695
   End
   Begin VB.ComboBox dataBit 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1440
      TabIndex        =   5
      Text            =   "8"
      Top             =   2160
      Width           =   1695
   End
   Begin VB.ComboBox baudRate 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      ItemData        =   "frmUsart.frx":0019
      Left            =   1440
      List            =   "frmUsart.frx":001B
      TabIndex        =   4
      Text            =   "9600"
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Line Line1 
      X1              =   3600
      X2              =   3600
      Y1              =   120
      Y2              =   3720
   End
   Begin VB.Label Label12 
      Caption         =   "输入串口："
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   23
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label11 
      Caption         =   "输出串口："
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   22
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "波特率："
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4080
      TabIndex        =   21
      Top             =   1560
      Width           =   840
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "数据位："
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4080
      TabIndex        =   20
      Top             =   2160
      Width           =   840
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "停止位："
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4080
      TabIndex        =   19
      Top             =   2760
      Width           =   840
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "校验："
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4080
      TabIndex        =   18
      Top             =   3360
      Width           =   630
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "端口："
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4080
      TabIndex        =   17
      Top             =   960
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "端口："
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   360
      TabIndex        =   10
      Top             =   960
      Width           =   630
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "校验："
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   360
      TabIndex        =   3
      Top             =   3360
      Width           =   630
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "停止位："
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   360
      TabIndex        =   2
      Top             =   2760
      Width           =   840
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "数据位："
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   360
      TabIndex        =   1
      Top             =   2160
      Width           =   840
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "波特率："
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   360
      TabIndex        =   0
      Top             =   1560
      Width           =   840
   End
End
Attribute VB_Name = "frmUsart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim i As Integer
    If frmMain.sp.PortOpen = True Then                  '先关闭串口
        MsgBox ("请先断开连接")
    Else
        For i = 1 To 16 Step 1
            frmMain.sp.CommPort = i
            On Error Resume Next                            '说明当一个运行时错误发生时，控件转到紧接着发生错误的语句之后的语句，并在此继续运行。访问对象时要使用这种形式而不使用 On Error GoTo。
            frmMain.sp.PortOpen = True
            If Err.Number <> 8002 Then                      '无效的串口号。这样可以检测到虚拟串口，如果用Err.Number = 0的话检测不到虚拟串口
                Me.comPort.AddItem "COM" & i                   '生成串口选择列表
                Me.comPortOut.AddItem "COM" & i
            End If
            frmMain.sp.PortOpen = False
        Next i
    End If
    Me.comPort.Text = "COM" + Trim(str(dataDim.comPort))
    Me.baudRate.AddItem ("1200")
    Me.baudRate.AddItem ("2400")
    Me.baudRate.AddItem ("4800")
    Me.baudRate.AddItem ("9600")
    Me.baudRate.AddItem ("14400")
    Me.baudRate.AddItem ("19200")
    Me.baudRate.AddItem ("38400")
    Me.baudRate.AddItem ("56000")
    Me.baudRate.AddItem ("57600")
    Me.baudRate.AddItem ("115200")
    Me.baudRate.AddItem ("194000")
    For i = 5 To 8
        Me.dataBit.AddItem (Trim(str(i)))
    Next
    Me.stopBit.AddItem ("1")
    Me.stopBit.AddItem ("2")
    Me.cr.AddItem ("NONE")
    Me.cr.AddItem ("ODD")
    Me.cr.AddItem ("EVEN")
    
    Me.comPort.Text = "COM" + Trim(str(dataDim.comPort))
    Me.baudRate.Text = Trim(str(dataDim.baudRate))
    Me.dataBit.Text = Trim(str(dataDim.dataBit))
    Me.stopBit.Text = Trim(str(dataDim.stopBit))
    Me.cr.Text = dataDim.cr
    
    Me.comPortOut.Text = "COM" + Trim(str(dataDim.comPortOut))
    Me.baudRateOut.AddItem ("1200")
    Me.baudRateOut.AddItem ("2400")
    Me.baudRateOut.AddItem ("4800")
    Me.baudRateOut.AddItem ("9600")
    Me.baudRateOut.AddItem ("14400")
    Me.baudRateOut.AddItem ("19200")
    Me.baudRateOut.AddItem ("38400")
    Me.baudRateOut.AddItem ("56000")
    Me.baudRateOut.AddItem ("57600")
    Me.baudRateOut.AddItem ("115200")
    Me.baudRateOut.AddItem ("194000")
    For i = 5 To 8
        Me.dataBitOut.AddItem (Trim(str(i)))
    Next
    Me.stopBitOut.AddItem ("1")
    Me.stopBitOut.AddItem ("2")
    Me.crOut.AddItem ("NONE")
    Me.crOut.AddItem ("ODD")
    Me.crOut.AddItem ("EVEN")
    
    Me.comPortOut.Text = "COM" + Trim(str(dataDim.comPortOut))
    Me.baudRateOut.Text = Trim(str(dataDim.baudRateOut))
    Me.dataBitOut.Text = Trim(str(dataDim.dataBitOut))
    Me.stopBitOut.Text = Trim(str(dataDim.stopBitOut))
    Me.crOut.Text = dataDim.crOut
    Me.outEn.value = IIf(dataDim.outEn, 1, 0)
End Sub

Private Sub ok_Click()
On Error GoTo errorHandler
    If (frmMain.sp.PortOpen) Then
        frmMain.LabInfo = "请先关闭串口"
    Else
        dataDim.comPort = Val(Right(Me.comPort.Text, Len(Me.comPort.Text) - 3))
        dataDim.baudRate = Val(Me.baudRate.Text)
        dataDim.dataBit = Val(Me.dataBit.Text)
        dataDim.stopBit = Val(Me.stopBit.Text)
        dataDim.cr = Me.cr.Text
        frmMain.sp.CommPort = dataDim.comPort
        frmMain.sp.Settings = Trim(str(dataDim.baudRate)) + "," + Left(dataDim.cr, 1) + "," + Trim(str(dataDim.dataBit)) + "," + Trim(str(dataDim.stopBit))
    End If
    If (frmMain.spOut.PortOpen) Then
        frmMain.spOut.PortOpen = False
    End If
    dataDim.comPortOut = Val(Right(Me.comPortOut.Text, Len(Me.comPortOut.Text) - 3))
    dataDim.baudRateOut = Val(Me.baudRateOut.Text)
    dataDim.dataBitOut = Val(Me.dataBitOut.Text)
    dataDim.stopBitOut = Val(Me.stopBitOut.Text)
    dataDim.crOut = Me.crOut.Text
    frmMain.spOut.CommPort = dataDim.comPortOut
    frmMain.spOut.Settings = Trim(str(dataDim.baudRateOut)) + "," + Left(dataDim.crOut, 1) + "," + Trim(str(dataDim.dataBitOut)) + "," + Trim(str(dataDim.stopBitOut))
    dataDim.outEn = Me.outEn.value
    If (dataDim.outEn) Then frmMain.spOut.PortOpen = True
    Unload Me
    Exit Sub
errorHandler:
    MsgBox (Error$)
    Exit Sub
End Sub
