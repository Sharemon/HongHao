VERSION 5.00
Begin VB.Form frmToler 
   Caption         =   "Form1"
   ClientHeight    =   5190
   ClientLeft      =   6105
   ClientTop       =   5640
   ClientWidth     =   8550
   LinkTopic       =   "Form1"
   ScaleHeight     =   5190
   ScaleWidth      =   8550
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   5520
      TabIndex        =   34
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   2400
      TabIndex        =   32
      Top             =   510
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   14
      Left            =   4800
      TabIndex        =   29
      Top             =   3720
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   13
      Left            =   4800
      TabIndex        =   28
      Top             =   3240
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   12
      Left            =   4800
      TabIndex        =   27
      Top             =   2760
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   11
      Left            =   4800
      TabIndex        =   26
      Top             =   2280
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   10
      Left            =   4800
      TabIndex        =   25
      Top             =   1800
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   19
      Left            =   6480
      TabIndex        =   24
      Top             =   3720
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   18
      Left            =   6480
      TabIndex        =   23
      Top             =   3240
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   17
      Left            =   6480
      TabIndex        =   22
      Top             =   2760
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   16
      Left            =   6480
      TabIndex        =   21
      Top             =   2280
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   15
      Left            =   6480
      TabIndex        =   20
      Top             =   1800
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   1440
      TabIndex        =   11
      Top             =   1800
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   3120
      TabIndex        =   10
      Top             =   1800
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   1440
      TabIndex        =   9
      Top             =   2280
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   3120
      TabIndex        =   8
      Top             =   2280
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   1440
      TabIndex        =   7
      Top             =   2760
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   3120
      TabIndex        =   6
      Top             =   2760
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   1440
      TabIndex        =   5
      Top             =   3240
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   7
      Left            =   3120
      TabIndex        =   4
      Top             =   3240
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   8
      Left            =   1440
      TabIndex        =   3
      Top             =   3720
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   9
      Left            =   3120
      TabIndex        =   2
      Top             =   3720
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   3600
      TabIndex        =   1
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   6960
      TabIndex        =   0
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   210
      Left            =   4200
      TabIndex        =   33
      Top             =   600
      Width           =   630
   End
   Begin VB.Line Line13 
      X1              =   6480
      X2              =   6480
      Y1              =   1800
      Y2              =   1200
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   7
      Left            =   5280
      TabIndex        =   31
      Top             =   1320
      Width           =   900
   End
   Begin VB.Line Line14 
      X1              =   8160
      X2              =   8160
      Y1              =   1800
      Y2              =   1200
   End
   Begin VB.Line Line15 
      X1              =   4800
      X2              =   8160
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   8
      Left            =   6960
      TabIndex        =   30
      Top             =   1320
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   210
      Left            =   240
      TabIndex        =   19
      Top             =   600
      Width           =   630
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   360
      TabIndex        =   18
      Top             =   1920
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   360
      TabIndex        =   17
      Top             =   2400
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   360
      TabIndex        =   16
      Top             =   2880
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   360
      TabIndex        =   15
      Top             =   3360
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   360
      TabIndex        =   14
      Top             =   3840
      Width           =   900
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   1800
      TabIndex        =   13
      Top             =   1320
      Width           =   900
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   6
      Left            =   3480
      TabIndex        =   12
      Top             =   1320
      Width           =   900
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   240
      Y1              =   1200
      Y2              =   4200
   End
   Begin VB.Line Line2 
      X1              =   240
      X2              =   4800
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line3 
      X1              =   1440
      X2              =   240
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line4 
      X1              =   1440
      X2              =   240
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line5 
      X1              =   1440
      X2              =   240
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Line Line6 
      X1              =   1440
      X2              =   240
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line7 
      X1              =   1440
      X2              =   240
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Line Line8 
      X1              =   1440
      X2              =   240
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Line Line9 
      X1              =   1440
      X2              =   1440
      Y1              =   1800
      Y2              =   1200
   End
   Begin VB.Line Line10 
      X1              =   3120
      X2              =   3120
      Y1              =   1800
      Y2              =   1200
   End
   Begin VB.Line Line11 
      X1              =   4800
      X2              =   4800
      Y1              =   1800
      Y2              =   1200
   End
End
Attribute VB_Name = "frmToler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Select Case Len(Text3.Text)
Case 1, 2
    Dim m As String
    m = Format(Text3.Text, "00")
    Data2_3(1) = Asc(Left(m, 1))
    Data2_3(2) = Asc(Right(m, 1))
    WriteINI "Custom", "DataBit2", Data2_3(1), App.Path & "\Config.ini"
    WriteINI "Custom", "DataBit3", Data2_3(2), App.Path & "\Config.ini"
Case Else
    MsgBox IIf(Lang, "Data property number must be between 1 to 99!", "输出数据属性值为1到99的整数！"), vbCritical, Me.Caption
    Exit Sub
End Select
FrmMain.Timer8.Enabled = False
FrmMain.Timer8.Interval = Val(Text2.Text)
FrmMain.Timer8.Enabled = True
WriteINI "Custom", "timer8", Val(Text2.Text), App.Path & "\Config.ini"
For I = 0 To 4
DelayTime(I + 1) = Val(Text1(2 * I))
Tolerance(I + 1) = Val(Text1(2 * I + 1)) / 100
ZeroToler(I + 1) = Val(Text1(I + 10)) / 100
DepartToler(I + 1) = Val(Text1(I + 15)) / 100
If RANGe = RangArry(I + 1) Then
DelayTimeW = DelayTime(I + 1)
ToleranceW = Tolerance(I + 1)
DepartTolerW = DepartToler(I + 1)
ZeroTolerW = ZeroToler(I + 1)
End If
Next I
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Caption = IIf(Lang, "Threshold Setting", "校准及门限设定")
Command1.Caption = IIf(Lang, "OK", "确定")
Command2.Caption = IIf(Lang, "Cancel", "取消")
Label1.Caption = IIf(Lang, "Cali_Data_Sending_Delay(ms)", "校准数据发送间隔(毫秒)")
Label3.Caption = IIf(Lang, "Data property", "输出数据属性")
Label2(5).Caption = IIf(Lang, "Time_delay", "稳定时间") & "(s)"
Label2(6).Caption = IIf(Lang, "Threshold", "稳定门限") & "(%)"
Label2(7).Caption = IIf(Lang, "Zero_Range", "零位范围") & "(%)"
Label2(8).Caption = IIf(Lang, "", "同值门限") & "(%)"
Label2(0).Caption = "100mV"
Label2(1).Caption = "1   V"
Label2(2).Caption = "10  V"
Label2(3).Caption = "100 V"
Label2(4).Caption = "1000V"
For I = 0 To 4
Text1(2 * I) = DelayTime(I + 1)
Text1(2 * I + 1) = FormatNumber(Tolerance(I + 1) * 100, 4, vbTrue)
Text1(I + 10) = FormatNumber(ZeroToler(I + 1) * 100, 4, vbTrue)
Text1(I + 15) = FormatNumber(DepartToler(I + 1) * 100, 4, vbTrue)
Next I
Text2.Text = FrmMain.Timer8.Interval
Text3.Text = CStr(Chr(Data2_3(1))) & CStr(Chr(Data2_3(2)))
End Sub
