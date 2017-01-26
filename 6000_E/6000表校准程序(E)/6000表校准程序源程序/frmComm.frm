VERSION 5.00
Begin VB.Form frmComm 
   Caption         =   "Form1"
   ClientHeight    =   4080
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   ScaleHeight     =   4080
   ScaleWidth      =   6495
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   0
      Top             =   3720
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Check1"
      Height          =   375
      Left            =   3360
      TabIndex        =   25
      Top             =   120
      Value           =   1  'Checked
      Width           =   2775
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   375
      Left            =   240
      TabIndex        =   24
      Top             =   120
      Width           =   2775
   End
   Begin VB.Frame Frame2 
      Caption         =   "6000万用表参数--输出"
      Height          =   2730
      Left            =   3360
      TabIndex        =   13
      Top             =   600
      Width           =   2730
      Begin VB.ComboBox Combo11 
         Height          =   300
         ItemData        =   "frmComm.frx":0000
         Left            =   1335
         List            =   "frmComm.frx":000D
         TabIndex        =   18
         Text            =   "Combo6"
         Top             =   2175
         Width           =   1200
      End
      Begin VB.ComboBox Combo10 
         Height          =   300
         Left            =   1335
         TabIndex        =   17
         Text            =   "Combo5"
         Top             =   1680
         Width           =   1200
      End
      Begin VB.ComboBox Combo9 
         Height          =   300
         ItemData        =   "frmComm.frx":0029
         Left            =   1320
         List            =   "frmComm.frx":002B
         TabIndex        =   16
         Text            =   "Combo4"
         Top             =   1200
         Width           =   1200
      End
      Begin VB.ComboBox Combo8 
         Height          =   300
         ItemData        =   "frmComm.frx":002D
         Left            =   1335
         List            =   "frmComm.frx":002F
         TabIndex        =   15
         Text            =   "Combo3"
         Top             =   765
         Width           =   1200
      End
      Begin VB.ComboBox Combo7 
         Height          =   300
         ItemData        =   "frmComm.frx":0031
         Left            =   1335
         List            =   "frmComm.frx":0033
         TabIndex        =   14
         Text            =   "Combo2"
         Top             =   300
         Width           =   1200
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "校验方式"
         Height          =   240
         Left            =   285
         TabIndex        =   23
         Top             =   2190
         Width           =   960
      End
      Begin VB.Label Label9 
         Caption         =   "停止位"
         Height          =   240
         Left            =   255
         TabIndex        =   22
         Top             =   1230
         Width           =   840
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "数据位"
         Height          =   240
         Left            =   285
         TabIndex        =   21
         Top             =   1755
         Width           =   720
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "波特率"
         Height          =   240
         Left            =   255
         TabIndex        =   20
         Top             =   795
         Width           =   720
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "串口号"
         Height          =   240
         Left            =   255
         TabIndex        =   19
         Top             =   345
         Width           =   720
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   4880
      TabIndex        =   12
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   2160
      TabIndex        =   11
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "6000万用表参数--输入"
      Height          =   2730
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   2730
      Begin VB.ComboBox Combo2 
         Height          =   300
         ItemData        =   "frmComm.frx":0035
         Left            =   1335
         List            =   "frmComm.frx":0037
         TabIndex        =   1
         Text            =   "Combo2"
         Top             =   300
         Width           =   1200
      End
      Begin VB.ComboBox Combo3 
         Height          =   300
         ItemData        =   "frmComm.frx":0039
         Left            =   1335
         List            =   "frmComm.frx":003B
         TabIndex        =   2
         Text            =   "Combo3"
         Top             =   765
         Width           =   1200
      End
      Begin VB.ComboBox Combo4 
         Height          =   300
         ItemData        =   "frmComm.frx":003D
         Left            =   1320
         List            =   "frmComm.frx":003F
         TabIndex        =   3
         Text            =   "Combo4"
         Top             =   1200
         Width           =   1200
      End
      Begin VB.ComboBox Combo5 
         Height          =   300
         Left            =   1335
         TabIndex        =   4
         Text            =   "Combo5"
         Top             =   1680
         Width           =   1200
      End
      Begin VB.ComboBox Combo6 
         Height          =   300
         ItemData        =   "frmComm.frx":0041
         Left            =   1335
         List            =   "frmComm.frx":004E
         TabIndex        =   5
         Text            =   "Combo6"
         Top             =   2175
         Width           =   1200
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "串口号"
         Height          =   240
         Left            =   255
         TabIndex        =   10
         Top             =   345
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "波特率"
         Height          =   240
         Left            =   255
         TabIndex        =   9
         Top             =   795
         Width           =   720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "数据位"
         Height          =   240
         Left            =   285
         TabIndex        =   8
         Top             =   1755
         Width           =   720
      End
      Begin VB.Label Label5 
         Caption         =   "停止位"
         Height          =   240
         Left            =   255
         TabIndex        =   7
         Top             =   1230
         Width           =   840
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "校验方式"
         Height          =   240
         Left            =   285
         TabIndex        =   6
         Top             =   2190
         Width           =   960
      End
   End
End
Attribute VB_Name = "frmComm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
CommIn = IIf(Check1.Value, True, False)
End Sub

Private Sub Check2_Click()
CommOut = IIf(Check2.Value, True, False)
End Sub

Private Sub Command1_Click()
'On Error GoTo errhdl
If FrmMain.MSComm.PortOpen = False And Check2.Value = 1 Then
Port2 = Val(Combo7.Text)
FrmMain.MSComm.CommPort = Port2
FrmMain.MSComm.PortOpen = True
OutPut = True
'frmMain.Check1.Visible = True
End If

If FrmMain.MSComm1.PortOpen = False And Check1.Value = 1 Then
Port1 = Val(Combo2.Text)
FrmMain.MSComm1.CommPort = Port1
FrmMain.MSComm1.PortOpen = True
FrmMain.MSComm1.OutPut = Trim("%00;03" & vbCrLf)
End If

If Check2.Value = 0 Then showhideWK False

Unload Me
ReadAuto = True
FrmMain.Refresh


'errhdl: If Err.Number <> 0 Then MsgBox IIf(Lang, "Err found the port settings, please try again!", "串口设置有误，请重新设置！")
'errhdl: If Err.Number <> 0 Then MsgBox IIf(Lang, "Error code:" & Err.Number & vbNewLine & Err.Description, "错误号:" & Err.Number & vbNewLine & Err.Description)

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
If FrmMain.MSComm.PortOpen <> False Then FrmMain.MSComm.PortOpen = False
If FrmMain.MSComm1.PortOpen <> False Then FrmMain.MSComm1.PortOpen = False
Frame1.Caption = IIf(Lang, "Parameters of 6000 DMM--Instruction Port", "6000万用表参数--指令口")
Frame2.Caption = IIf(Lang, "Parameters of 6000 DMM--Adjust Port", "6000万用表参数--校准口")
Me.Caption = IIf(Lang, "Comm Settings", "串口设置")
'Label1.Caption = IIf(Lang, "  Please choose appropriate" & Chr(10) & Chr(10) & "parameters in the right box then" & Chr(10) & Chr(10) & "click the command button to output" & Chr(10) & Chr(10) & "data to 6000 DMM.", "  请在右侧选定适当的通讯参数后点击" & Chr(10) & Chr(10) & "确定，数据将同步输出到6000数字万用" & Chr(10) & Chr(10) & "表上。")
Label2.Caption = IIf(Lang, "Port Number", "串口号")
Label3.Caption = IIf(Lang, "Baud rate", "波特率")
Label5.Caption = IIf(Lang, "Stop bit", "停止位")
Label4.Caption = IIf(Lang, "Data bits", "数据位")
Label6.Caption = IIf(Lang, "Check mode", "校验方式")

Label7.Caption = IIf(Lang, "Port Number", "串口号")
Label8.Caption = IIf(Lang, "Baud rate", "波特率")
Label9.Caption = IIf(Lang, "Stop bit", "停止位")
Label10.Caption = IIf(Lang, "Data bits", "数据位")
Label11.Caption = IIf(Lang, "Check mode", "校验方式")

Check1.Caption = IIf(Lang, "Enable input from 6000DMM", "开启6000表输入")
Check2.Caption = IIf(Lang, "Enable output from 6000DMM", "开启6000表输出")

Command1.Caption = IIf(Lang, "OK", "确定")
Command2.Caption = IIf(Lang, "Cancel", "取消")
For i = 0 To 15
Combo2.AddItem i + 1
Combo7.AddItem i + 1
Next i
Combo2.Text = Port1
Combo3.Clear
Combo4.Clear
Combo5.Clear
Combo6.Clear
Combo7.Text = Port2
Combo8.Clear
Combo9.Clear
Combo10.Clear
Combo11.Clear
Combo3.AddItem 1200
Combo3.AddItem 2400
Combo3.AddItem 4800
Combo3.AddItem 9600
Combo3.AddItem 19200
Combo8.AddItem 1200
Combo8.AddItem 2400
Combo8.AddItem 4800
Combo8.AddItem 9600
Combo8.AddItem 19200
Combo4.AddItem 1
Combo4.AddItem 2
Combo9.AddItem 1
Combo9.AddItem 2
Combo5.AddItem 7
Combo5.AddItem 8
Combo10.AddItem 7
Combo10.AddItem 8
Combo6.AddItem IIf(Lang, "Odd parity", "奇校验")
Combo6.AddItem IIf(Lang, "Even parity", "偶校验")
Combo6.AddItem IIf(Lang, "None", "无校验")
Combo11.AddItem IIf(Lang, "Odd parity", "奇校验")
Combo11.AddItem IIf(Lang, "Even parity", "偶校验")
Combo11.AddItem IIf(Lang, "None", "无校验")
Combo3.Text = 9600
Combo4.Text = 1
Combo5.Text = 8
Combo6.Text = IIf(Lang, "None", "无校验")
Combo8.Text = 9600
Combo9.Text = 1
Combo10.Text = 8
Combo11.Text = IIf(Lang, "None", "无校验")
Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
Me.SetFocus
Timer1.Enabled = False
End Sub
