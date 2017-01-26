VERSION 5.00
Begin VB.Form KEI 
   Caption         =   "Form1"
   ClientHeight    =   3480
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5685
   LinkTopic       =   "Form1"
   ScaleHeight     =   3480
   ScaleWidth      =   5685
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   3840
      TabIndex        =   12
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   3840
      TabIndex        =   11
      Top             =   720
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   2730
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   2730
      Begin VB.ComboBox Combo6 
         Height          =   300
         ItemData        =   "KEI.frx":0000
         Left            =   1335
         List            =   "KEI.frx":000D
         TabIndex        =   5
         Text            =   "Combo6"
         Top             =   2175
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
      Begin VB.ComboBox Combo4 
         Height          =   300
         ItemData        =   "KEI.frx":0029
         Left            =   1320
         List            =   "KEI.frx":002B
         TabIndex        =   3
         Text            =   "Combo4"
         Top             =   1200
         Width           =   1200
      End
      Begin VB.ComboBox Combo3 
         Height          =   300
         ItemData        =   "KEI.frx":002D
         Left            =   1335
         List            =   "KEI.frx":002F
         TabIndex        =   2
         Text            =   "Combo3"
         Top             =   765
         Width           =   1200
      End
      Begin VB.ComboBox Combo2 
         Height          =   300
         ItemData        =   "KEI.frx":0031
         Left            =   1335
         List            =   "KEI.frx":0033
         TabIndex        =   1
         Text            =   "Combo2"
         Top             =   300
         Width           =   1200
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "校验方式"
         Height          =   240
         Left            =   285
         TabIndex        =   10
         Top             =   2190
         Width           =   960
      End
      Begin VB.Label Label5 
         Caption         =   "停止位"
         Height          =   240
         Left            =   255
         TabIndex        =   9
         Top             =   1230
         Width           =   840
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
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "波特率"
         Height          =   240
         Left            =   255
         TabIndex        =   7
         Top             =   795
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "串口号"
         Height          =   240
         Left            =   255
         TabIndex        =   6
         Top             =   345
         Width           =   720
      End
   End
End
Attribute VB_Name = "KEI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo errhdl
Port0 = CLng(Combo2.Text)
FrmMain.MSComm0.CommPort = Port0
Unload Me
errhdl: If err.Number <> 0 Then MsgBox IIf(Lang, "Error code:" & err.Number & vbNewLine & err.Description, "错误号:" & err.Number & vbNewLine & err.Description) & "," & Me.Caption
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Caption = IIf(Lang, "Configuration of KEITHLEY2000", "吉时利2000表设置")
Command1.Caption = IIf(Lang, "OK", "确定")
Command2.Caption = IIf(Lang, "Cancel", "取消")
Label2.Caption = IIf(Lang, "Port Number", "串口号")
Label3.Caption = IIf(Lang, "Baud rate", "波特率")
Label5.Caption = IIf(Lang, "Stop bit", "停止位")
Label4.Caption = IIf(Lang, "Data bits", "数据位")
Label6.Caption = IIf(Lang, "Check mode", "校验方式")
For i = 0 To 15
Combo2.AddItem i + 1
Next i
Combo2.Text = Port0
Combo3.Clear
Combo4.Clear
Combo5.Clear
Combo6.Clear
Combo3.AddItem 1200
Combo3.AddItem 2400
Combo3.AddItem 4800
Combo3.AddItem 9600
Combo3.AddItem 19200
Combo4.AddItem 1
Combo4.AddItem 2
Combo5.AddItem 7
Combo5.AddItem 8
Combo6.AddItem IIf(Lang, "Odd parity", "奇校验")
Combo6.AddItem IIf(Lang, "Even parity", "偶校验")
Combo6.AddItem IIf(Lang, "None", "无校验")
Combo3.Text = 9600
Combo4.Text = 1
Combo5.Text = 8
Combo6.Text = IIf(Lang, "None", "无校验")
End Sub
