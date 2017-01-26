VERSION 5.00
Begin VB.Form Adjust 
   AutoRedraw      =   -1  'True
   ClientHeight    =   5760
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5145
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5760
   ScaleWidth      =   5145
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2400
      TabIndex        =   23
      Top             =   4530
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2400
      TabIndex        =   21
      Top             =   3930
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   3600
      TabIndex        =   19
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1200
      TabIndex        =   18
      Top             =   5160
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   9
      Left            =   3120
      TabIndex        =   10
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   8
      Left            =   1440
      TabIndex        =   9
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   7
      Left            =   3120
      TabIndex        =   8
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   3120
      TabIndex        =   6
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   1440
      TabIndex        =   5
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   3120
      TabIndex        =   4
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   1440
      TabIndex        =   3
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   3120
      TabIndex        =   2
      Top             =   1320
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   1440
      TabIndex        =   1
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Line Line13 
      X1              =   4800
      X2              =   4800
      Y1              =   1320
      Y2              =   720
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "通用除常数因子:"
      Height          =   210
      Left            =   480
      TabIndex        =   22
      Top             =   4620
      Width           =   1575
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "通用乘常数因子:"
      Height          =   210
      Left            =   480
      TabIndex        =   20
      Top             =   4005
      Width           =   1575
   End
   Begin VB.Line Line11 
      X1              =   4800
      X2              =   4800
      Y1              =   1320
      Y2              =   720
   End
   Begin VB.Line Line10 
      X1              =   3120
      X2              =   3120
      Y1              =   1320
      Y2              =   720
   End
   Begin VB.Line Line9 
      X1              =   1440
      X2              =   1440
      Y1              =   1320
      Y2              =   720
   End
   Begin VB.Line Line8 
      X1              =   1440
      X2              =   240
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Line Line7 
      X1              =   1440
      X2              =   240
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line6 
      X1              =   1440
      X2              =   240
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Line Line5 
      X1              =   1440
      X2              =   240
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line4 
      X1              =   1440
      X2              =   240
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line3 
      X1              =   1440
      X2              =   240
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line2 
      X1              =   240
      X2              =   4800
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   240
      Y1              =   720
      Y2              =   3720
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
      Left            =   3360
      TabIndex        =   17
      Top             =   840
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
      TabIndex        =   16
      Top             =   840
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
      Index           =   3
      Left            =   360
      TabIndex        =   14
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
      Index           =   2
      Left            =   360
      TabIndex        =   13
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
      Index           =   1
      Left            =   360
      TabIndex        =   12
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
      Index           =   0
      Left            =   360
      TabIndex        =   11
      Top             =   1440
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   210
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   630
   End
End
Attribute VB_Name = "Adjust"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
If Val(Text3.Text) = 0 Then
MsgBox IIf(Lang, "DivCons can not be ZERO!", "除常数因子不能为零！"), vbCritical
Exit Sub
End If
For I = 1 To 5
Adjnum(I).POS = Val(Text1((I - 1) * 2).Text)
Adjnum(I).Neg = Val(Text1(I * 2 - 1).Text)
Next
MultCons = Val(Text2.Text)
DivCons = Val(Text3.Text)
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Command1.Caption = IIf(Lang, "OK", "确 定")
Command2.Caption = IIf(Lang, "Cancel", "取 消")
Adjust.Caption = IIf(Lang, "Set adjust numbers", "设定修正值")
Label1.Caption = IIf(Lang, "Please enter the adjust number for each RANGE", "请输入各量程的修正值")
Label2(0).Caption = "100mV"
Label2(1).Caption = "1V"
Label2(2).Caption = "10V"
Label2(3).Caption = "100V"
Label2(4).Caption = "1000V"
Label2(5).Caption = IIf(Lang, "Positive", "正")
Label2(6).Caption = IIf(Lang, "Negative", "负")
For I = 0 To 4
Text1(I * 2).Text = Format$(Adjnum(I + 1).POS, "#0.00000000")
Text1(I * 2 + 1).Text = Format$(Adjnum(I + 1).Neg, "#0.00000000")
Next
Text2.Text = FormatNumber(MultCons, 7, vbTrue)
Text3.Text = FormatNumber(DivCons, 7, vbTrue)
Label3.Caption = IIf(Lang, "MultCons:", "乘常数因子:")
Label4.Caption = IIf(Lang, "DivCons:", "除常数因子:")
End Sub

Private Function ImportNum(KeyIn As Integer, ValidateString As String, Editable As Boolean) As Integer
Dim ValidateList As String
Dim KeyOut As Integer
If Editable = True Then
ValidateList = UCase(ValidateString) & Chr(8)
Else
ValidateList = UCase(ValidateString)
End If
If InStr(1, ValidateList, UCase(Chr(KeyIn)), 1) > 0 Then
KeyOut = KeyIn
Else
KeyOut = 0
Beep
End If
ImportNum = KeyOut
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Select Case RANGe
Case 0.1
    Filter = Filter00
Case 1
    Filter = Filter01
Case 10, 100, 1000
    Filter = Filter02
End Select
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = ImportNum(KeyAscii, "0123456789" & Chr(13) & "." & "-", True)
End Sub
