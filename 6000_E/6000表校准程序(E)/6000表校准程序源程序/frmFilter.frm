VERSION 5.00
Begin VB.Form frmFilter 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "对话框标题"
   ClientHeight    =   4200
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4095
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   4095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton CancelButton 
      Caption         =   "Command2"
      Height          =   495
      Left            =   2520
      TabIndex        =   6
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "Command1"
      Height          =   495
      Left            =   840
      TabIndex        =   5
      Top             =   3600
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
      Index           =   1
      Left            =   2040
      TabIndex        =   0
      Top             =   840
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
      Left            =   2040
      TabIndex        =   1
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   2040
      TabIndex        =   3
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
      Index           =   5
      Left            =   2040
      TabIndex        =   4
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
      Index           =   3
      Left            =   2040
      TabIndex        =   2
      Top             =   1800
      Width           =   1695
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
      Index           =   4
      Left            =   720
      TabIndex        =   13
      Top             =   2880
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
      Index           =   3
      Left            =   720
      TabIndex        =   12
      Top             =   2400
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
      Index           =   2
      Left            =   720
      TabIndex        =   11
      Top             =   1920
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
      Index           =   1
      Left            =   720
      TabIndex        =   10
      Top             =   1440
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
      Index           =   0
      Left            =   720
      TabIndex        =   9
      Top             =   960
      Width           =   900
   End
   Begin VB.Line Line8 
      X1              =   360
      X2              =   360
      Y1              =   840
      Y2              =   3240
   End
   Begin VB.Line Line7 
      X1              =   360
      X2              =   2040
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line6 
      X1              =   360
      X2              =   2040
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Line Line5 
      X1              =   360
      X2              =   2040
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line4 
      X1              =   360
      X2              =   2040
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line3 
      X1              =   360
      X2              =   2040
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line1 
      X1              =   360
      X2              =   2040
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line Line11 
      X1              =   3720
      X2              =   3720
      Y1              =   840
      Y2              =   240
   End
   Begin VB.Line Line10 
      X1              =   2040
      X2              =   2040
      Y1              =   840
      Y2              =   240
   End
   Begin VB.Line Line9 
      X1              =   360
      X2              =   360
      Y1              =   840
      Y2              =   240
   End
   Begin VB.Line Line2 
      X1              =   360
      X2              =   3720
      Y1              =   240
      Y2              =   240
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
      Left            =   2400
      TabIndex        =   8
      Top             =   360
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
      Left            =   720
      TabIndex        =   7
      Top             =   360
      Width           =   900
   End
   Begin VB.Line Line12 
      X1              =   360
      X2              =   3720
      Y1              =   3480
      Y2              =   3480
   End
End
Attribute VB_Name = "frmFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CancelButton_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Caption = IIf(Lang, "Filter setting", "设定滤波值")
CancelButton.Caption = IIf(Lang, "Cancel", "取消")
OKButton.Caption = IIf(Lang, "OK", "确定")
Label2(5).Caption = IIf(Lang, "RANGe", "量程")
Label2(6).Caption = IIf(Lang, "Filter", "滤波值(次)")
For I = 1 To 5
Label2(I - 1).Caption = 10 ^ (I - 2) & "V"
Text1(I).Text = FilterArry(I)
Next I
Label2(0).Caption = "100mV"
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

Private Sub OKButton_Click()
For I = 1 To 5
FilterArry(I) = Val(Text1(I).Text)
If RANGe = RangArry(I) Then Filter = FilterArry(I)
Next
Cnt = 0
ss = 0
ReDim numArry(1 To Filter)
Unload Me
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = ImportNum(KeyAscii, "0123456789", True)
End Sub
