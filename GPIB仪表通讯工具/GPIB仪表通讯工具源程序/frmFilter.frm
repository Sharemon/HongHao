VERSION 5.00
Begin VB.Form frmFilter 
   Caption         =   "ÂË²¨ÉèÖÃ"
   ClientHeight    =   8250
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4035
   LinkTopic       =   "Form1"
   ScaleHeight     =   8250
   ScaleWidth      =   4035
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton save 
      Caption         =   "±£ ´æ"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   30
      Top             =   6840
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cancel 
      Caption         =   "È¡ Ïû"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   2160
      TabIndex        =   29
      Top             =   7560
      Width           =   1400
   End
   Begin VB.TextBox filter 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   11
      Left            =   2160
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   6960
      Width           =   1215
   End
   Begin VB.TextBox filter 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   10
      Left            =   2160
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   6465
      Width           =   1215
   End
   Begin VB.TextBox filter 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   9
      Left            =   2160
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   5970
      Width           =   1215
   End
   Begin VB.TextBox filter 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   8
      Left            =   2160
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   5475
      Width           =   1215
   End
   Begin VB.TextBox filter 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   7
      Left            =   2160
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   4965
      Width           =   1215
   End
   Begin VB.ComboBox rangeChoice 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      ItemData        =   "frmFilter.frx":0000
      Left            =   2040
      List            =   "frmFilter.frx":000D
      TabIndex        =   9
      Top             =   240
      Width           =   1335
   End
   Begin VB.TextBox filter 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   6
      Left            =   2160
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   4470
      Width           =   1215
   End
   Begin VB.TextBox filter 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   5
      Left            =   2160
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   3975
      Width           =   1215
   End
   Begin VB.CommandButton ok 
      Caption         =   "È· ¶¨"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   240
      TabIndex        =   5
      Top             =   7560
      Width           =   1400
   End
   Begin VB.TextBox filter 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   4
      Left            =   2160
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox filter 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   3
      Left            =   2160
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   2985
      Width           =   1215
   End
   Begin VB.TextBox filter 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   2
      Left            =   2160
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   2490
      Width           =   1215
   End
   Begin VB.TextBox filter 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   1
      Left            =   2160
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1995
      Width           =   1215
   End
   Begin VB.TextBox filter 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   2160
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Á¿³ÌÑ¡Ôñ£º"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   28
      Top             =   240
      Width           =   1215
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   3840
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label2 
      Caption         =   "ÂË²¨´ÎÊý"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   27
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Á¿³Ì"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   26
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label rangeShow 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   25
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label rangeShow 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   600
      TabIndex        =   24
      Top             =   2055
      Width           =   975
   End
   Begin VB.Label rangeShow 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   600
      TabIndex        =   23
      Top             =   2535
      Width           =   975
   End
   Begin VB.Label rangeShow 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   600
      TabIndex        =   22
      Top             =   3030
      Width           =   975
   End
   Begin VB.Label rangeShow 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   600
      TabIndex        =   21
      Top             =   3525
      Width           =   975
   End
   Begin VB.Label rangeShow 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   600
      TabIndex        =   20
      Top             =   4500
      Width           =   975
   End
   Begin VB.Label rangeShow 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   600
      TabIndex        =   19
      Top             =   4995
      Width           =   975
   End
   Begin VB.Label rangeShow 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   600
      TabIndex        =   18
      Top             =   5475
      Width           =   975
   End
   Begin VB.Label rangeShow 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   600
      TabIndex        =   17
      Top             =   5970
      Width           =   975
   End
   Begin VB.Label rangeShow 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   600
      TabIndex        =   16
      Top             =   6465
      Width           =   975
   End
   Begin VB.Label rangeShow 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   600
      TabIndex        =   15
      Top             =   6960
      Width           =   975
   End
   Begin VB.Label rangeShow 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   600
      TabIndex        =   6
      Top             =   4005
      Width           =   975
   End
End
Attribute VB_Name = "frmFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cancel_Click()
    Unload Me
End Sub

Private Sub save_Click()
    Call savePara
End Sub

Private Sub filter_Change(Index As Integer)
    If (Not IsNumeric(Me.filter(Index).Text) And Len(Me.filter(Index).Text) > 0) Then
        Me.filter(Index).Text = Left(Me.filter(Index).Text, Len(Me.filter(Index).Text) - 1)
        Me.filter(Index).SelStart = Len(Me.filter(Index).Text)
    End If
End Sub

Private Sub Form_Load()
    If (frmMain.connectFlag) Then
        Me.rangeChoice.ListIndex = frmMain.VIO
    Else
        Me.rangeChoice.ListIndex = 0
    End If
End Sub

Private Sub ok_Click()
    Call savePara
    Unload Me
End Sub

Private Sub savePara()
    Dim i As Integer
    For i = 0 To 11
        If (Val(Me.filter(i).Text) = 0) Then Me.filter(i).Text = "1"
        dataDim.filter(Me.rangeChoice.ListIndex, i) = Val(Me.filter(i).Text)
    Next
End Sub

Private Sub rangeChoice_Click()
    Dim i As Integer
    Select Case Me.rangeChoice.ListIndex
    Case 0
        Me.rangeShow(8).Visible = False
        Me.rangeShow(9).Visible = False
        Me.rangeShow(10).Visible = False
        Me.rangeShow(11).Visible = False
        Me.filter(8).Visible = False
        Me.filter(9).Visible = False
        Me.filter(10).Visible = False
        Me.filter(11).Visible = False
        Me.rangeShow(0).Caption = "100uV"
        Me.rangeShow(1).Caption = "1mV"
        Me.rangeShow(2).Caption = "10mV"
        Me.rangeShow(3).Caption = "100mV"
        Me.rangeShow(4).Caption = "1V"
        Me.rangeShow(5).Caption = "10V"
        Me.rangeShow(6).Caption = "100V"
        Me.rangeShow(7).Caption = "1000V"
        Me.Height = 6800
        Me.ok.Top = 5570
        Me.cancel.Top = 5570
        Me.save.Top = 5570
        For i = 0 To 7
            Me.filter(i).Text = dataDim.filter(0, i)
        Next
    Case 1
        Me.rangeShow(8).Visible = True
        Me.rangeShow(9).Visible = False
        Me.rangeShow(10).Visible = False
        Me.rangeShow(11).Visible = False
        Me.filter(8).Visible = True
        Me.filter(9).Visible = False
        Me.filter(10).Visible = False
        Me.filter(11).Visible = False
        Me.rangeShow(0).Caption = "1uA"
        Me.rangeShow(1).Caption = "10uA"
        Me.rangeShow(2).Caption = "100uA"
        Me.rangeShow(3).Caption = "1mA"
        Me.rangeShow(4).Caption = "10mA"
        Me.rangeShow(5).Caption = "100mA"
        Me.rangeShow(6).Caption = "1A"
        Me.rangeShow(7).Caption = "3A"
        Me.rangeShow(8).Caption = "10A"
        Me.Height = 7270
        Me.ok.Top = 6080
        Me.cancel.Top = 6080
        Me.save.Top = 6080
        For i = 0 To 8
            Me.filter(i).Text = dataDim.filter(1, i)
        Next
    Case 2
        Me.rangeShow(8).Visible = True
        Me.rangeShow(9).Visible = True
        Me.rangeShow(10).Visible = True
        Me.rangeShow(11).Visible = True
        Me.filter(8).Visible = True
        Me.filter(9).Visible = True
        Me.filter(10).Visible = True
        Me.filter(11).Visible = True
        Me.rangeShow(0).Caption = "1¦¸"
        Me.rangeShow(1).Caption = "10¦¸"
        Me.rangeShow(2).Caption = "100¦¸"
        Me.rangeShow(3).Caption = "1K¦¸"
        Me.rangeShow(4).Caption = "10K¦¸"
        Me.rangeShow(5).Caption = "100K¦¸"
        Me.rangeShow(6).Caption = "1M¦¸"
        Me.rangeShow(7).Caption = "10M¦¸"
        Me.rangeShow(8).Caption = "100M¦¸"
        Me.rangeShow(9).Caption = "1G¦¸"
        Me.rangeShow(10).Caption = "10G¦¸"
        Me.rangeShow(11).Caption = "100G¦¸"
        Me.Height = 8820
        Me.ok.Top = 7560
        Me.cancel.Top = 7560
        Me.save.Top = 7560
        For i = 0 To 11
            Me.filter(i).Text = dataDim.filter(2, i)
        Next
    End Select
End Sub

Private Sub rangeChoice_DropDown()
    Call savePara
End Sub
