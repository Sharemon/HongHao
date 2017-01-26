VERSION 5.00
Begin VB.Form frmGPIB 
   Caption         =   "GPIBÉèÖÃ"
   ClientHeight    =   2580
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3750
   LinkTopic       =   "Form2"
   ScaleHeight     =   2580
   ScaleWidth      =   3750
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
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
      Left            =   2040
      TabIndex        =   5
      Top             =   1800
      Width           =   1400
   End
   Begin VB.TextBox GPIBnum 
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
      Left            =   2040
      TabIndex        =   4
      Text            =   "0"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox GPIBAddress 
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
      Left            =   2040
      TabIndex        =   2
      Text            =   "22"
      Top             =   360
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
      TabIndex        =   0
      Top             =   1800
      Width           =   1400
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "GPIBÐòºÅ£º"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
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
      Top             =   1080
      Width           =   1110
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "GPIBµØÖ·£º"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
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
      Top             =   360
      Width           =   1110
   End
End
Attribute VB_Name = "frmGPIB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.GPIBAddress.Text = dataDim.GPIBaddr
    Me.GPIBnum.Text = dataDim.GPIBnum
End Sub

Private Sub GPIBAddress_Change()
    If ((Not IsNumeric(Me.GPIBAddress.Text)) And (Len(Me.GPIBAddress.Text) > 0)) Then
        Me.GPIBAddress.Text = Left(Me.GPIBAddress.Text, Len(Me.GPIBAddress.Text) - 1)
        Me.GPIBAddress.SelStart = Len(Me.GPIBAddress.Text)
    End If
End Sub

Private Sub GPIBnum_Change()
    If ((Not IsNumeric(Me.GPIBnum.Text)) And (Len(Me.GPIBnum.Text) > 0)) Then
        Me.GPIBnum.Text = Left(Me.GPIBnum.Text, Len(Me.GPIBnum.Text) - 1)
        Me.GPIBnum.SelStart = Len(Me.GPIBnum.Text)
    End If
End Sub

Private Sub ok_Click()
    dataDim.GPIBaddr = Val(Me.GPIBAddress.Text)
    dataDim.GPIBnum = Val(Me.GPIBnum.Text)
    Unload Me
End Sub
