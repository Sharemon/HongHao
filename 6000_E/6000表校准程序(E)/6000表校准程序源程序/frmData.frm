VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmData 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Data Window"
   ClientHeight    =   8205
   ClientLeft      =   13485
   ClientTop       =   3750
   ClientWidth     =   4335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8205
   ScaleWidth      =   4335
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   3720
      Top             =   7680
   End
   Begin RichTextLib.RichTextBox RText0 
      Height          =   7335
      Left            =   360
      TabIndex        =   3
      Top             =   240
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   12938
      _Version        =   393217
      Enabled         =   0   'False
      ReadOnly        =   -1  'True
      TextRTF         =   $"frmData.frx":0000
   End
   Begin RichTextLib.RichTextBox RText1 
      Height          =   7335
      Left            =   1080
      TabIndex        =   0
      Top             =   240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   12938
      _Version        =   393217
      Enabled         =   0   'False
      ReadOnly        =   -1  'True
      TextRTF         =   $"frmData.frx":009D
   End
   Begin RichTextLib.RichTextBox RText2 
      Height          =   7335
      Left            =   2520
      TabIndex        =   4
      Top             =   240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   12938
      _Version        =   393217
      Enabled         =   0   'False
      ReadOnly        =   -1  'True
      TextRTF         =   $"frmData.frx":013A
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   360
      TabIndex        =   5
      Top             =   7680
      Width           =   90
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   360
      TabIndex        =   2
      Top             =   7680
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   3240
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   90
   End
End
Attribute VB_Name = "frmData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    frmData.RText0.Text = ""
    frmData.RText1.Text = ""
    frmData.RText2.Text = ""
    frmData.RText0.Text = "No." & Chr(10)
    For I = 0 To 28
        PostMessage frmData.RText0.hwnd, WM_KEYDOWN, VK_NEXT, 0
        frmData.RText0.SelStart = Len(frmData.RText0.Text)
        frmData.RText0.SelText = Format(I, "00") & Chr(10)
    Next
    frmData.RText1 = "正向" & IIf(RANGe = 0.1, "(mV)", "(V)") & Chr(10)
    frmData.RText2 = "负向" & IIf(RANGe = 0.1, "(mV)", "(V)") & Chr(10)

End Sub

Private Sub Timer1_Timer()
    Me.Left = FrmMain.Left + FrmMain.Width
    Me.Top = FrmMain.Top
    Timer1.Enabled = False
End Sub
