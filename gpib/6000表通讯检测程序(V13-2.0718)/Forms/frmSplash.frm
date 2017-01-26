VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6465
   ClientLeft      =   6750
   ClientTop       =   2880
   ClientWidth     =   6060
   ClipControls    =   0   'False
   DrawMode        =   14  'Copy Pen
   FillStyle       =   3  'Vertical Line
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":0442
   ScaleHeight     =   6465
   ScaleWidth      =   6060
   StartUpPosition =   1  '所有者中心
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   5640
      Top             =   0
   End
   Begin VB.CommandButton CmdCh 
      Caption         =   "中 文"
      Height          =   372
      Left            =   1320
      TabIndex        =   2
      Top             =   6000
      Width           =   852
   End
   Begin VB.CommandButton CmdEx 
      Caption         =   "退 出"
      Height          =   372
      Left            =   3720
      TabIndex        =   1
      Top             =   6000
      Width           =   852
   End
   Begin VB.CommandButton CmdEn 
      Caption         =   "English"
      Height          =   372
      Left            =   2520
      TabIndex        =   0
      Top             =   6000
      Width           =   852
   End
   Begin VB.Image Image1 
      Height          =   6480
      Left            =   0
      Picture         =   "frmSplash.frx":11CA0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6105
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdCh_Click()
LogoLan = 0
FrmMain.Show
Unload Me
End Sub

Private Sub CmdEn_Click()
LogoLan = 1
FrmMain.Show
Unload Me
End Sub

Private Sub CmdEx_Click()
Unload Me
End Sub

Private Sub Form_Activate()
SetFormToAlpha Me.hWnd, 220
End Sub

Private Sub Form_Load()
On Error Resume Next
CmdCh.Top = frmSplash.Height - 930
CmdEn.Top = frmSplash.Height - 930
CmdEx.Top = frmSplash.Height - 930
CmdCh.Left = (frmSplash.Width - (CmdCh.Width * 3) - 800) / 2
CmdEn.Left = CmdCh.Left + CmdCh.Width + 400
CmdEx.Left = CmdCh.Left + (CmdCh.Width * 2) + 800
    SkinH_Attach
    SkinH_AttachEx "china.she", ""
End Sub

Private Sub Timer1_Timer()
Select Case Getsyslan
    Case "简体中文"
        LogoLan = 0
        FrmMain.Show
        Unload Me
    Case Else
        FrmMain.CmdMode.Caption = "Detection Mode"
        LogoLan = 1
        FrmMain.Show
        Unload Me
End Select
End Sub
