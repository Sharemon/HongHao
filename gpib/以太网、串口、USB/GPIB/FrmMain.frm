VERSION 5.00
Begin VB.Form GPIB 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "GPIB"
   ClientHeight    =   7236
   ClientLeft      =   4032
   ClientTop       =   4416
   ClientWidth     =   6636
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7236
   ScaleWidth      =   6636
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdreceive 
      Caption         =   "����"
      Height          =   735
      Left            =   2460
      TabIndex        =   10
      Top             =   6360
      Width           =   1695
   End
   Begin VB.CommandButton cmdclc 
      Caption         =   "���"
      Height          =   735
      Left            =   4800
      TabIndex        =   9
      Top             =   6360
      Width           =   1695
   End
   Begin VB.CommandButton cmdsend 
      Caption         =   "����"
      Height          =   735
      Left            =   120
      TabIndex        =   8
      Top             =   6360
      Width           =   1695
   End
   Begin VB.Frame Frame2 
      Caption         =   "������"
      Height          =   3255
      Left            =   120
      TabIndex        =   6
      Top             =   3000
      Width           =   6375
      Begin VB.TextBox txtreceive 
         Height          =   2775
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   360
         Width           =   6135
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "������"
      Height          =   1095
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   6375
      Begin VB.ComboBox Combosend 
         Height          =   300
         ItemData        =   "FrmMain.frx":08CA
         Left            =   120
         List            =   "FrmMain.frx":08E0
         TabIndex        =   5
         Text            =   "*IDN?"
         Top             =   480
         Width           =   6135
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "����"
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6375
      Begin VB.TextBox txtaddr 
         BeginProperty Font 
            Name            =   "����"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1080
         TabIndex        =   2
         Text            =   "16"
         Top             =   510
         Width           =   1452
      End
      Begin VB.CommandButton cmdlink 
         Caption         =   "����"
         Height          =   735
         Left            =   3960
         TabIndex        =   1
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "�豸��ַ��"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   975
      End
      Begin VB.Shape islink 
         BorderColor     =   &H80000016&
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   3240
         Shape           =   3  'Circle
         Top             =   540
         Width           =   375
      End
   End
End
Attribute VB_Name = "GPIB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'GPIB��ַ�����豸������->CONTEC Devices->Common Setting->Diagnose��ȡ

Option Explicit

Const BDINDEX = 0
Const NO_SECONDARY_ADDR = 0
Const TIMEOUT = T10s
Const EOTMODE = 1
Const EOSMODE = 0

Const ARRAYSIZE = 100               ' ����ռ�

Dim ResByte As Integer
Dim Dev As Integer
Dim Valuestr As String * ARRAYSIZE

Dim ErrMsg As String * 100
Dim ErrorMnemonic

Dim Receivable As Boolean
Dim LinkorNot As Boolean

Private Sub cmdclc_Click()
    txtreceive = ""
End Sub

Private Sub cmdlink_Click()
On Error GoTo HadErr
    Dim pad As Integer
    pad = txtaddr
    Call ibdev(BDINDEX, pad%, NO_SECONDARY_ADDR, TIMEOUT, EOTMODE, EOSMODE, Dev%)       '�豸���Ӻ���
    Call ibwrt(Dev%, "*IDN?")               '�Է������ݣ����޴�������ʾ������
    If (ibsta And EERR) Then
        islink.FillColor = vbRed
        MsgBox "�����豸��ַ��", vbExclamation, "Error"
    Else
        islink.FillColor = vbGreen
        LinkorNot = True
        Call ibclr(Dev%)                    '���Է��͵ķ���ֵ���
    End If
HadErr: Exit Sub
End Sub

Private Sub cmdreceive_Click()
    If LinkorNot Then
        If Receivable Then
            If Right(Combosend.Text, 1) = "?" Then          '�жϷ��͵��Ƿ�Ϊ���ʾ䣬�������򲻵��ö�ȡ����
                Call ibrd(Dev%, Valuestr)                   '��ȡ����ֵ����
                txtreceive = txtreceive & Mid(Valuestr, 1, ibcntl) & vbCrLf
            Else
                txtreceive = txtreceive & "NULL" & vbCrLf
            End If
            Receivable = False
        End If
    Else
        MsgBox "�������ӣ�", vbExclamation, "Error"
    End If
End Sub

Private Sub cmdsend_Click()
Combosend.Text = Trim(Combosend.Text)
    If LinkorNot Then
        If Combosend.Text <> "" Then
            Call ibwrt(Dev%, Combosend)                     '���ͺ���
            Receivable = True
        End If
    Else
        MsgBox "�������ӣ�", vbExclamation, "Error"
    End If
End Sub

Private Sub Form_Load()
    Receivable = False
    LinkorNot = False
End Sub

Private Sub txtreceive_Change()
txtreceive.SelStart = Len(txtreceive)
End Sub
