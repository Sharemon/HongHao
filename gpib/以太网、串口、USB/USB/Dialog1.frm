VERSION 5.00
Begin VB.Form GPIB 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "GPIB������"
   ClientHeight    =   4044
   ClientLeft      =   4032
   ClientTop       =   4416
   ClientWidth     =   6228
   Icon            =   "Dialog1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4044
   ScaleWidth      =   6228
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdclc 
      Caption         =   "���"
      Height          =   372
      Left            =   5040
      TabIndex        =   9
      Top             =   3480
      Width           =   972
   End
   Begin VB.ComboBox Combosend 
      Height          =   288
      ItemData        =   "Dialog1.frx":16C02
      Left            =   1080
      List            =   "Dialog1.frx":16C18
      TabIndex        =   8
      Top             =   960
      Width           =   4932
   End
   Begin VB.CommandButton cmdreceive 
      Caption         =   "����"
      Height          =   372
      Left            =   3720
      TabIndex        =   7
      Top             =   3480
      Width           =   1092
   End
   Begin VB.CommandButton cmdsend 
      Caption         =   "����"
      Height          =   372
      Left            =   2400
      TabIndex        =   6
      Top             =   3480
      Width           =   1092
   End
   Begin VB.TextBox txtreceive 
      Height          =   1692
      Left            =   1080
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   1560
      Width           =   4932
   End
   Begin VB.CommandButton cmdlink 
      Caption         =   "����"
      Height          =   372
      Left            =   1080
      TabIndex        =   2
      Top             =   3480
      Width           =   1092
   End
   Begin VB.TextBox txtaddr 
      Height          =   372
      Left            =   1080
      TabIndex        =   1
      Text            =   "16"
      Top             =   360
      Width           =   1452
   End
   Begin VB.Shape islink 
      BorderColor     =   &H80000016&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   372
      Left            =   360
      Shape           =   3  'Circle
      Top             =   3480
      Width           =   372
   End
   Begin VB.Label Label3 
      Caption         =   "���գ�"
      Height          =   492
      Left            =   360
      TabIndex        =   5
      Top             =   1560
      Width           =   492
   End
   Begin VB.Label Label2 
      Caption         =   "���ͣ�"
      Height          =   372
      Left            =   360
      TabIndex        =   3
      Top             =   960
      Width           =   492
   End
   Begin VB.Label Label1 
      Caption         =   "�豸��ַ��"
      Height          =   252
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   852
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
