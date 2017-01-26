VERSION 5.00
Begin VB.Form frmChoose 
   Caption         =   "请选择"
   ClientHeight    =   3375
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3255
   LinkTopic       =   "Form1"
   ScaleHeight     =   3375
   ScaleWidth      =   3255
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   495
      Left            =   1680
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2400
      Left            =   840
      ScaleHeight     =   2370
      ScaleWidth      =   465
      TabIndex        =   7
      Top             =   480
      Width           =   495
      Begin VB.Shape Shape1 
         BorderColor     =   &H00C000C0&
         FillColor       =   &H00C000C0&
         FillStyle       =   0  'Solid
         Height          =   180
         Left            =   0
         Top             =   0
         Width           =   480
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   1140
      Left            =   1680
      TabIndex        =   6
      Top             =   1560
      Width           =   1380
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   180
      Index           =   4
      Left            =   240
      TabIndex        =   5
      Top             =   2520
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   180
      Index           =   3
      Left            =   240
      TabIndex        =   4
      Top             =   2040
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   180
      Index           =   2
      Left            =   240
      TabIndex        =   3
      Top             =   1560
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   180
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   180
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   540
   End
End
Attribute VB_Name = "frmChoose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub Command1_Click()
On Error GoTo errhdl
Select Case Stall
Case 2
If Shape1.Top = 800 Then
Digits = 6
WriteString Winsock1, "VOLT:DC:NPLC 1"
ElseIf Shape1.Top = 1600 Then
Digits = 7
WriteString Winsock1, "VOLT:DC:NPLC 1"
Else
MsgBox IIf(Lang, "Without any operation, the median will not be changed", "没有进行任何操作，位数将不作更改!")
End If

Case 3
If Shape1.Top = 0 Then
RANGe = 0.1
ReDim cmdstr(1 To 10)
cmdstrs = Split(GetIni("cmdstr", "RANGe120", "37,48,48,59,49,50,59,48,48,13", App.Path & "\Config.ini"), ",")
For i = LBound(cmdstr) To UBound(cmdstr)
    cmdstr(i) = Val(cmdstrs(i - 1))
Next
FrmMain.MSComm.OutPut = cmdstr

ElseIf Shape1.Top = 1100 Then
RANGe = 1
ReDim cmdstr(1 To 10)
cmdstrs = Split(GetIni("cmdstr", "RANGe1", "37,48,48,59,49,50,59,48,49,13", App.Path & "\Config.ini"), ",")
For i = LBound(cmdstr) To UBound(cmdstr)
    cmdstr(i) = Val(cmdstrs(i - 1))
Next
FrmMain.MSComm.OutPut = cmdstr
    
ElseIf Shape1.Top = 2400 - frmChoose.Shape1.Height Then
RANGe = 10
ReDim cmdstr(1 To 10)
cmdstrs = Split(GetIni("cmdstr", "RANGe30", "37,48,48,59,49,50,59,48,50,13", App.Path & "\Config.ini"), ",")
For i = LBound(cmdstr) To UBound(cmdstr)
    cmdstr(i) = Val(cmdstrs(i - 1))
Next
FrmMain.MSComm.OutPut = cmdstr

End If

If RANGe = 0.1 Or RANGe = 1 Or RANGe = 10 Or RANGe = 100 Or RANGe = 1000 Then
    If Instru = 0 Then
        WriteString Winsock1, "VOLT:DC:RANG " & RANGe
        WriteString Winsock1, "VOLT:DC:RANG?"
        RANGe = Val(ReadString(Winsock1))
    Else
        FrmMain.MSComm0.OutPut = ":VOLTage:DC:RANGe " & RANGe & vbCr
    End If
    For i = 0 To 4
        FrmMain.shpRange(i).FillColor = FrmMain.ShpErrOff.FillColor
        FrmMain.Label4(i).ForeColor = FrmMain.ShpErrOff.FillColor
        If RANGe = RangArry(i + 1) Then
            FrmMain.shpRange(i).FillColor = FrmMain.ShpErrOn.FillColor
            FrmMain.Label4(i).ForeColor = vbBlack
        End If
    Next
Select Case RANGe
Case 0.1
Filter = Filter00
Case 1
Filter = Filter01
Case 10, 100, 1000
Filter = Filter02
End Select
Cnt = 0
ss = 0
ReDim numArry(1 To Filter)
ReadAuto = True
FrmMain.Timer6.Enabled = True
End If

Case 4
If Shape1.Top = 0 Then
    Digits = 4
ElseIf Shape1.Top = 750 Then
    Digits = 5
ElseIf Shape1.Top = 1550 Then
    Digits = 6
ElseIf Shape1.Top = 2400 - frmChoose.Shape1.Height Then
    Digits = 7
End If
If FrmMain.Check1.Value = 0 Then
    FrmMain.MSComm0.OutPut = ":VOLTage:DC:DIGits " & Didits & vbCr
Else
    'If Instru = 0 Then
        'Select Case Digits
        'Case 4, 5
            'WriteString Winsock1, "VOLT:DC:NPLC 1"
        'Case 6, 7
            'WriteString Winsock1, "VOLT:DC:NPLC 2"
        'End Select
    'Else
        'frmMain.MSComm0.OutPut = ":VOLTage:DC:DIGits " & Digits & vbCr
    'End If
Select Case Digits
Case 4
    ReDim cmdstr(1 To 10)
    cmdstrs = Split(GetIni("cmdstr", "Digits5", "37,48,48,59,49,48,59,48,53,13", App.Path & "\Config.ini"), ",")
    For i = LBound(cmdstr) To UBound(cmdstr)
        cmdstr(i) = Val(cmdstrs(i - 1))
    Next
Case 5
    ReDim cmdstr(1 To 10)
    cmdstrs = Split(GetIni("cmdstr", "Digits6", "37,48,48,59,49,48,59,48,54,13", App.Path & "\Config.ini"), ",")
    For i = LBound(cmdstr) To UBound(cmdstr)
        cmdstr(i) = Val(cmdstrs(i - 1))
    Next
Case 6
    ReDim cmdstr(1 To 10)
    cmdstrs = Split(GetIni("cmdstr", "Digits7", "37,48,48,59,49,48,59,48,55,13", App.Path & "\Config.ini"), ",")
    For i = LBound(cmdstr) To UBound(cmdstr)
        cmdstr(i) = Val(cmdstrs(i - 1))
    Next
Case 7
    ReDim cmdstr(1 To 10)
    cmdstrs = Split(GetIni("cmdstr", "Digits8", "37,48,48,59,49,48,59,48,56,13", App.Path & "\Config.ini"), ",")
    For i = LBound(cmdstr) To UBound(cmdstr)
        cmdstr(i) = Val(cmdstrs(i - 1))
    Next
End Select
FrmMain.MSComm.OutPut = cmdstr
End If

Case 5
If Shape1.Top = 0 Then
RANGe = 0.1
ElseIf Shape1.Top = 550 Then
RANGe = 1
ElseIf Shape1.Top = 1100 Then
RANGe = 10
ElseIf Shape1.Top = 1650 Then
RANGe = 100
ElseIf Shape1.Top = 2400 - frmChoose.Shape1.Height Then
RANGe = 1000
End If

If Instru = 0 Then
    WriteString Winsock1, "VOLT:DC:RANG " & RANGe
    WriteString Winsock1, "VOLT:DC:RANG?"
    RANGe = Val(ReadString(Winsock1))
Else
    FrmMain.MSComm0.OutPut = ":VOLTage:DC:RANGe " & RANGe & vbCr
End If
For i = 0 To 4
    FrmMain.shpRange(i).FillColor = FrmMain.ShpErrOff.FillColor
    FrmMain.Label4(i).ForeColor = FrmMain.ShpErrOff.FillColor
    If RANGe = RangArry(i + 1) Then
        FrmMain.shpRange(i).FillColor = FrmMain.ShpErrOn.FillColor
        FrmMain.Label4(i).ForeColor = vbBlack
    End If
Next
Select Case RANGe
Case 0.1
Filter = Filter00
Case 1
Filter = Filter01
Case 10, 100, 1000
Filter = Filter02
End Select
Cnt = 0
ss = 0
ReDim numArry(1 To Filter)
End Select
ReadAuto = True
FrmMain.Timer6.Enabled = True
Stall = 0
Unload frmChoose
errhdl: If err.Number <> 0 Then MsgBox "Warning:Error" & err.Number & Chr(13) & err.Description & "," & Me.Caption
End Sub

Private Sub Form_Activate()
ReDim uVirtKey(1 To 3)
uVirtKey(1) = &H26
uVirtKey(2) = &H28
uVirtKey(3) = &H20
Dim Modifiers As Long
preWinProc = GetWindowLong(frmChoose.hwnd, GWL_WNDPROC)
SetWindowLong frmChoose.hwnd, GWL_WNDPROC, AddressOf WndProc
For i = 1 To 3
RegisterHotKey frmChoose.hwnd, 20 + i, Modifiers, uVirtKey(i)
Next i
End Sub

Private Sub Form_Load()

Me.Caption = IIf(Lang, "Please choose", "请选择")
If KeyChoose = 1 Then
Label2.Caption = IIf(Lang, "Use the arrow keys to select the digits", "请用方向键选择位数")
Select Case FrmMain.Check1.Value
Case 0
    If Instru = 0 Then
        Label1(1).Visible = True
        Label1(3).Visible = True
        Label1(1).Caption = "5.5"
        Label1(3).Caption = "6.5"
        Stall = 2
        Label1(1).Caption = Picture1.Top + 800
        Label1(3).Caption = Picture1.Top + 1600
    Else
        For i = 0 To 3
            Label1(i).Top = Picture1.Top + i * 745
            Label1(i).Visible = True
            Label1(i).Caption = CStr(i + 4)
        Next
        Stall = 4
    End If
Case 1
    For i = 0 To 3
        Label1(i).Visible = True
        Label1(i).Top = Picture1.Top + i * 745
        Label1(i).Caption = CStr(i + 5)
    Next
    Stall = 4
End Select


ElseIf KeyChoose = 2 Then

Label2.Caption = IIf(Lang, "Use the arrow keys to select the Range", "请用方向键选择量程")
Select Case FrmMain.Check1.Value
Case 0
    For i = 0 To 4
        Label1(i).Visible = True
        Label1(i).Top = Picture1.Top + i * 550
        If i <> 0 Then Label1(i).Caption = CStr(10 ^ (i - 1)) & "V"
    Next
    Label1(0).Caption = "100mV"
    Stall = 5
Case 1
    For i = 1 To 3
        Label1(i).Visible = True
    Next
    Label1(1).Caption = "120mV"
    Label1(2).Caption = "1V"
    Label1(3).Caption = "30V"
    For i = 1 To 3
    Label1(i).Top = Picture1.Top + (i - 1) * 1100
    Next
    Stall = 3
End Select

Else
Unload frmChoose
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
SetWindowLong frmChoose.hwnd, GWL_WNDPROC, preWinProc
For i = 1 To 3
UnregisterHotKey frmChoose.hwnd, uVirtKey(i)
Next i
End Sub
