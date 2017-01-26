VERSION 5.00
Object = "{CFA7AFF4-3242-4269-9172-7389D695AE01}#1.0#0"; "StoneXP.ocx"
Begin VB.Form frm_Main 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "GPIB调试器"
   ClientHeight    =   5145
   ClientLeft      =   2490
   ClientTop       =   1395
   ClientWidth     =   5205
   Icon            =   "frm_Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "frm_Main.frx":000C
   ScaleHeight     =   5145
   ScaleWidth      =   5205
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "调试"
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   3135
      Left            =   120
      TabIndex        =   15
      Top             =   1800
      Width           =   4935
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1320
         TabIndex        =   5
         Text            =   "*IDN?"
         Top             =   360
         Width           =   3495
      End
      Begin StoneXP.XPButton XPButton6 
         Height          =   375
         Left            =   3720
         TabIndex        =   9
         Top             =   720
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "清除(&C)"
         MouseIcon       =   "frm_Main.frx":034E
         MousePointer    =   99
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1815
         Left            =   120
         ScaleHeight     =   1785
         ScaleWidth      =   4665
         TabIndex        =   16
         Top             =   1200
         Width           =   4695
      End
      Begin StoneXP.XPButton XPButton5 
         Height          =   375
         Left            =   2280
         TabIndex        =   8
         Top             =   720
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "发送/读取(&F)"
         MouseIcon       =   "frm_Main.frx":0668
         MousePointer    =   99
      End
      Begin StoneXP.XPButton XPButton4 
         Height          =   375
         Left            =   1200
         TabIndex        =   7
         Top             =   720
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "读取(&R)"
         MouseIcon       =   "frm_Main.frx":0982
         MousePointer    =   99
      End
      Begin StoneXP.XPButton XPButton3 
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "发送(&S)"
         MouseIcon       =   "frm_Main.frx":0C9C
         MousePointer    =   99
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "发送命令:"
         BeginProperty Font 
            Name            =   "隶书"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   1080
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "仪器控制"
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   1575
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   4935
      Begin VB.ComboBox Cb_GPIB_Type 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox txt_GPIBAddr 
         Height          =   285
         Left            =   1440
         TabIndex        =   1
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox txt_remoteHost 
         Height          =   285
         Left            =   1440
         TabIndex        =   2
         Top             =   1080
         Width           =   1935
      End
      Begin StoneXP.XPButton XPButton8 
         Height          =   375
         Left            =   3480
         TabIndex        =   4
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "断开连接(&D)"
         MouseIcon       =   "frm_Main.frx":0FB6
         MousePointer    =   99
      End
      Begin StoneXP.XPButton XPButton1 
         Height          =   375
         Left            =   3480
         TabIndex        =   3
         Top             =   480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "连接仪器(&R)"
         MouseIcon       =   "frm_Main.frx":12D0
         MousePointer    =   99
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "连接方式:"
         BeginProperty Font 
            Name            =   "隶书"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   1080
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "GPIB地址:"
         BeginProperty Font 
            Name            =   "隶书"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   240
         TabIndex        =   12
         Top             =   720
         Width           =   1080
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "网络地址:"
         BeginProperty Font 
            Name            =   "隶书"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   5
         Left            =   240
         TabIndex        =   11
         Top             =   1080
         Width           =   1080
      End
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Height          =   195
      Left            =   8280
      TabIndex        =   14
      Top             =   7080
      Width           =   3060
   End
End
Attribute VB_Name = "frm_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Filename - Simple.frm
'
' This application demonstrates how to read from and write to the
' Tektronix PS2520G Programmable Power Supply using GPIB.
' Download by http://www.codefans.net
' This sample application is comprised of three basic parts:
'
' 1. Initialization
' 2. Main Body
' 3. Cleanup
'
' The Initialization portion consists of getting a handle to a
' device and then clearing the device.
'
' In the Main Body, this application queries a device for its
' identification code by issuing the '*IDN?' command. Many
' instruments respond to this command with an identification string.
' Note, 488.2 compliant devices are required to respond to this
' command.
'
' The last step, Cleanup, takes the device offline.
' Download by http://www.codefans.net
    Option Explicit
    

    
    Const BDINDEX = 0                   ' Board Index
    Dim PRIMARY_ADDR_OF_PPS       ' Primary address of device
    Const NO_SECONDARY_ADDR = 0         ' Secondary address of device
    Const timeOUT = T10s                ' Timeout value = 10 seconds
    Const EOTMODE = 1                   ' Enable the END message
    Const EOSMODE = 0                   ' Disable the EOS mode
    
    Const ARRAYSIZE = 1024              ' Size of read buffer
    Private Type CMD_str
        cmd(20) As String
    End Type
    Dim cmd As CMD_str
    Dim ErrorMnemonic
    Dim ErrMsg As String * 100
    Dim Dev As Integer
    Dim ValueStr As String * ARRAYSIZE
    
    Public iniFile As String
    
    Dim actual As Long

    Dim isConnect As Boolean
    Dim isFailed As Boolean
    Dim remoteHost As String
    Dim Current_X As Integer
    Dim Current_Y As Integer
    Dim icx As Integer
    Dim I As Integer
    
Private Sub GPIBCleanup(msg$)
On Error Resume Next
    ' After each GPIB call, the application checks whether the call
    ' succeeded. If an NI-488.2 call fails, the GPIB driver sets the
    ' corresponding bit in the global status variable. If the call
    ' failed, this procedure prints an error message, takes the device
    ' offline and exits.

    ErrorMnemonic = Array("EDVR", "ECIC", "ENOL", "EADR", "EARG", _
                          "ESAC", "EABO", "ENEB", "EDMA", "", _
                          "EOIP", "ECAP", "EFSO", "", "EBUS", _
                          "ESTB", "ESRQ", "", "", "", "ETAB")
    ErrMsg$ = msg$ & Chr(13) & "ibsta = &H" & Hex(ibsta) & Chr(13) _
              & "iberr = " & iberr & " <" & ErrorMnemonic(iberr) & ">"
    MsgBox ErrMsg$, vbCritical, "Error"

    ilonl Dev%, 0
    isConnect = False
    'Call RSTGPIB
    
End Sub

Sub GpibSend(ByVal CH As String)

    If isConnect = False Then Exit Sub
    
    Select Case Cb_GPIB_Type.Text
        Case "NI_GPIB"
            ilwrt Dev%, CH, Len(CH)
            If (ibsta And EERR) Then
                Call GPIBCleanup("Unable to request data from Power Supply")
            End If
        Case Else
            iwrite Dev%, CH + Chr$(10), Len(CH), 1, 0&
    End Select
End Sub

Function GPIBRead(ByVal I As Integer, Optional lyreCH As String) As String
    On Error Resume Next
    Dim CH As String
    
    If isConnect = False Then Exit Function
    
    CH = Space(I)
    Select Case Cb_GPIB_Type.Text
        Case "NI_GPIB"
            
    
            Call ibrd(Dev, CH)
    
            If (ibsta And EERR) Then
                Call GPIBCleanup("Unable to read from device")
            End If
        Case Else
            iread Dev%, CH, 2000, 0, actual
            'iread(id, strres, 80, 0&, actual)
    End Select
    
    CH = Trim(CH)
    
    GPIBRead = CH
End Function

Private Sub Cb_GPIB_Type_Click()
    On Error Resume Next
    Select Case Cb_GPIB_Type.Text
        Case "NI_GPIB", "Agilent_GPIB"
            txt_remoteHost.Enabled = False
 
        Case Else
            txt_remoteHost.Enabled = True

    End Select
    txt_GPIBAddr.SetFocus
    
End Sub











Private Sub Form_Load()
    On Error Resume Next

    Dim ss As String
    
    With Cb_GPIB_Type
        .AddItem "NI_GPIB"
        .AddItem "Agilent_LAN"
        .AddItem "Agilent_GPIB"
        .Text = "Agilent_LAN"
    End With
    
    With Combo1
        .AddItem "*IDN?"
        .AddItem "*CLS"
        .Text = "*IDN?"
    End With
    
    Me.Caption = " " & App.Title & "_V" & App.Major & "." & App.Minor & "." & App.Revision

    iniFile = IIf(Right(App.Path, 1) = "\", App.Path, App.Path & "\") & "11713A.ini"
End Sub

Sub RSTGPIB(ByVal lyIDN As Boolean)
    On Error GoTo ErrorHandler
    
    Dim CH As String
    
    Picture1.ForeColor = vbBlue
    Select Case Cb_GPIB_Type.Text
        Case "NI_GPIB"
        
            Dev% = ildev(BDINDEX, PRIMARY_ADDR_OF_PPS, NO_SECONDARY_ADDR, _
                         timeOUT, EOTMODE, EOSMODE)
            If (ibsta And EERR) Then
                ErrMsg = "Unable to open device" & Chr(13) & "ibsta = &H" _
                          & Hex(ibsta) & Chr(13) & "iberr = " & iberr
                MsgBox ErrMsg, vbCritical, "Error"
                Exit Sub
            End If
            'Debug.Print Dev
            ilclr Dev%
            If (ibsta And EERR) Then
                Call GPIBCleanup("Unable to clear device")
                isConnect = False
                Exit Sub
            Else
                isConnect = True
            End If
            
            If lyIDN = True Then
                GpibSend "*IDN?"
                CH = Space(2000)
                Call ibrd(Dev, CH)
                CH = Trim(CH)
                If CH = "" Then
                    isConnect = False
                    Exit Sub
                End If
                MsgBox "与仪器通信成功!" & vbCrLf & CH, vbInformation + vbOKOnly
            End If
        Case "Agilent_GPIB"
            Dim dvm As Integer
            Dev% = iopen("gpib0," & PRIMARY_ADDR_OF_PPS)
            'Debug.Print Dev
            If Dev = 0 Then
                isConnect = False
                Exit Sub
            Else
                isConnect = True
            End If
            Call itimeout(Dev, 10000)
            If lyIDN = True Then
                GpibSend "*IDN?"
                CH = Space(2000)
                Call iread(Dev%, CH, 2000, 0, actual)
                CH = Trim(CH)
                MsgBox "与仪器通信成功!" & vbCrLf & CH, vbInformation + vbOKOnly
            End If
        Case Else
            'Dim dvm As Integer
            Dev% = iopen("lan[" & remoteHost & "]:hpib9," & PRIMARY_ADDR_OF_PPS)
            'Debug.Print Dev
            If Dev = 0 Then
                isConnect = False
                Exit Sub
            Else
                isConnect = True
            End If
            Call itimeout(Dev, 10000)
            If lyIDN = True Then
                GpibSend "*IDN?"
                CH = Space(2000)
                Call iread(Dev%, CH, 2000, 0, actual)
                CH = Trim(CH)
                MsgBox "与仪器通信成功!" & vbCrLf & CH, vbInformation + vbOKOnly
            End If
        End Select
        Picture1.Print CH
        icx = 1
        'Me.Caption = "网络分析仪控制器-" & CH
        Exit Sub
        
ErrorHandler:

    MsgBox "Unable to request data from Equipment! " & vbCrLf & "Err Number=" & Err.Number & vbCrLf & Err.Description, vbCritical + vbOKOnly, "Error"

    If Dev <> 0 Then
        'Call iclose(Dev)
    End If

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    If isConnect = True Then
        iclose Dev
        ibonl32 Dev, 0
    End If

    'End
    
End Sub




Private Sub txt_GPIBAddr_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 13 Then
       If txt_remoteHost.Enabled = True Then
            txt_remoteHost.SetFocus
        Else
            XPButton1.SetFocus
        End If
    End If
End Sub

Private Sub txt_remoteHost_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then XPButton1.SetFocus
End Sub

Private Sub XPButton1_Click()
        
    Select Case Cb_GPIB_Type.Text
        Case "NI_GPIB"
            PRIMARY_ADDR_OF_PPS = Val(txt_GPIBAddr.Text)
            
            If PRIMARY_ADDR_OF_PPS = 0 Then Exit Sub

            

        Case "Agilent_LAN"
            PRIMARY_ADDR_OF_PPS = Val(txt_GPIBAddr.Text)
            
            If PRIMARY_ADDR_OF_PPS = 0 Then Exit Sub
            
            
            remoteHost = txt_remoteHost.Text
            
            If remoteHost = "" Then Exit Sub
        Case "Agilent_GPIB"
            PRIMARY_ADDR_OF_PPS = Val(txt_GPIBAddr.Text)
            
            If PRIMARY_ADDR_OF_PPS = 0 Then Exit Sub
    End Select
    
    Call RSTGPIB(False)

End Sub


Private Sub XPButton3_Click()
    If isConnect = True Then
        GpibSend Combo1.Text
        For I = 0 To Combo1.ListCount - 1
            If Combo1.list(I) = Combo1.Text Then Exit Sub
        Next
        Combo1.AddItem Combo1.Text
    End If
End Sub

Private Sub XPButton4_Click()
    Dim CH As String
    If isConnect = True Then
        If icx = 8 Then
            Picture1.Cls
            icx = 0
        End If
        Current_Y = icx * 200
        icx = icx + 1
        If icx Mod 2 = 1 Then
            Picture1.ForeColor = vbBlue
        Else
            Picture1.ForeColor = vbRed
        End If
        CH = Space(2000)
        CH = GPIBRead(2000)
        Picture1.CurrentY = Current_Y
        Picture1.Print CH
        
    End If
End Sub

Private Sub XPButton5_Click()
    Dim CH As String
    If isConnect = True Then
        GpibSend Combo1.Text
        For I = 0 To Combo1.ListCount - 1
            If Combo1.list(I) = Combo1.Text Then GoTo jump
        
        Next
        Combo1.AddItem Combo1.Text
jump:
        
        If icx = 8 Then
            Picture1.Cls
            icx = 0
        End If
        Current_Y = icx * 200
        icx = icx + 1
        If icx Mod 2 = 1 Then
            Picture1.ForeColor = vbBlue
        Else
            Picture1.ForeColor = vbRed
        End If
        Picture1.CurrentY = Current_Y
        CH = Space(2000)
        CH = GPIBRead(2000)
        Picture1.Print CH
        
    End If
End Sub

Private Sub XPButton6_Click()
    Picture1.Cls
    icx = 0
End Sub

Private Sub XPButton8_Click()
    On Error GoTo errJump
    
    If isConnect = False Then Exit Sub
    Select Case Cb_GPIB_Type.Text
        Case "NI_GPIB"
            ilonl Dev%, 0
        Case Else
            Call iclose(Dev%)
    End Select
    isConnect = False
    Picture1.Cls
    icx = 0
    MsgBox "已与设备断开连接！", vbInformation + vbOKOnly
    
    Exit Sub
errJump:

    MsgBox Err.Number & vbCrLf & Err.Description

End Sub


Function FormatStr(ByVal lyCH As String, ByVal lyLong As Integer) As String
    Dim lybck As String
    If InStr(1, lyCH, ">") <> 0 Then
        FormatStr = Right(lyCH, Len(lyCH) - InStr(1, lyCH, ">"))
        lybck = Left(lyCH, InStr(1, lyCH, ">"))
        FormatStr = Format(FormatStr, "0.000000000000000000000000000000")
        FormatStr = lybck & Left(FormatStr, lyLong)
    Else
        FormatStr = Right(lyCH, Len(lyCH) - InStr(1, lyCH, "<"))
        lybck = Left(lyCH, InStr(1, lyCH, "<"))
        FormatStr = Format(FormatStr, "0.000000000000000000000000000000")
        FormatStr = lybck & Left(FormatStr, lyLong)
    End If
        
End Function


