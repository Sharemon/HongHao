Attribute VB_Name = "modTray"
Public Lan As Integer
Public ISDataShow As Boolean '用于表示FrmDataShow是否显示
Public DataOfHex As String
Public DataOfBin As String
Public DataOfASCII As String
Public DatasSaveStyle As Integer '1、2、3三种情况，用于判断数据是否保存、单个保存、连续保存
Public RecordNumber As Long '记录数据条数
Public OpenFileName As String '储存路径
Public BinOrDeng As Boolean
Public ReceiveCounts As Long
Public ReceiveTrueCounts As Long

Option Explicit
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
    (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" _
    (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long

Public Const TRAY_CALLBACK = (&H400 + 1001&)
Public Const GWL_WNDPROC = -4

Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Const WM_RBUTTONDBLCLK = &H206

Public Enum WIN_STATUS
    STA_MIN
    STA_NORMAL
End Enum

Public glWinRet As Long
Public OrgWinRet As Long
Public Status As WIN_STATUS '保存窗体状态


Public Function CallbackMsgs(ByVal wHwnd As Long, ByVal wMsg As Long, ByVal wp_id As Long, ByVal lp_id As Long) As Long
On Error Resume Next
    If wMsg = TRAY_CALLBACK Then
        With FrmMain
            Select Case CLng(lp_id)
                Case WM_RBUTTONUP '右键
                    .PopupMenu .TrayMenu, , , , .MenuShow
                Case WM_LBUTTONUP '左键
                    If .WindowState = vbMinimized Then
                        Status = STA_NORMAL
                        .Visible = True
                        .WindowState = vbNormal
                    Else
                        Status = STA_MIN
                        .WindowState = vbMinimized
                        .Visible = False
                    End If
            End Select
        End With
    End If
    CallbackMsgs = CallWindowProc(glWinRet, wHwnd, wMsg, wp_id, lp_id)
End Function



Public Function DEC_to_BIN(Dec As Long, MinimumDigits As Integer) As String '''''''十进制转二进制
Dim ExtraDigitsNeeded As Integer
DEC_to_BIN = ""
Do While Dec > 0
DEC_to_BIN = Dec Mod 2 & DEC_to_BIN
Dec = Dec \ 2
Loop
ExtraDigitsNeeded = MinimumDigits - Len(DEC_to_BIN)
If ExtraDigitsNeeded > 0 Then
    DEC_to_BIN = String(ExtraDigitsNeeded, "0") & DEC_to_BIN
End If
End Function

