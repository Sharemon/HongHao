Attribute VB_Name = "modReboot"
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Private Declare Function Process32First Lib "kernel32" (ByVal hSnapShot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapShot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal blnheritHandle As Long, ByVal dwAppProcessId As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal ApphProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Sub CloseHandle Lib "kernel32" (ByVal hPass As Long)
Private Const TH32CS_SNAPPROCESS = &H2&
Private Const WM_CLOSE = &H10

Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * 260
End Type

'关闭指定名称的进程
Private Sub KillProcess(sProcess As String)
    Dim lSnapShot As Long
    Dim lNextProcess As Long
    Dim tPE As PROCESSENTRY32
    lSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0&)
    If lSnapShot <> -1 Then
        tPE.dwSize = Len(tPE)
        lNextProcess = Process32First(lSnapShot, tPE)
        Do While lNextProcess
            If LCase$(sProcess) = LCase$(Left(tPE.szExeFile, InStr(1, tPE.szExeFile, Chr(0)) - 1)) Then
                Dim lProcess As Long
                Dim lExitCode As Long
                lProcess = OpenProcess(1, False, tPE.th32ProcessID)
                TerminateProcess lProcess, lExitCode
                CloseHandle lProcess
            End If
            lNextProcess = Process32Next(lSnapShot, tPE)
        Loop
        CloseHandle (lSnapShot)
    End If
End Sub


Sub Main()
Dim UnloadMode As Integer
UnloadMode = Val(GetIni("UnloadMode", "Unload", "-1", App.Path & "\Config.ini"))
KillProcess ("6000DDM Calitor.exe")
delay 3000
If UnloadMode = 0 Then Shell App.Path & "\6000DDM Calitor.exe", vbNormalFocus
Exit Sub
End Sub
