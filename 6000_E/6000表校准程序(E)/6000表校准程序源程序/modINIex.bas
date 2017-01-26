Attribute VB_Name = "modIni"
Option Explicit
'读出自定义INI文件
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
'写入自定义INI文件
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
'读出自定义INI文件中的单个区段间的所有键名和值
Public Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
'读出自定义INI所有区段名
Public Declare Function GetPrivateProfileSectionNames Lib "kernel32.dll" Alias "GetPrivateProfileSectionNamesA" (ByVal lpszReturnBuffer As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Public Function GetIni(ByVal SectionName As String, ByVal KeyWord As String, ByVal DefString As String, ByVal FileName As String) As String
    Dim ResultString As String * 255
    If GetPrivateProfileString(ByVal SectionName, ByVal KeyWord, vbNullString, ByVal ResultString, ByVal Len(ResultString), ByVal FileName) > 0 Then '关键词的值不为空
        GetIni = Left(ResultString, InStr(ResultString, Chr(0)) - 1)
    Else    '将缺省值写入INI文件
        WritePrivateProfileString SectionName, KeyWord, DefString, FileName
        GetIni = DefString
    End If
End Function

Public Function GetKeyWord(ByVal SectionName As String, ByVal DefString As String, ByVal FileName As String) As String
    Dim szBuf As String * 255
    If GetPrivateProfileSection(ByVal SectionName, ByVal szBuf, Len(szBuf), ByVal FileName) > 0 Then
        '同时获取键名和值
        GetKeyWord = Left(szBuf, InStr(szBuf, Chr(0)) - 1)
    Else
        WritePrivateProfileString SectionName, DefString, vbNullString, FileName
        GetKeyWord = DefString
    End If
End Function

Public Function GetKey(ByVal SectionName As String, ByVal DefString As String, ByVal FileName As String) As String
    Dim szBuf As String * 255, ResultString As String
    ResultString = GetKeyWord(ByVal SectionName, ByVal szBuf, ByVal FileName)
    If InStr(ResultString, "=") <> 0 Then
        GetKey = Left(ResultString, InStr(ResultString, "=") - 1)
    Else
        GetKey = DefString
    End If
End Function

Public Sub WriteINI(ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String)
WritePrivateProfileString lpApplicationName, lpKeyName, lpString, lpFileName
End Sub
