Attribute VB_Name = "SysLan"
Option Explicit

Public LogoLan As Integer

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function GetSystemDefaultLCID Lib "kernel32" () As Long

Public Type OSVERSIONINFO
  dwOSVersionInfoSize As Long
  dwMajorVersion As Long
  dwMinorVersion As Long
  dwBuildNumber As Long
  dwPlatformId As Long
  szCSDVersion As String * 128
  wServicePackMinor As Integer
  wProductType As Byte
  OsName As String
  OsLanguage As String
End Type

Public Function Getsyslan() As String
  Dim Ver As OSVERSIONINFO
  Ver.dwOSVersionInfoSize = 148
  GetVersionEx Ver
  With Ver
      Dim LocaleID As Long
      LocaleID = GetSystemDefaultLCID
      Select Case LocaleID
          Case &H404
              .OsLanguage = "繁体中文"
          Case &H804
              .OsLanguage = "简体中文"
          Case &H409
              .OsLanguage = "英文"
          Case Else
              .OsLanguage = "其他"
      End Select
  End With
  Getsyslan = Ver.OsLanguage
End Function


