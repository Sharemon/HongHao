VERSION 5.00
Object = "{1C98F15C-068A-11D4-98C2-00108301CB39}#2.0#0"; "AGT3494A.OCX"
Begin VB.Form frmLearnString 
   Caption         =   "Form1"
   ClientHeight    =   2730
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5220
   LinkTopic       =   "Form1"
   ScaleHeight     =   2730
   ScaleWidth      =   5220
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtModel 
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   1680
      Width           =   2535
   End
   Begin Agt3494ALib.Agt3494A Agt3494A1 
      Left            =   2640
      Top             =   120
      _ExtentX        =   953
      _ExtentY        =   847
      Address         =   "GPIB::22"
   End
   Begin VB.TextBox txtAddress 
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   600
      Width           =   2655
   End
   Begin VB.CommandButton cmdSetIO 
      Caption         =   "Set I/O"
      Height          =   495
      Left            =   3360
      TabIndex        =   3
      Top             =   480
      Width           =   1575
   End
   Begin VB.CommandButton cmdSendSettings 
      Caption         =   "Send Settings"
      Height          =   495
      Left            =   3360
      TabIndex        =   1
      Top             =   2040
      Width           =   1575
   End
   Begin VB.CommandButton cmdGetSettings 
      Caption         =   "Get Settings"
      Height          =   495
      Left            =   3360
      TabIndex        =   0
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label label2 
      Caption         =   "Instrument"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   1320
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Address"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "frmLearnString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' """""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
'    Copyright © 2000 Agilent Technologies Inc. All rights
'    reserved.
'
' You have a royalty-free right to use, modify, reproduce and distribute
' the Sample Application Files (and/or any modified version) in any way
' you find useful, provided that you agree that Agilent has no
' warranty,  obligations or liability for any Sample Application Files.
'
' Agilent Technologies provides programming examples for illustration only,
' This sample program assumes that you are familiar with the programming
' language being demonstrated and the tools used to create and debug
' procedures. Agilent support engineers can help explain the
' functionality of Agilent software components and associated
' commands, but they will not modify these samples to provide added
' functionality or construct procedures to meet your specific needs.
' """"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
''' -------------------------------------------------------------------------
''' Project Name: learnString
'''
''' Description:   Demonstrates how to query the instrument for its
'''                settings. From the instrument responses this program
'''                will build a command string that can be sent to
'''                the instrument to set the instrument to the previous
'''                settings.
'''
'''                Copyright  ©  2000 Agilent Technologies, Inc.
'''
''' Date            Developer
''' May 12, 2000    Agilent Technologies
''' -------------------------------------------------------------------------

Dim learnString As String

Private Sub cmdGetSettings_Click()
    ' retrieve the 34401A learn string from instrument
    ' store in string and write to debug window.
    Dim id() As String
    
    On Error GoTo GetSettingsError
    
    EnableButtons False
    
    Agt3494A1.Address = txtAddress.Text
    
    ' Gets the instrument model number
    Agt3494A1.Output "*idn?"
    Agt3494A1.Enter id
    txtModel.Text = id(1)
    txtModel.Refresh
    
    If InStr(1, id(1), "34420") Then
        learnString = Get34420ALearnString(Agt3494A1)
    Else
        learnString = Get34401ALearnString(Agt3494A1)
    End If
    
    EnableButtons True
    Exit Sub
    
GetSettingsError:
    MsgBox "GetSettingsError, " & Err.Description
    EnableButtons True
End Sub

Private Sub cmdSendSettings_Click()
    ' if the learnString is not empty, the string will
    ' be send to the instrument.
    
    EnableButtons False
    
    If Len(learnString) > 10 Then
        ' allow 30 seconds for RS232, because it is slow
        Agt3494A1.Timeout = 30000
        Agt3494A1.Output learnString
        Agt3494A1.Timeout = 10000
    Else
        MsgBox "Learn string empty"
        Debug.Print learnString
    End If

    EnableButtons True
End Sub

Private Sub cmdSetIO_Click()
    ' display the find address dialog
    ' and update the address text box when done
    EnableButtons False
    
    With Agt3494A1
        .Address = txtAddress
        .ShowConnectDialog
       txtAddress.Text = .Address
    End With
    
    EnableButtons True
End Sub

Private Sub Form_Load()
    ' set the address from the 3494A control
    txtAddress.Text = Agt3494A1.Address
    
End Sub

Sub EnableButtons(ByVal enable As Boolean)

    If enable Then
        cmdGetSettings.Enabled = True
        cmdSendSettings.Enabled = True
        cmdSetIO.Enabled = True
    Else
        cmdGetSettings.Enabled = False
        cmdSendSettings.Enabled = False
        cmdSetIO.Enabled = False
    End If
End Sub
