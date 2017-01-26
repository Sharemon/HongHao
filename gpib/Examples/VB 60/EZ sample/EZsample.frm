VERSION 5.00
Object = "{1C98F15C-068A-11D4-98C2-00108301CB39}#2.0#0"; "AGT3494A.OCX"
Begin VB.Form Form1 
   Caption         =   "Agilent 34401A Sample"
   ClientHeight    =   6435
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6255
   LinkTopic       =   "Form1"
   ScaleHeight     =   6435
   ScaleWidth      =   6255
   StartUpPosition =   3  'Windows Default
   Begin Agt3494ALib.Agt3494A Agt3494A1 
      Left            =   120
      Top             =   3240
      _ExtentX        =   953
      _ExtentY        =   847
      Address         =   "GPIB::22"
   End
   Begin VB.CommandButton cmdClearDisplay 
      Caption         =   "Clear Display"
      Height          =   375
      Left            =   4320
      TabIndex        =   5
      Top             =   3240
      Width           =   1815
   End
   Begin VB.TextBox txtDisplay 
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Text            =   " My 34401A"
      Top             =   2640
      Width           =   3855
   End
   Begin VB.CommandButton cmdDisplay 
      Caption         =   "Send to Display"
      Height          =   375
      Left            =   4320
      TabIndex        =   4
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Caption         =   "Measurements"
      Height          =   2175
      Left            =   120
      TabIndex        =   13
      Top             =   4080
      Width           =   5895
      Begin VB.TextBox txtReading 
         Height          =   375
         Index           =   0
         Left            =   2040
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtReading 
         Height          =   375
         Index           =   1
         Left            =   2040
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtReading 
         Height          =   375
         Index           =   2
         Left            =   2040
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1560
         Width           =   1215
      End
      Begin VB.CommandButton cmdGetReadings 
         Caption         =   "Get 3 Readings (dc)"
         Height          =   375
         Left            =   3600
         TabIndex        =   7
         Top             =   1200
         Width           =   1815
      End
      Begin VB.CommandButton cmdOneReading 
         Caption         =   "One Reading (ac)"
         Height          =   375
         Left            =   3600
         TabIndex        =   6
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Reading 1"
         Height          =   375
         Left            =   480
         TabIndex        =   19
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Reading 2"
         Height          =   375
         Left            =   480
         TabIndex        =   18
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Reading 3"
         Height          =   375
         Left            =   480
         TabIndex        =   17
         Top             =   1560
         Width           =   1335
      End
   End
   Begin VB.TextBox txtAddress 
      BackColor       =   &H8000000F&
      Height          =   375
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   480
      Width           =   3855
   End
   Begin VB.CommandButton cmdSetIO 
      Caption         =   "Set I/O"
      Height          =   375
      Left            =   4320
      TabIndex        =   0
      Top             =   480
      Width           =   1815
   End
   Begin VB.CommandButton cmdRevision 
      Caption         =   "Revision"
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      Top             =   1920
      Width           =   1815
   End
   Begin VB.TextBox txtVersion 
      Height          =   375
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1920
      Width           =   3855
   End
   Begin VB.TextBox txtID 
      Height          =   375
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1200
      Width           =   3855
   End
   Begin VB.CommandButton CmdID 
      Caption         =   "Get ID string"
      Height          =   375
      Left            =   4320
      TabIndex        =   1
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label7 
      Caption         =   "Send to Display"
      Height          =   255
      Left            =   240
      TabIndex        =   21
      Top             =   2400
      Width           =   3735
   End
   Begin VB.Label Label6 
      Caption         =   "Address"
      Height          =   375
      Left            =   240
      TabIndex        =   20
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "Revision"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   1680
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "ID string"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   960
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'' """""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
''  Copyright © 1999, 2000 Agilent Technologies Inc.  All rights reserved.
''
'' You have a royalty-free right to use, modify, reproduce and distribute
'' the Sample Application Files (and/or any modified version) in any way
'' you find useful, provided that you agree that Agilent Technologies has no
'' warranty,  obligations or liability for any Sample Application Files.
''
'' Agilent Technologies provides programming examples for illustration only,
'' This sample program assumes that you are familiar with the programming
'' language being demonstrated and the tools used to create and debug
'' procedures. Agilent Technologies support engineers can help explain the
'' functionality of Agilent Technologies software components and associated
'' commands, but they will not modify these samples to provide added
'' functionality or construct procedures to meet your specific needs.
'' """"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""


'************************************************************************
' Note;
'   To use RS232, you must first set the instrument to
'   remote with this command:
'   Agt3494A1.Output "Syst:Rem"
'************************************************************************
Private Sub cmdGetReadings_Click()
    ' Configure the multimeter for dc voltage readings,
    ' 10V range, 0.1V resolution, and 3 readings
    Dim Reading(2) As Double
    Dim strtemp As String
  
'""""""""""""""""""""""""""""""""""""
'   Include this line for RS232 only
   Agt3494A1.Output "Syst:Rem"

    With Agt3494A1
        .Output ":CONF:VOLT:DC 10, 0.1"
        .Output "SAMP:COUN 3"
        ' for RS232 only, a delay may be needed before the Read
        delay 200
        .Output "Read?"
        .Enter Reading
    End With
    
    txtReading(0) = Reading(0)
    txtReading(1) = Reading(1)
    txtReading(2) = Reading(2)
        
End Sub

Private Sub CmdID_Click()
    Dim result As String
    Dim IO_type As String
    
    Agt3494A1.Connect
    
    
    IO_type = Agt3494A1.IO.IOType
    
    
    Agt3494A1.Output "*idn?"
    Agt3494A1.Enter result
    
    txtID.Text = result
    
End Sub

Private Sub cmdOneReading_Click()
    ' Set the multimeter for ac voltage reading,
    ' Use default values
    ' Get the reading and put it in first text box
    Dim Reading As Double
    
    
'""""""""""""""""""""""""""""""""""""
'   Include this line for RS232 only
   Agt3494A1.Output "Syst:Rem"

    With Agt3494A1
        .Output "Measure:Voltage:AC?"
        .Enter Reading
    End With
    
    txtReading(0) = Reading
    
    txtReading(1) = ""
    txtReading(2) = ""
    
    Exit Sub

    
End Sub

Private Sub cmdRevision_Click()
    ' Gets the hardware revision from the instrument
    Dim result As String
    
    Agt3494A1.Output ":Syst:Vers?"
    Agt3494A1.Enter result
    
    txtVersion.Text = result
End Sub

Private Sub cmdSetIO_Click()
    ' set the I/O address to the text box in case the
    ' user changed it.
    ' bring up the IO dialog and save any changes to the
    ' text box
    
    Agt3494A1.Address = txtAddress
    
    Agt3494A1.ShowConnectDialog
    
    txtAddress = Agt3494A1.Address

        
End Sub

Private Sub Form_Load()
    ' load the address of the control into the text box
    txtAddress = Agt3494A1.Address
End Sub
Private Sub cmdDisplay_Click()
    ' Send a message to the multimeters display,
    ' and generate a beep
    Agt3494A1.Output ":syst:beep;:disp:text " & "'" & txtDisplay.Text & "'"
    
End Sub
Private Sub cmdClearDisplay_Click()
    ' Clear the display
    Agt3494A1.Output "Display:text:Clear"

End Sub


