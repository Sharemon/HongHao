VERSION 5.00
Object = "{1C98F15C-068A-11D4-98C2-00108301CB39}#2.0#0"; "agt3494A.ocx"
Begin VB.Form frmStatReg 
   Caption         =   "Form1"
   ClientHeight    =   3570
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3570
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   495
      Left            =   2760
      TabIndex        =   6
      Top             =   2880
      Width           =   1695
   End
   Begin VB.TextBox txtData 
      Height          =   2295
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Timer tmrPollForSRQ 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3720
      Top             =   1920
   End
   Begin Agt3494ALib.Agt3494A Agt3494A1 
      Left            =   2880
      Top             =   1800
      _ExtentX        =   1085
      _ExtentY        =   873
      Address         =   "GPIB0::22::INSTR"
      Timeout         =   10000
   End
   Begin VB.TextBox txtAddress 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   2175
   End
   Begin VB.CommandButton cmdSetIO 
      Caption         =   "Set I/O"
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton cmdStartReading 
      Caption         =   "Start Readings"
      Height          =   495
      Left            =   2760
      TabIndex        =   0
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Data"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Address"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmStatReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
''' """""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
'''  Copyright © 2000 Agilent Technologies Inc.  All rights reserved.
'''
''' You have a royalty-free right to use, modify, reproduce and distribute
''' the Sample Application Files (and/or any modified version) in any way
''' you find useful, provided that you agree that Agilent Technologies has no
''' warranty,  obligations or liability for any Sample Application Files.
'''
''' Agilent Technologies provides programming examples for illustration only,
''' This sample program assumes that you are familiar with the programming
''' language being demonstrated and the tools used to create and debug
''' procedures. Agilent Technologies support engineers can help explain the
''' functionality of Agilent Technologies software components and associated
''' commands, but they will not modify these samples to provide added
''' functionality or construct procedures to meet your specific needs.
''' """"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
'

'*************************************************************
' The following example shows how you can use the multimeter's status
' registers to determine when a command sequence is completed. For
' more information see "The SCPI Status Model" in the Agilent 34401A
' User Guide
'
' NOTE: Polling will only work with a GPIB card. It does not
'       work for RS232 or with USB to GPIB convertors available
'       at the time of this writing.
'
' Sequence of Operation;
'   1. The meter is cleared and set to give an SRQ when its
'      operation is complete
'   2. Enable the GPIB port to look for an SRQ so we can
'      poll (repeatedly query) for it.
'   2. The meter is set for dc, and multiple readings. This will
'      take about 5 seconds for 10 readings
'   3. We start the reading with INIT. This will put the
'      data into memory.  When the meter is finished, it
'      will set SRQ.
'   4. Enable the timer that polls for SRQ every second.
'   5. When SRQ is detected, then get the reading from the
'      meter with the routine ReadData.
'
'Created on:   03/28/00
'Module:       frmStatReg
'Project:      StatReg_SRQ
'*************************************************************
Dim GPIBserver As IIOGPIB
Dim statusValue As Byte
Dim numberReadings As Long

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSetIO_Click()

    ' Set the io control address to the text box address
    ' we do this so the user can change address in text box
    ' and it will be reflected in the dialog
    Agt3494A1.Address = txtAddress.Text
    
    ' show the communication dialog
    Agt3494A1.ShowConnectDialog
    
    ' Put the address from the communication dialog in text box
    txtAddress.Text = Agt3494A1.Address
End Sub

Private Sub cmdStartReading_Click()
    ' Call the routine that sets up the meter
    
    cmdStartReading.Enabled = False
    
    startReadings
    
    
    cmdStartReading.Enabled = True
End Sub

Private Sub Form_Load()
    ' Load the forms address text box with persistant
    ' instrument address from agt3494A object
    txtAddress.Text = Agt3494A1.Address
End Sub

Private Sub startReadings()

    Dim Average As Double
    Dim MinReading As Double
    Dim MaxReading As Double
    Dim Value As Integer
    Dim Mask As Integer
    Dim Task As Integer
    
    On Error GoTo StartReadingsError
    
    ' Clear out text box for the data so we can see
    ' when new data arrives
    txtData.Text = ""
    txtData.Refresh
    
    ' reset the reference to the GPIB server
    Set GPIBserver = Nothing
    
    '""""""""""""""""""""""""""""""""""""""""""""""""""
    ' Setup dmm to return an SRQ event when readings are complete
    With Agt3494A1
        ' Set the address from users text box
        .Address = txtAddress.Text
        .Output "*RST"          ' Reset dmm
        .Output "*CLS"          ' Clear dmm status registers
        .Output "*ESE 1"        ' Enable 'operation complete bit to
                                '  set 'standard event' bit in status byte
        .Output "*SRE 32"       ' Enable 'standard event' bit in status
                                '  byte to pull the IEEE-488 SRQ line
        .Output "*OPC?"         ' Assure syncronization
        .Enter Value
    End With
    
    '""""""""""""""""""""""""""""""""""""""""""""""""""
    ' Enable the SRQ polling so we can poll for SRQ
    SetSRQforPolling
    txtData.Refresh
    
    '"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
    ' Configure the meter to take readings
    ' and initiate the readings (source is set to immediate by default)
    numberReadings = 10
    With Agt3494A1
        .Output "Configure:Voltage:dc 10"   ' set dmm to 10 volt dc range"
        .Output "Voltage:DC:NPLC 10"        ' set integration time to 10 Power line cycles (PLC)"
        .Output "Trigger:count" & Str$(numberReadings) ' set dmm to accept multiple triggers
        .Output "Init"                      ' Place dmm in 'wait-for-trigger' state
        .Output "*OPC"                      ' Set 'operation complete' bit in standard
                                            ' event registers when measurement is complete
    End With
    
    '""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
    
    ' enable the timer to poll gpib once per second
    tmrPollForSRQ.Enabled = True
    
    ' give message that meter is initialized
    ' give message that configuration is done
    txtData.Text = "Meter configured and " & vbCrLf & _
                    "Initialized"

    Exit Sub
    
StartReadingsError:
    Debug.Print "Start Readings Error = "; Err.Description
End Sub

Private Sub tmrPollForSRQ_Timer()
    ' This timer when enabled will poll the GPIB
    ' and return the status byte to indicate if SRQ is set
    
    On Error GoTo pollerror
    
    statusValue = 0
    ' this will also change the timeout
    statusValue = GPIBserver.WaitOnSRQ(100)
    
    ' Set the timeout back to 10 sec,
    ' the GPIBserver is set in SetSRQforPolling routine
    GPIBserver.Timeout = 10000
    If statusValue And 64 Then       ' SRQ from Operation complete
        ' Turn off the timer and stop polling.
        tmrPollForSRQ.Enabled = False
        Debug.Print "SRQ is set, getting data"
        ' Get the Data, the meter is ready
        ReadData
    End If
    
    Exit Sub
    
pollerror:
     Debug.Print "No SRQ yet, Poll error = "; Err.Description

End Sub
Private Sub SetSRQforPolling()
    ' Enble the GPIB for polling
    
    ' to have access to the GPIB polling event we use
    ' the io object to return a GPIB server object that
    ' supports additional capabilities
    ' GPIBserver is a global variable for this module
    
    'Agt3494A1.Connect     ' this not needed if agt3494A object is called previously
    Set GPIBserver = Nothing
    Set GPIBserver = Agt3494A1.IO
    
    ' Set the GPIB I/O for polling
    GPIBserver.EnableSRQ IO_Event_Queued
    
End Sub

Private Sub ReadData()
    ' Once the SRQ is detected, this routine will
    ' get the data from the meter
    ' Called by: PollForSRQTimer_Timer
    '
    Dim readings() As Double
    Dim i As Long
    
    On Error GoTo ReadDataError
    
    ' dimension the array for the number of readings
    ReDim readings(numberReadings - 1)
    
    With Agt3494A1
        .Output "Fetch?"        ' Query for the data in memory
        .Enter readings         ' get the data and parse into the array
    End With
    
    ' Insert data into text box
    txtData.Text = ""
    For i = 0 To numberReadings - 1
        txtData.SelText = readings(i) & " Vdc" & vbCrLf
    Next i
    
    Exit Sub
    
ReadDataError:
    Debug.Print "ReadData Error = "; Err.Description
    
End Sub

Private Sub Form_Terminate()

    On Error Resume Next
    
    Set GPIBserver = Nothing

End Sub


