Attribute VB_Name = "GPIBGLOBAL"
'==========================================
' 32-bit Visual Basic Language Interface
' Global valiables define
' CONTEC Co.,Ltd.  2001.06.07
'==========================================

Option Explicit

'--------------------
' Global variables
'--------------------
Global ibsta As Integer
Global iberr As Integer
Global ibcnt As Integer
Global ibcntl As Long
Global buf As String
Global bytebuf() As Byte


'--------------------
' Needed to register for GPIB Global Thread.
'--------------------
Global Longibsta As Long
Global Longiberr As Long
Global Longibcnt As Long
Global GPIBglobalsRegistered As Integer


'--------------------
' Command defines
'--------------------
Global Const UNL = &H3F
Global Const UNT = &H5F
Global Const GTL = &H1
Global Const SDC = &H4
Global Const PPC = &H5
Global Const GGET = &H8
Global Const TCT = &H9
Global Const LLO = &H11
Global Const DCL = &H14
Global Const PPU = &H15
Global Const SPE = &H18
Global Const SPD = &H19
Global Const PPE = &H60
Global Const PPD = &H70


'--------------------
' Status bit mask in ibsta
'--------------------
Global Const EERR = &H8000      ' Error detected
Global Const TIMO = &H4000      ' Timeout occured
Global Const EEND = &H2000      ' EOI or EOS detected
Global Const SRQI = &H1000      ' SRQ detected by CIC
Global Const RQS = &H800        ' Device requesting any service
Global Const CMPL = &H100       ' I/O completed
Global Const LOK = &H80         ' Local lockout state
Global Const RREM = &H40        ' Remote enable state
Global Const CIC = &H20         ' Controller-in-Charge state
Global Const AATN = &H10        ' Attention line asserted state
Global Const TACS = &H8         ' Talker active state
Global Const LACS = &H4         ' Listener active state
Global Const DTAS = &H2         ' Device trigger state
Global Const DCAS = &H1         ' Device clear state


'--------------------
' Error messages in iberr
'--------------------
Global Const EDVR = 0           ' System error
Global Const ECIC = 1           ' Function requires GPIB board to be CIC
Global Const ENOL = 2           ' Write function detected no Listeners
Global Const EADR = 3           ' Interface board not addressed correctly
Global Const EARG = 4           ' Invalid argument to function call
Global Const ESAC = 5           ' Function requires GPIB board to be SAC
Global Const EABO = 6           ' I/O operation aborted
Global Const ENEB = 7           ' Non-existent interface board
Global Const EDMA = 8           ' Error performing DMA
Global Const EOIP = 10          ' I/O operation started before previous operation completed
Global Const ECAP = 11          ' No capability for intended operation
Global Const EFSO = 12          ' File system operation error
Global Const EBUS = 14          ' Command error during device call
Global Const ESTB = 15          ' Serial poll status byte lost
Global Const ESRQ = 16          ' SRQ remains asserted
Global Const ETAB = 20          ' The return buffer is full
Global Const ELCK = 21          ' Address or board is locked


'--------------------
' EOS mode bits
'--------------------
Global Const BIN = &H1000       ' EOS compare with eight bit
Global Const XEOS = &H800       ' Send EOI with EOS byte
Global Const REOS = &H400       ' Terminate read on EOS


'--------------------
' Timeout values and meanings
'--------------------
Global Const TNONE = 0      ' Infinite timeout (disabled)
Global Const T10us = 1      ' Timeout of 10 uSec
Global Const T30us = 2      ' Timeout of 30 uSec
Global Const T100us = 3     ' Timeout of 100 uSec
Global Const T300us = 4     ' Timeout of 300 uSec
Global Const T1ms = 5       ' Timeout of 1 mSec
Global Const T3ms = 6       ' Timeout of 3 mSec
Global Const T10ms = 7      ' Timeout of 10 mSec
Global Const T30ms = 8      ' Timeout of 30 mSec
Global Const T100ms = 9     ' Timeout of 100 mSec
Global Const T300ms = 10    ' Timeout of 300 mSec
Global Const T1s = 11       ' Timeout of 1 Sec
Global Const T3s = 12       ' Timeout of 3 Sec
Global Const T10s = 13      ' Timeout of 10 Sec
Global Const T30s = 14      ' Timeout of 30 Sec
Global Const T100s = 15     ' Timeout of 100 Sec
Global Const T300s = 16     ' Timeout of 300 Sec
Global Const T1000s = 17    ' Timeout of 1000 Sec


'--------------------
' Secondary address setting
'--------------------
Global Const ALL_SAD = -1   ' No Secondary address use
Global Const NO_SAD = 0     ' All Secondary address check


'--------------------
' ibconfig parameter defines
'--------------------
Global Const IbcPAD = &H1             ' Primary Address setting
Global Const IbcSAD = &H2             ' Secondary Address setting
Global Const IbcTMO = &H3             ' Timeout Value setting
Global Const IbcEOT = &H4             ' Send EOI with last data byte setting
Global Const IbcPPC = &H5             ' Parallel Poll Configure setting
Global Const IbcREADDR = &H6          ' Repeat Addressing setting
Global Const IbcAUTOPOLL = &H7        ' Disable Auto Serial Polling setting
Global Const IbcCICPROT = &H8         ' Use the CIC Protocol setting (Not supported by CONTEC)
Global Const IbcIRQ = &H9             ' Use PIO for I/O setting (Not supported by CONTEC)
Global Const IbcSC = &HA              ' System Controller setting
Global Const IbcSRE = &HB             ' Assert SRE on device calls setting
Global Const IbcEOSrd = &HC           ' Terminate reads on EOS setting
Global Const IbcEOSwrt = &HD          ' Send EOI with EOS character setting
Global Const IbcEOScmp = &HE          ' Use 7 or 8-bit EOS compare setting
Global Const IbcEOSchar = &HF         ' The EOS character setting
Global Const IbcPP2 = &H10            ' Use Parallel Poll Mode 2 setting
Global Const IbcTIMING = &H11         ' NORMAL, HIGH, or VERY_HIGH timing setting
Global Const IbcDMA = &H12            ' Use DMA for I/O setting
Global Const IbcReadAdjust = &H13     ' Swap bytes during an ibrd setting
Global Const IbcWriteAdjust = &H14    ' Swap bytes during an ibwrt setting
Global Const IbcSendLLO = &H17        ' Enable/disable the sending of LLO setting
Global Const IbcSPollTime = &H18      ' Set the timeout value for serial polls setting
Global Const IbcPPollTime = &H19      ' Set the parallel poll length period setting
Global Const IbcEndBitIsNormal = &H1A ' Remove EOS from END bit of IBSTA setting
Global Const IbcUnAddr = &H1B         ' Enable/disable device unaddressing setting
Global Const IbcHSCableLength = &H1F  ' Enable/disable high-speed handshaking setting (Not supported by CONTEC)
Global Const IbcIst = &H20            ' Set the IST bit setting
Global Const IbcRsv = &H21            ' Set the RSV bit setting


'--------------------
' ibask parameter defines
'--------------------
Global Const IbaPAD = &H1             ' Primary Address setting
Global Const IbaSAD = &H2             ' Secondary Address setting
Global Const IbaTMO = &H3             ' Timeout Value setting
Global Const IbaEOT = &H4             ' Send EOI with last data byte? setting
Global Const IbaPPC = &H5             ' Parallel Poll Configure setting
Global Const IbaREADDR = &H6          ' Repeat Addressing setting
Global Const IbaAUTOPOLL = &H7        ' Disable Auto Serial Polling setting
Global Const IbaCICPROT = &H8         ' Use the CIC Protocol setting (Not supported by CONTEC)
Global Const IbaIRQ = &H9             ' Use PIO for I/O setting (Not supported by CONTEC)
Global Const IbaSC = &HA              ' System Controller setting
Global Const IbaSRE = &HB             ' Assert SRE on device calls setting
Global Const IbaEOSrd = &HC           ' Terminate reads on EOS setting
Global Const IbaEOSwrt = &HD          ' Send EOI with EOS character setting
Global Const IbaEOScmp = &HE          ' Use 7 or 8-bit EOS compare setting
Global Const IbaEOSchar = &HF         ' The EOS character setting
Global Const IbaPP2 = &H10            ' Use Parallel Poll Mode 2 setting
Global Const IbaTIMING = &H11         ' NORMAL, HIGH, or VERY_HIGH timing setting
Global Const IbaDMA = &H12            ' Use DMA for I/O setting
Global Const IbaReadAdjust = &H13     ' Swap bytes during an ibrd setting
Global Const IbaWriteAdjust = &H14    ' Swap bytes during an ibwrt setting
Global Const IbaSendLLO = &H17        ' Enable/disable the sending of LLO setting
Global Const IbaSPollTime = &H18      ' Set the timeout value for serial polls setting
Global Const IbaPPollTime = &H19      ' Set the parallel poll length period setting
Global Const IbaEndBitIsNormal = &H1A ' Remove EOS from END bit of ibsta setting
Global Const IbaUnAddr = &H1B         ' Enable/disable device unaddressing setting
Global Const IbaHSCableLength = &H1F  ' Enable/disable high-speed handshaking setting (Not supported by CONTEC)
Global Const IbaIst = &H20            ' Set the IST bit setting
Global Const IbaRsv = &H21            ' Set the RSV bit setting
Global Const IbaBNA = &H200           ' A device's access board setting


'--------------------
' iblines parameter defines
'--------------------
Global Const ValidEOI = &H80
Global Const ValidATN = &H40
Global Const ValidSRQ = &H20
Global Const ValidREN = &H10
Global Const ValidIFC = &H8
Global Const ValidNRFD = &H4
Global Const ValidNDAC = &H2
Global Const ValidDAV = &H1
Global Const BusEOI = &H8000
Global Const BusATN = &H4000
Global Const BusSRQ = &H2000
Global Const BusREN = &H1000
Global Const BusIFC = &H800
Global Const BusNRFD = &H400
Global Const BusNDAC = &H200
Global Const BusDAV = &H100


'--------------------
' iblines parameter defines
'--------------------
Global Const NULLend = &H0            ' Do nothing at the end of a transfer
Global Const NLend = &H1              ' Send NL with EOI after a transfer
Global Const DABend = &H2             ' Send EOI with the last DAB


'--------------------
' Values used by the 488.2 Receive command
'--------------------
Global Const STOPend = &H100          ' Stop the read on EOI

'--------------------
' NOADDR define
'--------------------
Global Const NOADDR = &HFFFF


'--------------------
' ibnotify error code
'--------------------
Global Const IBNOTIFY_REARM_FAILED = &HE00A003F


