VERSION 5.00
Begin VB.Form MAIN 
   Caption         =   "请选择I/O接口"
   ClientHeight    =   3432
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   2784
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3432
   ScaleWidth      =   2784
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton GPIBlinker 
      Caption         =   "GPIB连接器"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   360
      TabIndex        =   2
      Top             =   1920
      Width           =   2052
   End
   Begin VB.CommandButton USBlinker 
      Caption         =   "USB连接器"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   360
      TabIndex        =   1
      Top             =   1080
      Width           =   2052
   End
   Begin VB.CommandButton LANlinker 
      Caption         =   "以太网连接器"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   2052
   End
   Begin VB.Label Label1 
      Caption         =   "Agilent34410A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   480
      TabIndex        =   3
      Top             =   2880
      Width           =   1932
   End
End
Attribute VB_Name = "MAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub GPIBlinker_Click()
GPIB.Show
LAN.Hide
USB.Hide
End Sub

Private Sub LANlinker_Click()
LAN.Show
USB.Hide
GPIB.Hide
End Sub

Private Sub USBlinker_Click()
USB.Show
LAN.Hide
GPIB.Hide
End Sub
