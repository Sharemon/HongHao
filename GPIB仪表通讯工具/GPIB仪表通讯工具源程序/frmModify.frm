VERSION 5.00
Begin VB.Form frmModify 
   Caption         =   "ÐÞÕýÉèÖÃ"
   ClientHeight    =   7665
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   8475
   LinkTopic       =   "Form1"
   ScaleHeight     =   7665
   ScaleWidth      =   8475
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CheckBox AllValid 
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   7.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7800
      TabIndex        =   87
      Top             =   720
      Width           =   375
   End
   Begin VB.CommandButton cancel 
      Caption         =   "È¡ Ïû"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   6600
      TabIndex        =   73
      Top             =   6960
      Width           =   1400
   End
   Begin VB.TextBox standard 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Index           =   11
      Left            =   1320
      TabIndex        =   66
      Text            =   "Text1"
      Top             =   5880
      Width           =   1320
   End
   Begin VB.TextBox standard 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   1320
      TabIndex        =   65
      Text            =   "Text1"
      Top             =   5435
      Width           =   1320
   End
   Begin VB.TextBox standard 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   1320
      TabIndex        =   64
      Text            =   "Text1"
      Top             =   4997
      Width           =   1320
   End
   Begin VB.TextBox standard 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   1320
      TabIndex        =   63
      Text            =   "Text1"
      Top             =   4559
      Width           =   1320
   End
   Begin VB.TextBox standard 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   1320
      TabIndex        =   62
      Text            =   "Text1"
      Top             =   4121
      Width           =   1320
   End
   Begin VB.TextBox standard 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   1320
      TabIndex        =   61
      Text            =   "Text1"
      Top             =   3683
      Width           =   1320
   End
   Begin VB.TextBox standard 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   1320
      TabIndex        =   60
      Text            =   "Text1"
      Top             =   3245
      Width           =   1320
   End
   Begin VB.TextBox standard 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   1320
      TabIndex        =   59
      Text            =   "Text1"
      Top             =   2807
      Width           =   1320
   End
   Begin VB.TextBox standard 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   1320
      TabIndex        =   58
      Text            =   "Text1"
      Top             =   2369
      Width           =   1320
   End
   Begin VB.TextBox standard 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   1320
      TabIndex        =   57
      Text            =   "Text1"
      Top             =   1931
      Width           =   1320
   End
   Begin VB.TextBox standard 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1320
      TabIndex        =   56
      Text            =   "Text1"
      Top             =   1493
      Width           =   1320
   End
   Begin VB.TextBox standard 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Index           =   0
      Left            =   1320
      TabIndex        =   55
      Text            =   "Text1"
      Top             =   1080
      Width           =   1320
   End
   Begin VB.TextBox div 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6240
      TabIndex        =   52
      Text            =   "Text2"
      Top             =   6480
      Width           =   1215
   End
   Begin VB.TextBox multi 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1920
      TabIndex        =   50
      Text            =   "Text1"
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton ok 
      Caption         =   "È· ¶¨"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   4680
      TabIndex        =   0
      Top             =   6960
      Width           =   1400
   End
   Begin VB.CheckBox valid 
      Caption         =   "ÓÐÐ§"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   11
      Left            =   6840
      TabIndex        =   48
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CheckBox valid 
      Caption         =   "ÓÐÐ§"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   10
      Left            =   6840
      TabIndex        =   47
      Top             =   5440
      Width           =   1215
   End
   Begin VB.CheckBox valid 
      Caption         =   "ÓÐÐ§"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   9
      Left            =   6840
      TabIndex        =   46
      Top             =   5004
      Width           =   1215
   End
   Begin VB.CheckBox valid 
      Caption         =   "ÓÐÐ§"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   8
      Left            =   6840
      TabIndex        =   45
      Top             =   4568
      Width           =   1215
   End
   Begin VB.CheckBox valid 
      Caption         =   "ÓÐÐ§"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   7
      Left            =   6840
      TabIndex        =   44
      Top             =   4132
      Width           =   1215
   End
   Begin VB.CheckBox valid 
      Caption         =   "ÓÐÐ§"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   6
      Left            =   6840
      TabIndex        =   43
      Top             =   3696
      Width           =   1215
   End
   Begin VB.CheckBox valid 
      Caption         =   "ÓÐÐ§"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   5
      Left            =   6840
      TabIndex        =   42
      Top             =   3260
      Width           =   1215
   End
   Begin VB.CheckBox valid 
      Caption         =   "ÓÐÐ§"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   4
      Left            =   6840
      TabIndex        =   41
      Top             =   2824
      Width           =   1215
   End
   Begin VB.CheckBox valid 
      Caption         =   "ÓÐÐ§"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   3
      Left            =   6840
      TabIndex        =   40
      Top             =   2388
      Width           =   1215
   End
   Begin VB.CheckBox valid 
      Caption         =   "ÓÐÐ§"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   6840
      TabIndex        =   39
      Top             =   1952
      Width           =   1215
   End
   Begin VB.CheckBox valid 
      Caption         =   "ÓÐÐ§"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   6840
      TabIndex        =   38
      Top             =   1516
      Width           =   1215
   End
   Begin VB.TextBox real 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Index           =   11
      Left            =   4080
      TabIndex        =   36
      Text            =   "Text1"
      Top             =   5880
      Width           =   1320
   End
   Begin VB.TextBox real 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   4080
      TabIndex        =   35
      Text            =   "Text1"
      Top             =   5435
      Width           =   1320
   End
   Begin VB.TextBox real 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   4080
      TabIndex        =   34
      Text            =   "Text1"
      Top             =   4997
      Width           =   1320
   End
   Begin VB.TextBox real 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   4080
      TabIndex        =   33
      Text            =   "Text1"
      Top             =   4559
      Width           =   1320
   End
   Begin VB.TextBox real 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   4080
      TabIndex        =   32
      Text            =   "Text1"
      Top             =   4121
      Width           =   1320
   End
   Begin VB.TextBox real 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   4080
      TabIndex        =   31
      Text            =   "Text1"
      Top             =   3683
      Width           =   1320
   End
   Begin VB.TextBox real 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   4080
      TabIndex        =   30
      Text            =   "Text1"
      Top             =   3245
      Width           =   1320
   End
   Begin VB.TextBox real 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   4080
      TabIndex        =   29
      Text            =   "Text1"
      Top             =   2807
      Width           =   1320
   End
   Begin VB.TextBox real 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   4080
      TabIndex        =   28
      Text            =   "Text1"
      Top             =   2369
      Width           =   1320
   End
   Begin VB.TextBox real 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   4080
      TabIndex        =   27
      Text            =   "Text1"
      Top             =   1931
      Width           =   1320
   End
   Begin VB.TextBox real 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   4080
      TabIndex        =   26
      Text            =   "Text1"
      Top             =   1493
      Width           =   1320
   End
   Begin VB.TextBox real 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Index           =   0
      Left            =   4080
      TabIndex        =   25
      Text            =   "Text1"
      Top             =   1080
      Width           =   1320
   End
   Begin VB.ComboBox VIOChoice 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      ItemData        =   "frmModify.frx":0000
      Left            =   1680
      List            =   "frmModify.frx":000D
      TabIndex        =   53
      Text            =   "Combo1"
      Top             =   120
      Width           =   1695
   End
   Begin VB.CheckBox valid 
      Caption         =   "ÓÐÐ§"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   6840
      TabIndex        =   37
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   11
      Left            =   2880
      TabIndex        =   86
      Top             =   5880
      Width           =   1100
   End
   Begin VB.Label Label9 
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   2880
      TabIndex        =   85
      Top             =   5440
      Width           =   1100
   End
   Begin VB.Label Label9 
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   2880
      TabIndex        =   84
      Top             =   5004
      Width           =   1100
   End
   Begin VB.Label Label9 
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   2880
      TabIndex        =   83
      Top             =   4568
      Width           =   1100
   End
   Begin VB.Label Label9 
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   2880
      TabIndex        =   82
      Top             =   4132
      Width           =   1100
   End
   Begin VB.Label Label9 
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   2880
      TabIndex        =   81
      Top             =   3696
      Width           =   1100
   End
   Begin VB.Label Label9 
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   2880
      TabIndex        =   80
      Top             =   3260
      Width           =   1100
   End
   Begin VB.Label Label9 
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   2880
      TabIndex        =   79
      Top             =   2824
      Width           =   1100
   End
   Begin VB.Label Label9 
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   2880
      TabIndex        =   78
      Top             =   2388
      Width           =   1100
   End
   Begin VB.Label Label9 
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   2880
      TabIndex        =   77
      Top             =   1952
      Width           =   1100
   End
   Begin VB.Label Label9 
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   2880
      TabIndex        =   76
      Top             =   1516
      Width           =   1100
   End
   Begin VB.Label Label9 
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   2880
      TabIndex        =   75
      Top             =   1080
      Width           =   1100
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "ppm"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3000
      TabIndex        =   74
      Top             =   720
      Width           =   465
   End
   Begin VB.Label Label11 
      Caption         =   "³ýÊýÒò×Ó£º"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   72
      Top             =   6480
      Width           =   1215
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "ÊÇ·ñÓÐÐ§"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6840
      TabIndex        =   71
      Top             =   720
      Width           =   840
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "¸ºÐÞÕýÏµÊý"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4080
      TabIndex        =   70
      Top             =   720
      Width           =   1050
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "ÕýÐÞÕýÏµÊý"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1320
      TabIndex        =   69
      Top             =   720
      Width           =   1050
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Á¿³Ì"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   360
      TabIndex        =   68
      Top             =   720
      Width           =   420
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   8160
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label4 
      Caption         =   "ÏîÄ¿Ñ¡Ôñ£º"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   67
      Top             =   165
      Width           =   1215
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "ppm"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5760
      TabIndex        =   54
      Top             =   720
      Width           =   465
   End
   Begin VB.Label Label2 
      Caption         =   "³ýÊýÒò×Ó£º"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8280
      TabIndex        =   51
      Top             =   18480
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "³ËÊýÒò×Ó£º"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   49
      Top             =   6480
      Width           =   1215
   End
   Begin VB.Label ppm 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   11
      Left            =   5640
      TabIndex        =   24
      Top             =   5880
      Width           =   1100
   End
   Begin VB.Label ppm 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   10
      Left            =   5640
      TabIndex        =   23
      Top             =   5445
      Width           =   1100
   End
   Begin VB.Label ppm 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   9
      Left            =   5640
      TabIndex        =   22
      Top             =   5010
      Width           =   1100
   End
   Begin VB.Label ppm 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   8
      Left            =   5640
      TabIndex        =   21
      Top             =   4575
      Width           =   1100
   End
   Begin VB.Label ppm 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   7
      Left            =   5640
      TabIndex        =   20
      Top             =   4125
      Width           =   1100
   End
   Begin VB.Label ppm 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   6
      Left            =   5640
      TabIndex        =   19
      Top             =   3690
      Width           =   1100
   End
   Begin VB.Label ppm 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   5
      Left            =   5640
      TabIndex        =   18
      Top             =   3255
      Width           =   1100
   End
   Begin VB.Label ppm 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   4
      Left            =   5640
      TabIndex        =   17
      Top             =   2820
      Width           =   1100
   End
   Begin VB.Label ppm 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   3
      Left            =   5640
      TabIndex        =   16
      Top             =   2385
      Width           =   1100
   End
   Begin VB.Label ppm 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   2
      Left            =   5640
      TabIndex        =   15
      Top             =   1950
      Width           =   1100
   End
   Begin VB.Label ppm 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   5640
      TabIndex        =   14
      Top             =   1515
      Width           =   1100
   End
   Begin VB.Label ppm 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   5640
      TabIndex        =   13
      Top             =   1080
      Width           =   1100
   End
   Begin VB.Label label 
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   300
      TabIndex        =   12
      Top             =   5880
      Width           =   800
   End
   Begin VB.Label label 
      Caption         =   "11"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   300
      TabIndex        =   11
      Top             =   5440
      Width           =   800
   End
   Begin VB.Label label 
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   300
      TabIndex        =   10
      Top             =   5004
      Width           =   800
   End
   Begin VB.Label label 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   300
      TabIndex        =   9
      Top             =   4568
      Width           =   800
   End
   Begin VB.Label label 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   300
      TabIndex        =   8
      Top             =   4132
      Width           =   800
   End
   Begin VB.Label label 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   300
      TabIndex        =   7
      Top             =   3696
      Width           =   800
   End
   Begin VB.Label label 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   300
      TabIndex        =   6
      Top             =   3260
      Width           =   800
   End
   Begin VB.Label label 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   300
      TabIndex        =   5
      Top             =   2824
      Width           =   800
   End
   Begin VB.Label label 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   300
      TabIndex        =   4
      Top             =   2388
      Width           =   800
   End
   Begin VB.Label label 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   300
      TabIndex        =   3
      Top             =   1952
      Width           =   800
   End
   Begin VB.Label label 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   300
      TabIndex        =   2
      Top             =   1516
      Width           =   800
   End
   Begin VB.Label label 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   300
      TabIndex        =   1
      Top             =   1080
      Width           =   800
   End
End
Attribute VB_Name = "frmModify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AllValid_Click()
    Dim i As Integer
    For i = 0 To 11
        Me.valid(i).value = Me.AllValid.value
    Next
End Sub

Private Sub cancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    If (frmMain.connectFlag) Then
        Me.VIOChoice.ListIndex = frmMain.VIO
    Else
        Me.VIOChoice.ListIndex = 0
    End If
    Me.multi.Text = dataDim.multi(Me.VIOChoice.ListIndex, 0)
    Me.div.Text = dataDim.div(Me.VIOChoice.ListIndex, 0)
End Sub


Private Sub ok_Click()
    Call savePara
    Unload Me
    Call frmMain.Set_Memory
End Sub

Private Sub real_Change(Index As Integer)
    Dim temp As Double
    temp = Round((Val(Me.real(Index).Text) - 1), 10) * 1000000
    ppm(Index).Caption = Trim(str(temp))
End Sub

Private Sub standard_Change(Index As Integer)
    Dim temp As Double
    temp = Round((Val(Me.standard(Index).Text) - 1), 10) * 1000000
    Label9(Index).Caption = Trim(str(temp))
End Sub

Private Sub savePara()
    Dim i As Integer
    For i = 0 To 11
        dataDim.modiPara(Me.VIOChoice.ListIndex, i, 0) = Val(Me.standard(i).Text)
        dataDim.modiPara(Me.VIOChoice.ListIndex, i, 1) = Val(Me.real(i).Text)
        dataDim.valid(Me.VIOChoice.ListIndex, i) = Me.valid(i).value
    Next
    dataDim.multi(Me.VIOChoice.ListIndex, 0) = Val(Me.multi.Text)
    dataDim.div(Me.VIOChoice.ListIndex, 0) = Val(Me.div.Text)
End Sub

Private Sub VIOChoice_Click()
Dim i As Integer
    Select Case Me.VIOChoice.ListIndex
    Case 0
        Me.label(8).Visible = False
        Me.label(9).Visible = False
        Me.label(10).Visible = False
        Me.label(11).Visible = False
        Me.standard(8).Visible = False
        Me.standard(9).Visible = False
        Me.standard(10).Visible = False
        Me.standard(11).Visible = False
        Me.Label9(8).Visible = False
        Me.Label9(9).Visible = False
        Me.Label9(10).Visible = False
        Me.Label9(11).Visible = False
        Me.real(8).Visible = False
        Me.real(9).Visible = False
        Me.real(10).Visible = False
        Me.real(11).Visible = False
        Me.ppm(8).Visible = False
        Me.ppm(9).Visible = False
        Me.ppm(10).Visible = False
        Me.ppm(11).Visible = False
        Me.valid(8).Visible = False
        Me.valid(9).Visible = False
        Me.valid(10).Visible = False
        Me.valid(11).Visible = False
        Me.label(0).Caption = "100uV"
        Me.label(1).Caption = "1mV"
        Me.label(2).Caption = "10mV"
        Me.label(3).Caption = "100mV"
        Me.label(4).Caption = "1V"
        Me.label(5).Caption = "10V"
        Me.label(6).Caption = "100V"
        Me.label(7).Caption = "1000V"
        Me.Height = 6340
        Me.ok.Top = 5040
        Me.cancel.Top = 5040
        Me.multi.Top = 5040 - 480
        Me.div.Top = 5040 - 480
        Me.Label1.Top = 5040 - 480
        Me.Label11.Top = 5040 - 480
        Me.AllValid.value = 0
        For i = 0 To 7
            Me.standard(i).Text = Format(dataDim.modiPara(0, i, 0), "0.00000000")
            Me.real(i).Text = Format(dataDim.modiPara(0, i, 1), "0.00000000")
            Me.valid(i).value = dataDim.valid(0, i)
        Next
        Me.multi.Text = dataDim.multi(0, 0)
        Me.div.Text = dataDim.div(0, 0)
        For i = 0 To 7
            If (Me.valid(i).value = 0) Then Exit Sub
        Next
        Me.AllValid.value = 1
    Case 1
        Me.label(8).Visible = True
        Me.label(9).Visible = False
        Me.label(10).Visible = False
        Me.label(11).Visible = False
        Me.standard(8).Visible = True
        Me.standard(9).Visible = False
        Me.standard(10).Visible = False
        Me.standard(11).Visible = False
        Me.Label9(8).Visible = True
        Me.Label9(9).Visible = False
        Me.Label9(10).Visible = False
        Me.Label9(11).Visible = False
        Me.real(8).Visible = True
        Me.real(9).Visible = False
        Me.real(10).Visible = False
        Me.real(11).Visible = False
        Me.ppm(8).Visible = True
        Me.ppm(9).Visible = False
        Me.ppm(10).Visible = False
        Me.ppm(11).Visible = False
        Me.valid(8).Visible = True
        Me.valid(9).Visible = False
        Me.valid(10).Visible = False
        Me.valid(11).Visible = False
        Me.label(0).Caption = "1uA"
        Me.label(1).Caption = "10uA"
        Me.label(2).Caption = "100uA"
        Me.label(3).Caption = "1mA"
        Me.label(4).Caption = "10mA"
        Me.label(5).Caption = "100mA"
        Me.label(6).Caption = "1A"
        Me.label(7).Caption = "3A"
        Me.label(8).Caption = "10A"
        Me.Height = 6820
        Me.ok.Top = 5520
        Me.cancel.Top = 5520
        Me.multi.Top = 5040
        Me.div.Top = 5040
        Me.Label1.Top = 5040
        Me.Label11.Top = 5040
        Me.AllValid.value = 0
        For i = 0 To 8
            Me.standard(i).Text = Format(dataDim.modiPara(1, i, 0), "0.00000000")
            Me.real(i).Text = Format(dataDim.modiPara(1, i, 1), "0.00000000")
            Me.valid(i).value = dataDim.valid(1, i)
        Next
        Me.multi.Text = dataDim.multi(1, 0)
        Me.div.Text = dataDim.div(1, 0)
        For i = 0 To 8
            If (Me.valid(i).value = 0) Then Exit Sub
        Next
        Me.AllValid.value = 1
    Case 2
        Me.label(8).Visible = True
        Me.label(9).Visible = True
        Me.label(10).Visible = True
        Me.label(11).Visible = True
        Me.standard(8).Visible = True
        Me.standard(9).Visible = True
        Me.standard(10).Visible = True
        Me.standard(11).Visible = True
        Me.Label9(8).Visible = True
        Me.Label9(9).Visible = True
        Me.Label9(10).Visible = True
        Me.Label9(11).Visible = True
        Me.real(8).Visible = True
        Me.real(9).Visible = True
        Me.real(10).Visible = True
        Me.real(11).Visible = True
        Me.ppm(8).Visible = True
        Me.ppm(9).Visible = True
        Me.ppm(10).Visible = True
        Me.ppm(11).Visible = True
        Me.valid(8).Visible = True
        Me.valid(9).Visible = True
        Me.valid(10).Visible = True
        Me.valid(11).Visible = True
        Me.label(0).Caption = "1¦¸"
        Me.label(1).Caption = "10¦¸"
        Me.label(2).Caption = "100¦¸"
        Me.label(3).Caption = "1K¦¸"
        Me.label(4).Caption = "10K¦¸"
        Me.label(5).Caption = "100K¦¸"
        Me.label(6).Caption = "1M¦¸"
        Me.label(7).Caption = "10M¦¸"
        Me.label(8).Caption = "100M¦¸"
        Me.label(9).Caption = "1G¦¸"
        Me.label(10).Caption = "10G¦¸"
        Me.label(11).Caption = "100G¦¸"
        Me.Height = 8230
        Me.ok.Top = 6960
        Me.cancel.Top = 6960
        Me.multi.Top = 6480
        Me.div.Top = 6480
        Me.Label1.Top = 6480
        Me.Label11.Top = 6480
        Me.AllValid.value = 0
        For i = 0 To 11
            Me.standard(i).Text = Format(dataDim.modiPara(2, i, 0), "0.00000000")
            Me.real(i).Text = Format(dataDim.modiPara(2, i, 1), "0.00000000")
            Me.valid(i).value = dataDim.valid(2, i)
        Next
        Me.multi.Text = dataDim.multi(2, 0)
        Me.div.Text = dataDim.div(2, 0)
        For i = 0 To 11
            If (Me.valid(i).value = 0) Then Exit Sub
        Next
        Me.AllValid.value = 1
    End Select
End Sub

Private Sub VIOChoice_DropDown()
    Call savePara
End Sub
