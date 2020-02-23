VERSION 5.00
Begin VB.Form frmMapa 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   ClientHeight    =   5640
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8880
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   8880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   610
      TabIndex        =   0
      Top             =   5190
      Width           =   3135
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   158
      Left            =   550
      Top             =   1520
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   157
      Left            =   550
      Top             =   1200
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   156
      Left            =   550
      Top             =   840
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   155
      Left            =   550
      Top             =   480
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   154
      Left            =   5160
      Top             =   4320
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   153
      Left            =   4500
      Top             =   4300
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   152
      Left            =   5160
      Top             =   3960
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   151
      Left            =   4520
      Top             =   3960
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   150
      Left            =   5160
      Top             =   3600
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   149
      Left            =   4500
      Top             =   3600
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   148
      Left            =   6480
      Top             =   3240
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   147
      Left            =   5800
      Top             =   3240
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   146
      Left            =   6480
      Top             =   2900
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   145
      Left            =   5800
      Top             =   2920
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   144
      Left            =   6480
      Top             =   2580
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   143
      Left            =   5800
      Top             =   2600
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   142
      Left            =   6480
      Top             =   2200
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   141
      Left            =   5800
      Top             =   2200
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   140
      Left            =   6480
      Top             =   1920
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   139
      Left            =   5800
      Top             =   1920
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   138
      Left            =   6480
      Top             =   1560
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   137
      Left            =   5800
      Top             =   1560
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   136
      Left            =   6480
      Top             =   1200
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   135
      Left            =   5800
      Top             =   1200
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   134
      Left            =   6480
      Top             =   840
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   133
      Left            =   5800
      Top             =   840
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   132
      Left            =   6480
      Top             =   480
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   131
      Left            =   5800
      Top             =   480
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   130
      Left            =   550
      Top             =   4620
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   129
      Left            =   550
      Top             =   4320
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   128
      Left            =   550
      Top             =   3960
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   127
      Left            =   550
      Top             =   3600
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   126
      Left            =   550
      Top             =   3240
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   125
      Left            =   550
      Top             =   2925
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   124
      Left            =   550
      Top             =   2538
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   123
      Left            =   550
      Top             =   2200
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   122
      Left            =   550
      Top             =   1880
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   117
      Left            =   7800
      Top             =   4620
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   116
      Left            =   7150
      Top             =   4620
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   115
      Left            =   7800
      Top             =   4320
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   114
      Left            =   7150
      Top             =   4320
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   113
      Left            =   7800
      Top             =   3960
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   112
      Left            =   7150
      Top             =   3960
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   111
      Left            =   7800
      Top             =   3600
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   110
      Left            =   7150
      Top             =   3600
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   109
      Left            =   7800
      Top             =   3240
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   108
      Left            =   7150
      Top             =   3240
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   107
      Left            =   7150
      Top             =   2900
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   106
      Left            =   7800
      Top             =   2900
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   105
      Left            =   7800
      Top             =   2580
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   104
      Left            =   7150
      Top             =   2580
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   103
      Left            =   7800
      Top             =   2200
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   102
      Left            =   7150
      Top             =   2200
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   101
      Left            =   7800
      Top             =   1900
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   100
      Left            =   7150
      Top             =   1900
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   99
      Left            =   7800
      Top             =   1560
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   98
      Left            =   7150
      Top             =   1560
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   97
      Left            =   7800
      Top             =   1200
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   96
      Left            =   7150
      Top             =   1200
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   95
      Left            =   7800
      Top             =   840
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   94
      Left            =   7150
      Top             =   840
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   93
      Left            =   7800
      Top             =   480
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   92
      Left            =   7150
      Top             =   480
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   91
      Left            =   1880
      Top             =   1200
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   90
      Left            =   3180
      Top             =   1200
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   89
      Left            =   3840
      Top             =   480
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   87
      Left            =   5160
      Top             =   480
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   86
      Left            =   2520
      Top             =   1200
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   84
      Left            =   3840
      Top             =   4680
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   83
      Left            =   2520
      Top             =   840
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   85
      Left            =   3200
      Top             =   840
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   88
      Left            =   4560
      Top             =   480
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   82
      Left            =   3180
      Top             =   1560
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   81
      Left            =   5160
      Top             =   1880
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   80
      Left            =   5160
      Top             =   2200
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   79
      Left            =   5160
      Top             =   2580
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   78
      Left            =   5160
      Top             =   3290
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   77
      Left            =   5160
      Top             =   2930
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   76
      Left            =   4520
      Top             =   2930
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   75
      Left            =   3840
      Top             =   2520
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   74
      Left            =   4520
      Top             =   2580
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   73
      Left            =   4520
      Top             =   2200
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   72
      Left            =   4520
      Top             =   1880
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   71
      Left            =   4520
      Top             =   1560
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   70
      Left            =   3840
      Top             =   1900
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   69
      Left            =   3840
      Top             =   2200
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   68
      Left            =   3200
      Top             =   2200
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   66
      Left            =   3840
      Top             =   1520
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   65
      Left            =   5160
      Top             =   1560
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   64
      Left            =   5160
      Top             =   1200
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   63
      Left            =   4500
      Top             =   1200
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   62
      Left            =   3840
      Top             =   840
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   61
      Left            =   4500
      Top             =   840
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   67
      Left            =   3180
      Top             =   1880
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   60
      Left            =   3200
      Top             =   480
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   59
      Left            =   2520
      Top             =   480
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   58
      Left            =   1880
      Top             =   480
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   57
      Left            =   1200
      Top             =   480
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   56
      Left            =   4500
      Top             =   3290
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   55
      Left            =   6480
      Top             =   3600
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   54
      Left            =   6480
      Top             =   3960
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   53
      Left            =   6480
      Top             =   4320
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   52
      Left            =   5160
      Top             =   4680
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   51
      Left            =   3840
      Top             =   2930
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   50
      Left            =   3180
      Top             =   2538
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   49
      Left            =   2520
      Top             =   2200
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   48
      Left            =   1880
      Top             =   2200
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   47
      Left            =   2520
      Top             =   1880
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   46
      Left            =   2520
      Top             =   1560
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   45
      Left            =   1880
      Top             =   1560
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   44
      Left            =   3840
      Top             =   1200
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   43
      Left            =   5160
      Top             =   840
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   42
      Left            =   1200
      Top             =   840
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   41
      Left            =   1200
      Top             =   1200
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   40
      Left            =   1200
      Top             =   1560
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   39
      Left            =   1880
      Top             =   1880
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   38
      Left            =   1200
      Top             =   1880
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   37
      Left            =   1200
      Top             =   2200
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   36
      Left            =   1200
      Top             =   2538
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   35
      Left            =   1200
      Top             =   2925
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   34
      Left            =   1880
      Top             =   2538
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   33
      Left            =   2520
      Top             =   2538
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   32
      Left            =   3840
      Top             =   3290
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   31
      Left            =   5840
      Top             =   3630
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   30
      Left            =   5830
      Top             =   3960
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   29
      Left            =   5820
      Top             =   4300
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   28
      Left            =   1200
      Top             =   3270
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   27
      Left            =   4500
      Top             =   4650
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   26
      Left            =   3180
      Top             =   2930
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   25
      Left            =   3200
      Top             =   3290
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   24
      Left            =   2520
      Top             =   2930
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   23
      Left            =   1880
      Top             =   2925
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   22
      Left            =   1870
      Top             =   3280
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   21
      Left            =   3850
      Top             =   3610
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   20
      Left            =   3840
      Top             =   3960
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   19
      Left            =   3840
      Top             =   4320
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   18
      Left            =   3200
      Top             =   3630
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   17
      Left            =   3200
      Top             =   3960
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   16
      Left            =   1200
      Top             =   3610
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   15
      Left            =   1870
      Top             =   3610
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   14
      Left            =   2520
      Top             =   3610
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   13
      Left            =   1880
      Top             =   855
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   12
      Left            =   1200
      Top             =   4640
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   11
      Left            =   1200
      Top             =   4320
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   10
      Left            =   1200
      Top             =   3960
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   9
      Left            =   1870
      Top             =   3960
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   8
      Left            =   2520
      Top             =   3960
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   7
      Left            =   2520
      Top             =   3280
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   6
      Left            =   3190
      Top             =   4320
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   5
      Left            =   3200
      Top             =   4650
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   4
      Left            =   2520
      Top             =   4650
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   3
      Left            =   1870
      Top             =   4650
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   390
      Index           =   2
      Left            =   1870
      Top             =   4290
      Width           =   630
   End
   Begin VB.Image cmdSalir 
      Height          =   300
      Left            =   7080
      Top             =   5160
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   390
      Index           =   1
      Left            =   2520
      Top             =   4220
      Width           =   630
   End
End
Attribute VB_Name = "frmMapa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMDSalir_Click()
Call General_Set_Wav(SND_CLICK)
Unload Me
End Sub

Private Sub Form_Load()
Me.Picture = General_Load_Picture_From_Resource("73.gif")
End Sub
Private Sub Image1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Dim i As Integer
Dim MapName As String
Dim file As String
For i = 1 To 158
    Select Case Index
    
        Case i
            file = Get_Extract(Scripts, "WorldMapData.dat")
            MapName = GetVar(file, "WorldData", Image1(i).Index)
            
            If Not MapName = "" Then
                Label1.Caption = MapName
            Else
                Label1.Caption = "Mapa desconocido"
            End If
            Delete_File file
    End Select
Next i
End Sub
