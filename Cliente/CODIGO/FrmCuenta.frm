VERSION 5.00
Begin VB.Form frmCuenta 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PJ 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1150
      Index           =   1
      Left            =   913
      MouseIcon       =   "FrmCuenta.frx":0000
      Picture         =   "FrmCuenta.frx":144A
      ScaleHeight     =   77
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   49
      TabIndex        =   8
      Top             =   3150
      Width           =   735
   End
   Begin VB.PictureBox PJ 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1150
      Index           =   2
      Left            =   2280
      MouseIcon       =   "FrmCuenta.frx":16E5
      Picture         =   "FrmCuenta.frx":2B2F
      ScaleHeight     =   77
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   49
      TabIndex        =   7
      Top             =   3150
      Width           =   735
   End
   Begin VB.PictureBox PJ 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   1150
      Index           =   3
      Left            =   3645
      MouseIcon       =   "FrmCuenta.frx":2DCA
      Picture         =   "FrmCuenta.frx":4214
      ScaleHeight     =   77
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   49
      TabIndex        =   6
      Top             =   3120
      Width           =   735
   End
   Begin VB.PictureBox PJ 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1150
      Index           =   4
      Left            =   5040
      MouseIcon       =   "FrmCuenta.frx":44AF
      Picture         =   "FrmCuenta.frx":58F9
      ScaleHeight     =   77
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   49
      TabIndex        =   5
      Top             =   3150
      Width           =   735
   End
   Begin VB.PictureBox PJ 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1150
      Index           =   5
      Left            =   913
      MouseIcon       =   "FrmCuenta.frx":5B94
      Picture         =   "FrmCuenta.frx":6FDE
      ScaleHeight     =   77
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   49
      TabIndex        =   4
      Top             =   5805
      Width           =   735
   End
   Begin VB.PictureBox PJ 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1150
      Index           =   8
      Left            =   5055
      MouseIcon       =   "FrmCuenta.frx":7279
      Picture         =   "FrmCuenta.frx":86C3
      ScaleHeight     =   77
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   49
      TabIndex        =   3
      Top             =   5805
      Width           =   735
   End
   Begin VB.PictureBox PJ 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1150
      Index           =   7
      Left            =   3660
      MouseIcon       =   "FrmCuenta.frx":895E
      Picture         =   "FrmCuenta.frx":9DA8
      ScaleHeight     =   77
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   49
      TabIndex        =   2
      Top             =   5805
      Width           =   735
   End
   Begin VB.PictureBox PJ 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1150
      Index           =   6
      Left            =   2280
      MouseIcon       =   "FrmCuenta.frx":A043
      Picture         =   "FrmCuenta.frx":B48D
      ScaleHeight     =   77
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   49
      TabIndex        =   1
      Top             =   5805
      Width           =   735
   End
   Begin VB.Label Logged 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   8
      Left            =   5040
      TabIndex        =   48
      Top             =   7920
      Width           =   855
   End
   Begin VB.Label Logged 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   7
      Left            =   3600
      TabIndex        =   47
      Top             =   7920
      Width           =   855
   End
   Begin VB.Label Logged 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   6
      Left            =   2160
      TabIndex        =   46
      Top             =   7920
      Width           =   855
   End
   Begin VB.Label Logged 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   5
      Left            =   840
      TabIndex        =   45
      Top             =   7920
      Width           =   855
   End
   Begin VB.Label Logged 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   4
      Left            =   4995
      TabIndex        =   44
      Top             =   5400
      Width           =   855
   End
   Begin VB.Label Logged 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   3
      Left            =   3600
      TabIndex        =   43
      Top             =   5400
      Width           =   855
   End
   Begin VB.Label Logged 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   2
      Left            =   2160
      TabIndex        =   42
      Top             =   5400
      Width           =   855
   End
   Begin VB.Label Logged 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   1
      Left            =   840
      TabIndex        =   41
      Top             =   5400
      Width           =   855
   End
   Begin VB.Image Marco 
      Height          =   1470
      Left            =   11640
      Picture         =   "FrmCuenta.frx":B728
      Top             =   9000
      Width           =   1065
   End
   Begin VB.Image Image1 
      Height          =   285
      Left            =   10414
      Top             =   960
      Width           =   1365
   End
   Begin VB.Image cmdcrearpj 
      Height          =   285
      Left            =   10414
      Top             =   1395
      Width           =   1365
   End
   Begin VB.Image cmdborrarpj 
      Height          =   285
      Left            =   10414
      Top             =   1815
      Width           =   1365
   End
   Begin VB.Image cmdcambiarpass 
      Height          =   285
      Left            =   10414
      Top             =   2265
      Width           =   1365
   End
   Begin VB.Image Command1 
      Height          =   285
      Left            =   10414
      Top             =   7725
      Width           =   1365
   End
   Begin VB.Label raza 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   4968
      TabIndex        =   40
      Top             =   7800
      Width           =   855
   End
   Begin VB.Label raza 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   3585
      TabIndex        =   39
      Top             =   7800
      Width           =   855
   End
   Begin VB.Label raza 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   2190
      TabIndex        =   38
      Top             =   7800
      Width           =   855
   End
   Begin VB.Label raza 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   855
      TabIndex        =   37
      Top             =   7800
      Width           =   855
   End
   Begin VB.Label raza 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   4938
      TabIndex        =   36
      Top             =   5160
      Width           =   855
   End
   Begin VB.Label raza 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   3576
      TabIndex        =   35
      Top             =   5160
      Width           =   855
   End
   Begin VB.Label raza 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   2185
      TabIndex        =   34
      Top             =   5160
      Width           =   855
   End
   Begin VB.Label raza 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   808
      TabIndex        =   33
      Top             =   5160
      Width           =   855
   End
   Begin VB.Label nombre 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nada"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   1
      Left            =   1065
      TabIndex        =   32
      Top             =   4500
      Width           =   435
   End
   Begin VB.Label nombre 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nada"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   2
      Left            =   2445
      TabIndex        =   31
      Top             =   4500
      Width           =   435
   End
   Begin VB.Label nombre 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nada"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   3
      Left            =   3825
      TabIndex        =   30
      Top             =   4500
      Width           =   435
   End
   Begin VB.Label nombre 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nada"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   4
      Left            =   5205
      TabIndex        =   29
      Top             =   4500
      Width           =   435
   End
   Begin VB.Label nombre 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nada"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   5
      Left            =   1065
      TabIndex        =   28
      Top             =   7170
      Width           =   435
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   720
      TabIndex        =   27
      Top             =   4680
      Width           =   1035
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   1
      Left            =   1200
      TabIndex        =   26
      Top             =   4920
      Width           =   75
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   2
      Left            =   2125
      TabIndex        =   25
      Top             =   4920
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   3
      Left            =   3480
      TabIndex        =   24
      Top             =   4920
      Width           =   1035
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   4
      Left            =   5340
      TabIndex        =   23
      Top             =   4920
      Width           =   60
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   5
      Left            =   765
      TabIndex        =   22
      Top             =   7560
      Width           =   1035
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   2160
      TabIndex        =   21
      Top             =   4680
      Width           =   915
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   3480
      TabIndex        =   20
      Top             =   4680
      Width           =   1035
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   4
      Left            =   4860
      TabIndex        =   19
      Top             =   4680
      Width           =   1035
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   5
      Left            =   840
      TabIndex        =   18
      Top             =   7350
      Width           =   915
   End
   Begin VB.Label nombre 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nada"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   6
      Left            =   2445
      TabIndex        =   17
      Top             =   7170
      Width           =   435
   End
   Begin VB.Label nombre 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nada"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   7
      Left            =   3855
      TabIndex        =   16
      Top             =   7170
      Width           =   435
   End
   Begin VB.Label nombre 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nada"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   8
      Left            =   5205
      TabIndex        =   15
      Top             =   7170
      Width           =   435
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   6
      Left            =   2160
      TabIndex        =   14
      Top             =   7350
      Width           =   915
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   7
      Left            =   3555
      TabIndex        =   13
      Top             =   7350
      Width           =   900
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   8
      Left            =   4938
      TabIndex        =   12
      Top             =   7350
      Width           =   915
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   6
      Left            =   2160
      TabIndex        =   11
      Top             =   7560
      Width           =   915
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   7
      Left            =   3555
      TabIndex        =   10
      Top             =   7560
      Width           =   900
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   8
      Left            =   4920
      TabIndex        =   9
      Top             =   7560
      Width           =   915
   End
   Begin VB.Label NombreAccount 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   7755
      TabIndex        =   0
      Top             =   2100
      Width           =   75
   End
End
Attribute VB_Name = "frmCuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdborrarpj_Click()
    Call General_Set_Wav(SND_CLICK)

    If PJName = "Nada" Then
    MsgBox "¡No seleccionaste ningun personaje a borrar!", vbCritical, "Winter-AO Ultimate Cuentas"
    Exit Sub
    End If
    
    If frmMain.Winsock1.State <> sckClosed Then
        frmMain.Winsock1.Close
        DoEvents
    End If

    If PJClickeado < 1 Then
        MsgBox "Para borrar un personaje, antes tienes que seleccionarlo"
        Exit Sub
    End If
    
    frmBorrarPj.Show
    Exit Sub
End Sub

Private Sub cmdcambiarpass_Click()
Call General_Set_Wav(SND_CLICK)

If frmMain.Winsock1.State <> sckClosed Then
    frmMain.Winsock1.Close
    DoEvents
End If
    frmNewPassword.Show
    Exit Sub
End Sub
Private Sub cmdcrearpj_Click()
Call General_Set_Wav(SND_CLICK)

        If frmMain.Winsock1.State <> sckClosed Then
            frmMain.Winsock1.Close
            DoEvents
        End If
        
        If Cuenta.CantPJ = 8 Then
            MsgBox "No tienes más espacio para continuar creando personajes."
            Exit Sub
        End If
        
        EstadoLogin = Dados
        
        frmMain.Winsock1.Connect CurServerIp, CurServerPort
        
        If Audio.MusicActivated = True Then
            General_Set_Song 29, True
        End If
        Me.Visible = False
        Exit Sub
End Sub
Private Sub Command1_Click()
Call General_Set_Wav(SND_CLICK)
frmMain.Winsock1.Close
Unload Me
frmConnect.Show
End Sub
Private Sub Form_Load()
Me.Caption = Form_Caption
Me.Picture = General_Load_Picture_From_Resource("46.gif")
NombreAccount.Caption = Cuenta.name
End Sub
Private Sub Image1_Click()
    Call General_Set_Wav(SND_CLICK)
    
    If PJClickeado < 1 Then
        MsgBox "Selecciona antes un personaje."
        Exit Sub
    End If
    
    If frmMain.Winsock1.State <> sckClosed Then
        frmMain.Winsock1.Close
        DoEvents
    End If
            
    EstadoLogin = E_MODO.Normal
            
    frmMain.Winsock1.Connect CurServerIp, CurServerPort
    Exit Sub
End Sub

Private Sub nombre_Click(Index As Integer)
    PJClickeado = Index
    PJName = frmCuenta.Nombre(Index).Caption
End Sub

Private Sub PJ_Click(Index As Integer)
    Call General_Set_Wav(SND_CLICK)
        PJClickeado = Index
        PJName = frmCuenta.Nombre(Index).Caption
        
    Select Case Index
        Case 1
            Marco.Top = 199
            Marco.Left = 50
        Case 2
            Marco.Top = 199
            Marco.Left = 141
        Case 3
            Marco.Top = 199
            Marco.Left = 233
        Case 4
            Marco.Top = 199
            Marco.Left = 326
        Case 5
            Marco.Top = 377
            Marco.Left = 50
        Case 6
            Marco.Top = 377
            Marco.Left = 141
        Case 7
            Marco.Top = 377
            Marco.Left = 233
        Case 8
            Marco.Top = 377
            Marco.Left = 326
    End Select
End Sub
Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image1.Picture = General_Load_Picture_From_Resource("28.gif")
End Sub
Private Sub cmdcrearpj_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdcrearpj.Picture = General_Load_Picture_From_Resource("27.gif")
End Sub
Private Sub cmdborrarpj_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdborrarpj.Picture = General_Load_Picture_From_Resource("17.gif")
End Sub
Private Sub command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Command1.Picture = General_Load_Picture_From_Resource("22.gif")
End Sub
Private Sub cmdcambiarpass_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdcambiarpass.Picture = General_Load_Picture_From_Resource("26.gif")
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image1.Picture = LoadPicture("")
    cmdcrearpj.Picture = LoadPicture("")
    cmdborrarpj.Picture = LoadPicture("")
    Command1.Picture = LoadPicture("")
    cmdcambiarpass.Picture = LoadPicture("")
End Sub

Private Sub PJ_dblClick(Index As Integer)
    Call General_Set_Wav(SND_CLICK)
    
    If frmMain.Winsock1.State <> sckClosed Then
        frmMain.Winsock1.Close
        DoEvents
    End If
    If PJClickeado < 1 Then
        MsgBox "Selecciona antes un personaje."
        Exit Sub
    End If
            
    EstadoLogin = E_MODO.Normal
            
    frmMain.Winsock1.Connect CurServerIp, CurServerPort
    Exit Sub
End Sub
