VERSION 5.00
Begin VB.Form frmCrearPersonaje 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCrearPersonaje.frx":0000
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Cabeza 
      Height          =   315
      Left            =   2400
      TabIndex        =   41
      Top             =   7920
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.PictureBox PlayerView 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2475
      ScaleHeight     =   345
      ScaleWidth      =   795
      TabIndex        =   40
      Top             =   7530
      Width           =   825
   End
   Begin VB.ComboBox lstProfesion 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":4710F
      Left            =   3240
      List            =   "frmCrearPersonaje.frx":47146
      Style           =   2  'Dropdown List
      TabIndex        =   31
      Top             =   6720
      Width           =   2220
   End
   Begin VB.ComboBox lstGenero 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":471E0
      Left            =   3240
      List            =   "frmCrearPersonaje.frx":471EA
      Style           =   2  'Dropdown List
      TabIndex        =   30
      Top             =   5880
      Width           =   2220
   End
   Begin VB.ComboBox lstRaza 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":471FD
      Left            =   3240
      List            =   "frmCrearPersonaje.frx":47213
      Style           =   2  'Dropdown List
      TabIndex        =   29
      Top             =   4920
      Width           =   2220
   End
   Begin VB.ComboBox lstHogar 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":47246
      Left            =   3240
      List            =   "frmCrearPersonaje.frx":4724D
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   3840
      Width           =   2205
   End
   Begin VB.TextBox txtNombre 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Left            =   720
      TabIndex        =   0
      Top             =   2520
      Width           =   4815
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   2160
      Top             =   7560
      Width           =   255
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   3360
      Top             =   7560
      Width           =   255
   End
   Begin VB.Label modCarisma 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2730
      TabIndex        =   39
      Top             =   6720
      Width           =   45
   End
   Begin VB.Label modInteligencia 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2730
      TabIndex        =   38
      Top             =   5370
      Width           =   45
   End
   Begin VB.Label modAgilidad 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2730
      TabIndex        =   37
      Top             =   4680
      Width           =   45
   End
   Begin VB.Label modConstitucion 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2790
      TabIndex        =   36
      Top             =   6030
      Width           =   75
   End
   Begin VB.Label modfuerza 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2790
      TabIndex        =   35
      Top             =   3990
      Width           =   75
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   42
      Left            =   11595
      MouseIcon       =   "frmCrearPersonaje.frx":47257
      MousePointer    =   99  'Custom
      Top             =   6480
      Width           =   285
   End
   Begin VB.Image command1 
      Height          =   255
      Index           =   43
      Left            =   11040
      MouseIcon       =   "frmCrearPersonaje.frx":473A9
      MousePointer    =   99  'Custom
      Top             =   6360
      Width           =   135
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   21
      Left            =   11325
      TabIndex        =   34
      Top             =   6360
      Width           =   270
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "+3"
      ForeColor       =   &H00FFFF80&
      Height          =   195
      Left            =   2760
      TabIndex        =   33
      Top             =   3240
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label Puntos 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   11160
      TabIndex        =   32
      Top             =   7080
      Width           =   270
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   3
      Left            =   8280
      MouseIcon       =   "frmCrearPersonaje.frx":474FB
      MousePointer    =   99  'Custom
      Top             =   2040
      Width           =   255
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   5
      Left            =   8280
      MouseIcon       =   "frmCrearPersonaje.frx":4764D
      MousePointer    =   99  'Custom
      Top             =   2640
      Width           =   255
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   7
      Left            =   8280
      MouseIcon       =   "frmCrearPersonaje.frx":4779F
      MousePointer    =   99  'Custom
      Top             =   3000
      Width           =   255
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   9
      Left            =   8280
      MouseIcon       =   "frmCrearPersonaje.frx":478F1
      MousePointer    =   99  'Custom
      Top             =   3480
      Width           =   255
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   11
      Left            =   8280
      MouseIcon       =   "frmCrearPersonaje.frx":47A43
      MousePointer    =   99  'Custom
      Top             =   4080
      Width           =   255
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   13
      Left            =   8280
      MouseIcon       =   "frmCrearPersonaje.frx":47B95
      MousePointer    =   99  'Custom
      Top             =   4560
      Width           =   255
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   15
      Left            =   8280
      MouseIcon       =   "frmCrearPersonaje.frx":47CE7
      MousePointer    =   99  'Custom
      Top             =   5040
      Width           =   270
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   17
      Left            =   10920
      MouseIcon       =   "frmCrearPersonaje.frx":47E39
      MousePointer    =   99  'Custom
      Top             =   4440
      Width           =   270
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   19
      Left            =   8280
      MouseIcon       =   "frmCrearPersonaje.frx":47F8B
      MousePointer    =   99  'Custom
      Top             =   5640
      Width           =   255
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   21
      Left            =   8280
      MouseIcon       =   "frmCrearPersonaje.frx":480DD
      MousePointer    =   99  'Custom
      Top             =   6120
      Width           =   255
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   23
      Left            =   8280
      MouseIcon       =   "frmCrearPersonaje.frx":4822F
      MousePointer    =   99  'Custom
      Top             =   6600
      Width           =   255
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   25
      Left            =   10920
      MouseIcon       =   "frmCrearPersonaje.frx":48381
      MousePointer    =   99  'Custom
      Top             =   2520
      Width           =   255
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   27
      Left            =   8355
      MouseIcon       =   "frmCrearPersonaje.frx":484D3
      MousePointer    =   99  'Custom
      Top             =   7095
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   1
      Left            =   8280
      MouseIcon       =   "frmCrearPersonaje.frx":48625
      MousePointer    =   99  'Custom
      Top             =   1560
      Width           =   255
   End
   Begin VB.Image command1 
      Height          =   255
      Index           =   0
      Left            =   8880
      MouseIcon       =   "frmCrearPersonaje.frx":48777
      MousePointer    =   99  'Custom
      Top             =   1560
      Width           =   300
   End
   Begin VB.Image command1 
      Height          =   255
      Index           =   2
      Left            =   8880
      MouseIcon       =   "frmCrearPersonaje.frx":488C9
      MousePointer    =   99  'Custom
      Top             =   2040
      Width           =   315
   End
   Begin VB.Image command1 
      Height          =   375
      Index           =   4
      Left            =   8880
      MouseIcon       =   "frmCrearPersonaje.frx":48A1B
      MousePointer    =   99  'Custom
      Top             =   2520
      Width           =   315
   End
   Begin VB.Image command1 
      Height          =   255
      Index           =   6
      Left            =   8880
      MouseIcon       =   "frmCrearPersonaje.frx":48B6D
      MousePointer    =   99  'Custom
      Top             =   3000
      Width           =   300
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   8
      Left            =   8880
      MouseIcon       =   "frmCrearPersonaje.frx":48CBF
      MousePointer    =   99  'Custom
      Top             =   3600
      Width           =   315
   End
   Begin VB.Image command1 
      Height          =   255
      Index           =   10
      Left            =   8880
      MouseIcon       =   "frmCrearPersonaje.frx":48E11
      MousePointer    =   99  'Custom
      Top             =   4080
      Width           =   285
   End
   Begin VB.Image command1 
      Height          =   255
      Index           =   12
      Left            =   8880
      MouseIcon       =   "frmCrearPersonaje.frx":48F63
      MousePointer    =   99  'Custom
      Top             =   4605
      Width           =   285
   End
   Begin VB.Image command1 
      Height          =   240
      Index           =   14
      Left            =   8880
      MouseIcon       =   "frmCrearPersonaje.frx":490B5
      MousePointer    =   99  'Custom
      Top             =   5040
      Width           =   255
   End
   Begin VB.Image command1 
      Height          =   120
      Index           =   16
      Left            =   11520
      MouseIcon       =   "frmCrearPersonaje.frx":49207
      MousePointer    =   99  'Custom
      Top             =   4440
      Width           =   255
   End
   Begin VB.Image command1 
      Height          =   240
      Index           =   18
      Left            =   8880
      MouseIcon       =   "frmCrearPersonaje.frx":49359
      MousePointer    =   99  'Custom
      Top             =   5640
      Width           =   255
   End
   Begin VB.Image command1 
      Height          =   255
      Index           =   20
      Left            =   8880
      MouseIcon       =   "frmCrearPersonaje.frx":494AB
      MousePointer    =   99  'Custom
      Top             =   6120
      Width           =   285
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   22
      Left            =   8880
      MouseIcon       =   "frmCrearPersonaje.frx":495FD
      MousePointer    =   99  'Custom
      Top             =   6600
      Width           =   285
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   24
      Left            =   11520
      MouseIcon       =   "frmCrearPersonaje.frx":4974F
      MousePointer    =   99  'Custom
      Top             =   2520
      Width           =   255
   End
   Begin VB.Image command1 
      Height          =   120
      Index           =   26
      Left            =   9000
      MouseIcon       =   "frmCrearPersonaje.frx":498A1
      MousePointer    =   99  'Custom
      Top             =   7080
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   270
      Index           =   28
      Left            =   8970
      MouseIcon       =   "frmCrearPersonaje.frx":499F3
      MousePointer    =   99  'Custom
      Top             =   7485
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   270
      Index           =   29
      Left            =   8355
      MouseIcon       =   "frmCrearPersonaje.frx":49B45
      MousePointer    =   99  'Custom
      Top             =   7470
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   30
      Left            =   11640
      MouseIcon       =   "frmCrearPersonaje.frx":49C97
      MousePointer    =   99  'Custom
      Top             =   6000
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   150
      Index           =   31
      Left            =   11040
      MouseIcon       =   "frmCrearPersonaje.frx":49DE9
      MousePointer    =   99  'Custom
      Top             =   6000
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   255
      Index           =   32
      Left            =   11610
      MouseIcon       =   "frmCrearPersonaje.frx":49F3B
      MousePointer    =   99  'Custom
      Top             =   5400
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   33
      Left            =   10995
      MouseIcon       =   "frmCrearPersonaje.frx":4A08D
      MousePointer    =   99  'Custom
      Top             =   5400
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   34
      Left            =   11610
      MouseIcon       =   "frmCrearPersonaje.frx":4A1DF
      MousePointer    =   99  'Custom
      Top             =   4905
      Width           =   135
   End
   Begin VB.Image command1 
      Height          =   150
      Index           =   35
      Left            =   10995
      MouseIcon       =   "frmCrearPersonaje.frx":4A331
      MousePointer    =   99  'Custom
      Top             =   4920
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   36
      Left            =   11610
      MouseIcon       =   "frmCrearPersonaje.frx":4A483
      MousePointer    =   99  'Custom
      Top             =   4080
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   37
      Left            =   10995
      MouseIcon       =   "frmCrearPersonaje.frx":4A5D5
      MousePointer    =   99  'Custom
      Top             =   4080
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   120
      Index           =   38
      Left            =   11640
      MouseIcon       =   "frmCrearPersonaje.frx":4A727
      MousePointer    =   99  'Custom
      Top             =   3480
      Width           =   255
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   39
      Left            =   11010
      MouseIcon       =   "frmCrearPersonaje.frx":4A879
      MousePointer    =   99  'Custom
      Top             =   3495
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   255
      Index           =   40
      Left            =   11640
      MouseIcon       =   "frmCrearPersonaje.frx":4A9CB
      MousePointer    =   99  'Custom
      Top             =   2910
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   41
      Left            =   11055
      MouseIcon       =   "frmCrearPersonaje.frx":4AB1D
      MousePointer    =   99  'Custom
      Top             =   3000
      Width           =   135
   End
   Begin VB.Image boton 
      Height          =   645
      Index           =   2
      Left            =   1440
      MouseIcon       =   "frmCrearPersonaje.frx":4AC6F
      MousePointer    =   99  'Custom
      Top             =   6840
      Width           =   780
   End
   Begin VB.Image boton 
      Height          =   735
      Index           =   1
      Left            =   120
      MouseIcon       =   "frmCrearPersonaje.frx":4ADC1
      MousePointer    =   99  'Custom
      Top             =   7560
      Width           =   1725
   End
   Begin VB.Image boton 
      Height          =   570
      Index           =   0
      Left            =   3960
      MouseIcon       =   "frmCrearPersonaje.frx":4AF13
      MousePointer    =   99  'Custom
      Top             =   7560
      Width           =   1800
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   20
      Left            =   11325
      TabIndex        =   28
      Top             =   3000
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   19
      Left            =   11325
      TabIndex        =   27
      Top             =   3480
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   18
      Left            =   11325
      TabIndex        =   26
      Top             =   4095
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   17
      Left            =   11325
      TabIndex        =   25
      Top             =   4920
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   16
      Left            =   11325
      TabIndex        =   24
      Top             =   5400
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   15
      Left            =   11325
      TabIndex        =   23
      Top             =   6000
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   14
      Left            =   8685
      TabIndex        =   22
      Top             =   7560
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   13
      Left            =   8685
      TabIndex        =   21
      Top             =   7080
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   12
      Left            =   11325
      TabIndex        =   20
      Top             =   2520
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   11
      Left            =   8640
      TabIndex        =   19
      Top             =   6600
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   10
      Left            =   8640
      TabIndex        =   18
      Top             =   6120
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   9
      Left            =   8640
      TabIndex        =   17
      Top             =   5640
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   11325
      TabIndex        =   16
      Top             =   4440
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   8640
      TabIndex        =   15
      Top             =   5040
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   8640
      TabIndex        =   14
      Top             =   4560
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   8640
      TabIndex        =   13
      Top             =   4080
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   8640
      TabIndex        =   12
      Top             =   3480
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   8640
      TabIndex        =   11
      Top             =   3000
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   8640
      TabIndex        =   10
      Top             =   2520
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   8640
      TabIndex        =   9
      Top             =   1560
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   8640
      TabIndex        =   8
      Top             =   2040
      Width           =   270
   End
   Begin VB.Label lbCarisma 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2280
      TabIndex        =   6
      Top             =   6720
      Width           =   225
   End
   Begin VB.Label lbSabiduria 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   3120
      TabIndex        =   5
      Top             =   3240
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Label lbInteligencia 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2310
      TabIndex        =   4
      Top             =   5370
      Width           =   210
   End
   Begin VB.Label lbConstitucion 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2310
      TabIndex        =   3
      Top             =   6060
      Width           =   225
   End
   Begin VB.Label lbAgilidad 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2310
      TabIndex        =   2
      Top             =   4680
      Width           =   225
   End
   Begin VB.Label lbFuerza 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2325
      TabIndex        =   1
      Top             =   3990
      Width           =   210
   End
End
Attribute VB_Name = "frmCrearPersonaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public SkillPoints As Byte

Function CheckData() As Boolean
If UserRaza = "" Then
    MsgBox "Seleccione la raza del personaje."
    Exit Function
End If

If UserSexo = "" Then
    MsgBox "Seleccione el sexo del personaje."
    Exit Function
End If

If UserClase = "" Then
    MsgBox "Seleccione la clase del personaje."
    Exit Function
End If

If UserClase = "" Then
    MsgBox "Seleccione la clase del personaje."
    Exit Function
End If

If Cabeza.Text = "" Then
    MsgBox "Seleccione la cabeza de su personaje."
    Exit Function
End If

If SkillPoints > 0 Then
    MsgBox "Asigne los skillpoints del personaje."
    Exit Function
End If

Dim i As Integer
For i = 1 To NUMATRIBUTOS
    If UserAtributos(i) = 0 Then
        MsgBox "Los atributos del personaje son invalidos."
        Exit Function
    End If
Next i

CheckData = True


End Function

Private Sub boton_Click(Index As Integer)
Call Audio.PlayWave(SND_CLICK)
Select Case Index
    Case 0
        
        Dim i As Integer
        Dim k As Object
        i = 1
        For Each k In Skill
            UserSkills(i) = k.Caption
            i = i + 1
        Next
        
        UserName = txtNombre.Text
        
        If Len(txtNombre.Text) >= 11 Then
    MsgBox "¡¡ El nombre no puede superar los 11 caracteres !!"
    Exit Sub
End If
        If Len(txtNombre.Text) < 3 Then
    MsgBox "¡¡ El nombre debe de tener almenos 3 caracteres o mas !!"
    Exit Sub
End If
        If Right$(UserName, 1) = " " Then
                UserName = RTrim$(UserName)
                MsgBox "Nombre invalido, se han removido los espacios al final del nombre"
        End If
        
        UserRaza = lstRaza.List(lstRaza.listIndex)
        UserSexo = lstGenero.List(lstGenero.listIndex)
        UserClase = lstProfesion.List(lstProfesion.listIndex)
        
        UserAtributos(1) = Val(lbFuerza.Caption)
        UserAtributos(2) = Val(lbInteligencia.Caption)
        UserAtributos(3) = Val(lbAgilidad.Caption)
        UserAtributos(4) = Val(lbCarisma.Caption)
        UserAtributos(5) = Val(lbConstitucion.Caption)
        
        UserHogar = lstHogar.List(lstHogar.listIndex)
        
        'Barrin 3/10/03
        If CheckData() Then
            frmPasswdSinPadrinos.Show vbModal, Me
        End If
        
    Case 1
        If Musica Then
Call Extract_File2(Midi, App.Path & "\ARCHIVOS", "2.mid", Windows_Temp_Dir, False)
            Call Audio.PlayMIDI("2.mid")
            Delete_File (Windows_Temp_Dir & CStr(currentMidi) & ".mid")
        End If
        
        Me.Picture = General_Load_Picture_From_Resource("Conectar.gif")
        Me.Visible = False
        
        
    Case 2

        Call Audio.PlayWave(SND_DICE)

        Call TirarDados
      
End Select


End Sub


Function RandomNumber(ByVal LowerBound As Variant, ByVal UpperBound As Variant) As Single

Randomize timer

RandomNumber = (UpperBound - LowerBound + 1) * Rnd + LowerBound
If RandomNumber > UpperBound Then RandomNumber = UpperBound

End Function


Private Sub TirarDados()

        Call SendData("TIRDAD")


End Sub

Private Sub Command1_Click(Index As Integer)

Call Audio.PlayWave(SND_CLICK)


Dim indice
If Index Mod 2 = 0 Then
    If SkillPoints > 0 Then
        indice = Index \ 2
        Skill(indice).Caption = Val(Skill(indice).Caption) + 1
        SkillPoints = SkillPoints - 1
    End If
Else
    If SkillPoints < 10 Then
        
        indice = Index \ 2
        If Val(Skill(indice).Caption) > 0 Then
            Skill(indice).Caption = Val(Skill(indice).Caption) - 1
            SkillPoints = SkillPoints + 1
        End If
    End If
End If

puntos.Caption = SkillPoints
End Sub

Private Sub Form_Load()
SkillPoints = 10
puntos.Caption = SkillPoints
Me.Picture = General_Load_Picture_From_Resource("CP-Interface.gif")


Dim i As Integer
lstProfesion.Clear
For i = LBound(ListaClases) To UBound(ListaClases)
    lstProfesion.AddItem ListaClases(i)
Next i

lstProfesion.listIndex = 1

Call TirarDados
End Sub


Private Sub Image1_Click()
If Cabeza.listIndex < Cabeza.ListCount - 1 Then Cabeza.listIndex = Cabeza.listIndex + 1
End Sub

Private Sub Image2_Click()
If Cabeza.listIndex > 0 Then Cabeza.listIndex = Cabeza.listIndex - 1
End Sub
Private Sub txtNombre_Change()
txtNombre.Text = LTrim(txtNombre.Text)
End Sub
Private Sub txtnombre_click()
Call frmReglamentoName.Show(vbModal, frmCrearPersonaje)
End Sub

Private Sub cabeza_Click()
MiCabeza = Val(Cabeza.List(Cabeza.listIndex))
Call DibujarCPJ(MiCuerpo, MiCabeza)
End Sub

Private Sub lstGenero_Click()
Call DameOpciones
End Sub

Private Sub lstRaza_Click()
Call DameOpciones

Select Case (lstRaza.List(lstRaza.listIndex))
    Case Is = "Humano"
        modfuerza.Caption = "+ 2"
        modConstitucion.Caption = "+ 2"
        modAgilidad.Caption = ""
        modInteligencia.Caption = "+ 1"
        modCarisma.Caption = ""
    Case Is = "Elfo"
        modfuerza.Caption = "- 1"
        modConstitucion.Caption = ""
        modAgilidad.Caption = "+ 2"
        modInteligencia.Caption = "+ 3"
        modCarisma.Caption = "+ 2"
    Case Is = "Elfo Oscuro"
        modfuerza.Caption = "+ 1"
        modConstitucion.Caption = "+ 1"
        modAgilidad.Caption = "+ 1"
        modInteligencia.Caption = "+ 2"
        modCarisma.Caption = "+ 1"
    Case Is = "Enano"
        modfuerza.Caption = "+ 3"
        modConstitucion.Caption = "+ 4"
        modAgilidad.Caption = "- 1"
        modInteligencia.Caption = "- 7"
        modCarisma.Caption = "- 1"
    Case Is = "Orco"
        modfuerza.Caption = "+ 5"
        modConstitucion.Caption = "+ 3"
        modAgilidad.Caption = "- 2"
        modInteligencia.Caption = "- 6"
        modCarisma.Caption = "- 2"
    Case Is = "Gnomo"
        modfuerza.Caption = "- 4"
        modAgilidad.Caption = "+ 3"
        modInteligencia.Caption = "+ 4"
        modCarisma.Caption = "+ 1"
        modConstitucion.Caption = "- 1"
End Select
End Sub
