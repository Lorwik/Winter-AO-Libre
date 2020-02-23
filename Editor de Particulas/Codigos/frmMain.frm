VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Particle Editor - RincondelAO.com.ar - By Lorwik"
   ClientHeight    =   9705
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   11490
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9705
   ScaleWidth      =   11490
   StartUpPosition =   1  'CenterOwner
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmdOpenStreamFile 
      Caption         =   "&Open Stream File"
      Height          =   255
      Left            =   240
      TabIndex        =   99
      Top             =   600
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Frame Frame4 
      Caption         =   "Lista de particulas"
      Height          =   4455
      Left            =   0
      TabIndex        =   97
      Top             =   120
      Width           =   2415
      Begin VB.ListBox List2 
         BackColor       =   &H00C0C0C0&
         Height          =   4155
         ItemData        =   "frmMain.frx":000C
         Left            =   45
         List            =   "frmMain.frx":000E
         TabIndex        =   98
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Visor de grh"
      Height          =   2775
      Left            =   7920
      TabIndex        =   96
      Top             =   6600
      Width           =   3495
      Begin VB.PictureBox invpic 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         Height          =   2415
         Left            =   120
         ScaleHeight     =   2415
         ScaleWidth      =   3255
         TabIndex        =   4
         Top             =   240
         Width           =   3255
      End
   End
   Begin VB.Frame frameGrhs 
      Caption         =   "Grh Parameters"
      Height          =   6075
      Left            =   9480
      TabIndex        =   88
      Top             =   0
      Width           =   2010
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Borrar"
         Height          =   255
         Left            =   600
         TabIndex        =   93
         Top             =   3840
         Width           =   615
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Añadir"
         Height          =   255
         Left            =   0
         TabIndex        =   92
         Top             =   3840
         Width           =   615
      End
      Begin VB.ListBox lstSelGrhs 
         BackColor       =   &H00C0C0C0&
         Height          =   1620
         Left            =   120
         TabIndex        =   91
         Top             =   4320
         Width           =   1770
      End
      Begin VB.ListBox lstGrhs 
         BackColor       =   &H00C0C0C0&
         Height          =   3180
         Left            =   120
         TabIndex        =   90
         Top             =   480
         Width           =   1740
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Limpiar"
         Height          =   255
         Left            =   1200
         TabIndex        =   89
         Top             =   3840
         Width           =   615
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Grhs Seleccionados"
         Height          =   195
         Left            =   240
         TabIndex        =   95
         Top             =   4080
         Width           =   1425
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lista de Grh"
         Height          =   195
         Left            =   60
         TabIndex        =   94
         Top             =   255
         Width           =   855
      End
   End
   Begin VB.Frame frameGravity 
      BorderStyle     =   0  'None
      Caption         =   "Gravity Settings"
      Height          =   1095
      Left            =   330
      TabIndex        =   81
      Top             =   7110
      Width           =   1935
      Begin VB.TextBox txtGravStrength 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1440
         MaxLength       =   4
         TabIndex        =   84
         Text            =   "5"
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox txtBounceStrength 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1440
         MaxLength       =   4
         TabIndex        =   83
         Text            =   "1"
         Top             =   720
         Width           =   375
      End
      Begin VB.CheckBox chkGravity 
         Caption         =   "Gravity Influence"
         Height          =   255
         Left            =   120
         TabIndex        =   82
         Top             =   180
         Width           =   1575
      End
      Begin VB.Label Label64 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gravity Strength:"
         Height          =   195
         Left            =   120
         TabIndex        =   86
         Top             =   465
         Width           =   1185
      End
      Begin VB.Label Label65 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bounce Strength:"
         Height          =   195
         Left            =   120
         TabIndex        =   85
         Top             =   705
         Width           =   1245
      End
   End
   Begin VB.Frame frameMovement 
      BorderStyle     =   0  'None
      Caption         =   "Movement Settings"
      Height          =   1935
      Left            =   315
      TabIndex        =   70
      Top             =   7095
      Width           =   1935
      Begin VB.TextBox move_x1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1440
         MaxLength       =   4
         TabIndex        =   76
         Text            =   "0"
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox move_x2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1440
         MaxLength       =   4
         TabIndex        =   75
         Text            =   "0"
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox move_y1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1440
         MaxLength       =   4
         TabIndex        =   74
         Text            =   "0"
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox move_y2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1440
         MaxLength       =   4
         TabIndex        =   73
         Text            =   "0"
         Top             =   1560
         Width           =   375
      End
      Begin VB.CheckBox chkYMove 
         Caption         =   "Y Movement"
         Height          =   255
         Left            =   120
         TabIndex        =   72
         Top             =   1080
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CheckBox chkXMove 
         Caption         =   "X Movement"
         Height          =   255
         Left            =   120
         TabIndex        =   71
         Top             =   240
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.Label Label60 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Movement Y2:"
         Height          =   195
         Left            =   120
         TabIndex        =   80
         Top             =   1605
         Width           =   1035
      End
      Begin VB.Label Label61 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Movement Y1:"
         Height          =   195
         Left            =   120
         TabIndex        =   79
         Top             =   1365
         Width           =   1035
      End
      Begin VB.Label Label62 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Movement X2:"
         Height          =   195
         Left            =   120
         TabIndex        =   78
         Top             =   765
         Width           =   1035
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Movement X1:"
         Height          =   195
         Left            =   120
         TabIndex        =   77
         Top             =   525
         Width           =   1035
      End
   End
   Begin VB.Frame frameSpinSettings 
      BorderStyle     =   0  'None
      Caption         =   "Spin Settings"
      Height          =   1095
      Left            =   345
      TabIndex        =   64
      Top             =   7095
      Width           =   1935
      Begin VB.TextBox spin_speedL 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   67
         Text            =   "1"
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox spin_speedH 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   66
         Text            =   "1"
         Top             =   720
         Width           =   495
      End
      Begin VB.CheckBox chkSpin 
         Caption         =   "Spin"
         Height          =   255
         Left            =   105
         TabIndex        =   65
         Top             =   240
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.Label Label58 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Spin Speed (L):"
         Height          =   195
         Left            =   120
         TabIndex        =   69
         Top             =   525
         Width           =   1095
      End
      Begin VB.Label Label59 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Spin Speed (H):"
         Height          =   195
         Left            =   120
         TabIndex        =   68
         Top             =   765
         Width           =   1125
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Particle Duration"
      Height          =   855
      Left            =   330
      TabIndex        =   60
      Top             =   7125
      Width           =   1935
      Begin VB.CheckBox chkNeverDies 
         Caption         =   "Never Dies"
         Height          =   255
         Left            =   120
         TabIndex        =   62
         Top             =   240
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.TextBox life 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   61
         Text            =   "10"
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label57 
         AutoSize        =   -1  'True
         Caption         =   "Life:"
         Height          =   195
         Left            =   120
         TabIndex        =   63
         Top             =   525
         Width           =   300
      End
   End
   Begin VB.Frame frmSettings 
      BorderStyle     =   0  'None
      Height          =   2190
      Left            =   960
      TabIndex        =   27
      Top             =   7080
      Width           =   6600
      Begin VB.TextBox txRad 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5115
         TabIndex        =   102
         Text            =   "0"
         Top             =   1200
         Width           =   495
      End
      Begin VB.TextBox txtry 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3150
         MaxLength       =   4
         TabIndex        =   44
         Text            =   "0"
         Top             =   1635
         Width           =   495
      End
      Begin VB.CheckBox chkresize 
         Caption         =   "Resize"
         Height          =   195
         Left            =   1920
         TabIndex        =   43
         Top             =   1920
         Width           =   1245
      End
      Begin VB.CheckBox chkAlphaBlend 
         Caption         =   "Alpha Blend"
         Height          =   255
         Left            =   3930
         TabIndex        =   42
         Top             =   1560
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.TextBox fric 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5115
         MaxLength       =   4
         TabIndex        =   41
         Text            =   "5"
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox life2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5115
         MaxLength       =   4
         TabIndex        =   40
         Text            =   "50"
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox life1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5115
         MaxLength       =   4
         TabIndex        =   39
         Text            =   "10"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox vecy2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3150
         MaxLength       =   4
         TabIndex        =   38
         Text            =   "0"
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox vecy1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3150
         MaxLength       =   4
         TabIndex        =   37
         Text            =   "-50"
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox vecx2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3150
         MaxLength       =   4
         TabIndex        =   36
         Text            =   "10"
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox vecx1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3150
         MaxLength       =   4
         TabIndex        =   35
         Text            =   "-10"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtAngle 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   34
         Text            =   "0"
         Top             =   1605
         Width           =   495
      End
      Begin VB.TextBox txtY2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   33
         Text            =   "0"
         Top             =   1200
         Width           =   495
      End
      Begin VB.TextBox txtY1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   32
         Text            =   "0"
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox txtX2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   31
         Text            =   "0"
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox txtX1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   30
         Text            =   "0"
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox txtPCount 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   29
         Text            =   "20"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtrx 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3150
         MaxLength       =   4
         TabIndex        =   28
         Text            =   "0"
         Top             =   1395
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Radio:"
         Height          =   255
         Left            =   3915
         TabIndex        =   103
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Resize X:"
         Height          =   195
         Left            =   1950
         TabIndex        =   59
         Top             =   1440
         Width           =   675
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Resize Y:"
         Height          =   195
         Left            =   1950
         TabIndex        =   58
         Top             =   1680
         Width           =   675
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "X2:"
         Height          =   195
         Left            =   120
         TabIndex        =   57
         Top             =   765
         Width           =   240
      End
      Begin VB.Label Label46 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "X1:"
         Height          =   195
         Left            =   120
         TabIndex        =   56
         Top             =   525
         Width           =   240
      End
      Begin VB.Label lblPCount 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "# of Particles:"
         Height          =   195
         Left            =   120
         TabIndex        =   55
         Top             =   285
         Width           =   975
      End
      Begin VB.Label Label47 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Y1:"
         Height          =   195
         Left            =   120
         TabIndex        =   54
         Top             =   1005
         Width           =   240
      End
      Begin VB.Label Label48 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Y2:"
         Height          =   195
         Left            =   120
         TabIndex        =   53
         Top             =   1245
         Width           =   240
      End
      Begin VB.Label Label49 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Friction:"
         Height          =   195
         Left            =   3915
         TabIndex        =   52
         Top             =   885
         Width           =   555
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Life Range (H):"
         Height          =   195
         Left            =   3915
         TabIndex        =   51
         Top             =   525
         Width           =   1080
      End
      Begin VB.Label Label51 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Life Range (L):"
         Height          =   195
         Left            =   3915
         TabIndex        =   50
         Top             =   285
         Width           =   1050
      End
      Begin VB.Label Label52 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vector Y2"
         Height          =   195
         Left            =   1950
         TabIndex        =   49
         Top             =   1005
         Width           =   705
      End
      Begin VB.Label Label53 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vector Y1:"
         Height          =   195
         Left            =   1950
         TabIndex        =   48
         Top             =   765
         Width           =   750
      End
      Begin VB.Label Label54 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vector X2:"
         Height          =   195
         Left            =   1950
         TabIndex        =   47
         Top             =   525
         Width           =   750
      End
      Begin VB.Label Label55 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vector X1:"
         Height          =   195
         Left            =   1950
         TabIndex        =   46
         Top             =   285
         Width           =   750
      End
      Begin VB.Label Label56 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Angle:"
         Height          =   195
         Left            =   120
         TabIndex        =   45
         Top             =   1650
         Width           =   450
      End
   End
   Begin VB.Frame frameColorSettings 
      BorderStyle     =   0  'None
      Caption         =   "Color Tint Settings"
      Height          =   2175
      Left            =   375
      TabIndex        =   15
      Top             =   7035
      Width           =   3975
      Begin VB.HScrollBar RScroll 
         Height          =   255
         Left            =   360
         Max             =   255
         TabIndex        =   23
         Top             =   1800
         Width           =   3015
      End
      Begin VB.HScrollBar GScroll 
         Height          =   255
         Left            =   360
         Max             =   255
         TabIndex        =   22
         Top             =   1500
         Width           =   3015
      End
      Begin VB.HScrollBar BScroll 
         Height          =   255
         Left            =   360
         Max             =   255
         TabIndex        =   21
         Top             =   1200
         Width           =   3015
      End
      Begin VB.ListBox lstColorSets 
         Height          =   840
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   1215
      End
      Begin VB.PictureBox picColor 
         BackColor       =   &H00FFFFFF&
         Height          =   855
         Left            =   1440
         ScaleHeight     =   795
         ScaleWidth      =   2355
         TabIndex        =   19
         Top             =   240
         Width           =   2415
      End
      Begin VB.TextBox txtR 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3480
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   18
         Text            =   "0"
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox txtG 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3480
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   17
         Text            =   "0"
         Top             =   1500
         Width           =   375
      End
      Begin VB.TextBox txtB 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3480
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   16
         Text            =   "0"
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "R:"
         Height          =   195
         Left            =   120
         TabIndex        =   26
         Top             =   1200
         Width           =   165
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "G:"
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   1500
         Width           =   165
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "B:"
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   1800
         Width           =   150
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Particle Speed"
      Height          =   855
      Left            =   435
      TabIndex        =   12
      Top             =   7170
      Width           =   1935
      Begin VB.TextBox speed 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         TabIndex        =   13
         Text            =   "0.5"
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         Caption         =   "Render Delay:"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   1020
      End
   End
   Begin VB.Frame frmfade 
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   2235
      Left            =   120
      TabIndex        =   6
      Top             =   7080
      Width           =   7680
      Begin VB.TextBox txtfout 
         Height          =   300
         Left            =   1320
         TabIndex        =   8
         Text            =   "0"
         Top             =   405
         Width           =   645
      End
      Begin VB.TextBox txtfin 
         Height          =   285
         Left            =   1320
         TabIndex        =   7
         Text            =   "0"
         Top             =   90
         Width           =   630
      End
      Begin VB.Label Label36 
         Caption         =   "Note: The time a particle remains alive is set in the Duration Tab"
         Height          =   585
         Left            =   90
         TabIndex        =   11
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label37 
         Caption         =   "Fade out time"
         Height          =   300
         Left            =   60
         TabIndex        =   10
         Top             =   405
         Width           =   1215
      End
      Begin VB.Label Label38 
         Caption         =   "Fade in time"
         Height          =   180
         Left            =   60
         TabIndex        =   9
         Top             =   120
         Width           =   1245
      End
   End
   Begin VB.PictureBox renderer 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6525
      Left            =   2400
      ScaleHeight     =   435
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   469
      TabIndex        =   5
      Top             =   0
      Width           =   7035
      Begin MSComDlg.CommonDialog ComDlg 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Desaparecer"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   6120
      Width           =   2175
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Nueva Particula"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   4680
      Width           =   2175
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Guardar Particula"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   5160
      Width           =   2175
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Vista Previa"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   5640
      Width           =   2175
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   2670
      Left            =   0
      TabIndex        =   87
      Top             =   6720
      Width           =   7845
      _ExtentX        =   13838
      _ExtentY        =   4710
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   8
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Configuracion de Particula"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Gravedad"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Movimiento"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Vueltas"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Velocidad"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Duracion"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Color "
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab8 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Fade"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Adaptado y Traducido por Lorwik www.RincondelAO.com.ar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   101
      Top             =   9360
      Width           =   11415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "FPS:"
      Height          =   375
      Left            =   9480
      TabIndex        =   100
      Top             =   6240
      Width           =   1695
   End
   Begin VB.Menu mnuarchivo 
      Caption         =   "&Archivo"
      Begin VB.Menu mnunueva 
         Caption         =   "&Nueva particula"
      End
      Begin VB.Menu mnuabrir 
         Caption         =   "&Abrir archivo.."
      End
      Begin VB.Menu mnuguardar 
         Caption         =   "&Guardar"
      End
      Begin VB.Menu mnusalir 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu Mas 
      Caption         =   "Más"
      Begin VB.Menu sobre 
         Caption         =   "Sobre..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
'--> Current Stream File <--
Public CurStreamFile As String

Private Sub Command4_Click()
Dim loopc As Long
Dim StreamFile As String
Dim Bypass As Boolean
Dim retval
CurStreamFile = App.Path & "\INIT\Particles.ini"

If General_File_Exists(CurStreamFile, vbNormal) = True Then
    retval = MsgBox("¡El archivo " & CurStreamFile & " ya existe!" & vbCrLf & "¿Deseas sobreescribirlo?", vbYesNoCancel Or vbQuestion)
    If retval = vbNo Then
        Bypass = False
    ElseIf retval = vbCancel Then
        Exit Sub
    ElseIf retval = vbYes Then
        StreamFile = CurStreamFile
        Bypass = True
    End If
End If

If Bypass = False Then
    With ComDlg
        .Filter = "*.ini (Stream Data Files)|*.ini"
        .ShowSave
        StreamFile = .FileName
    End With
    
    If General_File_Exists(StreamFile, vbNormal) = True Then
        retval = MsgBox("¡El archivo " & StreamFile & " ya existe!" & vbCrLf & "¿Desea sobreescribirlo?", vbYesNo Or vbQuestion)
        If retval = vbNo Then
            Exit Sub
        End If
    End If
End If

Dim GrhListing As String
Dim i As Long

'Check for existing data file and kill it
If General_File_Exists(StreamFile, vbNormal) Then Kill StreamFile

'Write particle data to Particles.ini
General_Var_Write StreamFile, "INIT", "Total", Val(TotalStreams)

For loopc = 1 To TotalStreams
    General_Var_Write StreamFile, Val(loopc), "Name", StreamData(loopc).Name
    General_Var_Write StreamFile, Val(loopc), "NumOfParticles", Val(StreamData(loopc).NumOfParticles)
    General_Var_Write StreamFile, Val(loopc), "X1", Val(StreamData(loopc).x1)
    General_Var_Write StreamFile, Val(loopc), "Y1", Val(StreamData(loopc).y1)
    General_Var_Write StreamFile, Val(loopc), "X2", Val(StreamData(loopc).x2)
    General_Var_Write StreamFile, Val(loopc), "Y2", Val(StreamData(loopc).y2)
    General_Var_Write StreamFile, Val(loopc), "Angle", Val(StreamData(loopc).angle)
    General_Var_Write StreamFile, Val(loopc), "VecX1", Val(StreamData(loopc).vecx1)
    General_Var_Write StreamFile, Val(loopc), "VecX2", Val(StreamData(loopc).vecx2)
    General_Var_Write StreamFile, Val(loopc), "VecY1", Val(StreamData(loopc).vecy1)
    General_Var_Write StreamFile, Val(loopc), "VecY2", Val(StreamData(loopc).vecy2)
    General_Var_Write StreamFile, Val(loopc), "Life1", Val(StreamData(loopc).life1)
    General_Var_Write StreamFile, Val(loopc), "Life2", Val(StreamData(loopc).life2)
    General_Var_Write StreamFile, Val(loopc), "Friction", Val(StreamData(loopc).friction)
    General_Var_Write StreamFile, Val(loopc), "Spin", Val(StreamData(loopc).spin)
    General_Var_Write StreamFile, Val(loopc), "Spin_SpeedL", Val(StreamData(loopc).spin_speedL)
    General_Var_Write StreamFile, Val(loopc), "Spin_SpeedH", Val(StreamData(loopc).spin_speedH)
    General_Var_Write StreamFile, Val(loopc), "Grav_Strength", Val(StreamData(loopc).grav_strength)
    General_Var_Write StreamFile, Val(loopc), "Bounce_Strength", Val(StreamData(loopc).bounce_strength)
    
    General_Var_Write StreamFile, Val(loopc), "AlphaBlend", Val(StreamData(loopc).AlphaBlend)
    General_Var_Write StreamFile, Val(loopc), "Gravity", Val(StreamData(loopc).gravity)
    
    General_Var_Write StreamFile, Val(loopc), "XMove", Val(StreamData(loopc).XMove)
    General_Var_Write StreamFile, Val(loopc), "YMove", Val(StreamData(loopc).YMove)
    General_Var_Write StreamFile, Val(loopc), "move_x1", Val(StreamData(loopc).move_x1)
    General_Var_Write StreamFile, Val(loopc), "move_x2", Val(StreamData(loopc).move_x2)
    General_Var_Write StreamFile, Val(loopc), "move_y1", Val(StreamData(loopc).move_y1)
    General_Var_Write StreamFile, Val(loopc), "move_y2", Val(StreamData(loopc).move_y2)
    General_Var_Write StreamFile, Val(loopc), "Radio", Val(StreamData(loopc).Radio)
    General_Var_Write StreamFile, Val(loopc), "life_counter", Val(StreamData(loopc).life_counter)
    General_Var_Write StreamFile, Val(loopc), "Speed", Str(StreamData(loopc).speed)
    
    General_Var_Write StreamFile, Val(loopc), "resize", CInt(StreamData(loopc).grh_resize)
    General_Var_Write StreamFile, Val(loopc), "rx", StreamData(loopc).grh_resizex
    General_Var_Write StreamFile, Val(loopc), "ry", StreamData(loopc).grh_resizey
    
    General_Var_Write StreamFile, Val(loopc), "NumGrhs", Val(StreamData(loopc).NumGrhs)
    
    GrhListing = vbNullString
    For i = 1 To StreamData(loopc).NumGrhs
        GrhListing = GrhListing & StreamData(loopc).grh_list(i) & ","
    Next i
    
    General_Var_Write StreamFile, Val(loopc), "Grh_List", GrhListing
    
    General_Var_Write StreamFile, Val(loopc), "ColorSet1", StreamData(loopc).colortint(0).r & "," & StreamData(loopc).colortint(0).g & "," & StreamData(loopc).colortint(0).B
    General_Var_Write StreamFile, Val(loopc), "ColorSet2", StreamData(loopc).colortint(1).r & "," & StreamData(loopc).colortint(1).g & "," & StreamData(loopc).colortint(1).B
    General_Var_Write StreamFile, Val(loopc), "ColorSet3", StreamData(loopc).colortint(2).r & "," & StreamData(loopc).colortint(2).g & "," & StreamData(loopc).colortint(2).B
    General_Var_Write StreamFile, Val(loopc), "ColorSet4", StreamData(loopc).colortint(3).r & "," & StreamData(loopc).colortint(3).g & "," & StreamData(loopc).colortint(3).B
    
Next loopc

'Report the results
If TotalStreams > 1 Then
    MsgBox TotalStreams & " Particulas guardadas en: " & vbCrLf & StreamFile, vbInformation
Else
    MsgBox TotalStreams & " Particulas guardadas en: " & vbCrLf & StreamFile, vbInformation
End If

'Set DataChanged variable to false
DataChanged = False
CurStreamFile = StreamFile
End Sub

Private Sub Command5_Click()
Dim Nombre As String
Dim NewStreamNumber As Integer
Dim grhlist(0) As Long

'Get name for new stream
Nombre = InputBox("Por favor inserte un nombre a la particula", "New Stream")

If Nombre = "" Then Exit Sub

'Set new stream #
NewStreamNumber = List2.ListCount + 1

'Add stream to combo box
List2.AddItem Nombre

'Add 1 to TotalStreams
TotalStreams = TotalStreams + 1

grhlist(0) = 19751
'Add stream data to StreamData array
StreamData(NewStreamNumber).Name = Nombre
StreamData(NewStreamNumber).NumOfParticles = 20
StreamData(NewStreamNumber).x1 = 0
StreamData(NewStreamNumber).y1 = 0
StreamData(NewStreamNumber).x2 = 0
StreamData(NewStreamNumber).y2 = 0
StreamData(NewStreamNumber).angle = 0
StreamData(NewStreamNumber).vecx1 = -20
StreamData(NewStreamNumber).vecx2 = 20
StreamData(NewStreamNumber).vecy1 = -20
StreamData(NewStreamNumber).vecy2 = 20
StreamData(NewStreamNumber).life1 = 10
StreamData(NewStreamNumber).life2 = 50
StreamData(NewStreamNumber).friction = 8
StreamData(NewStreamNumber).spin_speedL = 0.1
StreamData(NewStreamNumber).spin_speedH = 0.1
StreamData(NewStreamNumber).grav_strength = 2
StreamData(NewStreamNumber).bounce_strength = -5
StreamData(NewStreamNumber).speed = 0.5
StreamData(NewStreamNumber).AlphaBlend = 1
StreamData(NewStreamNumber).gravity = 0
StreamData(NewStreamNumber).XMove = 0
StreamData(NewStreamNumber).YMove = 0
StreamData(NewStreamNumber).move_x1 = 0
StreamData(NewStreamNumber).move_x2 = 0
StreamData(NewStreamNumber).move_y1 = 0
StreamData(NewStreamNumber).move_y2 = 0
StreamData(NewStreamNumber).life_counter = -1
StreamData(NewStreamNumber).NumGrhs = 1
StreamData(NewStreamNumber).grh_list = grhlist()

'Select the new stream type in the combo box
List2.ListIndex = NewStreamNumber - 1

End Sub

Private Sub Command6_Click()
If List2.ListIndex < 0 Then Exit Sub
Call CargarParticulasLista
End Sub

Private Sub Command8_Click()
engine.Particle_Group_Remove_All
End Sub

Private Sub Form_Load()
       lstColorSets.AddItem "Bottom Left"
    lstColorSets.AddItem "Top Left"
    lstColorSets.AddItem "Bottom Right"
    lstColorSets.AddItem "Top Right"
    frmSettings.Visible = True
    frmfade.Visible = False
    frameColorSettings.Visible = False
    Frame2.Visible = False
    Frame1.Visible = False
    frameSpinSettings.Visible = False
    frameMovement.Visible = False
    frameGravity.Visible = False
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    HookSurfaceHwnd Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
End
End Sub

Private Sub Form_Terminate()
End
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Label2_Click()
End
End Sub



Private Sub List2_Click()
Call CargarParticulasLista
End Sub
Sub CargarParticulasLista()
Dim loopc As Long
Dim DataTemp As Boolean
DataTemp = DataChanged

'Set the values
txtPCount.Text = StreamData(List2.ListIndex + 1).NumOfParticles
txtX1.Text = StreamData(List2.ListIndex + 1).x1
txtY1.Text = StreamData(List2.ListIndex + 1).y1
txtX2.Text = StreamData(List2.ListIndex + 1).x2
txtY2.Text = StreamData(List2.ListIndex + 1).y2
txtAngle.Text = StreamData(List2.ListIndex + 1).angle
vecx1.Text = StreamData(List2.ListIndex + 1).vecx1
vecx2.Text = StreamData(List2.ListIndex + 1).vecx2
vecy1.Text = StreamData(List2.ListIndex + 1).vecy1
vecy2.Text = StreamData(List2.ListIndex + 1).vecy2
life1.Text = StreamData(List2.ListIndex + 1).life1
life2.Text = StreamData(List2.ListIndex + 1).life2
fric.Text = StreamData(List2.ListIndex + 1).friction
chkSpin.value = StreamData(List2.ListIndex + 1).spin
spin_speedL.Text = StreamData(List2.ListIndex + 1).spin_speedL
spin_speedH.Text = StreamData(List2.ListIndex + 1).spin_speedH
txtGravStrength.Text = StreamData(List2.ListIndex + 1).grav_strength
txtBounceStrength.Text = StreamData(List2.ListIndex + 1).bounce_strength
chkAlphaBlend.value = StreamData(List2.ListIndex + 1).AlphaBlend
chkGravity.value = StreamData(List2.ListIndex + 1).gravity
txtrx.Text = StreamData(List2.ListIndex + 1).grh_resizex
txtry.Text = StreamData(List2.ListIndex + 1).grh_resizey
chkXMove.value = StreamData(List2.ListIndex + 1).XMove
chkYMove.value = StreamData(List2.ListIndex + 1).YMove
move_x1.Text = StreamData(List2.ListIndex + 1).move_x1
move_x2.Text = StreamData(List2.ListIndex + 1).move_x2
move_y1.Text = StreamData(List2.ListIndex + 1).move_y1
move_y2.Text = StreamData(List2.ListIndex + 1).move_y2
txRad.Text = StreamData(List2.ListIndex + 1).Radio

If StreamData(List2.ListIndex + 1).grh_resize = True Then
    chkresize = vbChecked
Else
    chkresize = vbUnchecked
End If

If StreamData(List2.ListIndex + 1).life_counter = -1 Then
    life.Enabled = False
    chkNeverDies.value = vbChecked
Else
    life.Enabled = True
    life.Text = StreamData(List2.ListIndex + 1).life_counter
    chkNeverDies.value = vbUnchecked
End If

speed.Text = StreamData(List2.ListIndex + 1).speed

lstSelGrhs.Clear

For loopc = 1 To StreamData(List2.ListIndex + 1).NumGrhs
    lstSelGrhs.AddItem StreamData(List2.ListIndex + 1).grh_list(loopc)
Next loopc

DataChanged = DataTemp

indexs = frmMain.List2.ListIndex + 1

General_Particle_Create indexs, 50, 50

End Sub

Private Sub List2_KeyUp(KeyCode As Integer, Shift As Integer)
Dim loopc As Long
Dim DataTemp As Boolean
DataTemp = DataChanged

'Set the values
txtPCount.Text = StreamData(List2.ListIndex + 1).NumOfParticles
txtX1.Text = StreamData(List2.ListIndex + 1).x1
txtY1.Text = StreamData(List2.ListIndex + 1).y1
txtX2.Text = StreamData(List2.ListIndex + 1).x2
txtY2.Text = StreamData(List2.ListIndex + 1).y2
txtAngle.Text = StreamData(List2.ListIndex + 1).angle
vecx1.Text = StreamData(List2.ListIndex + 1).vecx1
vecx2.Text = StreamData(List2.ListIndex + 1).vecx2
vecy1.Text = StreamData(List2.ListIndex + 1).vecy1
vecy2.Text = StreamData(List2.ListIndex + 1).vecy2
life1.Text = StreamData(List2.ListIndex + 1).life1
life2.Text = StreamData(List2.ListIndex + 1).life2
fric.Text = StreamData(List2.ListIndex + 1).friction
chkSpin.value = StreamData(List2.ListIndex + 1).spin
spin_speedL.Text = StreamData(List2.ListIndex + 1).spin_speedL
spin_speedH.Text = StreamData(List2.ListIndex + 1).spin_speedH
txtGravStrength.Text = StreamData(List2.ListIndex + 1).grav_strength
txtBounceStrength.Text = StreamData(List2.ListIndex + 1).bounce_strength

chkAlphaBlend.value = StreamData(List2.ListIndex + 1).AlphaBlend
chkGravity.value = StreamData(List2.ListIndex + 1).gravity

chkXMove.value = StreamData(List2.ListIndex + 1).XMove
chkYMove.value = StreamData(List2.ListIndex + 1).YMove
move_x1.Text = StreamData(List2.ListIndex + 1).move_x1
move_x2.Text = StreamData(List2.ListIndex + 1).move_x2
move_y1.Text = StreamData(List2.ListIndex + 1).move_y1
move_y2.Text = StreamData(List2.ListIndex + 1).move_y2
txRad.Text = StreamData(List2.ListIndex + 1).Radio

lstSelGrhs.Clear

For loopc = 1 To StreamData(List2.ListIndex + 1).NumGrhs
    lstSelGrhs.AddItem StreamData(List2.ListIndex + 1).grh_list(loopc)
Next loopc

End Sub

Private Sub lstGrhs_Click()
invpic.Cls
engine.GrhRenderToHdc lstGrhs.List(lstGrhs.ListIndex), invpic.hdc, 2, 2, True
End Sub

Private Sub mnuabrir_Click()
Call cmdOpenStreamFile_Click
End Sub

Private Sub cmdOpenStreamFile_Click()
Dim sFile As String

With ComDlg
    .Filter = "*.ini (Stream Data Files)|*.ini"
    .ShowOpen
    sFile = .FileName
End With

LoadStreamFile sFile
CurStreamFile = sFile

End Sub

Private Sub mnuguardar_Click()
Call Command4_Click
End Sub

Private Sub mnunueva_Click()
Call Command5_Click
End Sub

Private Sub mnusalir_Click()
End
End Sub

Private Sub mnusobre_Click()
Form1.Show vbModal, Me
End Sub

Private Sub sobre_Click()
Form1.Show
End Sub

Private Sub TabStrip1_Click()
Select Case TabStrip1.SelectedItem.index
Case 1:
    frmSettings.Visible = True
    frameColorSettings.Visible = False
    Frame2.Visible = False
    Frame1.Visible = False
    frameSpinSettings.Visible = False
    frameMovement.Visible = False
    frameGravity.Visible = False
    frmfade.Visible = False
Case 2:
    frmSettings.Visible = False
    frameColorSettings.Visible = False
    Frame2.Visible = False
    Frame1.Visible = False
    frameSpinSettings.Visible = False
    frameMovement.Visible = False
    frameGravity.Visible = True
    frmfade.Visible = False
Case 3:
    frmSettings.Visible = False
    frameColorSettings.Visible = False
    Frame2.Visible = False
    Frame1.Visible = False
    frameSpinSettings.Visible = False
    frameMovement.Visible = True
    frameGravity.Visible = False
    frmfade.Visible = False
Case 4:
    frmSettings.Visible = False
    frameColorSettings.Visible = False
    Frame2.Visible = False
    Frame1.Visible = False
    frameSpinSettings.Visible = True
    frameMovement.Visible = False
    frameGravity.Visible = False
    frmfade.Visible = False
Case 5:
    frmSettings.Visible = False
    frameColorSettings.Visible = False
    Frame2.Visible = True
    Frame1.Visible = False
    frameSpinSettings.Visible = False
    frameMovement.Visible = False
    frameGravity.Visible = False
    frmfade.Visible = False
Case 6:
    frmSettings.Visible = False
    frameColorSettings.Visible = False
    Frame2.Visible = False
    Frame1.Visible = True
    frameSpinSettings.Visible = False
    frameMovement.Visible = False
    frameGravity.Visible = False
    frmfade.Visible = False
Case 7:
    frmSettings.Visible = False
    frameColorSettings.Visible = True
    Frame2.Visible = False
    Frame1.Visible = False
    frameSpinSettings.Visible = False
    frameMovement.Visible = False
    frameGravity.Visible = False
    frmfade.Visible = False
Case 8:
    frmSettings.Visible = False
    frameColorSettings.Visible = False
    Frame2.Visible = False
    Frame1.Visible = False
    frameSpinSettings.Visible = False
    frameMovement.Visible = False
    frameGravity.Visible = False
    frmfade.Visible = True
End Select
End Sub

Private Sub txRad_Change()
On Error Resume Next
StreamData(frmMain.List2.ListIndex + 1).Radio = Val(txRad.Text)
End Sub

Private Sub txtrx_Change()
On Error Resume Next
StreamData(frmMain.List2.ListIndex + 1).grh_resizex = txtrx.Text
End Sub

Private Sub txtry_Change()
On Error Resume Next
StreamData(frmMain.List2.ListIndex + 1).grh_resizey = txtry.Text
End Sub

Private Sub vecx1_GotFocus()

vecx1.SelStart = 0
vecx1.SelLength = Len(vecx1.Text)

End Sub

Private Sub vecx1_Change()
On Error Resume Next
DataChanged = True

StreamData(frmMain.List2.ListIndex + 1).vecx1 = vecx1.Text
End Sub

Private Sub vecx2_GotFocus()

vecx2.SelStart = 0
vecx2.SelLength = Len(vecx2.Text)

End Sub

Private Sub vecx2_Change()
On Error Resume Next
DataChanged = True

StreamData(frmMain.List2.ListIndex + 1).vecx2 = vecx2.Text
End Sub

Private Sub vecy1_GotFocus()

vecy1.SelStart = 0
vecy1.SelLength = Len(vecy1.Text)

End Sub

Private Sub vecy1_Change()
On Error Resume Next
DataChanged = True

StreamData(frmMain.List2.ListIndex + 1).vecy1 = vecy1.Text
End Sub

Private Sub vecy2_GotFocus()

vecy2.SelStart = 0
vecy2.SelLength = Len(vecy2.Text)

End Sub

Private Sub vecy2_Change()
On Error Resume Next
DataChanged = True

StreamData(frmMain.List2.ListIndex + 1).vecy2 = vecy2.Text
End Sub

Private Sub life1_GotFocus()

life1.SelStart = 0
life1.SelLength = Len(life1.Text)

End Sub

Private Sub life1_Change()
On Error Resume Next
DataChanged = True

StreamData(frmMain.List2.ListIndex + 1).life1 = life1.Text
End Sub

Private Sub life2_GotFocus()

life2.SelStart = 0
life2.SelLength = Len(life2.Text)

End Sub

Private Sub life2_Change()
On Error Resume Next
DataChanged = True

StreamData(frmMain.List2.ListIndex + 1).life2 = life2.Text
End Sub

Private Sub fric_GotFocus()

fric.SelStart = 0
fric.SelLength = Len(fric.Text)

End Sub

Private Sub fric_Change()
On Error Resume Next
DataChanged = True

StreamData(frmMain.List2.ListIndex + 1).friction = fric.Text
End Sub

Private Sub spin_speedL_GotFocus()

spin_speedL.SelStart = 0
spin_speedL.SelLength = Len(spin_speedH.Text)

End Sub

Private Sub spin_speedL_Change()
On Error Resume Next
DataChanged = True

StreamData(frmMain.List2.ListIndex + 1).spin_speedL = spin_speedL.Text
End Sub

Private Sub spin_speedH_GotFocus()

spin_speedH.SelStart = 0
spin_speedH.SelLength = Len(spin_speedH.Text)

End Sub

Private Sub spin_speedH_Change()
On Error Resume Next
DataChanged = True

StreamData(frmMain.List2.ListIndex + 1).spin_speedH = spin_speedH.Text
End Sub

Private Sub txtPCount_GotFocus()

txtPCount.SelStart = 0
txtPCount.SelLength = Len(txtPCount.Text)

End Sub

Private Sub txtPCount_Change()
On Error Resume Next
DataChanged = True

StreamData(frmMain.List2.ListIndex + 1).NumOfParticles = txtPCount.Text
End Sub

Private Sub txtX1_Change()
On Error Resume Next
DataChanged = True

StreamData(frmMain.List2.ListIndex + 1).x1 = txtX1.Text
End Sub

Private Sub txtX1_GotFocus()

txtX1.SelStart = 0
txtX1.SelLength = Len(txtX1.Text)

End Sub

Private Sub txtY1_Change()
On Error Resume Next
DataChanged = True

StreamData(frmMain.List2.ListIndex + 1).y1 = txtY1.Text
End Sub

Private Sub txtY1_GotFocus()

txtY1.SelStart = 0
txtY1.SelLength = Len(txtY1.Text)

End Sub

Private Sub txtX2_Change()
On Error Resume Next
DataChanged = True

StreamData(frmMain.List2.ListIndex + 1).x2 = txtX2.Text
End Sub

Private Sub txtX2_GotFocus()

txtX2.SelStart = 0
txtX2.SelLength = Len(txtX2.Text)

End Sub

Private Sub txtY2_Change()
On Error Resume Next
DataChanged = True

StreamData(frmMain.List2.ListIndex + 1).y2 = txtY2.Text
End Sub

Private Sub txtY2_GotFocus()

txtY2.SelStart = 0
txtY2.SelLength = Len(txtY2.Text)

End Sub

Private Sub txtAngle_Change()
On Error Resume Next
DataChanged = True

StreamData(frmMain.List2.ListIndex + 1).angle = txtAngle.Text
End Sub

Private Sub txtAngle_GotFocus()

txtAngle.SelStart = 0
txtAngle.SelLength = Len(txtAngle.Text)

End Sub

Private Sub txtGravStrength_Change()
On Error Resume Next
DataChanged = True

StreamData(frmMain.List2.ListIndex + 1).grav_strength = txtGravStrength.Text
End Sub

Private Sub txtGravStrength_GotFocus()

txtGravStrength.SelStart = 0
txtGravStrength.SelLength = Len(txtGravStrength.Text)

End Sub

Private Sub txtBounceStrength_Change()
On Error Resume Next
DataChanged = True

StreamData(frmMain.List2.ListIndex + 1).bounce_strength = txtBounceStrength.Text
End Sub

Private Sub txtBounceStrength_GotFocus()

txtBounceStrength.SelStart = 0
txtBounceStrength.SelLength = Len(txtBounceStrength.Text)

End Sub

Private Sub move_x1_Change()
On Error Resume Next
DataChanged = True

StreamData(frmMain.List2.ListIndex + 1).move_x1 = move_x1.Text
End Sub

Private Sub move_x1_GotFocus()

move_x1.SelStart = 0
move_x1.SelLength = Len(move_x1.Text)

End Sub

Private Sub move_x2_Change()
On Error Resume Next
DataChanged = True

StreamData(frmMain.List2.ListIndex + 1).move_x2 = move_x2.Text
End Sub

Private Sub move_x2_GotFocus()

move_x2.SelStart = 0
move_x2.SelLength = Len(move_x2.Text)

End Sub

Private Sub move_y1_Change()
On Error Resume Next
DataChanged = True

StreamData(frmMain.List2.ListIndex + 1).move_y1 = move_y1.Text
End Sub

Private Sub move_y1_GotFocus()

move_y1.SelStart = 0
move_y1.SelLength = Len(move_y1.Text)

End Sub

Private Sub move_y2_Change()
On Error Resume Next
DataChanged = True

StreamData(frmMain.List2.ListIndex + 1).move_y2 = move_y2.Text
End Sub

Private Sub move_y2_GotFocus()

move_y2.SelStart = 0
move_y2.SelLength = Len(move_y2.Text)

End Sub


Private Sub chkAlphaBlend_Click()

DataChanged = True

StreamData(frmMain.List2.ListIndex + 1).AlphaBlend = chkAlphaBlend.value
End Sub

Private Sub chkGravity_Click()

DataChanged = True

StreamData(frmMain.List2.ListIndex + 1).gravity = chkGravity.value

If chkGravity.value = vbChecked Then
    txtGravStrength.Enabled = True
    txtBounceStrength.Enabled = True
Else
    txtGravStrength.Enabled = False
    txtBounceStrength.Enabled = False
End If

End Sub

Private Sub chkXMove_Click()

DataChanged = True

StreamData(frmMain.List2.ListIndex + 1).XMove = chkXMove.value

If chkXMove.value = vbChecked Then
    move_x1.Enabled = True
    move_x2.Enabled = True
Else
    move_x1.Enabled = False
    move_x2.Enabled = False
End If

End Sub

Private Sub chkYMove_Click()

DataChanged = True

StreamData(frmMain.List2.ListIndex + 1).YMove = chkYMove.value

If chkYMove.value = vbChecked Then
    move_y1.Enabled = True
    move_y2.Enabled = True
Else
    move_y1.Enabled = False
    move_y2.Enabled = False
End If

End Sub

Private Sub BScroll_Change()
On Error Resume Next
DataChanged = True


StreamData(frmMain.List2.ListIndex + 1).colortint(lstColorSets.ListIndex).B = BScroll.value
txtB.Text = BScroll.value

picColor.BackColor = RGB(txtB.Text, txtG.Text, txtR.Text)

End Sub

Private Sub chkNeverDies_Click()

DataChanged = True

If chkNeverDies.value = vbChecked Then
    life.Enabled = False
    StreamData(frmMain.List2.ListIndex + 1).life_counter = -1
Else
    life.Enabled = True
    StreamData(frmMain.List2.ListIndex + 1).life_counter = life.Text
End If
End Sub

Private Sub chkSpin_Click()

DataChanged = True

StreamData(frmMain.List2.ListIndex + 1).spin = chkSpin.value

If chkSpin.value = vbChecked Then
    spin_speedL.Enabled = True
    spin_speedH.Enabled = True
Else
    spin_speedL.Enabled = False
    spin_speedH.Enabled = False
End If

End Sub



Private Sub GScroll_Change()
On Error Resume Next
DataChanged = True


StreamData(frmMain.List2.ListIndex + 1).colortint(lstColorSets.ListIndex).g = GScroll.value
txtG.Text = GScroll.value

picColor.BackColor = RGB(txtB.Text, txtG.Text, txtR.Text)

End Sub

Private Sub life_Change()
On Error Resume Next
DataChanged = True

StreamData(frmMain.List2.ListIndex + 1).life_counter = life.Text
End Sub

Private Sub life_GotFocus()

life.SelStart = 0
life.SelLength = Len(life.Text)

End Sub

Private Sub lstColorSets_Click()

Dim DataTemp As Boolean
DataTemp = DataChanged

RScroll.value = StreamData(frmMain.List2.ListIndex + 1).colortint(lstColorSets.ListIndex).r
GScroll.value = StreamData(frmMain.List2.ListIndex + 1).colortint(lstColorSets.ListIndex).g
BScroll.value = StreamData(frmMain.List2.ListIndex + 1).colortint(lstColorSets.ListIndex).B

DataChanged = DataTemp

End Sub

Private Sub RScroll_Change()
On Error Resume Next
DataChanged = True


StreamData(frmMain.List2.ListIndex + 1).colortint(lstColorSets.ListIndex).r = RScroll.value
txtR.Text = RScroll.value

picColor.BackColor = RGB(txtB.Text, txtG.Text, txtR.Text)

End Sub

Private Sub speed_Change()
On Error Resume Next
DataChanged = True

'Arrange decimal separator
Dim temp As String
temp = General_Field_Read(1, speed.Text, 44)
If Not temp = "" Then
    speed.Text = temp & "." & Right(speed.Text, Len(speed.Text) - Len(temp) - 1)
    speed.SelStart = Len(speed.Text)
    speed.SelLength = 0
End If
StreamData(frmMain.List2.ListIndex + 1).speed = Val(speed.Text)
End Sub

Private Sub speed_GotFocus()

speed.SelStart = 0
speed.SelLength = Len(speed.Text)

End Sub

Private Sub lstSelGrhs_DblClick()

Call cmdDelete_Click

End Sub
Private Sub cmdDelete_Click()
Dim loopc As Long

If lstSelGrhs.ListIndex >= 0 Then lstSelGrhs.RemoveItem lstSelGrhs.ListIndex

StreamData(List2.ListIndex + 1).NumGrhs = lstSelGrhs.ListCount

If StreamData(List2.ListIndex + 1).NumGrhs = 0 Then
    Erase StreamData(List2.ListIndex + 1).grh_list
Else
    ReDim StreamData(List2.ListIndex + 1).grh_list(1 To lstSelGrhs.ListCount)
End If

For loopc = 1 To StreamData(List2.ListIndex + 1).NumGrhs
    StreamData(List2.ListIndex + 1).grh_list(loopc) = lstSelGrhs.List(loopc - 1)
Next loopc

End Sub
Private Sub lstSelGrhs_Click()
Dim GrhInfo

Dim filepath As String
Dim src_x As Long
Dim src_y As Long
Dim src_width As Long
Dim src_height As Long
Dim framecount As Long


If framecount <= 0 Then Exit Sub

invpic.Cls
engine.GrhRenderToHdc lstSelGrhs.List(lstSelGrhs.ListIndex), invpic.hdc, 2, 2, True

End Sub



Private Sub lstGrhs_DblClick()

Call cmdAdd_Click

End Sub
Private Sub cmdAdd_Click()

Dim loopc As Long

If lstGrhs.ListIndex >= 0 Then lstSelGrhs.AddItem lstGrhs.List(lstGrhs.ListIndex)

StreamData(List2.ListIndex + 1).NumGrhs = lstSelGrhs.ListCount

ReDim StreamData(List2.ListIndex + 1).grh_list(1 To lstSelGrhs.ListCount)

For loopc = 1 To StreamData(List2.ListIndex + 1).NumGrhs
    StreamData(List2.ListIndex + 1).grh_list(loopc) = lstSelGrhs.List(loopc - 1)
Next loopc

End Sub
Private Sub chkresize_Click()
If chkresize.value = vbChecked Then
    StreamData(frmMain.List2.ListIndex + 1).grh_resize = True
Else
   StreamData(frmMain.List2.ListIndex + 1).grh_resize = False
End If
End Sub

Private Sub LoadStreamFile(StreamFile As String)
On Error Resume Next
    Dim loopc As Long
    
    '****************************
    'load stream types
    '****************************
    TotalStreams = Val(General_Var_Get(StreamFile, "INIT", "Total"))
    
    'resize StreamData array
    ReDim StreamData(1 To TotalStreams) As Stream
    
    'clear combo box
    List2.Clear
    
    Dim i As Long
    Dim GrhListing As String
    'fill StreamData array with info from Particles.ini
    For loopc = 1 To TotalStreams
        StreamData(loopc).Name = General_Var_Get(StreamFile, Val(loopc), "Name")
        StreamData(loopc).NumOfParticles = General_Var_Get(StreamFile, Val(loopc), "NumOfParticles")
        StreamData(loopc).x1 = General_Var_Get(StreamFile, Val(loopc), "X1")
        StreamData(loopc).y1 = General_Var_Get(StreamFile, Val(loopc), "Y1")
        StreamData(loopc).x2 = General_Var_Get(StreamFile, Val(loopc), "X2")
        StreamData(loopc).y2 = General_Var_Get(StreamFile, Val(loopc), "Y2")
        StreamData(loopc).angle = General_Var_Get(StreamFile, Val(loopc), "Angle")
        StreamData(loopc).vecx1 = General_Var_Get(StreamFile, Val(loopc), "VecX1")
        StreamData(loopc).vecx2 = General_Var_Get(StreamFile, Val(loopc), "VecX2")
        StreamData(loopc).vecy1 = General_Var_Get(StreamFile, Val(loopc), "VecY1")
        StreamData(loopc).vecy2 = General_Var_Get(StreamFile, Val(loopc), "VecY2")
        StreamData(loopc).life1 = General_Var_Get(StreamFile, Val(loopc), "Life1")
        StreamData(loopc).life2 = General_Var_Get(StreamFile, Val(loopc), "Life2")
        StreamData(loopc).friction = General_Var_Get(StreamFile, Val(loopc), "Friction")
        StreamData(loopc).spin = General_Var_Get(StreamFile, Val(loopc), "Spin")
        StreamData(loopc).spin_speedL = General_Var_Get(StreamFile, Val(loopc), "Spin_SpeedL")
        StreamData(loopc).spin_speedH = General_Var_Get(StreamFile, Val(loopc), "Spin_SpeedH")
        StreamData(loopc).AlphaBlend = General_Var_Get(StreamFile, Val(loopc), "AlphaBlend")
        StreamData(loopc).gravity = General_Var_Get(StreamFile, Val(loopc), "Gravity")
        StreamData(loopc).grav_strength = General_Var_Get(StreamFile, Val(loopc), "Grav_Strength")
        StreamData(loopc).bounce_strength = General_Var_Get(StreamFile, Val(loopc), "Bounce_Strength")
        StreamData(loopc).XMove = General_Var_Get(StreamFile, Val(loopc), "XMove")
        StreamData(loopc).YMove = General_Var_Get(StreamFile, Val(loopc), "YMove")
        StreamData(loopc).move_x1 = General_Var_Get(StreamFile, Val(loopc), "move_x1")
        StreamData(loopc).move_x2 = General_Var_Get(StreamFile, Val(loopc), "move_x2")
        StreamData(loopc).move_y1 = General_Var_Get(StreamFile, Val(loopc), "move_y1")
        StreamData(loopc).move_y2 = General_Var_Get(StreamFile, Val(loopc), "move_y2")
        StreamData(loopc).Radio = Val(General_Var_Get(StreamFile, Val(loopc), "Radio"))
        StreamData(loopc).life_counter = General_Var_Get(StreamFile, Val(loopc), "life_counter")
        StreamData(loopc).speed = Val(General_Var_Get(StreamFile, Val(loopc), "Speed"))
        StreamData(loopc).grh_resize = Val(General_Var_Get(StreamFile, Val(loopc), "resize"))
        StreamData(loopc).grh_resizex = Val(General_Var_Get(StreamFile, Val(loopc), "rx"))
        StreamData(loopc).grh_resizey = Val(General_Var_Get(StreamFile, Val(loopc), "ry"))
        StreamData(loopc).NumGrhs = General_Var_Get(StreamFile, Val(loopc), "NumGrhs")
        
        ReDim StreamData(loopc).grh_list(1 To StreamData(loopc).NumGrhs)
        GrhListing = General_Var_Get(StreamFile, Val(loopc), "Grh_List")
        
        For i = 1 To StreamData(loopc).NumGrhs
            StreamData(loopc).grh_list(i) = General_Field_Read(Str(i), GrhListing, 44)
        Next i
        
        Dim TempSet As String
        Dim ColorSet As Long
        
        For ColorSet = 1 To 4
            TempSet = General_Var_Get(StreamFile, Val(loopc), "ColorSet" & ColorSet)
            StreamData(loopc).colortint(ColorSet - 1).r = General_Field_Read(1, TempSet, 44)
            StreamData(loopc).colortint(ColorSet - 1).g = General_Field_Read(2, TempSet, 44)
            StreamData(loopc).colortint(ColorSet - 1).B = General_Field_Read(3, TempSet, 44)
        Next ColorSet
        
        'fill stream type combo box
        List2.AddItem loopc & " - " & StreamData(loopc).Name
    Next loopc
    
    'set list box index to 1st item
    List2.ListIndex = 0

End Sub
