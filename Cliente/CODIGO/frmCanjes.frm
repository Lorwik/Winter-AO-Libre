VERSION 5.00
Begin VB.Form frmcanjes 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Canjes de Puntos"
   ClientHeight    =   8580
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11400
   LinkTopic       =   "Form1"
   ScaleHeight     =   8580
   ScaleWidth      =   11400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame7 
      BackColor       =   &H00000000&
      Caption         =   "Monturas"
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
      Height          =   2775
      Left            =   120
      TabIndex        =   24
      Top             =   5760
      Width           =   9015
      Begin VB.PictureBox Picture14 
         Height          =   495
         Left            =   6240
         Picture         =   "frmcanjes.frx":0000
         ScaleHeight     =   400
         ScaleMode       =   0  'User
         ScaleWidth      =   435
         TabIndex        =   35
         Top             =   1680
         Width           =   495
      End
      Begin VB.PictureBox Picture13 
         Height          =   495
         Left            =   3120
         Picture         =   "frmcanjes.frx":0844
         ScaleHeight     =   400
         ScaleMode       =   0  'User
         ScaleWidth      =   435
         TabIndex        =   32
         Top             =   1680
         Width           =   495
      End
      Begin VB.PictureBox Picture12 
         Height          =   495
         Left            =   6240
         Picture         =   "frmcanjes.frx":1087
         ScaleHeight     =   400
         ScaleMode       =   0  'User
         ScaleWidth      =   435
         TabIndex        =   31
         Top             =   240
         Width           =   495
      End
      Begin VB.PictureBox Picture11 
         Height          =   495
         Left            =   120
         Picture         =   "frmcanjes.frx":1CCB
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   29
         Top             =   1680
         Width           =   495
      End
      Begin VB.PictureBox Picture10 
         Height          =   495
         Left            =   3120
         Picture         =   "frmcanjes.frx":2D0F
         ScaleHeight     =   400
         ScaleMode       =   0  'User
         ScaleWidth      =   435
         TabIndex        =   27
         Top             =   360
         Width           =   495
      End
      Begin VB.PictureBox Picture9 
         Height          =   495
         Left            =   120
         Picture         =   "frmcanjes.frx":3951
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   25
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Montura de Tigre Amarillo [Sagrado]: MaxDef=25/MinDef=25 Equitacion: 100              40 Puntos"
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
         Height          =   1095
         Left            =   6840
         TabIndex        =   36
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00FFFFFF&
         X1              =   6120
         X2              =   6120
         Y1              =   120
         Y2              =   2760
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Montura de Dragon Rojo [Sagrado]: MaxDef=30/MinDef=30 Equitacion: 100              45 Puntos"
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
         Height          =   1095
         Left            =   3720
         TabIndex        =   34
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Montura de Dragon Amarillo [Sagrado]: MaxDef=30/MinDef=30 Equitacion: 100              45 Puntos"
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
         Height          =   1095
         Left            =   6840
         TabIndex        =   33
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label13 
         BackColor       =   &H00000000&
         Caption         =   "Montura de Buey [Sagrado]: MaxDef=25/MinDef=25 Equitacion:100          40 Puntos"
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
         Height          =   975
         Left            =   720
         TabIndex        =   30
         Top             =   1680
         Width           =   2055
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   9120
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Montura de Dragon Negro [Sagrado]: MaxDef=35/MinDef=35 Equitacion: 100              50 Puntos"
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
         Height          =   1095
         Left            =   3720
         TabIndex        =   28
         Top             =   240
         Width           =   2175
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00FFFFFF&
         X1              =   3000
         X2              =   3000
         Y1              =   120
         Y2              =   2760
      End
      Begin VB.Label Label10 
         BackColor       =   &H00000000&
         Caption         =   "Montura de Preclitus [Sagrado]: MaxDef=25/MinDef=25 Equitacion: 100         40 Puntos"
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
         Height          =   975
         Left            =   720
         TabIndex        =   26
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00000000&
      Caption         =   "Armaduras y Tunicas"
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
      Height          =   2775
      Left            =   3360
      TabIndex        =   21
      Top             =   1560
      Width           =   5775
      Begin VB.PictureBox Picture18 
         Height          =   495
         Left            =   2880
         Picture         =   "frmcanjes.frx":4593
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   45
         Top             =   1560
         Width           =   495
      End
      Begin VB.PictureBox Picture16 
         Height          =   495
         Left            =   120
         Picture         =   "frmcanjes.frx":51D5
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   39
         Top             =   1560
         Width           =   495
      End
      Begin VB.PictureBox Picture15 
         Height          =   495
         Left            =   2880
         Picture         =   "frmcanjes.frx":5CF7
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   37
         Top             =   240
         Width           =   495
      End
      Begin VB.PictureBox Picture8 
         Height          =   495
         Left            =   120
         Picture         =   "frmcanjes.frx":6939
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   22
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label22 
         BackColor       =   &H00000000&
         Caption         =   "Tunica LwK [Sagrado]: MaxDef=46/MinDef=40 55 Puntos"
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
         Height          =   975
         Left            =   3600
         TabIndex        =   46
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Label Label18 
         BackColor       =   &H00000000&
         Caption         =   "Armadura de Altair (Bajos) [Sagrado]: MaxDef=70/MinDef=65 55 Puntos"
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
         Height          =   1095
         Left            =   720
         TabIndex        =   40
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   5760
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Label Label17 
         BackColor       =   &H00000000&
         Caption         =   "Tunica Infernal (Bajos) [Sagrado]: MaxDef=46/MinDef=40 55 Puntos"
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
         Height          =   975
         Left            =   3480
         TabIndex        =   38
         Top             =   240
         Width           =   2055
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00FFFFFF&
         X1              =   2760
         X2              =   2760
         Y1              =   120
         Y2              =   2760
      End
      Begin VB.Label Label9 
         BackColor       =   &H00000000&
         Caption         =   "Armadura del Logouth [Sagrado]: MaxDef=70/MinDef=65 55 Puntos"
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
         Height          =   975
         Left            =   720
         TabIndex        =   23
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00000000&
      Caption         =   "Escudos"
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
      Height          =   2655
      Left            =   120
      TabIndex        =   15
      Top             =   3000
      Width           =   3255
      Begin VB.PictureBox Picture6 
         Height          =   495
         Left            =   120
         Picture         =   "frmcanjes.frx":6DAF
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   19
         Top             =   1560
         Width           =   495
      End
      Begin VB.PictureBox Picture5 
         Height          =   495
         Left            =   120
         Picture         =   "frmcanjes.frx":7193
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   17
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label7 
         BackColor       =   &H00000000&
         Caption         =   "Escudo Desintegrador[Sagrado]: MaxDef=30/MinDef=30 35 Puntos"
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
         Height          =   855
         Left            =   840
         TabIndex        =   20
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   3960
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Label Label6 
         BackColor       =   &H00000000&
         Caption         =   "Escudo de Torre + 1: MaxDef=24/MinDef=24 5 Puntos"
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
         Height          =   975
         Left            =   720
         TabIndex        =   18
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00000000&
      Caption         =   "Amuletos Magicos"
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
      Height          =   1335
      Left            =   120
      TabIndex        =   12
      Top             =   1560
      Width           =   3255
      Begin VB.PictureBox Picture1 
         Height          =   495
         Left            =   120
         Picture         =   "frmcanjes.frx":79D5
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   13
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Pendiente del Sacrificio: Con este Pendiente al morir solo perderas el Pendiente.                     25 Puntos"
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
         Height          =   975
         Left            =   720
         TabIndex        =   14
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      Caption         =   "Cascos y Gorros"
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
      Height          =   1335
      Left            =   3360
      TabIndex        =   7
      Top             =   4320
      Width           =   5775
      Begin VB.PictureBox Picture3 
         Height          =   495
         Left            =   3000
         Picture         =   "frmcanjes.frx":80F1
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   10
         Top             =   240
         Width           =   495
      End
      Begin VB.PictureBox Picture4 
         Height          =   495
         Left            =   120
         Picture         =   "frmcanjes.frx":8D33
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   8
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label4 
         BackColor       =   &H00000000&
         Caption         =   "Gorro de Defensa Magica (+20)                     MaxDef=25/MinDef=20 15 Puntos"
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
         Height          =   975
         Left            =   3600
         TabIndex        =   11
         Top             =   240
         Width           =   2055
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   2880
         X2              =   2880
         Y1              =   120
         Y2              =   1320
      End
      Begin VB.Label Label5 
         BackColor       =   &H00000000&
         Caption         =   "Casco Bikingo: MaxDef=50/MinDef=45 10 Puntos"
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
         Height          =   735
         Left            =   720
         TabIndex        =   9
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Armas"
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
      Height          =   1455
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   9015
      Begin VB.PictureBox Picture17 
         Height          =   495
         Left            =   5880
         Picture         =   "frmcanjes.frx":9975
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   43
         Top             =   360
         Width           =   495
      End
      Begin VB.PictureBox Picture7 
         Height          =   495
         Left            =   3000
         Picture         =   "frmcanjes.frx":A5B7
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   5
         Top             =   360
         Width           =   495
      End
      Begin VB.PictureBox Picture2 
         Height          =   495
         Left            =   120
         Picture         =   "frmcanjes.frx":ADF9
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   3
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label21 
         BackColor       =   &H00000000&
         Caption         =   "Arco Argentum [Sagrado]: MinHit=14/MaxHit=22 50 Puntos"
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
         Height          =   855
         Left            =   6480
         TabIndex        =   44
         Top             =   240
         Width           =   1935
      End
      Begin VB.Line Line12 
         BorderColor     =   &H00FFFFFF&
         X1              =   5760
         X2              =   5760
         Y1              =   120
         Y2              =   1440
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   2760
         X2              =   2760
         Y1              =   120
         Y2              =   1440
      End
      Begin VB.Label Label8 
         BackColor       =   &H00000000&
         Caption         =   "Vara Infernal [Sagrado]: MinHit=5/MaxHit=15 40 Puntos"
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
         Height          =   855
         Left            =   3600
         TabIndex        =   6
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Caption         =   "Espada Argentum [Sagrado]: MinHit=25/MaxHit=29                                   50 Puntos"
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
         Height          =   975
         Left            =   720
         TabIndex        =   4
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Información"
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
      Height          =   8295
      Left            =   9240
      TabIndex        =   0
      Top             =   240
      Width           =   2055
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   $"frmcanjes.frx":B63B
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
         Height          =   2535
         Left            =   120
         TabIndex        =   42
         Top             =   5640
         Width           =   1815
      End
      Begin VB.Line Line11 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   2040
         Y1              =   5520
         Y2              =   5520
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   $"frmcanjes.frx":B734
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
         Height          =   2895
         Left            =   120
         TabIndex        =   41
         Top             =   2760
         Width           =   1815
      End
      Begin VB.Line Line10 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   2040
         Y1              =   2640
         Y2              =   2640
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   2040
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "Para ver los puntos disponibles escribe el comando /est"
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
         Height          =   615
         Left            =   120
         TabIndex        =   16
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Para canjear tus puntos debes de hacer click en el Item deseado, una vez hecho esta operacion no podras volver atras !! "
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
         Height          =   1575
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Image Image2 
      Height          =   225
      Left            =   10920
      MouseIcon       =   "frmcanjes.frx":B800
      MousePointer    =   99  'Custom
      Picture         =   "frmcanjes.frx":C4CA
      Top             =   0
      Width           =   420
   End
End
Attribute VB_Name = "frmcanjes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    HookSurfaceHwnd Me
End Sub

Private Sub Image1_Click()

End Sub

Private Sub Image2_Click()
frmcanjes.Visible = False
End Sub

Private Sub Picture1_Click()
Call SendData("KOTO1")
frmcanjes.Visible = False
End Sub

Private Sub Picture10_Click()
Call SendData("DONA2")
frmcanjes.Visible = False
End Sub

Private Sub Picture11_Click()
Call SendData("DONA3")
frmcanjes.Visible = False
End Sub

Private Sub Picture12_Click()
Call SendData("DONA5")
frmcanjes.Visible = False
End Sub

Private Sub Picture13_Click()
Call SendData("DONA4")
frmcanjes.Visible = False
End Sub

Private Sub Picture14_Click()
Call SendData("DONA6")
frmcanjes.Visible = False
End Sub

Private Sub Picture15_Click()
Call SendData("DONA7")
frmcanjes.Visible = False
End Sub

Private Sub Picture16_Click()
Call SendData("DONA8")
frmcanjes.Visible = False
End Sub

Private Sub Picture17_Click()
Call SendData("DONA9")
frmcanjes.Visible = False
End Sub

Private Sub Picture18_Click()
Call SendData("KWLF1")
frmcanjes.Visible = False
End Sub

Private Sub Picture2_Click()
Call SendData("KOTO2")
frmcanjes.Visible = False
End Sub

Private Sub Picture3_Click()
Call SendData("KOTO3")
frmcanjes.Visible = False
End Sub

Private Sub Picture4_Click()
Call SendData("KOTO4")
frmcanjes.Visible = False
End Sub

Private Sub Picture5_Click()
Call SendData("KOTO5")
frmcanjes.Visible = False
End Sub

Private Sub Picture6_Click()
Call SendData("DONA1")
frmcanjes.Visible = False
End Sub

Private Sub Picture7_Click()
Call SendData("KOTO7")
frmcanjes.Visible = False
End Sub

Private Sub Picture8_Click()
Call SendData("KOTO8")
frmcanjes.Visible = False
End Sub

Private Sub Picture9_Click()
Call SendData("KOTO9")
frmcanjes.Visible = False
End Sub
