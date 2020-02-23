VERSION 5.00
Begin VB.Form frmCargando 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Argentum"
   ClientHeight    =   2265
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6000
   Icon            =   "frmCargando.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2265
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   -840
      Picture         =   "frmCargando.frx":628A
      ScaleHeight     =   465
      ScaleWidth      =   12000
      TabIndex        =   2
      Top             =   1800
      Width           =   12000
      Begin VB.Image P2 
         Height          =   480
         Left            =   1560
         Picture         =   "frmCargando.frx":BF09
         ToolTipText     =   "Cuerpos"
         Top             =   0
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image P4 
         Height          =   480
         Left            =   3720
         Picture         =   "frmCargando.frx":C3C9
         ToolTipText     =   "NPC's"
         Top             =   0
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image P3 
         Height          =   480
         Left            =   2640
         Picture         =   "frmCargando.frx":D00B
         ToolTipText     =   "Cabezas"
         Top             =   0
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image P1 
         Height          =   480
         Left            =   120
         Picture         =   "frmCargando.frx":D84F
         ToolTipText     =   "Base de Datos"
         Top             =   0
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image P5 
         Height          =   480
         Left            =   4800
         Picture         =   "frmCargando.frx":E093
         ToolTipText     =   "Objetos"
         Top             =   0
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label L 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BdD"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   960
         TabIndex        =   8
         Top             =   120
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label L 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Body"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   2040
         TabIndex        =   7
         Top             =   120
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label L 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Head"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   2
         Left            =   3120
         TabIndex        =   6
         Top             =   120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label L 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NPC's"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   3
         Left            =   4200
         TabIndex        =   5
         Top             =   120
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Label L 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OBJ's"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   4
         Left            =   5280
         TabIndex        =   4
         Top             =   120
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.Label L 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Trig."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   5
         Left            =   6360
         TabIndex        =   3
         Top             =   120
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Image P6 
         Height          =   480
         Left            =   5880
         Picture         =   "frmCargando.frx":E8D7
         ToolTipText     =   "Función de Trigger"
         Top             =   0
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   1800
      Left            =   0
      ScaleHeight     =   1800
      ScaleWidth      =   6000
      TabIndex        =   0
      Top             =   0
      Width           =   6000
      Begin VB.Label X 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "..."
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
         Height          =   255
         Left            =   1080
         TabIndex        =   9
         Top             =   1560
         Width           =   4095
      End
      Begin VB.Label verX 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "v?.?.?"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   255
         TabIndex        =   1
         Top             =   0
         Width           =   555
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFFFFF&
         BorderStyle     =   0  'Transparent
         DrawMode        =   3  'Not Merge Pen
         FillColor       =   &H00FF80FF&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   -120
         Shape           =   4  'Rounded Rectangle
         Top             =   -120
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmCargando"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

