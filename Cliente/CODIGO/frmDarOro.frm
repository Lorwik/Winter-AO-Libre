VERSION 5.00
Begin VB.Form frmDarOro 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Dar Oro"
   ClientHeight    =   2595
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   2820
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   2820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Nombre 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   0
      TabIndex        =   3
      Top             =   1320
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   255
      Left            =   1440
      TabIndex        =   2
      Top             =   2280
      Width           =   1335
   End
   Begin VB.TextBox Cantidad 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Text            =   "0"
      Top             =   1920
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Si el nombre es separado, no olvides sustituir los espacios en blanco por ""+"" (sin comillas)"
      Height          =   735
      Left            =   0
      TabIndex        =   7
      Top             =   480
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "¿A quien deseas entregar el oro?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   1080
      Width           =   3015
   End
   Begin VB.Label Label2 
      Caption         =   "Escriba la cantidad deseada"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   1680
      Width           =   2535
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Dar ORO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   2775
   End
End
Attribute VB_Name = "frmDarOro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Call SendData("/Daroro " & Nombre.Text & " " & Cantidad.Text)
Unload Me
End Sub
Private Sub Command2_Click()
Unload Me
End Sub
