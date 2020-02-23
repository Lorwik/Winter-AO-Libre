VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Acerca de WoldEditor"
   ClientHeight    =   4350
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   4365
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   Picture         =   "frmAbout.frx":628A
   ScaleHeight     =   3002.447
   ScaleMode       =   0  'User
   ScaleWidth      =   4098.96
   StartUpPosition =   2  'CenterScreen
   Begin WorldEditor.lvButtons_H cmdOK 
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   3840
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "&Aceptar"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Modificado, mejorado y adaptado a DX8 por: Lorwik www.rincondelAO.com.ar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   735
      Left            =   360
      TabIndex        =   8
      Top             =   3000
      Width           =   3735
   End
   Begin VB.Label lblCred 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Mejoras por:                       About, Loopzer, Salvito"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   480
      Index           =   1
      Left            =   720
      TabIndex        =   7
      Top             =   1320
      Width           =   3075
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "WorldEditor"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   1200
      TabIndex        =   6
      Top             =   0
      Width           =   2970
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Versión"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   270
      Left            =   1560
      TabIndex        =   5
      Top             =   480
      Width           =   810
   End
   Begin VB.Label lblCred 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Basado en códigos de BaronSoft, Dunga, Maraxus y Morgolock"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   450
      Index           =   0
      Left            =   960
      TabIndex        =   3
      Top             =   1800
      Width           =   2565
   End
   Begin VB.Label lblCred 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "LaVolpe Button (c) by LaVolpe"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   285
      Index           =   4
      Left            =   120
      TabIndex        =   2
      Top             =   2880
      Width           =   4125
   End
   Begin VB.Label lblCred 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Agradecimientos especiales: Dunga, Manikke, Maraxus, Kiko, Koke y todos ;)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   645
      Index           =   3
      Left            =   960
      TabIndex        =   1
      Top             =   2280
      Width           =   2565
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   112.686
      X2              =   3944.016
      Y1              =   2567.61
      Y2              =   2567.61
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   112.686
      X2              =   3944.016
      Y1              =   2567.61
      Y2              =   2567.61
   End
   Begin VB.Label lblCred 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Programado por ^[GS]^"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   240
      Index           =   2
      Left            =   720
      TabIndex        =   0
      Top             =   1080
      Width           =   3075
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFF00&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      DrawMode        =   9  'Not Mask Pen
      FillColor       =   &H00C0C0FF&
      FillStyle       =   0  'Solid
      Height          =   2655
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   1080
      Width           =   4095
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'**************************************************************
Option Explicit

Private Sub cmdOK_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
  Unload Me
End Sub

Private Sub Form_Load()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    Me.Caption = "Acerca de " & App.Title
    lblVersion.Caption = "Versión " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.Title
End Sub

