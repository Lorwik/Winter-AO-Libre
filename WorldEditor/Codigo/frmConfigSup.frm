VERSION 5.00
Begin VB.Form frmConfigSup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuración Avanzada de Superficie "
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6315
   Icon            =   "frmConfigSup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   6315
   StartUpPosition =   2  'CenterScreen
   Begin WorldEditor.lvButtons_H cmdAceptar 
      Height          =   375
      Left            =   4200
      TabIndex        =   15
      Top             =   2280
      Width           =   1935
      _extentx        =   3413
      _extenty        =   661
      caption         =   "&Aceptar"
      capalign        =   2
      backstyle       =   2
      cgradient       =   0
      font            =   "frmConfigSup.frx":628A
      mode            =   0
      value           =   0
      cback           =   -2147483633
   End
   Begin VB.CommandButton cmdDM 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1605
      Width           =   240
   End
   Begin VB.CommandButton cmdDM 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   4005
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1605
      Width           =   240
   End
   Begin VB.CommandButton cmdDM 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   4005
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1200
      Width           =   240
   End
   Begin VB.CommandButton cmdDM 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1200
      Width           =   240
   End
   Begin VB.TextBox DMLargo 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   330
      Left            =   4245
      TabIndex        =   8
      Text            =   "0"
      Top             =   1560
      Width           =   1620
   End
   Begin VB.TextBox DMAncho 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   330
      Left            =   4245
      TabIndex        =   7
      Text            =   "0"
      Top             =   1155
      Width           =   1620
   End
   Begin VB.CheckBox DespMosaic 
      Appearance      =   0  'Flat
      Caption         =   "Desplazamiento de Mosaico"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   3240
      TabIndex        =   6
      Top             =   840
      Width           =   2880
   End
   Begin VB.TextBox mAncho 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   330
      Left            =   1080
      TabIndex        =   1
      Text            =   "1"
      Top             =   1200
      Width           =   1905
   End
   Begin VB.TextBox mLargo 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   330
      Left            =   1080
      TabIndex        =   0
      Text            =   "1"
      Top             =   1560
      Width           =   1905
   End
   Begin VB.CheckBox MOSAICO 
      Appearance      =   0  'Flat
      Caption         =   "Mosaico"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   165
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Largo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   3360
      TabIndex        =   14
      Top             =   1560
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Ancho"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   3360
      TabIndex        =   13
      Top             =   1200
      Width           =   525
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   255
      X2              =   6110
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ALERTA: Algunos superficies tienen un limite de mosaico, mirar la vista previa antes de colocar."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   6015
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Ancho"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   240
      TabIndex        =   4
      Top             =   1275
      Width           =   525
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Largo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   240
      TabIndex        =   3
      Top             =   1620
      Width           =   480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   240
      X2              =   6110
      Y1              =   2160
      Y2              =   2160
   End
End
Attribute VB_Name = "frmConfigSup"
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

Private Sub cmdDM_Click(index As Integer)
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************

On Error Resume Next
Select Case index
        Case 0
            DMAncho.Text = Str(Val(DMAncho.Text) + 1)
        Case 1
            DMAncho.Text = Str(Val(DMAncho.Text) - 1)
        Case 2
            DMLargo.Text = Str(Val(DMLargo.Text) - 1)
        Case 3
            DMLargo.Text = Str(Val(DMLargo.Text) + 1)
End Select
End Sub

Private Sub Form_Deactivate()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

Me.Hide
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

If UnloadMode <> 0 Then
    Cancel = True
    Me.Hide
End If
End Sub

Private Sub DespMosaic_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 26/05/06
'*************************************************
If LenB(DMAncho.Text) = 0 Then DMAncho.Text = "0"
If LenB(DMLargo.Text) = 0 Then DMLargo.Text = "0"
End Sub


Private Sub mAncho_KeyPress(KeyAscii As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
' Impedir que se ingrese un valor no numerico
If KeyAscii <> 8 And IsNumeric(Chr(KeyAscii)) = False Then KeyAscii = 0
End Sub

Private Sub mLargo_KeyPress(KeyAscii As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
' Impedir que se ingrese un valor no numerico
If KeyAscii <> 8 And IsNumeric(Chr(KeyAscii)) = False Then KeyAscii = 0
End Sub

Private Sub cmdAceptar_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Me.Hide
End Sub

Private Sub MOSAICO_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 26/05/06
'*************************************************
If LenB(mAncho.Text) = 0 Then mAncho.Text = "0"
If LenB(mLargo.Text) = 0 Then mLargo.Text = "0"
End Sub
