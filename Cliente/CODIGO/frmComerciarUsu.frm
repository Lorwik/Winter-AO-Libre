VERSION 5.00
Begin VB.Form frmComerciarUsu 
   BorderStyle     =   0  'None
   ClientHeight    =   6915
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7350
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   461
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   3930
      Left            =   600
      TabIndex        =   7
      Top             =   1440
      Width           =   2730
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Height          =   3930
      Left            =   3930
      TabIndex        =   5
      Top             =   1440
      Width           =   2730
   End
   Begin VB.OptionButton optQue 
      Caption         =   "Oro"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   5640
      TabIndex        =   4
      Top             =   1080
      Width           =   915
   End
   Begin VB.OptionButton optQue 
      Caption         =   "Objeto"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   4080
      TabIndex        =   3
      Top             =   1080
      Value           =   -1  'True
      Width           =   915
   End
   Begin VB.TextBox txtCant 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
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
      Height          =   195
      Left            =   5160
      TabIndex        =   2
      Text            =   "1"
      Top             =   5550
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   510
      Left            =   600
      ScaleHeight     =   510
      ScaleWidth      =   540
      TabIndex        =   0
      Top             =   345
      Width           =   540
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Height          =   255
      Left            =   1800
      TabIndex        =   6
      Top             =   5550
      Width           =   1335
   End
   Begin VB.Image cmdAceptar 
      Height          =   135
      Left            =   1200
      Top             =   6120
      Width           =   1575
   End
   Begin VB.Image cmdRechazar 
      Height          =   255
      Left            =   1200
      Top             =   6360
      Width           =   1575
   End
   Begin VB.Image Command2 
      Height          =   375
      Left            =   4440
      Top             =   6360
      Width           =   1695
   End
   Begin VB.Image cmdOfrecer 
      Height          =   255
      Left            =   4560
      Top             =   6000
      Width           =   1575
   End
   Begin VB.Label lblEstadoResp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Esperando respuesta..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2280
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   2490
   End
End
Attribute VB_Name = "frmComerciarUsu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************
' frmComerciarUsu.frm
'
'**************************************************************

'**************************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'**************************************************************************

Option Explicit

Private Sub cmdAceptar_Click()
    Call WriteUserCommerceOk
End Sub

Private Sub cmdOfrecer_Click()

If optQue(0).value = True Then
    If List1.ListIndex < 0 Then Exit Sub
    If List1.ItemData(List1.ListIndex) <= 0 Then Exit Sub
    
'    If Val(txtCant.Text) > List1.ItemData(List1.ListIndex) Or _
'        Val(txtCant.Text) <= 0 Then Exit Sub
ElseIf optQue(1).value = True Then
'    If Val(txtCant.Text) > UserGLD Then
'        Exit Sub
'    End If
End If

If optQue(0).value = True Then
    Call WriteUserCommerceOffer(List1.ListIndex + 1, Val(txtCant.Text))
ElseIf optQue(1).value = True Then
    Call WriteUserCommerceOffer(FLAGORO, Val(txtCant.Text))
Else
    Exit Sub
End If

lblEstadoResp.Visible = True
End Sub

Private Sub cmdRechazar_Click()
    Call WriteUserCommerceReject
End Sub

Private Sub Command2_Click()
    Call WriteUserCommerceEnd
End Sub

Private Sub Form_Deactivate()
'Me.SetFocus
'Picture1.SetFocus

End Sub

Private Sub Form_Load()
'Carga las imagenes...?
lblEstadoResp.Visible = False
Me.Picture = General_Load_Picture_From_Resource("82.gif")
End Sub

Private Sub Form_LostFocus()
Me.SetFocus
Picture1.SetFocus

End Sub

Private Sub list1_Click()
    If Inventario.GrhIndex(List1.ListIndex + 1) <> 0 Then
        DibujaGrh Inventario.GrhIndex(List1.ListIndex + 1)
    End If
End Sub

Public Sub DibujaGrh(Grh As Integer)
Call DrawGrhtoHdc(Picture1.hDC, Grh, 0, 0, False)
End Sub

Private Sub List2_Click()
If List2.ListIndex >= 0 And OtroInventario(List2.ListIndex + 1).GrhIndex <> 0 Then
    DibujaGrh OtroInventario(List2.ListIndex + 1).GrhIndex
    Label3.Caption = List2.ItemData(List2.ListIndex)
    cmdAceptar.Enabled = True
    cmdRechazar.Enabled = True
Else
    cmdAceptar.Enabled = False
    cmdRechazar.Enabled = False
End If

End Sub

Private Sub optQue_Click(Index As Integer)
Select Case Index
Case 0
    List1.Enabled = True
Case 1
    List1.Enabled = False
End Select

End Sub

Private Sub txtCant_Change()
    If Val(txtCant.Text) < 1 Then txtCant.Text = "1"
    
    If Val(txtCant.Text) > 2147483647 Then txtCant.Text = "2147483647"
End Sub

Private Sub txtCant_KeyDown(KeyCode As Integer, Shift As Integer)
If Not ((KeyCode >= 48 And KeyCode <= 57) Or KeyCode = vbKeyBack Or _
        KeyCode = vbKeyDelete Or (KeyCode >= 37 And KeyCode <= 40)) Then
    'txtCant = KeyCode
    KeyCode = 0
End If

End Sub

Private Sub txtCant_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = vbKeyBack Or _
        KeyAscii = vbKeyDelete Or (KeyAscii >= 37 And KeyAscii <= 40)) Then
    'txtCant = KeyCode
    KeyAscii = 0
End If

End Sub

'[/Alejo]

