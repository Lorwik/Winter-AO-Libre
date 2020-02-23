VERSION 5.00
Begin VB.Form frmBancoObj 
   BackColor       =   &H80000000&
   BorderStyle     =   0  'None
   ClientHeight    =   7950
   ClientLeft      =   0
   ClientTop       =   -180
   ClientWidth     =   6915
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   530
   ScaleMode       =   0  'User
   ScaleWidth      =   461
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox TCantidad 
      Alignment       =   2  'Center
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
      Height          =   210
      Left            =   3780
      MaxLength       =   7
      TabIndex        =   9
      Text            =   "0"
      Top             =   1560
      Width           =   1080
   End
   Begin VB.TextBox TName 
      Alignment       =   2  'Center
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
      Height          =   210
      Left            =   2400
      MaxLength       =   7
      TabIndex        =   8
      Top             =   1560
      Width           =   1080
   End
   Begin VB.TextBox CantidadOro 
      Alignment       =   2  'Center
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
      Height          =   210
      Left            =   3675
      MaxLength       =   7
      TabIndex        =   7
      Text            =   "1"
      Top             =   1080
      Width           =   600
   End
   Begin VB.TextBox cantidad 
      Alignment       =   2  'Center
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
      Height          =   210
      Left            =   3015
      MaxLength       =   5
      TabIndex        =   6
      Text            =   "1"
      Top             =   4035
      Width           =   870
   End
   Begin VB.PictureBox PicBancoInv 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   3795
      Left            =   285
      ScaleHeight     =   3765
      ScaleWidth      =   2520
      TabIndex        =   4
      Top             =   2370
      Width           =   2550
   End
   Begin VB.PictureBox PicInv 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   3795
      Left            =   4080
      ScaleHeight     =   16.617
      ScaleMode       =   0  'User
      ScaleWidth      =   861.935
      TabIndex        =   3
      Top             =   2400
      Width           =   2535
   End
   Begin VB.Image Image2 
      Height          =   285
      Left            =   5040
      Tag             =   "0"
      Top             =   1500
      Width           =   1425
   End
   Begin VB.Label lblUserGld 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   3855
      TabIndex        =   5
      Top             =   750
      Width           =   135
   End
   Begin VB.Image imgDepositarOro 
      Height          =   210
      Left            =   480
      Tag             =   "0"
      Top             =   1200
      Width           =   1410
   End
   Begin VB.Image imgRetirarOro 
      Height          =   285
      Left            =   5040
      Tag             =   "0"
      Top             =   1170
      Width           =   1425
   End
   Begin VB.Image imgCerrar 
      Height          =   255
      Left            =   6525
      Tag             =   "0"
      Top             =   240
      Width           =   255
   End
   Begin VB.Image Image1 
      Height          =   255
      Index           =   1
      Left            =   3300
      MousePointer    =   99  'Custom
      Top             =   3720
      Width           =   255
   End
   Begin VB.Image Image1 
      Height          =   255
      Index           =   0
      Left            =   3360
      MousePointer    =   99  'Custom
      Top             =   4320
      Width           =   255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Index           =   1
      Left            =   2160
      TabIndex        =   2
      Top             =   6990
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Index           =   2
      Left            =   2160
      TabIndex        =   1
      Top             =   7245
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Index           =   0
      Left            =   2160
      TabIndex        =   0
      Top             =   6750
      Width           =   750
   End
End
Attribute VB_Name = "frmBancoObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LasActionBuy As Boolean
Public LastIndex1 As Integer
Public LastIndex2 As Integer
Public NoPuedeMover As Boolean

Private Sub cantidad_Change()

    If Val(cantidad.Text) < 1 Then
        cantidad.Text = 1
    End If
    
    If Val(cantidad.Text) > MAX_INVENTORY_OBJS Then
        cantidad.Text = MAX_INVENTORY_OBJS
    End If

End Sub

Private Sub cantidad_KeyPress(KeyAscii As Integer)
    If (KeyAscii <> 8) Then
        If (KeyAscii <> 6) And (KeyAscii < 48 Or KeyAscii > 57) Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub CantidadOro_Change()
    If Val(CantidadOro.Text) < 1 Then
        cantidad.Text = 1
    End If
End Sub

Private Sub CantidadOro_KeyPress(KeyAscii As Integer)
    If (KeyAscii <> 8) Then
        If (KeyAscii <> 6) And (KeyAscii < 48 Or KeyAscii > 57) Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub Form_Load()
    'Cargamos la interfaz
    Me.Picture = General_Load_Picture_From_Resource("30.gif")
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Button = vbLeftButton) Then Call Auto_Drag(Me.hwnd)
End Sub
Private Sub Image1_Click(Index As Integer)
    
    Call General_Set_Wav(SND_CLICK)
    
    If InvBanco(Index).SelectedItem = 0 Then Exit Sub
    
    If Not IsNumeric(cantidad.Text) Then Exit Sub
    
    Select Case Index
        Case 0
            LastIndex1 = InvBanco(0).SelectedItem
            LasActionBuy = True
            Call WriteBankExtractItem(InvBanco(0).SelectedItem, cantidad.Text)
            
       Case 1
            LastIndex2 = InvBanco(1).SelectedItem
            LasActionBuy = False
            Call WriteBankDeposit(InvBanco(1).SelectedItem, cantidad.Text)
    End Select

End Sub


Private Sub Image2_Click()
    If TName.Text = "" Then
        MsgBox "Escriba el nombre del usuario al que deseas transferir."
        Exit Sub
    End If
    
    If TCantidad.Text = "" Then
        MsgBox "Escriba la cantidad de oro que deseas transferir."
        Exit Sub
    End If
    
    If TCantidad.Text = "0" Then
        MsgBox "Cantidad invalidad."
        Exit Sub
    End If
    
    Call WriteTransferencia(TName.Text, TCantidad)
End Sub

Private Sub imgDepositarOro_Click()
    Call WriteBankDepositGold(Val(CantidadOro.Text))
End Sub

Private Sub imgRetirarOro_Click()
    Call WriteBankExtractGold(Val(CantidadOro.Text))
End Sub

Private Sub PicBancoInv_Click()

    If InvBanco(0).SelectedItem <> 0 Then
        With UserBancoInventory(InvBanco(0).SelectedItem)
            Label1(0).Caption = .Name
            
            Select Case .OBJType
                Case 2, 32
                    Label1(1).Caption = "Máx Golpe:" & .MaxHit
                    Label1(2).Caption = "Mín Golpe:" & .MinHit
                    Label1(1).Visible = True
                    Label1(2).Visible = True
                    
                Case 3, 16, 17
                    Label1(1).Caption = "Defensa:" & .Def
                    Label1(1).Visible = True
                    Label1(2).Visible = True
                    
                Case Else
                    Label1(1).Visible = False
                    Label1(2).Visible = False
                    
            End Select
            
        End With
        
    Else
        Label1(0).Caption = ""
        Label1(1).Visible = False
        Label1(2).Visible = False
    End If

End Sub

Private Sub PicInv_Click()
    
    If InvBanco(1).SelectedItem <> 0 Then
        With Inventario
            Label1(0).Caption = .ItemName(InvBanco(1).SelectedItem)
            
            Select Case .OBJType(InvBanco(1).SelectedItem)
                Case eObjType.otWeapon, eObjType.otFlechas
                    Label1(1).Caption = "Máx Golpe:" & .MaxHit(InvBanco(1).SelectedItem)
                    Label1(2).Caption = "Mín Golpe:" & .MinHit(InvBanco(1).SelectedItem)
                    Label1(1).Visible = True
                    Label1(2).Visible = True
                    
                Case eObjType.otcasco, eObjType.otArmadura, eObjType.otescudo ' 3, 16, 17
                    Label1(1).Caption = "Defensa:" & .Def(InvBanco(1).SelectedItem)
                    Label1(1).Visible = True
                    Label1(2).Visible = True
                    
                Case Else
                    Label1(1).Visible = False
                    Label1(2).Visible = False
                    
            End Select
            
        End With
    Else
        Label1(0).Caption = ""
        Label1(1).Visible = False
        Label1(2).Visible = False
    End If
End Sub
Private Sub imgCerrar_Click()
    Call WriteBankEnd
    NoPuedeMover = False
End Sub

