VERSION 5.00
Begin VB.Form frmPanelGm 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Panel GM"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   4785
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   4785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Responder"
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   4440
      Width           =   4095
   End
   Begin VB.TextBox txtResp 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1095
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   3120
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.CommandButton cmd1 
      Caption         =   ">> Enviar el mensaje <<"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   4200
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   4920
      Width           =   3855
   End
   Begin VB.ListBox List1 
      Height          =   2010
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   4560
   End
   Begin VB.CommandButton cmdActualiza 
      Caption         =   "&Actualiza"
      Height          =   315
      Left            =   3840
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.ComboBox cboListaUsus 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   3675
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1095
      Left            =   120
      TabIndex        =   8
      Top             =   3120
      Width           =   4575
   End
   Begin VB.Label Label2 
      Caption         =   "0"
      Height          =   255
      Left            =   2160
      TabIndex        =   5
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label11 
      Caption         =   "Usuarios Online:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   600
      Width           =   1695
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   120
      X2              =   4680
      Y1              =   490
      Y2              =   490
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   120
      X2              =   4680
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   120
      X2              =   4680
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Menu mnuUsuario 
      Caption         =   "Usuario"
      Visible         =   0   'False
      Begin VB.Menu mnuBorrar 
         Caption         =   "Borrar mensaje"
      End
      Begin VB.Menu mnuIra 
         Caption         =   "Ir al usuario"
      End
      Begin VB.Menu mnuTraer 
         Caption         =   "Traer el usuario"
      End
   End
   Begin VB.Menu mnuChar 
      Caption         =   "Personaje"
      Begin VB.Menu cmdAccion 
         Caption         =   "Echar"
         Index           =   0
      End
      Begin VB.Menu cmdAccion 
         Caption         =   "Sumonear"
         Index           =   2
      End
      Begin VB.Menu cmdAccion 
         Caption         =   "Ir a"
         Index           =   3
      End
      Begin VB.Menu cmdAccion 
         Caption         =   "Desbanear"
         Index           =   12
      End
      Begin VB.Menu cmdAccion 
         Caption         =   "IP del personaje"
         Index           =   13
      End
      Begin VB.Menu cmdAccion 
         Caption         =   "Revivir"
         Index           =   21
      End
      Begin VB.Menu cmdBan 
         Caption         =   "Banear"
         Begin VB.Menu mnuBan 
            Caption         =   "Personaje"
            Index           =   1
         End
         Begin VB.Menu mnuBan 
            Caption         =   "Personaje e IP"
            Index           =   19
         End
      End
      Begin VB.Menu mnuEncarcelar 
         Caption         =   "Encarcelar"
         Begin VB.Menu mnuCarcel 
            Caption         =   "Encarcelar Usuario"
            Index           =   60
         End
      End
      Begin VB.Menu mnuInfo 
         Caption         =   "Información"
         Begin VB.Menu mnuAccion 
            Caption         =   "General"
            Index           =   8
         End
         Begin VB.Menu mnuAccion 
            Caption         =   "Inventario"
            Index           =   9
         End
         Begin VB.Menu mnuAccion 
            Caption         =   "Skills"
            Index           =   10
         End
         Begin VB.Menu mnuAccion 
            Caption         =   "Bóveda"
            Index           =   18
         End
      End
      Begin VB.Menu mnuSilenciar 
         Caption         =   "Silenciar"
         Begin VB.Menu mnuSilencio 
            Caption         =   "Silenciar Usuario"
            Index           =   60
         End
      End
   End
   Begin VB.Menu cmdHerramientas 
      Caption         =   "Herramientas"
      Begin VB.Menu mnuHerramientas 
         Caption         =   "Insertar comentario"
         Index           =   4
      End
      Begin VB.Menu mnuHerramientas 
         Caption         =   "Enviar hora"
         Index           =   5
      End
      Begin VB.Menu mnuHerramientas 
         Caption         =   "Limpiar Mapa"
         Index           =   15
      End
      Begin VB.Menu mnuHerramientas 
         Caption         =   "Usuarios trabajando"
         Index           =   23
      End
      Begin VB.Menu mnuHerramientas 
         Caption         =   "Bloquear tile"
         Index           =   26
      End
      Begin VB.Menu IP 
         Caption         =   "Direcciónes de IP"
         Index           =   0
         Begin VB.Menu mnuIP 
            Caption         =   "Banear una IP"
            Index           =   17
         End
         Begin VB.Menu mnuIP 
            Caption         =   "Lista de IPs baneadas"
            Index           =   25
         End
      End
   End
   Begin VB.Menu Admin 
      Caption         =   "Administración"
      Index           =   0
      Begin VB.Menu mnuAdmin 
         Caption         =   "Grabar personajes"
         Index           =   28
      End
      Begin VB.Menu mnuAdmin 
         Caption         =   "Iniciar WorldSave"
         Index           =   29
      End
      Begin VB.Menu mnuAdmin 
         Caption         =   "Limpiar el mundo"
         Index           =   34
      End
      Begin VB.Menu Ambiente 
         Caption         =   "Estado climático"
         Index           =   0
         Begin VB.Menu mnuAmbiente 
            Caption         =   "Iniciar o detener una lluvia"
            Index           =   31
         End
         Begin VB.Menu mnuAmbiente 
            Caption         =   "Hacer Amanecer"
            Index           =   36
         End
         Begin VB.Menu mnuAmbiente 
            Caption         =   "Hacer de Dia"
            Index           =   37
         End
         Begin VB.Menu mnuAmbiente 
            Caption         =   "Hacer de Tarde"
            Index           =   38
         End
         Begin VB.Menu mnuAmbiente 
            Caption         =   "Hacer de Noche"
            Index           =   39
         End
      End
   End
   Begin VB.Menu Reload 
      Caption         =   "Reload"
      Index           =   0
      Begin VB.Menu mnuReload 
         Caption         =   "Reload Objetos"
         Index           =   1
      End
      Begin VB.Menu mnuReload 
         Caption         =   "Reload Server.ini"
         Index           =   2
      End
      Begin VB.Menu mnuReload 
         Caption         =   "Reload Mapas"
         Index           =   3
      End
      Begin VB.Menu mnuReload 
         Caption         =   "Reload Hechizos"
         Index           =   4
      End
      Begin VB.Menu mnuReload 
         Caption         =   "Reload Motd"
         Index           =   5
      End
      Begin VB.Menu mnuReload 
         Caption         =   "Reload Npc"
         Index           =   6
      End
      Begin VB.Menu mnuReload 
         Caption         =   "Reload Reiniciar"
         Index           =   7
      End
      Begin VB.Menu mnuReload 
         Caption         =   "Reload Guilds"
         Index           =   8
      End
      Begin VB.Menu mnuReload 
         Caption         =   "Reload BanIP"
         Index           =   9
      End
   End
End
Attribute VB_Name = "frmPanelGm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MAX_GM_MSG = 300

Private MisMSG(0 To MAX_GM_MSG) As String
Private Apunt(0 To MAX_GM_MSG) As Integer

Dim lista As New Collection
Dim Nick As String
Public LastStr As String

Private Sub cmd1_Click()
If txtResp.Text = "" Then
    MsgBox "Debes ingresar una respuesta!"
    Exit Sub
ElseIf Len(txtResp.Text) < 10 Then
    MsgBox "El mensaje debe tener al menos 10 carácteres!"
    Exit Sub
End If

SendData ("/RESPONDER " & ReadField(1, List1.List(List1.listIndex), Asc(";")) & "¬" & txtResp)

txtResp.Visible = False
txtResp.Text = ""
cmd1.Visible = False
SendData ("SOSDONE" & List1.List(List1.listIndex))
List1.RemoveItem List1.listIndex
Label1.Caption = ""
List1.Clear
End Sub

Private Sub cmdAccion_Click(Index As Integer)

Dim tmp As String

Nick = Replace(cboListaUsus.Text, " ", "+")

Select Case Index

Case 0 '/ECHAR nick
    Call SendData("/ECHAR " & Nick)
Case 1 '/ban motivo@nick
    tmp = InputBox("¿Motivo?", "Ingrese el motivo")
    If MsgBox("¿Está seguro que desea banear al personaje """ & cboListaUsus.Text & """?", vbYesNo + vbQuestion) = vbYes Then
        Call SendData("/BAN " & tmp & "@" & cboListaUsus.Text)
    End If
Case 2 '/sum nick
    Call SendData("/SUM " & Nick)
Case 3 '/ira nick
    Call SendData("/IRA " & Nick)
Case 4 '/rem
    tmp = InputBox("¿Comentario?", "Ingrese comentario")
    Call SendData("/REM " & tmp)
Case 5 '/hora
    Call SendData("/HORA")
Case 6 '/donde nick
    Call SendData("/DONDE " & Nick)
Case 7 '/nene
    tmp = InputBox("¿En qué mapa?", "")
    Call SendData("/NENE " & Trim(tmp))
Case 8 '/info nick
    Call SendData("/INFO " & Nick)
Case 9 '/inv nick
    Call SendData("/INV " & Nick)
Case 10 '/skills nick
    Call SendData("/SKILLS " & Nick)
Case 11 '/carcel minutos nick
    tmp = InputBox("¿Minutos a encarcelar? (hasta 60)", "")
    If MsgBox("¿Esta seguro que desea encarcelar al personaje """ & Nick & """?", vbYesNo + vbQuestion) = vbYes Then
        Call SendData("/CARCEL " & tmp & " " & Nick)
    End If
Case 12 '/unban nick
    If MsgBox("¿Esta seguro que desea removerle el ban al personaje """ & Nick & """?", vbYesNo + vbQuestion) = vbYes Then
        Call SendData("/UNBAN " & Nick)
    End If
Case 13 '/nick2ip nick
    Call SendData("/NICK2IP " & Nick)
Case 14 '/sameip nick
    Call SendData("/SAMEIP " & Nick)
Case 15
    tmp = InputBox("¿Mapa?", "")
    Call SendData("/CLEANMAP " & Trim(tmp))
Case 16 '/att nick
    Call SendData("/ATT " & Nick)
Case 17
    tmp = InputBox("Escriba la dirección IP a banear", "")
    If MsgBox("¿Esta seguro que desea banear la IP """ & tmp & """?", vbYesNo + vbQuestion) = vbYes Then
        Call SendData("/BANIP " & tmp)
    End If
Case 18 '/bov nick
    Call SendData("/BOV " & Nick)
Case 19
    If MsgBox("¿Esta seguro que desea banear la IP del personaje """ & Nick & """?", vbYesNo + vbQuestion) = vbYes Then
        Call SendData("/BANIP " & Nick)
    End If
Case 20 '/infofami nick
    Call SendData("/INFOFAMI " & Nick)
Case 21 '/revivir nick
    Call SendData("/REVIVIR " & Nick)
Case 22
    Call SendData("/HMR " & Nick)
Case 23
    Call SendData("/TRABAJANDO")
Case 24
    Call SendData("/ENGRUPO")
Case 25
    Call SendData("/BANIPLIST")
Case 26
    Call SendData("/BLOQ")
Case 27
    Call SendData("/APAGAR")
Case 28
    Call SendData("/GRABAR")
Case 29
    Call SendData("/DOBACKUP")
Case 30
    Call SendData("/ONLINEMAP")
Case 31
    Call SendData("/LLUVIA")
Case 32
    Call SendData("/NOCHE")
Case 33
    Call SendData("/CurrentUser.Pausa")
Case 34
    Call SendData("/LIMPIAR")
Case 35 '/carcel minutos nick
    tmp = InputBox("¿Minutos a silenciar? (hasta 60)", "")
    If MsgBox("¿Esta seguro que desea silenciar al personaje """ & Nick & """?", vbYesNo + vbQuestion) = vbYes Then
        Call SendData("/SILENCIO " & tmp & " " & Nick)
    End If
Case 36
    Call SendData("/AMANECER")
Case 37
    Call SendData("/DIA")
Case 38
    Call SendData("/TARDE")
Case 39
    Call SendData("/NOCHE")
End Select

Nick = ""

End Sub

Private Sub cmdActualiza_Click()
Call SendData("LISTUSU")

End Sub

Private Sub cmdCerrar_Click()
Call MensajeBorrarTodos
Me.Visible = False
List1.Clear
End Sub

Private Sub cmdRepetir_Click()
If LastStr <> "" Then Call SendData(LastStr)
End Sub

Private Sub cmdTarget_Click()
'Call AddtoRichTextBox(frmMain.RecTxt, "Haz click sobre el personaje...", 100, 100, 120, 0, 0)
'frmMain.MousePointer = 2
'frmMain.PanelSelect = True
End Sub

Private Sub cmdOnline_Click()
Call SendData("/ONLINE")
With List1
    .Visible = True
End With


mnuIra.Enabled = True
mnuTraer.Enabled = True

End Sub



Private Sub Command1_Click()
If List1.listIndex < 0 Then Exit Sub

txtResp.Visible = True
txtResp.Text = ""
cmd1.Visible = True
End Sub

Private Sub Form_Load()
List1.Clear
Label2.Caption = frmMain.online.Caption
Call cmdActualiza_Click
Label1.Caption = ""
txtResp.Text = ""

txtResp.Visible = False

End Sub
Public Sub CrearGMmSg(Nick As String, msg As String)
If List1.ListCount < MAX_GM_MSG Then
        List1.AddItem Nick & "-" & List1.ListCount
        MisMSG(List1.ListCount - 1) = msg
        Apunt(List1.ListCount - 1) = List1.ListCount - 1
End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Call MensajeBorrarTodos
Me.Visible = False
List1.Clear
End Sub



Private Sub mnuAccion_Click(Index As Integer)
Call cmdAccion_Click(Index)
End Sub

Private Sub mnuAdmin_Click(Index As Integer)
Call cmdAccion_Click(Index)
End Sub

Private Sub mnuAmbiente_Click(Index As Integer)
Call cmdAccion_Click(Index)
End Sub

Private Sub mnuBan_Click(Index As Integer)
Call cmdAccion_Click(Index)
End Sub

Private Sub mnuCarcel_Click(Index As Integer)

If Index = 60 Then
    Call cmdAccion_Click(11)
    Exit Sub
End If

Call SendData("/CARCEL " & Index & " " & cboListaUsus.Text)

End Sub

Private Sub mnuSilencio_Click(Index As Integer)

If Index = 60 Then
    Call cmdAccion_Click(35)
    Exit Sub
End If

Call SendData("/SILENCIO " & Index & " " & cboListaUsus.Text)

End Sub

Private Sub mnuCompressChars_Click()
    Call SendData("/ZIPCHARS")
End Sub

Private Sub mnuHerramientas_Click(Index As Integer)
Call cmdAccion_Click(Index)
End Sub

Public Sub MensajePoner(ByVal Nick As String, ByVal Mensaje As String)
On Error Resume Next
lista.Add Mensaje, Nick
End Sub

Public Sub MensajeBorrarTodos()
Do While lista.Count > 0
    Call lista.Remove(lista.Count)
Loop
End Sub

Private Sub list1_Click()
On Error Resume Next
Dim ind As Integer
ind = Val(ReadField(2, List1.List(List1.listIndex), Asc(";")))
Label1.Caption = "Usuario: " & List1.Text
End Sub


Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbRightButton Then
    PopupMenu mnuUsuario
End If

End Sub


Private Sub mnuBorrar_Click()

If List1.Visible Then
    If List1.listIndex < 0 Then Exit Sub
    SendData ("SOSDON" & Nick)
    List1.RemoveItem List1.listIndex
End If

End Sub

Private Sub mnuIP_Click(Index As Integer)
Call cmdAccion_Click(Index)
End Sub

Private Sub mnuIRa_Click()

If List1.Visible Then
    SendData ("/IRA " & Nick)
End If

End Sub

Private Sub mnuInvalida_Click()

If List1.Visible Then
    If List1.listIndex < 0 Then Exit Sub
    SendData ("SOSINV" & Nick)
    List1.RemoveItem List1.listIndex
End If

End Sub

Private Sub mnuManual_Click()

If List1.Visible Then
    If List1.listIndex < 0 Then Exit Sub
    SendData ("SOSMAN" & Nick)
    List1.RemoveItem List1.listIndex
End If

End Sub
Private Sub mnuReload_Click(Index As Integer)

Select Case Index
    Case 1 'Reload objetos
        Call SendData("/RELOADOBJ")
    Case 2 'Reload server.ini
        Call SendData("/RELOADSINI")
    Case 3 'Reload mapas
        MsgBox "Deshabilitado"
    Case 4 'Reload hechizos
        Call SendData("/RELOADHECHIZOS")
    Case 5 'Reload motd
        Call SendData("/RELOADMOTD")
    Case 6 'Reload npcs
        Call SendData("/RELOADNPCS")
    Case 7 'Reload sockets
        If MsgBox("¿Estas seguro de que deseas reiniciar?", vbYesNo, "Advertencia") = vbYes Then _
            Call SendData("/REINICIAR")
    Case 9 'Reload Guilds
        Call SendData("/RELOADGUILD")
    Case 10 'Reload otros
        Call SendData("/BANIPRELOAD")
End Select

End Sub

Private Sub mnuStartUp_Click()

Dim TempApp As String
TempApp = InputBox("Ingrese el nombre del ejecutable que desea iniciar en el servidor.", "")
Call SendData("/INICIAR " & TempApp)

End Sub

Private Sub mnuKill_Click()

Dim TempApp As String
TempApp = InputBox("Ingrese el nombre del proceso que desea matar en el servidor.", "")
Call SendData("/KILLAPP " & TempApp)

End Sub

Private Sub mnutraer_Click()

If List1.Visible Then
SendData ("/SUM " & Nick)
Else
SendData ("/SUM " & Nick)
End If
End Sub

Private Sub list1_dblClick()
On Error Resume Next

If List1.Visible Then
    SendData ("/IRA " & Nick)
    SendData ("SOSDON" & Nick)
Else
    SendData ("SOSDON" & Nick)
End If

List1.Clear
Me.Visible = False

End Sub


