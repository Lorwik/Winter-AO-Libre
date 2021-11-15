Attribute VB_Name = "Mod_TCP"
Option Explicit
Public Warping As Boolean
Public LlegaronSkills As Boolean
Public LlegaronAtrib As Boolean
Public LlegoFama As Boolean
Public Function PuedoQuitarFoco() As Boolean
PuedoQuitarFoco = True
End Function

Sub HandleData(ByVal Rdata As String)
    On Error Resume Next
    
    Dim RetVal As Variant
    Dim X As Integer
    Dim Y As Integer
    Dim CharIndex As Integer
    Dim tempint As Integer
    Dim tempstr As String
    Dim slot As Integer
    Dim MapNumber As String
    Dim i As Integer, k As Integer
    Dim cad$, Index As Integer, m As Integer
    Dim T() As String
    
    Dim tstr As String
    Dim tstr2 As String
    
    
    Dim sData As String
    sData = UCase$(Rdata)
        If Left$(sData, 4) = "INVI" Then CartelInvisibilidad = Right$(sData, Len(sData) - 4)
        If Left$(sData, 4) = "INMO" Then CartelParalisis = Right$(sData, Len(sData) - 4)
        
         If UCase$(Left$(Rdata, 2)) = "QL" Then
        Unload frmQuests
        frmQuests.lstQuests.Clear
        
        For i = 1 To 10
            tstr = ReadField(i, Right$(Rdata, Len(Rdata) - 2), Asc("-"))
            
            If tstr = "0" Then
                frmQuests.lstQuests.AddItem "-"
            Else
                frmQuests.lstQuests.AddItem tstr
            End If
        Next i
        
        frmQuests.Show , frmMain
        Exit Sub
    ElseIf UCase$(Left$(Rdata, 2)) = "QI" Then
        tstr = Right$(Rdata, Len(Rdata) - 2)
        
        frmQuests.lblNombre.Caption = ReadField(1, tstr, Asc("-"))
        frmQuests.lblDescripcion.Caption = ReadField(2, tstr, Asc("-"))
        frmQuests.lblCriaturas.Caption = ReadField(3, tstr, Asc("-"))
        Exit Sub
    End If

        
    Select Case sData
    
    Case "BUENO"
            TimerPing(2) = GetTickCount()
            Call AddtoRichTextBox(frmMain.RecTxt, "Ping: " & (TimerPing(2) - TimerPing(1)) & " ms", 255, 0, 0, True, False, False)
    
        Case "LOGGED"            ' >>>>> LOGIN :: LOGGED
        
           Dim normal As Byte
           Dim temp As Byte
           temp = RandomNumber(11, 19)
           frmMain.lblTemp.Caption = temp & "º"
           frmMain.lblTemp.ForeColor = vbGreen

             
            logged = True
            UserCiego = False
            EngineRun = True
            IScombate = False
            UserDescansar = False
            Nombres = True
            If frmCrearPersonaje.Visible Then
                Unload frmPasswdSinPadrinos
                Unload frmCrearPersonaje
                Unload frmConnect
                frmMain.Show
            End If
            Call SetConnected
             
            bTecho = IIf(MapData(UserPos.X, UserPos.Y).Trigger = 1 Or _
            MapData(UserPos.X, UserPos.Y).Trigger = 2 Or _
            MapData(UserPos.X, UserPos.Y).Trigger = 5 Or _
            MapData(UserPos.X, UserPos.Y).Trigger = 7 Or _
            MapData(UserPos.X, UserPos.Y).Trigger = 4, True, False)
            Exit Sub
            
        Case "QTDL"              ' >>>>> Quitar Dialogos :: QTDL
            Call Dialogos.BorrarDialogos
            Exit Sub

 
        Case "NAVEG"
            UserNavegando = Not UserNavegando
            Exit Sub
        Case "ET"
            UserEquitando = Not UserEquitando
            Exit Sub
        Case "FINOK" ' Graceful exit ;))
            frmMain.Socket1.Disconnect
            frmMain.Visible = False
            logged = False
            UserParalizado = False
            IScombate = False
            pausa = False
            UserMeditar = False
            UserDescansar = False
            UserNavegando = False
            UserEquitando = False
            frmConnect.Visible = True
            Call Reloguear
            Call Audio.StopWave
            bRain = False
            bFogata = False
            SkillPoints = 0
            frmMain.Label1.Visible = False
            Call Dialogos.BorrarDialogos
            For i = 1 To LastChar
                charlist(i).invisible = False
            Next i
            
            bK = 0
            Exit Sub
        Case "FINCOMOK"          ' >>>>> Finaliza Comerciar :: FINCOMOK
            frmComerciar.List1(0).Clear
            frmComerciar.List1(1).Clear
            NPCInvDim = 0
            Unload frmComerciar
            Comerciando = False
            Exit Sub
             Case "FINSUBOK"          ' >>>>> Finaliza Comerciar :: FINCOMOK
            frmSubasta.List1(1).Clear
            Unload frmSubasta
            Exit Sub
        '[KEVIN]**************************************************************
        '-----------------------------------------------------------------------------
        Case "FINBANOK"          ' >>>>> Finaliza Banco :: FINBANOK
            frmBancoObj.List1(0).Clear
            frmBancoObj.List1(1).Clear
            NPCInvDim = 0
            Unload frmBancoObj
            Comerciando = False
            Exit Sub
        '[/KEVIN]***********************************************************************
        '------------------------------------------------------------------------------
         Case "INITSUB"           ' >>>>> Inicia Comerciar :: INITCOM
            i = 1
            Do While i <= MAX_INVENTORY_SLOTS
                If Inventario.OBJIndex(i) <> 0 Then
                        frmSubasta.List1(1).AddItem Inventario.ItemName(i)
                Else
                        frmSubasta.List1(1).AddItem "Nada"
                End If
                i = i + 1
            Loop
            'Comerciando = True
            frmSubasta.Show , frmMain
            Exit Sub
        Case "INITCOM"           ' >>>>> Inicia Comerciar :: INITCOM
            i = 1
            Do While i <= MAX_INVENTORY_SLOTS
                If Inventario.OBJIndex(i) <> 0 Then
                        frmComerciar.List1(1).AddItem Inventario.ItemName(i)
                Else
                        frmComerciar.List1(1).AddItem "Nada"
                End If
                i = i + 1
            Loop
            Comerciando = True
             
            frmComerciar.Show , frmMain
            Call Audio.PlayWave(RandomNumber(175, 181))
            
            Delete_File (Windows_Temp_Dir & "175.wav")

            Exit Sub
                
                 
        '[KEVIN]-----------------------------------------------
        '**************************************************************
        'Lorwik
                Case "INITBANKO"
            frmBanco.Show , frmMain
            Exit Sub
        Case "INITBANCO"           ' >>>>> Inicia Comerciar :: INITBANCO
            Dim II As Integer
            II = 1
            Do While II <= MAX_INVENTORY_SLOTS
                If Inventario.OBJIndex(II) <> 0 Then
                        frmBancoObj.List1(1).AddItem Inventario.ItemName(II)
                Else
                        frmBancoObj.List1(1).AddItem "Nada"
                End If
                II = II + 1
            Loop
            
            
            i = 1
            Do While i <= UBound(UserBancoInventory)
                If UserBancoInventory(i).OBJIndex <> 0 Then
                        frmBancoObj.List1(0).AddItem UserBancoInventory(i).name
                Else
                        frmBancoObj.List1(0).AddItem "Nada"
                End If
                i = i + 1
            Loop
            Comerciando = True
            frmBancoObj.Show , frmMain
            Exit Sub
        '---------------------------------------------------------------
        '[/KEVIN]******************
        '[Alejo]
        Case "INITCOMUSU"
            If frmComerciarUsu.List1.ListCount > 0 Then frmComerciarUsu.List1.Clear
            If frmComerciarUsu.List2.ListCount > 0 Then frmComerciarUsu.List2.Clear
            
            For i = 1 To MAX_INVENTORY_SLOTS
                If Inventario.OBJIndex(i) <> 0 Then
                        frmComerciarUsu.List1.AddItem Inventario.ItemName(i)
                        frmComerciarUsu.List1.ItemData(frmComerciarUsu.List1.NewIndex) = Inventario.Amount(i)
                Else
                        frmComerciarUsu.List1.AddItem "Nada"
                        frmComerciarUsu.List1.ItemData(frmComerciarUsu.List1.NewIndex) = 0
                End If
            Next i
            Comerciando = True
            frmComerciarUsu.Show , frmMain
        Case "FINCOMUSUOK"
            frmComerciarUsu.List1.Clear
            frmComerciarUsu.List2.Clear
            
            Unload frmComerciarUsu
            Comerciando = False
            '[/Alejo]
        Case "RECPASSOK"
            Call MsgBox("¡¡¡El password fue enviado con éxito!!!", vbApplicationModal + vbDefaultButton1 + vbInformation + vbOKOnly, "Envio de password")

            frmMain.Socket1.Disconnect


            Exit Sub
         Case "RECPASSER"
            Call MsgBox("¡¡¡No coinciden los datos con los del personaje en el servidor, el password no ha sido enviado.!!!", vbApplicationModal + vbDefaultButton1 + vbInformation + vbOKOnly, "Envio de password")

            frmMain.Socket1.Disconnect


            Exit Sub
        Case "SFH"
            frmHerrero.Show , frmMain
            Exit Sub
        Case "SFC"
            frmCarp.Show , frmMain
            Exit Sub
        Case "N1" ' <--- Npc ataco y fallo
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_CRIATURA_FALLA_GOLPE, 255, 0, 0, True, False, False)
            Exit Sub
        Case "6" ' <--- Npc mata al usuario
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_CRIATURA_MATADO, 255, 0, 0, True, False, False)
            Exit Sub
        Case "7" ' <--- Ataque rechazado con el escudo
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_RECHAZO_ATAQUE_ESCUDO, 255, 0, 0, True, False, False)
            Exit Sub
        Case "8" ' <--- Ataque rechazado con el escudo
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_USUARIO_RECHAZO_ATAQUE_ESCUDO, 255, 0, 0, True, False, False)
            Exit Sub
        Case "U1" ' <--- User ataco y fallo el golpe
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_FALLADO_GOLPE, 255, 0, 0, True, False, False)
            Exit Sub
        Case "SEGON" '  <--- Activa el seguro
            Call frmMain.DibujarSeguro
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_SEGURO_ACTIVADO, 0, 255, 0, True, False, False)
            Exit Sub
        Case "SEGOFF" ' <--- Desactiva el seguro
            Call frmMain.DesDibujarSeguro
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_SEGURO_DESACTIVADO, 255, 0, 0, True, False, False)
            Exit Sub
        Case "PN"     ' <--- Pierde Nobleza
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PIERDE_NOBLEZA, 255, 0, 0, False, False, False)
            Exit Sub
        Case "M!"     ' <--- Usa meditando
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_USAR_MEDITANDO, 255, 0, 0, False, False, False)
            Exit Sub
    End Select

    Select Case Left(sData, 1)
    Case "³"
      Rdata = Right$(Rdata, Len(Rdata) - 1)
      NumUsers = Rdata
      Exit Sub
        Case "+"              ' >>>>> Mover Char >>> +
            Rdata = Right$(Rdata, Len(Rdata) - 1)


            CharIndex = Val(ReadField(1, Rdata, Asc(",")))
            X = Val(ReadField(2, Rdata, Asc(",")))
            Y = Val(ReadField(3, Rdata, Asc(",")))


            'antigua codificacion del mensaje (decodificada x un chitero)
            'CharIndex = Asc(Mid$(Rdata, 1, 1)) * 64 + (Asc(Mid$(Rdata, 2, 1)) And &HFC&) / 4

            ' CONSTANTES TODO: De donde sale el 40-49 ?
            
            If charlist(CharIndex).Fx >= 40 And charlist(CharIndex).Fx <= 49 Then   'si esta meditando
                charlist(CharIndex).Fx = 0
                charlist(CharIndex).FxLoopTimes = 0
            End If
            
            ' CONSTANTES TODO: Que es .priv ?
            
            If charlist(CharIndex).priv = 0 Then
                Call DoPasosFx(CharIndex)
            End If

            Call MoveCharbyPos(CharIndex, X, Y)
            
            Call RefreshAllChars
            Exit Sub
        Case "*", "_"             ' >>>>> Mover NPC >>> *
            Rdata = Right$(Rdata, Len(Rdata) - 1)
            

            CharIndex = Val(ReadField(1, Rdata, Asc(",")))
            X = Val(ReadField(2, Rdata, Asc(",")))
            Y = Val(ReadField(3, Rdata, Asc(",")))

            
            'antigua codificacion del mensaje (decodificada x un chitero)
            'CharIndex = Asc(Mid$(Rdata, 1, 1)) * 64 + (Asc(Mid$(Rdata, 2, 1)) And &HFC&) / 4
            
'            If charlist(CharIndex).Body.Walk(1).GrhIndex = 4747 Then
'                Debug.Print "hola"
'            End If
            
            ' CONSTANTES TODO: De donde sale el 40-49 ?
            
            If charlist(CharIndex).Fx >= 40 And charlist(CharIndex).Fx <= 49 Then   'si esta meditando
                charlist(CharIndex).Fx = 0
                charlist(CharIndex).FxLoopTimes = 0
            End If
            
            ' CONSTANTES TODO: Que es .priv ?
            
            If charlist(CharIndex).priv = 0 Then
                Call DoPasosFx(CharIndex)
            End If
            
            Call MoveCharbyPos(CharIndex, X, Y)
            'Call MoveCharbyPos(CharIndex, Asc(Mid$(Rdata, 3, 1)), Asc(Mid$(Rdata, 4, 1)))
            
            Call RefreshAllChars
            Exit Sub
    
    End Select
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Select Case Left$(sData, 2)
        Case "AS"
            tstr = mid$(sData, 3, 1)
            k = Val(Right$(sData, Len(sData) - 3))
            
            Select Case tstr
                Case "M": UserMinMAN = Val(Right$(sData, Len(sData) - 3))
                Case "H": UserMinHP = Val(Right$(sData, Len(sData) - 3))
                Case "S": UserMinSTA = Val(Right$(sData, Len(sData) - 3))
                Case "G": UserGLD = Val(Right$(sData, Len(sData) - 3))
                Case "E": UserExp = Val(Right$(sData, Len(sData) - 3))
            End Select
            
            frmMain.exp.Caption = UserExp & "/" & UserPasarNivel
frmMain.ExpShp.Width = (((UserExp / 100) / (UserPasarNivel / 100)) * 111)
 
            frmMain.lblPorcLvl.Caption = Round(CDbl(UserExp) * CDbl(100) / CDbl(UserPasarNivel), 2) & "%"
            frmMain.Hpshp.Width = (((UserMinHP / 100) / (UserMaxHP / 100)) * 94)
             frmMain.HpBar.Caption = "" & UserMinHP & " / " & UserMaxHP & ""
             
            If UserMaxMAN > 0 Then
                frmMain.MANShp.Width = (((UserMinMAN + 1 / 100) / (UserMaxMAN + 1 / 100)) * 94)
            Else
                frmMain.MANShp.Width = 0
            End If
            frmMain.ManaBar.Caption = "" & UserMinMAN & " / " & UserMaxMAN & ""
            frmMain.STAShp.Width = (((UserMinSTA / 100) / (UserMaxSTA / 100)) * 94)
        
            frmMain.GldLbl.Caption = UserGLD
            frmMain.GldLbl2.Caption = CalculateK(UserGLD)
            frmMain.LvlLbl.Caption = UserLvl
            
            If UserMinHP = 0 Then
                UserEstado = 1
            Else
                UserEstado = 0
            End If
            
            Exit Sub
            
        Case "CM"              ' >>>>> Cargar Mapa :: CM
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            UserMap = ReadField(1, Rdata, 44)
            'Obtiene la version del mapa

            
            If FileExist(DirMapas & "Mapa" & UserMap & ".map", vbNormal) Then
                Open DirMapas & "Mapa" & UserMap & ".map" For Binary As #1
                Seek #1, 1
                Get #1, , tempint
                Close #1
                    'Si es la vers correcta cambiamos el mapa
                    Call SwitchMap(UserMap)
            Else
                'no encontramos el mapa en el hd
                MsgBox "Error en los mapas, algun archivo ha sido modificado o esta dañado."
                Call LiberarObjetosDX
                Call UnloadAllForms
                End
            End If
            Exit Sub
        
        Case "PU"                 ' >>>>> Actualiza Posición Usuario :: PU
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            MapData(UserPos.X, UserPos.Y).CharIndex = 0
            UserPos.X = CInt(ReadField(1, Rdata, 44))
            UserPos.Y = CInt(ReadField(2, Rdata, 44))
            MapData(UserPos.X, UserPos.Y).CharIndex = UserCharIndex
            charlist(UserCharIndex).Pos = UserPos
            Call DibujarMiniMapaUser
            frmMain.Coord.Caption = UserMap
            frmMain.Coord2.Caption = UserPos.X
            frmMain.coord3.Caption = UserPos.Y
            Exit Sub
        
        Case "N2" ' <<--- Npc nos impacto (Ahorramos ancho de banda)
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            i = Val(ReadField(1, Rdata, 44))
            Select Case i
                Case bCabeza
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_CABEZA & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bBrazoIzquierdo
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_BRAZO_IZQ & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bBrazoDerecho
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_BRAZO_DER & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bPiernaIzquierda
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_PIERNA_IZQ & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bPiernaDerecha
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_PIERNA_DER & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bTorso
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_TORSO & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
            End Select
            Exit Sub
        Case "U2" ' <<--- El user ataco un npc e impacato
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_CRIATURA_1 & Rdata & MENSAJE_2, 255, 0, 0, True, False, False)
            Exit Sub
        Case "U3" ' <<--- El user ataco un user y falla
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & Rdata & MENSAJE_ATAQUE_FALLO, 255, 0, 0, True, False, False)
            Exit Sub
        Case "N4" ' <<--- user nos impacto
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            i = Val(ReadField(1, Rdata, 44))
            Select Case i
                Case bCabeza
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & ReadField(3, Rdata, 44) & MENSAJE_RECIVE_IMPACTO_CABEZA & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bBrazoIzquierdo
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & ReadField(3, Rdata, 44) & MENSAJE_RECIVE_IMPACTO_BRAZO_IZQ & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bBrazoDerecho
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & ReadField(3, Rdata, 44) & MENSAJE_RECIVE_IMPACTO_BRAZO_DER & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bPiernaIzquierda
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & ReadField(3, Rdata, 44) & MENSAJE_RECIVE_IMPACTO_PIERNA_IZQ & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bPiernaDerecha
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & ReadField(3, Rdata, 44) & MENSAJE_RECIVE_IMPACTO_PIERNA_DER & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bTorso
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & ReadField(3, Rdata, 44) & MENSAJE_RECIVE_IMPACTO_TORSO & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
            End Select
            Exit Sub
        Case "N5" ' <<--- impactamos un user
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            i = Val(ReadField(1, Rdata, 44))
            Select Case i
                Case bCabeza
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & ReadField(3, Rdata, 44) & MENSAJE_PRODUCE_IMPACTO_CABEZA & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bBrazoIzquierdo
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & ReadField(3, Rdata, 44) & MENSAJE_PRODUCE_IMPACTO_BRAZO_IZQ & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bBrazoDerecho
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & ReadField(3, Rdata, 44) & MENSAJE_PRODUCE_IMPACTO_BRAZO_DER & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bPiernaIzquierda
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & ReadField(3, Rdata, 44) & MENSAJE_PRODUCE_IMPACTO_PIERNA_IZQ & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bPiernaDerecha
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & ReadField(3, Rdata, 44) & MENSAJE_PRODUCE_IMPACTO_PIERNA_DER & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bTorso
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & ReadField(3, Rdata, 44) & MENSAJE_PRODUCE_IMPACTO_TORSO & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
            End Select
            Exit Sub
              Case "|?"
            Call frmGMAyuda.Show(vbModeless, frmMain)
            Exit Sub
            Case "|G" 'Guerras
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            If Rdata = 1 Then UserGuerra = True: frmMain.GuerraPos.Visible = True
            If Rdata = 0 Then UserGuerra = False: frmMain.GuerraPos.Visible = False
            Exit Sub
        Case "||"                 ' >>>>> Dialogo de Usuarios y NPCs :: ||
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            Dim iuser As Integer
            iuser = Val(ReadField(3, Rdata, 176))
            
            If iuser > 0 Then
                Dialogos.CrearDialogo ReadField(2, Rdata, 176), iuser, Val(ReadField(1, Rdata, 176))
            Else
                If PuedoQuitarFoco Then
                    AddtoRichTextBox frmMain.RecTxt, ReadField(1, Rdata, 126), Val(ReadField(2, Rdata, 126)), Val(ReadField(3, Rdata, 126)), Val(ReadField(4, Rdata, 126)), Val(ReadField(5, Rdata, 126)), Val(ReadField(6, Rdata, 126))
                End If
            End If

            Exit Sub
        Case "|+"                 ' >>>>> Consola de clan y NPCs :: |+
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            
            iuser = Val(ReadField(3, Rdata, 176))

            If iuser = 0 Then
                If PuedoQuitarFoco And Not DialogosClanes.Activo Then
                    AddtoRichTextBox frmMain.RecTxt, ReadField(1, Rdata, 126), Val(ReadField(2, Rdata, 126)), Val(ReadField(3, Rdata, 126)), Val(ReadField(4, Rdata, 126)), Val(ReadField(5, Rdata, 126)), Val(ReadField(6, Rdata, 126))
                ElseIf DialogosClanes.Activo Then
                    DialogosClanes.PushBackText ReadField(1, Rdata, 126)
                End If
            End If

            Exit Sub

        Case "!!"                ' >>>>> Msgbox :: !!
            If PuedoQuitarFoco Then
                Rdata = Right$(Rdata, Len(Rdata) - 2)
                frmMensaje.msg.Caption = Rdata
                frmMensaje.Show
            End If
            Exit Sub
        Case "IU"                ' >>>>> Indice de Usuario en Server :: IU
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            UserIndex = Val(Rdata)
            Exit Sub
        Case "IP"                ' >>>>> Indice de Personaje de Usuario :: IP
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            UserCharIndex = Val(Rdata)
            UserPos = charlist(UserCharIndex).Pos
             Call DibujarMiniMapaUser
            frmMain.Coord.Caption = UserMap
            frmMain.Coord2.Caption = UserPos.X
            frmMain.coord3.Caption = UserPos.Y
            Exit Sub
        Case "CC"              ' >>>>> Crear un Personaje :: CC
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            CharIndex = ReadField(4, Rdata, 44)
            X = ReadField(5, Rdata, 44)
            Y = ReadField(6, Rdata, 44)
            charlist(CharIndex).Aura_Index = Val(ReadField(2, Rdata, 44))
            charlist(CharIndex).Fx = Val(ReadField(9, Rdata, 44))
            charlist(CharIndex).FxLoopTimes = Val(ReadField(10, Rdata, 44))
            charlist(CharIndex).Nombre = ReadField(12, Rdata, 44)
            charlist(CharIndex).Criminal = Val(ReadField(13, Rdata, 44))
            charlist(CharIndex).priv = Val(ReadField(14, Rdata, 44))
            
            Call MakeChar(CharIndex, ReadField(1, Rdata, 44), ReadField(2, Rdata, 44), ReadField(3, Rdata, 44), X, Y, Val(ReadField(7, Rdata, 44)), Val(ReadField(8, Rdata, 44)), Val(ReadField(11, Rdata, 44)))
            
            Call RefreshAllChars
            Exit Sub
            
        Case "BP"             ' >>>>> Borrar un Personaje :: BP
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            Call EraseChar(Val(Rdata))
            Call Dialogos.QuitarDialogo(Val(Rdata))
            Call RefreshAllChars
            Exit Sub
        Case "MP"             ' >>>>> Mover un Personaje :: MP
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            CharIndex = Val(ReadField(1, Rdata, 44))
            
            If charlist(CharIndex).Fx >= 40 And charlist(CharIndex).Fx <= 49 Then   'si esta meditando
                charlist(CharIndex).Fx = 0
                charlist(CharIndex).FxLoopTimes = 0
            End If
            
            If charlist(CharIndex).priv = 0 Then
                Call DoPasosFx(CharIndex)
            End If
            
            Call MoveCharbyPos(CharIndex, ReadField(2, Rdata, 44), ReadField(3, Rdata, 44))
            
            Call RefreshAllChars
            Exit Sub
        Case "CP"             ' >>>>> Cambiar Apariencia Personaje :: CP
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            
            CharIndex = Val(ReadField(1, Rdata, 44))
            charlist(CharIndex).muerto = Val(ReadField(3, Rdata, 44)) = 500
            charlist(CharIndex).Body = BodyData(Val(ReadField(2, Rdata, 44)))
            charlist(CharIndex).Head = HeadData(Val(ReadField(3, Rdata, 44)))
            charlist(CharIndex).Heading = Val(ReadField(4, Rdata, 44))
            charlist(CharIndex).Fx = Val(ReadField(7, Rdata, 44))
            charlist(CharIndex).FxLoopTimes = Val(ReadField(8, Rdata, 44))
            tempint = Val(ReadField(5, Rdata, 44))
            If tempint <> 0 Then charlist(CharIndex).Arma = WeaponAnimData(tempint)
            tempint = Val(ReadField(6, Rdata, 44))
            If tempint <> 0 Then charlist(CharIndex).Escudo = ShieldAnimData(tempint)
            tempint = Val(ReadField(9, Rdata, 44))
            If tempint <> 0 Then charlist(CharIndex).Casco = CascoAnimData(tempint)

            Call RefreshAllChars
            Exit Sub
        Case "HO"            ' >>>>> Crear un Objeto
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            X = Val(ReadField(2, Rdata, 44))
            Y = Val(ReadField(3, Rdata, 44))
            'ID DEL OBJ EN EL CLIENTE
            MapData(X, Y).ObjGrh.GrhIndex = Val(ReadField(1, Rdata, 44))
            MapData(X, Y).ObjName = ReadField(4, Rdata, 44)
            InitGrh MapData(X, Y).ObjGrh, MapData(X, Y).ObjGrh.GrhIndex
            Exit Sub
        Case "BO"           ' >>>>> Borrar un Objeto
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            X = Val(ReadField(1, Rdata, 44))
            Y = Val(ReadField(2, Rdata, 44))
            MapData(X, Y).ObjGrh.GrhIndex = 0
            MapData(X, Y).ObjName = ""
            Exit Sub
        Case "BQ"           ' >>>>> Bloquear Posición
            Dim b As Byte
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            MapData(Val(ReadField(1, Rdata, 44)), Val(ReadField(2, Rdata, 44))).Blocked = Val(ReadField(3, Rdata, 44))
            Exit Sub
            Case "N~"           ' >>>>> Nombre del Mapa
Rdata = Right$(Rdata, Len(Rdata) - 2)
frmMain.lblMapaName.Caption = Rdata
Exit Sub
        Case "TM"           ' >>>>> Play un MIDI :: TM
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            currentMidi = Val(ReadField(1, Rdata, 45))
            
            If Musica Then
                If currentMidi <> 0 Then
                    Rdata = Right$(Rdata, Len(Rdata) - Len(ReadField(1, Rdata, 45)))
                    If Len(Rdata) > 0 Then
                    Call Extract_File2(Midi, App.Path & "\ARCHIVOS", CStr(currentMidi) & ".mid", Windows_Temp_Dir, False)
                        Call Audio.PlayMIDI(CStr(currentMidi) & ".mid", Val(Right$(Rdata, Len(Rdata) - 1)))
                        Delete_File (Windows_Temp_Dir & CStr(currentMidi) & ".mid")
                    Else
                    Call Extract_File2(Midi, App.Path & "\ARCHIVOS", CStr(currentMidi) & ".mid", Windows_Temp_Dir, False)
                        Call Audio.PlayMIDI(CStr(currentMidi) & ".mid")
                        Delete_File (Windows_Temp_Dir & CStr(currentMidi) & ".mid")
                    End If
                End If
            End If
            Exit Sub

    
        Case "TW"          ' >>>>> Play un WAV :: TW
            If Sound Then
                Rdata = Right$(Rdata, Len(Rdata) - 2)
                 Call Audio.PlayWave(Rdata & ".wav")
            End If
            Exit Sub

Case "PF" ' >>>>> Label de fuerza
Rdata = Right$(Rdata, Len(Rdata) - 2)
frmMain.Fuerza.Caption = Rdata

Exit Sub

Case "PG" ' >>>>> Label de Agilidad
Rdata = Right$(Rdata, Len(Rdata) - 2)
frmMain.Agilidad.Caption = Rdata

Exit Sub

Exit Sub

        Case "GL" 'Lista de guilds
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            Call frmGuildAdm.ParseGuildList(Rdata)
            Exit Sub
        Case "FO"          ' >>>>> Play un WAV :: TW
            bFogata = True
            If FogataBufferIndex = 0 Then
                FogataBufferIndex = Audio.PlayWave("fuego.wav", LoopStyle.Enabled)
              
            End If
            Exit Sub
        Case "CA"
            CambioDeArea Asc(mid$(sData, 3, 1)), Asc(mid$(sData, 4, 1))
            Exit Sub
    End Select
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Select Case Left$(sData, 3)
    
    Case "FRX"
Dim frio As Byte
temp = RandomNumber(1, 10)
frmMain.lblTemp.Caption = temp & "º"
frmMain.lblTemp.ForeColor = vbBlue

Case "CLX"
   Dim calor As Byte
temp = RandomNumber(20, 40)
frmMain.lblTemp.Caption = temp & "º"
frmMain.lblTemp.ForeColor = vbRed

Case "CRT"
If PanelsitoY Then
Exit Sub
Else
Call frmUsu.Show(vbModeless, frmMain)
End If

    Case "LWK" ' Crea Aura Sobre El Char
        Rdata = Right$(Rdata, Len(Rdata) - 3)
        CharIndex = Val(ReadField(1, Rdata, 44))
        charlist(CharIndex).Aura_Index = Val(ReadField(2, Rdata, 44))
        Call InitGrh(charlist(CharIndex).Aura, Val(ReadField(2, Rdata, 44)))

    Case "VAL"                  ' >>>>> Validar Cliente :: VAL
         Dim ValString As String
         Rdata = Right$(Rdata, Len(Rdata) - 3)
         bK = CLng(ReadField(1, Rdata, Asc(",")))
         bRK = ReadField(2, Rdata, Asc(","))
         ValString = ReadField(3, Rdata, Asc(","))
         CargarCabezas
            

     If EstadoLogin = normal Or EstadoLogin = CrearNuevoPj Or EstadoLogin = loginaccount Then
            Call login(ValidarLoginMSG(CInt(bRK)))
            ElseIf EstadoLogin = Dados Then
                frmCrearPersonaje.Show vbModal
            ElseIf EstadoLogin = CrearAccount Then
                frmCrearAccount.Show vbModal
            End If
    Exit Sub
    
        Case "BKW"                  ' >>>>> Pausa :: BKW
            pausa = Not pausa
            Exit Sub
            
       
                                    Case "ZRE"
                                    If MPTres Then
                            'Lorwik - Si el usuario murio reproducimos el MP3 y paramos el midi
                            Musica = False
                 Audio.StopMidi
                                Windows_Temp_Dir = General_Get_Temp_Dir
                               
 Set MP3P = New clsMP3Player
    Call Extract_File2(mp3, App.Path & "\ARCHIVOS\", "2.mp3", Windows_Temp_Dir, False)
    MP3P.stopMP3
    MP3P.mp3file = Windows_Temp_Dir & "2.mp3"
    MP3P.playMP3
    MP3P.Volume = 1000
    Delete_File (Windows_Temp_Dir & "2.mp3")
    End If
    Exit Sub
    Case "XPR"
    MP3P.stopMP3
    Musica = True
                            Call Extract_File2(Midi, App.Path & "\ARCHIVOS", CStr(currentMidi) & ".mid", Windows_Temp_Dir, False)
    Call Audio.PlayMIDI(CStr(currentMidi) & ".mid")
    Delete_File (Windows_Temp_Dir & CStr(currentMidi) & ".mid")
    Exit Sub
        Case "QDL"                  ' >>>>> Quitar Dialogo :: QDL
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            Call Dialogos.QuitarDialogo(Val(Rdata))
            Exit Sub
        Case "CFX"                  ' >>>>> Mostrar FX sobre Personaje :: CFX
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            CharIndex = Val(ReadField(1, Rdata, 44))
            charlist(CharIndex).Fx = Val(ReadField(2, Rdata, 44))
            charlist(CharIndex).FxLoopTimes = Val(ReadField(3, Rdata, 44))
            Exit Sub
        Case "AYM"                  ' >>>>> Pone Mensaje en Cola GM :: AYM
            Dim N As String, n2 As String
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            N = ReadField(2, Rdata, 176)
            n2 = ReadField(1, Rdata, 176)
            frmPanelGm.CrearGMmSg N, n2
            frmPanelGm.Show , frmMain
            Exit Sub
                    Case "ARM" ' fuerza y armaduras/escus/cascos en labels
        Rdata = Right$(Rdata, Len(Rdata) - 3)
       
        With frmMain
                .Arma = ReadField(1, Rdata, Asc(","))
                .Armadura = ReadField(2, Rdata, Asc(","))
                .Casco = ReadField(3, Rdata, Asc(","))
                .Escudo = ReadField(4, Rdata, Asc(","))
        End With
'Lorwik> Este es el sistema noche de WAO Clasico
       Case "NUB"
        Rdata = Right$(Rdata, Len(Rdata) - 3)
       If Rdata = 1 Then
            Amanecer = 0
            Atardecer = 0
            Anochecer = 1
       End If
       If Rdata = 0 Then
            Anochecer = 0
       End If
       Exit Sub
 
       Case "MAÑ" 'Mañana
        Rdata = Right$(Rdata, Len(Rdata) - 3)
       If Rdata = True Then
            Amanecer = 1
            Atardecer = 0
            Anochecer = 0
       End If
       If Rdata = False Then
            Amanecer = False
       End If
       Exit Sub
 
       Case "TAR" 'Tarde
       Rdata = Right$(Rdata, Len(Rdata) - 3)
       If Rdata = 1 Then
          Amanecer = 0
          Atardecer = 1
          Anochecer = 0
       End If
       If Rdata = 0 Then
          Atardecer = 0
       End If
       Exit Sub
 
        Case "MDI" 'Dia
        Rdata = Right$(Rdata, Len(Rdata) - 3)
         If Rdata = 1 Then
            Amanecer = 0
            Atardecer = 0
            Anochecer = 0
         End If
         Exit Sub
            
        Case "EST"                  ' >>>>> Actualiza Estadisticas de Usuario :: EST
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            UserMaxHP = Val(ReadField(1, Rdata, 44))
            UserMinHP = Val(ReadField(2, Rdata, 44))
            UserMaxMAN = Val(ReadField(3, Rdata, 44))
            UserMinMAN = Val(ReadField(4, Rdata, 44))
            UserMaxSTA = Val(ReadField(5, Rdata, 44))
            UserMinSTA = Val(ReadField(6, Rdata, 44))
            frmMain.HpBar.Caption = "" & UserMinHP & " / " & UserMaxHP & ""
            frmMain.ManaBar.Caption = "" & UserMinMAN & " / " & UserMaxMAN & ""
            frmMain.StaBar.Caption = "" & UserMinSTA & " / " & UserMaxSTA & ""
            UserGLD = Val(ReadField(7, Rdata, 44))
            UserLvl = Val(ReadField(8, Rdata, 44))
            UserPasarNivel = Val(ReadField(9, Rdata, 44))
            UserExp = Val(ReadField(10, Rdata, 44))
                        'Lorwik
            UserGLDBOV = Val(ReadField(11, Rdata, 44))
            UserBOVItem = Val(ReadField(12, Rdata, 44))
           
            If frmBanco.Visible Then
                frmBanco.lblInfo.Caption = "Bienvenido a la cadena de finanzas Goliath. Tienes " & UserGLD & " monedas de oro en tu billetera y en tu cuenta tienes " & UserGLDBOV & " Monedas de oro. y " & UserBOVItem & " items en tu Boveda. ¿Cómo te puedo ayudar?"
            End If
            frmMain.exp.Caption = UserExp & "/" & UserPasarNivel
            frmMain.ExpShp.Width = (((UserExp / 100) / (UserPasarNivel / 100)) * 111)
 
            frmMain.lblPorcLvl.Caption = Round(CDbl(UserExp) * CDbl(100) / CDbl(UserPasarNivel), 2) & "%"
            frmMain.Hpshp.Width = (((UserMinHP / 100) / (UserMaxHP / 100)) * 94)
            
            If UserMaxMAN > 0 Then
                frmMain.MANShp.Width = (((UserMinMAN + 1 / 100) / (UserMaxMAN + 1 / 100)) * 94)
            Else
                frmMain.MANShp.Width = 0
            End If
            
            frmMain.STAShp.Width = (((UserMinSTA / 100) / (UserMaxSTA / 100)) * 94)
        
            frmMain.GldLbl.Caption = UserGLD
            frmMain.GldLbl2.Caption = CalculateK(UserGLD)
            frmMain.LvlLbl.Caption = UserLvl
            
            If UserMinHP = 0 Then
                UserEstado = 1
            Else
                UserEstado = 0
            End If
        
            Exit Sub
            
        Case "T01"                  ' >>>>> TRABAJANDO :: TRA
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            UsingSkill = Val(Rdata)
            frmMain.MousePointer = 2
            Select Case UsingSkill
                Case Magia
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_MAGIA, 100, 100, 120, 0, 0)
                Case Pesca
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_PESCA, 100, 100, 120, 0, 0)
                Case Robar
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_ROBAR, 100, 100, 120, 0, 0)
                Case Talar
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_TALAR, 100, 100, 120, 0, 0)
                Case Mineria
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_MINERIA, 100, 100, 120, 0, 0)
                Case FundirMetal
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_FUNDIRMETAL, 100, 100, 120, 0, 0)
                Case Proyectiles
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_PROYECTILES, 100, 100, 120, 0, 0)
            End Select
            Exit Sub
            'Lorwik> Sistema de textos en el cliente extraido de EAO.(Cuando las cosas no son mias las digo ¬¬)
        Case "PRE"
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            Call txtReceived(ReadField(1, Rdata, 44), ReadField(2, Rdata, 44), ReadField(3, Rdata, 44), ReadField(4, Rdata, 44), ReadField(5, Rdata, 44), ReadField(6, Rdata, 44))
            Exit Sub
        Case "PRB"
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            Call txtReceivedB(ReadField(1, Rdata, 44), ReadField(2, Rdata, 44), ReadField(3, Rdata, 44), ReadField(4, Rdata, 44), ReadField(5, Rdata, 44), ReadField(6, Rdata, 44))
            Exit Sub
        Case "PRT"
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            Call txtReceivedT(ReadField(1, Rdata, 44), ReadField(2, Rdata, 44), ReadField(3, Rdata, 44), ReadField(4, Rdata, 44), ReadField(5, Rdata, 44), ReadField(6, Rdata, 44))
            Exit Sub
         '/Lorwik> Aqui se acaba
            
        Case "CSI"                 ' >>>>> Actualiza Slot Inventario :: CSI
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            slot = ReadField(1, Rdata, 44)
            
            If ReadField(3, Rdata, 44) = "(None)" Then
                Inventario.SetItem slot, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0
            Else
            
            Call Inventario.SetItem(slot, ReadField(2, Rdata, 44), ReadField(4, Rdata, 44), ReadField(5, Rdata, 44), Val(ReadField(6, Rdata, 44)), Val(ReadField(7, Rdata, 44)), _
                                    Val(ReadField(8, Rdata, 44)), Val(ReadField(9, Rdata, 44)), Val(ReadField(10, Rdata, 44)), Val(ReadField(11, Rdata, 44)), ReadField(3, Rdata, 44), ReadField(12, Rdata, 44))
End If
            


            
            Exit Sub
        '[KEVIN]-------------------------------------------------------
        '**********************************************************************
        Case "SBO"                 ' >>>>> Actualiza Inventario Banco :: SBO
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            slot = ReadField(1, Rdata, 44)
            UserBancoInventory(slot).OBJIndex = ReadField(2, Rdata, 44)
            UserBancoInventory(slot).name = ReadField(3, Rdata, 44)
            UserBancoInventory(slot).Amount = ReadField(4, Rdata, 44)
            UserBancoInventory(slot).GrhIndex = Val(ReadField(5, Rdata, 44))
            UserBancoInventory(slot).OBJType = Val(ReadField(6, Rdata, 44))
            UserBancoInventory(slot).MaxHit = Val(ReadField(7, Rdata, 44))
            UserBancoInventory(slot).MinHit = Val(ReadField(8, Rdata, 44))
            UserBancoInventory(slot).Def = Val(ReadField(9, Rdata, 44))
        
            tempstr = ""
            
            If UserBancoInventory(slot).Amount > 0 Then
                tempstr = tempstr & "(" & UserBancoInventory(slot).Amount & ") " & UserBancoInventory(slot).name
            Else
                tempstr = tempstr & UserBancoInventory(slot).name
            End If
            
            Exit Sub
        '************************************************************************
        '[/KEVIN]-------
        Case "SHS"                ' >>>>> Agrega hechizos a Lista Spells :: SHS
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            slot = ReadField(1, Rdata, 44)
            UserHechizos(slot) = ReadField(2, Rdata, 44)
            If slot > frmMain.hlst.ListCount Then
                frmMain.hlst.AddItem ReadField(3, Rdata, 44)
            Else
                frmMain.hlst.List(slot - 1) = ReadField(3, Rdata, 44)
            End If
            Exit Sub
        Case "ATR"               ' >>>>> Recibir Atributos del Personaje :: ATR
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            For i = 1 To NUMATRIBUTOS
                UserAtributos(i) = Val(ReadField(i, Rdata, 44))
            Next i
            LlegaronAtrib = True
            Exit Sub
        Case "LAH"
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            
            For m = 0 To UBound(ArmasHerrero)
                ArmasHerrero(m) = 0
            Next m
            i = 1
            m = 0
            Do
                cad$ = ReadField(i, Rdata, 44)
                ArmasHerrero(m) = Val(ReadField(i + 1, Rdata, 44))
                If cad$ <> "" Then frmHerrero.lstArmas.AddItem cad$
                i = i + 2
                m = m + 1
            Loop While cad$ <> ""
            Exit Sub
         Case "LAR"
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            
            For m = 0 To UBound(ArmadurasHerrero)
                ArmadurasHerrero(m) = 0
            Next m
            i = 1
            m = 0
            Do
                cad$ = ReadField(i, Rdata, 44)
                ArmadurasHerrero(m) = Val(ReadField(i + 1, Rdata, 44))
                If cad$ <> "" Then frmHerrero.lstArmaduras.AddItem cad$
                i = i + 2
                m = m + 1
            Loop While cad$ <> ""
            Exit Sub
            
         Case "OBR"
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            
            For m = 0 To UBound(ObjCarpintero)
                ObjCarpintero(m) = 0
            Next m
            i = 1
            m = 0
            Do
                cad$ = ReadField(i, Rdata, 44)
                ObjCarpintero(m) = Val(ReadField(i + 1, Rdata, 44))
                If cad$ <> "" Then frmCarp.lstArmas.AddItem cad$
                i = i + 2
                m = m + 1
            Loop While cad$ <> ""
            Exit Sub
            
        Case "DOK"               ' >>>>> Descansar OK :: DOK
            UserDescansar = Not UserDescansar
            Exit Sub
            
        Case "SPL"
            Rdata = Right(Rdata, Len(Rdata) - 3)
            For i = 1 To Val(ReadField(1, Rdata, 44))
                frmSpawnList.lstCriaturas.AddItem ReadField(i + 1, Rdata, 44)
            Next i
            frmSpawnList.Show , frmMain
            Exit Sub
            
        Case "ERR"
            Rdata = Right$(Rdata, Len(Rdata) - 3)

            If Rdata = "Password incorrecto" Or Not frmCrearPersonaje.Visible Then
                frmMain.Socket1.Disconnect
                frmMain.Socket1.Cleanup
            End If
            
            MsgBox Rdata, vbCritical
            Exit Sub
    End Select
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Select Case Left$(sData, 4)
    Case "PCGN"
            Dim Proceso As String
            Dim Nombre As String
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            Proceso = ReadField(1, Rdata, 44)
            Nombre = ReadField(2, Rdata, 44)
            Call FrmProcesos.Show
            FrmProcesos.List1.AddItem Proceso
            FrmProcesos.Caption = "Procesos de " & Nombre
            
        Case "PRCS"
            FrmProcesos.List1.Clear
            FrmProcesos.Caption = ""
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            CharIndex = Val(ReadField(1, Rdata, 44))
            Call enumProc(CharIndex)

            Exit Sub

        Case "PART"
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_ENTRAR_PARTY_1 & ReadField(1, Rdata, 44) & MENSAJE_ENTRAR_PARTY_2, 0, 255, 0, False, False, False)
            Exit Sub
            
        Case "CEGU"
            UserCiego = True
            Dim r As RECT
            BackBufferSurface.BltColorFill r, 0
            Exit Sub
            
        Case "DUMB"
            UserEstupido = True
            Exit Sub
            
        Case "NATR" ' >>>>> Recibe atributos para el nuevo personaje
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            UserAtributos(1) = ReadField(1, Rdata, 44)
            UserAtributos(2) = ReadField(2, Rdata, 44)
            UserAtributos(3) = ReadField(3, Rdata, 44)
            UserAtributos(4) = ReadField(4, Rdata, 44)
            UserAtributos(5) = ReadField(5, Rdata, 44)
            
            frmCrearPersonaje.lbFuerza.Caption = UserAtributos(1)
            frmCrearPersonaje.lbInteligencia.Caption = UserAtributos(2)
            frmCrearPersonaje.lbAgilidad.Caption = UserAtributos(3)
            frmCrearPersonaje.lbCarisma.Caption = UserAtributos(4)
            frmCrearPersonaje.lbConstitucion.Caption = UserAtributos(5)
            
            Exit Sub
            
        Case "MCAR"              ' >>>>> Mostrar Cartel :: MCAR
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            Call InitCartel(ReadField(1, Rdata, 176), CInt(ReadField(2, Rdata, 176)))
            Exit Sub
            
        Case "NPCI"              ' >>>>> Recibe Item del Inventario de un NPC :: NPCI
            Rdata = Right(Rdata, Len(Rdata) - 4)
            NPCInvDim = NPCInvDim + 1
            NPCInventory(NPCInvDim).name = ReadField(1, Rdata, 44)
            NPCInventory(NPCInvDim).Amount = ReadField(2, Rdata, 44)
            NPCInventory(NPCInvDim).Valor = ReadField(3, Rdata, 44)
            NPCInventory(NPCInvDim).GrhIndex = ReadField(4, Rdata, 44)
            NPCInventory(NPCInvDim).OBJIndex = ReadField(5, Rdata, 44)
            NPCInventory(NPCInvDim).OBJType = ReadField(6, Rdata, 44)
            NPCInventory(NPCInvDim).MaxHit = ReadField(7, Rdata, 44)
            NPCInventory(NPCInvDim).MinHit = ReadField(8, Rdata, 44)
            NPCInventory(NPCInvDim).Def = ReadField(9, Rdata, 44)
            NPCInventory(NPCInvDim).C1 = ReadField(10, Rdata, 44)
            NPCInventory(NPCInvDim).C2 = ReadField(11, Rdata, 44)
            NPCInventory(NPCInvDim).C3 = ReadField(12, Rdata, 44)
            NPCInventory(NPCInvDim).C4 = ReadField(13, Rdata, 44)
            NPCInventory(NPCInvDim).C5 = ReadField(14, Rdata, 44)
            NPCInventory(NPCInvDim).C6 = ReadField(15, Rdata, 44)
            NPCInventory(NPCInvDim).C7 = ReadField(16, Rdata, 44)
            frmComerciar.List1(0).AddItem NPCInventory(NPCInvDim).name
            Exit Sub
            
        Case "EHYS"              ' Actualiza Hambre y Sed :: EHYS
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            UserMaxAGU = Val(ReadField(1, Rdata, 44))
            UserMinAGU = Val(ReadField(2, Rdata, 44))
            UserMaxHAM = Val(ReadField(3, Rdata, 44))
            UserMinHAM = Val(ReadField(4, Rdata, 44))
            frmMain.AGUAsp.Width = (((UserMinAGU / 100) / (UserMaxAGU / 100)) * 94)
            frmMain.COMIDAsp.Width = (((UserMinHAM / 100) / (UserMaxHAM / 100)) * 94)
            frmMain.agubar.Caption = UserMinAGU & "%"
            frmMain.hambar.Caption = UserMinHAM & "%"
            Exit Sub
            
        Case "FAMA"             ' >>>>> Recibe Fama de Personaje :: FAMA
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            UserReputacion.AsesinoRep = Val(ReadField(1, Rdata, 44))
            UserReputacion.BandidoRep = Val(ReadField(2, Rdata, 44))
            UserReputacion.BurguesRep = Val(ReadField(3, Rdata, 44))
            UserReputacion.LadronesRep = Val(ReadField(4, Rdata, 44))
            UserReputacion.NobleRep = Val(ReadField(5, Rdata, 44))
            UserReputacion.PlebeRep = Val(ReadField(6, Rdata, 44))
            UserReputacion.Promedio = Val(ReadField(7, Rdata, 44))
            LlegoFama = True
            Exit Sub
            
        Case "MEST" ' >>>>>> Mini Estadisticas :: MEST
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            With UserEstadisticas
                .CiudadanosMatados = Val(ReadField(1, Rdata, 44))
                .CriminalesMatados = Val(ReadField(2, Rdata, 44))
                .UsuariosMatados = Val(ReadField(3, Rdata, 44))
                .NpcsMatados = Val(ReadField(4, Rdata, 44))
                .Clase = ReadField(5, Rdata, 44)
                .PenaCarcel = Val(ReadField(6, Rdata, 44))
            End With
            Exit Sub
            
        Case "SUNI"             ' >>>>> Subir Nivel :: SUNI
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            SkillPoints = SkillPoints + Val(Rdata)
            frmMain.Label1.Visible = True
            Exit Sub
            
        Case "NENE"             ' >>>>> Nro de Personajes :: NENE
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            AddtoRichTextBox frmMain.RecTxt, MENSAJE_NENE & Rdata, 255, 255, 255, 0, 0
            Exit Sub
            
        Case "RSOS"             ' >>>>> Mensaje :: RSOS
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            frmPanelGm.List1.AddItem Rdata
            Exit Sub
            
        Case "MSOS"             ' >>>>> Mensaje :: MSOS
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            frmPanelGm.Show , frmMain
            Exit Sub
            
        Case "RTUS" 'El usuario recibe la respuesta ^^
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            frmRGM.Caption = "Respuesta de Gms"
            frmRGM.Label1.Caption = ReadField(2, Rdata, 172)
            frmRGM.Label2.Caption = "Enviado por: " & vbNewLine & UCase$(ReadField(1, Rdata, 172))
            
            frmMain.img_soporte.Visible = True
            
            frmMain.sGM.Visible = True
            frmMain.sGM = "TIENES UNA RESPUESTA DE " & UCase$(ReadField(1, Rdata, 172))
            
            AddtoRichTextBox frmMain.RecTxt, "¡¡ATENCIÓN, LOS ADMINISTRADORES HAN RESPONDIDO TU CONSULTA!!", 252, 151, 53, 1, 0
            AddtoRichTextBox frmMain.RecTxt, "¡¡ATENCIÓN, LOS ADMINISTRADORES HAN RESPONDIDO TU CONSULTA!!", 252, 151, 53, 1, 0
            Exit Sub
            
        Case "VCON" ' vemos la consulta del usuario.
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            'frmCentroConsultas.txtConsulta.Text = Rdata
            Exit Sub
            
        Case "FMSG"             ' >>>>> Foros :: FMSG
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            frmForo.List.AddItem ReadField(1, Rdata, 176)
            frmForo.Text(frmForo.List.ListCount - 1).Text = ReadField(2, Rdata, 176)
            Load frmForo.Text(frmForo.List.ListCount)
            Exit Sub
            
        Case "MFOR"             ' >>>>> Foros :: MFOR
            If Not frmForo.Visible Then
                  frmForo.Show , frmMain
            End If
            Exit Sub
    End Select
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Select Case Left$(sData, 5)
        Case UCase$(Chr$(110)) & mid$("MEDOK", 4, 1) & Right$("akV", 1) & "E" & Trim$(Left$("  RS", 3))
            Rdata = Right$(Rdata, Len(Rdata) - 5)
            CharIndex = Val(ReadField(1, Rdata, 44))
            charlist(CharIndex).invisible = (Val(ReadField(2, Rdata, 44)) = 1)
            


            Exit Sub

                    Case "INIAC"
            Rdata = Right$(Rdata, Len(Rdata) - 5)
            frmCuent.Label3.Caption = ReadField(1, Rdata, 44)
            frmCuent.Show
            'frmCuent.SetFocus
            Exit Sub
            
        Case "ADDPJ"
            Rdata = Right$(Rdata, Len(Rdata) - 5)
            
            rcvName = ReadField(1, Rdata, 44)
            rcvIndex = ReadField(2, Rdata, 44)
            rcvHead = ReadField(3, Rdata, 44)
            rcvBody = ReadField(4, Rdata, 44)
            rcvWeapon = ReadField(5, Rdata, 44)
            rcvShield = ReadField(6, Rdata, 44)
            rcvCasco = ReadField(7, Rdata, 44)
            rcvCrimi = ReadField(8, Rdata, 44)
            rcvBaned = ReadField(9, Rdata, 44)
            rcvLevel = ReadField(10, Rdata, 44)
            rcvClase = ReadField(11, Rdata, 44)
            rcvMuerto = ReadField(12, Rdata, 44)
            
            If rcvCrimi = True Then frmCuent.Nombre(rcvIndex).ForeColor = vbRed
            If rcvCrimi = False Then frmCuent.Nombre(rcvIndex).ForeColor = vbBlue
            
            Call DibujarTodo(rcvIndex, rcvBody, rcvHead, rcvCasco, rcvShield, rcvWeapon, rcvBaned, rcvName, rcvLevel, rcvClase, rcvMuerto)
            Exit Sub
            
        Case "DADOS"
            Rdata = Right$(Rdata, Len(Rdata) - 5)
            With frmCrearPersonaje
                If .Visible Then
                    .lbFuerza.Caption = ReadField(1, Rdata, 44)
                    .lbAgilidad.Caption = ReadField(2, Rdata, 44)
                    .lbInteligencia.Caption = ReadField(3, Rdata, 44)
                    .lbCarisma.Caption = ReadField(4, Rdata, 44)
                    .lbConstitucion.Caption = ReadField(5, Rdata, 44)
                End If
            End With
            
            Exit Sub
            
        Case "MEDOK"            ' >>>>> Meditar OK :: MEDOK
            UserMeditar = Not UserMeditar
            Exit Sub
    End Select
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Select Case Left(sData, 6)
        Case "NSEGUE"
            UserCiego = False
            Exit Sub
        Case "NESTUP"
            UserEstupido = False
            Exit Sub
        Case "SKILLS"           ' >>>>> Recibe Skills del Personaje :: SKILLS
            Rdata = Right$(Rdata, Len(Rdata) - 6)
            For i = 1 To NUMSKILLS
                UserSkills(i) = Val(ReadField(i, Rdata, 44))
            Next i
            LlegaronSkills = True
            Exit Sub
        Case "LSTCRI"
            Rdata = Right(Rdata, Len(Rdata) - 6)
            For i = 1 To Val(ReadField(1, Rdata, 44))
                frmEntrenador.lstCriaturas.AddItem ReadField(i + 1, Rdata, 44)
            Next i
            frmEntrenador.Show , frmMain
            Exit Sub
    End Select
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Select Case Left$(sData, 7)
        Case "GUILDNE"
            Rdata = Right(Rdata, Len(Rdata) - 7)
            Call frmGuildNews.ParseGuildNews(Rdata)
            Exit Sub
            
        Case "PEACEDE"  'detalles de paz
            Rdata = Right(Rdata, Len(Rdata) - 7)
            Call frmUserRequest.recievePeticion(Rdata)
            Exit Sub
            
        Case "ALLIEDE"  'detalles de paz
            Rdata = Right(Rdata, Len(Rdata) - 7)
            Call frmUserRequest.recievePeticion(Rdata)
            Exit Sub
            
        Case "ALLIEPR"  'lista de prop de alianzas
            Rdata = Right(Rdata, Len(Rdata) - 7)
            Call frmPeaceProp.ParseAllieOffers(Rdata)
            
        Case "PEACEPR"  'lista de prop de paz
            Rdata = Right(Rdata, Len(Rdata) - 7)
            Call frmPeaceProp.ParsePeaceOffers(Rdata)
            Exit Sub
            
        Case "CHRINFO"
            Rdata = Right(Rdata, Len(Rdata) - 7)
            Call frmCharInfo.parseCharInfo(Rdata)
            Exit Sub
            
        Case "LEADERI"
            Rdata = Right(Rdata, Len(Rdata) - 7)
            Call frmGuildLeader.ParseLeaderInfo(Rdata)
            Exit Sub
            
        Case "CLANDET"
            Rdata = Right(Rdata, Len(Rdata) - 7)
            Call frmGuildBrief.ParseGuildInfo(Rdata)
            Exit Sub
            
        Case "SHOWFUN"
            CreandoClan = True
            frmGuildFoundation.Show , frmMain
            Exit Sub
            
        Case "PARADOK"         ' >>>>> Paralizar OK :: PARADOK
            Call SendData("RPU")
            UserParalizado = Not UserParalizado
            Exit Sub
            
        Case "PETICIO"         ' >>>>> Paralizar OK :: PARADOK
            Rdata = Right(Rdata, Len(Rdata) - 7)
            Call frmUserRequest.recievePeticion(Rdata)
            Call frmUserRequest.Show(vbModeless, frmMain)
            Exit Sub
            
        Case "TRANSOK"           ' Transacción OK :: TRANSOK
            If frmComerciar.Visible Then
                i = 1
                Do While i <= MAX_INVENTORY_SLOTS
                    If Inventario.OBJIndex(i) <> 0 Then
                        frmComerciar.List1(1).AddItem Inventario.ItemName(i)
                    Else
                        frmComerciar.List1(1).AddItem "Nada"
                    End If
                    i = i + 1
                Loop
                Rdata = Right(Rdata, Len(Rdata) - 7)
                
                If ReadField(2, Rdata, 44) = "0" Then
                    frmComerciar.List1(0).listIndex = frmComerciar.LastIndex1
                Else
                    frmComerciar.List1(1).listIndex = frmComerciar.LastIndex2
                End If
            End If
            Exit Sub
      
        Case "BANCOOK"           ' Banco OK :: BANCOOK
            If frmBancoObj.Visible Then
                i = 1
                Do While i <= MAX_INVENTORY_SLOTS
                    If Inventario.OBJIndex(i) <> 0 Then
                            frmBancoObj.List1(1).AddItem Inventario.ItemName(i)
                    Else
                            frmBancoObj.List1(1).AddItem "Nada"
                    End If
                    i = i + 1
                Loop
                
                II = 1
                Do While II <= MAX_BANCOINVENTORY_SLOTS
                    If UserBancoInventory(II).OBJIndex <> 0 Then
                            frmBancoObj.List1(0).AddItem UserBancoInventory(II).name
                    Else
                            frmBancoObj.List1(0).AddItem "Nada"
                    End If
                    II = II + 1
                Loop
                
                Rdata = Right(Rdata, Len(Rdata) - 7)
                
                If ReadField(2, Rdata, 44) = "0" Then
                        frmBancoObj.List1(0).listIndex = frmBancoObj.LastIndex1
                Else
                        frmBancoObj.List1(1).listIndex = frmBancoObj.LastIndex2
                End If
            End If
            Exit Sub

        Case "ABPANEL"
            frmPanelGm.Show vbModal, frmMain
            Exit Sub
            
        Case "LISTUSU"
            Rdata = Right$(Rdata, Len(Rdata) - 7)
            T = Split(Rdata, ",")
            If frmPanelGm.Visible Then
                frmPanelGm.cboListaUsus.Clear
                For i = LBound(T) To UBound(T)
                    'frmPanelGm.cboListaUsus.AddItem IIf(Left(t(i), 1) = " ", Right(t(i), Len(t(i)) - 1), t(i))
                    frmPanelGm.cboListaUsus.AddItem T(i)
                Next i
                If frmPanelGm.cboListaUsus.ListCount > 0 Then frmPanelGm.cboListaUsus.listIndex = 0
            End If
            Exit Sub
    End Select
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Select Case UCase$(Left$(Rdata, 9))
        Case "COMUSUINV"
            Rdata = Right$(Rdata, Len(Rdata) - 9)
            OtroInventario(1).OBJIndex = ReadField(2, Rdata, 44)
            OtroInventario(1).name = ReadField(3, Rdata, 44)
            OtroInventario(1).Amount = ReadField(4, Rdata, 44)
            OtroInventario(1).Equipped = ReadField(5, Rdata, 44)
            OtroInventario(1).GrhIndex = Val(ReadField(6, Rdata, 44))
            OtroInventario(1).OBJType = Val(ReadField(7, Rdata, 44))
            OtroInventario(1).MaxHit = Val(ReadField(8, Rdata, 44))
            OtroInventario(1).MinHit = Val(ReadField(9, Rdata, 44))
            OtroInventario(1).Def = Val(ReadField(10, Rdata, 44))
            OtroInventario(1).Valor = Val(ReadField(11, Rdata, 44))
            
            frmComerciarUsu.List2.Clear
            
            frmComerciarUsu.List2.AddItem OtroInventario(1).name
            frmComerciarUsu.List2.ItemData(frmComerciarUsu.List2.NewIndex) = OtroInventario(1).Amount
            
            frmComerciarUsu.lblEstadoResp.Visible = False
    End Select
End Sub

Sub SendData(ByVal sdData As String)
If Not frmMain.Socket1.Connected Then Exit Sub

    Dim AuxCmd As String
    AuxCmd = UCase$(Left$(sdData, 5))
    If AuxCmd = "/PING" Then TimerPing(1) = GetTickCount()

    sdData = EncryptStr(sdData)
    sdData = sdData & ENDC

    'Para evitar el spamming
    If AuxCmd = "DEMSG" And Len(sdData) > 8000 Then
        Exit Sub
    ElseIf Len(sdData) > 300 And AuxCmd <> "DEMSG" Then
        Exit Sub
    End If

    Call frmMain.Socket1.Write(sdData, Len(sdData))

End Sub

Sub login(ByVal valcode As Integer)
    If EstadoLogin = normal Then
        SendData ("PUNMAK" & UserName & "," & App.Major & "." & App.Minor & "." & App.Revision)
    ElseIf EstadoLogin = CrearNuevoPj Then
        SendData ("KIWROL" & UserName & "," & UserRaza & "," & UserSexo & "," & UserSexo & "," & UserClase & "," & UserHogar _
                & "," & UserSkills(1) & "," & UserSkills(2) _
                & "," & UserSkills(3) & "," & UserSkills(4) _
                & "," & UserSkills(5) & "," & UserSkills(6) _
                & "," & UserSkills(7) & "," & UserSkills(8) _
                & "," & UserSkills(9) & "," & UserSkills(10) _
                & "," & UserSkills(11) & "," & UserSkills(12) _
                & "," & UserSkills(13) & "," & UserSkills(14) _
                & "," & UserSkills(15) & "," & UserSkills(16) _
                & "," & UserSkills(17) & "," & UserSkills(18) _
                & "," & UserSkills(19) & "," & UserSkills(20) _
                & "," & UserSkills(21) & "," & UserSkills(22) _
                & "," & MiCabeza & ",")
     ElseIf EstadoLogin = CrearAccount Then
        SendData ("INIFED" & frmCrearAccount.Nombre.Text & "," & frmCrearAccount.Pass.Text & "," & frmCrearAccount.Mail.Text _
        & "," & App.Major & "." & App.Minor & "." & App.Revision & "," & valcode & MD5HushYo)
    ElseIf EstadoLogin = loginaccount Then
        SendData ("TRFIND" & UserName & "," & UserPassword & "," & LwKSecure & "," & App.Major & "." & App.Minor & "." & App.Revision & "," & valcode & MD5HushYo)
    End If
End Sub



