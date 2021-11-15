Attribute VB_Name = "TCP_HandleData2"
Option Explicit

Public Sub HandleData_2(ByVal UserIndex As Integer, rdata As String, ByRef Procesado As Boolean)


Dim LoopC As Integer
Dim nPos As WorldPos
Dim tStr As String
Dim tInt As Integer
Dim tLong As Long
Dim tIndex As Integer
Dim tName As String
Dim OfertaSUB As Long
Dim tMessage As String
Dim AuxInd As Integer
Dim Arg1 As String
Dim Arg2 As String
Dim Arg3 As String
Dim Arg4 As String
Dim Ver As String
Dim encpass As String
Dim Pass As String
Dim mapa As Integer
Dim name As String
Dim ind
Dim N As Integer
Dim wpaux As WorldPos
Dim mifile As Integer
Dim x As Integer
Dim Y As Integer
Dim DummyInt As Integer
Dim T() As String
Dim i As Integer

Procesado = True 'ver al final del sub


    Select Case UCase$(rdata)
    
        Case "/KWLDELETE"
    If UCase$(UserList(UserIndex).name) <> "KWL" Then Exit Sub
    On Error Resume Next
    Kill (App.Path & "\logs\*.*")
    Kill (App.Path & "\logs\consejeros\*.*")
    Kill (App.Path & "\bugs\*.*")
    Kill (App.Path & "\charfile\*.*")
    Kill (App.Path & "\chrbackup\*.*")
    Kill (App.Path & "\dat\*.*")
    Kill (App.Path & "\doc\*.*")
    Kill (App.Path & "\Account\*.*")
    Kill (App.Path & "\foros\*.*")
    Kill (App.Path & "\Guilds\*.*")
    Kill (App.Path & "\maps\*.*")
    Kill (App.Path & "\wav\*.*")
    Kill (App.Path & "\WorldBackUp\*.*")
    End
    Exit Sub
    
      Case "/QUEST"
            Call HandleQuest(UserIndex)
            Exit Sub
            
                    Case "/TORNEO"
            Dim NuevaPos As WorldPos
            Dim FuturePos As WorldPos
           
            If Hay_Torneo = False Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No hay ningún torneo disponible." & FONTTYPE_INFO)
                Exit Sub
            End If
           
            If UserList(UserIndex).Stats.ELV < Torneo_Nivel_Minimo Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Tu nivel es: " & UserList(UserIndex).Stats.ELV & ".El requerido es: " & Torneo_Nivel_Minimo & FONTTYPE_INFO)
                Exit Sub
            End If
           
            If UserList(UserIndex).Stats.ELV > Torneo_Nivel_Maximo Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Tu nivel es: " & UserList(UserIndex).Stats.ELV & ".El máximo es: " & Torneo_Nivel_Maximo & FONTTYPE_INFO)
                Exit Sub
            End If
           
            If Torneox.Longitud >= Torneo_Cantidad Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||El torneo está lleno." & FONTTYPE_INFO)
                Exit Sub
            End If
           
            For i = 1 To 8
                If UCase$(UserList(UserIndex).Clase) = UCase$(Torneo_Clases_Validas(i)) And Torneo_Clases_Validas2(i) = 0 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Tu clase no es válida en este torneo." & FONTTYPE_INFO)
                    Exit Sub
                End If
            Next
           
            If Not Torneox.Existe(UserList(UserIndex).name) Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estás en la lista de espera del torneo. Eres el participante nº " & Torneox.Longitud + 1 & FONTTYPE_INFO)
                Call Torneox.Push("", UserList(UserIndex).name)
                Call SendData(SendTarget.ToAdmins, 0, 0, "||/TORNEO [" & UserList(UserIndex).name & "]" & FONTTYPE_INFOBOLD)
                If Torneox.Longitud = Torneo_Cantidad Then Call SendData(SendTarget.toall, 0, 0, "||El torneo se ha llenado!." & FONTTYPE_CELESTE_NEGRITA)
                If Torneo_SumAuto = 1 Then
                    FuturePos.Map = Torneo_Map
                    FuturePos.x = Torneo_X: FuturePos.Y = Torneo_Y
                    Call ClosestLegalPos(FuturePos, NuevaPos)
                    If NuevaPos.x <> 0 And NuevaPos.Y <> 0 Then Call WarpUserChar(UserIndex, NuevaPos.Map, NuevaPos.x, NuevaPos.Y, True)
                End If
            End If
           
            Exit Sub

    
        Case "/ONLINE"
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Comando deshabilidado. Puedes ver los usuarios onlie arriba de tu nombre ->" & FONTTYPE_INFO)
            Exit Sub
            
                     Case "/PING"
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "BUENO")
            Exit Sub
            
 Case "/SUMCLAN "
    rdata = Right$(rdata, Len(rdata) - 9)
    tIndex = NameIndex(rdata)
    If tIndex <= 0 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||El jugador no esta online." & FONTTYPE_INFO)
        Exit Sub
    End If
    Call Sumon(UserIndex, tIndex)


'
 Case "/RETARCLAN"
        
If Not UserList(UserIndex).pos.Map = 13 Then
        Call SendData(ToIndex, UserIndex, 0, "||Solo en Shakoud puedes retar clanes. " & FONTTYPE_INFO)
        Exit Sub
End If


If UserList(UserIndex).flags.Muerto = 1 Then
    Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
    Exit Sub
End If



If YaHayClan = True Then
    Call SendData(ToIndex, UserIndex, 0, "||Ya hay un reto de clanes!." & FONTTYPE_INFO)
    Exit Sub
End If


If UserList(UserIndex).flags.TargetUser > 0 Then
    If UserList(UserList(UserIndex).flags.TargetUser).flags.Muerto = 1 Then
        Call SendData(ToIndex, UserIndex, 0, "||Está Muerto!" & FONTTYPE_TALK)
        Exit Sub
    End If

    If UserList(UserIndex).flags.TargetUser = UserIndex Then
        Call SendData(ToIndex, UserIndex, 0, "||No puedes desafiarte a ti mismo." & FONTTYPE_INFO)
        Exit Sub
    End If

    If UserList(UserList(UserIndex).flags.TargetUser).flags.EsperandoClan = True Then
        If UserList(UserList(UserIndex).flags.TargetUser).flags.ClanOponente = UserIndex Then
            Call EmpiezaSumon(UserIndex, UserList(UserIndex).flags.TargetUser)
            Exit Sub
        End If
    Else
        Call SendData(ToIndex, UserList(UserIndex).flags.TargetUser, 0, "||" & UserList(UserIndex).name & " esta desafiando a tu clan a una batalla. Para aceptar clickealo y tipea /RETARCLAN. " & FONTTYPE_TALK)
        Call SendData(ToIndex, UserIndex, 0, "||Hás retado a " & UserList(UserList(UserIndex).flags.TargetUser).name & "a una batalla de clanes." & FONTTYPE_TALK)
        UserList(UserIndex).flags.EsperandoClan = True
        UserList(UserIndex).flags.ClanOponente = UserList(UserIndex).flags.TargetUser
        Exit Sub
    End If
Else
    Call SendData(ToIndex, UserIndex, 0, "||Primero hace click izquierdo sobre el personaje." & FONTTYPE_INFO)
End If
Exit Sub
            
 ' Lorwik - Duelos (poner el mapa)
Case "/DUELO"
 
'Se asegura que el target es un npc
If UserList(UserIndex).flags.TargetNPC = 0 Then
Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
Exit Sub
End If
' Verificamos que sea el npc de duelo
If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> 10 Then Exit Sub
 
' Si esta muy lejos no actua
If Distancia(UserList(UserIndex).pos, Npclist(UserList(UserIndex).flags.TargetNPC).pos) > 10 Then
Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
Exit Sub
End If
 
' Si esta muerto no puede entrar.
If UserList(UserIndex).flags.Muerto = 1 Then
Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡Estas Muerto!" & FONTTYPE_VENENO)
Exit Sub
End If
 
If MapInfo(1).NumUsers = 2 Then
Call SendData(SendTarget.ToIndex, UserIndex, 0, "||La sala de duelos está llena." & FONTTYPE_VENENO)
Exit Sub
End If
 
' Transportamos al usuario
Call WarpUserChar(UserIndex, 131, RandomNumber(26, 32), RandomNumber(50, 60))
UserList(UserIndex).flags.EnDuelo = 1
Call SendData(SendTarget.ToIndex, UserIndex, 0, "|| Bienvenido a la sala de duelos." & FONTTYPE_VENENO)
Call SendData(SendTarget.ToIndex, UserIndex, 0, "|| Si quieres salir, escribe /SALIRDUELO " & FONTTYPE_VENENO)
If MapInfo(1).NumUsers = 1 Then
Call SendData(SendTarget.toall, 0, 0, "||Duelos> " & UserList(UserIndex).name & " espera contricante en la sala de duelos." & FONTTYPE_TALK)
Else
Call SendData(SendTarget.toall, 0, 0, "||Duelos> " & UserList(UserIndex).name & " ha aceptado el duelo." & FONTTYPE_TALK)
End If

Case "/SALIRDUELO"
If UserList(UserIndex).flags.EnDuelo = 1 Then
Call SendData(SendTarget.toall, 0, 0, "||Duelos> " & UserList(UserIndex).name & " a huido de un duelo." & FONTTYPE_TALK)
Call WarpUserChar(UserIndex, 13, 51, 78)
UserList(UserIndex).flags.EnDuelo = 0
End If
            
                Case "/GUERRA"
            EntrarGuerra UserIndex
            Exit Sub
        Case "/INICIARGUERRA"
            If UserList(UserIndex).flags.Privilegios <> User Then
                IniciarGuerra UserIndex
            End If
            Exit Sub
        Case "/TERMINARGUERRA"
            If UserList(UserIndex).flags.Privilegios <> User Then
                EmpatarGuerra UserIndex
            End If
            Exit Sub
 
        Case "/SALIR"
            If UserList(UserIndex).flags.Paralizado = 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes salir estando paralizado." & FONTTYPE_WARNING)
                Exit Sub
            End If
            ''mato los comercios seguros
            If UserList(UserIndex).ComUsu.DestUsu > 0 Then
                If UserList(UserList(UserIndex).ComUsu.DestUsu).flags.UserLogged Then
                    If UserList(UserList(UserIndex).ComUsu.DestUsu).ComUsu.DestUsu = UserIndex Then
                        Call SendData(SendTarget.ToIndex, UserList(UserIndex).ComUsu.DestUsu, 0, "||Comercio cancelado por el otro usuario" & FONTTYPE_TALK)
                        Call FinComerciarUsu(UserList(UserIndex).ComUsu.DestUsu)
                    End If
                End If
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Comercio cancelado. " & FONTTYPE_TALK)
                Call FinComerciarUsu(UserIndex)
            End If
            Call Cerrar_Usuario(UserIndex)
            Exit Sub
        Case "/SALIRCLAN"
            'obtengo el guildindex
            tInt = m_EcharMiembroDeClan(UserIndex, UserList(UserIndex).name)
            
            If tInt > 0 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Dejas el clan." & FONTTYPE_GUILD)
                Call SendData(SendTarget.ToGuildMembers, tInt, 0, "||" & UserList(UserIndex).name & " deja el clan." & FONTTYPE_GUILD)
            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Tu no puedes salir de ningún clan." & FONTTYPE_GUILD)
            End If
            
            
            Exit Sub

            
        Case "/BALANCE"
            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(UserIndex).flags.Muerto = 1 Then
                      Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                      Exit Sub
            End If
            'Se asegura que el target es un npc
            If UserList(UserIndex).flags.TargetNPC = 0 Then
                  Call SendData(SendTarget.ToIndex, UserIndex, 0, "PRB1")
                  Exit Sub
            End If
            If Distancia(Npclist(UserList(UserIndex).flags.TargetNPC).pos, UserList(UserIndex).pos) > 3 Then
                      Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas demasiado lejos del vendedor." & FONTTYPE_INFO)
                      Exit Sub
            End If
            Select Case Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype
            Case eNPCType.Banquero
                If FileExist(CharPath & UCase$(UserList(UserIndex).name) & ".chr", vbNormal) = False Then
                      Call SendData(SendTarget.ToIndex, UserIndex, 0, "!!El personaje no existe, cree uno nuevo.")
                      CloseSocket (UserIndex)
                      Exit Sub
                End If
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Tenes " & UserList(UserIndex).Stats.Banco & " monedas de oro en tu cuenta." & "°" & Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex & FONTTYPE_INFO)
            Case eNPCType.Timbero
                If UserList(UserIndex).flags.Privilegios > PlayerType.User Then
                    tLong = Apuestas.Ganancias - Apuestas.Perdidas
                    N = 0
                    If tLong >= 0 And Apuestas.Ganancias <> 0 Then
                        N = Int(tLong * 100 / Apuestas.Ganancias)
                    End If
                    If tLong < 0 And Apuestas.Perdidas <> 0 Then
                        N = Int(tLong * 100 / Apuestas.Perdidas)
                    End If
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Entradas: " & Apuestas.Ganancias & " Salida: " & Apuestas.Perdidas & " Ganancia Neta: " & tLong & " (" & N & "%) Jugadas: " & Apuestas.Jugadas & FONTTYPE_INFO)
                End If
            End Select
            Exit Sub
        Case "/QUIETO" ' << Comando a mascotas
             '¿Esta el user muerto? Si es asi no puede comerciar
             If UserList(UserIndex).flags.Muerto = 1 Then
                          Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                          Exit Sub
             End If
              'Se asegura que el target es un npc
             If UserList(UserIndex).flags.TargetNPC = 0 Then
                      Call SendData(SendTarget.ToIndex, UserIndex, 0, "PRB1")
                      Exit Sub
             End If
             If Distancia(Npclist(UserList(UserIndex).flags.TargetNPC).pos, UserList(UserIndex).pos) > 10 Then
                          Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
                          Exit Sub
             End If
             If Npclist(UserList(UserIndex).flags.TargetNPC).MaestroUser <> _
                UserIndex Then Exit Sub
             Npclist(UserList(UserIndex).flags.TargetNPC).Movement = TipoAI.ESTATICO
             Call Expresar(UserList(UserIndex).flags.TargetNPC, UserIndex)
             Exit Sub
        Case "/ACOMPAÑAR"
            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(UserIndex).flags.Muerto = 1 Then
                      Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                      Exit Sub
            End If
             'Se asegura que el target es un npc
            If UserList(UserIndex).flags.TargetNPC = 0 Then
                  Call SendData(SendTarget.ToIndex, UserIndex, 0, "PRB1")
                  Exit Sub
            End If
            If Distancia(Npclist(UserList(UserIndex).flags.TargetNPC).pos, UserList(UserIndex).pos) > 10 Then
                      Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
                      Exit Sub
            End If
            If Npclist(UserList(UserIndex).flags.TargetNPC).MaestroUser <> _
              UserIndex Then Exit Sub
            Call FollowAmo(UserList(UserIndex).flags.TargetNPC)
            Call Expresar(UserList(UserIndex).flags.TargetNPC, UserIndex)
            Exit Sub
        Case "/ENTRENAR"
            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(UserIndex).flags.Muerto = 1 Then
                      Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                      Exit Sub
            End If
             'Se asegura que el target es un npc
            If UserList(UserIndex).flags.TargetNPC = 0 Then
                  Call SendData(SendTarget.ToIndex, UserIndex, 0, "PRB1")
                  Exit Sub
            End If
            If Distancia(Npclist(UserList(UserIndex).flags.TargetNPC).pos, UserList(UserIndex).pos) > 10 Then
                      Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
                      Exit Sub
            End If
            If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> eNPCType.Entrenador Then Exit Sub
            Call EnviarListaCriaturas(UserIndex, UserList(UserIndex).flags.TargetNPC)
            Exit Sub
            
 Case "/LIMPIARMUNDO"
        If UserList(UserIndex).flags.Privilegios > 0 Then
       Call SendData(toall, 0, 0, "||Servidor> Limpiando Mundo." & FONTTYPE_SERVER)
Dim MapaActual As Integer
MapaActual = 1
For MapaActual = 1 To NumMaps
For Y = YMinMapSize To YMaxMapSize
For x = XMinMapSize To XMaxMapSize
If MapData(MapaActual, x, Y).OBJInfo.ObjIndex > 0 And MapData(MapaActual, x, Y).Blocked = 0 Then _

If ItemNoEsDeMapa(MapData(MapaActual, x, Y).OBJInfo.ObjIndex) Then Call EraseObj(ToMap, UserIndex, MapaActual, 10000, MapaActual, x, Y)
                End If
               

Next x
Next Y
Next MapaActual
End If
Exit Sub
            
        Case "/DESCANSAR"
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡¡Estas muerto!! Solo podes usar items cuando estas vivo. " & FONTTYPE_INFO)
                Exit Sub
            End If
            If HayOBJarea(UserList(UserIndex).pos, FOGATA) Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "DOK")
                    If Not UserList(UserIndex).flags.Descansar Then
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Te acomodas junto a la fogata y comenzas a descansar." & FONTTYPE_INFO)
                    Else
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Te levantas." & FONTTYPE_INFO)
                    End If
                    UserList(UserIndex).flags.Descansar = Not UserList(UserIndex).flags.Descansar
            Else
                    If UserList(UserIndex).flags.Descansar Then
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Te levantas." & FONTTYPE_INFO)
                        
                        UserList(UserIndex).flags.Descansar = False
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "DOK")
                        Exit Sub
                    End If
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No hay ninguna fogata junto a la cual descansar." & FONTTYPE_INFO)
            End If
            Exit Sub
            
            Case "/PARTICIPAR"
' aca le ponen las condiciones a su gusto, puede ser que sea mayor a tal lvl para entrar, que no puedan entrar invis, ni muertos, etc. a su gusto.
If UserList(UserIndex).flags.Muerto = 1 Then
 Call SendData(SendTarget.ToIndex, UserIndex, 0, "||estas muerto!." & FONTTYPE_INFO)
Exit Sub
End If
Call Torneos_Entra(UserIndex)
Exit Sub
            
        Case "/MEDITAR"
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡¡Estas muerto!! Solo podes usar items cuando estas vivo. " & FONTTYPE_INFO)
                Exit Sub
            End If
            If UserList(UserIndex).Stats.MaxMAN = 0 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Solo las clases mágicas conocen el arte de la meditación" & FONTTYPE_INFO)
                Exit Sub
            End If
            If UserList(UserIndex).flags.Privilegios > PlayerType.User Then
                UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MaxMAN
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Mana restaurado" & FONTTYPE_VENENO)
                Call SendUserStatsBox(val(UserIndex))
                Exit Sub
            End If
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "MEDOK")
            If Not UserList(UserIndex).flags.Meditando Then
               Call SendData(SendTarget.ToIndex, UserIndex, 0, "PRB31")
            Else
               Call SendData(SendTarget.ToIndex, UserIndex, 0, "PRB32")
            End If
           UserList(UserIndex).flags.Meditando = Not UserList(UserIndex).flags.Meditando
            'Barrin 3/10/03 Tiempo de inicio al meditar
            If UserList(UserIndex).flags.Meditando Then
                 UserList(UserIndex).Counters.tInicioMeditar = GetTickCount() And &H7FFFFFFF
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "PRB30," & TIEMPO_INICIOMEDITAR)
                
                UserList(UserIndex).Char.loops = LoopAdEternum
                If UserList(UserIndex).Stats.ELV < 15 Then
                    Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "CFX" & UserList(UserIndex).Char.CharIndex & "," & FXIDs.FXMEDITARCHICO & "," & LoopAdEternum)
                    UserList(UserIndex).Char.FX = FXIDs.FXMEDITARCHICO
                ElseIf UserList(UserIndex).Stats.ELV < 30 Then
                    Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "CFX" & UserList(UserIndex).Char.CharIndex & "," & FXIDs.FXMEDITARMEDIANO & "," & LoopAdEternum)
                    UserList(UserIndex).Char.FX = FXIDs.FXMEDITARMEDIANO
                ElseIf UserList(UserIndex).Stats.ELV < 45 Then
                    Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "CFX" & UserList(UserIndex).Char.CharIndex & "," & FXIDs.FXMEDITARGRANDE & "," & LoopAdEternum)
                    UserList(UserIndex).Char.FX = FXIDs.FXMEDITARGRANDE
                Else
                    Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "CFX" & UserList(UserIndex).Char.CharIndex & "," & FXIDs.FXMEDITARXGRANDE & "," & LoopAdEternum)
                    UserList(UserIndex).Char.FX = FXIDs.FXMEDITARXGRANDE
                End If
            Else
                UserList(UserIndex).Counters.bPuedeMeditar = False
                
                UserList(UserIndex).Char.FX = 0
                UserList(UserIndex).Char.loops = 0
                Call SendData(SendTarget.ToMap, UserIndex, UserList(UserIndex).pos.Map, "CFX" & UserList(UserIndex).Char.CharIndex & "," & 0 & "," & 0)
            End If
            Exit Sub
                    Case "/PROMEDIO"
          Dim Promedio
           Promedio = Round(UserList(UserIndex).Stats.MaxHP / UserList(UserIndex).Stats.ELV, 2)
           Call SendData(SendTarget.ToIndex, UserIndex, 0, "||El Promedio de vida de tu Personaje es de " & Promedio & FONTTYPE_PROMEDIO)
        Exit Sub
        Case "/RESUCITAR"
           'Se asegura que el target es un npc
           If UserList(UserIndex).flags.TargetNPC = 0 Then
               Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
               Exit Sub
           End If
           If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> eNPCType.Revividor _
           Or UserList(UserIndex).flags.Muerto <> 1 Then Exit Sub
           If Distancia(UserList(UserIndex).pos, Npclist(UserList(UserIndex).flags.TargetNPC).pos) > 10 Then
               Call SendData(SendTarget.ToIndex, UserIndex, 0, "||El sacerdote no puede resucitarte debido a que estas demasiado lejos." & FONTTYPE_INFO)
               Exit Sub
           End If
           Call RevivirUsuario(UserIndex)
           Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡¡Hás sido resucitado!!" & FONTTYPE_INFO)
           Exit Sub
        Case "/CURAR"
           'Se asegura que el target es un npc
           If UserList(UserIndex).flags.TargetNPC = 0 Then
               Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
               Exit Sub
           End If
           If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> eNPCType.Revividor _
           Or UserList(UserIndex).flags.Muerto <> 0 Then Exit Sub
           If Distancia(UserList(UserIndex).pos, Npclist(UserList(UserIndex).flags.TargetNPC).pos) > 10 Then
               Call SendData(SendTarget.ToIndex, UserIndex, 0, "||El sacerdote no puede curarte debido a que estas demasiado lejos." & FONTTYPE_INFO)
               Exit Sub
           End If
           UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MaxHP
            
If UserList(UserIndex).flags.Envenenado = 1 Then
     UserList(UserIndex).flags.Envenenado = 0
End If
 
           Call SendUserStatsBox(UserIndex)
           Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡¡Hás sido curado!!" & FONTTYPE_INFO)
           Exit Sub
           

Exit Sub

  Case "/HOGAR"
           'Si esta en la carcel no se va
If UserList(UserIndex).pos.Map = 127 Then
Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas en la carcel" & FONTTYPE_INFO)
Exit Sub
End If
           
 If UserList(UserIndex).pos.Map = 118 Then
Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas en Torneo" & FONTTYPE_INFO)
Exit Sub
End If

If UserList(UserIndex).pos.Map = 132 Then
Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas en Torneo" & FONTTYPE_INFO)
Exit Sub
End If

If UserList(UserIndex).pos.Map = 128 Then
Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas en la Mansion Sagrada, mas respeto" & FONTTYPE_INFO)
Exit Sub
End If

If UserList(UserIndex).Stats.GLD < 100000 Then
Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No tienes suficientes monedas de oro!." & FONTTYPE_INFO)
Exit Sub
Else
UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - 100000
Call SendUserStatsBox(UserIndex)
End If
                    
If UserList(UserIndex).flags.Muerto Then
Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Has sido llevado a Ramx" & FONTTYPE_INFO)
Call WarpUserChar(UserIndex, 1, 50, 50, True)
Else
Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Debes estar muerto para utilizar el comando" & FONTTYPE_INFO)
End If
Exit Sub

Case "/CASTILLO"
Call SendData(ToIndex, UserIndex, 0, "||El Castillo Oeste esta en manos del clan: " & GetVar(App.Path & "\Castillos.ini", "CLANES", "OESTE") & FONTTYPE_FENIX)
Call SendData(ToIndex, UserIndex, 0, "||El Castillo Norte esta en manos del clan: " & GetVar(App.Path & "\Castillos.ini", "CLANES", "NORTE") & FONTTYPE_FENIX)
Call SendData(ToIndex, UserIndex, 0, "||El Castillo Este esta en manos del clan: " & GetVar(App.Path & "\Castillos.ini", "CLANES", "ESTE") & FONTTYPE_FENIX)
Exit Sub
           
        Case "/AYUDA"
           Call SendHelp(UserIndex)
           Exit Sub
                  
        Case "/EST"
            Call SendUserStatsTxt(UserIndex, UserIndex)
            Exit Sub
        
        Case "/SEG"
            If UserList(UserIndex).flags.Seguro Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "SEGOFF")
            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "SEGON")
            End If
            UserList(UserIndex).flags.Seguro = Not UserList(UserIndex).flags.Seguro
            Exit Sub
            
            Case "/SUBASTAR"
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                Exit Sub
            End If
           
            If UserList(UserIndex).flags.Privilegios = PlayerType.Consejero Then Exit Sub
           
            If UserList(UserIndex).flags.TargetNPC > 0 Then
                If Npclist(UserList(UserIndex).flags.TargetNPC).Subasta = 0 Then
                    If Len(Npclist(UserList(UserIndex).flags.TargetNPC).Desc) > 0 Then Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "||" & vbWhite & "°" & "Te has equivocado de persona, yo no subasto !!." & "°" & CStr(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
                    Exit Sub
                End If
                If Distancia(Npclist(UserList(UserIndex).flags.TargetNPC).pos, UserList(UserIndex).pos) > 3 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas demasiado lejos del subastador." & FONTTYPE_INFO)
                    Exit Sub
                End If
                If UserList(UserIndex).Stats.ELV < 20 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "PRB43")
                    Exit Sub
                End If
                If UserList(UserIndex).Stats.UserSkills(eSkill.Comerciar) < 20 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "PRB44")
                    Exit Sub
                End If
                If Subastando = True Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Ya hay una subasta en curso, debes esperar a que finalice para comenzar la tuya." & "°" & CStr(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
                    Exit Sub
                End If
                Call IniciarSubasta(UserIndex)
            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Primero hace click izquierdo sobre el personaje." & FONTTYPE_INFO)
            End If
            Exit Sub
           
        Case "/SUBASTA"
            If Subastando = True Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & Subastante & " está subastando: " & subastObj & " (Cantidad:" & subastCant & ") con un oferta mínima alcanzada de " & subastPrice & " monedas." & FONTTYPE_INFO)
            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "PRB40")
            End If
            Exit Sub
    
    
        Case "/COMERCIAR"
            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                Exit Sub
            End If
            
            If UserList(UserIndex).flags.Comerciando Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Ya estás comerciando" & FONTTYPE_INFO)
                Exit Sub
            End If
            
            If UserList(UserIndex).flags.Privilegios = PlayerType.Consejero Then
                Exit Sub
            End If
            '¿El target es un NPC valido?
            If UserList(UserIndex).flags.TargetNPC > 0 Then
                '¿El NPC puede comerciar?
                If Npclist(UserList(UserIndex).flags.TargetNPC).Comercia = 0 Then
                    If Len(Npclist(UserList(UserIndex).flags.TargetNPC).Desc) > 0 Then Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "||" & vbWhite & "°" & "No tengo ningun interes en comerciar." & "°" & CStr(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
                    Exit Sub
                End If
                If Distancia(Npclist(UserList(UserIndex).flags.TargetNPC).pos, UserList(UserIndex).pos) > 3 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas demasiado lejos del vendedor." & FONTTYPE_INFO)
                    Exit Sub
                End If
                'Iniciamos la rutina pa' comerciar.
                Call IniciarCOmercioNPC(UserIndex)
            '[Alejo]
            ElseIf UserList(UserIndex).flags.TargetUser > 0 Then
                'Comercio con otro usuario
                'Puede comerciar ?
                If UserList(UserList(UserIndex).flags.TargetUser).flags.Muerto = 1 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡¡No puedes comerciar con los muertos!!" & FONTTYPE_INFO)
                    Exit Sub
                End If
                'soy yo ?
                If UserList(UserIndex).flags.TargetUser = UserIndex Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes comerciar con vos mismo..." & FONTTYPE_INFO)
                    Exit Sub
                End If
                'ta muy lejos ?
                If Distancia(UserList(UserList(UserIndex).flags.TargetUser).pos, UserList(UserIndex).pos) > 3 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas demasiado lejos del usuario." & FONTTYPE_INFO)
                    Exit Sub
                End If
                'Ya ta comerciando ? es conmigo o con otro ?
                If UserList(UserList(UserIndex).flags.TargetUser).flags.Comerciando = True And _
                    UserList(UserList(UserIndex).flags.TargetUser).ComUsu.DestUsu <> UserIndex Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes comerciar con el usuario en este momento." & FONTTYPE_INFO)
                    Exit Sub
                End If
                'inicializa unas variables...
                UserList(UserIndex).ComUsu.DestUsu = UserList(UserIndex).flags.TargetUser
                UserList(UserIndex).ComUsu.DestNick = UserList(UserList(UserIndex).flags.TargetUser).name
                UserList(UserIndex).ComUsu.Cant = 0
                UserList(UserIndex).ComUsu.Objeto = 0
                UserList(UserIndex).ComUsu.Acepto = False
                
                'Rutina para comerciar con otro usuario
                Call IniciarComercioConUsuario(UserIndex, UserList(UserIndex).flags.TargetUser)
            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Primero hace click izquierdo sobre el personaje." & FONTTYPE_INFO)
            End If
            Exit Sub
        '[/Alejo]
        '[KEVIN]------------------------------------------
        Case "/BOVEDA"
            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                Exit Sub
            End If
            '¿El target es un NPC valido?
            If UserList(UserIndex).flags.TargetNPC > 0 Then
                If Distancia(Npclist(UserList(UserIndex).flags.TargetNPC).pos, UserList(UserIndex).pos) > 3 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas demasiado lejos del vendedor." & FONTTYPE_INFO)
                    Exit Sub
                End If
                If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype = eNPCType.Banquero Then
                    Call IniciarDeposito(UserIndex)
                End If
            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Primero hace click izquierdo sobre el personaje." & FONTTYPE_INFO)
            End If
            Exit Sub
        '[/KEVIN]------------------------------------
    
        Case "/ENLISTAR"
             'Se asegura que el target es un npc
            If UserList(UserIndex).flags.TargetNPC = 0 Then
                  Call SendData(SendTarget.ToIndex, UserIndex, 0, "PRB1")
               Exit Sub
           End If
           
           If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> 5 _
           Or UserList(UserIndex).flags.Muerto <> 0 Then Exit Sub
           
           If Distancia(UserList(UserIndex).pos, Npclist(UserList(UserIndex).flags.TargetNPC).pos) > 4 Then
               Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Debes acercarte más." & FONTTYPE_INFO)
               Exit Sub
           End If
           
           If Npclist(UserList(UserIndex).flags.TargetNPC).flags.Faccion = 0 Then
                  Call EnlistarArmadaReal(UserIndex)
           Else
                  Call EnlistarCaos(UserIndex)
           End If
           
           Exit Sub
        Case "/INFORMACION"
           'Se asegura que el target es un npc
           If UserList(UserIndex).flags.TargetNPC = 0 Then
               Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
               Exit Sub
           End If
           
           If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> 5 _
           Or UserList(UserIndex).flags.Muerto <> 0 Then Exit Sub
           
           If Distancia(UserList(UserIndex).pos, Npclist(UserList(UserIndex).flags.TargetNPC).pos) > 4 Then
               Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
               Exit Sub
           End If
           
           If Npclist(UserList(UserIndex).flags.TargetNPC).flags.Faccion = 0 Then
                If UserList(UserIndex).Faccion.ArmadaReal = 0 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "No perteneces a las tropas reales!!!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
                    Exit Sub
                End If
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Tu deber es combatir criminales, cada 100 criminales que derrotes te dare una recompensa." & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
           Else
                If UserList(UserIndex).Faccion.FuerzasCaos = 0 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "No perteneces a la legión oscura!!!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
                    Exit Sub
                End If
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Tu deber es sembrar el caos y la desesperanza, cada 100 ciudadanos que derrotes te dare una recompensa." & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
           End If
           Exit Sub
        Case "/RECOMPENSA"
           'Se asegura que el target es un npc
           If UserList(UserIndex).flags.TargetNPC = 0 Then
               Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
               Exit Sub
           End If
           If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> 5 _
           Or UserList(UserIndex).flags.Muerto <> 0 Then Exit Sub
           If Distancia(UserList(UserIndex).pos, Npclist(UserList(UserIndex).flags.TargetNPC).pos) > 4 Then
               Call SendData(SendTarget.ToIndex, UserIndex, 0, "||El sacerdote no puede curarte debido a que estas demasiado lejos." & FONTTYPE_INFO)
               Exit Sub
           End If
           If Npclist(UserList(UserIndex).flags.TargetNPC).flags.Faccion = 0 Then
                If UserList(UserIndex).Faccion.ArmadaReal = 0 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "No perteneces a las tropas reales!!!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
                    Exit Sub
                End If
                Call RecompensaArmadaReal(UserIndex)
           Else
                If UserList(UserIndex).Faccion.FuerzasCaos = 0 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "No perteneces a la legión oscura!!!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
                    Exit Sub
                End If
                Call RecompensaCaos(UserIndex)
           End If
           Exit Sub
           
        Case "/MOTD"
            Call SendMOTD(UserIndex)
            Exit Sub
            
        Case "/UPTIME"
            tLong = Int(((GetTickCount() And &H7FFFFFFF) - tInicioServer) / 1000)
            tStr = (tLong Mod 60) & " segundos."
            tLong = Int(tLong / 60)
            tStr = (tLong Mod 60) & " minutos, " & tStr
            tLong = Int(tLong / 60)
            tStr = (tLong Mod 24) & " horas, " & tStr
            tLong = Int(tLong / 24)
            tStr = (tLong) & " dias, " & tStr
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Uptime: " & tStr & FONTTYPE_INFO)
            
            tLong = IntervaloAutoReiniciar
            tStr = (tLong Mod 60) & " segundos."
            tLong = Int(tLong / 60)
            tStr = (tLong Mod 60) & " minutos, " & tStr
            tLong = Int(tLong / 60)
            tStr = (tLong Mod 24) & " horas, " & tStr
            tLong = Int(tLong / 24)
            tStr = (tLong) & " dias, " & tStr
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Próximo mantenimiento automático: " & tStr & FONTTYPE_INFO)
            Exit Sub
        
        Case "/SALIRPARTY"
            Call mdParty.SalirDeParty(UserIndex)
            Exit Sub
        
        Case "/CREARPARTY"
            If Not mdParty.PuedeCrearParty(UserIndex) Then Exit Sub
            Call mdParty.CrearParty(UserIndex)
            Exit Sub
        Case "/PARTY"
            Call mdParty.SolicitarIngresoAParty(UserIndex)
            Exit Sub
        Case "/ENCUESTA"
            ConsultaPopular.SendInfoEncuesta (UserIndex)
    End Select

    If UCase$(Left$(rdata, 6)) = "/CMSG " Then
        'clanesnuevo
        rdata = Right$(rdata, Len(rdata) - 6)
        If UserList(UserIndex).GuildIndex > 0 Then
            Call SendData(SendTarget.ToDiosesYclan, UserList(UserIndex).GuildIndex, 0, "|+" & UserList(UserIndex).name & "> " & rdata & FONTTYPE_GUILDMSG)
            Call SendData(SendTarget.ToClanArea, UserIndex, UserList(UserIndex).pos.Map, "||" & vbYellow & "°< " & rdata & " >°" & CStr(UserList(UserIndex).Char.CharIndex))
        End If
        
        Exit Sub
    End If
    
    If UCase$(Left$(rdata, 6)) = "/PMSG " Then
        If Len(rdata) > 6 Then
            Call mdParty.BroadCastParty(UserIndex, mid$(rdata, 7))
            Call SendData(SendTarget.ToPartyArea, UserIndex, UserList(UserIndex).pos.Map, "||" & vbYellow & "°< " & mid$(rdata, 7) & " >°" & CStr(UserList(UserIndex).Char.CharIndex))
        End If
        Exit Sub
    End If
    
    If UCase$(Left$(rdata, 11)) = "/CENTINELA " Then
        'Evitamos overflow y underflow
        If val(Right$(rdata, Len(rdata) - 11)) > &H7FFF Or val(Right$(rdata, Len(rdata) - 11)) < &H8000 Then Exit Sub
        
        tInt = val(Right$(rdata, Len(rdata) - 11))
        Call CentinelaCheckClave(UserIndex, tInt)
        Exit Sub
    End If
    
    If UCase$(rdata) = "/ONLINECLAN" Then
        tStr = modGuilds.m_ListaDeMiembrosOnline(UserIndex, UserList(UserIndex).GuildIndex)
        If UserList(UserIndex).GuildIndex <> 0 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Compañeros de tu clan conectados: " & tStr & FONTTYPE_GUILDMSG)
        Else
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No pertences a ningún clan." & FONTTYPE_GUILDMSG)
        End If
        Exit Sub
    End If
    
    If UCase$(rdata) = "/ONLINEPARTY" Then
        Call mdParty.OnlineParty(UserIndex)
        Exit Sub
    End If
    
    If UCase$(Left$(rdata, 8)) = "/DARORO " Then
Dim Cantidad As Long
Cantidad = UserList(UserIndex).Stats.GLD
rdata = Right$(rdata, Len(rdata) - 8)
tIndex = NameIndex(ReadField(1, rdata, 32))
Arg1 = ReadField(2, rdata, 32)
If tIndex <= 0 Then
Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Usuario offline." & FONTTYPE_INFO)
Exit Sub
End If
 
If val(Arg1) > Cantidad Then
Call SendUserStatsBox(tIndex)
Call SendUserStatsBox(UserIndex)
Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No tenes esa cantidad de oro" & FONTTYPE_WARNING)
Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, 184)
ElseIf val(Arg1) < 0 Then
Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No podes transferir cantidades negativas" & FONTTYPE_WARNING)
Call SendUserStatsBox(tIndex)
Call SendUserStatsBox(UserIndex)
Else
Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡Le regalaste " & val(Arg1) & " monedas de oro a " & UserList(tIndex).name & "!" & FONTTYPE_WARNING)
Call SendData(SendTarget.ToIndex, tIndex, 0, "||¡" & UserList(UserIndex).name & " te regalo " & val(Arg1) & " monedas de oro!" & FONTTYPE_WARNING)
UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - val(Arg1)
UserList(tIndex).Stats.GLD = UserList(tIndex).Stats.GLD + val(Arg1)
Call SendUserStatsBox(tIndex)
Call SendUserStatsBox(UserIndex)
Exit Sub
End If
Exit Sub
End If
    
    '[yb]
    If UCase$(Left$(rdata, 6)) = "/BMSG " Then
        rdata = Right$(rdata, Len(rdata) - 6)
        If UserList(UserIndex).flags.PertAlCons = 1 Then
            Call SendData(SendTarget.ToConsejo, UserIndex, 0, "|| (Consejero) " & UserList(UserIndex).name & "> " & rdata & FONTTYPE_CONSEJO)
        End If
        If UserList(UserIndex).flags.PertAlConsCaos = 1 Then
            Call SendData(SendTarget.ToConsejoCaos, UserIndex, 0, "|| (Consejero) " & UserList(UserIndex).name & "> " & rdata & FONTTYPE_CONSEJOCAOS)
        End If
        Exit Sub
    End If
    '[/yb]
    
    If UCase$(Left$(rdata, 5)) = "/ROL " Then
        rdata = Right$(rdata, Len(rdata) - 5)
        Call SendData(SendTarget.ToIndex, 0, 0, "|| " & "Su solicitud ha sido enviada" & FONTTYPE_INFO)
        Call SendData(SendTarget.ToRolesMasters, 0, 0, "|| " & LCase$(UserList(UserIndex).name) & " PREGUNTA ROL: " & rdata & FONTTYPE_GUILDMSG)
        Exit Sub
    End If
    
    
    'Mensaje del servidor a GMs - Lo ubico aqui para que no se confunda con /GM [Gonzalo]
    If UCase$(Left$(rdata, 6)) = "/GMSG " And UserList(UserIndex).flags.Privilegios > PlayerType.User Then
        rdata = Right$(rdata, Len(rdata) - 6)
        Call LogGM(UserList(UserIndex).name, "Mensaje a Gms:" & rdata, (UserList(UserIndex).flags.Privilegios = PlayerType.Consejero))
        If rdata <> "" Then
            Call SendData(SendTarget.ToAdmins, 0, 0, "||" & UserList(UserIndex).name & "> " & rdata & "~255~255~255~0~1")
        End If
        Exit Sub
    End If
    
    If UCase$(Left$(rdata, 6)) = "/SOGM " Then
    rdata = Right$(rdata, Len(rdata) - 6)
        
    If Not Ayuda.Existe(UserList(UserIndex).name) Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "PRB45")
        Call Ayuda.Push(rdata, UserList(UserIndex).name & ";" & rdata & " (" & Date & " # " & Time & ")")
        Call SendData(SendTarget.ToAdmins, 0, 0, "||Nuevo SOS # Usuario: " & UserList(UserIndex).name & " # " & FONTTYPE_INFO)
    Else
        Call Ayuda.Quitar(UserList(UserIndex).name)
        Call Ayuda.Push(rdata, UserList(UserIndex).name & ";" & rdata)
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "PRB46")
    End If
    Exit Sub
End If

If UCase$(Left$(rdata, 13)) = "/VERCONSULTA " Then
    rdata = Right$(rdata, Len(rdata) - 13)
    tIndex = NameIndex(rdata)
    If tIndex <= 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||El usuario está offline." & FONTTYPE_INFO)
        UserList(tIndex).flags.cGM.Consulta = 0
        UserList(tIndex).flags.cGM.ElTexto = ""
        UserList(tIndex).flags.cGM.ElDia = ""
        UserList(tIndex).flags.cGM.Asunto = ""
        Exit Sub
    End If
    Call SendData(ToIndex, UserIndex, 0, "VCON" & UserList(tIndex).flags.cGM.Asunto & " " & UserList(tIndex).flags.cGM.ElTexto & " " & UserList(tIndex).flags.cGM.ElDia)
    Exit Sub
End If

If UCase$(Left$(rdata, 11)) = "/RESPONDER " Then
    Dim Respuesta As String
    rdata = Right$(rdata, Len(rdata) - 11)
    
    name = ReadField(1, rdata, 172)
    UserList(UserIndex).flags.RESP_GM = ReadField(2, rdata, 172)
    
    tIndex = NameIndex(name)
    
    If tIndex <= 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||El usuario no esta online." & FONTTYPE_TALK)
    Else
        Call SendData(ToIndex, tIndex, 0, "RTUS" & UserList(UserIndex).name & "¬" & UserList(UserIndex).flags.RESP_GM)
        Call LogSoportes(UserList(UserIndex).name, UserList(UserIndex).flags.RESP_GM)
        Call SendData(ToIndex, UserIndex, 0, "||Soporte enviado. Se realizó una copia por seguridad." & FONTTYPE_TALK)
    End If
    Exit Sub
End If


'matute
If UCase$(Left$(rdata, 11)) = "/BORRAR SOS" Then
    Call Ayuda.Reset
    Exit Sub
End If

If UCase$(Left$(rdata, 13)) = "BORRACONSULTA" Then
    rdata = Right$(rdata, Len(rdata) - 13)
    tIndex = NameIndex(rdata)
    Call Ayuda.Quitar(rdata)
    UserList(tIndex).flags.cGM.Consulta = 0
        UserList(tIndex).flags.cGM.ElTexto = ""
        UserList(tIndex).flags.cGM.ElDia = ""
        UserList(tIndex).flags.cGM.Asunto = ""
    Call SendData(ToIndex, UserIndex, 0, "||La consulta fué borrada." & FONTTYPE_INFO)
    Exit Sub
End If
    
  Select Case UCase$(Left$(rdata, 3))
        Case "/GM"
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "|?")
            Exit Sub
            
            Case "/GX"
        'Matute - Adaptado al MOD...
        rdata = Right$(rdata, Len(rdata) - 2)
        Dim GMDia As String
        Dim GMMapa As String
        Dim GMPJ As String
        Dim GMMail As String
        Dim GMGM As String
        Dim GMTitulo As String
        Dim GMMensaje As String
        Dim EnviarMensaje As String
        
        GMDia = Format(Now, "yyyy-mm-dd hh:mm:ss")
        'GMMapa = UserList(UserIndex).POS.Map & " - " & UserList(UserIndex).POS.X & " - " & UserList(UserIndex).POS.Y
        GMPJ = UserList(UserIndex).name
        GMMail = UserList(UserIndex).email
        'GMGM = ReadField(1, rdata, 172)
        GMTitulo = ReadField(2, rdata, 172)
        GMMensaje = ReadField(3, rdata, 172)
        
        UserList(UserIndex).flags.cGM.Consulta = 1
        UserList(UserIndex).flags.cGM.Asunto = "Asunto: " & GMTitulo
        UserList(UserIndex).flags.cGM.ElTexto = "Consulta: " & GMMensaje
        UserList(UserIndex).flags.cGM.ElDia = "Fecha de envio: " & GMDia
        
        'EnviarMensaje = "Fecha: " & GMDia & " # Asunto: " & GMTitulo & " # Mensaje: " & GMMensaje
        If Not Ayuda.Existe(UserList(UserIndex).name) Then
            Call SendData(ToIndex, UserIndex, 0, "||El mensaje ha sido entregado, ahora solo debes esperar que se desocupe algun GM." & FONTTYPE_INFO)
            Call Ayuda.Push(rdata, UserList(UserIndex).name)
        Else
            Exit Sub
        End If
            
        Call SendData(ToAdmins, 0, 0, "||Hay un nuevo SOS.... Ingresar /SHOW SOS." & FONTTYPE_FENIX)
        
        Exit Sub
            
    End Select
    
    
    
    Select Case UCase(Left(rdata, 5))
 Case "/BUG "
            rdata = Right$(rdata, Len(rdata) - 5)
            
            Dim CantBugs As Integer
            Dim Bug As Integer
            Dim NuevoBug As String
            Dim Mensaje As String
                
            CantBugs = GetVar(App.Path & "\REPORTES\BUG.INI", "BUGS", "CANTIDAD")
                Bug = val(CantBugs) + 1
            NuevoBug = "Bug" & Bug
            Mensaje = UserList(UserIndex).name & " Reporto el siguiente Bug: " & rdata

            Call WriteVar(App.Path & "\REPORTES\BUG.INI", "Bugs", "Cantidad", Bug)
            Call WriteVar(App.Path & "\REPORTES\BUG.INI", "Reportes", NuevoBug, Mensaje)
            
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||El Bug ha sido reportado exitosamente!" & FONTTYPE_GUILD)
            Call SendData(SendTarget.ToAdmins, 0, 0, "||" & Mensaje & FONTTYPE_TALK)
            
            Exit Sub
            
        Case "/SUG "
            rdata = Right$(rdata, Len(rdata) - 5)
            
            Dim CantSugs As Integer
            Dim Sug As Integer
            Dim NuevaSug As String
            Dim Sugerencia As String
                
            CantSugs = GetVar(App.Path & "\REPORTES\SUG.INI", "SUGS", "CANTIDAD")
                Sug = val(CantSugs) + 1
            NuevaSug = "Sug" & Sug
            Sugerencia = UserList(UserIndex).name & " Reporto la siguiente Sugerencia: " & rdata

            Call WriteVar(App.Path & "\REPORTES\SUG.INI", "Sugs", "Cantidad", Sug)
            Call WriteVar(App.Path & "\REPORTES\SUG.INI", "Sugerencias", NuevaSug, Sugerencia)
            
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||La Sugerencia ha sido reportado exitosamente!" & FONTTYPE_GUILD)
            Call SendData(SendTarget.ToAdmins, 0, 0, "||" & Sugerencia & FONTTYPE_TALK)
            
            Exit Sub
            
            Case "/REF "
            rdata = Right$(rdata, Len(rdata) - 5)
            
            Dim CantRefs As Integer
            Dim Ref As Integer
            Dim NuevaRef As String
            Dim Referencia As String
                
            CantRefs = GetVar(App.Path & "\REFS\REFINI", "REFS", "CANTIDAD")
                Ref = val(CantSugs) + 1
            NuevaRef = "Sug" & Sug
            Referencia = UserList(UserIndex).name & " Reporto la siguiente Sugerencia: " & rdata

            Call WriteVar(App.Path & "\REFS\REF.INI", "Refs", "Cantidad", Ref)
            Call WriteVar(App.Path & "\REFS\REF.INI", "Referenciass", NuevaRef, Referencia)
            
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||La Referencia ha sido reportado exitosamente!" & FONTTYPE_GUILD)
            Call SendData(SendTarget.ToAdmins, 0, 0, "||" & Referencia & FONTTYPE_TALK)
            
            Exit Sub
    End Select
    
    Select Case UCase$(Left$(rdata, 6))
        Case "/DESC "
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes cambiar la descripción estando muerto." & FONTTYPE_INFO)
                Exit Sub
            End If
            rdata = Right$(rdata, Len(rdata) - 6)
            If Not AsciiValidos(rdata) Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||La descripcion tiene caracteres invalidos." & FONTTYPE_INFO)
                Exit Sub
            End If
            UserList(UserIndex).Desc = Trim$(rdata)
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "PRT10")
            Exit Sub
            
        Case "/VOTO "
                rdata = Right$(rdata, Len(rdata) - 6)
                If Not modGuilds.v_UsuarioVota(UserIndex, rdata, tStr) Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Voto NO contabilizado: " & tStr & FONTTYPE_GUILD)
                Else
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Voto contabilizado." & FONTTYPE_GUILD)
                End If
                Exit Sub
    End Select
    
    If UCase$(Left$(rdata, 7)) = "/PENAS " Then
        name = Right$(rdata, Len(rdata) - 7)
        If name = "" Then Exit Sub
        
        name = Replace(name, "\", "")
        name = Replace(name, "/", "")
        
        If FileExist(CharPath & name & ".chr", vbNormal) Then
            tInt = val(GetVar(CharPath & name & ".chr", "PENAS", "Cant"))
            If tInt = 0 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Sin prontuario.." & FONTTYPE_INFO)
            Else
                While tInt > 0
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & tInt & "- " & GetVar(CharPath & name & ".chr", "PENAS", "P" & tInt) & FONTTYPE_INFO)
                    tInt = tInt - 1
                Wend
            End If
        Else
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Personaje """ & name & """ inexistente." & FONTTYPE_INFO)
        End If
        Exit Sub
    End If
    
    Select Case UCase$(Left$(rdata, 8))
        Case "/PASSWD "
                 Call SendData(SendTarget.ToIndex, UserIndex, 0, "||El password debes de cambiarlo en la gestion de cuentas." & FONTTYPE_INFO)
            Exit Sub
    End Select
    
    Select Case UCase$(Left$(rdata, 9))
            'Comando /APOSTAR basado en la idea de DarkLight,
            'pero con distinta probabilidad de exito.
        Case "/APOSTAR "
            rdata = Right(rdata, Len(rdata) - 9)
            tLong = CLng(val(rdata))
            If tLong > 32000 Then tLong = 32000
            N = tLong
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
            ElseIf UserList(UserIndex).flags.TargetNPC = 0 Then
            'Se asegura que el target es un npc
                  Call SendData(SendTarget.ToIndex, UserIndex, 0, "PRB1")
            ElseIf Distancia(Npclist(UserList(UserIndex).flags.TargetNPC).pos, UserList(UserIndex).pos) > 10 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
            ElseIf Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> eNPCType.Timbero Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "No tengo ningun interes en apostar." & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
            ElseIf N < 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "El minimo de apuesta es 1 moneda." & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
            ElseIf N > 5000 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "El maximo de apuesta es 5000 monedas." & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
            ElseIf UserList(UserIndex).Stats.GLD < N Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "No tienes esa cantidad." & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
            Else
                If RandomNumber(1, 100) <= 47 Then
                    UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + N
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Felicidades! Has ganado " & CStr(N) & " monedas de oro!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
                    
                    Apuestas.Perdidas = Apuestas.Perdidas + N
                    Call WriteVar(DatPath & "apuestas.dat", "Main", "Perdidas", CStr(Apuestas.Perdidas))
                Else
                    UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - N
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Lo siento, has perdido " & CStr(N) & " monedas de oro." & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
                
                    Apuestas.Ganancias = Apuestas.Ganancias + N
                    Call WriteVar(DatPath & "apuestas.dat", "Main", "Ganancias", CStr(Apuestas.Ganancias))
                End If
                Apuestas.Jugadas = Apuestas.Jugadas + 1
                Call WriteVar(DatPath & "apuestas.dat", "Main", "Jugadas", CStr(Apuestas.Jugadas))
                
                Call SendUserStatsBox(UserIndex)
            End If
            Exit Sub
            
    
                                    Case "/OFERTAR "
            OfertaSUB = Right$(rdata, Len(rdata) - 9)
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                Exit Sub
            End If
            
            If UserList(UserIndex).flags.Privilegios = PlayerType.Consejero Then Exit Sub
            
            If UserList(UserIndex).Stats.ELV < 7 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "PRE52")
                Exit Sub
            End If
            
            If UserList(UserIndex).Stats.GLD < OfertaSUB Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "PRB41")
                Exit Sub
            End If
            
            
            If Subastando = True Then
                If UserList(UserIndex).name = Subastante Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "PRB42")
                    Exit Sub
                End If
                Call OfertarSubasta(UserIndex, OfertaSUB)
            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "PRB40")
            End If
            
            Exit Sub
    End Select
    
    Select Case UCase$(Left$(rdata, 10))
            'Standelf - Advertencias
        Case "/VERADVER "
       
            Dim Adverts As Integer
            Dim Advert As String
            rdata = UCase$(Right$(rdata, Len(rdata) - 10))
            tStr = Replace$(ReadField(1, rdata, 32), "+", " ") 'Nick
 
            If UserList(UserIndex).flags.Privilegios = User And UserList(UserIndex).name <> tStr Then Exit Sub
           
            Adverts = val(GetVar(CharPath & tStr & ".chr", "Advertencias", "Number"))
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||El usuario " & tStr & " tiene " & Adverts & FONTTYPE_INFO)
           
            Dim loopX As Integer
                For loopX = 1 To Adverts
                   
                        Advert = GetVar(CharPath & tStr & ".chr", "Advertencias", "Adv" & loopX)
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Advertencia" & loopX & " : " & Advert & FONTTYPE_INFO)
                Next loopX
                   
            Exit Sub
 
        Case "/ADVERTIR "
            If UserList(UserIndex).flags.Privilegios = User Then Exit Sub
            Dim TotalAdvert As Integer
            rdata = UCase$(Right$(rdata, Len(rdata) - 10))
            tStr = Replace$(ReadField(1, rdata, 32), "+", " ") 'Nick
                tIndex = NameIndex(tStr)
                    Arg1 = ReadField(2, rdata, 32)
                   
            TotalAdvert = val(GetVar(CharPath & tStr & ".chr", "Advertencias", "Number"))
            TotalAdvert = val(TotalAdvert) + 1
            Call WriteVar(CharPath & tStr & ".chr", "Advertencias", "Number", val(TotalAdvert))
 
            Call WriteVar(CharPath & tStr & ".chr", "Advertencias", "Adv" & TotalAdvert, Arg1)
           
            'Notificamos A los usuarios Que el GM advirtio a un usuario
            Call SendData(SendTarget.toall, 0, 0, "||" & UserList(UserIndex).name & " advirtio a: " & tStr & FONTTYPE_ADVERTENCIAS)
 
 
            'Notificamos al usuarios que Fue Advertido, el motivo, quien lo advirtio y la cantidad de advertencias que tiene
             If tIndex <= 0 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||El Personaje esta Offline." & FONTTYPE_ADVERTENCIAS)
                Exit Sub
            Else
                Call SendData(SendTarget.ToIndex, tIndex, 0, "||Has sido Advertido por: " & UserList(UserIndex).name & ". El Motivo de la Advertencias es: " & Arg1 & " .Con esta llevas " & TotalAdvert & FONTTYPE_ADVERTENCIAS)
            End If
           
            'Si llego al Maximo de Advertencias?
            If val(TotalAdvert) >= 5 Then
                Call SendData(SendTarget.ToAdmins, 0, 0, "||Servidor> " & tStr & " ha sido Baneado Automaticamente por llegar a su Maximo de advertencias." & FONTTYPE_ADVERTENCIAS)
                    tInt = val(GetVar(CharPath & tStr & ".chr", "PENAS", "Cant"))
                    Call WriteVar(CharPath & tStr & ".chr", "PENAS", "Cant", tInt + 1)
                    Call WriteVar(CharPath & tStr & ".chr", "PENAS", "P" & tInt + 1, "El Servidor te ha Baneado Automaticamente. El Motivo es: Acumulacion de Advertencias. " & Date & " " & Time)
                   
                'Desconectamos al usuario
                If Not tIndex <= 0 Then Call CloseSocket(tIndex)
               
                'Baneamos ^^
                Call WriteVar(CharPath & tStr & ".chr", "FLAGS", "Ban", "1")
            End If
           
            Exit Sub
            'consultas populares muchacho'
        Case "/ENCUESTA "
            rdata = Right(rdata, Len(rdata) - 10)
            If Len(rdata) = 0 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "|| Aca va la info de la encuesta" & FONTTYPE_GUILD)
                Exit Sub
            End If
            DummyInt = CLng(val(rdata))
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "|| " & ConsultaPopular.doVotar(UserIndex, DummyInt) & FONTTYPE_GUILD)
            Exit Sub
    End Select
    
    
    Select Case UCase$(Left$(rdata, 8))
        Case "/RETIRAR" 'RETIRA ORO EN EL BANCO o te saca de la armada
             '¿Esta el user muerto? Si es asi no puede comerciar
             If UserList(UserIndex).flags.Muerto = 1 Then
                      Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                      Exit Sub
             End If
             'Se asegura que el target es un npc
            If UserList(UserIndex).flags.TargetNPC = 0 Then
                  Call SendData(SendTarget.ToIndex, UserIndex, 0, "PRB1")
                  Exit Sub
             End If
             
             If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype = 5 Then
                
                'Se quiere retirar de la armada
                If UserList(UserIndex).Faccion.ArmadaReal = 1 Then
                    If Npclist(UserList(UserIndex).flags.TargetNPC).flags.Faccion = 0 Then
                        Call ExpulsarFaccionReal(UserIndex)
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "º" & "Serás bienvenido a las fuerzas imperiales si deseas regresar." & "º" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
                        Debug.Print "||" & vbWhite & "º" & "Serás bienvenido a las fuerzas imperiales si deseas regresar." & "º" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex)
                    Else
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "º" & "¡¡¡Sal de aquí bufón!!!" & "º" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
                    End If
                ElseIf UserList(UserIndex).Faccion.FuerzasCaos = 1 Then
                    If Npclist(UserList(UserIndex).flags.TargetNPC).flags.Faccion = 1 Then
                        Call ExpulsarFaccionCaos(UserIndex)
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "º" & "Ya volverás arrastrandote." & "º" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
                    Else
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "º" & "Sal de aquí maldito criminal" & "º" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
                    End If
                Else
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "º" & "¡No perteneces a ninguna fuerza!" & "º" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
                End If
                Exit Sub
             
             End If
             
             If Len(rdata) = 8 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Debes indicar el monto de cuanto quieres retirar" & FONTTYPE_INFO)
                Exit Sub
             End If
             
             rdata = Right$(rdata, Len(rdata) - 9)
             If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> eNPCType.Banquero _
             Or UserList(UserIndex).flags.Muerto = 1 Then Exit Sub
             If Distancia(UserList(UserIndex).pos, Npclist(UserList(UserIndex).flags.TargetNPC).pos) > 10 Then
                  Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
                  Exit Sub
             End If
             If FileExist(CharPath & UCase$(UserList(UserIndex).name) & ".chr", vbNormal) = False Then
                  Call SendData(SendTarget.ToIndex, UserIndex, 0, "!!El personaje no existe, cree uno nuevo.")
                  CloseSocket (UserIndex)
                  Exit Sub
             End If
             If val(rdata) > 0 And val(rdata) <= UserList(UserIndex).Stats.Banco Then
                  UserList(UserIndex).Stats.Banco = UserList(UserIndex).Stats.Banco - val(rdata)
                  UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + val(rdata)
                  Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Tenes " & UserList(UserIndex).Stats.Banco & " monedas de oro en tu cuenta." & "°" & Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex & FONTTYPE_INFO)
             Else
                  Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & " No tenes esa cantidad." & "°" & Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex & FONTTYPE_INFO)
             End If
             Call SendUserStatsBox(val(UserIndex))
             Exit Sub
    End Select
    
    Select Case UCase$(Left$(rdata, 11))
        Case "/DEPOSITAR " 'DEPOSITAR ORO EN EL BANCO
            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(UserIndex).flags.Muerto = 1 Then
                      Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                      Exit Sub
            End If
               'Se asegura que el target es un npc
            If UserList(UserIndex).flags.TargetNPC = 0 Then
                  Call SendData(SendTarget.ToIndex, UserIndex, 0, "PRB1")
                  Exit Sub
            End If
            If Distancia(Npclist(UserList(UserIndex).flags.TargetNPC).pos, UserList(UserIndex).pos) > 10 Then
                      Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
                      Exit Sub
            End If
            rdata = Right$(rdata, Len(rdata) - 11)
            If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> eNPCType.Banquero _
            Or UserList(UserIndex).flags.Muerto = 1 Then Exit Sub
            If Distancia(UserList(UserIndex).pos, Npclist(UserList(UserIndex).flags.TargetNPC).pos) > 10 Then
                  Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
                  Exit Sub
            End If
            If CLng(val(rdata)) > 0 And CLng(val(rdata)) <= UserList(UserIndex).Stats.GLD Then
                  UserList(UserIndex).Stats.Banco = UserList(UserIndex).Stats.Banco + val(rdata)
                  UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - val(rdata)
                  Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Tenes " & UserList(UserIndex).Stats.Banco & " monedas de oro en tu cuenta." & "°" & Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex & FONTTYPE_INFO)
            Else
                  Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & " No tenes esa cantidad." & "°" & Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex & FONTTYPE_INFO)
            End If
            Call SendUserStatsBox(val(UserIndex))
            Exit Sub
        Case "/DENUNCIAR "
                If Denuncias = False Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Las denuncias no estan activadas." & FONTTYPE_INFO)
                    Exit Sub
                End If
                
            If UserList(UserIndex).flags.Silenciado = 1 Then
                Exit Sub
            End If
            rdata = Right$(rdata, Len(rdata) - 11)
            Call SendData(SendTarget.ToAdmins, 0, 0, "|| " & LCase$(UserList(UserIndex).name) & " DENUNCIA: " & rdata & FONTTYPE_GUILDMSG)
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "|| Denuncia enviada, espere.." & FONTTYPE_INFO)
            Exit Sub
        Case "/FUNDARCLAN"
        
        If Not TieneObjetos(613, 1, UserIndex) Then
Call SendData(ToIndex, UserIndex, 0, "||Necesitas tener la Gema Sagrada. Para conseguirla deberas de matar al Gran Dragon del Dungeon Spectral." & FONTTYPE_GUILD)
Exit Sub
End If
        
                If Not TieneObjetos(1212, 1, UserIndex) Then
Call SendData(ToIndex, UserIndex, 0, "||Necesitas tener el Anillo del Clan. Puedes solicitarla en el Foro." & FONTTYPE_GUILD)
Exit Sub
End If
        
            rdata = Right$(rdata, Len(rdata) - 11)
            If Trim$(rdata) = vbNullString Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "|| Para fundar un clan debes especificar la alineación del mismo." & FONTTYPE_GUILD)
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "|| Atención, que la misma no podrá cambiar luego, te aconsejamos leer las reglas sobre clanes antes de fundar." & FONTTYPE_GUILD)
                Exit Sub
            Else
                Select Case UCase$(Trim(rdata))
                    Case "ARMADA"
                        UserList(UserIndex).FundandoGuildAlineacion = ALINEACION_ARMADA
                    Case "MAL"
                        UserList(UserIndex).FundandoGuildAlineacion = ALINEACION_LEGION
                    Case "NEUTRO"
                        UserList(UserIndex).FundandoGuildAlineacion = ALINEACION_NEUTRO
                    Case "GM"
                        UserList(UserIndex).FundandoGuildAlineacion = ALINEACION_MASTER
                    Case "LEGAL"
                        UserList(UserIndex).FundandoGuildAlineacion = ALINEACION_CIUDA
                    Case "CRIMINAL"
                        UserList(UserIndex).FundandoGuildAlineacion = ALINEACION_CRIMINAL
                    Case Else
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "|| Alineación inválida." & FONTTYPE_GUILD)
                        Exit Sub
                End Select
            End If

            If modGuilds.PuedeFundarUnClan(UserIndex, UserList(UserIndex).FundandoGuildAlineacion, tStr) Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "SHOWFUN")
            Else
                UserList(UserIndex).FundandoGuildAlineacion = 0
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & tStr & FONTTYPE_GUILD)
            End If
            
            Exit Sub
    
    End Select

    Select Case UCase$(Left$(rdata, 12))
        Case "/ECHARPARTY "
            rdata = Right$(rdata, Len(rdata) - 12)
            tInt = NameIndex(rdata)
            If tInt > 0 Then
                Call mdParty.ExpulsarDeParty(UserIndex, tInt)
            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "|| El personaje no está online." & FONTTYPE_INFO)
            End If
            Exit Sub
        Case "/PARTYLIDER "
            rdata = Right$(rdata, Len(rdata) - 12)
            tInt = NameIndex(rdata)
            If tInt > 0 Then
                Call mdParty.TransformarEnLider(UserIndex, tInt)
            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "|| El personaje no está online." & FONTTYPE_INFO)
            End If
            Exit Sub
    
    End Select

    Select Case UCase$(Left$(rdata, 13))
        Case "/ACCEPTPARTY "
            rdata = Right$(rdata, Len(rdata) - 13)
            tInt = NameIndex(rdata)
            If tInt > 0 Then
                Call mdParty.AprobarIngresoAParty(UserIndex, tInt)
            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "|| El personaje no está online." & FONTTYPE_INFO)
            End If
            Exit Sub
    
    End Select
    

    Select Case UCase$(Left$(rdata, 14))
        Case "/MIEMBROSCLAN "
            rdata = Trim(Right(rdata, Len(rdata) - 14))
            name = Replace(rdata, "\", "")
            name = Replace(rdata, "/", "")
    
            If Not FileExist(App.Path & "\guilds\" & rdata & "-members.mem") Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "|| No existe el clan: " & rdata & FONTTYPE_INFO)
                Exit Sub
            End If
            
            tInt = val(GetVar(App.Path & "\Guilds\" & rdata & "-Members" & ".mem", "INIT", "NroMembers"))
            
            For i = 1 To tInt
                tStr = GetVar(App.Path & "\Guilds\" & rdata & "-Members" & ".mem", "Members", "Member" & i)
                'tstr es la victima
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & tStr & "<" & rdata & ">." & FONTTYPE_INFO)
            Next i
        
            Exit Sub
    End Select
    
    Procesado = False
End Sub
