Attribute VB_Name = "TCP_HandleData1"
Option Explicit

Public Sub HandleData_1(ByVal userindex As Integer, rdata As String, ByRef Procesado As Boolean)


Dim LoopC As Integer
Dim nPos As WorldPos
Dim tStr As String
Dim tInt As Integer
Dim tLong As Long
Dim tIndex As Integer
Dim tName As String
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

       Select Case UCase$(Left$(rdata, 1))
        Case ";" 'Hablar
            rdata = Right$(rdata, Len(rdata) - 1)
            If InStr(rdata, "°") Then
                Exit Sub
            End If
        
            '[Consejeros]
            If UserList(userindex).flags.Privilegios = PlayerType.Consejero Then
                Call LogGM(UserList(userindex).name, "Dijo: " & rdata, True)
            End If
            
            ind = UserList(userindex).Char.CharIndex
            
            If rdata = "xD" Then
                Call SendData(ToPCArea, userindex, UserList(userindex).pos.Map, "CFX" & UserList(userindex).Char.CharIndex & "," & FXXD & "," & 20)
                Exit Sub
            ElseIf rdata = ":)" Then
                Call SendData(ToPCArea, userindex, UserList(userindex).pos.Map, "CFX" & UserList(userindex).Char.CharIndex & "," & FXFE & "," & 20)
                Exit Sub
            ElseIf rdata = ":(" Then
                Call SendData(ToPCArea, userindex, UserList(userindex).pos.Map, "CFX" & UserList(userindex).Char.CharIndex & "," & FXTR & "," & 20)
                Exit Sub
            ElseIf rdata = ":@" Then
                Call SendData(ToPCArea, userindex, UserList(userindex).pos.Map, "CFX" & UserList(userindex).Char.CharIndex & "," & FXHOT & "," & 20)
                Exit Sub
            ElseIf rdata = ":'(" Then
                Call SendData(ToPCArea, userindex, UserList(userindex).pos.Map, "CFX" & UserList(userindex).Char.CharIndex & "," & FXBUA & "," & 20)
                Exit Sub
            ElseIf rdata = "(H)" Then
                Call SendData(ToPCArea, userindex, UserList(userindex).pos.Map, "CFX" & UserList(userindex).Char.CharIndex & "," & FXEA & "," & 20)
                Exit Sub
            ElseIf rdata = ":S" Then
                Call SendData(ToPCArea, userindex, UserList(userindex).pos.Map, "CFX" & UserList(userindex).Char.CharIndex & "," & FXCON & "," & 20)
                Exit Sub
            ElseIf rdata = ":|" Then
                Call SendData(ToPCArea, userindex, UserList(userindex).pos.Map, "CFX" & UserList(userindex).Char.CharIndex & "," & FXWTF & "," & 20)
            Exit Sub
        End If
            'piedra libre para todos los compas!
            If UserList(userindex).flags.Oculto > 0 Then
                UserList(userindex).flags.Oculto = 0
                If UserList(userindex).flags.Invisible = 0 Then
                    Call SendData(SendTarget.ToMap, 0, UserList(userindex).pos.Map, "NOVER" & UserList(userindex).Char.CharIndex & ",0")
                    Call SendData(SendTarget.ToIndex, userindex, 0, "PRB39")
                End If
            End If
            
                                For LoopC = 1 To Len(rdata)
If LCase$(mid$(rdata, LoopC, 9)) = "~" Then
Call SendData(SendTarget.ToIndex, userindex, 0, "||El Texto tiene caracteres invalidos." & FONTTYPE_INFO)
Exit Sub
End If
Next LoopC
            
                                
            'Admin (Reservado y solo exclusivo para el pro de Lorwik xD)
              If UserList(userindex).flags.Privilegios = PlayerType.Admin Then
               Call SendData(ToPCArea, userindex, UserList(userindex).pos.Map, "||" & vbCyan & "°" & rdata & "°" & str(ind))
                                Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "||" & UserList(userindex).name & "> " & rdata & FONTTYPE_LORWIAD)
                                
                    'Game Masters
                ElseIf UserList(userindex).flags.Privilegios > 0 Then
                    Call SendData(ToPCArea, userindex, UserList(userindex).pos.Map, "||" & vbGreen & "°" & rdata & "°" & str(ind))
                                Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "||" & UserList(userindex).name & "> " & rdata & FONTTYPE_LORWIKG)
            'Muerto
                ElseIf UserList(userindex).flags.Muerto = 1 Then
                    Call SendData(ToDeadArea, userindex, UserList(userindex).pos.Map, "||" & vbCyan & "°" & rdata & "°" & str(ind))
                                Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "||" & UserList(userindex).name & "> " & rdata & FONTTYPE_LORWIKM)
            'Usuario Comun
                Else
                    Call SendData(ToPCArea, userindex, UserList(userindex).pos.Map, "||" & vbWhite & "°" & rdata & "°" & str(ind))
                    Call SendData(ToPCArea, userindex, UserList(userindex).pos.Map, "||" & UserList(userindex).name & "> " & rdata & FONTTYPE_LORWIK)

            End If


        Exit Sub
            
            Case ":"
            If UserList(userindex).flags.Muerto = 1 Then
                    Call SendData(SendTarget.ToIndex, userindex, 0, "||¡¡Estas muerto!! Los muertos no pueden comunicarse con el mundo de los vivos. " & FONTTYPE_INFO)
                    Exit Sub
            End If
            rdata = Right$(rdata, Len(rdata) - 1)
           
            
For LoopC = 1 To Len(rdata)
If LCase$(mid$(rdata, LoopC, 9)) = "~" Then
Call SendData(SendTarget.ToIndex, userindex, 0, "||El Texto tiene caracteres invalidos." & FONTTYPE_INFO)
Exit Sub
End If
Next LoopC

  rdata = " " & rdata & " "
             rdata = Replace(rdata, " ~", " ")
             rdata = mid(rdata, 2, Len(rdata) - 2)

           
            If glob = False Then
            Call SendData(SendTarget.ToIndex, userindex, 0, "PRB50")
            Exit Sub
            End If
            

            Call SendData(SendTarget.toall, 0, 0, "||" & UserList(userindex).name & "> " & rdata & FONTTYPE_PARTY)
            
        Case "-" 'Gritar
            If UserList(userindex).flags.Muerto = 1 Then
                    Call SendData(SendTarget.ToIndex, userindex, 0, "||¡¡Estas muerto!! Los muertos no pueden comunicarse con el mundo de los vivos. " & FONTTYPE_INFO)
                    Exit Sub
            End If
            rdata = Right$(rdata, Len(rdata) - 1)
            If InStr(rdata, "°") Then
                Exit Sub
            End If
            '[Consejeros]
            If UserList(userindex).flags.Privilegios = PlayerType.Consejero Then
                Call LogGM(UserList(userindex).name, "Grito: " & rdata, True)
            End If
    
            'piedra libre para todos los compas!
            If UserList(userindex).flags.Oculto > 0 Then
                UserList(userindex).flags.Oculto = 0
                If UserList(userindex).flags.Invisible = 0 Then
                    Call SendData(SendTarget.ToMap, 0, UserList(userindex).pos.Map, "NOVER" & UserList(userindex).Char.CharIndex & ",0")
                   Call SendData(SendTarget.ToIndex, userindex, 0, "PRB39")
                End If
            End If
    
    
            ind = UserList(userindex).Char.CharIndex
            Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "||" & vbRed & "°" & rdata & "°" & str(ind))
            Exit Sub
        Case "\" 'Susurrar al oido
            If UserList(userindex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, userindex, 0, "||¡¡Estas muerto!! Los muertos no pueden comunicarse con el mundo de los vivos. " & FONTTYPE_INFO)
                Exit Sub
            End If
            rdata = Right$(rdata, Len(rdata) - 1)
            tName = ReadField(1, rdata, 32)
            
            'A los dioses y admins no vale susurrarles si no sos uno vos mismo (así no pueden ver si están conectados o no)
            If (EsDios(tName) Or EsAdmin(tName)) And UserList(userindex).flags.Privilegios < PlayerType.Dios Then
                Call SendData(SendTarget.ToIndex, userindex, 0, "||No puedes susurrarle a los Dioses y Admins." & FONTTYPE_INFO)
                Exit Sub
            End If
            
            'A los Consejeros y SemiDioses no vale susurrarles si sos un PJ común.
            If UserList(userindex).flags.Privilegios = PlayerType.User And (EsSemiDios(tName) Or EsConsejero(tName)) Then
                Call SendData(SendTarget.ToIndex, userindex, 0, "||No puedes susurrarle a los GMs" & FONTTYPE_INFO)
                Exit Sub
            End If
            
            tIndex = NameIndex(tName)
            If tIndex <> 0 Then
                If Len(rdata) <> Len(tName) Then
                    tMessage = Right$(rdata, Len(rdata) - (1 + Len(tName)))
                Else
                    tMessage = " "
                End If
                If Not EstaPCarea(userindex, tIndex) Then
                    Call SendData(SendTarget.ToIndex, userindex, 0, "||Estas muy lejos del usuario." & FONTTYPE_INFO)
                    Exit Sub
                End If
                ind = UserList(userindex).Char.CharIndex
                If InStr(tMessage, "°") Then
                    Exit Sub
                End If
                
                '[Consejeros]
                If UserList(userindex).flags.Privilegios = PlayerType.Consejero Then
                    Call LogGM(UserList(userindex).name, "Le dijo a '" & UserList(tIndex).name & "' " & tMessage, True)
                End If
    
                Call SendData(SendTarget.ToIndex, userindex, UserList(userindex).pos.Map, "||" & vbBlue & "°" & tMessage & "°" & str(ind))
                Call SendData(SendTarget.ToIndex, tIndex, UserList(userindex).pos.Map, "||" & vbBlue & "°" & tMessage & "°" & str(ind))
                '[CDT 17-02-2004]
                If UserList(userindex).flags.Privilegios < PlayerType.SemiDios Then
                    Call SendData(SendTarget.ToAdminsAreaButConsejeros, userindex, UserList(userindex).pos.Map, "||" & vbYellow & "°" & "a " & UserList(tIndex).name & "> " & tMessage & "°" & str(ind))
                End If
                '[/CDT]
                Exit Sub
            End If
            Call SendData(SendTarget.ToIndex, userindex, 0, "||Usuario inexistente. " & FONTTYPE_INFO)
            Exit Sub
            
        Case "M" 'Moverse
        'Lorwik> Gracias a EAO pude mejorarlo xD
           Dim dummy As Long
            Dim TempTick As Long
            Dim TiempoDeWalk As Long
            
            If UserList(userindex).flags.Equitando = 1 Then
                TiempoDeWalk = 38
            Else
                TiempoDeWalk = 34
            End If
            
            If UserList(userindex).flags.Muerto = 1 Then
                TiempoDeWalk = 38
            Else
                TiempoDeWalk = 34
            End If
            
            If UserList(userindex).flags.TimesWalk >= TiempoDeWalk Then
                TempTick = GetTickCount And &H7FFFFFFF
                dummy = (TempTick - UserList(userindex).flags.StartWalk)
                If dummy < 6050 Then
                    If TempTick - UserList(userindex).flags.CountSH > 90000 Then
                        UserList(userindex).flags.CountSH = 0
                    End If
                    If Not UserList(userindex).flags.CountSH = 0 Then
                        dummy = 126000 \ dummy
                        Call LogHackAttemp("Tramposo SH: " & UserList(userindex).name & " , " & dummy)
                        Call SendData(SendTarget.ToAdmins, 0, 0, "PRT9," & UserList(userindex).name)
                        'Call CloseSocket(userindex)
                        Exit Sub
                    Else
                        UserList(userindex).flags.CountSH = TempTick
                    End If
                End If
                UserList(userindex).flags.StartWalk = TempTick
                UserList(userindex).flags.TimesWalk = 0
            End If
            
            UserList(userindex).flags.TimesWalk = UserList(userindex).flags.TimesWalk + 1
            
            rdata = Right$(rdata, Len(rdata) - 1)
            
            'salida parche
            If UserList(userindex).Counters.Saliendo And UserList(userindex).flags.Ban = 0 Then
                Call SendData(SendTarget.ToIndex, userindex, 0, "PRT8")
                UserList(userindex).Counters.Saliendo = False
                UserList(userindex).Counters.Salir = 0
            End If
            
            If UserList(userindex).flags.Paralizado = 0 Then
                If Not UserList(userindex).flags.Descansar And Not UserList(userindex).flags.Meditando Then
                    Call MoveUserChar(userindex, val(rdata))
                ElseIf UserList(userindex).flags.Descansar Then
                    UserList(userindex).flags.Descansar = False
                    Call SendData(SendTarget.ToIndex, userindex, 0, "DOK")
                    Call SendData(SendTarget.ToIndex, userindex, 0, "||Has dejado de descansar." & FONTTYPE_INFO)
                    Call MoveUserChar(userindex, val(rdata))
                ElseIf UserList(userindex).flags.Meditando Then
                    UserList(userindex).flags.Meditando = False
                    Call SendData(SendTarget.ToIndex, userindex, 0, "MEDOK")
                    Call SendData(SendTarget.ToIndex, userindex, 0, "PRE38")
                    UserList(userindex).Char.FX = 0
                    UserList(userindex).Char.loops = 0
                    Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "CFX" & UserList(userindex).Char.CharIndex & "," & 0 & "," & 0)
                End If
            Else    'paralizado
                '[CDT 17-02-2004] (<- emmmmm ?????)
                If Not UserList(userindex).flags.UltimoMensaje = 1 Then
                    Call SendData(SendTarget.ToIndex, userindex, 0, "PRE39")
                    UserList(userindex).flags.UltimoMensaje = 1
                End If
                '[/CDT]
                UserList(userindex).flags.CountSH = 0
            End If
            
            If UserList(userindex).flags.Oculto = 1 Then
                If UCase$(UserList(userindex).Clase) <> "LADRON" Then
                    UserList(userindex).flags.Oculto = 0
                    If UserList(userindex).flags.Invisible = 0 Then
                        Call SendData(SendTarget.ToIndex, userindex, 0, "PRB39")
                        Call SendData(SendTarget.ToMap, 0, UserList(userindex).pos.Map, "NOVER" & UserList(userindex).Char.CharIndex & ",0")
                        'If Guilds(UserList(UserIndex).GuildIndex).GuildName <> "" Then Call SendData(SendTarget.toguildmembers, UserIndex, UserList(UserIndex).Pos.Map, "PUDVE" & UserList(UserIndex).Char.CharIndex & ",0")
                    End If
                End If
            End If
            
            If UserList(userindex).flags.Muerto = 1 Then
                Call Empollando(userindex)
            Else
                UserList(userindex).flags.EstaEmpo = 0
                UserList(userindex).EmpoCont = 0
            End If
            Exit Sub
    End Select
    
     Select Case UCase$(Left$(rdata, 6))
    Case "ACHEAT"
            rdata = Right$(rdata, Len(rdata) - 6)
            If UserList(userindex).flags.Privilegios = PlayerType.User Then
                'UserList(UserIndex).flags.Ban = 1
                Call SendData(SendTarget.toall, 0, 0, "PRT7," & UserList(userindex).name)
                Call Cerrar_Usuario(userindex, 99)
                Call LogCheat(userindex, rdata)
            End If
            Exit Sub
    End Select
    
    Select Case UCase$(rdata)
        Case "RPU" 'Pedido de actualizacion de la posicion
            Call SendData(SendTarget.ToIndex, userindex, 0, "PU" & UserList(userindex).pos.x & "," & UserList(userindex).pos.Y)
            Exit Sub
        Case "AT"
            If UserList(userindex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, userindex, 0, "||¡¡No podes atacar a nadie porque estas muerto!!. " & FONTTYPE_INFO)
                Exit Sub
            End If
            If Not UserList(userindex).flags.ModoCombate Then
               Call SendData(SendTarget.ToIndex, userindex, 0, "PRE40")
            Else
                If UserList(userindex).Invent.WeaponEqpObjIndex > 0 Then
                    If ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).proyectil = 1 Then
                        Call SendData(SendTarget.ToIndex, userindex, 0, "||No podés usar asi esta arma." & FONTTYPE_INFO)
                        Exit Sub
                    End If
                End If
                Call UsuarioAtaca(userindex)
                
                'piedra libre para todos los compas!
                If UserList(userindex).flags.Oculto > 0 And UserList(userindex).flags.AdminInvisible = 0 Then
                    UserList(userindex).flags.Oculto = 0
                    If UserList(userindex).flags.Invisible = 0 Then
                        Call SendData(SendTarget.ToMap, 0, UserList(userindex).pos.Map, "NOVER" & UserList(userindex).Char.CharIndex & ",0")
                       Call SendData(SendTarget.ToIndex, userindex, 0, "PRB39")
                    End If
                End If
                
            End If
            Exit Sub
        Case "AG"
            If UserList(userindex).flags.Muerto = 1 Then
                    Call SendData(SendTarget.ToIndex, userindex, 0, "||¡¡Estas muerto!! Los muertos no pueden tomar objetos. " & FONTTYPE_INFO)
                    Exit Sub
            End If
            '[Consejeros]
            If UserList(userindex).flags.Privilegios = PlayerType.Consejero And Not UserList(userindex).flags.EsRolesMaster Then
                Call SendData(SendTarget.ToIndex, userindex, 0, "||No puedes tomar ningun objeto. " & FONTTYPE_INFO)
                Exit Sub
            End If
            Call GetObj(userindex)
            Exit Sub
         Case "TAB" 'Entrar o salir modo combate
            If UserList(userindex).flags.ModoCombate Then
                SendData SendTarget.ToIndex, userindex, 0, "PRE41"
            Else
                SendData SendTarget.ToIndex, userindex, 0, "PRE42"
            End If
            UserList(userindex).flags.ModoCombate = Not UserList(userindex).flags.ModoCombate
            Exit Sub
        Case "SEG" 'Activa / desactiva el seguro
            If UserList(userindex).flags.Seguro Then
                Call SendData(SendTarget.ToIndex, userindex, 0, "PRE43")
            Else
                Call SendData(SendTarget.ToIndex, userindex, 0, "SEGON")
                UserList(userindex).flags.Seguro = Not UserList(userindex).flags.Seguro
            End If
            Exit Sub
        Case "ACTUALIZAR"
            Call SendData(SendTarget.ToIndex, userindex, 0, "PU" & UserList(userindex).pos.x & "," & UserList(userindex).pos.Y)
            Exit Sub
        Case "GLINFO"
            tStr = SendGuildLeaderInfo(userindex)
            If tStr = vbNullString Then
                Call SendData(SendTarget.ToIndex, userindex, 0, "GL" & SendGuildsList(userindex))
            Else
                Call SendData(SendTarget.ToIndex, userindex, 0, "LEADERI" & tStr)
            End If
            Exit Sub
        Case "ATRI"
            Call EnviarAtrib(userindex)
            Exit Sub
        Case "FAMA"
            Call EnviarFama(userindex)
            Exit Sub
        Case "ESKI"
            Call EnviarSkills(userindex)
            Exit Sub
        Case "FEST" 'Mini estadisticas :)
            Call EnviarMiniEstadisticas(userindex)
            Exit Sub
        Case "FINSUB"
            'User sale del modo COMERCIO
            'UserList(UserIndex).flags.Comerciando = False
            Subastando = False
            Call SendData(SendTarget.ToIndex, userindex, 0, "FINSUBOK")
            Exit Sub
        '[Alejo]
        Case "FINCOM"
            'User sale del modo COMERCIO
            UserList(userindex).flags.Comerciando = False
            Call SendData(SendTarget.ToIndex, userindex, 0, "FINCOMOK")
            Exit Sub
        Case "FINCOMUSU"
            'Sale modo comercio Usuario
            If UserList(userindex).ComUsu.DestUsu > 0 And _
                UserList(UserList(userindex).ComUsu.DestUsu).ComUsu.DestUsu = userindex Then
                Call SendData(SendTarget.ToIndex, UserList(userindex).ComUsu.DestUsu, 0, "||" & UserList(userindex).name & " ha dejado de comerciar con vos." & FONTTYPE_TALK)
                Call FinComerciarUsu(UserList(userindex).ComUsu.DestUsu)
            End If
            
            Call FinComerciarUsu(userindex)
            Exit Sub
        '[KEVIN]---------------------------------------
        '******************************************************
                Case "INIBOV"
            Call SendUserStatsBox(userindex)
            Call IniciarDeposito(userindex)
            Exit Sub
        Case "FINBAN"
            'User sale del modo BANCO
            UserList(userindex).flags.Comerciando = False
            Call SendData(SendTarget.ToIndex, userindex, 0, "FINBANOK")
            Exit Sub
        '-------------------------------------------------------
        '[/KEVIN]**************************************
        Case "COMUSUOK"
            'Aceptar el cambio
            Call AceptarComercioUsu(userindex)
            Exit Sub
        Case "COMUSUNO"
            'Rechazar el cambio
            If UserList(userindex).ComUsu.DestUsu > 0 Then
                If UserList(UserList(userindex).ComUsu.DestUsu).flags.UserLogged Then
                    Call SendData(SendTarget.ToIndex, UserList(userindex).ComUsu.DestUsu, 0, "||" & UserList(userindex).name & " ha rechazado tu oferta." & FONTTYPE_TALK)
                    Call FinComerciarUsu(UserList(userindex).ComUsu.DestUsu)
                End If
            End If
            Call SendData(SendTarget.ToIndex, userindex, 0, "||Has rechazado la oferta del otro usuario." & FONTTYPE_TALK)
            Call FinComerciarUsu(userindex)
            Exit Sub
        '[/Alejo]
    
    
    End Select
    
    Select Case UCase$(Left$(rdata, 5))
    
  Case "KOTO1"
  Dim Sarasa As Obj
Sarasa.Amount = 1
Sarasa.ObjIndex = 1000
If UserList(userindex).Stats.PuntosTorneo < 25 Then
Call SendData(SendTarget.ToIndex, userindex, 0, "||No tienes suficientes puntos de torneo!." & FONTTYPE_INFO)
Else
UserList(userindex).Stats.PuntosTorneo = UserList(userindex).Stats.PuntosTorneo - 25
Call MeterItemEnInventario(userindex, Sarasa)
End If
Exit Sub
    
Case "KOTO2"
Sarasa.Amount = 1
Sarasa.ObjIndex = 991

If UserList(userindex).Stats.PuntosTorneo < 50 Then
Call SendData(SendTarget.ToIndex, userindex, 0, "||No tienes suficientes puntos de torneo!." & FONTTYPE_INFO)
Else
UserList(userindex).Stats.PuntosTorneo = UserList(userindex).Stats.PuntosTorneo - 50
Call MeterItemEnInventario(userindex, Sarasa)
End If
Exit Sub

Case "KOTO3"
Sarasa.Amount = 1
Sarasa.ObjIndex = 1017

If UserList(userindex).Stats.PuntosTorneo < 15 Then
Call SendData(SendTarget.ToIndex, userindex, 0, "||No tienes suficientes puntos de torneo!." & FONTTYPE_INFO)
Else
UserList(userindex).Stats.PuntosTorneo = UserList(userindex).Stats.PuntosTorneo - 15
Call MeterItemEnInventario(userindex, Sarasa)
End If
Exit Sub

Case "KOTO4"
Sarasa.Amount = 1
Sarasa.ObjIndex = 1012

If UserList(userindex).Stats.PuntosTorneo < 10 Then
Call SendData(SendTarget.ToIndex, userindex, 0, "||No tienes suficientes puntos de torneo!." & FONTTYPE_INFO)
Else
UserList(userindex).Stats.PuntosTorneo = UserList(userindex).Stats.PuntosTorneo - 10
Call MeterItemEnInventario(userindex, Sarasa)
End If
Exit Sub

Case "KOTO5"
Sarasa.Amount = 1
Sarasa.ObjIndex = 999

If UserList(userindex).Stats.PuntosTorneo < 5 Then
Call SendData(SendTarget.ToIndex, userindex, 0, "||No tienes suficientes puntos de torneo!." & FONTTYPE_INFO)
Else
UserList(userindex).Stats.PuntosTorneo = UserList(userindex).Stats.PuntosTorneo - 5
Call MeterItemEnInventario(userindex, Sarasa)
End If
Exit Sub

Case "KOTO6"

If UserList(userindex).Stats.PuntosTorneo < 5 Then
Call SendData(SendTarget.ToIndex, userindex, 0, "||No tienes suficientes puntos de torneo!." & FONTTYPE_INFO)
Else
UserList(userindex).Stats.PuntosTorneo = UserList(userindex).Stats.PuntosTorneo - 5
UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD + 500000

End If
Exit Sub

Case "KOTO7"
Sarasa.Amount = 1
Sarasa.ObjIndex = 1021

If UserList(userindex).Stats.PuntosTorneo < 40 Then
Call SendData(SendTarget.ToIndex, userindex, 0, "||No tienes suficientes puntos de torneo!." & FONTTYPE_INFO)
Else
UserList(userindex).Stats.PuntosTorneo = UserList(userindex).Stats.PuntosTorneo - 40
Call MeterItemEnInventario(userindex, Sarasa)
End If
Exit Sub

Case "KOTO8"
Sarasa.Amount = 1
Sarasa.ObjIndex = 1055

If UserList(userindex).Stats.PuntosTorneo < 55 Then
Call SendData(SendTarget.ToIndex, userindex, 0, "||No tienes suficientes puntos de torneo!." & FONTTYPE_INFO)
Else
UserList(userindex).Stats.PuntosTorneo = UserList(userindex).Stats.PuntosTorneo - 55
Call MeterItemEnInventario(userindex, Sarasa)
End If
Exit Sub

Case "KOTO9"
Sarasa.Amount = 1
Sarasa.ObjIndex = 1066

If UserList(userindex).Stats.PuntosTorneo < 40 Then
Call SendData(SendTarget.ToIndex, userindex, 0, "||No tienes suficientes puntos de torneo!." & FONTTYPE_INFO)
Else
UserList(userindex).Stats.PuntosTorneo = UserList(userindex).Stats.PuntosTorneo - 40
Call MeterItemEnInventario(userindex, Sarasa)
End If
Exit Sub

Case "DONA1"
Sarasa.Amount = 1
Sarasa.ObjIndex = 1245

If UserList(userindex).Stats.PuntosTorneo < 35 Then
Call SendData(SendTarget.ToIndex, userindex, 0, "||No tienes suficientes puntos de torneo!." & FONTTYPE_INFO)
Else
UserList(userindex).Stats.PuntosTorneo = UserList(userindex).Stats.PuntosTorneo - 35
Call MeterItemEnInventario(userindex, Sarasa)
End If
Exit Sub

Case "DONA2"
Sarasa.Amount = 1
Sarasa.ObjIndex = 1261

If UserList(userindex).Stats.PuntosTorneo < 50 Then
Call SendData(SendTarget.ToIndex, userindex, 0, "||No tienes suficientes puntos de torneo!." & FONTTYPE_INFO)
Else
UserList(userindex).Stats.PuntosTorneo = UserList(userindex).Stats.PuntosTorneo - 50
Call MeterItemEnInventario(userindex, Sarasa)
End If
Exit Sub

Case "DONA3"
Sarasa.Amount = 1
Sarasa.ObjIndex = 1082

If UserList(userindex).Stats.PuntosTorneo < 40 Then
Call SendData(SendTarget.ToIndex, userindex, 0, "||No tienes suficientes puntos de torneo!." & FONTTYPE_INFO)
Else
UserList(userindex).Stats.PuntosTorneo = UserList(userindex).Stats.PuntosTorneo - 40
Call MeterItemEnInventario(userindex, Sarasa)
End If
Exit Sub

Case "DONA4"
Sarasa.Amount = 1
Sarasa.ObjIndex = 1263

If UserList(userindex).Stats.PuntosTorneo < 40 Then
Call SendData(SendTarget.ToIndex, userindex, 0, "||No tienes suficientes puntos de torneo!." & FONTTYPE_INFO)
Else
UserList(userindex).Stats.PuntosTorneo = UserList(userindex).Stats.PuntosTorneo - 40
Call MeterItemEnInventario(userindex, Sarasa)
End If
Exit Sub

Case "DONA5"
Sarasa.Amount = 1
Sarasa.ObjIndex = 1262

If UserList(userindex).Stats.PuntosTorneo < 40 Then
Call SendData(SendTarget.ToIndex, userindex, 0, "||No tienes suficientes puntos de torneo!." & FONTTYPE_INFO)
Else
UserList(userindex).Stats.PuntosTorneo = UserList(userindex).Stats.PuntosTorneo - 40
Call MeterItemEnInventario(userindex, Sarasa)
End If
Exit Sub

Case "DONA6"
Sarasa.Amount = 1
Sarasa.ObjIndex = 1264

If UserList(userindex).Stats.PuntosTorneo < 40 Then
Call SendData(SendTarget.ToIndex, userindex, 0, "||No tienes suficientes puntos de torneo!." & FONTTYPE_INFO)
Else
UserList(userindex).Stats.PuntosTorneo = UserList(userindex).Stats.PuntosTorneo - 40
Call MeterItemEnInventario(userindex, Sarasa)
End If
Exit Sub

Case "DONA7"
Sarasa.Amount = 1
Sarasa.ObjIndex = 1039

If UserList(userindex).Stats.PuntosTorneo < 55 Then
Call SendData(SendTarget.ToIndex, userindex, 0, "||No tienes suficientes puntos de torneo!." & FONTTYPE_INFO)
Else
UserList(userindex).Stats.PuntosTorneo = UserList(userindex).Stats.PuntosTorneo - 55
Call MeterItemEnInventario(userindex, Sarasa)
End If
Exit Sub

Case "DONA8"
Sarasa.Amount = 1
Sarasa.ObjIndex = 1265

If UserList(userindex).Stats.PuntosTorneo < 55 Then
Call SendData(SendTarget.ToIndex, userindex, 0, "||No tienes suficientes puntos de torneo!." & FONTTYPE_INFO)
Else
UserList(userindex).Stats.PuntosTorneo = UserList(userindex).Stats.PuntosTorneo - 55
Call MeterItemEnInventario(userindex, Sarasa)
End If
Exit Sub

Case "DONA9"
Sarasa.Amount = 1
Sarasa.ObjIndex = 1266

If UserList(userindex).Stats.PuntosTorneo < 50 Then
Call SendData(SendTarget.ToIndex, userindex, 0, "||No tienes suficientes puntos de torneo!." & FONTTYPE_INFO)
Else
UserList(userindex).Stats.PuntosTorneo = UserList(userindex).Stats.PuntosTorneo - 50
Call MeterItemEnInventario(userindex, Sarasa)
End If
Exit Sub

Case "KWLF1"
Sarasa.Amount = 1
Sarasa.ObjIndex = 1267

If UserList(userindex).Stats.PuntosTorneo < 55 Then
Call SendData(SendTarget.ToIndex, userindex, 0, "||No tienes suficientes puntos de torneo!." & FONTTYPE_INFO)
Else
UserList(userindex).Stats.PuntosTorneo = UserList(userindex).Stats.PuntosTorneo - 55
Call MeterItemEnInventario(userindex, Sarasa)
End If
Exit Sub

End Select
    
    Select Case UCase$(Left$(rdata, 2))
    '    Case "/Z"
    '        Dim Pos As WorldPos, Pos2 As WorldPos
    '        Dim O As Obj
    '
    '        For LoopC = 1 To 100
    '            Pos = UserList(UserIndex).Pos
    '            O.Amount = 1
    '            O.ObjIndex = iORO
    '            'Exit For
    '            Call TirarOro(100000, UserIndex)
    '            'Call Tilelibre(Pos, Pos2)
    '            'If Pos2.x = 0 Or Pos2.y = 0 Then Exit For
    '
    '            'Call MakeObj(SendTarget.ToMap, 0, UserList(UserIndex).Pos.Map, O, Pos2.Map, Pos2.x, Pos2.y)
    '        Next LoopC
    '
    '        Exit Sub
    
    
        Case "TI" 'Tirar item
                If UserList(userindex).flags.Navegando = 1 Or _
                   UserList(userindex).flags.Muerto = 1 Or _
                   (UserList(userindex).flags.Privilegios = PlayerType.Consejero And Not UserList(userindex).flags.EsRolesMaster) Then Exit Sub
                   '[Consejeros]
                
                rdata = Right$(rdata, Len(rdata) - 2)
                Arg1 = ReadField(1, rdata, 44)
                Arg2 = ReadField(2, rdata, 44)
                If val(Arg1) = FLAGORO Then
                    
                    Call TirarOro(val(Arg2), userindex)
                    
                    Call SendUserStatsBox(userindex)
                    Exit Sub
                Else
                    If val(Arg1) <= MAX_INVENTORY_SLOTS And val(Arg1) > 0 Then
                        If UserList(userindex).Invent.Object(val(Arg1)).ObjIndex = 0 Then
                                Exit Sub
                        End If
                        Call DropObj(userindex, val(Arg1), val(Arg2), UserList(userindex).pos.Map, UserList(userindex).pos.x, UserList(userindex).pos.Y)
                    Else
                        Exit Sub
                    End If
                End If
                Exit Sub

        Case "HL" ' Lanzar hechizo
            If UserList(userindex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, userindex, 0, "||¡¡Estas muerto!!." & FONTTYPE_INFO)
                Exit Sub
            End If
            rdata = Right$(rdata, Len(rdata) - 2)
            UserList(userindex).flags.Hechizo = val(rdata)
            Exit Sub
        Case "LC" 'Click izquierdo
            rdata = Right$(rdata, Len(rdata) - 2)
            Arg1 = ReadField(1, rdata, 44)
            Arg2 = ReadField(2, rdata, 44)
            If Not Numeric(Arg1) Or Not Numeric(Arg2) Then Exit Sub
            x = CInt(Arg1)
            Y = CInt(Arg2)
            Call LookatTile(userindex, UserList(userindex).pos.Map, x, Y)
            Exit Sub
        Case "RC" 'Click derecho
            rdata = Right$(rdata, Len(rdata) - 2)
            Arg1 = ReadField(1, rdata, 44)
            Arg2 = ReadField(2, rdata, 44)
            If Not Numeric(Arg1) Or Not Numeric(Arg2) Then Exit Sub
            x = CInt(Arg1)
            Y = CInt(Arg2)
            Call Accion(userindex, UserList(userindex).pos.Map, x, Y)
            Exit Sub
        Case "KU"
            If UserList(userindex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, userindex, 0, "||¡¡Estas muerto!!." & FONTTYPE_INFO)
                Exit Sub
            End If
    
            rdata = Right$(rdata, Len(rdata) - 2)
            Select Case val(rdata)
                Case Robar
                    Call SendData(SendTarget.ToIndex, userindex, 0, "T01" & Robar)
                Case Magia
                    Call SendData(SendTarget.ToIndex, userindex, 0, "T01" & Magia)
                Case Domar
                    Call SendData(SendTarget.ToIndex, userindex, 0, "T01" & Domar)
                Case Ocultarse
                    If UserList(userindex).flags.Navegando = 1 Then
                        '[CDT 17-02-2004]
                        If Not UserList(userindex).flags.UltimoMensaje = 3 Then
                            Call SendData(SendTarget.ToIndex, userindex, 0, "||No podes ocultarte si estas navegando." & FONTTYPE_INFO)
                            UserList(userindex).flags.UltimoMensaje = 3
                        End If
                        '[/CDT]
                        Exit Sub
                    End If
                    
                    If UserList(userindex).flags.Oculto = 1 Then
                        '[CDT 17-02-2004]
                        If Not UserList(userindex).flags.UltimoMensaje = 2 Then
                            Call SendData(SendTarget.ToIndex, userindex, 0, "||Ya estas oculto." & FONTTYPE_INFO)
                            UserList(userindex).flags.UltimoMensaje = 2
                        End If
                        '[/CDT]
                        Exit Sub
                    End If
                    
                    Call DoOcultarse(userindex)
            End Select
            Exit Sub
    
    End Select
    
    Select Case UCase$(Left$(rdata, 3))
         Case "UMH" ' Usa macro de hechizos
            Call SendData(SendTarget.ToAdmins, userindex, 0, "||" & UserList(userindex).name & " fue expulsado por Anti-macro de hechizos " & FONTTYPE_VENENO)
            Call SendData(SendTarget.ToIndex, userindex, 0, "ERR Has sido expulsado por usar macro de hechizos. Recomendamos leer el reglamento sobre el tema macros" & FONTTYPE_INFO)
            Call CloseSocket(userindex)
            Exit Sub
        Case "USA"
            rdata = Right$(rdata, Len(rdata) - 3)
            If val(rdata) <= MAX_INVENTORY_SLOTS And val(rdata) > 0 Then
                If UserList(userindex).Invent.Object(val(rdata)).ObjIndex = 0 Then Exit Sub
            Else
                Exit Sub
            End If
            If UserList(userindex).flags.Meditando Then
                Call SendData(SendTarget.ToIndex, userindex, 0, "M!")
                Exit Sub
            End If
            Call UseInvItem(userindex, val(rdata))
            Exit Sub
        Case "CNS" ' Construye herreria
            rdata = Right$(rdata, Len(rdata) - 3)
            x = CInt(rdata)
            If x < 1 Then Exit Sub
            If ObjData(x).SkHerreria = 0 Then Exit Sub
            Call HerreroConstruirItem(userindex, x)
            Exit Sub
        Case "CNC" ' Construye carpinteria
        rdata = Right$(rdata, Len(rdata) - 3)
        x = ReadField(1, rdata, 44)
        If x < 1 Or ObjData(x).SkCarpinteria = 0 Then Exit Sub
        Call CarpinteroConstruirItem(userindex, x, val(ReadField(2, rdata, 44)))
        Exit Sub
        Case "WLC" 'Click izquierdo en modo trabajo
            rdata = Right$(rdata, Len(rdata) - 3)
            Arg1 = ReadField(1, rdata, 44)
            Arg2 = ReadField(2, rdata, 44)
            Arg3 = ReadField(3, rdata, 44)
            If Arg3 = "" Or Arg2 = "" Or Arg1 = "" Then Exit Sub
            If Not Numeric(Arg1) Or Not Numeric(Arg2) Or Not Numeric(Arg3) Then Exit Sub
            
            x = CInt(Arg1)
            Y = CInt(Arg2)
            tLong = CInt(Arg3)
            
            If UserList(userindex).flags.Muerto = 1 Or _
               UserList(userindex).flags.Descansar Or _
               UserList(userindex).flags.Meditando Or _
               Not InMapBounds(UserList(userindex).pos.Map, x, Y) Then Exit Sub
            
            If Not InRangoVision(userindex, x, Y) Then
                Call SendData(SendTarget.ToIndex, userindex, 0, "PU" & UserList(userindex).pos.x & "," & UserList(userindex).pos.Y)
                Exit Sub
            End If
            
            Select Case tLong
            
            Case Proyectiles
                Dim TU As Integer, tN As Integer
                'Nos aseguramos que este usando un arma de proyectiles
                If Not IntervaloPermiteAtacar(userindex, False) Or Not IntervaloPermiteUsarArcos(userindex) Then
                    Exit Sub
                End If

                DummyInt = 0

                If UserList(userindex).Invent.WeaponEqpObjIndex = 0 Then
                    DummyInt = 1
                ElseIf UserList(userindex).Invent.WeaponEqpSlot < 1 Or UserList(userindex).Invent.WeaponEqpSlot > MAX_INVENTORY_SLOTS Then
                    DummyInt = 1
                ElseIf UserList(userindex).Invent.MunicionEqpSlot < 1 Or UserList(userindex).Invent.MunicionEqpSlot > MAX_INVENTORY_SLOTS Then
                    DummyInt = 1
                ElseIf UserList(userindex).Invent.MunicionEqpObjIndex = 0 Then
                    DummyInt = 1
                ElseIf ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).proyectil <> 1 Then
                    DummyInt = 2
                ElseIf ObjData(UserList(userindex).Invent.MunicionEqpObjIndex).OBJType <> eOBJType.otFlechas Then
                    DummyInt = 1
                ElseIf UserList(userindex).Invent.Object(UserList(userindex).Invent.MunicionEqpSlot).Amount < 1 Then
                    DummyInt = 1
                End If
                
                If DummyInt <> 0 Then
                    If DummyInt = 1 Then
                        Call SendData(SendTarget.ToIndex, userindex, 0, "||No tenes municiones." & FONTTYPE_INFO)
                    End If
                    Call Desequipar(userindex, UserList(userindex).Invent.MunicionEqpSlot)
                    Call Desequipar(userindex, UserList(userindex).Invent.WeaponEqpSlot)
                    Exit Sub
                End If
                
                DummyInt = 0
               'Quitamos stamina
                If UserList(userindex).Stats.MinSta >= 10 Then
                     Call QuitarSta(userindex, RandomNumber(1, 10))
                Else
                     Call SendData(SendTarget.ToIndex, userindex, 0, "PRE44")
                     Exit Sub
                End If
                 
                Call LookatTile(userindex, UserList(userindex).pos.Map, Arg1, Arg2)
                
                TU = UserList(userindex).flags.TargetUser
                tN = UserList(userindex).flags.TargetNPC
                
                'Sólo permitimos atacar si el otro nos puede atacar también
                If TU > 0 Then
                    If Abs(UserList(UserList(userindex).flags.TargetUser).pos.Y - UserList(userindex).pos.Y) > RANGO_VISION_Y Then
                        Call SendData(SendTarget.ToIndex, userindex, 0, "||Estas demasiado lejos para atacar." & FONTTYPE_WARNING)
                        Exit Sub
                    End If
                ElseIf tN > 0 Then
                    If Abs(Npclist(UserList(userindex).flags.TargetNPC).pos.Y - UserList(userindex).pos.Y) > RANGO_VISION_Y Then
                        Call SendData(SendTarget.ToIndex, userindex, 0, "||Estas demasiado lejos para atacar." & FONTTYPE_WARNING)
                        Exit Sub
                    End If
                End If
                
                
                If TU > 0 Then
                    'Previene pegarse a uno mismo
                    If TU = userindex Then
                        Call SendData(SendTarget.ToIndex, userindex, 0, "||¡No puedes atacarte a vos mismo!" & FONTTYPE_INFO)
                        DummyInt = 1
                        Exit Sub
                    End If
                End If
    
                If DummyInt = 0 Then
                    'Saca 1 flecha
                    DummyInt = UserList(userindex).Invent.MunicionEqpSlot
                    Call QuitarUserInvItem(userindex, UserList(userindex).Invent.MunicionEqpSlot, 1)
                    If DummyInt < 1 Or DummyInt > MAX_INVENTORY_SLOTS Then Exit Sub
                    If UserList(userindex).Invent.Object(DummyInt).Amount > 0 Then
                        UserList(userindex).Invent.Object(DummyInt).Equipped = 1
                        UserList(userindex).Invent.MunicionEqpSlot = DummyInt
                        UserList(userindex).Invent.MunicionEqpObjIndex = UserList(userindex).Invent.Object(DummyInt).ObjIndex
                        Call UpdateUserInv(False, userindex, UserList(userindex).Invent.MunicionEqpSlot)
                    Else
                        Call UpdateUserInv(False, userindex, DummyInt)
                        UserList(userindex).Invent.MunicionEqpSlot = 0
                        UserList(userindex).Invent.MunicionEqpObjIndex = 0
                    End If
                    '-----------------------------------
                End If

                If tN > 0 Then
                    If Npclist(tN).Attackable <> 0 Then
                        Call UsuarioAtacaNpc(userindex, tN)
                    End If
                ElseIf TU > 0 Then
                    If UserList(userindex).flags.Seguro Then
                        If Not Criminal(TU) Then
                            Call SendData(SendTarget.ToIndex, userindex, 0, "||¡Para atacar ciudadanos desactiva el seguro!" & FONTTYPE_FIGHT)
                            Exit Sub
                        End If
                    End If
                    Call UsuarioAtacaUsuario(userindex, TU)
                End If
                
            Case Magia
                If MapInfo(UserList(userindex).pos.Map).MagiaSinEfecto > 0 Then
                    Call SendData(SendTarget.ToIndex, userindex, 0, "||Una fuerza oscura te impide canalizar tu energía" & FONTTYPE_FIGHT)
                    Exit Sub
                End If
                Call LookatTile(userindex, UserList(userindex).pos.Map, x, Y)
                
                'MmMmMmmmmM
                Dim wp2 As WorldPos
                wp2.Map = UserList(userindex).pos.Map
                wp2.x = x
                wp2.Y = Y
                                
                If UserList(userindex).flags.Hechizo > 0 Then
                    If IntervaloPermiteLanzarSpell(userindex) Then
                        Call LanzarHechizo(UserList(userindex).flags.Hechizo, userindex)
                    '    UserList(UserIndex).flags.PuedeLanzarSpell = 0
                        UserList(userindex).flags.Hechizo = 0
                    End If
                Else
                   Call SendData(SendTarget.ToIndex, userindex, 0, "PRE45")
                End If
                
                'If Distancia(UserList(UserIndex).Pos, wp2) > 10 Then
                If (Abs(UserList(userindex).pos.x - wp2.x) > 9 Or Abs(UserList(userindex).pos.Y - wp2.Y) > 8) Then
                    Dim txt As String
                    txt = "Ataque fuera de rango de " & UserList(userindex).name & "(" & UserList(userindex).pos.Map & "/" & UserList(userindex).pos.x & "/" & UserList(userindex).pos.Y & ") ip: " & UserList(userindex).ip & " a la posicion (" & wp2.Map & "/" & wp2.x & "/" & wp2.Y & ") "
                    If UserList(userindex).flags.Hechizo > 0 Then
                        txt = txt & ". Hechizo: " & Hechizos(UserList(userindex).Stats.UserHechizos(UserList(userindex).flags.Hechizo)).Nombre
                    End If
                    If MapData(wp2.Map, wp2.x, wp2.Y).userindex > 0 Then
                        txt = txt & " hacia el usuario: " & UserList(MapData(wp2.Map, wp2.x, wp2.Y).userindex).name
                    ElseIf MapData(wp2.Map, wp2.x, wp2.Y).NpcIndex > 0 Then
                        txt = txt & " hacia el NPC: " & Npclist(MapData(wp2.Map, wp2.x, wp2.Y).NpcIndex).name
                    End If
                    
                    Call LogCheating(txt)
                End If
                
            
            
            
            Case Pesca
                        
                AuxInd = UserList(userindex).Invent.HerramientaEqpObjIndex
                If AuxInd = 0 Then Exit Sub
                
                'If UserList(UserIndex).flags.PuedeTrabajar = 0 Then Exit Sub
                If Not IntervaloPermiteTrabajar(userindex) Then Exit Sub
                
                If AuxInd <> CAÑA_PESCA And AuxInd <> RED_PESCA Then
                    'Call Cerrar_Usuario(UserIndex)
                    ' Podemos llegar acá si el user equipó el anillo dsp de la U y antes del click
                    Exit Sub
                End If
                
                'Basado en la idea de Barrin
                'Comentario por Barrin: jah, "basado", caradura ! ^^
                If MapData(UserList(userindex).pos.Map, UserList(userindex).pos.x, UserList(userindex).pos.Y).trigger = 1 Then
                    Call SendData(SendTarget.ToIndex, userindex, 0, "||No puedes pescar desde donde te encuentras." & FONTTYPE_INFO)
                    Exit Sub
                End If
                
                If HayAgua(UserList(userindex).pos.Map, x, Y) Then
                    Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "TW" & SND_PESCAR)
                    
                    Select Case AuxInd
                    Case CAÑA_PESCA
                        Call DoPescar(userindex)
                    Case RED_PESCA
                        With UserList(userindex)
                            wpaux.Map = .pos.Map
                            wpaux.x = x
                            wpaux.Y = Y
                        End With
                        
                        If Distancia(UserList(userindex).pos, wpaux) > 2 Then
                            Call SendData(SendTarget.ToIndex, userindex, 0, "||Estás demasiado lejos para pescar." & FONTTYPE_INFO)
                            Exit Sub
                        End If
                        
                        Call DoPescarRed(userindex)
                    End Select
    
                Else
                    Call SendData(SendTarget.ToIndex, userindex, 0, "||No hay agua donde pescar busca un lago, rio o mar." & FONTTYPE_INFO)
                End If
                
            Case Robar
               If MapInfo(UserList(userindex).pos.Map).Pk Then
                    'If UserList(UserIndex).flags.PuedeTrabajar = 0 Then Exit Sub
                    If Not IntervaloPermiteTrabajar(userindex) Then Exit Sub
                    
                    Call LookatTile(userindex, UserList(userindex).pos.Map, x, Y)
                    
                    If UserList(userindex).flags.TargetUser > 0 And UserList(userindex).flags.TargetUser <> userindex Then
                       If UserList(UserList(userindex).flags.TargetUser).flags.Muerto = 0 Then
                            wpaux.Map = UserList(userindex).pos.Map
                            wpaux.x = val(ReadField(1, rdata, 44))
                            wpaux.Y = val(ReadField(2, rdata, 44))
                            If Distancia(wpaux, UserList(userindex).pos) > 2 Then
                                Call SendData(SendTarget.ToIndex, userindex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
                                Exit Sub
                            End If
                            '17/09/02
                            'No aseguramos que el trigger le permite robar
                            If MapData(UserList(UserList(userindex).flags.TargetUser).pos.Map, UserList(UserList(userindex).flags.TargetUser).pos.x, UserList(UserList(userindex).flags.TargetUser).pos.Y).trigger = eTrigger.ZONASEGURA Then
                                Call SendData(SendTarget.ToIndex, userindex, 0, "||No podes robar aquí." & FONTTYPE_WARNING)
                                Exit Sub
                            End If
                            If MapData(UserList(userindex).pos.Map, UserList(userindex).pos.x, UserList(userindex).pos.Y).trigger = eTrigger.ZONASEGURA Then
                                Call SendData(SendTarget.ToIndex, userindex, 0, "||No podes robar aquí." & FONTTYPE_WARNING)
                                Exit Sub
                            End If
                            
                            Call DoRobar(userindex, UserList(userindex).flags.TargetUser)
                       End If
                    Else
                        Call SendData(SendTarget.ToIndex, userindex, 0, "||No a quien robarle!." & FONTTYPE_INFO)
                    End If
                Else
                    Call SendData(SendTarget.ToIndex, userindex, 0, "||¡No podes robarle en zonas seguras!." & FONTTYPE_INFO)
                End If
            Case Talar
            
                               If MapInfo(UserList(userindex).pos.Map).Pk = False Then
                    Call SendData(ToIndex, userindex, 0, "||No está permitido talar en zonas seguras!" & FONTTYPE_INFO)
                    Exit Sub
                End If
                
                'If UserList(UserIndex).flags.PuedeTrabajar = 0 Then Exit Sub
                If Not IntervaloPermiteTrabajar(userindex) Then Exit Sub
                
                If UserList(userindex).Invent.HerramientaEqpObjIndex = 0 Then
                    Call SendData(SendTarget.ToIndex, userindex, 0, "||Deberías equiparte el hacha." & FONTTYPE_INFO)
                    Exit Sub
                End If
                
                If UserList(userindex).Invent.HerramientaEqpObjIndex <> HACHA_LEÑADOR Then
                    ' Call Cerrar_Usuario(UserIndex)
                    ' Podemos llegar acá si el user equipó el anillo dsp de la U y antes del click
                    Exit Sub
                End If
                
                AuxInd = MapData(UserList(userindex).pos.Map, x, Y).OBJInfo.ObjIndex
                If AuxInd > 0 Then
                    wpaux.Map = UserList(userindex).pos.Map
                    wpaux.x = x
                    wpaux.Y = Y
                    If Distancia(wpaux, UserList(userindex).pos) > 2 Then
                        Call SendData(SendTarget.ToIndex, userindex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    'Barrin 29/9/03
                    If Distancia(wpaux, UserList(userindex).pos) = 0 Then
                        Call SendData(SendTarget.ToIndex, userindex, 0, "||No podes talar desde allí." & FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    '¿Hay un arbol donde clickeo?
                    If ObjData(AuxInd).OBJType = eOBJType.otArboles Then
                        Call SendData(SendTarget.ToPCArea, CInt(userindex), UserList(userindex).pos.Map, "TW" & SND_TALAR)
                        Call DoTalar(userindex)
                    End If
                Else
                    Call SendData(SendTarget.ToIndex, userindex, 0, "||No hay ningun arbol ahi." & FONTTYPE_INFO)
                End If
            Case Mineria
                
                'If UserList(UserIndex).flags.PuedeTrabajar = 0 Then Exit Sub
                If Not IntervaloPermiteTrabajar(userindex) Then Exit Sub
                                
                If UserList(userindex).Invent.HerramientaEqpObjIndex = 0 Then Exit Sub
                
                If UserList(userindex).Invent.HerramientaEqpObjIndex <> PIQUETE_MINERO Then
                    ' Call Cerrar_Usuario(UserIndex)
                    ' Podemos llegar acá si el user equipó el anillo dsp de la U y antes del click
                    Exit Sub
                End If
                
                Call LookatTile(userindex, UserList(userindex).pos.Map, x, Y)
                
                AuxInd = MapData(UserList(userindex).pos.Map, x, Y).OBJInfo.ObjIndex
                If AuxInd > 0 Then
                    wpaux.Map = UserList(userindex).pos.Map
                    wpaux.x = x
                    wpaux.Y = Y
                    If Distancia(wpaux, UserList(userindex).pos) > 2 Then
                        Call SendData(SendTarget.ToIndex, userindex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
                        Exit Sub
                    End If
                    '¿Hay un yacimiento donde clickeo?
                    If ObjData(AuxInd).OBJType = eOBJType.otYacimiento Then
                        Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "TW" & SND_MINERO)
                        Call DoMineria(userindex)
                    Else
                        Call SendData(SendTarget.ToIndex, userindex, 0, "||Ahi no hay ningun yacimiento." & FONTTYPE_INFO)
                    End If
                Else
                    Call SendData(SendTarget.ToIndex, userindex, 0, "||Ahi no hay ningun yacimiento." & FONTTYPE_INFO)
                End If
            Case Domar
              'Modificado 25/11/02
              'Optimizado y solucionado el bug de la doma de
              'criaturas hostiles.
              Dim CI As Integer
              
              Call LookatTile(userindex, UserList(userindex).pos.Map, x, Y)
              CI = UserList(userindex).flags.TargetNPC
              
              If CI > 0 Then
                       If Npclist(CI).flags.Domable > 0 Then
                            wpaux.Map = UserList(userindex).pos.Map
                            wpaux.x = x
                            wpaux.Y = Y
                            If Distancia(wpaux, Npclist(UserList(userindex).flags.TargetNPC).pos) > 2 Then
                                  Call SendData(SendTarget.ToIndex, userindex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
                                  Exit Sub
                            End If
                            If Npclist(CI).flags.AttackedBy <> "" Then
                                  Call SendData(SendTarget.ToIndex, userindex, 0, "||No podés domar una criatura que está luchando con un jugador." & FONTTYPE_INFO)
                                  Exit Sub
                            End If
                            Call DoDomar(userindex, CI)
                        Else
                            Call SendData(SendTarget.ToIndex, userindex, 0, "||No podes domar a esa criatura." & FONTTYPE_INFO)
                        End If
              Else
                     Call SendData(SendTarget.ToIndex, userindex, 0, "||No hay ninguna criatura alli!." & FONTTYPE_INFO)
              End If
              
            Case FundirMetal
                'Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, X, Y)
                If Not IntervaloPermiteTrabajar(userindex) Then Exit Sub
                
                If UserList(userindex).flags.TargetObj > 0 Then
                    If ObjData(UserList(userindex).flags.TargetObj).OBJType = eOBJType.otFragua Then
                        ''chequeamos que no se zarpe duplicando oro
                        If UserList(userindex).Invent.Object(UserList(userindex).flags.TargetObjInvSlot).ObjIndex <> UserList(userindex).flags.TargetObjInvIndex Then
                            If UserList(userindex).Invent.Object(UserList(userindex).flags.TargetObjInvSlot).ObjIndex = 0 Or UserList(userindex).Invent.Object(UserList(userindex).flags.TargetObjInvSlot).Amount = 0 Then
                                Call SendData(SendTarget.ToIndex, userindex, 0, "||No tienes mas minerales" & FONTTYPE_INFO)
                                Exit Sub
                            End If
                            
                            ''FUISTE
                            'Call Ban(UserList(UserIndex).Name, "Sistema anti cheats", "Intento de duplicacion de items")
                            'Call LogCheating(UserList(UserIndex).Name & " intento crear minerales a partir de otros: FlagSlot/usaba/usoconclick/cantidad/IP:" & UserList(UserIndex).flags.TargetObjInvSlot & "/" & UserList(UserIndex).flags.TargetObjInvIndex & "/" & UserList(UserIndex).Invent.Object(UserList(UserIndex).flags.TargetObjInvSlot).ObjIndex & "/" & UserList(UserIndex).Invent.Object(UserList(UserIndex).flags.TargetObjInvSlot).Amount & "/" & UserList(UserIndex).ip)
                            'UserList(UserIndex).flags.Ban = 1
                            'Call SendData(SendTarget.ToAll, 0, 0, "||>>>> El sistema anti-cheats baneó a " & UserList(UserIndex).Name & " (intento de duplicación). Ip Logged. " & FONTTYPE_FIGHT)
                            Call SendData(SendTarget.ToIndex, userindex, 0, "ERRHas sido expulsado por el sistema anti cheats. Reconéctate.")
                            Call CloseSocket(userindex)
                            Exit Sub
                        End If
                        Call FundirMineral(userindex)
                    Else
                        Call SendData(SendTarget.ToIndex, userindex, 0, "||Ahi no hay ninguna fragua." & FONTTYPE_INFO)
                    End If
                Else
                    Call SendData(SendTarget.ToIndex, userindex, 0, "||Ahi no hay ninguna fragua." & FONTTYPE_INFO)
                End If
                
            Case Herreria
                Call LookatTile(userindex, UserList(userindex).pos.Map, x, Y)
                
                If UserList(userindex).flags.TargetObj > 0 Then
                    If ObjData(UserList(userindex).flags.TargetObj).OBJType = eOBJType.otYunque Then
                        Call EnivarArmasConstruibles(userindex)
                        Call EnivarArmadurasConstruibles(userindex)
                        Call SendData(SendTarget.ToIndex, userindex, 0, "SFH")
                    Else
                        Call SendData(SendTarget.ToIndex, userindex, 0, "||Ahi no hay ningun yunque." & FONTTYPE_INFO)
                    End If
                Else
                    Call SendData(SendTarget.ToIndex, userindex, 0, "||Ahi no hay ningun yunque." & FONTTYPE_INFO)
                End If
                
            End Select
            
            'UserList(UserIndex).flags.PuedeTrabajar = 0
            Exit Sub
        Case "CIG"
            rdata = Right$(rdata, Len(rdata) - 3)
            
            If modGuilds.CrearNuevoClan(rdata, userindex, UserList(userindex).FundandoGuildAlineacion, tStr) Then
                Call SendData(SendTarget.toall, 0, 0, "||" & UserList(userindex).name & " fundó el clan " & Guilds(UserList(userindex).GuildIndex).GuildName & " de alineación " & Alineacion2String(Guilds(UserList(userindex).GuildIndex).Alineacion) & "." & FONTTYPE_GUILD)
            Else
                Call SendData(SendTarget.ToIndex, userindex, 0, "||" & tStr & FONTTYPE_GUILD)
            End If
            
            Exit Sub
    End Select
    
    Select Case UCase$(Left$(rdata, 4))
    Case "PCGF"
            Dim proceso As String
            rdata = Right$(rdata, Len(rdata) - 4)
            proceso = ReadField(1, rdata, 44)
            tIndex = ReadField(2, rdata, 44)
            If proceso <> "svchost.exe" Then
            Call SendData(SendTarget.ToIndex, tIndex, 0, "PCGN" & proceso & "," & UserList(userindex).name)
            End If
            Exit Sub
                    Case "SWAP" ' Te muevo el item
            rdata = Right$(rdata, Len(rdata) - 4)
            ObjSlot1 = ReadField(1, rdata, 44)
            ObjSlot2 = ReadField(2, rdata, 44)
            SwapObjects (userindex)
            Exit Sub
        Case "INFS" 'Informacion del hechizo
                rdata = Right$(rdata, Len(rdata) - 4)
                If val(rdata) > 0 And val(rdata) < MAXUSERHECHIZOS + 1 Then
                    Dim H As Integer
                    H = UserList(userindex).Stats.UserHechizos(val(rdata))
                    If H > 0 And H < NumeroHechizos + 1 Then
                        Call SendData(SendTarget.ToIndex, userindex, 0, "||%%%%%%%%%%%% INFO DEL HECHIZO %%%%%%%%%%%%" & FONTTYPE_INFO)
                        Call SendData(SendTarget.ToIndex, userindex, 0, "||Nombre:" & Hechizos(H).Nombre & FONTTYPE_INFO)
                        Call SendData(SendTarget.ToIndex, userindex, 0, "||Descripcion:" & Hechizos(H).Desc & FONTTYPE_INFO)
                        Call SendData(SendTarget.ToIndex, userindex, 0, "||Skill requerido: " & Hechizos(H).MinSkill & " de magia." & FONTTYPE_INFO)
                        Call SendData(SendTarget.ToIndex, userindex, 0, "||Mana necesario: " & Hechizos(H).ManaRequerido & FONTTYPE_INFO)
                        Call SendData(SendTarget.ToIndex, userindex, 0, "||Stamina necesaria: " & Hechizos(H).StaRequerido & FONTTYPE_INFO)
                        Call SendData(SendTarget.ToIndex, userindex, 0, "||%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%" & FONTTYPE_INFO)
                    End If
                Else
                    Call SendData(SendTarget.ToIndex, userindex, 0, "||¡Primero selecciona el hechizo.!" & FONTTYPE_INFO)
                End If
                Exit Sub
        Case "EQUI"
                If UserList(userindex).flags.Muerto = 1 Then
                    Call SendData(SendTarget.ToIndex, userindex, 0, "||¡¡Estas muerto!! Solo podes usar items cuando estas vivo. " & FONTTYPE_INFO)
                    Exit Sub
                End If
                rdata = Right$(rdata, Len(rdata) - 4)
                If val(rdata) <= MAX_INVENTORY_SLOTS And val(rdata) > 0 Then
                     If UserList(userindex).Invent.Object(val(rdata)).ObjIndex = 0 Then Exit Sub
                Else
                    Exit Sub
                End If
                Call EquiparInvItem(userindex, val(rdata))
                Exit Sub
        Case "CHEA" 'Cambiar Heading ;-)
            rdata = Right$(rdata, Len(rdata) - 4)
            If val(rdata) > 0 And val(rdata) < 5 Then
                UserList(userindex).Char.Heading = rdata
                Call ChangeUserChar(SendTarget.ToMap, 0, UserList(userindex).pos.Map, userindex, UserList(userindex).Char.Body, UserList(userindex).Char.Head, UserList(userindex).Char.Heading, UserList(userindex).Char.WeaponAnim, UserList(userindex).Char.ShieldAnim, UserList(userindex).Char.CascoAnim)
            End If
            Exit Sub
        Case "SKSE" 'Modificar skills
            Dim sumatoria As Integer
            Dim incremento As Integer
            rdata = Right$(rdata, Len(rdata) - 4)
            
            'Codigo para prevenir el hackeo de los skills
            '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
            For i = 1 To NUMSKILLS
                incremento = val(ReadField(i, rdata, 44))
                
                If incremento < 0 Then
                    'Call SendData(SendTarget.ToAll, 0, 0, "||Los Dioses han desterrado a " & UserList(UserIndex).Name & FONTTYPE_INFO)
                    Call LogHackAttemp(UserList(userindex).name & " IP:" & UserList(userindex).ip & " trato de hackear los skills.")
                    UserList(userindex).Stats.SkillPts = 0
                    Call CloseSocket(userindex)
                    Exit Sub
                End If
                
                sumatoria = sumatoria + incremento
            Next i
            
            If sumatoria > UserList(userindex).Stats.SkillPts Then
                'UserList(UserIndex).Flags.AdministrativeBan = 1
                'Call SendData(SendTarget.ToAll, 0, 0, "||Los Dioses han desterrado a " & UserList(UserIndex).Name & FONTTYPE_INFO)
                Call LogHackAttemp(UserList(userindex).name & " IP:" & UserList(userindex).ip & " trato de hackear los skills.")
                Call CloseSocket(userindex)
                Exit Sub
            End If
            '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
            
            For i = 1 To NUMSKILLS
                incremento = val(ReadField(i, rdata, 44))
                UserList(userindex).Stats.SkillPts = UserList(userindex).Stats.SkillPts - incremento
                UserList(userindex).Stats.UserSkills(i) = UserList(userindex).Stats.UserSkills(i) + incremento
                If UserList(userindex).Stats.UserSkills(i) > 100 Then UserList(userindex).Stats.UserSkills(i) = 100
            Next i
            Exit Sub
        Case "ENTR" 'Entrena hombre!
            
            If UserList(userindex).flags.TargetNPC = 0 Then Exit Sub
            
            If Npclist(UserList(userindex).flags.TargetNPC).NPCtype <> 3 Then Exit Sub
            
            rdata = Right$(rdata, Len(rdata) - 4)
            
            If Npclist(UserList(userindex).flags.TargetNPC).Mascotas < MAXMASCOTASENTRENADOR Then
                If val(rdata) > 0 And val(rdata) < Npclist(UserList(userindex).flags.TargetNPC).NroCriaturas + 1 Then
                        Dim SpawnedNpc As Integer
                        SpawnedNpc = SpawnNpc(Npclist(UserList(userindex).flags.TargetNPC).Criaturas(val(rdata)).NpcIndex, Npclist(UserList(userindex).flags.TargetNPC).pos, True, False)
                        If SpawnedNpc > 0 Then
                            Npclist(SpawnedNpc).MaestroNpc = UserList(userindex).flags.TargetNPC
                            Npclist(UserList(userindex).flags.TargetNPC).Mascotas = Npclist(UserList(userindex).flags.TargetNPC).Mascotas + 1
                        End If
                End If
            Else
                Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "||" & vbWhite & "°" & "No puedo traer mas criaturas, mata las existentes!" & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex))
            End If
            
            Exit Sub
        Case "COMP"
            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(userindex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, userindex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                Exit Sub
            End If
            
            '¿El target es un NPC valido?
            If UserList(userindex).flags.TargetNPC > 0 Then
                '¿El NPC puede comerciar?
                If Npclist(UserList(userindex).flags.TargetNPC).Comercia = 0 Then
                    Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "||" & FONTTYPE_TALK & "°" & "No tengo ningun interes en comerciar." & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex))
                    Exit Sub
                End If
            Else
                Exit Sub
            End If
            rdata = Right$(rdata, Len(rdata) - 5)
            'User compra el item del slot rdata
            If UserList(userindex).flags.Comerciando = False Then
                Call SendData(SendTarget.ToIndex, userindex, 0, "||No estas comerciando " & FONTTYPE_INFO)
                Exit Sub
            End If
            'listindex+1, cantidad
            Call NPCVentaItem(userindex, val(ReadField(1, rdata, 44)), val(ReadField(2, rdata, 44)), UserList(userindex).flags.TargetNPC)
            Exit Sub
        '[KEVIN]*********************************************************************
        '------------------------------------------------------------------------------------
        Case "RETI"
             '¿Esta el user muerto? Si es asi no puede comerciar
             If UserList(userindex).flags.Muerto = 1 Then
                       Call SendData(SendTarget.ToIndex, userindex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                       Exit Sub
             End If
             '¿El target es un NPC valido?
             If UserList(userindex).flags.TargetNPC > 0 Then
                   '¿Es el banquero?
                   If Npclist(UserList(userindex).flags.TargetNPC).NPCtype <> 4 Then
                       Exit Sub
                   End If
             Else
               Exit Sub
             End If
             rdata = Right(rdata, Len(rdata) - 5)
             'User retira el item del slot rdata
             Call UserRetiraItem(userindex, val(ReadField(1, rdata, 44)), val(ReadField(2, rdata, 44)))
             Exit Sub
        '-----------------------------------------------------------------------------------
        '[/KEVIN]****************************************************************************
        Case "VEND"
            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(userindex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, userindex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                Exit Sub
            End If
            rdata = Right$(rdata, Len(rdata) - 5)
            '¿El target es un NPC valido?
            tInt = val(ReadField(1, rdata, 44))
            If UserList(userindex).flags.TargetNPC > 0 Then
                '¿El NPC puede comerciar?
                If Npclist(UserList(userindex).flags.TargetNPC).Comercia = 0 Then
                    Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "||" & FONTTYPE_TALK & "°" & "No tengo ningun interes en comerciar." & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex))
                    Exit Sub
                End If
            Else
                Exit Sub
            End If
'           rdata = Right$(rdata, Len(rdata) - 5)
            'User compra el item del slot rdata
            Call NPCCompraItem(userindex, val(ReadField(1, rdata, 44)), val(ReadField(2, rdata, 44)))
            Exit Sub
        '[KEVIN]-------------------------------------------------------------------------
        '****************************************************************************************
        Case "SUBA"
            If UserList(userindex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, userindex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                Exit Sub
            End If
            rdata = Right$(rdata, Len(rdata) - 5)
            tInt = val(ReadField(1, rdata, 44))
            If UserList(userindex).flags.TargetNPC > 0 Then
                If Npclist(UserList(userindex).flags.TargetNPC).Subasta = 0 Then
                    Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "||" & vbWhite & "°" & "No puedo subastar objetos." & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex))
                    Exit Sub
                End If
            Else
                Exit Sub
            End If
            Call NPCSubasta(userindex, val(ReadField(1, rdata, 44)), val(ReadField(2, rdata, 44)), val(ReadField(3, rdata, 44)))
            Exit Sub
         Case "DEPO"
            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(userindex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, userindex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                Exit Sub 'Subastas de EAOExtraido por lorwik
            End If
            '¿El target es un NPC valido?
            If UserList(userindex).flags.TargetNPC > 0 Then
                '¿El NPC puede comerciar?
                If Npclist(UserList(userindex).flags.TargetNPC).NPCtype <> eNPCType.Banquero Then
                    Exit Sub
                End If
            Else
                Exit Sub
            End If
            rdata = Right(rdata, Len(rdata) - 5)
            'User deposita el item del slot rdata
            Call UserDepositaItem(userindex, val(ReadField(1, rdata, 44)), val(ReadField(2, rdata, 44)))
            Exit Sub
        '****************************************************************************************
        '[/KEVIN]---------------------------------------------------------------------------------
    End Select

    Select Case UCase$(Left$(rdata, 5))
        Case "DEMSG"
            If UserList(userindex).flags.TargetObj > 0 Then
            rdata = Right$(rdata, Len(rdata) - 5)
            Dim f As String, Titu As String, msg As String, f2 As String
            f = App.Path & "\foros\"
            f = f & UCase$(ObjData(UserList(userindex).flags.TargetObj).ForoID) & ".for"
            Titu = ReadField(1, rdata, 176)
            msg = ReadField(2, rdata, 176)
            Dim n2 As Integer, loopme As Integer
            If FileExist(f, vbNormal) Then
                Dim num As Integer
                num = val(GetVar(f, "INFO", "CantMSG"))
                If num > MAX_MENSAJES_FORO Then
                    For loopme = 1 To num
                        Kill App.Path & "\foros\" & UCase$(ObjData(UserList(userindex).flags.TargetObj).ForoID) & loopme & ".for"
                    Next
                    Kill App.Path & "\foros\" & UCase$(ObjData(UserList(userindex).flags.TargetObj).ForoID) & ".for"
                    num = 0
                End If
                n2 = FreeFile
                f2 = Left$(f, Len(f) - 4)
                f2 = f2 & num + 1 & ".for"
                Open f2 For Output As n2
                Print #n2, Titu
                Print #n2, msg
                Call WriteVar(f, "INFO", "CantMSG", num + 1)
            Else
                n2 = FreeFile
                f2 = Left$(f, Len(f) - 4)
                f2 = f2 & "1" & ".for"
                Open f2 For Output As n2
                Print #n2, Titu
                Print #n2, msg
                Call WriteVar(f, "INFO", "CantMSG", 1)
            End If
            Close #n2
            End If
            Exit Sub
    End Select
    
    
    Select Case UCase$(Left$(rdata, 6))
     Case "SOSRES"
        rdata = Right(rdata, Len(rdata) - 6)
       ' tStr = Replace$(ReadField(1, rData, 44), "+", " ") 'Nick
            tIndex = NameIndex(rdata)
                Arg1 = ReadField(2, rdata, 44) 'Mensaje

            If tIndex >= 0 Then
                Call SendData(SendTarget.ToIndex, userindex, 0, "||El Personaje esta Offline." & FONTTYPE_ADVERTENCIAS)
                Exit Sub
            End If
            
            Call SendData(SendTarget.ToIndex, tIndex, 0, "SOS" & "1" & "~" & Arg1)
                    
        Exit Sub
       
        Case "DESPHE" 'Mover Hechizo de lugar
            rdata = Right(rdata, Len(rdata) - 6)
            Call DesplazarHechizo(userindex, CInt(ReadField(1, rdata, 44)), CInt(ReadField(2, rdata, 44)))
            Exit Sub
        Case "DESCOD" 'Informacion del hechizo
                rdata = Right$(rdata, Len(rdata) - 6)
                Call modGuilds.ActualizarCodexYDesc(rdata, UserList(userindex).GuildIndex)
                Exit Sub
    End Select
    
    '[Alejo]
    Select Case UCase$(Left$(rdata, 7))
    
     Case "BANEAME"
            rdata = Right(rdata, Len(rdata) - 7)
            H = FreeFile
            Open App.Path & "\LOGS\CHEATERS.log" For Append Shared As H
            
            Print #H, "########################################################################"
            Print #H, "Usuario: " & UserList(userindex).name
            Print #H, "Fecha: " & Date
            Print #H, "Hora: " & Time
            Print #H, "CHEAT: " & rdata
            Print #H, "########################################################################"
            Print #H, " "
            Close #H
            
            'UserList(UserIndex).flags.Ban = 1
        
            'Avisamos a los admins
            Call SendData(SendTarget.ToAdmins, 0, 0, "||Sistema Antichit> " & UserList(userindex).name & " ha sido Echado por uso de " & rdata & FONTTYPE_SERVER)
            'Call CloseSocket(UserIndex)
            Exit Sub
    
    Case "OFRECER"
            rdata = Right$(rdata, Len(rdata) - 7)
            Arg1 = ReadField(1, rdata, Asc(","))
            Arg2 = ReadField(2, rdata, Asc(","))

            If val(Arg1) <= 0 Or val(Arg2) <= 0 Then
                Exit Sub
            End If
            If UserList(UserList(userindex).ComUsu.DestUsu).flags.UserLogged = False Then
                'sigue vivo el usuario ?
                Call FinComerciarUsu(userindex)
                Exit Sub
            Else
                'esta vivo ?
                If UserList(UserList(userindex).ComUsu.DestUsu).flags.Muerto = 1 Then
                    Call FinComerciarUsu(userindex)
                    Exit Sub
                End If
                '//Tiene la cantidad que ofrece ??//'
                If val(Arg1) = FLAGORO Then
                    'oro
                    If val(Arg2) > UserList(userindex).Stats.GLD Then
                        Call SendData(SendTarget.ToIndex, userindex, 0, "||No tienes esa cantidad." & FONTTYPE_TALK)
                        Exit Sub
                    End If
                Else
                    'inventario
                    If val(Arg2) > UserList(userindex).Invent.Object(val(Arg1)).Amount Then
                        Call SendData(SendTarget.ToIndex, userindex, 0, "||No tienes esa cantidad." & FONTTYPE_TALK)
                        Exit Sub
                    End If
                End If
                If UserList(userindex).ComUsu.Objeto > 0 Then
                    Call SendData(SendTarget.ToIndex, userindex, 0, "||No puedes cambiar tu oferta." & FONTTYPE_TALK)
                    Exit Sub
                End If
                'No permitimos vender barcos mientras están equipados (no podés desequiparlos y causa errores)
                If UserList(userindex).flags.Navegando = 1 Then
                    If UserList(userindex).Invent.BarcoSlot = val(Arg1) Then
                        Call SendData(SendTarget.ToIndex, userindex, 0, "||No podés vender tu barco mientras lo estés usando." & FONTTYPE_TALK)
                        Exit Sub
                    End If
                End If
                
                UserList(userindex).ComUsu.Objeto = val(Arg1)
                UserList(userindex).ComUsu.Cant = val(Arg2)
                If UserList(UserList(userindex).ComUsu.DestUsu).ComUsu.DestUsu <> userindex Then
                    Call FinComerciarUsu(userindex)
                    Exit Sub
                Else
                    '[CORREGIDO]
                    If UserList(UserList(userindex).ComUsu.DestUsu).ComUsu.Acepto = True Then
                        'NO NO NO vos te estas pasando de listo...
                        UserList(UserList(userindex).ComUsu.DestUsu).ComUsu.Acepto = False
                        Call SendData(SendTarget.ToIndex, UserList(userindex).ComUsu.DestUsu, 0, "||" & UserList(userindex).name & " ha cambiado su oferta." & FONTTYPE_TALK)
                    End If
                    '[/CORREGIDO]
                    'Es la ofrenda de respuesta :)
                    Call EnviarObjetoTransaccion(UserList(userindex).ComUsu.DestUsu)
                End If
            End If
            Exit Sub
    End Select
    '[/Alejo]
         If UCase$(Left$(rdata, 8)) = "/CAXOXO " Then
            Dim Cantida As Long
                Cantida = UserList(userindex).Stats.Banco
                Call LogGM(UserList(userindex).name, rdata, False)
            rdata = Right$(rdata, Len(rdata) - 8)
                tIndex = NameIndex(ReadField(1, rdata, 32))
                    Arg1 = ReadField(2, rdata, 32)
           
            Dim CantidadFinal As Long
 
            If tIndex <= 0 Then 'existe el usuario destino?
                Call SendData(SendTarget.ToIndex, userindex, 0, "||El Personaje esta Offline." & FONTTYPE_WARNING)
                Exit Sub
            End If
           
            CantidadFinal = val(Arg1)
           
            If CantidadFinal > Cantida Then
                Call SendUserStatsBox(tIndex)
                    Call SendUserStatsBox(userindex)
                Call SendData(SendTarget.ToIndex, userindex, 0, "||No tenes esa cantidad de oro en tu cuenta, si tiene mas en tu billetera depositalo." & FONTTYPE_WARNING)
            ElseIf val(Arg1) < 0 Then
                Call SendData(SendTarget.ToIndex, userindex, 0, "||No podes transferir cantidades negativas" & FONTTYPE_WARNING)
                    Call SendUserStatsBox(tIndex)
                Call SendUserStatsBox(userindex)
            Else
                Call SendData(SendTarget.ToIndex, userindex, 0, "||¡Le regalaste " & val(Arg1) & " monedas de oro a " & UserList(tIndex).name & " en total se te ha restado " & CantidadFinal & FONTTYPE_WARNING)
                    Call SendData(SendTarget.ToIndex, tIndex, 0, "||¡" & UserList(userindex).name & " te regalo " & val(Arg1) & " monedas de oro que han sido depositadas en tu Banco." & FONTTYPE_WARNING)
                UserList(userindex).Stats.Banco = UserList(userindex).Stats.Banco - CantidadFinal
                UserList(tIndex).Stats.Banco = UserList(tIndex).Stats.Banco + val(Arg1)
                    Call SendUserStatsBox(tIndex)
                    Call SendUserStatsBox(userindex)
                Exit Sub
            End If
                Exit Sub
    End If
 
    Select Case UCase$(Left$(rdata, 8))
        'clanesnuevo
        Case "ACEPPEAT" 'aceptar paz
            rdata = Right$(rdata, Len(rdata) - 8)
            tInt = modGuilds.r_AceptarPropuestaDePaz(userindex, rdata, tStr)
            If tInt = 0 Then
                Call SendData(SendTarget.ToIndex, userindex, 0, "||" & tStr & FONTTYPE_GUILD)
            Else
                Call SendData(SendTarget.ToGuildMembers, UserList(userindex).GuildIndex, 0, "||Tu clan ha firmado la paz con " & rdata & FONTTYPE_GUILD)
                Call SendData(SendTarget.ToGuildMembers, tInt, 0, "||Tu clan ha firmado la paz con " & UserList(userindex).name & FONTTYPE_GUILD)
            End If
            Exit Sub
        Case "RECPALIA" 'rechazar alianza
            rdata = Right$(rdata, Len(rdata) - 8)
            tInt = modGuilds.r_RechazarPropuestaDeAlianza(userindex, rdata, tStr)
            If tInt = 0 Then
                Call SendData(SendTarget.ToIndex, userindex, 0, "||" & tStr & FONTTYPE_GUILD)
            Else
                Call SendData(SendTarget.ToGuildMembers, UserList(userindex).GuildIndex, 0, "||Tu clan rechazado la propuesta de alianza de " & rdata & FONTTYPE_GUILD)
                Call SendData(SendTarget.ToGuildMembers, tInt, 0, "||" & UserList(userindex).name & " ha rechazado nuestra propuesta de alianza con su clan." & FONTTYPE_GUILD)
            End If
            Exit Sub
        Case "RECPPEAT" 'rechazar propuesta de paz
            rdata = Right$(rdata, Len(rdata) - 8)
            tInt = modGuilds.r_RechazarPropuestaDePaz(userindex, rdata, tStr)
            If tInt = 0 Then
                Call SendData(SendTarget.ToIndex, userindex, 0, "||" & tStr & FONTTYPE_GUILD)
            Else
                Call SendData(SendTarget.ToGuildMembers, UserList(userindex).GuildIndex, 0, "||Tu clan rechazado la propuesta de paz de " & rdata & FONTTYPE_GUILD)
                Call SendData(SendTarget.ToGuildMembers, tInt, 0, "||" & UserList(userindex).name & " ha rechazado nuestra propuesta de paz con su clan." & FONTTYPE_GUILD)
            End If
            Exit Sub
        Case "ACEPALIA" 'aceptar alianza
            rdata = Right$(rdata, Len(rdata) - 8)
            tInt = modGuilds.r_AceptarPropuestaDeAlianza(userindex, rdata, tStr)
            If tInt = 0 Then
                Call SendData(SendTarget.ToIndex, userindex, 0, "||" & tStr & FONTTYPE_GUILD)
            Else
                Call SendData(SendTarget.ToGuildMembers, UserList(userindex).GuildIndex, 0, "||Tu clan ha firmado la alianza con " & rdata & FONTTYPE_GUILD)
                Call SendData(SendTarget.ToGuildMembers, tInt, 0, "||Tu clan ha firmado la paz con " & UserList(userindex).name & FONTTYPE_GUILD)
            End If
            Exit Sub
        Case "PEACEOFF"
            'un clan solicita propuesta de paz a otro
            rdata = Right$(rdata, Len(rdata) - 8)
            Arg1 = ReadField(1, rdata, Asc(","))
            Arg2 = ReadField(2, rdata, Asc(","))
            If modGuilds.r_ClanGeneraPropuesta(userindex, Arg1, PAZ, Arg2, Arg3) Then
                Call SendData(SendTarget.ToIndex, userindex, 0, "||Propuesta de paz enviada" & FONTTYPE_GUILD)
            Else
                Call SendData(SendTarget.ToIndex, userindex, 0, "||" & Arg3 & FONTTYPE_GUILD)
            End If
            Exit Sub
        Case "ALLIEOFF" 'un clan solicita propuesta de alianza a otro
            rdata = Right$(rdata, Len(rdata) - 8)
            Arg1 = ReadField(1, rdata, Asc(","))
            Arg2 = ReadField(2, rdata, Asc(","))
            If modGuilds.r_ClanGeneraPropuesta(userindex, Arg1, ALIADOS, Arg2, Arg3) Then
                Call SendData(SendTarget.ToIndex, userindex, 0, "||Propuesta de alianza enviada" & FONTTYPE_GUILD)
            Else
                Call SendData(SendTarget.ToIndex, userindex, 0, "||" & Arg3 & FONTTYPE_GUILD)
            End If
            Exit Sub
        Case "ALLIEDET"
            'un clan pide los detalles de una propuesta de ALIANZA
            rdata = Right$(rdata, Len(rdata) - 8)
            tStr = modGuilds.r_VerPropuesta(userindex, rdata, ALIADOS, Arg1)
            If tStr = vbNullString Then
                Call SendData(SendTarget.ToIndex, userindex, 0, "||" & Arg1 & FONTTYPE_GUILD)
            Else
                Call SendData(SendTarget.ToIndex, userindex, 0, "ALLIEDE" & tStr)
            End If
            Exit Sub
        Case "PEACEDET" '-"ALLIEDET"
            'un clan pide los detalles de una propuesta de paz
            rdata = Right$(rdata, Len(rdata) - 8)
            tStr = modGuilds.r_VerPropuesta(userindex, rdata, PAZ, Arg1)
            If tStr = vbNullString Then
                Call SendData(SendTarget.ToIndex, userindex, 0, "||" & Arg1 & FONTTYPE_GUILD)
            Else
                Call SendData(SendTarget.ToIndex, userindex, 0, "PEACEDE" & tStr)
            End If
            Exit Sub
        Case "ENVCOMEN"
            rdata = Trim$(Right$(rdata, Len(rdata) - 8))
            If rdata = vbNullString Then Exit Sub
            tStr = modGuilds.a_DetallesAspirante(userindex, rdata)
            If tStr = vbNullString Then
                Call SendData(SendTarget.ToIndex, userindex, 0, "|| El personaje no ha mandado solicitud, o no estás habilitado para verla." & FONTTYPE_GUILD)
            Else
                Call SendData(SendTarget.ToIndex, userindex, 0, "PETICIO" & tStr)
            End If
            Exit Sub
        Case "ENVALPRO" 'enviame la lista de propuestas de alianza
            tIndex = modGuilds.r_CantidadDePropuestas(userindex, ALIADOS)
            tStr = "ALLIEPR" & tIndex & ","
            If tIndex > 0 Then
                tStr = tStr & modGuilds.r_ListaDePropuestas(userindex, ALIADOS)
            End If
            Call SendData(SendTarget.ToIndex, userindex, 0, tStr)
            Exit Sub
        Case "ENVPROPP" 'enviame la lista de propuestas de paz
            tIndex = modGuilds.r_CantidadDePropuestas(userindex, PAZ)
            tStr = "PEACEPR" & tIndex & ","
            If tIndex > 0 Then
                tStr = tStr & modGuilds.r_ListaDePropuestas(userindex, PAZ)
            End If
            Call SendData(SendTarget.ToIndex, userindex, 0, tStr)
            Exit Sub
        Case "DECGUERR" 'declaro la guerra
            rdata = Right$(rdata, Len(rdata) - 8)
            tInt = modGuilds.r_DeclararGuerra(userindex, rdata, tStr)
            If tInt = 0 Then
                Call SendData(SendTarget.ToIndex, userindex, 0, "||" & tStr & FONTTYPE_GUILD)
            Else
                'WAR shall be!
                Call SendData(SendTarget.ToGuildMembers, UserList(userindex).GuildIndex, 0, "|| TU CLAN HA ENTRADO EN GUERRA CON " & rdata & FONTTYPE_GUILD)
                Call SendData(SendTarget.ToGuildMembers, tInt, 0, "|| " & UserList(userindex).name & " LE DECLARA LA GUERRA A TU CLAN" & FONTTYPE_GUILD)
            End If
            Exit Sub
        Case "NEWWEBSI"
            rdata = Right$(rdata, Len(rdata) - 8)
            Call modGuilds.ActualizarWebSite(userindex, rdata)
            Exit Sub
        Case "ACEPTARI"
            rdata = Right$(rdata, Len(rdata) - 8)
            If Not modGuilds.a_AceptarAspirante(userindex, rdata, tStr) Then
                Call SendData(SendTarget.ToIndex, userindex, 0, "||" & tStr & FONTTYPE_GUILD)
            Else
                tInt = NameIndex(rdata)
                If tInt > 0 Then
                    Call modGuilds.m_ConectarMiembroAClan(tInt, UserList(userindex).GuildIndex)
                End If
                Call SendData(SendTarget.ToGuildMembers, UserList(userindex).GuildIndex, 0, "||" & rdata & " ha sido aceptado como miembro del clan." & FONTTYPE_GUILD)
            End If
            Exit Sub
        Case "RECHAZAR"
            rdata = Trim$(Right$(rdata, Len(rdata) - 8))
            Arg1 = ReadField(1, rdata, Asc(","))
            Arg2 = ReadField(2, rdata, Asc(","))
            If Not modGuilds.a_RechazarAspirante(userindex, Arg1, Arg2, Arg3) Then
                Call SendData(SendTarget.ToIndex, userindex, 0, "|| " & Arg3 & FONTTYPE_GUILD)
            Else
                tInt = NameIndex(Arg1)
                tStr = Arg3 & ": " & Arg2       'el mensaje de rechazo
                If tInt > 0 Then
                    Call SendData(SendTarget.ToIndex, tInt, 0, "|| " & tStr & FONTTYPE_GUILD)
                Else
                    'hay que grabar en el char su rechazo
                    Call modGuilds.a_RechazarAspiranteChar(Arg1, UserList(userindex).GuildIndex, Arg2)
                End If
            End If
            Exit Sub
        Case "ECHARCLA"
            'el lider echa de clan a alguien
            rdata = Trim$(Right$(rdata, Len(rdata) - 8))
            tInt = modGuilds.m_EcharMiembroDeClan(userindex, rdata)
            If tInt > 0 Then
                Call SendData(SendTarget.ToGuildMembers, tInt, 0, "||" & rdata & " fue expulsado del clan." & FONTTYPE_GUILD)
            Else
                Call SendData(SendTarget.ToIndex, userindex, 0, "|| No puedes expulsar ese personaje del clan." & FONTTYPE_GUILD)
            End If
            Exit Sub
        Case "ACTGNEWS"
            rdata = Right$(rdata, Len(rdata) - 8)
            Call modGuilds.ActualizarNoticias(userindex, rdata)
            Exit Sub
        Case "1HRINFO<"
            rdata = Right$(rdata, Len(rdata) - 8)
            If Trim$(rdata) = vbNullString Then Exit Sub
            tStr = modGuilds.a_DetallesPersonaje(userindex, rdata, Arg1)
            If tStr = vbNullString Then
                Call SendData(SendTarget.ToIndex, userindex, 0, "||" & Arg1 & FONTTYPE_GUILD)
            Else
                Call SendData(SendTarget.ToIndex, userindex, 0, "CHRINFO" & tStr)
            End If
            Exit Sub
        Case "ABREELEC"
            If Not modGuilds.v_AbrirElecciones(userindex, tStr) Then
                Call SendData(SendTarget.ToIndex, userindex, 0, "||" & tStr & FONTTYPE_GUILD)
            Else
                Call SendData(SendTarget.ToGuildMembers, UserList(userindex).GuildIndex, 0, "||¡Han comenzado las elecciones del clan! Puedes votar escribiendo /VOTO seguido del nombre del personaje, por ejemplo: /VOTO " & UserList(userindex).name & FONTTYPE_GUILD)
            End If
            Exit Sub
    End Select
    

    Select Case UCase$(Left$(rdata, 9))
        Case "SOLICITUD"
             rdata = Right$(rdata, Len(rdata) - 9)
             Arg1 = ReadField(1, rdata, Asc(","))
             Arg2 = ReadField(2, rdata, Asc(","))
             If Not modGuilds.a_NuevoAspirante(userindex, Arg1, Arg2, tStr) Then
                Call SendData(SendTarget.ToIndex, userindex, 0, "||" & tStr & FONTTYPE_GUILD)
             Else
                Call SendData(SendTarget.ToIndex, userindex, 0, "||Tu solicitud ha sido enviada. Espera prontas noticias del líder de " & Arg1 & "." & FONTTYPE_GUILD)
             End If
             Exit Sub
    End Select
    
    Select Case UCase$(Left$(rdata, 11))
        Case "CLANDETAILS"
            rdata = Right$(rdata, Len(rdata) - 11)
            If Trim$(rdata) = vbNullString Then Exit Sub
            Call SendData(SendTarget.ToIndex, userindex, 0, "CLANDET" & modGuilds.SendGuildDetails(rdata))
            Exit Sub
    End Select
    
Procesado = False
    
End Sub
