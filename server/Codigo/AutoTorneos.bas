Attribute VB_Name = "AutoTorneos"
Option Explicit
Public Torneo_Activo As Boolean
Public Torneo_Esperando As Boolean
Private Torneo_Rondas As Integer
Private Torneo_Luchadores() As Integer
'*********Mapa del Torneo**********
Private Const MapTorneo As Integer = 118
'**********Esquinas****************
Private Const ES1X As Integer = 68
Private Const ES1Y As Integer = 76
Private Const ES2X As Integer = 88
Private Const ES2Y As Integer = 88
'**********Pos de espera***********
Private Const StopX As Integer = 24
Private Const StopY As Integer = 21
'*******Despues del torneo*********
Private Const ExitTorneo As Integer = 1
Private Const ExitTorneoY As Integer = 63
Private Const ExitTorneoX As Integer = 32
' estas son las pocisiones de las 2 esquinas de la zona de espera, en su mapa tienen que tener en la misma posicion las 2 esquinas.
Private Const X1 As Integer = 17
Private Const Y1 As Integer = 17
Private Const X2 As Integer = 32
Private Const Y2 As Integer = 25
     
Sub CancelarAutoTorneo()
On Error GoTo errorh:
    If (Not Torneo_Activo And Not Torneo_Esperando) Then Exit Sub
    Torneo_Activo = False
    Torneo_Esperando = False
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("AutoTorneo> El Torneo Automatico fues suspendido por la falta de participantes.", FontTypeNames.FONTTYPE_GUILD))
       
    Dim i As Integer
        For i = LBound(Torneo_Luchadores) To UBound(Torneo_Luchadores)
                If (Torneo_Luchadores(i) <> -1) Then
                    Dim NuevaPos As WorldPos
                    Dim FuturePos As WorldPos
                    FuturePos.map = ExitTorneo
                    FuturePos.X = ExitTorneoX: FuturePos.Y = ExitTorneoY
                    Call ClosestLegalPos(FuturePos, NuevaPos)
                    UserList(i).Stats.GLD = UserList(i).Stats.GLD + 100000
                    Call WriteUpdateUserStats(i)
                    If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then Call WarpUserChar(Torneo_Luchadores(i), NuevaPos.map, NuevaPos.X, NuevaPos.Y, True)
                        UserList(Torneo_Luchadores(i)).flags.AutoTorneo = False
                End If
        Next i
errorh:
    End Sub
Sub Rondas_Cancela()
On Error GoTo errorh
        If (Not Torneo_Activo And Not Torneo_Esperando) Then Exit Sub
        Torneo_Activo = False
        Torneo_Esperando = False
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("AutoTorneo> El GameMaster suspendio el torneo automatico.", FontTypeNames.FONTTYPE_GUILD))
     
     
        Dim i As Integer
        For i = LBound(Torneo_Luchadores) To UBound(Torneo_Luchadores)
                    If (Torneo_Luchadores(i) <> -1) Then
                            Dim NuevaPos As WorldPos
                      Dim FuturePos As WorldPos
                        FuturePos.map = ExitTorneo
                        FuturePos.X = ExitTorneoX: FuturePos.Y = ExitTorneoY
                        Call ClosestLegalPos(FuturePos, NuevaPos)
                        UserList(i).Stats.GLD = UserList(i).Stats.GLD + 100000
                        Call WriteUpdateUserStats(i)
                        If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then Call WarpUserChar(Torneo_Luchadores(i), NuevaPos.map, NuevaPos.X, NuevaPos.Y, True)
                        UserList(Torneo_Luchadores(i)).flags.AutoTorneo = False
                    End If
            Next i
errorh:
End Sub
Sub Rondas_UsuarioMuere(ByVal UserIndex As Integer, Optional Real As Boolean = True, Optional CambioMapa As Boolean = False)
On Error GoTo rondas_usuariomuere_errorh
            Dim i As Integer, Pos As Integer, j As Integer
            Dim combate As Integer, LI1 As Integer, LI2 As Integer
            Dim UI1 As Integer, UI2 As Integer
            
            If (Not Torneo_Activo) Then
                Exit Sub
            ElseIf (Torneo_Activo And Torneo_Esperando) Then
                For i = LBound(Torneo_Luchadores) To UBound(Torneo_Luchadores)
                If (Torneo_Luchadores(i) = UserIndex) Then
                        Torneo_Luchadores(i) = -1
                        Call WarpUserChar(UserIndex, ExitTorneo, ExitTorneoY, ExitTorneoX, True)
                        UserList(UserIndex).flags.AutoTorneo = False
                        Exit Sub
                End If
                Next i
                Exit Sub
            End If
     
            For Pos = LBound(Torneo_Luchadores) To UBound(Torneo_Luchadores)
                    If (Torneo_Luchadores(Pos) = UserIndex) Then Exit For
            Next Pos
     
            ' si no lo ha encontrado
            If (Torneo_Luchadores(Pos) <> UserIndex) Then Exit Sub
           
            ' Ojo con esta parte, aqui es donde verifica si el usuario esta en la posicion de espera del torneo, en estas cordenadas tienen que fijarse al crear su Mapa de torneos.
     
            If UserList(UserIndex).Pos.X >= X1 And UserList(UserIndex).Pos.X <= X2 And UserList(UserIndex).Pos.Y >= Y1 And UserList(UserIndex).Pos.Y <= Y2 Then
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Torneo: " & UserList(UserIndex).Name & " se fue del torneo mientras esperaba pelear", FontTypeNames.FONTTYPE_GUILD))
                Call WarpUserChar(UserIndex, ExitTorneo, ExitTorneoX, ExitTorneoY, True)
                UserList(UserIndex).flags.AutoTorneo = False
                Torneo_Luchadores(Pos) = -1
                Exit Sub
            End If
     
            combate = 1 + (Pos - 1) \ 2
     
            'ponemos li1 y li2 (luchador index) de los que combatian
            LI1 = 2 * (combate - 1) + 1
            LI2 = LI1 + 1
     
            'se informa a la gente
            If (Real) Then
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Torneo: " & UserList(UserIndex).Name & " esta fuera de combate!", FontTypeNames.FONTTYPE_TALK))
            Else
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Torneo: " & UserList(UserIndex).Name & " sale del combate!", FontTypeNames.FONTTYPE_TALK))
            End If
     
            'se le teleporta fuera si murio
            If (Real) Then
                Call WarpUserChar(UserIndex, ExitTorneo, ExitTorneoX, ExitTorneoY, True)
                UserList(UserIndex).flags.AutoTorneo = False
            ElseIf (Not CambioMapa) Then
                Call WarpUserChar(UserIndex, ExitTorneo, ExitTorneoX, ExitTorneoY, True)
                UserList(UserIndex).flags.AutoTorneo = False
            End If
     
            'se le borra de la lista y se mueve el segundo a li1
            If (Torneo_Luchadores(LI1) = UserIndex) Then
                    Torneo_Luchadores(LI1) = Torneo_Luchadores(LI2) 'cambiamos slot
                    Torneo_Luchadores(LI2) = -1
            Else
                    Torneo_Luchadores(LI2) = -1
            End If
     
        'si es la ultima ronda
        If (Torneo_Rondas = 1) Then
            Call WarpUserChar(Torneo_Luchadores(LI1), ExitTorneo, 51, 51, True)
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("GANADOR DEL TORNEO: " & UserList(Torneo_Luchadores(LI1)).Name, FontTypeNames.FONTTYPE_TALK))
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("PREMIO: " & 100000 * UBound(Torneo_Luchadores) & " de oro.", FontTypeNames.FONTTYPE_TALK))
     
             UserList(Torneo_Luchadores(LI1)).Stats.GLD = UserList(Torneo_Luchadores(LI1)).Stats.GLD + 100000 * UBound(Torneo_Luchadores)
             UserList(Torneo_Luchadores(LI1)).flags.AutoTorneo = False
             Call WriteUpdateUserStats(Torneo_Luchadores(LI1))
             Torneo_Activo = False
             Exit Sub
        Else
             'a su compañero se le teleporta dentro, condicional por seguridad
             Call WarpUserChar(Torneo_Luchadores(LI1), MapTorneo, StopX, StopY, True)
        End If
     
            'si es el ultimo combate de la ronda
            If (2 ^ Torneo_Rondas = 2 * combate) Then
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Torneo: Siguiente ronda!", FontTypeNames.FONTTYPE_TALK))
                Call SendData(SendTarget.ToAll, UserIndex, PrepareMessagePlayWave(265, 0, 0))
                Torneo_Rondas = Torneo_Rondas - 1
     
                'antes de llamar a la proxima ronda hay q copiar a los tipos
                For i = 1 To 2 ^ Torneo_Rondas
                        UI1 = Torneo_Luchadores(2 * (i - 1) + 1)
                        UI2 = Torneo_Luchadores(2 * i)
                        If (UI1 = -1) Then UI1 = UI2
                        Torneo_Luchadores(i) = UI1
                Next i
                ReDim Preserve Torneo_Luchadores(1 To 2 ^ Torneo_Rondas) As Integer
                Call Rondas_Combate(1)
                Exit Sub
            End If
     
            'vamos al siguiente combate
            Call Rondas_Combate(combate + 1)
rondas_usuariomuere_errorh:
End Sub
     
Sub Rondas_UsuarioDesconecta(ByVal UserIndex As Integer)
    On Error GoTo errorh
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Torneo: " & UserList(UserIndex).Name & " Se te ha penalizado con 200000 monedas de oro por desconectar del torneo.", FontTypeNames.FONTTYPE_TALK))
     
    If UserList(UserIndex).Stats.GLD >= 100000 Then
        UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - 100000
    End If
    
    Call WriteUpdateUserStats(UserIndex)
    Call Rondas_UsuarioMuere(UserIndex, False, False)
errorh:
End Sub
     
Sub Rondas_UsuarioCambiamapa(ByVal UserIndex As Integer)
On Error GoTo errorh
    Call Rondas_UsuarioMuere(UserIndex, False, True)
errorh:
End Sub
Sub Auto_Torneos(ByVal rondas As Integer)
On Error GoTo errorh
    If (Torneo_Activo) Then
        Exit Sub
    End If
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Autorneo: Va a comenzar un nuevo torneo de 1 vs 1 de " & val(2 ^ rondas) & " participantes!! Para entrar escriban '/PARTICIPAR' - Precio 100k - (No cae inventario)", FontTypeNames.FONTTYPE_TALK))
    Torneo_Rondas = rondas
    Torneo_Activo = True
    Torneo_Esperando = True
     
    ReDim Torneo_Luchadores(1 To 2 ^ rondas) As Integer
    Dim i As Integer
    For i = LBound(Torneo_Luchadores) To UBound(Torneo_Luchadores)
        Torneo_Luchadores(i) = -1
    Next i
errorh:
End Sub
     
Sub Torneos_Inicia(ByVal UserIndex As Integer, ByVal rondas As Integer)
On Error GoTo errorh
    If (Torneo_Activo) Then
        Call WriteConsoleMsg(UserIndex, "Ya hay un torneo!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
    End If
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Autorneo: Va a comenzar un nuevo torneo de 1 vs 1 de " & val(2 ^ rondas) & " participantes!! Para entrar escriban '/PARTICIPAR' - Precio 100k - (No cae inventario)", FontTypeNames.FONTTYPE_TALK))
    Torneo_Rondas = rondas
    Torneo_Activo = True
    Torneo_Esperando = True
     
    ReDim Torneo_Luchadores(1 To 2 ^ rondas) As Integer
    Dim i As Integer
    For i = LBound(Torneo_Luchadores) To UBound(Torneo_Luchadores)
        Torneo_Luchadores(i) = -1
    Next i
errorh:
End Sub
    
Sub Torneos_Entra(ByVal UserIndex As Integer)
On Error GoTo errorh
            Dim i As Integer
           
            If (Not Torneo_Activo) Then
                Call WriteConsoleMsg(UserIndex, "¡No hay ningun torneo!", FontTypeNames.FONTTYPE_INFO)
     
                    Exit Sub
            End If
           
            If (Not Torneo_Esperando) Then
                Call WriteConsoleMsg(UserIndex, "¡El torneo ya ha comenzado, te quedaste fuera!.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
           
           
            For i = LBound(Torneo_Luchadores) To UBound(Torneo_Luchadores)
                If (Torneo_Luchadores(i) = UserIndex) Then
                    Call WriteConsoleMsg(UserIndex, "¡Ya estas dentro!", FontTypeNames.FONTTYPE_WARNING)
                    Exit Sub
                    End If
            Next i
     
            If Not UserList(UserIndex).Stats.GLD >= 100000 Then
                Call WriteConsoleMsg(UserIndex, "¡Necesitas 100000 monedas de oro para participar!", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
     
            For i = LBound(Torneo_Luchadores) To UBound(Torneo_Luchadores)
            If (Torneo_Luchadores(i) = -1) Then
                Torneo_Luchadores(i) = UserIndex
                Dim NuevaPos As WorldPos
                Dim FuturePos As WorldPos
                FuturePos.map = MapTorneo
                FuturePos.X = StopX: FuturePos.Y = StopY
                Call ClosestLegalPos(FuturePos, NuevaPos)
                
                UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - 100000
                Call WriteUpdateUserStats(UserIndex)
                
                If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then Call WarpUserChar(Torneo_Luchadores(i), NuevaPos.map, NuevaPos.X, NuevaPos.Y, True)
                UserList(Torneo_Luchadores(i)).flags.AutoTorneo = True
                Call WriteConsoleMsg(UserIndex, "Estas dentro del torneo!", FontTypeNames.FONTTYPE_INFO)
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("AutoTorneo> Entra el participante " & UserList(UserIndex).Name, FontTypeNames.FONTTYPE_INFO))
     
                If (i = UBound(Torneo_Luchadores)) Then
                    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("AutoTorneo> Comienza el torneo", FontTypeNames.FONTTYPE_INFO))
                    Torneo_Esperando = False
                    Call Rondas_Combate(1)
                    End If
                Exit Sub
            End If
            Next i
errorh:
End Sub
Sub Rondas_Combate(combate As Integer)
On Error GoTo errorh
    Dim UI1 As Integer, UI2 As Integer
        UI1 = Torneo_Luchadores(2 * (combate - 1) + 1)
        UI2 = Torneo_Luchadores(2 * combate)
       
        If (UI2 = -1) Then
            UI2 = Torneo_Luchadores(2 * (combate - 1) + 1)
            UI1 = Torneo_Luchadores(2 * combate)
        End If
       
        If (UI1 = -1) Then
     
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("AutoTorneo> Uno de los participantes en combate se desconecto. ¡Combate anulado!", FontTypeNames.FONTTYPE_INFO))
     
            If (Torneo_Rondas = 1) Then
                If (UI2 <> -1) Then
                        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("AutoTorneo> El torneo ha finalizado. El ganador del torneo por eliminacion es: " & UserList(UI2).Name, FontTypeNames.FONTTYPE_INFO))
     
                    UserList(UI2).flags.AutoTorneo = False
                    ' dale_recompensa()
                    Torneo_Activo = False
                    Exit Sub
                End If
                    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("AutoTorneo> El torneo ha finalizado. No hay ganador. ", FontTypeNames.FONTTYPE_INFO))
                Exit Sub
            End If
            If (UI2 <> -1) Then
                        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("AutoTorneo> " & UserList(UI2).Name & " pasa a la siguiente ronda!", FontTypeNames.FONTTYPE_INFO))
     
       End If
            If (2 ^ Torneo_Rondas = 2 * combate) Then
     
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Torneo: Siguiente ronda!", FontTypeNames.FONTTYPE_INFO))
     
                Torneo_Rondas = Torneo_Rondas - 1
                'antes de llamar a la proxima ronda hay q copiar a los tipos
                Dim i As Integer, j As Integer
                For i = 1 To 2 ^ Torneo_Rondas
                    UI1 = Torneo_Luchadores(2 * (i - 1) + 1)
                    UI2 = Torneo_Luchadores(2 * i)
                    If (UI1 = -1) Then UI1 = UI2
                    Torneo_Luchadores(i) = UI1
                Next i
                ReDim Preserve Torneo_Luchadores(1 To 2 ^ Torneo_Rondas) As Integer
                Call Rondas_Combate(1)
                Exit Sub
            End If
            Call Rondas_Combate(combate + 1)
            Exit Sub
        End If
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("AutoTorneo> " & UserList(UI1).Name & " vs " & UserList(UI2).Name & ".", FontTypeNames.FONTTYPE_INFO))
     
        Call WarpUserChar(UI1, MapTorneo, ES1X, ES1Y, True)
        Call WarpUserChar(UI2, MapTorneo, ES2X, ES2Y, True)
errorh:
End Sub
