Attribute VB_Name = "Torneos"
Option Explicit
' Codigo: Torneos Automaticos 100%
' Autor: Joan Calderón - SaturoS.
Public Torneo_Activo As Boolean
Public Torneo_Esperando As Boolean
Private Torneo_Rondas As Integer
Private Torneo_Luchadores() As Integer
 
Private Const mapatorneo As Integer = 125
' esquinas superior isquierda del ring
Private Const esquina1x As Integer = 41
Private Const esquina1y As Integer = 50
' esquina inferior derecha del ring
Private Const esquina2x As Integer = 60
Private Const esquina2y As Integer = 50
' Donde esperan los tios
Private Const esperax As Integer = 52
Private Const esperay As Integer = 27
' Mapa desconecta
Private Const mapa_fuera As Integer = 1
Private Const fueraesperay As Integer = 50
Private Const fueraesperax As Integer = 50
 ' estas son las pocisiones de las 2 esquinas de la zona de espera, en su mapa tienen que tener en la misma posicion las 2 esquinas.
Private Const X1 As Integer = 36
Private Const X2 As Integer = 65
Private Const Y1 As Integer = 24
Private Const Y2 As Integer = 30

Sub Torneoauto_Cancela()
On Error GoTo errorh:
    If (Not Torneo_Activo And Not Torneo_Esperando) Then Exit Sub
    Torneo_Activo = False
    Torneo_Esperando = False
    Call SendData(SendTarget.toall, 0, 0, "||Torneo: Torneo Automatico cancelado por falta de participantes." & FONTTYPE_GUILD)
    Dim i As Integer
     For i = LBound(Torneo_Luchadores) To UBound(Torneo_Luchadores)
                If (Torneo_Luchadores(i) <> -1) Then
                 Dim NuevaPos As WorldPos
                  Dim FuturePos As WorldPos
                    FuturePos.Map = mapa_fuera
                    FuturePos.x = fueraesperax: FuturePos.Y = fueraesperay
                    Call ClosestLegalPos(FuturePos, NuevaPos)
                    If NuevaPos.x <> 0 And NuevaPos.Y <> 0 Then Call WarpUserChar(Torneo_Luchadores(i), NuevaPos.Map, NuevaPos.x, NuevaPos.Y, True)
                      UserList(Torneo_Luchadores(i)).flags.automatico = False
                End If
        Next i
errorh:
End Sub
Sub Rondas_Cancela()
On Error GoTo errorh
    If (Not Torneo_Activo And Not Torneo_Esperando) Then Exit Sub
    Torneo_Activo = False
    Torneo_Esperando = False
    Call SendData(SendTarget.toall, 0, 0, "||Torneo: Torneo Automatico cancelado por Game Master" & FONTTYPE_GUILD)
    Dim i As Integer
    For i = LBound(Torneo_Luchadores) To UBound(Torneo_Luchadores)
                If (Torneo_Luchadores(i) <> -1) Then
                        Dim NuevaPos As WorldPos
                  Dim FuturePos As WorldPos
                    FuturePos.Map = mapa_fuera
                    FuturePos.x = fueraesperax: FuturePos.Y = fueraesperay
                    Call ClosestLegalPos(FuturePos, NuevaPos)
                    If NuevaPos.x <> 0 And NuevaPos.Y <> 0 Then Call WarpUserChar(Torneo_Luchadores(i), NuevaPos.Map, NuevaPos.x, NuevaPos.Y, True)
                    UserList(Torneo_Luchadores(i)).flags.automatico = False
                End If
        Next i
errorh:
End Sub
Sub Rondas_UsuarioMuere(ByVal userindex As Integer, Optional Real As Boolean = True, Optional CambioMapa As Boolean = False)
On Error GoTo rondas_usuariomuere_errorh
        Dim i As Integer, pos As Integer, j As Integer
        Dim combate As Integer, LI1 As Integer, LI2 As Integer
        Dim UI1 As Integer, UI2 As Integer
If (Not Torneo_Activo) Then
                Exit Sub
            ElseIf (Torneo_Activo And Torneo_Esperando) Then
                For i = LBound(Torneo_Luchadores) To UBound(Torneo_Luchadores)
                    If (Torneo_Luchadores(i) = userindex) Then
                        Torneo_Luchadores(i) = -1
                        Call WarpUserChar(userindex, mapa_fuera, fueraesperay, fueraesperax, True)
                         UserList(userindex).flags.automatico = False
                        Exit Sub
                    End If
                Next i
                Exit Sub
            End If
 
        For pos = LBound(Torneo_Luchadores) To UBound(Torneo_Luchadores)
                If (Torneo_Luchadores(pos) = userindex) Then Exit For
        Next pos
 
        ' si no lo ha encontrado
        If (Torneo_Luchadores(pos) <> userindex) Then Exit Sub
       
 '  Ojo con esta parte, aqui es donde verifica si el usuario esta en la posicion de espera del torneo, en estas cordenadas tienen que fijarse al crear su Mapa de torneos.
 
If UserList(userindex).pos.x >= X1 And UserList(userindex).pos.x <= X2 And UserList(userindex).pos.Y >= Y1 And UserList(userindex).pos.Y <= Y2 Then
Call SendData(SendTarget.toall, 0, 0, "||Torneo: " & UserList(userindex).name & " se fue del torneo mientras esperaba pelear.!" & FONTTYPE_TALK)
Call WarpUserChar(userindex, mapa_fuera, fueraesperax, fueraesperay, True)
UserList(userindex).flags.automatico = False
Torneo_Luchadores(pos) = -1
Exit Sub
End If
 
        combate = 1 + (pos - 1) \ 2
 
        'ponemos li1 y li2 (luchador index) de los que combatian
        LI1 = 2 * (combate - 1) + 1
        LI2 = LI1 + 1
 
        'se informa a la gente
        If (Real) Then
                Call SendData(SendTarget.toall, 0, 0, "||Torneo: " & UserList(userindex).name & " pierde el combate!" & FONTTYPE_TALK)
        Else
                Call SendData(SendTarget.toall, 0, 0, "||Torneo: " & UserList(userindex).name & " se fue del combate!" & FONTTYPE_TALK)
        End If
 
        'se le teleporta fuera si murio
        If (Real) Then
                Call WarpUserChar(userindex, mapa_fuera, fueraesperax, fueraesperay, True)
                 UserList(userindex).flags.automatico = False
        ElseIf (Not CambioMapa) Then
             
                 Call WarpUserChar(userindex, mapa_fuera, fueraesperax, fueraesperay, True)
                  UserList(userindex).flags.automatico = False
        End If
 
        'se le borra de la lista y se mueve el segundo a li1
        If (Torneo_Luchadores(LI1) = userindex) Then
                Torneo_Luchadores(LI1) = Torneo_Luchadores(LI2) 'cambiamos slot
                Torneo_Luchadores(LI2) = -1
        Else
                Torneo_Luchadores(LI2) = -1
        End If
 
    'si es la ultima ronda
    If (Torneo_Rondas = 1) Then
        Call WarpUserChar(Torneo_Luchadores(LI1), mapa_fuera, 51, 51, True)
        Call SendData(SendTarget.toall, 0, 0, "||GANADOR DEL TORNEO: " & UserList(Torneo_Luchadores(LI1)).name & FONTTYPE_GUILD)
        Call SendData(SendTarget.toall, 0, 0, "||PREMIO: 1.000.000 de oro." & FONTTYPE_GUILD)
        UserList(Torneo_Luchadores(LI1)).Stats.GLD = UserList(Torneo_Luchadores(LI1)).Stats.GLD + 1000000
         UserList(Torneo_Luchadores(LI1)).flags.automatico = False
       Call SendUserStatsBox(Torneo_Luchadores(LI1))
        Torneo_Activo = False
        Exit Sub
    Else
        'a su compañero se le teleporta dentro, condicional por seguridad
        Call WarpUserChar(Torneo_Luchadores(LI1), mapatorneo, esperax, esperay, True)
    End If
 
               
        'si es el ultimo combate de la ronda
        If (2 ^ Torneo_Rondas = 2 * combate) Then
 
                Call SendData(SendTarget.toall, 0, 0, "||Torneo: Siguiente ronda!" & FONTTYPE_GUILD)
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
 
 
 
Sub Rondas_UsuarioDesconecta(ByVal userindex As Integer)
On Error GoTo errorh
Call SendData(SendTarget.toall, 0, 0, "||Torneo: " & UserList(userindex).name & " Ha desconectado en Torneo Automatico, se le penaliza con 2kk !!" & FONTTYPE_TALK)
 If UserList(userindex).Stats.GLD >= 2000000 Then
UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD - 2000000
End If
Call SendUserStatsBox(userindex)
        Call Rondas_UsuarioMuere(userindex, False, False)
errorh:
End Sub
 
 
 
Sub Rondas_UsuarioCambiamapa(ByVal userindex As Integer)
On Error GoTo errorh
        Call Rondas_UsuarioMuere(userindex, False, True)
errorh:
End Sub
 
Sub torneos_auto(ByVal rondas As Integer)
On Error GoTo errorh
If (Torneo_Activo) Then
               
                Exit Sub
        End If
        Call SendData(SendTarget.toall, 0, 0, "||Torneo: Esta empezando un nuevo torneo 1v1 de " & val(2 ^ rondas) & " participantes!! para participar pon /PARTICIPAR - (No cae inventario)" & FONTTYPE_GUILD)
        Call SendData(SendTarget.toall, 0, 0, "TW48")
       
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
 
Sub Torneos_Inicia(ByVal userindex As Integer, ByVal rondas As Integer)
On Error GoTo errorh
        If (Torneo_Activo) Then
                Call SendData(SendTarget.toindex, userindex, 0, "||Ya hay un torneo!." & FONTTYPE_INFO)
                Exit Sub
        End If
        Call SendData(SendTarget.toall, 0, 0, "||Torneo: Esta empezando un nuevo torneo 1v1 de " & val(2 ^ rondas) & " participantes!! para participar pon /PARTICIPAR - (No cae inventario)" & FONTTYPE_GUILD)
        Call SendData(SendTarget.toall, 0, 0, "TW48")
       
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
 
 
 
Sub Torneos_Entra(ByVal userindex As Integer)
On Error GoTo errorh
        Dim i As Integer
       
        If (Not Torneo_Activo) Then
                Call SendData(SendTarget.toindex, userindex, 0, "||No hay ningun torneo!." & FONTTYPE_INFO)
                Exit Sub
        End If
       
        If (Not Torneo_Esperando) Then
                Call SendData(SendTarget.toindex, userindex, 0, "||El torneo ya ha empezado, te quedaste fuera!." & FONTTYPE_INFO)
                Exit Sub
        End If
       
        For i = LBound(Torneo_Luchadores) To UBound(Torneo_Luchadores)
                If (Torneo_Luchadores(i) = userindex) Then
                        Call SendData(SendTarget.toindex, userindex, 0, "||Ya estas dentro!" & FONTTYPE_WARNING)
                        Exit Sub
                End If
        Next i
 
        For i = LBound(Torneo_Luchadores) To UBound(Torneo_Luchadores)
        If (Torneo_Luchadores(i) = -1) Then
                Torneo_Luchadores(i) = userindex
                 Dim NuevaPos As WorldPos
                  Dim FuturePos As WorldPos
                    FuturePos.Map = mapatorneo
                    FuturePos.x = esperax: FuturePos.Y = esperay
                    Call ClosestLegalPos(FuturePos, NuevaPos)
                   
                    If NuevaPos.x <> 0 And NuevaPos.Y <> 0 Then Call WarpUserChar(Torneo_Luchadores(i), NuevaPos.Map, NuevaPos.x, NuevaPos.Y, True)
                 UserList(Torneo_Luchadores(i)).flags.automatico = True
                 
                Call SendData(SendTarget.toindex, userindex, 0, "||Estas dentro del torneo!" & FONTTYPE_INFO)
               
                Call SendData(SendTarget.toall, 0, 0, "||Torneo: Entra el participante " & UserList(userindex).name & FONTTYPE_INFO)
                If (i = UBound(Torneo_Luchadores)) Then
                Call SendData(SendTarget.toall, 0, 0, "||Torneo: Empieza el torneo!" & FONTTYPE_GUILD)
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
        Call SendData(SendTarget.toall, 0, 0, "||Torneo: Combate anulado porque un participante involucrado se desconecto" & FONTTYPE_TALK)
        If (Torneo_Rondas = 1) Then
            If (UI2 <> -1) Then
                Call SendData(SendTarget.toall, 0, 0, "||Torneo: Torneo terminado. Ganador del torneo por eliminacion: " & UserList(UI2).name & FONTTYPE_GUILD)
                UserList(UI2).flags.automatico = False
                ' dale_recompensa()
                Torneo_Activo = False
                Exit Sub
            End If
            Call SendData(SendTarget.toall, 0, 0, "||Torneo: Torneo terminado. No hay ganador porque todos se fueron :(" & FONTTYPE_GUILD)
            Exit Sub
        End If
        If (UI2 <> -1) Then _
            Call SendData(SendTarget.toall, 0, 0, "||Torneo: " & UserList(UI2).name & " pasa a la siguiente ronda!" & FONTTYPE_TALK)
   
        If (2 ^ Torneo_Rondas = 2 * combate) Then
            Call SendData(SendTarget.toall, 0, 0, "||Torneo: Siguiente ronda!" & FONTTYPE_GUILD)
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
 
    Call SendData(SendTarget.toall, 0, 0, "||Torneo: " & UserList(UI1).name & " versus " & UserList(UI2).name & ". Esquinas!! Peleen!" & FONTTYPE_GUILD)
 
    Call WarpUserChar(UI1, mapatorneo, esquina1x, esquina1y, True)
    Call WarpUserChar(UI2, mapatorneo, esquina2x, esquina2y, True)
errorh:
End Sub
