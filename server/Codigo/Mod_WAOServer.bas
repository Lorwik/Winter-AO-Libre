Attribute VB_Name = "ModGuerras"
Option Explicit
 
Public HayGuerra As Boolean 'Temporal: Hay Guerra o No?
Public CiudadGuerra As Integer 'Temporal: En que ciudad es la Guerra?
Private TiempoGuerra As Integer 'Temporal: Tiempo Transcurrido
Private GuerrasAutomaticas As Boolean 'Temporal: Guerras Automaticas
Private PosicionNPC As WorldPos 'Temporal: Posicion del NPC
Private NPCGuerra As Integer 'Temporal: NPC Usado en Guerra
 
'Facccion Real:
Public Const NPC1 As Integer = 130 'NPC de La Faccion Real
Private Const MapaGuerra1 As Integer = 43 'Mapa de la Faccion Real
Private Const MapaGuerra1X As Byte = 50 'X del Mapa de la Faccion Real
Private Const MapaGuerra1Y As Byte = 50 'Y del Mapa de la Faccion Caos
 
'Faccion Caos:
Public Const NPC2 As Integer = 131 'NPC de La Faccion Caos
Private Const MapaGuerra2 As Integer = 70 'Mapa de la Faccion Caos
Private Const MapaGuerra2X As Byte = 50 'X del Mapa de la Faccion Real
Private Const MapaGuerra2Y As Byte = 50 'Y del Mapa de la Faccion Caos
 
Private Const TiempoEntreGuerra As Byte = 4 'Duración de entre una Guerra y otra (Minutos)
Private Const DuracionGuerra As Byte = 8 'Duración de Guerra (Minutos)
 
Private Const OroRecompenza As Long = 200000 'Oro de Recompenz
 
Private Const FONTGUERRA As String = "~255~180~180~1~0"
 

Public Sub IniciarGuerra(ByVal userindex As Integer)
    If userindex <> 0 Then
        If HayGuerra Then
            Call SendData(SendTarget.ToIndex, userindex, 0, "||Ya hay una Guerra Actualmente." & FONTGUERRA)
            Exit Sub
        End If
    End If
    
    HayGuerra = True
    TiempoGuerra = 0

    CiudadGuerra = RandomNumber(1, 2)
        Select Case CiudadGuerra
            Case 1 'Mapa de la Faccion Real
                MapInfo(MapaGuerra1).Pk = True
                    With PosicionNPC
                        .Map = MapaGuerra1
                        .x = MapaGuerra1X
                        .Y = MapaGuerra1Y
                    End With
                SpawnNpc NPC2, PosicionNPC, True, False
                CiudadGuerra = MapaGuerra1
                NPCGuerra = NPC2
                
            Case 2 'Mapa de la Faccion Caos
                MapInfo(MapaGuerra2).Pk = True
                    NPCGuerra = NPC1
                    With PosicionNPC
                        .Map = MapaGuerra2
                        .x = MapaGuerra2X
                        .Y = MapaGuerra2Y
                    End With
                SpawnNpc NPC1, PosicionNPC, True, False
                CiudadGuerra = MapaGuerra2
                NPCGuerra = NPC1
        End Select
        
    Call SendData(SendTarget.toall, 0, 0, "||La Guerra ha Comenzado, Para participar envia /GUERRA" & FONTGUERRA)

Exit Sub
End Sub

Public Sub TerminaGuerra(ByVal FaccionGanadora As String, MurioNPC As Boolean)
Dim UI As Integer, x As Integer, Y As Integer

    If FaccionGanadora = "Real" Then
        Call SendData(SendTarget.toall, 0, 0, "||La Guerra ha terminado, La facción Ganadora es la Armada Real, Los miembros de esta faccion reciben a cambio " & OroRecompenza & " Monedas de oro." & FONTGUERRA)
    ElseIf FaccionGanadora = "Caos" Then
        Call SendData(SendTarget.toall, 0, 0, "||La Guerra ha terminado, La facción Ganadora es la Legion Oscura, Los miembros de esta faccion reciben a cambio " & OroRecompenza & " Monedas de oro." & FONTGUERRA)
    End If

    For UI = 1 To NumUsers
        If UserList(UI).flags.Guerra = True Then
            If FaccionGanadora = "Caos" And Criminal(UI) Then
                    UserList(UI).Stats.GLD = UserList(UI).Stats.GLD + OroRecompenza
                    SendUserStatsBox UI
            End If
            If FaccionGanadora = "Real" And Not Criminal(UI) Then
                    UserList(UI).Stats.GLD = UserList(UI).Stats.GLD + OroRecompenza
                    SendUserStatsBox UI
            End If
            UserList(UI).flags.Guerra = False
        End If
    Next UI
    
    If Not MurioNPC Then
        For Y = 1 To 100
            For x = 1 To 100
                If MapData(CiudadGuerra, x, Y).NpcIndex > 0 Then
                    If Npclist(MapData(CiudadGuerra, x, Y).NpcIndex).Numero = NPCGuerra Then
                        Call QuitarNPC(MapData(CiudadGuerra, x, Y).NpcIndex)
                    End If
                End If
            Next x
        Next Y
    End If
    
    Call SendData(SendTarget.toall, 0, 0, "|G0")
    
    MapInfo(CiudadGuerra).Pk = True
    HayGuerra = False
    TiempoGuerra = 0
Exit Sub
End Sub

Public Sub TimeGuerra()
TiempoGuerra = TiempoGuerra + 1

    If Not HayGuerra And GuerrasAutomaticas Then
    
        If TiempoEntreGuerra - TiempoGuerra = 3 Or TiempoEntreGuerra - TiempoGuerra = 2 Or TiempoEntreGuerra - TiempoGuerra = 1 Then
            Call SendData(SendTarget.toall, 0, 0, "||Los Miembros de la Legion Ocura y la Armada Real Pelearan una Guerra en " & TiempoEntreGuerra - TiempoGuerra & " Minutos, Equipense y preparense para defender a su Reino! Grandes Riquezas les esperan a los Sobrevivientes Victoriosos." & FONTGUERRA)
        End If
        If val(TiempoGuerra) = TiempoEntreGuerra Then
            IniciarGuerra 0
            Exit Sub
        End If
    End If
    
    If HayGuerra Then
        If val(TiempoGuerra) = DuracionGuerra Then
            If CiudadGuerra = MapaGuerra1 Then
                TerminaGuerra "Caos", False
            Else
                TerminaGuerra "Real", False
            End If
        Else
            Call SendData(SendTarget.toall, 0, 0, "||Quedan " & (DuracionGuerra - TiempoGuerra) & " Minutos de Guerra. Para defender a tu Reino Envia /Guerra." & FONTGUERRA)
        End If
    End If
Exit Sub
End Sub

Public Sub EntrarGuerra(ByVal userindex As Integer)
    If Not HayGuerra Then
        Call SendData(SendTarget.ToIndex, userindex, 0, "||No Hay Ninguna Guerra Actualmente." & FONTGUERRA)
        Exit Sub
    End If
        
    If UserList(userindex).flags.Guerra = True Then
        Call SendData(SendTarget.ToIndex, userindex, 0, "||Ya estas participando de la Guerra." & FONTGUERRA)
        Exit Sub
    End If

    If CiudadGuerra = MapaGuerra1 Then
        WarpUserChar userindex, MapaGuerra1, MapaGuerra1X - 10, MapaGuerra1Y - 10, True
    Else
        WarpUserChar userindex, MapaGuerra2, MapaGuerra2X - 10, MapaGuerra2Y - 10, True
    End If
    
    Call SendData(SendTarget.ToIndex, userindex, 0, "||La Guerra ha Comenzado para ti, Defiende a tu faccion para recibir una recompenza." & FONTGUERRA)
    Call SendData(SendTarget.ToIndex, userindex, 0, "|G1")
    UserList(userindex).flags.Guerra = True
Exit Sub
End Sub

Public Sub GuerrasAuto(ByVal userindex As Integer, OnOff As Integer)
    If OnOff = 1 Then
        Call SendData(SendTarget.ToIndex, userindex, 0, "||Las Guerras Automaticas han sido Ativadas." & FONTGUERRA)
        GuerrasAutomaticas = True
    Else
        Call SendData(SendTarget.ToIndex, userindex, 0, "||Las Guerras Automaticas han sido Desativadas." & FONTGUERRA)
        GuerrasAutomaticas = False
    End If
Exit Sub
End Sub

Public Sub EmpatarGuerra(ByVal userindex As Integer)
Dim UI As Integer, x As Integer, Y As Integer

    Call SendData(SendTarget.toall, 0, 0, "||La Guerra ha terminado, Ninguna Facción resulto victoriosa." & FONTGUERRA)

    For UI = 1 To NumUsers
            UserList(UI).flags.Guerra = False
    Next UI
    
    For Y = 1 To 100
        For x = 1 To 100
            If MapData(CiudadGuerra, x, Y).NpcIndex > 0 Then
                If Npclist(MapData(CiudadGuerra, x, Y).NpcIndex).Numero = NPCGuerra Then
                    Call QuitarNPC(MapData(CiudadGuerra, x, Y).NpcIndex)
                End If
            End If
        Next x
    Next Y
    Call SendData(SendTarget.toall, 0, 0, "|G0")
    MapInfo(CiudadGuerra).Pk = True
    HayGuerra = False
    TiempoGuerra = 0
Exit Sub
End Sub

