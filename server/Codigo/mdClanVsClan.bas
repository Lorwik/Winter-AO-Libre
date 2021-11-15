Attribute VB_Name = "mdClanVsClan"
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'Fecha: 10 de diciembre del año 2009
'Creador: SaturoS - Joan Calderón.
'Descripcion: Codigo fuente retos clanes vs clanes con capitanes y sumon.
'Porfavor mantener este comentario y la autoria del codigo.
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

' arrays de participantes sumoneados
Public Clan1(1 To 3) As Integer
Public Clan2(1 To 3) As Integer
' variables de lideres de batalla de clanes
Public LiderClan1 As String
Public LiderClan2 As String
' nombres de los clanes participantes
Public NombreClan1 As Integer
Public NombreClan2 As Integer
' Mapa de guerra de clanes
Private Const MapaClan As Byte = 118
' posiciones de espera para sumonear
Private Const MapaClan1_x As Byte = 40
Private Const MapaClan2_x As Byte = 48
Private Const MapaClan1_y As Byte = 12
Private Const MapaClan2_y As Byte = 12
' posiciones de esquinas de cada clan para pelear
Private Const Esquina_x_Clan1 As Byte = 53
Private Const Esquina_y_Clan1 As Byte = 46
Private Const Esquina_x_Clan2 As Byte = 71
Private Const Esquina_y_Clan2 As Byte = 45
' mapa y posiciones cuando acaba la guerra
Private Const Mapa_Fuera As Byte = 1
Private Const Mapa_Fuera_x As Byte = 55
Private Const Mapa_Fuera_y As Byte = 47
' indica si ya hay una guerra de clanes
Public YaHayClan As Boolean
Public Sub EmpiezaSumon(ByVal userindex As Integer, ByVal userindex2 As Integer)
On Error GoTo errorclan
' enviamos el mensaje a cada lider para que sumonee
Call SendData(SendTarget.ToIndex, userindex, 0, "||Estas a punto de desatar una batalla con otro clan, tienes la oportunidad de sumonear a 3 integrantes de tu clan para unirse a esta batalla. Para sumonear tipea /SUMCLAN NICK" & FONTTYPE_INFO)
Call SendData(SendTarget.ToIndex, userindex2, 0, "||Estas a punto de desatar una batalla con otro clan, tienes la oportunidad de sumonear a 3 integrantes de tu clan para unirse a esta batalla. Para sumonear tipea /SUMCLAN NICK" & FONTTYPE_INFO)
 ' sumoneamos a los lideres al area de sumoneo de clan
Call WarpUserChar(userindex, MapaClan, MapaClan1_x, MapaClan1_y, True)
Call WarpUserChar(userindex2, MapaClan, MapaClan2_x, MapaClan2_y, True)
' ponemos en true las variables de sumonear para que los lideres puedan sumonear
UserList(userindex).flags.PuedeSumon = True
UserList(userindex).flags.PuedeSumon = True
' Nombre de cada lider que inicio la guerra
LiderClan1 = userindex
LiderClan2 = userindex2
' Nombre de los clanes
NombreClan1 = Guilds(UserList(userindex).GuildIndex).GuildName
NombreClan2 = Guilds(UserList(userindex2).GuildIndex).GuildName
' reseteamos los userindex de clanes
Dim i As Integer
For i = 1 To 3
Clan1(i) = -1
Clan2(i) = -1
Next i
' se avisa que ya hay una guerra de clanes, para que no se puedan hacer 2
YaHayClan = True
errorclan:
End Sub
Public Sub Sumon(ByVal userindex As Integer, ByVal Sumonear As Integer)
On Error GoTo errorclan
' Sumoneamos el clan numero 1
If UserList(userindex).flags.PuedeSumon = True And LiderClan1 = userindex Then
    If Guilds(UserList(Sumonear).GuildIndex).GuildName <> NombreClan1 Then
        Call SendData(SendTarget.ToIndex, userindex, 0, "||No esta en tu clan!" & FONTTYPE_INFO)
        Exit Sub
    End If
    
Dim i As Integer
For i = LBound(Clan1) To UBound(Clan1)
    If Clan1(i) = -1 Then
        Clan1(i) = Sumonear
        ' sumoneamos
        Call WarpUserChar(Sumonear, UserList(userindex).pos.Map, UserList(userindex).pos.x, UserList(userindex).pos.Y + 1, True)
        Call SendData(SendTarget.ToIndex, userindex, 0, "||Has elegido a " & UserList(Clan1(i)).name & " Para que te acompañe en esta batalla." & FONTTYPE_INFO)
        Call SendData(SendTarget.ToIndex, Clan1(i), 0, "||Has sido elegido para participar en la batalla de clanes." & FONTTYPE_INFO)
        Exit For
    End If
Next i

' si ya se alcanzo el limite se le avisa, y no puede sumonear mas
If Clan1(3) <> -1 Then
    Call SendData(SendTarget.ToIndex, userindex, 0, "||Has alcanzado el limite de participantes para tu clan." & FONTTYPE_INFO)
    UserList(userindex).flags.PuedeSumon = False
End If

End If

' Sumoneamos el clan numero 2
If UserList(userindex).flags.PuedeSumon = True And LiderClan2 = userindex Then
    If Guilds(UserList(Sumonear).GuildIndex).GuildName <> NombreClan2 Then
    Call SendData(SendTarget.ToIndex, userindex, 0, "||No esta en tu clan!" & FONTTYPE_INFO)
    Exit Sub
    End If
Dim x As Integer

For x = LBound(Clan2) To UBound(Clan2)
    If Clan2(x) = -1 Then
        Clan2(x) = Sumonear
        ' sumoneamos
        Call WarpUserChar(Sumonear, UserList(userindex).pos.Map, UserList(userindex).pos.x, UserList(userindex).pos.Y + 1, True)
        Call SendData(SendTarget.ToIndex, userindex, 0, "||Has elegido a " & UserList(Clan2(i)).name & " Para que te acompañe en esta batalla." & FONTTYPE_INFO)
        Call SendData(SendTarget.ToIndex, Clan2(i), 0, "||Has sido elegido para participar en la batalla de clanes." & FONTTYPE_INFO)
        Exit For
    End If
Next x

' si ya se alcanzo el limite se le avisa y no puede sumonear mas
If Clan2(3) <> -1 Then
    Call SendData(SendTarget.ToIndex, userindex, 0, "||Has alcanzado el limite de participantes para tu clan." & FONTTYPE_INFO)
    End If
End If

' cuando los dos clanes tengan el máximo de participantes, empieza la batalla.
If Clan1(3) <> -1 And Clan2(3) <> -1 Then
    Call EmpiezaClan
End If
errorclan:
End Sub
Public Sub EmpiezaClan()
On Error GoTo errorclan
Call SendData(SendTarget.toall, 0, 0, "||Clanes> Una nueva batalla de clanes ha comenzado. Se enfrentan el clan <" & NombreClan1 & "> VS <" & NombreClan2 & ">" & FONTTYPE_GUILD)
Dim i As Integer
For i = 1 To 3
    If UserList(i).flags.EnClanes = True Then
        ' Sumoneamos al primer clan a pelear
        Dim NuevaPos As WorldPos
        Dim FuturePos As WorldPos
        FuturePos.Map = MapaClan
        FuturePos.x = Esquina_x_Clan1: FuturePos.Y = Esquina_y_Clan1
        Call ClosestLegalPos(FuturePos, NuevaPos)
        If NuevaPos.x <> 0 And NuevaPos.Y <> 0 Then Call WarpUserChar(Clan1(i), NuevaPos.Map, NuevaPos.x, NuevaPos.Y, True)

        ' Sumoneamos al segundo clan a pelear
        Dim NuevaPos2 As WorldPos
        Dim FuturePos2 As WorldPos
        FuturePos2.Map = MapaClan
        FuturePos2.x = Esquina_x_Clan2: FuturePos2.Y = Esquina_y_Clan2
        Call ClosestLegalPos(FuturePos2, NuevaPos2)
        If NuevaPos2.x <> 0 And NuevaPos2.Y <> 0 Then Call WarpUserChar(Clan2(i), NuevaPos2.Map, NuevaPos2.x, NuevaPos2.Y, True)
    End If
Next i

' mandamos al primer lider a la esquina
FuturePos.Map = MapaClan
FuturePos.x = Esquina_x_Clan1: FuturePos.Y = Esquina_y_Clan1
Call ClosestLegalPos(FuturePos, NuevaPos)
If NuevaPos.x <> 0 And NuevaPos.Y <> 0 Then Call WarpUserChar(LiderClan1, NuevaPos.Map, NuevaPos.x, NuevaPos.Y, True)

' ahora mandamos al segundo lider a la esquina
Call ClosestLegalPos(FuturePos2, NuevaPos2)
If NuevaPos2.x <> 0 And NuevaPos2.Y <> 0 Then Call WarpUserChar(LiderClan2, NuevaPos2.Map, NuevaPos2.x, NuevaPos2.Y, True)
errorclan:
End Sub
Public Sub ClanMuere(ByVal userindex As Integer)
On Error GoTo errorclan
' cuando muere lo ponemos como muerto
UserList(userindex).flags.MuereClan = True
If UserList(Clan1(1)).flags.MuereClan = True And UserList(Clan1(2)).flags.MuereClan = True And UserList(Clan1(3)).flags.MuereClan = True And UserList(LiderClan1).flags.MuereClan = True Then
    Call SendData(SendTarget.toall, 0, 0, "||Clanes> El clan <" & NombreClan1 & "> Ha ganado la batalla de clanes!." & FONTTYPE_GUILD)
    Call SendData(SendTarget.toall, 0, 0, "||Clanes> El clan <" & NombreClan1 & "> Recibe como premio: Ponele aqui los premios." & FONTTYPE_GUILD)
    Dim i As Integer
    
    For i = LBound(Clan1) To UBound(Clan1)
        If UserList(Clan1(i)).flags.EnClanes = True Then
        'userlist(clan1(i)). 'bla bla bla para que le des el premio a cada uno del clan
        End If
    Next i
    
    Call TerminaClan
    
ElseIf UserList(Clan2(1)).flags.MuereClan = True And UserList(Clan2(2)).flags.MuereClan = True And UserList(Clan2(3)).flags.MuereClan = True And UserList(LiderClan2).flags.MuereClan = True Then
    Call SendData(SendTarget.toall, 0, 0, "||Clanes> El clan <" & NombreClan2 & "> Ha ganado la batalla de clanes!." & FONTTYPE_GUILD)
    Call SendData(SendTarget.toall, 0, 0, "||Clanes> El clan <" & NombreClan2 & "> Recibe como premio: Ponele aqui los premios." & FONTTYPE_GUILD)
    Dim x As Integer
    
    For x = LBound(Clan2) To UBound(Clan2)
        If UserList(Clan2(x)).flags.EnClanes = True Then
        'userlist(clan2(x)). 'bla bla bla para que le des el premio a cada uno del clan
        End If
    Next x
    
    Call TerminaClan
End If
errorclan:
End Sub
Public Sub ClanDesconecta(ByVal userindex As Integer)
On Error GoTo errorclan
' lo tiramos a ulla, y le quitamos 1kk por desgraciado.
Call SendData(SendTarget.toall, 0, 0, "||Clanes> " & UserList(userindex).name & " del clan <" & Guilds(UserList(userindex).GuildIndex).GuildName & "> Ha desconectado en la batalla de clanes. Se le penaliza con 1kk" & FONTTYPE_GUILD)
     
     If UserList(userindex).Stats.GLD >= 1000000 Then
        UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD - 1000000
    End If
    
Call WarpUserChar(userindex, Mapa_Fuera, Mapa_Fuera_x, Mapa_Fuera_y, True)
Call ClanMuere(userindex)
UserList(userindex).flags.EnClanes = False

errorclan:
End Sub
Public Sub TerminaClan()
On Error GoTo errorclan
Dim i As Integer
For i = 1 To 3
    ' Sumoneamos al primer clan a ciuad
        If UserList(i).flags.EnClanes = True Then
            Dim NuevaPos As WorldPos
            Dim FuturePos As WorldPos
            FuturePos.Map = Mapa_Fuera
            FuturePos.x = Mapa_Fuera_x: FuturePos.Y = Mapa_Fuera_y
            Call ClosestLegalPos(FuturePos, NuevaPos)
            If NuevaPos.x <> 0 And NuevaPos.Y <> 0 Then Call WarpUserChar(Clan1(i), NuevaPos.Map, NuevaPos.x, NuevaPos.Y, True)

            ' Sumoneamos al segundo clan a ciudad
            Dim NuevaPos2 As WorldPos
            Dim FuturePos2 As WorldPos
            FuturePos2.Map = Mapa_Fuera
            FuturePos2.x = Mapa_Fuera_x: FuturePos2.Y = Mapa_Fuera_y
            Call ClosestLegalPos(FuturePos2, NuevaPos2)
            If NuevaPos2.x <> 0 And NuevaPos2.Y <> 0 Then Call WarpUserChar(Clan2(i), NuevaPos2.Map, NuevaPos2.x, NuevaPos2.Y, True)

            'reseteamos todo
            UserList(i).flags.EnClanes = False
            UserList(i).flags.MuereClan = False
            UserList(i).flags.PuedeSumon = False
        End If
Next i

' mandamos a los lideres al mapa de afuera
FuturePos.Map = Mapa_Fuera
FuturePos.x = Mapa_Fuera_x: FuturePos.Y = Mapa_Fuera_y
Call ClosestLegalPos(FuturePos, NuevaPos)
If NuevaPos.x <> 0 And NuevaPos.Y <> 0 Then Call WarpUserChar(LiderClan1, NuevaPos.Map, NuevaPos.x, NuevaPos.Y, True)

FuturePos.Map = Mapa_Fuera
FuturePos.x = Mapa_Fuera_x: FuturePos.Y = Mapa_Fuera_y
Call ClosestLegalPos(FuturePos, NuevaPos)
If NuevaPos.x <> 0 And NuevaPos.Y <> 0 Then Call WarpUserChar(LiderClan2, NuevaPos.Map, NuevaPos.x, NuevaPos.Y, True)

'reseteamos a los lideres
UserList(LiderClan1).flags.EnClanes = False
UserList(LiderClan1).flags.MuereClan = False
UserList(LiderClan1).flags.PuedeSumon = False
UserList(LiderClan2).flags.EnClanes = False
UserList(LiderClan2).flags.MuereClan = False
UserList(LiderClan2).flags.PuedeSumon = False

'abrimos para que se haga otra guerra de clan
YaHayClan = False
errorclan:
End Sub

