Attribute VB_Name = "modNuevoTimer"

Option Explicit

'
' Las siguientes funciones devuelven TRUE o FALSE si el intervalo
' permite hacerlo. Si devuelve TRUE, setean automaticamente el
' timer para que no se pueda hacer la accion hasta el nuevo ciclo.
'

' CASTING DE HECHIZOS
Public Function IntervaloPermiteLanzarSpell(ByVal UserIndex As Integer, Optional ByVal Actualizar As Boolean = True) As Boolean
Dim TActual As Long

TActual = GetTickCount() And &H7FFFFFFF

If TActual - UserList(UserIndex).Counters.TimerLanzarSpell >= 40 * IntervaloUserPuedeCastear Then
    If Actualizar Then UserList(UserIndex).Counters.TimerLanzarSpell = TActual
    IntervaloPermiteLanzarSpell = True
Else
    IntervaloPermiteLanzarSpell = False
End If

End Function


Public Function IntervaloPermiteAtacar(ByVal UserIndex As Integer, Optional ByVal Actualizar As Boolean = True) As Boolean
Dim TActual As Long

TActual = GetTickCount() And &H7FFFFFFF

If TActual - UserList(UserIndex).Counters.TimerPuedeAtacar >= 40 * IntervaloUserPuedeAtacar Then
    If Actualizar Then UserList(UserIndex).Counters.TimerPuedeAtacar = TActual
    IntervaloPermiteAtacar = True
Else
    IntervaloPermiteAtacar = False
End If
End Function


' ATAQUE CUERPO A CUERPO
'Public Function IntervaloPermiteAtacar(ByVal UserIndex As Integer, Optional ByVal Actualizar As Boolean = True) As Boolean
'Dim TActual As Long
'
'TActual = GetTickCount() And &H7FFFFFFF''
'
'If TActual - UserList(UserIndex).Counters.TimerPuedeAtacar >= 40 * IntervaloUserPuedeAtacar Then
'    If Actualizar Then UserList(UserIndex).Counters.TimerPuedeAtacar = TActual
'    IntervaloPermiteAtacar = True
'Else
'    IntervaloPermiteAtacar = False
'End If
'End Function

' TRABAJO
Public Function IntervaloPermiteTrabajar(ByVal UserIndex As Integer, Optional ByVal Actualizar As Boolean = True) As Boolean
Dim TActual As Long

TActual = GetTickCount() And &H7FFFFFFF

If TActual - UserList(UserIndex).Counters.TimerPuedeTrabajar >= 40 * IntervaloUserPuedeTrabajar Then
    If Actualizar Then UserList(UserIndex).Counters.TimerPuedeTrabajar = TActual
    IntervaloPermiteTrabajar = True
Else
    IntervaloPermiteTrabajar = False
End If
End Function

' USAR OBJETOS
Public Function IntervaloPermiteUsar(ByVal UserIndex As Integer, Optional ByVal Actualizar As Boolean = True) As Boolean
Dim TActual As Long

TActual = GetTickCount() And &H7FFFFFFF

If TActual - UserList(UserIndex).Counters.TimerUsar >= IntervaloUserPuedeUsar Then
    If Actualizar Then UserList(UserIndex).Counters.TimerUsar = TActual
    IntervaloPermiteUsar = True
Else
    IntervaloPermiteUsar = False
End If

End Function

Public Function IntervaloPermiteUsarArcos(ByVal UserIndex As Integer, Optional ByVal Actualizar As Boolean = True) As Boolean
Dim TActual As Long

TActual = GetTickCount() And &H7FFFFFFF

If TActual - UserList(UserIndex).Counters.TimerUsar >= IntervaloFlechasCazadores Then
    If Actualizar Then UserList(UserIndex).Counters.TimerUsar = TActual
    IntervaloPermiteUsarArcos = True
Else
    IntervaloPermiteUsarArcos = False
End If

End Function

Sub ControlarPortalLum(ByVal UserIndex As Integer)
   
    If UserList(UserIndex).Counters.CreoTeleport = True Then
        Call EraseObj(ToMap, 0, UserList(UserIndex).flags.DondeTiroMap, MapData(UserList(UserIndex).flags.DondeTiroMap, UserList(UserIndex).flags.DondeTiroX, UserList(UserIndex).flags.DondeTiroY).OBJInfo.Amount, UserList(UserIndex).flags.DondeTiroMap, UserList(UserIndex).flags.DondeTiroX, UserList(UserIndex).flags.DondeTiroY) 'verificamos que destruye el objeto anterior.
        MapData(UserList(UserIndex).flags.DondeTiroMap, UserList(UserIndex).flags.DondeTiroX, UserList(UserIndex).flags.DondeTiroY).TileExit.Map = 0
        MapData(UserList(UserIndex).flags.DondeTiroMap, UserList(UserIndex).flags.DondeTiroX, UserList(UserIndex).flags.DondeTiroY).TileExit.X = 0
        MapData(UserList(UserIndex).flags.DondeTiroMap, UserList(UserIndex).flags.DondeTiroX, UserList(UserIndex).flags.DondeTiroY).TileExit.Y = 0
        UserList(UserIndex).flags.DondeTiroMap = ""
        UserList(UserIndex).flags.DondeTiroX = ""
        UserList(UserIndex).flags.DondeTiroY = ""
    End If
End Sub
