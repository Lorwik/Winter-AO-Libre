Attribute VB_Name = "modSubasta"
Public Sub IniciarSubasta(ByVal UserIndex As Integer)
 
On Error Resume Next
   
    Call UpdateUserInv(True, UserIndex, 0)
    Call SendUserStatsBox(UserIndex)
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "INITSUB")
   
    Subastando = True
 
End Sub
 
 
Sub NPCSubasta(ByVal UserIndex As Integer, ByVal Item As Integer, ByVal Cantidad As Integer, ByVal Precio As Long)
On Error GoTo errhandler
    Dim NpcIndex As Integer
   
    NpcIndex = UserList(UserIndex).flags.TargetNPC
   
    If ObjData(UserList(UserIndex).Invent.Object(Item).ObjIndex).Real = 1 Then
        If Npclist(NpcIndex).name <> "SR" Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Las armaduras faccionarias no pueden ser subastadas." & FONTTYPE_WARNING)
            Subastando = False
            Exit Sub
        End If
       
    ElseIf ObjData(UserList(UserIndex).Invent.Object(Item).ObjIndex).Caos = 1 Then
       
        If Npclist(NpcIndex).name <> "SC" Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Las armaduras faccionarias no pueden ser subastadas." & FONTTYPE_WARNING)
            Subastando = False
            Exit Sub
        End If
    End If
   
    If ObjData(UserList(UserIndex).Invent.Object(Item).ObjIndex).NoSubasta = 1 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "PRB55")
        Subastando = False
        Exit Sub
    End If
   
    If UserList(UserIndex).Invent.Object(Item).Amount > 0 And UserList(UserIndex).Invent.Object(Item).Equipped = 0 Then
        If Cantidad > 0 And Cantidad > UserList(UserIndex).Invent.Object(Item).Amount Then Cantidad = UserList(UserIndex).Invent.Object(Item).Amount
        Call NpcSubastaObj(UserIndex, CInt(Item), Cantidad, Precio)
    End If
Exit Sub
 
errhandler:
    Call LogError("Error en vender item: " & Err.Description)
End Sub
 
Sub NpcSubastaObj(ByVal UserIndex As Integer, ByVal ObjIndex As Integer, ByVal Cantidad As Integer, ByVal Precio As Long)
On Error GoTo errorh
    Dim obji As Integer
    Dim NpcIndex As Integer
         
    If Cantidad < 1 Then Exit Sub
   
    NpcIndex = UserList(UserIndex).flags.TargetNPC
    obji = UserList(UserIndex).Invent.Object(ObjIndex).ObjIndex
   
    If ObjData(obji).Newbie = 1 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No comercio objetos para newbies." & FONTTYPE_INFO)
        Subastando = False
        Exit Sub
    End If
     
    If obji = iORO Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||El npc no esta interesado en comprar ese objeto." & FONTTYPE_WARNING)
        Subastando = False
        Exit Sub
    End If
   
   
    Call QuitarUserInvItem(UserIndex, CByte(ObjIndex), Cantidad)
    Call UpdateUserInv(True, UserIndex, 0)
   
    Subastante = UserList(UserIndex).name
    subastCant = Cantidad
    subastObj = ObjData(obji).name
    subastPrice = Precio
    subastObjIndex = obji
    subastLastOffer = ""
   
    frmMain.tmrSubasta.Enabled = True
   
    UserList(UserIndex).flags.ParticipaSubasta = True
   
    Call SendData(SendTarget.ToAll, 0, 0, "PRE1," & Subastante & "," & subastObj & "," & subastCant & "," & subastPrice)
   
Exit Sub
 
errorh:
    Call LogError("Error en NPCSUBASTAOBJ. " & Err.Description)
End Sub
 
Sub OfertarSubasta(ByVal UserIndex As Integer, ByVal Oferta As Long)
 
If Oferta <= subastPrice Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "Tu oferta debe superar las " & subastPrice & " monedas de oro.")
    Exit Sub
End If
 
If subastLastOffer <> "" Then Call DevolverOro(subastLastOffer, subastPrice)
 
UserList(UserIndex).flags.ParticipaSubasta = True
 
subastLastOffer = UserList(UserIndex).name
subastPrice = Oferta
 
UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - Oferta
 
Call SendUserStatsBox(UserIndex)
Call SendData(SendTarget.tosubasta, 0, 0, "PRE2," & subastLastOffer & "," & subastPrice)
 
End Sub
 
Public Sub ResolverSubasta()
 
frmMain.tmrSubasta.Enabled = False
Subastando = False
 
If subastLastOffer <> "" Then
    Debug.Print "Termino la subasta."
    Call DarOroSubastante(Subastante, subastPrice)
    Call DarItemOferente(subastLastOffer)
Else
    Call SendData(SendTarget.ToAll, 0, 0, "PRE3")
    Call DevolverItem(Subastante)
End If
 
Call RestartSubasta
 
End Sub
 
Sub RestartSubasta()
 
Dim i As Integer
 
For i = 1 To LastUser
    If UserList(i).flags.ParticipaSubasta = True Then UserList(i).flags.ParticipaSubasta = False
Next i
 
End Sub
 
Public Sub DevolverItem(ByVal rctFin As String)
 
Dim LoopC As Integer
Dim MiObj As Obj
 
MiObj.Amount = subastCant
MiObj.ObjIndex = subastObjIndex
 
For LoopC = 1 To LastUser
   
    If UserList(LoopC).name = rctFin Then
        IndexOffer = LoopC
        Exit For
    End If
   
   
    IndexOffer = 0
   
Next LoopC
 
If IndexOffer <> 0 Then
    If Not MeterItemEnInventario(IndexOffer, MiObj) Then
        'Call UserDepositaItem(UserIndex, val(ReadField(1, rData, 44)), val(ReadField(2, rData, 44)))
        Call TirarItemAlPiso(UserList(IndexOffer).Pos, MiObj)
    End If
Else
    Call DepositarItemOffline(rctFin, subastObjIndex, subastCant)
End If
 
End Sub
 
Public Sub DarItemOferente(ByVal rctOfe As String)
 
Dim LoopC As Integer
Dim MiObj As Obj
 
MiObj.Amount = subastCant
MiObj.ObjIndex = subastObjIndex
 
For LoopC = 1 To LastUser
   
    If UserList(LoopC).name = rctOfe Then
        IndexOffer = LoopC
        Exit For
    End If
   
   
    IndexOffer = 0
   
Next LoopC
 
If IndexOffer <> 0 Then
    If Not MeterItemEnInventario(IndexOffer, MiObj) Then
        'Call UserDepositaItem(UserIndex, val(ReadField(1, rData, 44)), val(ReadField(2, rData, 44)))
        Call TirarItemAlPiso(UserList(IndexOffer).Pos, MiObj)
    End If
Else
    Call DepositarItemOffline(rctOfe, subastObjIndex, subastCant)
End If
 
End Sub
 
 
Public Sub DarOroSubastante(ByVal rctSubastante As String, Oro As Long)
 
Dim LoopC As Integer
Dim BancoOffline As Integer
Dim IndexOffer As Integer
 
 
For LoopC = 1 To LastUser
   
    If UserList(LoopC).name = rctSubastante Then
        IndexOffer = LoopC
        Exit For
    End If
   
   
    IndexOffer = 0
   
Next LoopC
 
If IndexOffer <> 0 Then
    UserList(IndexOffer).Stats.GLD = UserList(IndexOffer).Stats.GLD + Oro
    Call SendUserGold(IndexOffer)
    Call SendData(ToIndex, IndexOffer, 0, "PRE5," & Oro)
Else
    BancoOffline = val(GetVar(App.Path & "\Charfile\" & rctSubastante & ".chr", "STATS", "Banco"))
    Call WriteVar(App.Path & "\Charfile\" & rctSubastante & ".chr", "STATS", "Banco", val(BancoOffline + Oro))
End If
 
Call SendData(SendTarget.ToAll, 0, 0, "PRE6," & subastObj & "," & subastCant & "," & Oro)
 
End Sub
 
Public Sub DevolverOro(ByVal LastOffer As String, Price As Long)
 
Dim LoopC As Integer
Dim BancoOffline As Integer
Dim IndexOffer As Integer
 
For LoopC = 1 To LastUser
   
    If UserList(LoopC).name = LastOffer Then
        IndexOffer = LoopC
        Exit For
    End If
   
    IndexOffer = 0
   
Next LoopC
 
If IndexOffer <> 0 Then
    UserList(IndexOffer).Stats.GLD = UserList(IndexOffer).Stats.GLD + Price
    Call SendUserGold(IndexOffer)
    Call SendData(ToIndex, IndexOffer, 0, "PRE4")
Else
    BancoOffline = val(GetVar(App.Path & "\Charfile\" & LastOffer & ".chr", "STATS", "Banco"))
    Call WriteVar(App.Path & "\Charfile\" & LastOffer & ".chr", "STATS", "Banco", val(BancoOffline + Price))
End If
 
End Sub
 
Function EstaConectado(ByVal User As Integer) As Boolean
 
If UserList(User).ConnID <> -1 And UserList(User).flags.UserLogged Then
EstaConectado = True
Exit Function
End If
 
EstaConectado = False
End Function
 
Sub DepositarItemOffline(ByVal Comprador As String, ByVal ObjIndex As Integer, ByVal Cantidad As Integer)
 
Dim Slot As Integer
Dim obji As Integer
Dim Nitems As Integer
Dim LoopC As Integer
Dim ln As String
 
If Cantidad < 1 Then Exit Sub
 
obji = ObjIndex
 
 
Nitems = GetVar(App.Path & "\Charfile\" & Comprador & ".chr", "BancoInventory", "CantidadItems")
 
For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS
    ln = GetVar(App.Path & "\Charfile\" & Comprador & ".chr", "BancoInventory", "Obj" & LoopC)
    Debug.Print ln
    bancoObj(LoopC).ObjIndex = CInt(ReadField(1, ln, 45))
    bancoObj(LoopC).Amount = CInt(ReadField(2, ln, 45))
Next LoopC
 
'¿Ya tiene un objeto de este tipo?
Slot = 1
Do Until bancoObj(Slot).ObjIndex = obji And _
         bancoObj(Slot).Amount + Cantidad <= MAX_INVENTORY_OBJS
            Slot = Slot + 1
       
            If Slot > MAX_BANCOINVENTORY_SLOTS Then
                Exit Do
            End If
Loop
 
'Sino se fija por un slot vacio antes del slot devuelto
If Slot > MAX_BANCOINVENTORY_SLOTS Then
        Slot = 1
        Do Until bancoObj(Slot).ObjIndex = 0
            Slot = Slot + 1
 
            If Slot > MAX_BANCOINVENTORY_SLOTS Then
                Call ItemLastPos(Comprador, Cantidad, obji)
                Exit Sub
                Exit Do
            End If
        Loop
        If Slot <= MAX_BANCOINVENTORY_SLOTS Then Call WriteVar(App.Path & "\Charfile\" & Comprador & ".chr", "BancoInventory", "CantidadItems", Nitems + 1)
       
End If
 
If Slot <= MAX_BANCOINVENTORY_SLOTS Then 'Slot valido
    'Mete el obj en el slot
    If bancoObj(Slot).Amount + Cantidad <= MAX_INVENTORY_OBJS Then
       
        'Menor que MAX_INV_OBJS
        Call WriteVar(App.Path & "\Charfile\" & Comprador & ".chr", "BancoInventory", "Obj" & Slot, obji & "-" & bancoObj(Slot).Amount + Cantidad)
 
    Else
        Call ItemLastPos(Comprador, Cantidad, obji)
    End If
 
Else
    'Call QuitarUserInvItem(UserIndex, CByte(ObjIndex), Cantidad)
End If
 
End Sub
 
Sub ItemLastPos(ByVal PJ As String, sCant As Integer, objndX As Integer)
 
Dim DameLastPos As String
Dim LastPos As WorldPos
Dim MiObj As Obj
 
MiObj.Amount = sCant
MiObj.ObjIndex = objndX
 
DameLastPos = GetVar(App.Path & "\Charfile\" & PJ & ".chr", "INIT", "Position")
 
LastPos.Map = CInt(ReadField(1, DameLastPos, 45))
LastPos.X = CInt(ReadField(2, DameLastPos, 45))
LastPos.Y = CInt(ReadField(3, DameLastPos, 45))
 
Call TirarItemAlPiso(LastPos, MiObj)
 
End Sub

