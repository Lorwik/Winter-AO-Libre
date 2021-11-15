VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGrapchicalInventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
 
Dim InventoryOffset As Long             'Number of lines we scrolled down from topmost
Dim InvSelectedItem As Long             'Currently selected item
Dim MouseDownSelectedItem As Long

 
Dim InvSurface As DirectDrawSurface7            'DD Surface used to render everything
 
Dim UserInventory(1 To MAX_INVENTORY_SLOTS) As Inventory    'User's inventory
 
Dim WithEvents InventoryWindow As PictureBox    'Placeholder where to render the inventory
Attribute InventoryWindow.VB_VarHelpID = -1
Public Sub ScrollInventory(ByVal up As Boolean)
'***************************************************
'Scrolls the graphical inventory up / down
'***************************************************
    'Only move if there is something to show up /down
    If up Then
        If (InventoryOffset + (InventoryWindow.ScaleHeight \ TilePixelHeight)) * (InventoryWindow.ScaleWidth \ TilePixelWidth) < MAX_INVENTORY_SLOTS Then _
            InventoryOffset = InventoryOffset + 1
    Else
        If InventoryOffset > 0 Then _
            InventoryOffset = InventoryOffset - 1
    End If
   
    'Redraw the inventory
    DrawInventory
End Sub
 
Public Sub SelectGold()
'***************************************************
'Sets the gold as the currently selected item
'***************************************************
    Dim prevSelectedItem As Long
   
    'Store preivously selected item
    prevSelectedItem = InvSelectedItem
   
    'Select the gold
    InvSelectedItem = FLAGORO
   
    'Redraw old item to deselect it
    If prevSelectedItem <> FLAGORO Then
        Call DrawInventorySlot(prevSelectedItem)
    End If
End Sub
 
Public Property Get SelectedItem() As Long
'***************************************************
'Retrieves the selected item index
'***************************************************
    SelectedItem = InvSelectedItem
End Property
 
Public Property Get MaxHit(ByVal slot As Byte) As Integer
'***************************************************
'Retrieves the max hit of the selected item
'***************************************************
    MaxHit = UserInventory(slot).MaxHit
End Property
 
Public Property Get MinHit(ByVal slot As Byte) As Integer
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the min hit of the selected item
'***************************************************
    MinHit = UserInventory(slot).MinHit
End Property
 
Public Property Get Def(ByVal slot As Byte) As Integer
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the defense of the selected item
'***************************************************
    Def = UserInventory(slot).Def
End Property
 
Public Property Get GrhIndex(ByVal slot As Byte) As Integer

'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the grh index of the selected item
'***************************************************
    GrhIndex = UserInventory(slot).GrhIndex
End Property
 
Public Property Get Valor(ByVal slot As Byte) As Long
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the value of the selected item
'***************************************************
    Valor = UserInventory(slot).Valor
End Property
 
Public Property Get Amount(ByVal slot As Byte) As Long
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the selected item's amount
'***************************************************
    Amount = UserInventory(slot).Amount
End Property
 
Public Property Get OBJIndex(ByVal slot As Byte) As Integer
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the selected item's object index
'***************************************************
    OBJIndex = UserInventory(slot).OBJIndex
End Property
 
Public Property Get OBJType(ByVal slot As Byte) As Integer
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the selected item's object type
'***************************************************
    OBJType = UserInventory(slot).OBJType
End Property
 
Public Property Get ItemName(ByVal slot As Byte) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the selected item's name
'***************************************************
    ItemName = UserInventory(slot).Name
End Property
 
Public Property Get Equipped(ByVal slot As Byte) As Boolean
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves True if the item at the given pos is eqiupped
'***************************************************
    Equipped = UserInventory(slot).Equipped
End Property
 
Public Sub SetItem(ByVal slot As Byte, ByVal eOBJIndex As Integer, ByVal eAmount As Integer, ByVal eEquipped As Byte, _
                        ByVal eGrhIndex As Integer, ByVal eObjType As Integer, ByVal eMaxHit As Integer, ByVal eMinHit As Integer, _
                        ByVal eDef As Integer, ByVal eValor As Single, ByVal eName As String, Optional ByVal PuedeUsar As Byte = 1)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 12/04/06
'Sets all data for a given inventory slot
'***************************************************
    If slot < 1 Or slot > MAX_INVENTORY_SLOTS Then Exit Sub
   
    With UserInventory(slot)
        .Amount = eAmount
        .Def = eDef
        .Equipped = eEquipped
        .GrhIndex = eGrhIndex
        .MaxHit = eMaxHit
        .MinHit = eMinHit
        .Name = eName
        .OBJIndex = eOBJIndex
        .OBJType = eObjType
        .PuedeUsar = PuedeUsar
        .Valor = eValor
    End With
   
    'Render inventory slot (we don't need to render the whole inventory)
    Call DrawInventorySlot(slot)
End Sub
 
Private Function ClickItem(ByVal X As Long, ByVal Y As Long) As Long
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Selects the item clicked if it's valid and return's it's index
'***************************************************
    Dim TempItem As Long
    Dim temp_x As Long
    Dim temp_y As Long
   
    temp_x = X \ 32 'TilePixelWidth
    temp_y = Y \ 32 'TilePixelHeight
   
    TempItem = temp_x + (temp_y + InventoryOffset) * (InventoryWindow.ScaleWidth \ 32) + 1
   
    'Make sure it's within limits
    If TempItem <= MAX_INVENTORY_SLOTS Then
        ClickItem = TempItem
    End If

    
End Function
 
Private Sub DrawInventory()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Renders the inventory to the given PictureBox
'***************************************************
    Dim loopc As Long
    Dim SrcRect As RECT
    Dim TempRect As RECT
    Dim destRect As RECT
    
   
    With TempRect
        .Bottom = TilePixelHeight
        .Right = TilePixelWidth
    End With
   
    destRect = TempRect
   
    'Clear the inventory window
    InventoryWindow.Cls
   
    For loopc = InventoryOffset * (InventoryWindow.ScaleWidth \ TilePixelWidth) + 1 To UBound(UserInventory)
        'Clear the slot area
        InvSurface.BltFast 0, 0, SurfaceDB.Surface(123123), TempRect, DDBLTFAST_SRCCOLORKEY + DDBLTFAST_WAIT
       
        If UserInventory(loopc).GrhIndex Then
            'Get source rect
            With SrcRect
                .Left = GrhData(UserInventory(loopc).GrhIndex).sX
                .Top = GrhData(UserInventory(loopc).GrhIndex).sY
                .Right = .Left + TilePixelWidth
                .Bottom = .Top + TilePixelHeight
            End With
           
            'Render the item grh and the amount
            InvSurface.SetForeColor vbWhite
            InvSurface.BltFast 0, 0, SurfaceDB.Surface(GrhData(UserInventory(loopc).GrhIndex).FileNum), SrcRect, DDBLTFAST_SRCCOLORKEY + DDBLTFAST_WAIT
            InvSurface.DrawText 0, 0, UserInventory(loopc).Amount, False
           
            'If equipped we render "E"
            If UserInventory(loopc).Equipped Then
                InvSurface.SetForeColor vbYellow
                InvSurface.DrawText 20, 20, "E", False
            End If
            


           
            'Render a box around the selected item
            If InvSelectedItem = loopc Then
                InvSurface.SetForeColor vbGreen
                InvSurface.setDrawStyle DrawStyleConstants.vbDot
                InvSurface.DrawBox 0, 0, TilePixelWidth, TilePixelHeight
            End If
        End If
       
        'Render the item to the Inventory Window
        'Call InvSurface.BltToDC(InventoryWindow.hdc, TempRect, destRect)
       
        Call InventoryWindow.Refresh
       
        With destRect
            'Check if we have reached the right end of the inventory picturebox
            If .Left = (InventoryWindow.ScaleWidth - (InventoryWindow.ScaleWidth Mod TilePixelWidth) - TilePixelWidth) Then
                .Left = 0
                .Right = TilePixelWidth
                .Bottom = .Bottom + TilePixelHeight
                .Top = .Top + TilePixelHeight
            Else
                .Left = .Left + TilePixelWidth
                .Right = .Right + TilePixelWidth
            End If
        End With
    Next loopc
End Sub
 
Private Sub DrawInventorySlot(ByVal slot As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 12/04/06
'Renders a inventory slot to the given PictureBox
'***************************************************
    Dim loopc As Long
    Dim SrcRect As RECT
    Dim TempRect As RECT
    Dim destRect As RECT
   
    'Slot
    With TempRect
        .Bottom = TilePixelHeight
        .Right = TilePixelWidth
    End With
   
    Dim upperLeftSlot As Integer
    upperLeftSlot = InventoryOffset * (InventoryWindow.ScaleWidth \ TilePixelWidth) + 1
   
    'If not in renderable area we exit
    If slot < upperLeftSlot Then Exit Sub
   
    'Fondo
    With destRect
        .Top = ((slot - upperLeftSlot) \ (InventoryWindow.ScaleWidth \ TilePixelWidth)) * TilePixelHeight
        .Left = ((slot - 1) Mod (InventoryWindow.ScaleWidth \ TilePixelWidth)) * TilePixelWidth
        .Bottom = .Top + TilePixelHeight
        .Right = .Left + TilePixelWidth
    End With
   
    'Clear the slot area
    InvSurface.BltFast 0, 0, SurfaceDB.Surface(123123), destRect, DDBLTFAST_SRCCOLORKEY + DDBLTFAST_WAIT
   
    If UserInventory(slot).GrhIndex Then
        'Get source rect
        With SrcRect
            .Left = GrhData(UserInventory(slot).GrhIndex).sX
            .Top = GrhData(UserInventory(slot).GrhIndex).sY
            .Right = .Left + TilePixelWidth
            .Bottom = .Top + TilePixelHeight
        End With
       
        'Render the item grh and the amount
        InvSurface.SetForeColor vbWhite
        InvSurface.BltFast 0, 0, SurfaceDB.Surface(CLng(GrhData(UserInventory(slot).GrhIndex).FileNum)), SrcRect, DDBLTFAST_SRCCOLORKEY + DDBLTFAST_WAIT
       
        'Render a box around the selected item
        If InvSelectedItem = slot Then
            EfectoInventario 0, 0, slot
        End If
       
        InvSurface.DrawText 0, 0, UserInventory(slot).Amount, False
       
       
        'If equipped we render "E"
        If UserInventory(slot).Equipped Then
            InvSurface.SetForeColor vbYellow
            InvSurface.DrawText 20, 20, "E", False
        End If
        
 If UserInventory(slot).PuedeUsar = False Then
EfectNoUsa 0, 0, slot
            End If
       
    End If
   
    'Render the item to the Inventory Window
    InvSurface.BltToDC InventoryWindow.hDC, TempRect, destRect
   
    Call InventoryWindow.Refresh
End Sub
 
Public Sub Initialize(ByRef InvPic As PictureBox)
   
    'HAY QUE INICIAR EL DX8 desde aqui, como en la CARGA de graficos!!! y sonido!!
   
    Set InventoryWindow = InvPic
   
    'Make sure auto-redraw is set to true
    InventoryWindow.AutoRedraw = True
   
    'Set apropiate scale (pixel)
    InventoryWindow.ScaleMode = 3
   
    'initialize DX stuff
    Dim SurfaceDesc As DDSURFACEDESC2
   
    'Make sure DirectDraw was correctly initialized
    If DirectDraw Is Nothing Then Exit Sub
   
    'Set surface's description
    With SurfaceDesc
        .lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
        .lHeight = TilePixelHeight
        .lWidth = TilePixelWidth
    End With
 
    'Create inventory surface
    Set InvSurface = DirectDraw.CreateSurface(SurfaceDesc)
   
    'Prepare surfaces for text rendering
    InvSurface.SetFontTransparency True
    'TODO : Fonts should be in a separate class / collection
    Dim font As New StdFont
    Dim Ifnt As IFont
   
    font.Name = "Verdana"
    font.Bold = True
    font.Italic = False
    font.Size = 7
    font.Underline = False
    font.Strikethrough = False
   
    Set Ifnt = font
   
    InvSurface.SetFont Ifnt
'\TODO
   
    InvSelectedItem = ClickItem(1, 1)   'If there is anything there we select the top left item
 
End Sub
Public Sub InventoryWindow_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
'***************************************************
'Implements the mouse move event of the inventory picture box
'Displays a ToolTip of the item under the mouse
'***************************************************
    Dim temp_x As Integer
    Dim temp_y As Integer
    Dim TempItem As Integer
   
    temp_x = X \ 32 'TilePixelWidth
    temp_y = Y \ 32 'TilePixelHeight
   
    TempItem = temp_x + (temp_y + InventoryOffset) * (InventoryWindow.ScaleWidth \ 32) + 1
   
    If TempItem <= MAX_INVENTORY_SLOTS Then
        InventoryWindow.ToolTipText = UserInventory(TempItem).Name
        frmMain.ItemName.Caption = UserInventory(TempItem).Name
    End If
 
End Sub
Private Sub InventoryWindow_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Check outs which item was clicked
   
    'Store previously selected item
    Dim prevSelItem As Long
   
    prevSelItem = InvSelectedItem
 
    'Get the currently clickced item
    InvSelectedItem = ClickItem(CInt(X), CInt(Y))
    
      If InvSelectedItem <> MouseDownSelectedItem And MouseDownSelectedItem <> 0 And InvSelectedItem <> 0 Then
        Call SendData("SWAP" & InvSelectedItem & "," & MouseDownSelectedItem)
    End If
 
   
    'Update needed inventory slots
    If prevSelItem <> InvSelectedItem Then
        If prevSelItem <> 0 And prevSelItem <> FLAGORO Then Call DrawInventorySlot(prevSelItem)
        If InvSelectedItem Then Call DrawInventorySlot(InvSelectedItem)
    End If
   
End Sub
Private Sub EfectoInventario(ByVal X As Integer, ByVal Y As Integer, ByVal slot As Integer)
'*****************************************************************
'Draw the GRH with a tonality, in the position X and Y of the Inventory
'*****************************************************************
    Dim CurrentGrhIndex As Integer
    Dim SourceRect As RECT
 
        With SourceRect
            .Left = GrhData(UserInventory(slot).GrhIndex).sX
            .Top = GrhData(UserInventory(slot).GrhIndex).sY
            .Right = .Left + 32
            .Bottom = .Top + 32
        End With
 
        Dim upperLeftSlot As Integer
    upperLeftSlot = InventoryOffset * (InventoryWindow.ScaleWidth \ TilePixelWidth) + 1
 
   
    Dim src As DirectDrawSurface7
    Dim rDest As RECT
    Dim dArray() As Byte, sArray() As Byte
    Dim ddsdSrc As DDSURFACEDESC2, ddsdDest As DDSURFACEDESC2
    Dim Modo As Long
   
    Set src = SurfaceDB.Surface(CLng(GrhData(UserInventory(slot).GrhIndex).FileNum))
   
    src.GetSurfaceDesc ddsdSrc
    InvSurface.GetSurfaceDesc ddsdDest
   
    With rDest
        .Top = ((slot - upperLeftSlot) \ (InventoryWindow.ScaleWidth \ TilePixelWidth)) * TilePixelHeight
        .Left = ((slot - 1) Mod (InventoryWindow.ScaleWidth \ TilePixelWidth)) * TilePixelWidth
        .Bottom = .Top + TilePixelHeight
        .Right = .Left + TilePixelWidth
    End With
 
   
    Dim SrcLock As Boolean
    Dim DstLock As Boolean
   
    SrcLock = False
    DstLock = False
   
On Local Error GoTo HayErrorAlpha
   
    Call src.Lock(SourceRect, ddsdSrc, DDLOCK_WAIT, 0)
    SrcLock = True
    Call InvSurface.Lock(rDest, ddsdDest, DDLOCK_WAIT, 0)
    DstLock = True
   
    Call InvSurface.GetLockedArray(dArray())
    Call src.GetLockedArray(sArray())
   
   
If ddsdDest.ddpfPixelFormat.lGBitMask = &H3E0 Then
  Modo = 555
ElseIf ddsdDest.ddpfPixelFormat.lGBitMask = &H7E0 Then
  Modo = 565
Else
  MsgBox "Modo de vídeo no esta en 555 o 565 o algo falló."
  End
End If
 
'Alpha
    If Modo = 565 Then
        Call vbDABLcolorblend16565ck(ByVal VarPtr(sArray(SourceRect.Left * 2, SourceRect.Top)), ByVal VarPtr(dArray(X + X, Y)), 120, rDest.Right - rDest.Left, rDest.Bottom - rDest.Top, ddsdSrc.lPitch, ddsdDest.lPitch, 255, 230, 138)
    Else
        Call vbDABLcolorblend16555ck(ByVal VarPtr(sArray(SourceRect.Left * 2, SourceRect.Top)), ByVal VarPtr(dArray(X + X, Y)), 120, rDest.Right - rDest.Left, rDest.Bottom - rDest.Top, ddsdSrc.lPitch, ddsdDest.lPitch, 255, 230, 138)
    End If
    InvSurface.Unlock rDest
    DstLock = False
    src.Unlock SourceRect
    SrcLock = False
Exit Sub
 
HayErrorAlpha:
    If SrcLock Then src.Unlock SourceRect
    If DstLock Then InvSurface.Unlock rDest
End Sub

Private Sub EfectNoUsa(ByVal X As Integer, ByVal Y As Integer, ByVal slot As Integer)
'*****************************************************************
'Draw the GRH with a tonality, in the position X and Y of the Inventory
'*****************************************************************
    Dim CurrentGrhIndex As Integer
    Dim SourceRect As RECT
 
        With SourceRect
            .Left = GrhData(UserInventory(slot).GrhIndex).sX
            .Top = GrhData(UserInventory(slot).GrhIndex).sY
            .Right = .Left + 32
            .Bottom = .Top + 32
        End With
 
        Dim upperLeftSlot As Integer
    upperLeftSlot = InventoryOffset * (InventoryWindow.ScaleWidth \ TilePixelWidth) + 1
 
   
    Dim src As DirectDrawSurface7
    Dim rDest As RECT
    Dim dArray() As Byte, sArray() As Byte
    Dim ddsdSrc As DDSURFACEDESC2, ddsdDest As DDSURFACEDESC2
    Dim Modo As Long
   
    Set src = SurfaceDB.Surface(CLng(GrhData(UserInventory(slot).GrhIndex).FileNum))
   
    src.GetSurfaceDesc ddsdSrc
    InvSurface.GetSurfaceDesc ddsdDest
   
    With rDest
        .Top = ((slot - upperLeftSlot) \ (InventoryWindow.ScaleWidth \ TilePixelWidth)) * TilePixelHeight
        .Left = ((slot - 1) Mod (InventoryWindow.ScaleWidth \ TilePixelWidth)) * TilePixelWidth
        .Bottom = .Top + TilePixelHeight
        .Right = .Left + TilePixelWidth
    End With
 
   
    Dim SrcLock As Boolean
    Dim DstLock As Boolean
   
    SrcLock = False
    DstLock = False
   
On Local Error GoTo HayErrorAlpha
   
    Call src.Lock(SourceRect, ddsdSrc, DDLOCK_WAIT, 0)
    SrcLock = True
    Call InvSurface.Lock(rDest, ddsdDest, DDLOCK_WAIT, 0)
    DstLock = True
   
    Call InvSurface.GetLockedArray(dArray())
    Call src.GetLockedArray(sArray())
   
   
If ddsdDest.ddpfPixelFormat.lGBitMask = &H3E0 Then
  Modo = 555
ElseIf ddsdDest.ddpfPixelFormat.lGBitMask = &H7E0 Then
  Modo = 565
Else
  MsgBox "Modo de vídeo no esta en 555 o 565 o algo falló."
  End
End If
 
'Alpha
    If Modo = 565 Then
        Call vbDABLcolorblend16565ck(ByVal VarPtr(sArray(SourceRect.Left * 2, SourceRect.Top)), ByVal VarPtr(dArray(X + X, Y)), 120, rDest.Right - rDest.Left, rDest.Bottom - rDest.Top, ddsdSrc.lPitch, ddsdDest.lPitch, 250, 0, 0)
    Else
        Call vbDABLcolorblend16555ck(ByVal VarPtr(sArray(SourceRect.Left * 2, SourceRect.Top)), ByVal VarPtr(dArray(X + X, Y)), 120, rDest.Right - rDest.Left, rDest.Bottom - rDest.Top, ddsdSrc.lPitch, ddsdDest.lPitch, 250, 0, 0)
    End If
    InvSurface.Unlock rDest
    DstLock = False
    src.Unlock SourceRect
    SrcLock = False
Exit Sub
 
HayErrorAlpha:
    If SrcLock Then src.Unlock SourceRect
    If DstLock Then InvSurface.Unlock rDest
End Sub

 Private Sub InventoryWindow_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim TempItem As Long
 
 
    If Button = 2 Then 'click derecho
        TempItem = ClickItem(X, Y)
       
        If UserInventory(TempItem).GrhIndex Then
            MouseDownSelectedItem = TempItem
            Exit Sub
        End If
    End If
   
    MouseDownSelectedItem = 0
End Sub
