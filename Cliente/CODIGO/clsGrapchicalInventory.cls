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

Private InventoryOffset As Long             'Number of lines we scrolled down from topmost
Private InvSelectedItem As Long             'Currently selected item
Private MouseDownSelectedItem As Long

Private ownTilePixelWidth As Integer
Private ownTilePixelHeight As Integer
Private startdX As Integer
Private startdY As Integer

Private ShowValue As Boolean
Private ImgContainer As Boolean

Private Inventory() As Inventory    'User's inventory

Private WithEvents InventoryWindow As PictureBox    'Placeholder where to render the inventory
Attribute InventoryWindow.VB_VarHelpID = -1

Public Sub ScrollInventory(ByVal Up As Boolean)
    'Only move if there is something to show up /down
    If CanScroll(Up) Then
        InventoryOffset = InventoryOffset + IIf(Up, 1, -1)
    End If

End Sub

Public Function CanScroll(ByVal Up As Boolean)
    If Up Then
        CanScroll = (InventoryOffset + (InventoryWindow.ScaleHeight \ ownTilePixelHeight)) * _
                    (InventoryWindow.ScaleWidth \ ownTilePixelWidth) < UBound(Inventory)
    Else
        CanScroll = InventoryOffset > 0
    End If
End Function

Public Sub SelectGold()
    Dim prevSelectedItem As Long
    
    'Store preivously selected item
    prevSelectedItem = InvSelectedItem
    
    'Select the gold
    InvSelectedItem = FLAGORO
End Sub

Public Sub DeselectItem()
'***************************************************
'Author: ZaMa
'Last Modify Date: 30/11/2009
'Deselect the currently selected item
'***************************************************
    If InvSelectedItem = 0 Then Exit Sub
    
    Dim ClearSlot As Byte
    
    ClearSlot = InvSelectedItem
    
    'Select nothing
    InvSelectedItem = 0
    
    'Redraw old item to deselect it
    'Call DrawInventory
End Sub

Public Sub ChangeSlotItemAmount(ByVal slot As Byte, ByVal Amount As Long)
'***************************************************
'Author: ZaMa
'Created Date: 30/11/2009
'Last Modify Date: 09/12/2009 (Budi)
'Change the amount of the required item, and update the slot.
'***************************************************
    
    Inventory(slot).Amount = Amount
    
    'Update Amount
    'Call DrawInventory
End Sub

' Retrieves the selected item index
Public Property Get SelectedItem() As Byte: SelectedItem = InvSelectedItem: End Property
' Retrieves the max hit of the selected item
Public Property Get MaxHit(ByVal slot As Byte) As Integer: MaxHit = Inventory(slot).MaxHit: End Property
' Retrieves the min hit of the selected item
Public Property Get MinHit(ByVal slot As Byte) As Integer: MinHit = Inventory(slot).MinHit: End Property
' Retrieves the defense of the selected item
Public Property Get Def(ByVal slot As Byte) As Integer: Def = Inventory(slot).Def: End Property
' Retrieves the grh index of the selected item
Public Property Get GrhIndex(ByVal slot As Byte) As Integer: GrhIndex = Inventory(slot).GrhIndex: End Property
' Retrieves the value of the selected item
Public Property Get Valor(ByVal slot As Byte) As Long: Valor = Inventory(slot).Valor: End Property
' Retrieves the selected item's amount
Public Property Get Amount(ByVal slot As Byte) As Integer: Amount = Inventory(slot).Amount: End Property
' Retrieves the selected item's object index
Public Property Get OBJIndex(ByVal slot As Byte) As Integer: OBJIndex = Inventory(slot).OBJIndex: End Property
' Retrieves the selected item's object type
Public Property Get OBJType(ByVal slot As Byte) As Integer: OBJType = Inventory(slot).OBJType: End Property
' Retrieves the selected item's name
Public Property Get ItemName(ByVal slot As Byte) As String: ItemName = Inventory(slot).name: End Property
' Retrieves True if the item at the given pos is eqiupped
Public Property Get Equipped(ByVal slot As Byte) As Boolean: Equipped = Inventory(slot).Equipped: End Property

Public Property Get MaxObjs() As Byte: MaxObjs = UBound(Inventory): End Property
   
Public Sub SetMaxObjs(ByVal MaxObjs As Byte)
'***************************************************
'Author: Torres Patricio (Pato)
'Last Modify Date: 09/16/09
'Set the capacity of the Inventary
'***************************************************
    If UBound(Inventory) = MaxObjs Then Exit Sub
    
    ReDim Preserve Inventory(1 To MaxObjs) As Inventory
End Sub

Public Sub ToggleShowValue(ByVal bShow As Boolean)
'***************************************************
'Author: ZaMa
'Last Modify Date: 11/12
'Indicates if the amount of the items in the inventory is shown
'***************************************************
    ShowValue = bShow
    
    Dim slot As Integer
    ' Update Inventory
    For slot = 1 To Me.MaxObjs
       ' DrawInventory
    Next slot
End Sub

Public Sub SetItem(ByVal slot As Byte, ByVal eOBJIndex As Integer, ByVal eAmount As Long, ByVal eEquipped As Byte, _
                        ByVal eGrhIndex As Integer, ByVal eObjType As Integer, ByVal eMaxHit As Integer, ByVal eMinHit As Integer, _
                        ByVal Def As Integer, ByVal eValor As Single, ByVal eName As String, Optional ByVal PuedeUsar As Byte = 1)
    If slot < 1 Or slot > UBound(Inventory) Then Exit Sub
    
    With Inventory(slot)
        .Amount = eAmount
        .Def = Def
        .Equipped = eEquipped
        .GrhIndex = eGrhIndex
        .MaxHit = eMaxHit
        .MinHit = eMinHit
        .name = eName
        .OBJIndex = eOBJIndex
        .OBJType = eObjType
        .Valor = eValor
        .PuedeUsar = PuedeUsar
    End With
    
Call DrawInventory
End Sub

Private Function ClickItem(ByVal X As Long, ByVal Y As Long) As Long
    Dim TempItem As Long
    TempItem = X \ TilePixelWidth + (Y \ TilePixelHeight + InventoryOffset) * (InventoryWindow.ScaleWidth \ TilePixelWidth) + 1
   
        ClickItem = TempItem
        If Not ClickItem = 0 Then
            frmMain.ItemInfo.Caption = Inventory(TempItem).name
        Else
            frmMain.ItemInfo.Caption = "(nada)"
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
        .bottom = InventoryWindow.Height
        .Right = InventoryWindow.Width
    End With
    
    destRect = TempRect
    
    'Clear the inventory window
    DirectDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0
    DirectDevice.BeginScene
    
        For loopc = InventoryOffset * (InventoryWindow.ScaleWidth \ ownTilePixelWidth) + 1 To UBound(Inventory)
            
            startdX = 5 + ((loopc - 1) Mod (frmMain.picInv.Width / 32)) * 32
            startdY = ((loopc - 1) \ (frmMain.picInv.Width / 32)) * 32
            
            If Inventory(loopc).GrhIndex Then
            
                'Get source rect
                With SrcRect
                    .Left = GrhData(Inventory(loopc).GrhIndex).SX
                    .Top = GrhData(Inventory(loopc).GrhIndex).SY
                    .Right = .Left + ownTilePixelWidth
                    .bottom = .Top + ownTilePixelHeight
                End With
                
                'Renderizado de los items y Seleccionados
                If InvSelectedItem = loopc Then
                    If ImgContainer Then
                        Draw_FillBox startdX, startdY + 2, 30, 30, D3DColorXRGB(0, 0, 1), -65536
                        If Inventory(loopc).PuedeUsar = False Then
                            Device_Textured_Render startdX, startdY, SurfaceDB.Surface(CLng(GrhData(Inventory(loopc).GrhIndex).FileNum)), SrcRect, NoPuedeUsar
                        Else
                            Device_Textured_Render startdX, startdY, SurfaceDB.Surface(CLng(GrhData(Inventory(loopc).GrhIndex).FileNum)), SrcRect, LightIluminado
                        End If
                    End If
                Else
                    If Inventory(loopc).PuedeUsar = False Then
                        Device_Textured_Render startdX, startdY, SurfaceDB.Surface(CLng(GrhData(Inventory(loopc).GrhIndex).FileNum)), SrcRect, NoPuedeUsar
                    Else
                        Device_Textured_Render startdX, startdY, SurfaceDB.Surface(CLng(GrhData(Inventory(loopc).GrhIndex).FileNum)), SrcRect, LightOscurito
                    End If
                End If
                            
                'Cuando esta equipado "+"
                If Inventory(loopc).Equipped Then
                    Fonts_Render_String "+", startdX + 20, startdY - 2, -65536
                End If
                
                'Cantidad de items
                Fonts_Render_String Inventory(loopc).Amount, startdX - 2, startdY + 20, -1
                
            End If
        Next loopc
    DirectDevice.EndScene
    DirectDevice.Present destRect, ByVal 0, InventoryWindow.hwnd, ByVal 0
    
End Sub

Public Sub Initialize(ByRef DirectD3D As D3DX8, ByRef InvPic As PictureBox, ByVal MaxObjs As Byte, _
                            Optional ByVal TileWidth As Integer = 32, _
                            Optional ByVal TileHeight As Integer = 32, Optional ByVal startX As Integer = 0, _
                            Optional ByVal startY As Integer = 0, Optional ByVal bImgContainer As Boolean = True, _
                            Optional ByVal bShowText As Boolean = True)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 03/12/09
'Sets the reference to the picture box on which to render the inventory
'03/12/09: I added the optionals args FontSize, TileWidth, TileHeight, startX and startY. (Budi)
'***************************************************
    Set InventoryWindow = InvPic
    
    'Make sure auto-redraw is set to true
    InventoryWindow.AutoRedraw = True
    
    'Set apropiate scale (pixel)
    InventoryWindow.ScaleMode = 3
    
    ReDim Inventory(1 To MaxObjs) As Inventory
    
    'initialize DX stuff
    'Dim SurfaceDesc As DDSURFACEDESC2
    
    'Make sure DirectD3D was correctly initialized
    If DirectD3D Is Nothing Then Exit Sub

    ownTilePixelHeight = TileHeight
    ownTilePixelWidth = TileWidth
    startdX = startX
    startdY = startY
    
    ImgContainer = bImgContainer
    ShowValue = bShowText
    
'\TODO
    
    If ImgContainer Then _
        InvSelectedItem = ClickItem(1, 1)   'If there is anything there we select the top left item

End Sub

Public Sub UpdateInventory()
'*****************************
'By Lorwik
'Actualizacion del inventario
'*****************************
Call DrawInventory
DoEvents
Call DrawInventory
DoEvents
End Sub

Private Sub InventoryWindow_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim TempItem As Integer
    
    'Exit if it got outside the control's area
    If X < 0 Or Y < 0 Or X > InventoryWindow.Width Or Y > InventoryWindow.Height Then _
        Exit Sub
    
    TempItem = X \ ownTilePixelWidth + (Y \ ownTilePixelHeight + InventoryOffset) * (InventoryWindow.ScaleWidth \ ownTilePixelWidth) + 1
    
    If TempItem <= UBound(Inventory) Then
        InventoryWindow.ToolTipText = Inventory(TempItem).name
    End If
End Sub

Private Sub InventoryWindow_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    'Store previously selected item
    Dim prevSelItem As Long
    
    'Exit if it got outside the control's area
    If X < 0 Or Y < 0 Or X > InventoryWindow.Width Or Y > InventoryWindow.Height Then _
        Exit Sub
    
    prevSelItem = InvSelectedItem

    'Get the currently clickced item
    InvSelectedItem = ClickItem(CInt(X), CInt(Y))
    
    If InvSelectedItem <> MouseDownSelectedItem And MouseDownSelectedItem <> 0 And InvSelectedItem <> 0 Then
        Call Protocol.WriteSwapObjects(InvSelectedItem, MouseDownSelectedItem)
    End If
    
    'Update needed inventory slots
    If prevSelItem <> InvSelectedItem Then
        If prevSelItem <> 0 And prevSelItem <> FLAGORO Or InvSelectedItem Then Call DrawInventory
    End If
End Sub

Private Sub InventoryWindow_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim TempItem As Long
 
    If Button = 2 Then 'click derecho
        TempItem = ClickItem(X, Y)
       
        If Inventory(TempItem).GrhIndex Then
            MouseDownSelectedItem = TempItem
            Exit Sub
        End If
    End If
   
    MouseDownSelectedItem = 0
End Sub
