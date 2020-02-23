VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Base 0
Private Type tLight
    RGBLight As D3DCOLORVALUE
    LightAc As Boolean
    Xt As Byte
    Yt As Byte
    Rango As Byte
    LightLightId As Long
End Type
 
Private LightList() As tLight
Private LightCount As Byte
Dim VarC As Long
Dim VarL As Long
Public Function Create_Light_To_Map(ByVal Xt As Byte, ByVal Yt As Byte, Optional Rango As Byte = 3, Optional ByVal red As Byte = 255, Optional ByVal green = 255, Optional ByVal blue As Byte = 255)
    LightCount = LightCount + 1
   
    ReDim Preserve LightList(1 To LightCount) As tLight
   
    'Le damos color
    LightList(LightCount).RGBLight.r = red
    LightList(LightCount).RGBLight.g = green
    LightList(LightCount).RGBLight.b = blue
   
    'Alpha (Si borras esto RE KB!!)
    LightList(LightCount).RGBLight.A = 255
   
    'Un rango determinado
    LightList(LightCount).Rango = Rango
   
    'La activamos para que se pueda renderizar
    LightList(LightCount).LightAc = True
   
    'Pos en el mapa
    LightList(LightCount).Xt = Xt
    LightList(LightCount).Yt = Yt
   
    Call LightRender(LightCount)
End Function
Public Function Delete_Light_To_Map(ByVal X As Byte, ByVal Y As Byte)
   
    Dim i As Long
   
    For i = 1 To LightCount
        If LightList(i).Xt = X And LightList(i).Yt = Y Then
            Delete_Light_To_Index i
            Exit Function
        End If
    Next i
 
End Function
Public Function Delete_Light_To_Index(ByVal light_index As Byte)
   
    Dim min_x As Integer
    Dim min_y As Integer
    Dim max_x As Integer
    Dim max_y As Integer
    Dim Ya As Integer
    Dim Xa As Integer
 
    LightList(light_index).LightAc = False
 
    min_x = LightList(light_index).Xt - LightList(light_index).Rango
    max_x = LightList(light_index).Xt + LightList(light_index).Rango
    min_y = LightList(light_index).Yt - LightList(light_index).Rango
    max_y = LightList(light_index).Yt + LightList(light_index).Rango
       
    For Ya = min_y To max_y
        For Xa = min_x To max_x
            If InMapBounds(Xa, Ya) Then
                MapData(Xa, Ya).light_value(0) = 0
                MapData(Xa, Ya).light_value(1) = 0
                MapData(Xa, Ya).light_value(2) = 0
                MapData(Xa, Ya).light_value(3) = 0
            End If
        Next Xa
    Next Ya
   
End Function
Private Function LightCalculate(ByVal cRadio As Integer, ByVal LightX As Integer, ByVal LightY As Integer, ByVal XCoord As Integer, ByVal YCoord As Integer, TileLight As Long, LightColor As D3DCOLORVALUE, AmbientColor As D3DCOLORVALUE) As Long
    Dim XDist As Single
    Dim YDist As Single
    Dim VertexDist As Single
    Dim pRadio As Integer
   
    Dim CurrentColor As D3DCOLORVALUE
   
    pRadio = cRadio * 32
   
    XDist = LightX + 16 - XCoord
    YDist = LightY + 16 - YCoord
   
    VertexDist = Sqr(XDist * XDist + YDist * YDist)
   
    If VertexDist <= pRadio Then
        Call D3DXColorLerp(CurrentColor, LightColor, AmbientColor, VertexDist / pRadio)
        LightCalculate = D3DColorXRGB(Round(CurrentColor.r), Round(CurrentColor.g), Round(CurrentColor.b))
        'If TileLight > LightCalculate Then LightCalculate = TileLight
    Else
        LightCalculate = TileLight
    End If
End Function
 
Private Sub LightRender(ByVal light_index As Integer)
 
    If light_index = 0 Then Exit Sub
    If LightList(light_index).LightAc = False Then Exit Sub
   
    Dim min_x As Integer
    Dim min_y As Integer
    Dim max_x As Integer
    Dim max_y As Integer
    Dim color As Long
    Dim Ya As Integer
    Dim Xa As Integer
   
    Dim TileLight As D3DCOLORVALUE
    Dim AmbientColor As D3DCOLORVALUE
    Dim LightColor As D3DCOLORVALUE
   
    Dim XCoord As Integer
    Dim YCoord As Integer
   
    AmbientColor.r = 255
    AmbientColor.g = 255
    AmbientColor.b = 255
   
    LightColor = LightList(light_index).RGBLight
       
    min_x = LightList(light_index).Xt - LightList(light_index).Rango
    max_x = LightList(light_index).Xt + LightList(light_index).Rango
    min_y = LightList(light_index).Yt - LightList(light_index).Rango
    max_y = LightList(light_index).Yt + LightList(light_index).Rango
       
    For Ya = min_y To max_y
        For Xa = min_x To max_x
            If InMapBounds(Xa, Ya) Then
                XCoord = Xa * 32
                YCoord = Ya * 32
                MapData(Xa, Ya).light_value(1) = LightCalculate(LightList(light_index).Rango, LightList(light_index).Xt * 32, LightList(light_index).Yt * 32, XCoord, YCoord, MapData(Xa, Ya).light_value(1), LightColor, AmbientColor)
 
                XCoord = Xa * 32 + 32
                YCoord = Ya * 32
                MapData(Xa, Ya).light_value(3) = LightCalculate(LightList(light_index).Rango, LightList(light_index).Xt * 32, LightList(light_index).Yt * 32, XCoord, YCoord, MapData(Xa, Ya).light_value(3), LightColor, AmbientColor)
                       
                XCoord = Xa * 32
                YCoord = Ya * 32 + 32
                MapData(Xa, Ya).light_value(0) = LightCalculate(LightList(light_index).Rango, LightList(light_index).Xt * 32, LightList(light_index).Yt * 32, XCoord, YCoord, MapData(Xa, Ya).light_value(0), LightColor, AmbientColor)
   
                XCoord = Xa * 32 + 32
                YCoord = Ya * 32 + 32
                MapData(Xa, Ya).light_value(2) = LightCalculate(LightList(light_index).Rango, LightList(light_index).Xt * 32, LightList(light_index).Yt * 32, XCoord, YCoord, MapData(Xa, Ya).light_value(2), LightColor, AmbientColor)
               
            End If
        Next Xa
    Next Ya
End Sub
 
Public Function Light_Remove(ByVal light_index As Long) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'*****************************************************************
    'Make sure it's a legal index
    If Light_Check(light_index) Then
        Light_Destroy light_index
        Light_Remove = True
    End If
End Function
 
Public Function Light_Create(ByVal Xt As Integer, ByVal Yt As Integer, _
                            Optional ByVal Rango As Byte = 1, Optional ByVal LightId As Long, Optional ByVal red As Byte = 255, Optional ByVal green = 255, Optional ByVal blue As Byte = 255) As Long
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Returns the light_index if successful, else 0
'Edited by Juan Martín Sotuyo Dodero
'**************************************************************
    If InMapBounds(Xt, Yt) Then
        'Make sure there is no light in the given map pos
        'If Map_Light_Get(Xt, Yt) <> 0 Then
        '    Light_Create = 0
        '    Exit Function
        'End If
        Light_Create = Light_Next_Open
        Light_Make Light_Create, Xt, Yt, Rango, LightId, red, green, blue
    End If
End Function
 
Public Function Light_Move(ByVal light_index As Long, ByVal Xt As Integer, ByVal Yt As Integer) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Returns true if successful, else false
'**************************************************************
    'Make sure it's a legal CharIndex
    If Light_Check(light_index) Then
        'Make sure it's a legal move
        If InMapBounds(Xt, Yt) Then
       
            'Move it
            Light_Erase light_index
            LightList(light_index).Xt = Xt
            LightList(light_index).Yt = Yt
   
            Light_Move = True
           
        End If
    End If
End Function
 
Public Function Light_Move_By_Head(ByVal light_index As Long, ByVal Heading As Byte) As Boolean
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 15/05/2002
'Returns true if successful, else false
'**************************************************************
    Dim Xt As Integer
    Dim Yt As Integer
    Dim nX As Integer
    Dim nY As Integer
    Dim addy As Byte
    Dim addx As Byte
    'Check for valLightId heading
    If Heading < 1 Or Heading > 8 Then
        Light_Move_By_Head = False
        Exit Function
    End If
 
    'Make sure it's a legal CharIndex
    If Light_Check(light_index) Then
   
        Xt = LightList(light_index).Xt
        Yt = LightList(light_index).Yt
       
 
 
        Select Case Heading
            Case NORTH
                addy = -1
       
            Case EAST
                addx = 1
       
            Case SOUTH
                addy = 1
           
            Case WEST
                addx = -1
        End Select
       
        nX = Xt + addx
        nY = Yt + addy
       
        'Make sure it's a legal move
        If InMapBounds(nX, nY) Then
       
            'Move it
            Light_Erase light_index
 
            LightList(light_index).Xt = nX
            LightList(light_index).Yt = nY
   
            Light_Move_By_Head = True
           
        End If
    End If
End Function
 
Private Sub Light_Make(ByVal light_index As Long, ByVal Xt As Integer, ByVal Yt As Integer, _
                        ByVal Rango As Long, Optional ByVal LightId As Long, Optional ByVal red As Byte = 255, Optional ByVal green = 255, Optional ByVal blue As Byte = 255)
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'*****************************************************************
    'Update array size
    If light_index > VarL Then
        VarL = light_index
        ReDim Preserve LightList(1 To VarL)
    End If
    VarC = VarC + 1
   
    'Make LightAc
    LightList(light_index).LightAc = True
   
        'Le damos color
    LightList(light_index).RGBLight.r = red
    LightList(light_index).RGBLight.g = green
    LightList(light_index).RGBLight.b = blue
   
    'Alpha (Si borras esto RE KB!!)
    LightList(light_index).RGBLight.A = 255
   
    LightList(light_index).Xt = Xt
    LightList(light_index).Yt = Yt
    LightList(light_index).Rango = Rango
    'LightList(light_index).LightId = LightId
End Sub
 
Private Function Light_Check(ByVal light_index As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'**************************************************************
    'check light_index
    If light_index > 0 And light_index <= VarL Then
        If LightList(light_index).LightAc Then
            Light_Check = True
        End If
    End If
End Function
 
Public Sub Light_Render_All()
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Dim loop_counter As Long
           
    For loop_counter = 1 To VarC
       
        If LightList(loop_counter).LightAc Then
            LightRender loop_counter
        End If
   
    Next loop_counter
End Sub
 
Private Function Light_Next_Open() As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim loopc As Long
    
    If VarL + 1 = 0 Then
        Light_Next_Open = 1
        Exit Function
    End If
    
    loopc = 1
    Do Until LightList(loopc).LightAc = False
        If loopc = VarL Then
            Light_Next_Open = VarL + 1 + 1
            Exit Function
        End If
        loopc = loopc + 1
    Loop
    
    Light_Next_Open = loopc

Exit Function

ErrorHandler:
Light_Next_Open = 1
End Function
 
Public Function Light_Find(ByVal LightId As Long) As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'Find the index related to the handle
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim loopc As Long
   
    loopc = 1
    'Do Until LightList(loopc).LightId = LightId
        If loopc = VarL Then
            Light_Find = 0
            Exit Function
        End If
        loopc = loopc + 1
    'Loop
   
    Light_Find = loopc
Exit Function
ErrorHandler:
    Light_Find = 0
End Function
 
Public Function Light_Remove_All() As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'*****************************************************************
    Dim Index As Long
   
    For Index = 1 To VarL
        'Make sure it's a legal index
        If Light_Check(Index) Then
            Light_Destroy Index
        End If
    Next Index
   
    Light_Remove_All = True
End Function
 
Private Sub Light_Destroy(ByVal light_index As Long)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Dim temp As tLight
   
    Light_Erase light_index
   
    LightList(light_index) = temp
   
    'Update array size
    If light_index = VarL Then
        Do Until LightList(VarL).LightAc
            VarL = VarL - 1
            If VarL = 0 Then
                VarC = 0
                Exit Sub
            End If
        Loop
        ReDim Preserve LightList(1 To VarL)
    End If
    VarC = VarC - 1
End Sub
 
Private Sub Light_Erase(ByVal light_index As Long)
'***************************************'
'Author: Juan Martín Sotuyo Dodero
'Last modified: 3/31/2003
'Correctly erases a light
'***************************************'
    Dim min_x As Integer
    Dim min_y As Integer
    Dim max_x As Integer
    Dim max_y As Integer
    Dim X As Integer
    Dim Y As Integer
   
    'Set up light borders
    min_x = LightList(light_index).Xt - LightList(light_index).Rango
    min_y = LightList(light_index).Yt - LightList(light_index).Rango
    max_x = LightList(light_index).Xt + LightList(light_index).Rango
    max_y = LightList(light_index).Yt + LightList(light_index).Rango
   
    'ArRango corners
    'NE
    If InMapBounds(min_x, min_y) Then
        MapData(min_x, min_y).light_value(2) = 0
    End If
    'NW
    If InMapBounds(max_x, min_y) Then
        MapData(max_x, min_y).light_value(0) = 0
    End If
    'SW
    If InMapBounds(max_x, max_y) Then
        MapData(max_x, max_y).light_value(1) = 0
    End If
    'SE
    If InMapBounds(min_x, max_y) Then
        MapData(min_x, max_y).light_value(3) = 0
    End If
   
    'ArRango borders
    'Upper border
    For X = min_x + 1 To max_x - 1
        If InMapBounds(X, min_y) Then
            MapData(X, min_y).light_value(0) = 0
            MapData(X, min_y).light_value(2) = 0
        End If
    Next X
   
    'Lower border
    For X = min_x + 1 To max_x - 1
        If InMapBounds(X, max_y) Then
            MapData(X, max_y).light_value(1) = 0
            MapData(X, max_y).light_value(3) = 0
        End If
    Next X
   
    'Left border
    For Y = min_y + 1 To max_y - 1
        If InMapBounds(min_x, Y) Then
            MapData(min_x, Y).light_value(2) = 0
            MapData(min_x, Y).light_value(3) = 0
        End If
    Next Y
   
    'Right border
    For Y = min_y + 1 To max_y - 1
        If InMapBounds(max_x, Y) Then
            MapData(max_x, Y).light_value(0) = 0
            MapData(max_x, Y).light_value(1) = 0
        End If
    Next Y
   
    'Set the inner part of the light
    For X = min_x + 1 To max_x - 1
        For Y = min_y + 1 To max_y - 1
            If InMapBounds(X, Y) Then
                MapData(X, Y).light_value(0) = 0
                MapData(X, Y).light_value(1) = 0
                MapData(X, Y).light_value(2) = 0
                MapData(X, Y).light_value(3) = 0
            End If
        Next Y
    Next X
End Sub
 



