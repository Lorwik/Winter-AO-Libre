Option Explicit

Public MiCuerpo As Integer, MiCabeza As Integer, MinCabeza As Integer, MaxCabeza As Integer
 'Old fashion BitBlt function
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long

Sub DibujaPJ(Surface As DirectDrawSurface7, Grh As Grh, ByVal X As Integer, ByVal Y As Integer, index As Integer)
On Error Resume Next
Dim r1           As RECT, r2 As RECT, auxr As RECT
Dim iGrhIndex As Integer
If Grh.GrhIndex <= 0 Then Exit Sub
iGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
 
With r1
    .Right = GrhData(iGrhIndex).pixelWidth
    .Bottom = GrhData(iGrhIndex).pixelHeight
End With
 
With r2
   .Left = GrhData(iGrhIndex).sX
   .Top = GrhData(iGrhIndex).sY
   .Right = .Left + GrhData(iGrhIndex).pixelWidth
   .Bottom = .Top + GrhData(iGrhIndex).pixelHeight
End With
With auxr
    .Left = 0
  .Top = 0
   .Right = 150
  .Bottom = 150
End With
 
Surface.BltFast X, Y, SurfaceDB.Surface(GrhData(iGrhIndex).FileNum), r2, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
Surface.BltToDC frmCuenta.PJ(index).hDC, auxr, auxr
 
frmCuenta.PJ(index).Refresh
 
End Sub
Sub dibujaban(Surface As DirectDrawSurface7, index As Integer)
 
Dim r2 As RECT, auxr As RECT
 
With r2
   .Left = 0
   .Top = 0
   .Right = 20
   .Bottom = 20
End With
 
With auxr
    .Left = 0
  .Top = 0
   .Right = 150
  .Bottom = 150
End With
 
Surface.SetFontTransparency True
Surface.SetForeColor vbRed
frmCuenta.font.Size = 15
Surface.SetFont frmMain.font
'Surface.BltFast x, y, Surface, r2, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
Surface.DrawText 6, 60, "Banned", False
Surface.BltToDC frmCuenta.PJ(index).hDC, auxr, auxr
 
End Sub
 
Sub dibujamuerto(Surface As DirectDrawSurface7, index As Integer)
 
Dim r2 As RECT, auxr As RECT
 
With r2
   .Left = 0
   .Top = 0
   .Right = 20
   .Bottom = 20
End With
 
With auxr
    .Left = 0
  .Top = 0
   .Right = 150
  .Bottom = 150
End With
 
Surface.SetFontTransparency True
Surface.SetForeColor vbWhite
frmCuenta.font.Size = 6
Surface.SetFont frmCuenta.font
'Surface.BltFast x, y, Surface, r2, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
Surface.DrawText 5, 10, "MUERTO", False
Surface.BltToDC frmCuenta.PJ(index).hDC, auxr, auxr
 
End Sub
 
Sub DibujarTodo(ByVal index As Integer, Body As Integer, Head As Integer, Casco As Integer, Shield As Integer, Weapon As Integer, Baned As Integer, nombre As String, LVL As Integer, Clase As String, muerto As Integer)
 
Dim Grh As Grh
Dim Pos As Integer
Dim loopc As Integer
Dim r As RECT
Dim r2 As RECT
 
Dim YBody As Integer
Dim YYY As Integer
Dim XBody As Integer
Dim BBody As Integer
 
 
With r2
    .Left = 0
  .Top = 0
   .Right = 150
  .Bottom = 150
End With
 
    BackBufferSurface.BltColorFill r, vbBlack
 
If Baned = 1 Then
    Call dibujaban(BackBufferSurface, index)
End If
 
frmCuenta.nombre(index).Caption = nombre
 
frmCuenta.Label1(index).font = frmMain.font
frmCuenta.Label1(index).font = frmMain.font
 
frmCuenta.Label1(index).Caption = LVL
frmCuenta.Label2(index).Caption = Clase
 
XBody = 12
YBody = 15
BBody = 17
 
If muerto = 1 Then
    Body = 8
    Head = 500
    'Arma = 2
    Shield = 2
    Weapon = 2
    XBody = 10
    YBody = 35
    BBody = 16
    Call dibujamuerto(BackBufferSurface, index)
End If
 
Grh = BodyData(Body).Walk(3)
   
Call DibujaPJ(BackBufferSurface, Grh, XBody, YBody, index)
 
If muerto = 0 Then YYY = BodyData(Body).HeadOffset.Y
If muerto = 1 Then YYY = -9
 
Pos = YYY + GrhData(GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)).pixelHeight
Grh = HeadData(Head).Head(3)
   
Call DibujaPJ(BackBufferSurface, Grh, BBody, Pos, index)
   
If Casco <> 2 And Casco > 0 Then
    Grh = CascoAnimData(Casco).Head(3)
    Call DibujaPJ(BackBufferSurface, Grh, BBody, Pos, index)
End If
 
If Weapon <> 2 And Weapon > 0 Then
    Grh = WeaponAnimData(Weapon).WeaponWalk(3)
    Call DibujaPJ(BackBufferSurface, Grh, XBody, YBody, index)
End If
 
If Shield <> 2 And Shield > 0 Then
    Grh = ShieldAnimData(Shield).ShieldWalk(3)
    Call DibujaPJ(BackBufferSurface, Grh, XBody, BBody, index)
End If
   
End Sub

Private Sub DrawGrafico(Grh As Grh, ByVal X As Byte, ByVal Y As Byte)

Dim r2 As RECT, auxr As RECT
Dim iGrhIndex As Integer

    If Grh.GrhIndex <= 0 Then Exit Sub
    
    iGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
        
    With r2
        .Left = GrhData(iGrhIndex).sX
        .Top = GrhData(iGrhIndex).sY
        .Right = .Left + GrhData(iGrhIndex).pixelWidth
        .Bottom = .Top + GrhData(iGrhIndex).pixelHeight
    End With
    
    With auxr
        .Left = 0
        .Top = 0
        .Right = 50
        .Bottom = 65
    End With
    
    BackBufferSurface.BltFast X, Y, SurfaceDB.Surface(GrhData(iGrhIndex).FileNum), r2, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
    Call BackBufferSurface.BltToDC(frmCrearPersonaje.PlayerView.hDC, auxr, auxr)

End Sub

Sub DibujarCabeza(ByVal MyBody As Integer, ByVal MyHead As Integer)

Dim Grh As Grh
Dim Pos As Integer
Dim r2 As RECT

    With r2
        .Left = 0
        .Top = 0
        .Right = 50
        .Bottom = 65
    End With
    
    BackBufferSurface.BltColorFill r2, vbBlack
    
    'Grh = BodyData(MyBody).Walk(3)
    'Call DrawGrafico(Grh, 0, 0) '(Grh, 12, 15)
    
    'Pos = BodyData(MyBody).HeadOffset.Y + GrhData(GrhData(Grh.GrhIndex).Frames(1)).pixelHeight
    Grh = HeadData(MyHead).Head(3)
    Call DrawGrafico(Grh, 8, 3) '(Grh, 0, 0)
    
    frmCrearPersonaje.PlayerView.Refresh
    
End Sub

Sub DameOpciones()

Dim i As Integer

If frmCrearPersonaje.lstGenero.listIndex < 0 Or frmCrearPersonaje.lstRaza.listIndex < 0 Then
    frmCrearPersonaje.Cabeza.Enabled = False
ElseIf frmCrearPersonaje.lstGenero.listIndex <> -1 And frmCrearPersonaje.lstRaza.listIndex <> -1 Then
    frmCrearPersonaje.Cabeza.Enabled = True
End If

frmCrearPersonaje.Cabeza.Clear
    
Select Case frmCrearPersonaje.lstGenero.List(frmCrearPersonaje.lstGenero.listIndex)
   Case "Hombre"
        Select Case frmCrearPersonaje.lstRaza.List(frmCrearPersonaje.lstRaza.listIndex)
            Case "Humano"
                For i = 1 To 30
                    frmCrearPersonaje.Cabeza.AddItem i
                Next i
                MiCuerpo = 1
            Case "Elfo"
                For i = 101 To 113
                    If i = 113 Then i = 201
                    frmCrearPersonaje.Cabeza.AddItem i
                Next i
                MiCuerpo = 2
            Case "Elfo Oscuro"
                For i = 202 To 209
                    frmCrearPersonaje.Cabeza.AddItem i
                Next i
                MiCuerpo = 3
            Case "Enano"
                For i = 301 To 305
                    frmCrearPersonaje.Cabeza.AddItem i
                Next i
                MiCuerpo = 52
            Case "Gnomo"
                For i = 401 To 406
                    frmCrearPersonaje.Cabeza.AddItem i
                Next i
                MiCuerpo = 52
            Case Else
                UserHead = 1
                MiCuerpo = 1
        End Select
   Case "Mujer"
        Select Case frmCrearPersonaje.lstRaza.List(frmCrearPersonaje.lstRaza.listIndex)
            Case "Humano"
                For i = 70 To 76
                    frmCrearPersonaje.Cabeza.AddItem i
                Next i
                MiCuerpo = 1
            Case "Elfo"
                For i = 170 To 176
                    frmCrearPersonaje.Cabeza.AddItem i
                Next i
                MiCuerpo = 2
            Case "Elfo Oscuro"
                For i = 270 To 280
                    frmCrearPersonaje.Cabeza.AddItem i
                Next i
                MiCuerpo = 3
            Case "Gnomo"
                For i = 470 To 474
                    frmCrearPersonaje.Cabeza.AddItem i
                Next i
                MiCuerpo = 52
            Case "Enano"
                UserHead = RandomNumber(1, 3) + 369
                MiCuerpo = 52
            Case Else
                frmCrearPersonaje.Cabeza.AddItem "70"
                MiCuerpo = 1
        End Select
End Select

MinCabeza = Val(frmCrearPersonaje.Cabeza.List(0))
MaxCabeza = Val(frmCrearPersonaje.Cabeza.List(frmCrearPersonaje.Cabeza.ListCount - 1))
frmCrearPersonaje.PlayerView.Cls
If frmCrearPersonaje.lstGenero.listIndex < 0 Or frmCrearPersonaje.lstRaza.listIndex < 0 Then
    frmCrearPersonaje.Cabeza.listIndex = -1
    frmCrearPersonaje.CabA.Enabled = False
    frmCrearPersonaje.CabS.Enabled = False
ElseIf frmCrearPersonaje.Cabeza.ListCount > 0 And (frmCrearPersonaje.lstGenero.listIndex <> -1 And frmCrearPersonaje.lstRaza.listIndex <> -1) Then
    frmCrearPersonaje.Cabeza.listIndex = 0
    Call DibujarCabeza(MiCuerpo, MiCabeza)
        frmCrearPersonaje.CabA.Enabled = True
    frmCrearPersonaje.CabS.Enabled = True
End If


End Sub
