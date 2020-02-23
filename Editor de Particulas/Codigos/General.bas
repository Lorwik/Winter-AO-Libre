Attribute VB_Name = "Mod_General"
Option Explicit
Private Const CANT_GRH_INDEX As Long = 40000

Public Sub LoadGrhData()
On Error GoTo ErrorHandler
    Dim Grh As Long
    Dim Frame As Long
    Dim grhCount As Long
    Dim handle As Integer
    Dim fileVersion As Long
   
    'Open files
    handle = FreeFile()

    Open App.Path & "\INIT\Graficos.ind" For Binary Access Read As handle
    Seek #handle, 1
   
    'Get file version
    Get handle, , fileVersion
   
    'Get number of grhs
    Get handle, , grhCount
   
    'Resize arrays
    ReDim GrhData(1 To grhCount) As GrhData
    
    Get handle, , Grh

    While Not Grh <= 0
        With GrhData(Grh)
        
            'Get number of frames
            Get handle, , .NumFrames
            If .NumFrames <= 0 Then GoTo ErrorHandler
            

            ReDim .Frames(1 To GrhData(Grh).NumFrames)
           
            If .NumFrames > 1 Then
                'Read a animation GRH set
                For Frame = 1 To .NumFrames
                    Get handle, , .Frames(Frame)
                    If .Frames(Frame) <= 0 Or .Frames(Frame) > grhCount Then
                        GoTo ErrorHandler
                    End If
                Next Frame
               
                Get handle, , .speed
               
                If .speed <= 0 Then GoTo ErrorHandler
               
                'Compute width and height
                .pixelHeight = GrhData(.Frames(1)).pixelHeight
                If .pixelHeight <= 0 Then GoTo ErrorHandler
               
                .pixelWidth = GrhData(.Frames(1)).pixelWidth
                If .pixelWidth <= 0 Then GoTo ErrorHandler
               
                .TileWidth = GrhData(.Frames(1)).TileWidth
                If .TileWidth <= 0 Then GoTo ErrorHandler
               
                .TileHeight = GrhData(.Frames(1)).TileHeight
                If .TileHeight <= 0 Then GoTo ErrorHandler
            Else
                'Read in normal GRH data
                Get handle, , .FileNum
                If .FileNum <= 0 Then GoTo ErrorHandler
               
                Get handle, , GrhData(Grh).SX
                If .SX < 0 Then GoTo ErrorHandler
               
                Get handle, , .SY
                If .SY < 0 Then GoTo ErrorHandler
               
                Get handle, , .pixelWidth
                If .pixelWidth <= 0 Then GoTo ErrorHandler
               
                Get handle, , .pixelHeight
                If .pixelHeight <= 0 Then GoTo ErrorHandler
               
                'Compute width and height
                .TileWidth = .pixelWidth / 32
                .TileHeight = .pixelHeight / 32
               
                .Frames(1) = Grh
            End If
        End With
    Get handle, , Grh
    Wend
   
    Close handle
 
 Exit Sub
 
ErrorHandler:
End Sub

Public Function DirGraficos() As String
    DirGraficos = App.Path & "\GRAFICOS\"
End Function

Sub AddtoRichTextBox(ByRef RichTextBox As RichTextBox, ByVal Text As String, Optional ByVal red As Integer = -1, Optional ByVal green As Integer, Optional ByVal blue As Integer, Optional ByVal Bold As Boolean = False, Optional ByVal Italic As Boolean = False, Optional ByVal bCrLf As Boolean = False)
'******************************************
'Adds text to a Richtext box at the bottom.
'Automatically scrolls to new text.
'Text box MUST be multiline and have a 3D
'apperance!
'******************************************
    With RichTextBox
        If (Len(.Text)) > 10000 Then .Text = ""
        
        .SelStart = Len(RichTextBox.Text)
        .SelLength = 0
        
        .SelBold = Bold
        .SelItalic = Italic
        
        If Not red = -1 Then .SelColor = RGB(red, green, blue)
        
        .SelText = IIf(bCrLf, Text, Text & vbCrLf)
        
        RichTextBox.Refresh
    End With
End Sub

Public Function RandomNumber(ByVal LowerBound As Long, ByVal UpperBound As Long) As Long
    'Initialize randomizer
    Randomize Timer
    
    'Generate random number
    RandomNumber = (UpperBound - LowerBound) * Rnd + LowerBound
End Function

Sub UnloadAllForms()
On Error Resume Next

    Dim mifrm As Form
    
    For Each mifrm In Forms
        Unload mifrm
    Next
End Sub

Sub Main()
On Error Resume Next


    
    ChDrive App.Path
    ChDir App.Path
    
    AddtoRichTextBox frmCargando.Status, "Cargando Engine Grafico...."
    
    frmCargando.Show
    frmCargando.Refresh
    DoEvents
    AddtoRichTextBox frmCargando.Status, "Terminado carga de Engine Grafico con Extio.."
    AddtoRichTextBox frmCargando.Status, "Abriento interfaz. ¡Bienvenido!"
    AddtoRichTextBox frmCargando.Status, "www.RincondelAO.net"
    Call engine.Engine_Init
    Call LoadGrhData
    DoEvents
    AgregaGrH (1)
    Unload frmCargando
                   
    frmMain.Show

    'Inicialización de variables globales
    prgRun = True
    Dim pausa
    pausa = False
    Call CargarParticulas
    engine.Start
    
Exit Sub
ManejadorErrores:
    MsgBox "Ha ocurrido un error irreparable, el cliente se cerrará."
    Debug.Print "Contexto:" & Err.HelpContext & " Desc:" & Err.Description & " Fuente:" & Err.Source
    End
End Sub

Public Function General_Particle_Create(ByVal ParticulaInd As Long, ByVal X As Integer, ByVal Y As Integer, Optional ByVal particle_life As Long = 0) As Long

Dim rgb_list(0 To 3) As Long
rgb_list(0) = RGB(StreamData(ParticulaInd).colortint(0).r, StreamData(ParticulaInd).colortint(0).g, StreamData(ParticulaInd).colortint(0).B)
rgb_list(1) = RGB(StreamData(ParticulaInd).colortint(1).r, StreamData(ParticulaInd).colortint(1).g, StreamData(ParticulaInd).colortint(1).B)
rgb_list(2) = RGB(StreamData(ParticulaInd).colortint(2).r, StreamData(ParticulaInd).colortint(2).g, StreamData(ParticulaInd).colortint(2).B)
rgb_list(3) = RGB(StreamData(ParticulaInd).colortint(3).r, StreamData(ParticulaInd).colortint(3).g, StreamData(ParticulaInd).colortint(3).B)

General_Particle_Create = engine.Particle_Group_Create(X, Y, StreamData(ParticulaInd).grh_list, rgb_list(), StreamData(ParticulaInd).NumOfParticles, ParticulaInd, _
    StreamData(ParticulaInd).AlphaBlend, IIf(particle_life = 0, StreamData(ParticulaInd).life_counter, particle_life), StreamData(ParticulaInd).speed, , StreamData(ParticulaInd).x1, StreamData(ParticulaInd).y1, StreamData(ParticulaInd).angle, _
    StreamData(ParticulaInd).vecx1, StreamData(ParticulaInd).vecx2, StreamData(ParticulaInd).vecy1, StreamData(ParticulaInd).vecy2, _
    StreamData(ParticulaInd).life1, StreamData(ParticulaInd).life2, StreamData(ParticulaInd).friction, StreamData(ParticulaInd).spin_speedL, _
    StreamData(ParticulaInd).gravity, StreamData(ParticulaInd).grav_strength, StreamData(ParticulaInd).bounce_strength, StreamData(ParticulaInd).x2, _
    StreamData(ParticulaInd).y2, StreamData(ParticulaInd).XMove, StreamData(ParticulaInd).move_x1, StreamData(ParticulaInd).move_x2, StreamData(ParticulaInd).move_y1, _
    StreamData(ParticulaInd).move_y2, StreamData(ParticulaInd).YMove, StreamData(ParticulaInd).spin_speedH, StreamData(ParticulaInd).spin, _
    StreamData(ParticulaInd).grh_resize, StreamData(ParticulaInd).grh_resizex, StreamData(ParticulaInd).grh_resizey, _
    StreamData(ParticulaInd).Radio)


End Function

Sub CargarParticulas()
Dim StreamFile As String
Dim loopc As Long
Dim i As Long
Dim GrhListing As String
Dim TempSet As String
Dim ColorSet As Long
    
StreamFile = App.Path & "\INIT\Particles.ini"
TotalStreams = Val(General_Var_Get(StreamFile, "INIT", "Total"))

'resize StreamData array
ReDim StreamData(1 To TotalStreams) As Stream

    'fill StreamData array with info from Particles.ini
    For loopc = 1 To TotalStreams
        StreamData(loopc).Name = General_Var_Get(StreamFile, Val(loopc), "Name")
        StreamData(loopc).NumOfParticles = General_Var_Get(StreamFile, Val(loopc), "NumOfParticles")
        StreamData(loopc).x1 = General_Var_Get(StreamFile, Val(loopc), "X1")
        StreamData(loopc).y1 = General_Var_Get(StreamFile, Val(loopc), "Y1")
        StreamData(loopc).x2 = General_Var_Get(StreamFile, Val(loopc), "X2")
        StreamData(loopc).y2 = General_Var_Get(StreamFile, Val(loopc), "Y2")
        StreamData(loopc).angle = General_Var_Get(StreamFile, Val(loopc), "Angle")
        StreamData(loopc).vecx1 = General_Var_Get(StreamFile, Val(loopc), "VecX1")
        StreamData(loopc).vecx2 = General_Var_Get(StreamFile, Val(loopc), "VecX2")
        StreamData(loopc).vecy1 = General_Var_Get(StreamFile, Val(loopc), "VecY1")
        StreamData(loopc).vecy2 = General_Var_Get(StreamFile, Val(loopc), "VecY2")
        StreamData(loopc).life1 = General_Var_Get(StreamFile, Val(loopc), "Life1")
        StreamData(loopc).life2 = General_Var_Get(StreamFile, Val(loopc), "Life2")
        StreamData(loopc).friction = General_Var_Get(StreamFile, Val(loopc), "Friction")
        StreamData(loopc).spin = General_Var_Get(StreamFile, Val(loopc), "Spin")
        StreamData(loopc).spin_speedL = General_Var_Get(StreamFile, Val(loopc), "Spin_SpeedL")
        StreamData(loopc).spin_speedH = General_Var_Get(StreamFile, Val(loopc), "Spin_SpeedH")
        StreamData(loopc).AlphaBlend = General_Var_Get(StreamFile, Val(loopc), "AlphaBlend")
        StreamData(loopc).gravity = General_Var_Get(StreamFile, Val(loopc), "Gravity")
        StreamData(loopc).grav_strength = General_Var_Get(StreamFile, Val(loopc), "Grav_Strength")
        StreamData(loopc).bounce_strength = General_Var_Get(StreamFile, Val(loopc), "Bounce_Strength")
        StreamData(loopc).XMove = General_Var_Get(StreamFile, Val(loopc), "XMove")
        StreamData(loopc).YMove = General_Var_Get(StreamFile, Val(loopc), "YMove")
        StreamData(loopc).move_x1 = General_Var_Get(StreamFile, Val(loopc), "move_x1")
        StreamData(loopc).move_x2 = General_Var_Get(StreamFile, Val(loopc), "move_x2")
        StreamData(loopc).move_y1 = General_Var_Get(StreamFile, Val(loopc), "move_y1")
        StreamData(loopc).move_y2 = General_Var_Get(StreamFile, Val(loopc), "move_y2")
        StreamData(loopc).Radio = Val(General_Var_Get(StreamFile, Val(loopc), "Radio"))
        StreamData(loopc).life_counter = General_Var_Get(StreamFile, Val(loopc), "life_counter")
        StreamData(loopc).speed = Val(General_Var_Get(StreamFile, Val(loopc), "Speed"))
                StreamData(loopc).grh_resize = Val(General_Var_Get(StreamFile, Val(loopc), "resize"))
        StreamData(loopc).grh_resizex = Val(General_Var_Get(StreamFile, Val(loopc), "rx"))
        StreamData(loopc).grh_resizey = Val(General_Var_Get(StreamFile, Val(loopc), "ry"))
        StreamData(loopc).NumGrhs = General_Var_Get(StreamFile, Val(loopc), "NumGrhs")
        
        ReDim StreamData(loopc).grh_list(1 To StreamData(loopc).NumGrhs)
        GrhListing = General_Var_Get(StreamFile, Val(loopc), "Grh_List")
        
        For i = 1 To StreamData(loopc).NumGrhs
            StreamData(loopc).grh_list(i) = General_Field_Read(Str(i), GrhListing, 44)
        Next i
        StreamData(loopc).grh_list(i - 1) = StreamData(loopc).grh_list(i - 1)
        For ColorSet = 1 To 4
            TempSet = General_Var_Get(StreamFile, Val(loopc), "ColorSet" & ColorSet)
            StreamData(loopc).colortint(ColorSet - 1).r = General_Field_Read(1, TempSet, 44)
            StreamData(loopc).colortint(ColorSet - 1).g = General_Field_Read(2, TempSet, 44)
            StreamData(loopc).colortint(ColorSet - 1).B = General_Field_Read(3, TempSet, 44)
        Next ColorSet
            frmMain.List2.AddItem loopc & " - " & StreamData(loopc).Name
    Next loopc

End Sub
Public Function General_Random_Number(ByVal LowerBound As Long, ByVal UpperBound As Long) As Single
'*****************************************************************
'Author: Aaron Perkins
'Find a Random number between a range
'*****************************************************************
    Randomize Timer
    General_Random_Number = (UpperBound - LowerBound + 1) * Rnd + LowerBound
End Function
Public Sub General_Var_Write(ByVal file As String, ByVal Main As String, ByVal var As String, ByVal value As String)
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Writes a var to a text file
'*****************************************************************
    writeprivateprofilestring Main, var, value, file
End Sub


Public Function General_Var_Get(ByVal file As String, ByVal Main As String, ByVal var As String) As String
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Get a var to from a text file
'*****************************************************************
    Dim l As Long
    Dim Char As String
    Dim sSpaces As String 'Input that the program will retrieve
    Dim szReturn As String 'Default value if the string is not found
    
    szReturn = ""
    
    sSpaces = Space$(5000)
    
    getprivateprofilestring Main, var, szReturn, sSpaces, Len(sSpaces), file
    
    General_Var_Get = RTrim$(sSpaces)
    General_Var_Get = Left$(General_Var_Get, Len(General_Var_Get) - 1)
End Function

Public Function General_Field_Read(ByVal field_pos As Long, ByVal Text As String, ByVal delimiter As Byte) As String
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Gets a field from a delimited string
'*****************************************************************
    Dim i As Long
    Dim LastPos As Long
    Dim FieldNum As Long
    
    LastPos = 0
    FieldNum = 0
    For i = 1 To Len(Text)
        If delimiter = CByte(Asc(Mid$(Text, i, 1))) Then
            FieldNum = FieldNum + 1
            If FieldNum = field_pos Then
                General_Field_Read = Mid$(Text, LastPos + 1, (InStr(LastPos + 1, Text, Chr$(delimiter), vbTextCompare) - 1) - (LastPos))
                Exit Function
            End If
            LastPos = i
        End If
    Next i
    FieldNum = FieldNum + 1
    If FieldNum = field_pos Then
        General_Field_Read = Mid$(Text, LastPos + 1)
    End If
End Function



Public Sub AgregaGrH(ByVal numgrh As Integer)
Dim i As Long
Dim EsteIndex As Long
Dim CuentaIndex As Long

GrhData(numgrh).FileNum = 1
GrhData(numgrh).NumFrames = 1
GrhData(numgrh).pixelHeight = 32
GrhData(numgrh).pixelWidth = 32
GrhData(numgrh).Frames(1) = numgrh

CuentaIndex = -1
frmMain.lstGrhs.Clear
For i = 1 To 32000
    If GrhData(i).NumFrames = 1 Then
        frmMain.lstGrhs.AddItem i
        CuentaIndex = CuentaIndex + 1
    ElseIf GrhData(i).NumFrames > 1 Then
        frmMain.lstGrhs.AddItem i & " (animacion)"
        CuentaIndex = CuentaIndex + 1
    End If
    If i = numgrh Then
        EsteIndex = CuentaIndex
    End If
Next i
frmMain.lstGrhs.ListIndex = EsteIndex
End Sub

Public Function General_File_Exists(ByVal file_path As String, ByVal file_type As VbFileAttribute) As Boolean
    If Dir(file_path, file_type) = "" Then
        General_File_Exists = False
    Else
        General_File_Exists = True
    End If
End Function

Public Sub HookSurfaceHwnd(pic As Form)
    Call ReleaseCapture
    Call SendMessage(pic.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End Sub

Sub ReCargarParticulas()
Dim StreamFile As String
Dim loopc As Long
Dim i As Long
Dim GrhListing As String
Dim TempSet As String
Dim ColorSet As Long
    
StreamFile = App.Path & "\INIT\Particles.ini"
TotalStreams = Val(General_Var_Get(StreamFile, "INIT", "Total"))

'resize StreamData array
ReDim StreamData(1 To TotalStreams) As Stream

    'fill StreamData array with info from Particles.ini
    For loopc = 1 To TotalStreams
        StreamData(loopc).Name = General_Var_Get(StreamFile, Val(loopc), "Name")
        StreamData(loopc).NumOfParticles = General_Var_Get(StreamFile, Val(loopc), "NumOfParticles")
        StreamData(loopc).x1 = General_Var_Get(StreamFile, Val(loopc), "X1")
        StreamData(loopc).y1 = General_Var_Get(StreamFile, Val(loopc), "Y1")
        StreamData(loopc).x2 = General_Var_Get(StreamFile, Val(loopc), "X2")
        StreamData(loopc).y2 = General_Var_Get(StreamFile, Val(loopc), "Y2")
        StreamData(loopc).angle = General_Var_Get(StreamFile, Val(loopc), "Angle")
        StreamData(loopc).vecx1 = General_Var_Get(StreamFile, Val(loopc), "VecX1")
        StreamData(loopc).vecx2 = General_Var_Get(StreamFile, Val(loopc), "VecX2")
        StreamData(loopc).vecy1 = General_Var_Get(StreamFile, Val(loopc), "VecY1")
        StreamData(loopc).vecy2 = General_Var_Get(StreamFile, Val(loopc), "VecY2")
        StreamData(loopc).life1 = General_Var_Get(StreamFile, Val(loopc), "Life1")
        StreamData(loopc).life2 = General_Var_Get(StreamFile, Val(loopc), "Life2")
        StreamData(loopc).friction = General_Var_Get(StreamFile, Val(loopc), "Friction")
        StreamData(loopc).spin = General_Var_Get(StreamFile, Val(loopc), "Spin")
        StreamData(loopc).spin_speedL = General_Var_Get(StreamFile, Val(loopc), "Spin_SpeedL")
        StreamData(loopc).spin_speedH = General_Var_Get(StreamFile, Val(loopc), "Spin_SpeedH")
        StreamData(loopc).AlphaBlend = General_Var_Get(StreamFile, Val(loopc), "AlphaBlend")
        StreamData(loopc).gravity = General_Var_Get(StreamFile, Val(loopc), "Gravity")
        StreamData(loopc).grav_strength = General_Var_Get(StreamFile, Val(loopc), "Grav_Strength")
        StreamData(loopc).bounce_strength = General_Var_Get(StreamFile, Val(loopc), "Bounce_Strength")
        StreamData(loopc).XMove = General_Var_Get(StreamFile, Val(loopc), "XMove")
        StreamData(loopc).YMove = General_Var_Get(StreamFile, Val(loopc), "YMove")
        StreamData(loopc).move_x1 = General_Var_Get(StreamFile, Val(loopc), "move_x1")
        StreamData(loopc).move_x2 = General_Var_Get(StreamFile, Val(loopc), "move_x2")
        StreamData(loopc).move_y1 = General_Var_Get(StreamFile, Val(loopc), "move_y1")
        StreamData(loopc).move_y2 = General_Var_Get(StreamFile, Val(loopc), "move_y2")
        StreamData(loopc).Radio = Val(General_Var_Get(StreamFile, Val(loopc), "Radio"))
        StreamData(loopc).life_counter = General_Var_Get(StreamFile, Val(loopc), "life_counter")
        StreamData(loopc).speed = Val(General_Var_Get(StreamFile, Val(loopc), "Speed"))
                StreamData(loopc).grh_resize = Val(General_Var_Get(StreamFile, Val(loopc), "resize"))
        StreamData(loopc).grh_resizex = Val(General_Var_Get(StreamFile, Val(loopc), "rx"))
        StreamData(loopc).grh_resizey = Val(General_Var_Get(StreamFile, Val(loopc), "ry"))
        StreamData(loopc).NumGrhs = General_Var_Get(StreamFile, Val(loopc), "NumGrhs")
        
        ReDim StreamData(loopc).grh_list(1 To StreamData(loopc).NumGrhs)
        GrhListing = General_Var_Get(StreamFile, Val(loopc), "Grh_List")
        
        For i = 1 To StreamData(loopc).NumGrhs
            StreamData(loopc).grh_list(i) = General_Field_Read(Str(i), GrhListing, 44)
        Next i
        StreamData(loopc).grh_list(i - 1) = StreamData(loopc).grh_list(i - 1)
        For ColorSet = 1 To 4
            TempSet = General_Var_Get(StreamFile, Val(loopc), "ColorSet" & ColorSet)
            StreamData(loopc).colortint(ColorSet - 1).r = General_Field_Read(1, TempSet, 44)
            StreamData(loopc).colortint(ColorSet - 1).g = General_Field_Read(2, TempSet, 44)
            StreamData(loopc).colortint(ColorSet - 1).B = General_Field_Read(3, TempSet, 44)
        Next ColorSet
            frmMain.List2.AddItem loopc & " - " & StreamData(loopc).Name
    Next loopc

End Sub
Private Function Grh_Check(ByVal grh_index As Long) As Boolean
    If grh_index > 0 And grh_index <= CANT_GRH_INDEX Then
        Grh_Check = GrhData(grh_index).NumFrames
    End If
End Function
