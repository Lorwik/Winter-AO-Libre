VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDialogs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0
Private Const MAX_DIALOGS As Byte = 100
Private Const MAX_LENGTH As Byte = 20
Private Const MS_PER_CHAR As Byte = 100
Private Const MS_ADD_EXTRA As Integer = 5000

Private Type dialog
    TextLine()  As String
    X           As Integer
    Y           As Integer
    CharIndex   As Integer
    color       As Long
    Renderable  As Boolean
    Elevate As Integer
    Alpha As Integer
    Inicia As Boolean
    Efect As Boolean
    TickCountFadeOut As Long
    TickCountFadeIn As Long
    TickCOuntOutFinal As Long
End Type

Private Dialogs(MAX_DIALOGS - 1) As dialog
Private DialogCount As Byte
Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Sub Class_Initialize()
    DialogCount = 0
End Sub

Private Function BinarySearch(ByVal CharIndex As Integer) As Integer
Dim min As Long, max As Long, mid As Long
    min = 0
    max = DialogCount - 1
    
    Do While min <= max
        mid = (min + max) \ 2
        
        If Dialogs(mid).CharIndex < CharIndex Then
            min = mid + 1
        ElseIf Dialogs(mid).CharIndex > CharIndex Then
            max = mid - 1
        Else
            BinarySearch = mid
            Exit Function
        End If
    Loop
    BinarySearch = Not min
End Function

Public Sub CreateDialog(ByRef chat As String, ByVal CharIndex As Integer, ByRef color As Long, Optional ByVal Efecto1 As Boolean = False)
    Dim slot As Integer
    Dim i As Long
    
    slot = BinarySearch(CharIndex)
    
    If slot < 0 Then
        If DialogCount = MAX_DIALOGS Then Exit Sub
        slot = Not slot
        
        For i = DialogCount To slot + 1 Step -1
            Dialogs(i) = Dialogs(i - 1)
        Next i
        
        DialogCount = DialogCount + 1
    End If
    
    With Dialogs(slot)
        .CharIndex = CharIndex
        .color = ColorToDX8(color)
        .Renderable = False
        .TextLine = FormatChat(chat)
        .Elevate = 20
        .Inicia = True
        .Efect = Efecto1
        If .Efect = True Then .Alpha = 250
        If .Efect = False Then .Alpha = 0
        .TickCountFadeOut = GetTickCount
        .TickCountFadeIn = .TickCountFadeOut
        .TickCOuntOutFinal = .TickCountFadeOut
    End With
End Sub
Private Function FormatChat(ByRef chat As String) As String()
    Dim word As String
    Dim curPos As Long
    Dim Length As Long
    Dim acumLength As Long
    Dim lineLength As Long
    Dim wordLength As Long
    Dim curLine As Long
    Dim chatLines() As String
    
    curLine = 0
    curPos = 1
    Length = Len(chat)
    acumLength = 0
    lineLength = -1
    ReDim chatLines(FieldCount(chat, 32)) As String

    Do While acumLength < Length
        word = ReadField(curPos, chat, 32)
        
        wordLength = Len(word)
        
        If lineLength = -1 Then
            chatLines(curLine) = word
            
            lineLength = wordLength
            acumLength = wordLength
        Else
            If lineLength + wordLength + 1 > MAX_LENGTH Then
                curLine = curLine + 1
                chatLines(curLine) = word
                lineLength = wordLength
            Else
                chatLines(curLine) = chatLines(curLine) & " " & word
                
                lineLength = lineLength + wordLength + 1
            End If
            
            acumLength = acumLength + wordLength + 1
        End If
        curPos = curPos + 1
    Loop
    If curLine = 0 And Length < MAX_LENGTH Then
        chatLines(curLine) = String((MAX_LENGTH - Length) \ 4 + 1, " ") & chatLines(curLine)
    End If
    ReDim Preserve chatLines(curLine) As String
    
    FormatChat = chatLines
End Function


Public Sub UpdateDialogPos(ByVal X As Integer, ByVal Y As Integer, ByVal CharIndex As Integer)

    Dim slot As Integer
    
    slot = BinarySearch(CharIndex)
    
    If slot >= 0 Then
        With Dialogs(slot)
            .X = X - (MAX_LENGTH \ 2) * 3
            .Y = Y - (UBound(.TextLine()) * 3)
            .Renderable = True
        End With
    End If
End Sub

Public Sub Render()
On Error Resume Next
Dim i As Long, j As Long, Offset As Integer, removedDialogs As Long
    For i = 0 To DialogCount - 1
        With Dialogs(i - removedDialogs)
            If .Renderable Then
                Offset = -10 * UBound(.TextLine())
                
            If .Efect = False Then
                'Subir Letras
                If GetTickCount - .TickCountFadeIn > 18 Then
                    .Elevate = .Elevate - 1
                    .TickCountFadeIn = GetTickCount
                End If
                
                    If .Elevate >= 0 Then .Y = .Y + .Elevate

                'Alpha de Aparición
                If .Alpha <= 255 And .Inicia = True And GetTickCount - .TickCOuntOutFinal > 18 Then
                    .TickCOuntOutFinal = GetTickCount
                    If Not .Alpha = 255 Then .Alpha = .Alpha + 5
                    If .Alpha >= 255 Then .Inicia = False: .Alpha = 255
                End If
                                
                'Alpha de Desaparición
                If Not .Inicia And GetTickCount - .TickCountFadeOut > 10000 Then 'Give 10 Seconds to read Dialog
                    If .Alpha >= 0 And Not .Inicia And GetTickCount - .TickCOuntOutFinal > 18 Then
                        .Alpha = .Alpha - 5
                        .TickCOuntOutFinal = GetTickCount
                            'Si el Alpha es 0, removemos el dialogo
                            If .Alpha <= 0 Then
                                Call RemoveDialog(.CharIndex)
                                removedDialogs = removedDialogs + 1
                                Exit Sub
                            End If
                    End If
                End If
            Else
                'Subir Letras
                    If GetTickCount - .TickCountFadeIn > 18 Then
                        .Elevate = .Elevate - 1
                        .TickCountFadeIn = GetTickCount
                    End If
                    
                    .Y = .Y + (.Elevate)
                
                'Alpha de Desaparición
                If GetTickCount - .TickCountFadeOut > 1000 Then
                    If .Alpha >= 0 And GetTickCount - .TickCOuntOutFinal > 18 Then
                        .Alpha = .Alpha - 5
                        .TickCOuntOutFinal = GetTickCount
                            'Si el Alpha es 0, removemos el dialogo
                            If .Alpha <= 0 Then
                                Call RemoveDialog(.CharIndex)
                                removedDialogs = removedDialogs + 1
                                Exit Sub
                            End If
                    End If
                End If
            End If
                
                For j = 0 To UBound(.TextLine())

                    Fonts_Render_String .TextLine(j), .X, .Y + Offset, .color
                    Offset = Offset + 10 + 5
                Next j
                .Renderable = False
            End If
        End With
    Next i
End Sub

Public Sub RemoveDialog(ByVal CharIndex As Integer)
Dim slot As Integer, i As Long
    slot = BinarySearch(CharIndex)
    If slot < 0 Then Exit Sub
    For i = slot To MAX_DIALOGS - 2
        Dialogs(i) = Dialogs(i + 1)
    Next i
    DialogCount = DialogCount - 1
End Sub

Public Sub RemoveAllDialogs()
    DialogCount = 0
End Sub
