VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMP3Player"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
 
Private mvarmp3file As String
 
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
 
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal lBuffer As Long) As Long
 
Dim mVolume As Integer
 
Public Sub pauseMP3()
mciSendString "pause audio1", 0, 0, 0
End Sub
 
Public Sub playMP3()
Dim rtn As Long
Dim FileName As String
Dim lngRes As Long, strPath As String
 
strPath = String$(165, 0)  'Create a buffer
lngRes = GetShortPathName(mvarmp3file, strPath, 164)
FileName = Trim(left$(strPath, lngRes))
 
rtn = mciSendString("Open " & FileName & " Alias audio1", 0&, 0, 0)
rtn = mciSendString("play audio1", 0&, 0, 0)
 
End Sub
 
Public Sub stopMP3()
Dim rtn As Long
 
rtn = mciSendString("stop audio1", 0&, 0, 0)
rtn = mciSendString("close audio1", 0&, 0, 0)
 
End Sub
 
Public Property Let mp3file(ByVal vData As String)
    mvarmp3file = vData
End Property
 
Public Property Get mp3file() As String
    mp3file = mvarmp3file
End Property
 
Public Property Let Volume(ByVal nVolume As Integer)
    mciSendString "setaudio audio1 volume to " & nVolume, "", 0, 0
    mVolume = nVolume
End Property
 
Public Property Get Volume() As Integer
    Volume = mVolume
End Property
 
'setaudio MP3 volume to " + NewVolume
Public Function Position(secs As Boolean) As String
'Returns the track's current position as "mm:ss".
'Returns tracks's current position in seconds if secs = True
 
Static p As String * 30
Dim sec As Integer
Dim mins As Integer
 
mciSendString "set audio1 Time Format milliseconds", "", 0, 0
mciSendString "status audio1 position", p, Len(p), 0
sec = Round(Val(mid$(p, 1, Len(p))) / 1000)
 
If secs Then
    Position = sec
Else
    If sec < 60 Then Position = "0:" & Format(sec, "00")
    If sec > 59 Then
        mins = Int(sec / 60)
        sec = sec - (mins * 60)
        Position = Format(mins, "00") & ":" & Format(sec, "00")
    End If
End If
 
End Function
 
Public Function TrackLength(secs As Boolean) As String
'Returns the track's length as "mm:ss".
'Returns tracks's length in seconds if secs = True
 
Static l As String * 30
Dim sec As Integer
Dim mins As Integer

mciSendString "set audio1 Time Format milliseconds", "", 0, 0
mciSendString "status audio1 length", l, Len(l), 0
 
sec = Round(Val(mid$(l, 1, Len(l))) / 1000)
 
If secs Then
    TrackLength = sec
Else
    If sec < 60 Then
        TrackLenghth = "0:" & Format(sec, "00")
    Else
        mins = Int(sec / 60)
        sec = sec - (mins * 60)
        TrackLength = Format(mins, "00") & ":" & Format(sec, "00")
    End If
End If
 
End Function
 
Public Function IsItPlaying() As Boolean
'Returns true if the file is playing. False otherwise.
 
Static yn As String * 30
 
mciSendString "status audio1 mode", yn, Len(yn), 0
IsItPlaying = (mid$(yn, 1, 7) = "playing")
 
End Function



