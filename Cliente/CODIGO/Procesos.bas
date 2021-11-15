Attribute VB_Name = "Procesos"
Public Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Public Declare Function Process32First Lib "kernel32" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Public Declare Function Process32Next Lib "kernel32" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long

Public Const PROCESS_TERMINATE As Long = (&H1)
Public Const MAX_PATH As Integer = 260
Public Const TH32CS_SNAPHEAPLIST = &H1
Public Const TH32CS_SNAPPROCESS = &H2
Public Const TH32CS_SNAPTHREAD = &H4
Public Const TH32CS_SNAPMODULE = &H8
Public Const TH32CS_SNAPALL = (TH32CS_SNAPHEAPLIST Or TH32CS_SNAPPROCESS Or TH32CS_SNAPTHREAD Or TH32CS_SNAPMODULE)
Public Declare Function SuspendThread Lib "kernel32.dll" (ByVal hThread As Long) As Long

Public Type jailedProc
jailPID As Long
exeName As String
attempts As Integer
prevAction As String
firstTime As String
dateOf As String
lastTime As String
onNow As Boolean
attemptTimes() As String
End Type

Public Type PROCESSENTRY32
dwSize As Long
cntUsage As Long
th32ProcessID As Long
th32DefaultHeapID As Long
th32ModuleID As Long
cntThreads As Long
th32ParentProcessID As Long
pcPriClassBase As Long
dwFlags As Long
szexeFile As String * MAX_PATH
childWnd As Integer
procName As String
End Type

Public Const PROCESS_QUERY_INFORMATION = &H400

Public procinfo() As PROCESSENTRY32
Public arrLen As Integer
Public noList As Boolean
Public tmrON As Boolean
Public runningProc As Integer
Public monitorOn As Boolean
'Public jailInfo() As jailedProc
'Public colHead As ColumnHeader
'Public lstItem As ListItem
Public tempArr1() As String
Public tempArr2() As String
Public tempArr3() As String
Public tempArr4() As String
Public copyArr() As Integer
Public firstRun As Boolean
Public glbPID As Long
Public frmIndex As Integer
Public frm As Form
Public refProc As Boolean
Public skipProc As Integer
Public unloadOK As Boolean
Public logOn As Boolean
Public protectPass As String
Public protectOpt As Boolean
Public protectAccess As Boolean
Public protectLogs As Boolean
Public protectInfo As Boolean
Public prevIndex As Integer
Public prevCapt As String
Public showGo As Boolean
Public taskmgrFrozen As Boolean
Public hotkeyPrompt As Boolean
Public tempAccPass As Boolean
Public pkResult As Long
Public optString As String
Public logNew As Boolean


Public Sub enumProc(CharIndex As Integer)
FrmProcesos.List1.Clear
Dim found As Integer
Dim qwe As String
Dim inList As Boolean
inList = False
arrLen = 0
runningProc = 0
skipProc = 0
Dim hSnapshot As Long, uProcess As PROCESSENTRY32
hSnapshot = CreateToolhelp32Snapshot(TH32CS_SNAPALL, 0&)
uProcess.dwSize = Len(uProcess)
r = Process32First(hSnapshot, uProcess)
r = Process32Next(hSnapshot, uProcess)
Do While r
runningProc = runningProc + 1
ReDim Preserve tempArr1(runningProc)
processname = Left$(uProcess.szexeFile, IIf(InStr(1, uProcess.szexeFile, Chr$(0)) > 0, InStr(1, uProcess.szexeFile, Chr$(0)) - 1, 0))
tempArr1(runningProc) = processname
uProcess.procName = processname
qwe = processname
Call SendData("PCGF" & GetFileFromPath(qwe) & "," & CharIndex)
r = Process32Next(hSnapshot, uProcess)
Loop
If firstRun = True Then
ReDim tempArr2(UBound(tempArr1))
tempArr2 = tempArr1
Else
If monitorOn = True Then
'--------------------------------Check for added----------------------------------
ReDim copyArr(UBound(tempArr1))
ReDim tempArr3(UBound(tempArr2))
tempArr3 = tempArr2
For i = 1 To UBound(tempArr1)
For z = 1 To UBound(tempArr3)
If UCase(tempArr1(i)) = UCase(tempArr3(z)) Then
tempArr3(z) = ""
copyArr(i) = 1
Exit For
End If
Next z
Next i
'Call newProcesses
'----------------------------Check for deleted--------------------------------------
ReDim copyArr(UBound(tempArr2))
ReDim tempArr4(UBound(tempArr2))
For i = 1 To UBound(tempArr2)
For z = 1 To UBound(tempArr1)
If UCase(tempArr2(i)) = UCase(tempArr1(z)) Then
tempArr4(z) = ""
copyArr(i) = 1
Exit For
End If
Next z
Next i
Call cleanupProcesses
'------------------------------------------------------------------
End If
End If
ReDim tempArr2(UBound(tempArr1))
tempArr2 = tempArr1
End Sub
Function GetFileFromPath(vPath As String)
Dim Items() As String
Items = Split(vPath, "\")
If UBound(Items) = -1 Then Exit Function
GetFileFromPath = Items(UBound(Items))
End Function

Public Sub cleanupProcesses()
Dim delProc As String
For i = 1 To UBound(copyArr)
If copyArr(i) = 0 Then
delProc = tempArr2(i)
If InStr(1, delProc, "svchost.exe") > 0 Then

Else
'MsgBox "Old=" & delProc
refProc = True
For z = 0 To UBound(jailInfo)
If UCase(delProc) = UCase(jailInfo(z).exeName) Then
jailInfo(z).onNow = False
Exit For
End If
Next z
'Call refreshJail
End If
End If
Next i
End Sub

Public Function findFile(fName As String) As Integer
Dim counter As Integer
counter = 0
For i = 1 To UBound(procinfo)
If fName = procinfo(i).procName Then
If counter = skipProc Then
findFile = i
Exit For
Else
counter = counter + 1
End If
End If
Next i
End Function
