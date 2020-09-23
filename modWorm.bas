Attribute VB_Name = "modWorm"
Public Type Dot
 Left As Double
 Top As Double
 Visible As Boolean
End Type

Public Type typeApple
    Left As Double
    Top As Double
    Width As Long
    Height As Long
    Pic As Integer
    Visible As Boolean
End Type

Public bLoaded As Boolean 'Whether or not the game has been loaded yet

'Bitblt for the game's graphics
Declare Function BitBlt Lib "gdi32" ( _
        ByVal hDestDC As Long, _
        ByVal X As Long, _
        ByVal Y As Long, _
        ByVal nWidth As Long, _
        ByVal nHeight As Long, _
        ByVal hSrcDC As Long, _
        ByVal xSrc As Long, _
        ByVal ySrc As Long, _
        ByVal dwRop As Long) As Long

Public GameType As Long
Public Difficulty As Long
Public Speed As Long
Public Control As Long
Public Multiplier As Long
Public HighScore(0 To 5) As Long
Public Size As Long

'Used for collision detection
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long

'Wave Functions
Private Declare Function sndPlaySound Lib "WINMM.DLL" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

'Midi functions
Declare Function mciSendString Lib "WINMM.DLL" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As Any, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Declare Function mciGetErrorString Lib "WINMM.DLL" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal lBuffer As Long) As Long

'Ini Functions
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Function GetFromIni(strSectionHeader As String, strVariableName As String, strFileName As String) As String
    Dim strReturn As String
    strReturn = String(255, Chr(0))
    GetFromIni = Left$(strReturn, GetPrivateProfileString(strSectionHeader, ByVal strVariableName, "", strReturn, Len(strReturn), strFileName))
End Function
Function WriteIni(strSectionHeader As String, strVariableName As String, strValue As String, strFileName As String) As Integer
    WriteIni = WritePrivateProfileString(strSectionHeader, strVariableName, strValue, strFileName)
End Function
Public Sub Encode(strValue As String, strINIValue As String, strINILength As String, nSave As String)
'This is a simple encoding method that changes the values of characters to make editing the high score ini file at least a little harder
On Error Resume Next
    Dim strLength As String
    
    strLength = CStr(Len(strValue))
    
    If Len(strValue) < 10 Then
        strLength = "0" & strLength
    End If
    
    strValue = Eyncrypt(strValue)
    
    Dim strLength2 As String
    
    strLength2 = strLength
    strLength = Eyncrypt(strLength2)
    
    Call WriteIni("GEN", strINIValue, strValue, nSave)
    Call WriteIni("GEN", strINILength, strLength, nSave)
    
End Sub
Public Function Eyncrypt(sData As String) As String
On Error Resume Next
    Dim sTemp As String, sTemp1 As String
    Dim strBS As String
    Dim strBS2 As String
    'Randomly creates garbage characters before and after the actual value
    For i = 1 To 6
        strBS = strBS & Int(Rnd * 9)
        strBS2 = strBS2 & Int(Rnd * 9)
    Next 'i
    
    sData = strBS & sData & strBS2 'Combines the garbage and the actual string

    'Change the characters
    For iI% = 1 To Len(sData$)
        sTemp$ = Mid$(sData$, iI%, 1)
        lT = Asc(sTemp$) * 2
        sTemp1$ = sTemp1$ & Chr(lT)
    Next iI%
    Eyncrypt$ = sTemp1$
End Function
Public Function Decode(sData As String) As String
    Dim sTemp As String, sTemp1 As String

    For iI% = 1 To Len(sData$)
        sTemp$ = Mid$(sData$, iI%, 1)
        lT = Asc(sTemp$) \ 2
        sTemp1$ = sTemp1$ & Chr(lT)
    Next iI%
    Decode$ = sTemp1$
End Function
Public Function DecryptString(ByVal nSave As String, ByVal strField As String, ByVal strLength As String) As String
'Converts the encrypted string to an unencrypted string
Dim sLength As String
Dim iLength As Integer
sLength = GetFromIni("GEN", strLength, nSave)
sLength = Decode(sLength)
iLength = CInt(Mid$(sLength, 7, 2))


strField = GetFromIni("GEN", strField, nSave)
strField = Decode(strField)
strField = Mid$(strField, 7, iLength)

DecryptString = strField

End Function
Sub PlySound(strSound As String)
'Play a sound
Call sndPlaySound(App.Path & "\" & strSound & ".wav", 1)
End Sub
Sub PlayMidi(strMidi As String)
On Error Resume Next
strMidi = GetShortPath(App.Path) & "\" & strMidi & ".mid"
If strMidi = "" Then Exit Sub
Call mciSendString("play " & strMidi$, 0&, 0, 0)
End Sub
Sub StopMidi(strMidi As String)
On Error Resume Next
strMidi = GetShortPath(App.Path) & "\" & strMidi & ".mid"
If strMidi = "" Then Exit Sub
Call mciSendString("stop " & strMidi$, 0&, 0, 0)
End Sub
Public Function GetShortPath(strFileName As String) As String
    Dim lngRes As Long, strPath As String
    'Create a buffer
    strPath = String$(165, 0)
    'retrieve the short pathname
    lngRes = GetShortPathName(strFileName, strPath, 164)
    'remove all unnecessary chr$(0)'s
    GetShortPath = Left$(strPath, lngRes)
End Function
