Attribute VB_Name = "basCOMMON"
Option Explicit
Private Declare Function sndPlaySound Lib "winmm" Alias "sndPlaySoundA" (lpszSoundName As Any, ByVal uFlags As Long) As Long
Private Const SND_SYNC = &H0        ' Play synchronously (default).
Private Const SND_NODEFAULT = &H2    ' Do not use default sound.
Private Const SND_MEMORY = &H4
Private Const SND_LOOP = &H8         ' Loop the sound until next
Private Const SND_NOSTOP = &H10      ' Do not stop any currently
Private Const SND_ASYNC = &H1          '  play asynchronously
Private bytSound() As Byte ' Always store binary data in byte arrays!
Public Enum SoundFlags
    soundSYNC = SND_SYNC
    soundNO_DEFAULT = SND_NODEFAULT
    soundMEMORY = SND_MEMORY
    soundLOOP = SND_LOOP
    soundNO_STOP = SND_NOSTOP
    soundASYNC = SND_ASYNC
End Enum
Public Enum AppSounds
    sndEGG = 101
    sndFOLD = 102
End Enum
Public Sub PlayWaveRes(vntResourceID As AppSounds, Optional vntFlags As SoundFlags = soundASYNC)
    bytSound = LoadResData(vntResourceID, "WAVE")
    If IsMissing(vntFlags) Then
        vntFlags = SND_NODEFAULT Or SND_SYNC Or SND_MEMORY
    End If
    If (vntFlags And SND_MEMORY) = 0 Then
        vntFlags = vntFlags Or SND_MEMORY
    End If
    sndPlaySound bytSound(0), vntFlags
End Sub
'------------------------------------------------------------
' Author:  Clint M. LaFever [clint.m.lafever@cpmx.saic.com]
' Purpose:  Gets the first token off a delimited string.
'                 Returns the token and changes the passed string
'                with the first token removed.
' Parameters:  delimited string and delimiter
' Returns:  First Token off delimited string
' Date: April,30 1999 @ 11:14:01
'------------------------------------------------------------
Function GetToken(sSource As String, ByVal sDelim As String) As String
   Dim iDelimPos As Integer
    On Error GoTo ErrorGetToken
   '------------------------------------------------------------
   ' Find the first delimiter
   '------------------------------------------------------------
   iDelimPos = InStr(1, sSource, sDelim)
   '------------------------------------------------------------
   ' If no delimiter was found, return the existing
   ' string and set the source to an empty string.
   '------------------------------------------------------------
   If (iDelimPos = 0) Then
      GetToken = Trim$(sSource)
      sSource = ""
   '------------------------------------------------------------
   ' Otherwise, return everything to the left of the
   ' delimiter and return the source string with it
   ' removed.
   '------------------------------------------------------------
   Else
      GetToken = Trim$(Left$(sSource, iDelimPos - 1))
      sSource = Mid$(sSource, iDelimPos + 1)
   End If
   Exit Function
ErrorGetToken:
    MsgBox Err & ":Error in GetToken() Function.  Error Message:" & Error(Err), 48, "Warning"
    Exit Function
End Function
'------------------------------------------------------------
' Author:  Clint M. LaFever [clint.m.lafever@cpmx.saic.com]
' Purpose:  Counts the number of tokens that are in a delimited
'                string.
' Parameters:  delimited string and delimiter
' Returns:  Number of tokens.
' Date: April,30 1999 @ 11:14:24
'------------------------------------------------------------
Function CountTokens(ByVal sSource As String, ByVal sDelim As String) As Integer
   Dim iDelimPos As Integer
   Dim iCount As Integer
    On Error GoTo ErrorCountTokens
    '------------------------------------------------------------
    ' Number of tokens = 0 if the source string is
    ' empty
    '------------------------------------------------------------
   If sSource = "" Then
      CountTokens = 0
   '------------------------------------------------------------
   ' Otherwise number of tokens = number of delimiters
   '  1
   '------------------------------------------------------------
   Else
      iDelimPos = InStr(1, sSource, sDelim)
      Do Until iDelimPos = 0
         iCount = iCount + 1
         iDelimPos = InStr(iDelimPos + 1, sSource, sDelim)
      Loop
      CountTokens = iCount + 1
   End If
   Exit Function
ErrorCountTokens:
    MsgBox Err & ":Error in CountTokens() Function.  Error Message:" & Error(Err), 48, "Warning"
    Exit Function
End Function
'------------------------------------------------------------
' Author:  Clint LaFever [lafeverc@saic.com]
' Date: October,04 2000 @ 11:45:31
'------------------------------------------------------------
Public Function CheckForDup(ByVal dSTR As String, sSTR As String, delim As String) As Boolean
    On Error GoTo ErrorCheckForDup
    Dim x As Long, f As Boolean
    f = False
    For x = 0 To CountTokens(dSTR, delim) - 1
        If UCase(GetToken(dSTR, delim)) = UCase(sSTR) Then
            f = True
            Exit For
        End If
    Next x
    CheckForDup = f
    Exit Function
ErrorCheckForDup:
    CheckForDup = False
    MsgBox Err & ":Error in CheckForDup.  Error Message: " & Err.Description, vbCritical, "Warning"
    Exit Function
End Function

