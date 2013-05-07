Attribute VB_Name = "MIDIExtended"
Option Explicit

Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long
Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

'BAS file to play MIDI's  by Wpsjr1@syix.com
' www.syix.com/wpsjr1/index.html
'usage:
'
'
'Private Sub cmdPlay_Click()  'PLAY MIDI
'Call PlayMIDI("c:\midi\Grabbag.mid")
'End Sub
'
'Private Sub cmdStop_Click()  'STOP MIDI
'Call StopMIDI
'End Sub
'
'Private Sub cmdLoop_Click()  'LOOP MIDI
'Call PlayMIDI("c:\midi\grabbag.mid", True)
'End Sub
'
'Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'Call StopMIDI
'End Sub
'
'____________________________________________________________


Public Function PlayMIDI2(DriveDirFile As String, Optional loopIT As Boolean)

  Dim returnStr As String * 255
  Dim Shortpath As String
  Shortpath = Space(Len(DriveDirFile))
  Shortpath = DriveDirFile
 Dim x As Long
  'x = GetShortPathName(DriveDirFile, Shortpath, Len(Shortpath))
  'If x = 0 Then GoTo errorhandler

  'If x > Len(DriveDirFile) Then 'not a long filename
  '  Shortpath = DriveDirFile
 ' Else                          'it is a long filename
  '  Shortpath = Left(Shortpath, x) 'x is the length of the return buffer
  'End If

  x = mciSendString("close yada", returnStr, 255, 0) 'just in case
  'x = mciSendString("open " & Shortpath & " type sequencer alias background", 0, 0, 0)
  'x = mciSendString("play background ", 0, 0, frmMirage.hWnd)
  x = mciSendString("open " & Chr(34) & Shortpath & Chr(34) & " type sequencer alias yada", returnStr, 255, 0)

  If x <> 0 Then GoTo theEnd  'invalid filename or path

  x = mciSendString("play yada", returnStr, 255, 0)
  
  If x <> 0 Then GoTo theEnd  'device busy or not ready
  
  If Not loopIT Then Exit Function
 
  Do While DoEvents
    x = mciSendString("status yada mode", returnStr, 255, 0)
    If x <> 0 Then Exit Function 'StopMIDI() was pressed or error
    If Left(returnStr, 7) = "stopped" Then x = mciSendString("play yada from 1", returnStr, 255, 0)
  Loop
  
  Exit Function

theEnd:  'MIDI errorhandler
  returnStr = Space(255)
  x = mciGetErrorString(x, returnStr, 255)
  MsgBox Trim(returnStr), vbExclamation 'error message
  x = mciSendString("close yada", returnStr, 255, 0)
  Exit Function

errorhandler:
  MsgBox "Invalid Filename or Error.", vbInformation
End Function

Public Function StopMIDI2()
  Dim x&
  Dim returnStr As String * 255

  x = mciSendString("status yada mode", returnStr, 255, 0)
  If Left(returnStr, 7) = "playing" Then x = mciSendString("stop yada", returnStr, 255, 0)

  returnStr = Space(255)

  x = mciSendString("status yada mode", returnStr, 255, 0)
  If Left(returnStr, 7) = "stopped" Then x = mciSendString("close yada", returnStr, 255, 0)
End Function
