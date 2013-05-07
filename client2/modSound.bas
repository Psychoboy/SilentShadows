Attribute VB_Name = "modSound"
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function sndPlaySound2 Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Declare Function mciSendCommand Lib "winmm.dll" Alias "mciSendCommandA" (ByVal wDeviceID As Long, ByVal uMessage As Long, ByVal dwParam1 As Long, ByRef dwParam2 As Any) As Long
Private Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVallpstrBuffer As String, ByVal uLength As Long) As Long


Public Const SND_SYNC = &H0
Public Const SND_ASYNC = &H1
Public Const SND_NODEFAULT = &H2
Public Const SND_MEMORY = &H4
Public Const SND_LOOP = &H8
Public Const SND_NOSTOP = &H10
Private Type MCI_STATUS_PARAMS
   dwCallback As Long
   dwReturn As Long
   dwItem As Long
   dwTrack As Integer
End Type

Private Type MCI_OPEN_PARAMS
   dwCallback As Long
   wDeviceID As Long
   lpstrDeviceType As String
   lpstrElementName As String
   lpstrAlias As String
End Type

Private mciStatusParams As MCI_STATUS_PARAMS
Private mciOpenParams As MCI_OPEN_PARAMS
Private Const MCI_STATUS_MODE = &H4&
Private Const MCI_WAIT = &H2&
Private Const MCI_STATUS_ITEM = &H100&
Private Const MCI_OPEN = &H803
Private mciDeviceID As Integer
Private Const MCI_STATUS = &H814
Private Const MCI_OPEN_ELEMENT = &H200&
Dim songHandle As Long
Dim sampleHandle As Long
Dim sampleChannel As Long
Dim streamHandle As Long
Dim streamChannel As Long

Public Sub PlayMidi(Song As String)
'Dim i As Long
    'i = mciSendString("close all", 0, 0, 0)
    'If IHaveMusic = 0 Then Exit Sub
    'i = mciSendString("open " & Song & " type sequencer alias background", 0, 0, 0)
    'i = mciSendString("play background notify", 0, 0, frmMirage.hwnd)

    streamHandle = FSOUND_Stream_Open(Song, FSOUND_LOOP_NORMAL, 0, 0)
    Dim result As Boolean
    streamChannel = FSOUND_Stream_Play(FSOUND_FREE, streamHandle)


End Sub

Function MIDIPlaying() As Boolean
    mciStatusParams.dwItem = MCI_STATUS_MODE
        mciSendCommand 0, MCI_STATUS, MCI_WAIT Or MCI_STATUS_ITEM, mciStatusParams
       
       If mciStatusParams.dwReturn = 526 Then
        MIDIPlaying = True
        Exit Function
        Else
        MIDIPlaying = False
        Exit Function
       End If
End Function
Public Sub StopMidi()
Dim i As Long
  
    FSOUND_Stream_Stop streamHandle
'After a stream has been stopped, the channel is not active anymore
streamChannel = 0

End Sub

Public Sub PlaySound(Sound As String)
    If IHaveSFX = 0 Then Exit Sub
    Call sndPlaySound(App.Path & "\" & Sound, SND_ASYNC Or SND_NOSTOP)
End Sub
Public Sub PlaySound2(Sound As String)
    If IHaveSFX = 0 Then Exit Sub
    Sound = App.Path & Sound
    sampleHandle = FSOUND_Sample_Load(FSOUND_FREE, Sound, FSOUND_NORMAL, 0, 0)
    sampleChannel = FSOUND_PlaySound(FSOUND_FREE, sampleHandle)
    
    'Call sndPlaySound2(App.Path & "\" & Sound, SND_ASYNC Or SND_NOSTOP)
End Sub
Public Sub AmbS1(Sound As String)
    If IHaveSFX = 0 Then Exit Sub
    Sound = App.Path & Sound
    sampleHandle = FSOUND_Sample_Load(FSOUND_FREE, Sound, FSOUND_NORMAL, 0, 0)
    sampleChannel = FSOUND_PlaySound(FSOUND_FREE, sampleHandle)
    
    'Call sndPlaySound2(App.Path & "\" & Sound, SND_ASYNC Or SND_NOSTOP)
End Sub
Public Sub AmbS2(Sound As String)
    If IHaveSFX = 0 Then Exit Sub
    Sound = App.Path & Sound
    sampleHandle = FSOUND_Sample_Load(FSOUND_FREE, Sound, FSOUND_NORMAL, 0, 0)
    sampleChannel = FSOUND_PlaySound(FSOUND_FREE, sampleHandle)
    
    'Call sndPlaySound2(App.Path & "\" & Sound, SND_ASYNC Or SND_NOSTOP)
End Sub
Public Sub AmbS3(Sound As String)
    If IHaveSFX = 0 Then Exit Sub
    Sound = App.Path & Sound
    sampleHandle = FSOUND_Sample_Load(FSOUND_FREE, Sound, FSOUND_NORMAL, 0, 0)
    sampleChannel = FSOUND_PlaySound(FSOUND_FREE, sampleHandle)
    
    'Call sndPlaySound2(App.Path & "\" & Sound, SND_ASYNC Or SND_NOSTOP)
End Sub



