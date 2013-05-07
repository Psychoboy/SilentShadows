Attribute VB_Name = "meow"
Private Declare Function sndPlaySound2 Lib "winmm.dll" Alias "sndPlaySoundB" (ByVal lpszSoundName2 As String, ByVal uFlags2 As Long) As Long

Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long


Const SND_ASYNC = &H1       'ASYNC allows us to play waves with the ability to interrupt
Const SND_LOOP = &H8        'LOOP causes to sound to be continuously replayed
Const SND_NODEFAULT = &H2   'NODEFAULT causes no sound to be played if the wav can't be found
Const SND_SYNC = &H0        'SYNC plays a wave file without returning control to the calling program until it's finished
Const SND_NOSTOP = &H10     'NOSTOP ensures that we don't stop another wave from playing
Const SND_MEMORY = &H4      'MEMORY plays a wave file stored in memory


