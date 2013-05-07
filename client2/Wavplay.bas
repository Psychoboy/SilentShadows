Attribute VB_Name = "Wavplay"
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Const SND_SYNC = &H0         '  play synchronously (default)
Const SND_NODEFAULT = &H2



