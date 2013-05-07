Attribute VB_Name = "REs"
Private Const CCDEVICENAME = 32
Private Const CCFORMNAME = 32
Private Const DM_PELSWIDTH = &H80000
Private Const DM_PELSHEIGHT = &H100000

Private Declare Function EnumDisplaySettings Lib "user32" _
    Alias "EnumDisplaySettingsA" _
    (ByVal lpszDeviceName As Long, _
    ByVal iModeNum As Long, _
    lpDevMode As Any) As Boolean
    
Private Declare Function ChangeDisplaySettings Lib "user32" _
    Alias "ChangeDisplaySettingsA" _
    (lpDevMode As Any, _
    ByVal dwflags As Long) As Long

Private Type DEVMODE
    dmDeviceName As String * CCDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type


'this function below will change the resolution
Sub ChangeRes(iWidth As Single, iHeight As Single)
    Dim blnWorked As Boolean
    Dim i As Long
    Dim DevM As DEVMODE
    
    i = 0
    
    Do
        blnWorked = EnumDisplaySettings(0&, i, DevM)
        i = i + 1
    Loop Until (blnWorked = False)
        
    With DevM
        .dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT
        .dmPelsWidth = iWidth
        .dmPelsHeight = iHeight
    End With
    Call ChangeDisplaySettings(DevM, 0)
End Sub



