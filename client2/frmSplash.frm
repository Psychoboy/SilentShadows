VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form3"
   ClientHeight    =   7560
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7305
   LinkTopic       =   "Form3"
   Picture         =   "frmSplash.frx":0000
   ScaleHeight     =   7560
   ScaleWidth      =   7305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   10000
      Left            =   360
      Top             =   480
   End
   Begin VB.PictureBox Picture1 
      Height          =   615
      Left            =   360
      Picture         =   "frmSplash.frx":1EB1B
      ScaleHeight     =   555
      ScaleWidth      =   915
      TabIndex        =   0
      Top             =   7800
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Click to skip"
      Height          =   255
      Left            =   3120
      TabIndex        =   1
      Top             =   7320
      Width           =   975
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public CountDownsplash As Byte
Private sAppName As String, sAppPath As String

Private Sub Form_Click()
If CountDownsplash = 0 Then
frmSplash.Picture = Picture1.Picture
CountDownsplash = CountDownsplash + 1
Else
'Call Main
frmMainMenu.Visible = True
frmSplash.Timer1.Enabled = False
frmSplash.Hide
End If
End Sub

Private Sub Form_Load()
Dim result As Boolean

result = FSOUND_Init(44100, 32, 0)
If result Then
    'Successfully initialized
    Else
    'An error occured
    MsgBox "An error occured initializing fmod!" & vbCrLf & _
        FSOUND_GetErrorString(FSOUND_GetError), vbOKOnly
End If
Call PlayMidi(App.Path & "\music\15.ogg")
CountDownsplash = 0



End Sub

Private Sub Label1_Click()
If CountDownsplash = 0 Then
frmSplash.Picture = Picture1.Picture
CountDownsplash = CountDownsplash + 1
Else
frmMainMenu.Visible = True
Timer1.Enabled = False
frmSplash.Hide
End If
End Sub

Private Sub Timer1_Timer()
If CountDownsplash = 0 Then
frmSplash.Picture = Picture1.Picture
CountDownsplash = CountDownsplash + 1
Else
frmMainMenu.Visible = True
Timer1.Enabled = False
frmSplash.Hide
End If
End Sub
