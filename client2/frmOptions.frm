VERSION 5.00
Begin VB.Form frmOptions 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   6075
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3270
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   3270
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   2895
   End
   Begin VB.Image Image9 
      Height          =   840
      Left            =   120
      Picture         =   "frmOptions.frx":0000
      Top             =   4080
      Width           =   3105
   End
   Begin VB.Image Image6 
      Height          =   450
      Left            =   2520
      Picture         =   "frmOptions.frx":4FA0
      Top             =   2520
      Width           =   570
   End
   Begin VB.Image Image5 
      Height          =   450
      Left            =   2520
      Picture         =   "frmOptions.frx":79F0
      Top             =   1440
      Width           =   570
   End
   Begin VB.Image Image4 
      Height          =   450
      Left            =   120
      Picture         =   "frmOptions.frx":A440
      Top             =   2520
      Width           =   570
   End
   Begin VB.Image Image3 
      Height          =   450
      Left            =   120
      Picture         =   "frmOptions.frx":CFA3
      Top             =   1440
      Width           =   570
   End
   Begin VB.Image Image2 
      Height          =   840
      Left            =   120
      Picture         =   "frmOptions.frx":FB06
      Top             =   360
      Width           =   3105
   End
   Begin VB.Image Image1 
      Height          =   840
      Left            =   120
      Picture         =   "frmOptions.frx":13A9C
      Top             =   5160
      Width           =   3105
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "---"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1320
      TabIndex        =   3
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "SFX"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1200
      TabIndex        =   2
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Music"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1200
      TabIndex        =   1
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "---"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1320
      TabIndex        =   0
      Top             =   1680
      Width           =   615
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub

Private Sub HScroll1_Change()

End Sub

Private Sub Command2_Click()

End Sub

Private Sub Command3_Click()

End Sub

Private Sub Command4_Click()

End Sub

Private Sub Command5_Click()

End Sub

Private Sub MUSICSCRL_Change()
If MUSICSCRL.value = 1 Then
Label3.Caption = "On"
IHaveMusic = 1
Else
Label3.Caption = "Off"
IHaveMusic = 0
End If
End Sub

Private Sub SFXSCRL_Change()
If SFXSCRL.value = 1 Then
Label6.Caption = "On"
IHaveSFX = 1
Else
Label6.Caption = "Off"
IHaveSFX = 0
End If
End Sub

Private Sub Command6_Click()

End Sub

Private Sub Command7_Click()

End Sub

Private Sub Form_Load()
Label1.Caption = "Time:" & Time
End Sub

Private Sub Image1_Click()

If IHaveMusic = 0 Then
Call PlayMidi(9999)
Call StopMidi
FMUSIC_StopSong songHandle
FMUSIC_FreeSong songHandle
End If

If IHaveMusic = 1 Then result = FMUSIC_PlaySong(songHandle)
Unload Me
End Sub

Private Sub Image3_Click()
Label3.Caption = "Off"
IHaveMusic = 0
Call StopMidi
End Sub

Private Sub Image4_Click()
Label6.Caption = "Off"
IHaveSFX = 0
End Sub

Private Sub Image5_Click()
Label3.Caption = "On"
IHaveMusic = 1
End Sub

Private Sub Image6_Click()
Label6.Caption = "On"
IHaveSFX = 1
End Sub

Private Sub Image7_Click()
frmMirage.Height = 8595
End Sub

Private Sub Image8_Click()
frmMirage.Height = 10000
End Sub

Private Sub Image9_Click()
frmQSpell.Show
End Sub
