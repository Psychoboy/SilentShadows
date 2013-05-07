VERSION 5.00
Begin VB.Form frmChars 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Silent Shadows (Characters)"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6975
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   6975
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picDelChar 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   840
      Left            =   3000
      Picture         =   "frmChars.frx":0000
      ScaleHeight     =   840
      ScaleWidth      =   3105
      TabIndex        =   5
      Top             =   3720
      Width           =   3105
   End
   Begin VB.PictureBox picUseChar 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   840
      Left            =   3000
      Picture         =   "frmChars.frx":5158
      ScaleHeight     =   840
      ScaleWidth      =   3105
      TabIndex        =   4
      Top             =   2040
      Width           =   3105
   End
   Begin VB.PictureBox picNewChar 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   840
      Left            =   3000
      Picture         =   "frmChars.frx":A6B7
      ScaleHeight     =   840
      ScaleWidth      =   3105
      TabIndex        =   3
      Top             =   2880
      Width           =   3105
   End
   Begin VB.PictureBox picCancel 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   840
      Left            =   3000
      Picture         =   "frmChars.frx":FAFC
      ScaleHeight     =   840
      ScaleWidth      =   3105
      TabIndex        =   2
      Top             =   4560
      Width           =   3105
   End
   Begin VB.ListBox lstChars 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   1170
      ItemData        =   "frmChars.frx":13AF7
      Left            =   2280
      List            =   "frmChars.frx":13AF9
      TabIndex        =   1
      Top             =   720
      Width           =   4575
   End
   Begin VB.PictureBox picNewAccount 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   3000
      Picture         =   "frmChars.frx":13AFB
      ScaleHeight     =   750
      ScaleWidth      =   3000
      TabIndex        =   0
      Top             =   0
      Width           =   3000
   End
   Begin VB.Image Image1 
      Height          =   6000
      Left            =   0
      Picture         =   "frmChars.frx":180DE
      Top             =   0
      Width           =   2160
   End
End
Attribute VB_Name = "frmChars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub picCancel_Click()
    Call TcpDestroy
    frmLogin.Visible = True
    Me.Visible = False
End Sub

Private Sub picNewChar_Click()
    Call MenuState(MENU_STATE_NEWCHAR)
End Sub

Private Sub picUseChar_Click()
If frmMainMenu.fullscreen.value = 1 Then Call ChangeRes(800, 600)
    Call StopMidi
    Call MenuState(MENU_STATE_USECHAR)
End Sub

Private Sub picDelChar_Click()
Dim value As Long

    value = MsgBox("Are you sure you wish to delete this character?", vbYesNo, GAME_NAME)
    If value = vbYes Then
        Call MenuState(MENU_STATE_DELCHAR)
    End If
End Sub

