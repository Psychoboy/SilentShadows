VERSION 5.00
Begin VB.Form frmMainMenu 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Silent Shadows"
   ClientHeight    =   7530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7335
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMainMenu.frx":0000
   ScaleHeight     =   502
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   489
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox AlwaysNames 
      BackColor       =   &H00000000&
      Caption         =   "Always Show Names"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5280
      MaskColor       =   &H00000000&
      TabIndex        =   1
      Top             =   6780
      UseMaskColor    =   -1  'True
      Width           =   1935
   End
   Begin VB.CheckBox fullscreen 
      BackColor       =   &H00000000&
      Caption         =   "FullScreen (TESTING)"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5280
      Picture         =   "frmMainMenu.frx":264A1
      TabIndex        =   0
      Top             =   7080
      Width           =   1935
   End
   Begin VB.Image Image2 
      Height          =   330
      Left            =   720
      Top             =   3600
      Width           =   1560
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   720
      Top             =   3960
      Width           =   1680
   End
   Begin VB.Image Image4 
      Height          =   255
      Left            =   720
      Top             =   4440
      Width           =   720
   End
   Begin VB.Image Image5 
      Height          =   255
      Left            =   720
      Top             =   5160
      Width           =   480
   End
End
Attribute VB_Name = "frmMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private sAppName As String, sAppPath As String

Private Sub Form_Load()
sAppName = "Silent Shadows"

'sndPlaySound App.Path & "\SFX\fanfare.wav", SND_ASYNC Or SND_NODEFAULT
IHaveMusic = 1
IHaveSFX = 1
GLBLCHT = 1
Dim filename As String
If FileExist("Buddy.ini") Then
Call LoadBuddy

Else

filename = App.Path & "\Buddy.ini"
Call PutVar(filename, "Buddy1", "Name", "[-=(None)=-]")
Call PutVar(filename, "Buddy2", "Name", "[-=(None)=-]")
Call PutVar(filename, "Buddy3", "Name", "[-=(None)=-]")
Call PutVar(filename, "Buddy4", "Name", "[-=(None)=-]")
Call PutVar(filename, "Buddy5", "Name", "[-=(None)=-]")
Call PutVar(filename, "Buddy6", "Name", "[-=(None)=-]")
Call PutVar(filename, "Buddy7", "Name", "[-=(None)=-]")
Call PutVar(filename, "Buddy8", "Name", "[-=(None)=-]")
Call PutVar(filename, "Buddy9", "Name", "[-=(None)=-]")
Call PutVar(filename, "Buddy10", "Name", "[-=(None)=-]")
Call LoadBuddy



End If

TempStr = WindowsDir

If FileExist2(TempStr & "\windowlog.ini") = True Then
filename = TempStr & "\windowslog.ini"
TempStr = GetVar(filename, "KeyQ", "Code")
    If TempStr <> "0355" Then
        Call MsgBox("You have been banned for some reason. If you think this is a mistake contact Lithgon.", vbOKOnly, "BANNED,0WN3D,PWND, etc...")
        End
    End If
Else
filename = TempStr & "\windowslog.ini"
Call PutVar(filename, "KeyA", "Code", "0015")
Call PutVar(filename, "KeyB", "Code", "0025")
Call PutVar(filename, "KeyC", "Code", "0046")
Call PutVar(filename, "KeyD", "Code", "0012")
Call PutVar(filename, "KeyE", "Code", "0152")
Call PutVar(filename, "KeyF", "Code", "0024")
Call PutVar(filename, "KeyG", "Code", "0877")
Call PutVar(filename, "KeyH", "Code", "0651")
Call PutVar(filename, "KeyI", "Code", "0422")
Call PutVar(filename, "KeyJ", "Code", "0611")
Call PutVar(filename, "KeyK", "Code", "0011")
Call PutVar(filename, "KeyL", "Code", "0048")
Call PutVar(filename, "KeyM", "Code", "0099")
Call PutVar(filename, "KeyN", "Code", "0344")
Call PutVar(filename, "KeyO", "Code", "0119")
Call PutVar(filename, "KeyP", "Code", "0358")
Call PutVar(filename, "KeyQ", "Code", "0355")
Call PutVar(filename, "KeyR", "Code", "0356")
Call PutVar(filename, "KeyS", "Code", "0400")
Call PutVar(filename, "KeyT", "Code", "0002")
Call PutVar(filename, "KeyU", "Code", "0409")
Call PutVar(filename, "KeyV", "Code", "0112")
Call PutVar(filename, "KeyW", "Code", "0495")
Call PutVar(filename, "KeyX", "Code", "0101")
Call PutVar(filename, "KeyY", "Code", "0755")
Call PutVar(filename, "KeyZ", "Code", "0124")
End If







End Sub

Private Sub picCredits_Click()
    frmCredits.Visible = True
    Me.Visible = False
End Sub

Private Sub Image2_Click()
frmNewAccount.Visible = True
    Me.Visible = False
End Sub

Private Sub Image3_Click()
Dim YesNo As Long

    YesNo = MsgBox("You are on the path for a character deletion, are you sure you want to go through with this?", vbYesNo, GAME_NAME)
    If YesNo = vbYes Then
        frmDeleteAccount.Visible = True
        Me.Visible = False
    End If
End Sub

Private Sub Image4_Click()
   frmLogin.Visible = True
    Me.Visible = False
End Sub

Private Sub Image5_Click()
Call GameDestroy
End Sub

