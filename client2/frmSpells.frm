VERSION 5.00
Begin VB.Form frmSpells 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Spells"
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3495
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   3495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picPlayerSpells 
      Appearance      =   0  'Flat
      BackColor       =   &H00400000&
      FillColor       =   &H00FFFF80&
      ForeColor       =   &H80000002&
      Height          =   3405
      Left            =   0
      ScaleHeight     =   225
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   231
      TabIndex        =   0
      Top             =   0
      Width           =   3495
      Begin VB.ListBox lstSpells 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1920
         ItemData        =   "frmSpells.frx":0000
         Left            =   120
         List            =   "frmSpells.frx":0002
         TabIndex        =   1
         Top             =   720
         Width           =   3255
      End
      Begin VB.Image Image1 
         Height          =   450
         Left            =   1380
         Picture         =   "frmSpells.frx":0004
         Top             =   2760
         Width           =   570
      End
      Begin VB.Image Image6 
         Height          =   750
         Left            =   240
         Picture         =   "frmSpells.frx":34B8
         Top             =   0
         Width           =   3000
      End
      Begin VB.Image Image7 
         Height          =   450
         Left            =   600
         Picture         =   "frmSpells.frx":7005
         Top             =   2760
         Width           =   570
      End
      Begin VB.Image Image8 
         Height          =   450
         Left            =   2160
         Picture         =   "frmSpells.frx":9BBA
         Top             =   2760
         Width           =   570
      End
   End
End
Attribute VB_Name = "frmSpells"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image1_Click()
Dialog.Show vbModal
End Sub

Private Sub Image7_Click()
If Player(MyIndex).Spell(lstSpells.ListIndex + 1) > 0 Then
        If GetTickCount > Player(MyIndex).AttackTimer + 1000 Then
            If Player(MyIndex).Moving = 0 Then
                Call SendData("cast" & SEP_CHAR & lstSpells.ListIndex + 1 & SEP_CHAR & END_CHAR)
                Player(MyIndex).Attacking = 1
                Player(MyIndex).AttackTimer = GetTickCount
                Player(MyIndex).CastedSpell = YES
            Else
                Call AddText("Cannot cast while walking!", BrightRed)
            End If
        End If
    Else
        Call AddText("No spell here.", BrightRed)
    End If
End Sub

Private Sub Image8_Click()
frmSpells.Hide
End Sub
