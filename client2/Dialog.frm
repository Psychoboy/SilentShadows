VERSION 5.00
Begin VB.Form Dialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Forget Spell"
   ClientHeight    =   1080
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3540
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1080
   ScaleWidth      =   3540
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2100
      TabIndex        =   1
      Top             =   540
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   540
      Width           =   1215
   End
   Begin VB.Label label1 
      Caption         =   "Are you sure you want to forget this spell?"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   60
      Width           =   2955
   End
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
Dialog.Hide
End Sub

Private Sub OKButton_Click()
If Player(MyIndex).Spell(frmSpells.lstSpells.ListIndex + 1) > 0 Then
Call SendData("forgetspell" & SEP_CHAR & frmSpells.lstSpells.ListIndex + 1 & SEP_CHAR & END_CHAR)
End If
Dialog.Hide
frmSpells.Hide
End Sub
