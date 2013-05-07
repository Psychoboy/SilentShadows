VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Quick Inventory"
   ClientHeight    =   2430
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3945
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   3945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Which Quick Slot Do You Want It In?"
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   3735
      Begin VB.OptionButton Option3 
         BackColor       =   &H00000000&
         Caption         =   "Item3"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   2520
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton option2 
         BackColor       =   &H00000000&
         Caption         =   "Item2"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1320
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00000000&
         Caption         =   "Item1"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Image Image1 
      Height          =   840
      Left            =   480
      Picture         =   "Form2.frx":0000
      Top             =   1560
      Width           =   3105
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   3975
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image1_Click()
tEmp = 0
If Option1.value = True Then

frmMirage.ITEM1.Tag = TEMPITEMNUM
frmMirage.ITEM1X.Tag = TEMPINVSLOT + 1
Unload Me
End If

If option2.value = True Then

frmMirage.ITEM2.Tag = TEMPITEMNUM
frmMirage.ITEM2X.Tag = TEMPINVSLOT + 1
Unload Me
End If

If Option3.value = True Then

frmMirage.ITEM3.Tag = TEMPITEMNUM
frmMirage.ITEM3X.Tag = TEMPINVSLOT + 1
Unload Me
End If

End Sub
