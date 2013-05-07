VERSION 5.00
Begin VB.Form frmTrade 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Silent Shadows (Trade)"
   ClientHeight    =   7275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8895
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   8895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "INFO"
      Height          =   855
      Left            =   6960
      TabIndex        =   1
      Top             =   5520
      Width           =   1935
   End
   Begin VB.ListBox lstTrade 
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
      ForeColor       =   &H00FFC0C0&
      Height          =   3810
      ItemData        =   "frmTrade.frx":0000
      Left            =   2400
      List            =   "frmTrade.frx":0002
      TabIndex        =   0
      Top             =   840
      Width           =   5655
   End
   Begin VB.Image Image1 
      Height          =   7800
      Left            =   0
      Picture         =   "frmTrade.frx":0004
      Top             =   0
      Width           =   2160
   End
   Begin VB.Image Image5 
      Height          =   840
      Left            =   3840
      Picture         =   "frmTrade.frx":956A
      Top             =   5520
      Width           =   3105
   End
   Begin VB.Image Image4 
      Height          =   840
      Left            =   3840
      Picture         =   "frmTrade.frx":CE73
      Top             =   6360
      Width           =   3105
   End
   Begin VB.Image Image3 
      Height          =   840
      Left            =   3840
      Picture         =   "frmTrade.frx":10E6E
      Top             =   4680
      Width           =   3105
   End
   Begin VB.Image Image2 
      Height          =   750
      Left            =   3840
      Picture         =   "frmTrade.frx":150A8
      Top             =   0
      Width           =   3000
   End
End
Attribute VB_Name = "frmTrade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If frmTrade.lstTrade.ListCount > 0 Then
        Call SendData("shopitemdetails" & SEP_CHAR & frmTrade.lstTrade.ListIndex + 1 & SEP_CHAR & END_CHAR)
    End If

End Sub

Private Sub Image3_Click()
Dim i As Long

    For i = 1 To MAX_INV
        If GetPlayerInvItemNum(MyIndex, i) > 0 Then
            frmFixItem.cmbItem.AddItem i & ": " & Trim(Item(GetPlayerInvItemNum(MyIndex, i)).Name)
        Else
            frmFixItem.cmbItem.AddItem "Unused Slot"
        End If
    Next i
    frmFixItem.cmbItem.ListIndex = 0
    frmFixItem.Show vbModal
End Sub

Private Sub Image4_Click()
 Unload Me
End Sub

Private Sub Image5_Click()
If lstTrade.ListCount > 0 Then
        Call SendData("traderequest" & SEP_CHAR & lstTrade.ListIndex + 1 & SEP_CHAR & END_CHAR)
    End If
End Sub

Private Sub Label1_Click()

End Sub

