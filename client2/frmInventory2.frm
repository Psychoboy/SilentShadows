VERSION 5.00
Begin VB.Form frmInventory 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Inventory"
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3480
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   3480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picInv 
      Appearance      =   0  'Flat
      BackColor       =   &H00400000&
      ForeColor       =   &H80000008&
      Height          =   3405
      Left            =   0
      ScaleHeight     =   225
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   231
      TabIndex        =   0
      Top             =   0
      Width           =   3495
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   450
         Left            =   1080
         Picture         =   "frmInventory2.frx":0000
         ScaleHeight     =   450
         ScaleWidth      =   570
         TabIndex        =   3
         Top             =   2880
         Width           =   570
      End
      Begin VB.ListBox LSTINV 
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
         Height          =   2130
         ItemData        =   "frmInventory2.frx":336E
         Left            =   240
         List            =   "frmInventory2.frx":3370
         TabIndex        =   2
         Top             =   720
         Width           =   3015
      End
      Begin VB.VScrollBar invbar 
         Height          =   495
         Left            =   3240
         Max             =   50
         TabIndex        =   1
         Top             =   0
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Image Image2 
         Height          =   750
         Left            =   240
         Picture         =   "frmInventory2.frx":3372
         Top             =   0
         Width           =   3000
      End
      Begin VB.Image Image3 
         Height          =   450
         Left            =   0
         Picture         =   "frmInventory2.frx":7546
         Top             =   2880
         Width           =   570
      End
      Begin VB.Image Image5 
         Height          =   450
         Left            =   2880
         Picture         =   "frmInventory2.frx":A271
         Top             =   2880
         Width           =   570
      End
      Begin VB.Image ISELL 
         Height          =   450
         Left            =   1680
         Picture         =   "frmInventory2.frx":CFAE
         Top             =   2880
         Width           =   570
      End
      Begin VB.Image Image4 
         Height          =   450
         Left            =   600
         Picture         =   "frmInventory2.frx":FCFD
         Top             =   2880
         Width           =   570
      End
      Begin VB.Image Image23 
         Height          =   450
         Left            =   2280
         Picture         =   "frmInventory2.frx":12A7D
         Top             =   2880
         Width           =   570
      End
   End
End
Attribute VB_Name = "frmInventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Image23_Click()

If Item(GetPlayerInvItemNum(MyIndex, frmInventory.LSTINV.ListIndex + 1)).Type = ITEM_TYPE_CURRENCY Then
frmAmountSend.Show vbModal
Else

SendData ("GIVEITEM" & SEP_CHAR & frmInventory.LSTINV.ListIndex + 1 & SEP_CHAR & 0 & SEP_CHAR & END_CHAR)
End If
'sndPlaySound App.Path & "\SFX\close1.wav", SND_ASYNC Or SND_NODEFAULT
End Sub

Private Sub Image3_Click()
sndPlaySound App.Path & "\SFX\item1.wav", SND_ASYNC Or SND_NODEFAULT
Call SendUseItem(frmInventory.LSTINV.ListIndex + SELectINV + 1 + (frmInventory.invbar.value))
End Sub

Private Sub Inv1_Click()

End Sub

Private Sub Image4_Click()
sndPlaySound App.Path & "\SFX\item1.wav", SND_ASYNC Or SND_NODEFAULT
Dim value As Long
Dim InvNum As Long

    InvNum = frmInventory.LSTINV.ListIndex + SELectINV + 1 + (frmInventory.invbar.value)
    'If ITEM1X.Tag = frmMirage.LSTINV.ListIndex + SELectINV + 1 + (invbar.value) Then
    'ITEM1X.Tag = 0
    'ITEM1.Tag = 0
    'End If
    'If ITEM2X.Tag = frmMirage.LSTINV.ListIndex + SELectINV + 1 + (invbar.value) Then
    'ITEM2X.Tag = 0
    'ITEM2.Tag = 0
    'End If
    'If ITEM3X.Tag = frmMirage.LSTINV.ListIndex + SELectINV + 1 + (invbar.value) Then
    'ITEM3X.Tag = 0
    'ITEM3.Tag = 0
    'End If
    

    If GetPlayerInvItemNum(MyIndex, InvNum) > 0 And GetPlayerInvItemNum(MyIndex, InvNum) <= MAX_ITEMS Then
        If Item(GetPlayerInvItemNum(MyIndex, InvNum)).Type = ITEM_TYPE_CURRENCY Then
            ' Show them the drop dialog
            frmDrop.Show vbModal
        Else
            Call SendDropItem(frmInventory.LSTINV.ListIndex + SELectINV + 1 + (frmInventory.invbar.value), 0)
        End If
    End If
End Sub

Private Sub Image5_Click()
frmInventory.Hide
End Sub

Private Sub ISELL_Click()

 If Player(MyIndex).WeaponSlot = frmInventory.LSTINV.ListIndex + SELectINV + 1 + (frmInventory.invbar.value) Then
    Exit Sub
    End If
     If Player(MyIndex).ArmorSlot = frmInventory.LSTINV.ListIndex + SELectINV + 1 + (frmInventory.invbar.value) Then
    Exit Sub
    End If
     If Player(MyIndex).HelmetSlot = frmInventory.LSTINV.ListIndex + SELectINV + 1 + (frmInventory.invbar.value) Then
    Exit Sub
    End If
    If Player(MyIndex).ShieldSlot = frmInventory.LSTINV.ListIndex + SELectINV + 1 + (frmInventory.invbar.value) Then
    Exit Sub
    End If
   
    

sndPlaySound App.Path & "\SFX\item1.wav", SND_ASYNC Or SND_NODEFAULT
'Call SendSellItem(frmInventory.LSTINV.ListIndex + SELectINV + 1 + (frmInventory.invbar.value))
Call ApproveSellItem

End Sub
Private Sub ApproveSellItem()

frmSellApprove.amount = Val(Item(GetPlayerInvItemNum(MyIndex, frmInventory.LSTINV.ListIndex + 1)).SellValue)
frmSellApprove.Show vbModal
End Sub

Private Sub Label1_Click()
If frmInventory.LSTINV.ListCount > 0 Then
        Call SendData("itemdetails" & SEP_CHAR & frmInventory.LSTINV.ListIndex + 1 & SEP_CHAR & END_CHAR)
    End If
End Sub

Private Sub LSTINV_DblClick()
Call quicksl
End Sub

Private Sub Picture1_Click()
If frmInventory.LSTINV.ListCount > 0 Then
        Call SendData("itemdetails" & SEP_CHAR & frmInventory.LSTINV.ListIndex + 1 & SEP_CHAR & END_CHAR)
    End If
End Sub
