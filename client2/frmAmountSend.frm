VERSION 5.00
Begin VB.Form frmAmountSend 
   Caption         =   "Form3"
   ClientHeight    =   945
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   ScaleHeight     =   945
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   435
      Left            =   3660
      TabIndex        =   2
      Top             =   420
      Width           =   915
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   720
      TabIndex        =   0
      Text            =   "0"
      Top             =   540
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "How Much?"
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   240
      Width           =   2655
   End
End
Attribute VB_Name = "frmAmountSend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSend_Click()
If Text1.Text < 1 Then Text1.Text = 1

SendData ("GIVEITEM" & SEP_CHAR & frmInventory.LSTINV.ListIndex + 1 & SEP_CHAR & Text1.Text & SEP_CHAR & END_CHAR)
frmAmountSend.Hide
End Sub

