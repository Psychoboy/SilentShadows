VERSION 5.00
Begin VB.Form frmSellApprove 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Sell?"
   ClientHeight    =   1125
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3435
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1125
   ScaleWidth      =   3435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "NO"
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "YES"
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Do you really want to sell this Item?"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   360
      Width           =   2535
   End
   Begin VB.Label amount 
      Caption         =   "0"
      Height          =   255
      Left            =   1440
      TabIndex        =   1
      Top             =   0
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "This Item Sells For"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1335
   End
End
Attribute VB_Name = "frmSellApprove"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If frmInventory.LSTINV.ListCount > 0 Then
Call SendSellItem(frmInventory.LSTINV.ListIndex + 1)
End If
End Sub

Private Sub Command2_Click()
frmSellApprove.Hide
End Sub
