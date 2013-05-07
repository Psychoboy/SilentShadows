VERSION 5.00
Begin VB.Form frmBank 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Bank"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7320
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   7320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtAmount 
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Text            =   "1"
      Top             =   2160
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3120
      TabIndex        =   3
      Top             =   2880
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3120
      TabIndex        =   2
      Top             =   840
      Width           =   615
   End
   Begin VB.ListBox lstInventory 
      Height          =   3960
      Left            =   3840
      TabIndex        =   1
      Top             =   480
      Width           =   3255
   End
   Begin VB.ListBox lstBank 
      Height          =   3765
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   2775
   End
   Begin VB.Label Label3 
      Caption         =   "Inventory"
      Height          =   255
      Left            =   5040
      TabIndex        =   7
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Bank"
      Height          =   255
      Left            =   1080
      TabIndex        =   6
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Amount:"
      Height          =   255
      Left            =   3120
      TabIndex        =   5
      Top             =   1920
      Width           =   615
   End
End
Attribute VB_Name = "frmBank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

If txtAmount.Text < 1 Then
txtAmount.Text = 1
End If
If lstBank.ListCount > 0 Then
        Call SendData("takebank" & SEP_CHAR & lstBank.ListIndex + 1 & SEP_CHAR & txtAmount.Text & SEP_CHAR & END_CHAR)
    End If
    frmBank.Hide
Call SendData("bank" & SEP_CHAR & END_CHAR)
End Sub

Private Sub Command2_Click()
If txtAmount.Text < 1 Then
txtAmount.Text = 1
End If
If lstInventory.ListCount > 0 Then
        Call SendData("givebank" & SEP_CHAR & lstInventory.ListIndex + 1 & SEP_CHAR & txtAmount.Text & SEP_CHAR & END_CHAR)
    End If
frmBank.Hide
Call SendData("bank" & SEP_CHAR & END_CHAR)
End Sub

