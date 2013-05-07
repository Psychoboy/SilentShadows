VERSION 5.00
Begin VB.Form frmFishing 
   Caption         =   "Fishing Editor"
   ClientHeight    =   2565
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   ScaleHeight     =   2565
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Submit"
      Height          =   435
      Left            =   1140
      TabIndex        =   6
      Top             =   2100
      Width           =   2115
   End
   Begin VB.HScrollBar HScroll3 
      Height          =   255
      Left            =   120
      Max             =   1000
      TabIndex        =   5
      Top             =   1800
      Width           =   4515
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   255
      Left            =   120
      Max             =   1000
      TabIndex        =   3
      Top             =   1140
      Width           =   4515
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   315
      Left            =   120
      Max             =   1000
      TabIndex        =   0
      Top             =   480
      Width           =   4515
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   255
      Left            =   60
      TabIndex        =   4
      Top             =   1440
      Width           =   4515
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   4455
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   180
      TabIndex        =   1
      Top             =   60
      Width           =   4455
   End
End
Attribute VB_Name = "frmFishing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Fishing1 = HScroll1.value
Fishing2 = HScroll2.value
Fishing3 = HScroll3.value
frmFishing.Hide
End Sub

Private Sub HScroll1_Change()
Label1.Caption = HScroll1.value & ". " & Item(HScroll1.value).name
End Sub

Private Sub HScroll2_Change()
Label2.Caption = HScroll2.value & ". " & Item(HScroll2.value).name
End Sub

Private Sub HScroll3_Change()
Label3.Caption = HScroll3.value & ". " & Item(HScroll3.value).name
End Sub
