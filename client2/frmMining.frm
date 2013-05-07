VERSION 5.00
Begin VB.Form frmMining 
   Caption         =   "Form3"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Submit"
      Height          =   435
      Left            =   1080
      TabIndex        =   6
      Top             =   2760
      Width           =   1875
   End
   Begin VB.HScrollBar HScroll3 
      Height          =   375
      Left            =   180
      Max             =   1000
      TabIndex        =   5
      Top             =   2340
      Width           =   4155
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   375
      Left            =   180
      Max             =   1000
      TabIndex        =   3
      Top             =   1380
      Width           =   4095
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   315
      Left            =   240
      Max             =   1000
      TabIndex        =   1
      Top             =   540
      Width           =   4095
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   375
      Left            =   180
      TabIndex        =   4
      Top             =   1800
      Width           =   4095
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   315
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   3975
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   180
      Width           =   3915
   End
End
Attribute VB_Name = "frmMining"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Mining1 = HScroll1.value
Mining2 = HScroll2.value
Mining3 = HScroll3.value
frmMining.Hide
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
