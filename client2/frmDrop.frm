VERSION 5.00
Begin VB.Form frmDrop 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Silent Shadows (Drop Item)"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3690
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
   ScaleHeight     =   3765
   ScaleWidth      =   3690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "-100"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2040
      TabIndex        =   9
      Top             =   1560
      Width           =   615
   End
   Begin VB.Image Image8 
      Height          =   450
      Left            =   2040
      Picture         =   "frmDrop.frx":0000
      Top             =   1440
      Width           =   570
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "+100"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2040
      TabIndex        =   8
      Top             =   1080
      Width           =   615
   End
   Begin VB.Image Image7 
      Height          =   450
      Left            =   2040
      Picture         =   "frmDrop.frx":27A3
      Top             =   960
      Width           =   570
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "+10"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1080
      TabIndex        =   7
      Top             =   1080
      Width           =   615
   End
   Begin VB.Image Image6 
      Height          =   450
      Left            =   1080
      Picture         =   "frmDrop.frx":4F46
      Top             =   960
      Width           =   570
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "-10"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1080
      TabIndex        =   6
      Top             =   1560
      Width           =   615
   End
   Begin VB.Image Image5 
      Height          =   450
      Left            =   1080
      Picture         =   "frmDrop.frx":76E9
      Top             =   1440
      Width           =   570
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "-1"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   615
   End
   Begin VB.Image Image4 
      Height          =   450
      Left            =   120
      Picture         =   "frmDrop.frx":9E8C
      Top             =   1440
      Width           =   570
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "+1"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   615
   End
   Begin VB.Image Image3 
      Height          =   450
      Left            =   120
      Picture         =   "frmDrop.frx":C62F
      Top             =   960
      Width           =   570
   End
   Begin VB.Image Image2 
      Height          =   840
      Left            =   360
      Picture         =   "frmDrop.frx":EDD2
      Top             =   2880
      Width           =   3105
   End
   Begin VB.Image Image1 
      Height          =   840
      Left            =   360
      Picture         =   "frmDrop.frx":12DCD
      Top             =   2160
      Width           =   3105
   End
   Begin VB.Label lblAmmount 
      BackColor       =   &H00000000&
      Caption         =   "1"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Ammount"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Item"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lblName 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "+1K"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3000
      TabIndex        =   11
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "-1K"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3000
      TabIndex        =   10
      Top             =   1560
      Width           =   615
   End
   Begin VB.Image Image9 
      Height          =   450
      Left            =   3000
      Picture         =   "frmDrop.frx":16112
      Top             =   1440
      Width           =   570
   End
   Begin VB.Image Image10 
      Height          =   450
      Left            =   3000
      Picture         =   "frmDrop.frx":188B5
      Top             =   960
      Width           =   570
   End
End
Attribute VB_Name = "frmDrop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Ammount As Long

Private Sub Form_Load()
Dim InvNum As Long

    Ammount = 1
    InvNum = frmInventory.LSTINV.ListIndex + SELectINV + 1 + (frmInventory.invbar.value)
    
    frmDrop.lblName = Trim(Item(GetPlayerInvItemNum(MyIndex, InvNum)).name)
    Call ProcessAmmount
End Sub

Private Sub cmdOk_Click()

End Sub

Private Sub cmdCancel_Click()
    
End Sub

Private Sub cmdPlus1_Click()
    Ammount = Ammount + 1
    Call ProcessAmmount
End Sub

Private Sub cmdMinus1_Click()
    Ammount = Ammount - 1
    Call ProcessAmmount
End Sub

Private Sub cmdPlus10_Click()
    Ammount = Ammount + 10
    Call ProcessAmmount
End Sub

Private Sub cmdMinus10_Click()
    Ammount = Ammount - 10
    Call ProcessAmmount
End Sub

Private Sub cmdPlus100_Click()
    Ammount = Ammount + 100
    Call ProcessAmmount
End Sub

Private Sub cmdMinus100_Click()
    Ammount = Ammount - 100
    Call ProcessAmmount
End Sub

Private Sub cmdPlus1000_Click()
    Ammount = Ammount + 1000
    Call ProcessAmmount
End Sub

Private Sub cmdMinus1000_Click()
    Ammount = Ammount - 1000
    Call ProcessAmmount
End Sub

Private Sub ProcessAmmount()
Dim InvNum As Long

    InvNum = frmInventory.LSTINV.ListIndex + SELectINV + 1 + (frmInventory.invbar.value)
        
    ' Check if more then max and set back to max if so
    If Ammount > GetPlayerInvItemValue(MyIndex, InvNum) Then
        Ammount = GetPlayerInvItemValue(MyIndex, InvNum)
    End If
    
    ' Make sure its not 0
    If Ammount <= 0 Then
        Ammount = 1
    End If

    frmDrop.lblAmmount.Caption = Ammount & "/" & GetPlayerInvItemValue(MyIndex, InvNum)
End Sub

Private Sub Image1_Click()
Dim InvNum As Long

    InvNum = frmInventory.LSTINV.ListIndex + SELectINV + 1 + (frmInventory.invbar.value)
    
    Call SendDropItem(InvNum, Ammount)
    Unload Me
End Sub

Private Sub Image10_Click()
Ammount = Ammount + 1000
    Call ProcessAmmount
End Sub

Private Sub Image2_Click()
Unload Me
End Sub

Private Sub Image3_Click()
Ammount = Ammount + 1
    Call ProcessAmmount
End Sub

Private Sub Image4_Click()
Ammount = Ammount - 1
    Call ProcessAmmount
End Sub

Private Sub Image5_Click()
Ammount = Ammount - 10
    Call ProcessAmmount
End Sub

Private Sub Image6_Click()
Ammount = Ammount + 10
    Call ProcessAmmount
End Sub

Private Sub Image7_Click()
Ammount = Ammount + 100
    Call ProcessAmmount
End Sub

Private Sub Image8_Click()
Ammount = Ammount - 100
    Call ProcessAmmount
End Sub

Private Sub Image9_Click()
Ammount = Ammount - 1000
    Call ProcessAmmount
End Sub

Private Sub Label10_Click()
Ammount = Ammount + 1000
    Call ProcessAmmount
End Sub

Private Sub Label3_Click()
Ammount = Ammount + 1
    Call ProcessAmmount
End Sub

Private Sub Label4_Click()
Ammount = Ammount - 1
    Call ProcessAmmount
End Sub

Private Sub Label5_Click()
Ammount = Ammount - 10
    Call ProcessAmmount
End Sub

Private Sub Label6_Click()
Ammount = Ammount + 10
    Call ProcessAmmount
End Sub

Private Sub Label7_Click()
Ammount = Ammount + 100
    Call ProcessAmmount
End Sub

Private Sub Label8_Click()
Ammount = Ammount - 100
    Call ProcessAmmount
End Sub

Private Sub Label9_Click()
Ammount = Ammount - 1000
    Call ProcessAmmount
End Sub

