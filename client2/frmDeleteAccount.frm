VERSION 5.00
Begin VB.Form frmDeleteAccount 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Silent Shadows (Delete Account)"
   ClientHeight    =   4995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6975
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   6975
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   390
      IMEMode         =   3  'DISABLE
      Left            =   3600
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   3000
      Width           =   3375
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   390
      Left            =   3600
      MaxLength       =   20
      TabIndex        =   0
      Top             =   2400
      Width           =   3375
   End
   Begin VB.Image Image1 
      Height          =   6000
      Left            =   0
      Picture         =   "frmDeleteAccount.frx":0000
      Top             =   0
      Width           =   2160
   End
   Begin VB.Image Image4 
      Height          =   840
      Left            =   3120
      Picture         =   "frmDeleteAccount.frx":8DC7
      Top             =   4200
      Width           =   3105
   End
   Begin VB.Image Image3 
      Height          =   840
      Left            =   3120
      Picture         =   "frmDeleteAccount.frx":CDC2
      Top             =   3480
      Width           =   3105
   End
   Begin VB.Image Image2 
      Height          =   750
      Left            =   3120
      Picture         =   "frmDeleteAccount.frx":10CF9
      Top             =   0
      Width           =   3000
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter a account name and password of the account you wish to delete."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   975
      Left            =   2280
      TabIndex        =   3
      Top             =   1080
      Width           =   4575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   2400
      Width           =   855
   End
End
Attribute VB_Name = "frmDeleteAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Image3_Click()
If Trim(txtName.Text) <> "" And Trim(txtPassword.Text) <> "" Then
        Call MenuState(MENU_STATE_DELACCOUNT)
    End If
End Sub

Private Sub Image4_Click()
    frmMainMenu.Visible = True
    frmDeleteAccount.Visible = False
End Sub
