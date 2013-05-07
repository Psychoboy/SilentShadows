VERSION 5.00
Begin VB.Form frmNpcEditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Npc Editor"
   ClientHeight    =   8250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6885
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
   ScaleHeight     =   8250
   ScaleWidth      =   6885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   5880
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   60
      Top             =   2760
      Width           =   480
   End
   Begin VB.HScrollBar HScroll5 
      Height          =   375
      Left            =   3960
      Max             =   600
      TabIndex        =   59
      Top             =   3240
      Width           =   2415
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   5880
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   56
      Top             =   1920
      Width           =   480
   End
   Begin VB.HScrollBar HScroll4 
      Height          =   375
      Left            =   3960
      Max             =   600
      TabIndex        =   55
      Top             =   2400
      Width           =   2415
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   5880
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   52
      Top             =   1080
      Width           =   480
   End
   Begin VB.HScrollBar HScroll3 
      Height          =   375
      Left            =   3960
      Max             =   600
      TabIndex        =   51
      Top             =   1560
      Width           =   2415
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   375
      Left            =   3960
      Max             =   1000
      TabIndex        =   44
      Top             =   6240
      Value           =   1
      Width           =   2295
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   375
      Left            =   3960
      Max             =   32000
      TabIndex        =   43
      Top             =   7080
      Value           =   1
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   390
      Left            =   3000
      TabIndex        =   41
      Text            =   "0"
      Top             =   5280
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Test"
      Height          =   375
      Left            =   3000
      TabIndex        =   40
      Top             =   3480
      Width           =   615
   End
   Begin VB.TextBox txtSpawnSecs 
      Alignment       =   1  'Right Justify
      Height          =   390
      Left            =   3000
      TabIndex        =   39
      Text            =   "0"
      Top             =   4440
      Width           =   1815
   End
   Begin VB.TextBox txtAttackSay 
      Height          =   390
      Left            =   960
      TabIndex        =   37
      Top             =   600
      Width           =   3975
   End
   Begin VB.Timer tmrSprite 
      Interval        =   50
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox picSprites 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   120
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   35
      Top             =   9360
      Width           =   480
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4200
      TabIndex        =   34
      Top             =   7800
      Width           =   2415
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   0
      TabIndex        =   33
      Top             =   7800
      Width           =   2415
   End
   Begin VB.HScrollBar scrlValue 
      Height          =   375
      Left            =   840
      Max             =   32000
      TabIndex        =   31
      Top             =   6960
      Width           =   2055
   End
   Begin VB.HScrollBar scrlNum 
      Height          =   375
      Left            =   840
      Max             =   1000
      TabIndex        =   25
      Top             =   6000
      Value           =   1
      Width           =   2055
   End
   Begin VB.TextBox txtChance 
      Alignment       =   1  'Right Justify
      Height          =   390
      Left            =   3000
      TabIndex        =   24
      Text            =   "0"
      Top             =   4920
      Width           =   1815
   End
   Begin VB.HScrollBar scrlMAGI 
      Height          =   375
      Left            =   600
      Max             =   32000
      TabIndex        =   20
      Top             =   3000
      Width           =   2055
   End
   Begin VB.HScrollBar scrlSPEED 
      Height          =   375
      Left            =   960
      Max             =   255
      TabIndex        =   17
      Top             =   3480
      Width           =   1575
   End
   Begin VB.HScrollBar scrlDEF 
      Height          =   375
      Left            =   600
      Max             =   1000
      TabIndex        =   14
      Top             =   2520
      Width           =   2055
   End
   Begin VB.HScrollBar scrlSTR 
      Height          =   375
      Left            =   600
      Max             =   1000
      TabIndex        =   11
      Top             =   2040
      Width           =   2055
   End
   Begin VB.HScrollBar scrlRange 
      Height          =   375
      Left            =   840
      Max             =   255
      TabIndex        =   8
      Top             =   1560
      Value           =   1
      Width           =   2055
   End
   Begin VB.ComboBox cmbBehavior 
      Height          =   390
      ItemData        =   "frmNpcEditor.frx":0000
      Left            =   1320
      List            =   "frmNpcEditor.frx":0013
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   3960
      Width           =   3615
   End
   Begin VB.TextBox txtName 
      Height          =   390
      Left            =   960
      TabIndex        =   2
      Top             =   120
      Width           =   3975
   End
   Begin VB.HScrollBar scrlSprite 
      Height          =   375
      Left            =   720
      Max             =   600
      TabIndex        =   1
      Top             =   1080
      Width           =   1455
   End
   Begin VB.PictureBox picSprite 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   3480
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   0
      Top             =   1080
      Width           =   480
   End
   Begin VB.Label LblSprite4 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   375
      Left            =   5280
      TabIndex        =   62
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label Label28 
      Caption         =   "Accessory"
      Height          =   375
      Left            =   3960
      TabIndex        =   61
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Lblsprite3 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   375
      Left            =   5280
      TabIndex        =   58
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label Label26 
      Caption         =   "Weapon"
      Height          =   375
      Left            =   3960
      TabIndex        =   57
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label lblsprite2 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   375
      Left            =   5280
      TabIndex        =   54
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label24 
      Caption         =   "Head"
      Height          =   375
      Left            =   3960
      TabIndex        =   53
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label23 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   375
      Left            =   4920
      TabIndex        =   50
      Top             =   6600
      Width           =   1215
   End
   Begin VB.Label Label22 
      Caption         =   "Num"
      Height          =   375
      Left            =   3360
      TabIndex        =   49
      Top             =   6240
      Width           =   615
   End
   Begin VB.Label Label21 
      Caption         =   "Item2:"
      Height          =   375
      Left            =   3360
      TabIndex        =   48
      Top             =   5760
      Width           =   615
   End
   Begin VB.Label Label20 
      Height          =   375
      Left            =   4080
      TabIndex        =   47
      Top             =   5760
      Width           =   2775
   End
   Begin VB.Label Label19 
      Caption         =   "Value"
      Height          =   375
      Left            =   3360
      TabIndex        =   46
      Top             =   7080
      Width           =   615
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   375
      Left            =   4680
      TabIndex        =   45
      Top             =   7440
      Width           =   1575
   End
   Begin VB.Label Label17 
      Caption         =   "Drop Item2 Chance 1 out of"
      Height          =   375
      Left            =   0
      TabIndex        =   42
      Top             =   5280
      Width           =   2895
   End
   Begin VB.Label Label16 
      Caption         =   "Spawn Rate (in seconds)"
      Height          =   375
      Left            =   0
      TabIndex        =   38
      Top             =   4440
      Width           =   2895
   End
   Begin VB.Label Label14 
      Caption         =   "Say"
      Height          =   375
      Left            =   120
      TabIndex        =   36
      Top             =   600
      Width           =   735
   End
   Begin VB.Label lblValue 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   375
      Left            =   1440
      TabIndex        =   32
      Top             =   7320
      Width           =   1455
   End
   Begin VB.Label Label7 
      Caption         =   "Value"
      Height          =   375
      Left            =   120
      TabIndex        =   30
      Top             =   6960
      Width           =   615
   End
   Begin VB.Label lblItemName 
      Height          =   375
      Left            =   720
      TabIndex        =   29
      Top             =   5760
      Width           =   2535
   End
   Begin VB.Label Label11 
      Caption         =   "Item1:"
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   5760
      Width           =   615
   End
   Begin VB.Label Label9 
      Caption         =   "Num"
      Height          =   375
      Left            =   0
      TabIndex        =   27
      Top             =   6000
      Width           =   615
   End
   Begin VB.Label lblNum 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   375
      Left            =   1560
      TabIndex        =   26
      Top             =   6480
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Drop Item Chance 1 out of"
      Height          =   375
      Left            =   0
      TabIndex        =   23
      Top             =   4920
      Width           =   2895
   End
   Begin VB.Label Label12 
      Caption         =   "Exp."
      Height          =   375
      Left            =   120
      TabIndex        =   22
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label lblMAGI 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   375
      Left            =   2520
      TabIndex        =   21
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label Label10 
      Caption         =   "Sound"
      Height          =   375
      Left            =   120
      TabIndex        =   19
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label lblSPEED 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   375
      Left            =   2520
      TabIndex        =   18
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label Label8 
      Caption         =   "DEF"
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label lblDEF 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   375
      Left            =   2520
      TabIndex        =   15
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "STR"
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label lblSTR 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   375
      Left            =   2640
      TabIndex        =   12
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Range"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label lblRange 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   375
      Left            =   2880
      TabIndex        =   9
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Behavior"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Name"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "Body"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label lblSprite 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   1080
      Width           =   1095
   End
End
Attribute VB_Name = "frmNpcEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Call PlaySound("\sfx\npc" & frmNpcEditor.scrlSPEED.value & ".wav")
End Sub

Private Sub HScroll1_Change()
Label18.Caption = STR(HScroll1.value)
End Sub

Private Sub HScroll2_Change()
Label23.Caption = STR(HScroll2.value)
    If HScroll2.value > 0 Then
        Label20.Caption = Trim(Item(HScroll2.value).name)
    End If
End Sub

Private Sub HScroll3_Change()
lblsprite2.Caption = STR(HScroll3.value)
End Sub

Private Sub HScroll4_Change()
 Lblsprite3.Caption = STR(HScroll4.value)
End Sub

Private Sub HScroll5_Change()
 LblSprite4.Caption = STR(HScroll5.value)
End Sub

Private Sub scrlSprite_Change()
    lblSprite.Caption = STR(scrlSprite.value)
End Sub

Private Sub scrlRange_Change()
    lblRange.Caption = STR(scrlRange.value)
End Sub

Private Sub scrlSTR_Change()
Dim ffff As Long
    lblSTR.Caption = STR(scrlSTR.value)
    'ffff = (scrlSTR.value * scrlDEF.value)
    'lblStartHP.Caption = ffff
    'ffff = (scrlSTR.value * scrlDEF.value * 2)
    'lblExpGiven.Caption = ffff
End Sub

Private Sub scrlDEF_Change()
    lblDEF.Caption = STR(scrlDEF.value)
    'lblStartHP.Caption = STR(scrlSTR.value * scrlDEF.value)
    'lblExpGiven.Caption = STR(scrlSTR.value * scrlDEF.value * 2)
End Sub

Private Sub scrlSpeed_Change()
    lblSPEED.Caption = STR(scrlSPEED.value)
    'lblExpGiven.Caption = STR(scrlSTR.value * scrlDEF.value * 2)
End Sub

Private Sub scrlMAGI_Change()
    lblMAGI.Caption = STR(scrlMAGI.value)
End Sub

Private Sub scrlNum_Change()
    lblNum.Caption = STR(scrlNum.value)
    If scrlNum.value > 0 Then
        lblItemName.Caption = Trim(Item(scrlNum.value).name)
    End If
End Sub

Private Sub scrlValue_Change()
    lblValue.Caption = STR(scrlValue.value)
End Sub

Private Sub cmdOk_Click()
    Call NpcEditorOk
End Sub

Private Sub cmdCancel_Click()
    Call NpcEditorCancel
End Sub

Private Sub tmrSprite_Timer()
    Call NpcEditorBltSprite
End Sub
