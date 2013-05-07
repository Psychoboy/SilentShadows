VERSION 5.00
Begin VB.Form frmMapProperties 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Map Properties"
   ClientHeight    =   7560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8355
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
   ScaleHeight     =   7560
   ScaleWidth      =   8355
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Caption         =   "Doesnt Respawn"
      Height          =   270
      Left            =   120
      TabIndex        =   52
      Top             =   6120
      Width           =   2175
   End
   Begin VB.TextBox Text6 
      Height          =   390
      Left            =   3000
      TabIndex        =   50
      Top             =   4920
      Width           =   855
   End
   Begin VB.TextBox Text5 
      Height          =   390
      Left            =   3000
      TabIndex        =   49
      Top             =   4320
      Width           =   855
   End
   Begin VB.TextBox Text4 
      Height          =   390
      Left            =   3000
      TabIndex        =   48
      Top             =   3720
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Height          =   390
      Left            =   840
      TabIndex        =   43
      Top             =   4920
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Height          =   390
      Left            =   840
      TabIndex        =   42
      Top             =   4320
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   390
      Left            =   840
      TabIndex        =   41
      Top             =   3720
      Width           =   1815
   End
   Begin VB.ComboBox cmbNpc 
      Height          =   390
      Index           =   14
      Left            =   4200
      Style           =   2  'Dropdown List
      TabIndex        =   40
      Top             =   7080
      Width           =   4095
   End
   Begin VB.ComboBox cmbNpc 
      Height          =   390
      Index           =   13
      Left            =   4200
      Style           =   2  'Dropdown List
      TabIndex        =   39
      Top             =   6600
      Width           =   4095
   End
   Begin VB.ComboBox cmbNpc 
      Height          =   390
      Index           =   12
      Left            =   4200
      Style           =   2  'Dropdown List
      TabIndex        =   38
      Top             =   6120
      Width           =   4095
   End
   Begin VB.ComboBox cmbNpc 
      Height          =   390
      Index           =   11
      Left            =   4200
      Style           =   2  'Dropdown List
      TabIndex        =   37
      Top             =   5640
      Width           =   4095
   End
   Begin VB.ComboBox cmbNpc 
      Height          =   390
      Index           =   10
      Left            =   4200
      Style           =   2  'Dropdown List
      TabIndex        =   36
      Top             =   5160
      Width           =   4095
   End
   Begin VB.ComboBox cmbNpc 
      Height          =   390
      Index           =   9
      Left            =   4200
      Style           =   2  'Dropdown List
      TabIndex        =   35
      Top             =   4680
      Width           =   4095
   End
   Begin VB.ComboBox cmbNpc 
      Height          =   390
      Index           =   8
      Left            =   4200
      Style           =   2  'Dropdown List
      TabIndex        =   34
      Top             =   4200
      Width           =   4095
   End
   Begin VB.ComboBox cmbNpc 
      Height          =   390
      Index           =   7
      Left            =   4200
      Style           =   2  'Dropdown List
      TabIndex        =   33
      Top             =   3720
      Width           =   4095
   End
   Begin VB.ComboBox cmbNpc 
      Height          =   390
      Index           =   6
      Left            =   4200
      Style           =   2  'Dropdown List
      TabIndex        =   32
      Top             =   3240
      Width           =   4095
   End
   Begin VB.ComboBox cmbNpc 
      Height          =   390
      Index           =   5
      Left            =   4200
      Style           =   2  'Dropdown List
      TabIndex        =   31
      Top             =   2760
      Width           =   4095
   End
   Begin VB.ComboBox cmbShop 
      Height          =   390
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   30
      Top             =   2280
      Width           =   2415
   End
   Begin VB.HScrollBar scrlMusic 
      Height          =   375
      Left            =   960
      Max             =   255
      TabIndex        =   26
      Top             =   2760
      Value           =   1
      Width           =   2415
   End
   Begin VB.TextBox txtBootY 
      Alignment       =   1  'Right Justify
      Height          =   390
      Left            =   3360
      TabIndex        =   24
      Text            =   "0"
      Top             =   6120
      Width           =   735
   End
   Begin VB.TextBox txtBootX 
      Alignment       =   1  'Right Justify
      Height          =   390
      Left            =   3360
      TabIndex        =   23
      Text            =   "0"
      Top             =   5520
      Width           =   735
   End
   Begin VB.TextBox txtBootMap 
      Alignment       =   1  'Right Justify
      Height          =   390
      Left            =   1320
      TabIndex        =   20
      Text            =   "0"
      Top             =   5520
      Width           =   735
   End
   Begin VB.ComboBox cmbNpc 
      Height          =   390
      Index           =   4
      ItemData        =   "frmMapProperties.frx":0000
      Left            =   4200
      List            =   "frmMapProperties.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   2280
      Width           =   4095
   End
   Begin VB.ComboBox cmbNpc 
      Height          =   390
      Index           =   3
      ItemData        =   "frmMapProperties.frx":0004
      Left            =   4200
      List            =   "frmMapProperties.frx":0006
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   1800
      Width           =   4095
   End
   Begin VB.ComboBox cmbNpc 
      Height          =   390
      Index           =   2
      Left            =   4200
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   1320
      Width           =   4095
   End
   Begin VB.ComboBox cmbNpc 
      Height          =   390
      Index           =   1
      Left            =   4200
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   840
      Width           =   4095
   End
   Begin VB.ComboBox cmbNpc 
      Height          =   390
      Index           =   0
      Left            =   4200
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   360
      Width           =   4095
   End
   Begin VB.ComboBox cmbMoral 
      Height          =   390
      ItemData        =   "frmMapProperties.frx":0008
      Left            =   960
      List            =   "frmMapProperties.frx":0015
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   1800
      Width           =   2415
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   7080
      Width           =   3975
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   6600
      Width           =   3975
   End
   Begin VB.TextBox txtLeft 
      Alignment       =   1  'Right Justify
      Height          =   390
      Left            =   2640
      TabIndex        =   3
      Text            =   "0"
      Top             =   720
      Width           =   735
   End
   Begin VB.TextBox txtRight 
      Alignment       =   1  'Right Justify
      Height          =   390
      Left            =   2640
      TabIndex        =   4
      Text            =   "0"
      Top             =   1200
      Width           =   735
   End
   Begin VB.TextBox txtDown 
      Alignment       =   1  'Right Justify
      Height          =   390
      Left            =   960
      TabIndex        =   2
      Text            =   "0"
      Top             =   1200
      Width           =   735
   End
   Begin VB.TextBox txtUp 
      Alignment       =   1  'Right Justify
      Height          =   390
      Left            =   960
      TabIndex        =   1
      Text            =   "0"
      Top             =   720
      Width           =   735
   End
   Begin VB.TextBox txtName 
      Height          =   390
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
   Begin VB.Line Line4 
      X1              =   120
      X2              =   3960
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   3960
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   3960
      Y1              =   5400
      Y2              =   5400
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   3960
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      Caption         =   "Interval"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   51
      Top             =   3360
      Width           =   855
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Caption         =   "File Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   47
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label Label15 
      Caption         =   "Ambient 3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   46
      Top             =   4920
      Width           =   855
   End
   Begin VB.Label Label14 
      Caption         =   "Ambient 2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   45
      Top             =   4320
      Width           =   855
   End
   Begin VB.Label Label13 
      Caption         =   "Ambient 1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   44
      Top             =   3720
      Width           =   855
   End
   Begin VB.Label Label12 
      Caption         =   "Shop"
      Height          =   375
      Left            =   120
      TabIndex        =   29
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Caption         =   "NPC's"
      Height          =   375
      Left            =   4200
      TabIndex        =   28
      Top             =   0
      Width           =   4095
   End
   Begin VB.Label lblMusic 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   375
      Left            =   3480
      TabIndex        =   27
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label Label10 
      Caption         =   "Music"
      Height          =   375
      Left            =   120
      TabIndex        =   25
      Top             =   2760
      Width           =   735
   End
   Begin VB.Label Label9 
      Caption         =   "Boot Y"
      Height          =   375
      Left            =   2280
      TabIndex        =   22
      Top             =   6120
      Width           =   975
   End
   Begin VB.Label Label8 
      Caption         =   "Boot X"
      Height          =   375
      Left            =   2280
      TabIndex        =   21
      Top             =   5520
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "Boot Map"
      Height          =   375
      Left            =   120
      TabIndex        =   19
      Top             =   5520
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "Moral"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "Right"
      Height          =   375
      Left            =   1920
      TabIndex        =   11
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Left"
      Height          =   375
      Left            =   1920
      TabIndex        =   10
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Down"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Up"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Name"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmMapProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim X As Long, y As Long, i As Long

    txtName.Text = Trim(Map.name)
    txtUp.Text = STR(Map.Up)
    txtDown.Text = STR(Map.Down)
    txtLeft.Text = STR(Map.Left)
    txtRight.Text = STR(Map.Right)
    cmbMoral.ListIndex = Map.Moral
    scrlMusic.value = Map.Music
    txtBootMap.Text = STR(Map.BootMap)
    txtBootX.Text = STR(Map.BootX)
    txtBootY.Text = STR(Map.BootY)
    
    
    
    cmbShop.AddItem "No Shop"
    For X = 1 To MAX_SHOPS
        cmbShop.AddItem X & ": " & Trim(Shop(X).name)
    Next X
    cmbShop.ListIndex = Map.Shop
    
    For X = 1 To MAX_MAP_NPCS
        cmbNpc(X - 1).AddItem "No NPC"
    Next X
    
    For y = 1 To MAX_NPCS
        For X = 1 To MAX_MAP_NPCS
            cmbNpc(X - 1).AddItem y & ": " & Trim(Npc(y).name)
        Next X
    Next y
    
    For i = 1 To MAX_MAP_NPCS
        cmbNpc(i - 1).ListIndex = Map.Npc(i)
    Next i
End Sub

Private Sub scrlMusic_Change()
    lblMusic.Caption = STR(scrlMusic.value)
End Sub

Private Sub cmdOk_Click()
Dim X As Long, y As Long, i As Long

    Map.name = txtName.Text
    Map.Up = Val(txtUp.Text)
    Map.Down = Val(txtDown.Text)
    Map.Left = Val(txtLeft.Text)
    Map.Right = Val(txtRight.Text)
    Map.Moral = cmbMoral.ListIndex
    Map.Music = scrlMusic.value
    Map.BootMap = Val(txtBootMap.Text)
    Map.BootX = Val(txtBootX.Text)
    Map.BootY = Val(txtBootY.Text)
    Map.Shop = cmbShop.ListIndex
    For i = 1 To MAX_MAP_NPCS
        Map.Npc(i) = cmbNpc(i - 1).ListIndex
    Next i
    
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

