VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmServer 
   Caption         =   "Mirage Server"
   ClientHeight    =   4650
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   8325
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   8325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "End Event"
      Height          =   255
      Left            =   3720
      TabIndex        =   7
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Send Event"
      Height          =   255
      Left            =   3720
      TabIndex        =   6
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   3240
      Width           =   7575
   End
   Begin VB.Timer PlayerTimer 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   2160
      Top             =   240
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Global Off"
      Height          =   495
      Left            =   7680
      TabIndex        =   3
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Global On"
      Height          =   495
      Left            =   7680
      TabIndex        =   2
      Top             =   0
      Width           =   615
   End
   Begin VB.Timer tmrShutdown 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1680
      Top             =   240
   End
   Begin VB.Timer tmrGameAI 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1200
      Top             =   240
   End
   Begin VB.Timer tmrSpawnMapItems 
      Interval        =   1000
      Left            =   720
      Top             =   240
   End
   Begin VB.Timer tmrPlayerSave 
      Interval        =   60000
      Left            =   240
      Top             =   240
   End
   Begin VB.TextBox txtChat 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   2520
      Width           =   7455
   End
   Begin VB.TextBox txtText 
      Height          =   2295
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   7455
   End
   Begin MSWinsockLib.Winsock Socket 
      Index           =   0
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label E 
      Caption         =   "Event"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuShutdown 
         Caption         =   "Shutdown"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuDatabase 
      Caption         =   "&Database"
      Begin VB.Menu ClockOff 
         Caption         =   "Clock On"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuReloadClasses 
         Caption         =   "Reload Classes"
      End
      Begin VB.Menu mnuReloadRaces 
         Caption         =   "Reload Races"
      End
      Begin VB.Menu mnuReloadVerson 
         Caption         =   "Reload Version"
      End
   End
   Begin VB.Menu mnuLog 
      Caption         =   "&Log"
      Begin VB.Menu mnuServerLog 
         Caption         =   "Server Log"
      End
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ClockOff_Click()
If ClockOff.Checked = False Then
ClockOff.Checked = True
Else
ClockOff.Checked = False
End If
End Sub

Private Sub Command1_Click()
CanGlobal = 1
End Sub

Private Sub Command2_Click()
CanGlobal = 0
End Sub

Private Sub Command3_Click()
Call SendDataToAll("EVENT" & SEP_CHAR & Text1.text & SEP_CHAR & END_CHAR)
End Sub

Private Sub Command4_Click()
Call SendDataToAll("EVENT" & SEP_CHAR & "endz1" & SEP_CHAR & END_CHAR)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim lmsg As Long
    
    lmsg = X / Screen.TwipsPerPixelX
    Select Case lmsg
        Case WM_LBUTTONDBLCLK
            frmServer.WindowState = vbNormal
            frmServer.Show
    End Select
End Sub

Private Sub Form_Resize()
    If frmServer.WindowState = vbMinimized Then
        frmServer.Hide
    End If
End Sub

Private Sub Form_Terminate()
    Call DestroyServer
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call DestroyServer
End Sub

Private Sub mnuServerLog_Click()
    If mnuServerLog.Checked = True Then
        mnuServerLog.Checked = False
        ServerLog = False
    Else
        mnuServerLog.Checked = True
        ServerLog = True
    End If
End Sub

Private Sub PlayerTimer_Timer()
Dim i As Long

If PlayerI <= MAX_PLAYERS Then
'If IsPlaying(PlayerI) Then
'Call SavePlayer(PlayerI)
'Call PlayerMsgCombat(PlayerI, GetPlayerName(PlayerI) & " is now saved.", Yellow)
'End If
'PlayerI = PlayerI + 1
'End If
'If PlayerI >= MAX_PLAYERS Then
'PlayerI = 1
'PlayerTimer.Enabled = False
'tmrPlayerSave.Enabled = True
End If
End Sub


Private Sub tmrGameAI_Timer()
    Call ServerLogic
End Sub

Private Sub tmrPlayerSave_Timer()
    Call PlayerSaveTimer
End Sub

Private Sub tmrSpawnMapItems_Timer()
    Call CheckSpawnMapItems
End Sub

Private Sub txtText_GotFocus()
    txtChat.SetFocus
End Sub

Private Sub txtChat_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And Trim(txtChat.text) <> "" Then
        Call GlobalMsg(txtChat.text, White)
        Call TextAdd(frmServer.txtText, "Server: " & txtChat.text, True)
        txtChat.text = ""
    End If
End Sub

Private Sub tmrShutdown_Timer()
Static Secs As Long

    If Secs <= 0 Then Secs = 30
    Call GlobalMsgCombat("Server Shutdown in " & Secs & " seconds.", BrightBlue)
    Call TextAdd(frmServer.txtText, "Automated Server Shutdown in " & Secs & " seconds.", True)
    Secs = Secs - 1
    If Secs <= 0 Then
        tmrShutdown.Enabled = False
        Call DestroyServer
    End If
End Sub

Private Sub mnuShutdown_Click()
    tmrShutdown.Enabled = True
End Sub

Private Sub mnuExit_Click()
    Call DestroyServer
End Sub

Private Sub mnuReloadRaces_Click()
    Call LoadRaces
    Call TextAdd(frmServer.txtText, "All races reloaded.", True)
End Sub

Private Sub mnuReloadClasses_Click()
    Call LoadClasses
    Call TextAdd(frmServer.txtText, "All classes reloaded.", True)
End Sub

Private Sub Socket_ConnectionRequest(index As Integer, ByVal requestID As Long)
    Call AcceptConnection(index, requestID)
End Sub

Private Sub Socket_Accept(index As Integer, SocketId As Integer)
    Call AcceptConnection(index, SocketId)
End Sub

Private Sub Socket_DataArrival(index As Integer, ByVal bytesTotal As Long)
    If IsConnected(index) Then
        Call IncomingData(index, bytesTotal)
    End If
End Sub

Private Sub Socket_Close(index As Integer)
    Call CloseSocket(index)
End Sub


