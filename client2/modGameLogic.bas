Attribute VB_Name = "modGameLogic"
Option Explicit
Private sAppName As String, sAppPath As String

Private Const ERROR_ALREADY_EXISTS = 183&
Private Declare Function CreateMutex Lib "kernel32" Alias "CreateMutexA" (lpMutexAttributes As Any, ByVal bInitialOwner As Long, ByVal lpName As String) As Long
Private Declare Function ReleaseMutex Lib "kernel32" (ByVal hMutex As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long


Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Type POINTAPI
    X As Long
    y As Long
End Type

Public Const SRCAND = &H8800C6
Public Const SRCCOPY = &HCC0020
Public Const SRCPAINT = &HEE0086

Public Const VK_UP = &H26
Public Const VK_DOWN = &H28
Public Const VK_LEFT = &H25
Public Const VK_RIGHT = &H27
Public Const VK_SHIFT = &H10
Public Const VK_RETURN = &HD
Public Const VK_CONTROL = &H11

' Menu states
Public Const MENU_STATE_NEWACCOUNT = 0
Public Const MENU_STATE_DELACCOUNT = 1
Public Const MENU_STATE_LOGIN = 2
Public Const MENU_STATE_GETCHARS = 3
Public Const MENU_STATE_NEWCHAR = 4
Public Const MENU_STATE_ADDCHAR = 5
Public Const MENU_STATE_DELCHAR = 6
Public Const MENU_STATE_USECHAR = 7
Public Const MENU_STATE_INIT = 8

' Speed moving vars
Public Const WALK_SPEED = 4
Public Const RUN_SPEED = 8

' Game direction vars
Public DirUp As Boolean
Public DirDown As Boolean
Public DirLeft As Boolean
Public DirRight As Boolean
Public ShiftDown As Boolean
Public ControlDown As Boolean

' Game text buffer
Public MyText As String

' Index of actual player
Public MyIndex As Long

' Map animation #, used to keep track of what map animation is currently on
Public MapAnim As Byte
Public MapAnimTimer As Long

' Used to freeze controls when getting a new map
Public GettingMap As Boolean

' Used to check if in editor or not and variables for use in editor
Public InEditor As Boolean
Public EditorTileX As Long
Public EditorTileY As Long
Public EditorWarpMap As Long
Public EditorWarpX As Long
Public EditorWarpY As Long

' Used for map item editor
Public ItemEditorNum As Long
Public ItemEditorValue As Long

' Used for map key editor
Public KeyEditorNum As Long
Public KeyEditorTake As Long

' Used for map key opene ditor
Public KeyOpenEditorX As Long
Public KeyOpenEditorY As Long

Public Mining1 As Long
Public Mining2 As Long
Public Mining3 As Long
Public Fishing1 As Long
Public Fishing2 As Long
Public Fishing3 As Long

Public SignTitle As String
Public SignText As String

Public Deathsay As String

Public SoundFile As String
Public Deepness As Integer

Public SoundFileD As String
Public SoundFileH As String
Public SoundFileS As String

Public Quest1 As String
Public Quest2 As String
Public Quest3 As String
'for the map edit items
Public TileData As Integer


' Map for local use
Public SaveMap As MapRec
Public SaveMapItem(1 To MAX_MAP_ITEMS) As MapItemRec
Public SaveMapNpc(1 To MAX_MAP_NPCS) As MapNpcRec

' Used for index based editors
Public InItemsEditor As Boolean
Public InNpcEditor As Boolean
Public InShopEditor As Boolean
Public InSpellEditor As Boolean
Public EditorIndex As Long

' Game fps
Public GameFPS As Long

' Used for atmosphere
Public GameWeather As Long
Public GameTime As Long

Sub Main()
Dim i As Long

Dim lngMutexHandle As Long
Dim lngLastError As Long

 ' Createmutex with above "strange" name - use names, that you think no other programs will not use :)
lngMutexHandle = CreateMutex(ByVal 0&, 1, "YourUniqueProgramNameHere")
 ' Check if the mute is already created (by other instance of this program)
If Err.LastDllError = ERROR_ALREADY_EXISTS Then
    ' Release handles
    ReleaseMutex lngMutexHandle
    CloseHandle lngMutexHandle
     ' Here you can place your code when another instance is started
    MsgBox "There are another instance of this program already running."
    End
Else
     ' No other instance is detected
     ' So do your job here
End If

 emotetemp = 0

    SHeight = Screen.Height / 15
    SWidth = Screen.Width / 15
    'frmMainMenu.Caption = Screen.Height / 11.25
    ' Check if the maps directory is there, if its not make it
    If LCase(Dir(App.Path & "\maps", vbDirectory)) <> "maps" Then
        Call MkDir(App.Path & "\maps")
    End If
    
    ' Make sure we set that we aren't in the game
    InGame = False
    GettingMap = True
    InEditor = False
    InItemsEditor = False
    InNpcEditor = False
    InShopEditor = False
    
    ' Clear out players
    For i = 1 To MAX_PLAYERS
        Call ClearPlayer(i)
    Next i
    Call ClearTempTile
    
    frmSendGetData.Visible = True
    Call SetStatus("Initializing TCP settings...")
    Call TcpInit
    'frmMainMenu.Visible = True
    frmSplash.Visible = True
    frmSendGetData.Visible = False
    

End Sub

Sub SetStatus(ByVal Caption As String)
    frmSendGetData.lblStatus.Caption = Caption
End Sub

Sub MenuState(ByVal State As Long)
    frmSendGetData.Visible = True
    Call SetStatus("Connecting to server...")
    Select Case State
        Case MENU_STATE_NEWACCOUNT
            frmNewAccount.Visible = False
            If ConnectToServer = True Then
                Call SetStatus("Connected, sending new account information...")
                Call SendNewAccount(frmNewAccount.txtName.Text, frmNewAccount.txtPassword.Text)
                Call SendNewAccount(frmNewAccount.txtName.Text, frmNewAccount.txtPassword.Text)
                       
            End If
            
        Case MENU_STATE_DELACCOUNT
            frmDeleteAccount.Visible = False
            If ConnectToServer = True Then
                Call SetStatus("Connected, sending account deletion request ...")
                Call SendDelAccount(frmDeleteAccount.txtName.Text, frmDeleteAccount.txtPassword.Text)
            End If
        
        Case MENU_STATE_LOGIN
            frmLogin.Visible = False
            If ConnectToServer = True Then
                Call SetStatus("Connected, sending login information...")
                Call SendLogin(frmLogin.txtName.Text, frmLogin.txtPassword.Text)
            End If
        
        Case MENU_STATE_NEWCHAR
            frmChars.Visible = False
            Call SetStatus("Connected, getting available races...")
            Call SendGetRaces
            Call SetStatus("Connected, getting available classes...")
            Call SendGetClasses
            frmNewChar.Visible = True
            frmSendGetData.Visible = False
            

            
        Case MENU_STATE_ADDCHAR
            frmNewChar.Visible = False
            If ConnectToServer = True Then
                Call SetStatus("Connected, sending character addition data...")
                If frmNewChar.optMale.value = True Then
                    Call SendAddChar(frmNewChar.txtName, 0, frmNewChar.cmbClass.ListIndex, frmNewChar.cmbRace.ListIndex, frmChars.lstChars.ListIndex + 1, frmNewChar.head.ListIndex)
                    
                Else
                    Call SendAddChar(frmNewChar.txtName, 1, frmNewChar.cmbClass.ListIndex, frmNewChar.cmbRace.ListIndex, frmChars.lstChars.ListIndex + 1, frmNewChar.head.ListIndex)
                    
                End If
            End If
        
        Case MENU_STATE_DELCHAR
            frmChars.Visible = False
            If ConnectToServer = True Then
                Call SetStatus("Connected, sending character deletion request...")
                Call SendDelChar(frmChars.lstChars.ListIndex + 1)
            End If
            
        Case MENU_STATE_USECHAR
            frmChars.Visible = False
            If ConnectToServer = True Then
                Call SetStatus("Connected, sending char data...")
                Call SendUseChar(frmChars.lstChars.ListIndex + 1)
            End If
    End Select

    If Not IsConnected Then
    
        frmMainMenu.Visible = True
        frmSendGetData.Visible = False
        Call MsgBox("Sorry, the server seems to be down.  Please try to reconnect in a few minutes or visit www.silent-shadows.com", vbOKOnly, GAME_NAME)
    End If
End Sub

Sub GameInit()
    frmMirage.Visible = True
    frmSendGetData.Visible = False
    Call ResizeGUI
    Call InitDirectX
End Sub

Sub GameLoop()
Dim Tick As Long
Dim TickFPS As Long
Dim FPS As Long
Dim X As Long
Dim y As Long
Dim i As Long
Dim rec_back As RECT
    
    ' Set the focus
    frmMirage.picScreen.SetFocus
    
    ' Set font
    Call SetFont("Fixedsys", 18)
    
    ' Used for calculating fps
    TickFPS = GetTickCount
    FPS = 0
    
    Do While InGame
        Tick = GetTickCount
        
        ' Check to make sure they aren't trying to auto do anything
        If GetAsyncKeyState(VK_UP) >= 0 And DirUp = True Then DirUp = False
        If GetAsyncKeyState(VK_DOWN) >= 0 And DirDown = True Then DirDown = False
        If GetAsyncKeyState(VK_LEFT) >= 0 And DirLeft = True Then DirLeft = False
        If GetAsyncKeyState(VK_RIGHT) >= 0 And DirRight = True Then DirRight = False
        If GetAsyncKeyState(VK_CONTROL) >= 0 And ControlDown = True Then ControlDown = False
        If GetAsyncKeyState(VK_SHIFT) >= 0 And ShiftDown = True Then ShiftDown = False
        
        ' Check to make sure we are still connected
        If Not IsConnected Then InGame = False
        
        ' Check if we need to restore surfaces
        If NeedToRestoreSurfaces Then
            DD.RestoreAllSurfaces
            Call InitSurfaces
        End If
                
        ' Blit out tiles layers ground/anim1/anim2
        For y = 0 To MAX_MAPY
            For X = 0 To MAX_MAPX
                Call BltTile(X, y)
            Next X
        Next y
                    
        ' Blit out the items
        For i = 1 To MAX_MAP_ITEMS
            If MapItem(i).num > 0 Then
                Call BltItem(i)
            End If
        Next i
        
        ' Blit out the npcs
        For i = 1 To MAX_MAP_NPCS
            Call BltNpc(i)
        Next i
        
        ' Blit out players
        For i = 1 To MAX_PLAYERS
            If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                Call BltPlayer(i)
            End If
        Next i
                
       '  Blit out tile layer fringe
        For y = 0 To MAX_MAPY
            For X = 0 To MAX_MAPX
                Call BltFringeTile(X, y)
            Next X
        Next y
                
        ' Lock the backbuffer so we can draw text and names
        TexthDC = DD_BackBuffer.GetDC
        Dim TmpR As RECT, TmpP As POINTAPI
        DX.GetWindowRect frmMirage.picScreen.hwnd, TmpR
        GetCursorPos TmpP
        TmpP.X = TmpP.X - TmpR.Left
        TmpP.y = TmpP.y - TmpR.top
        
        
        If frmMainMenu.AlwaysNames.value = 0 Then
        For i = 1 To MAX_PLAYERS
            If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                With Player(i)
                    If TmpP.X > (.X * PIC_X) + .XOffset Then
                        If TmpP.X < (.X * PIC_X) + .XOffset + PIC_X Then
                            If TmpP.y > (.y * PIC_Y) + .YOffset Then
                                If TmpP.y < (.y * PIC_Y) + .YOffset + PIC_Y Then
                                    Call BltPlayerName(i)
                                End If
                            End If
                        End If
                    End If
                End With
            End If
        Next i
        Else
        For i = 1 To MAX_PLAYERS
            If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                Call BltPlayerName(i)
            End If
        Next i
        End If
        
        
        'Draw NPC Names
        For i = LBound(MapNpc) To UBound(MapNpc)
            If MapNpc(i).num > 0 Then
                With MapNpc(i)
                    If TmpP.X > (.X * PIC_X) + .XOffset Then
                        If TmpP.X < (.X * PIC_X) + .XOffset + PIC_X Then
                            If TmpP.y > (.y * PIC_Y) + .YOffset Then
                                If TmpP.y < (.y * PIC_Y) + .YOffset + PIC_Y Then
                                    BltMapNPCName i
                                End If
                            End If
                        End If
                    End If
                End With
            End If
        Next
                
        ' Blit out attribs if in editor
        If InEditor Then
            For y = 0 To MAX_MAPY
                For X = 0 To MAX_MAPX
                    With Map.Tile(X, y)
                        If .Type = TILE_TYPE_BLOCKED Then Call DrawText(TexthDC, X * PIC_X + 8, y * PIC_Y + 8, "B", QBColor(BrightRed))
                        If .Type = TILE_TYPE_WARP Then Call DrawText(TexthDC, X * PIC_X + 8, y * PIC_Y + 8, "W", QBColor(BrightBlue))
                        If .Type = TILE_TYPE_ITEM Then Call DrawText(TexthDC, X * PIC_X + 8, y * PIC_Y + 8, "I", QBColor(White))
                        If .Type = TILE_TYPE_NPCAVOID Then Call DrawText(TexthDC, X * PIC_X + 8, y * PIC_Y + 8, "N", QBColor(White))
                        If .Type = TILE_TYPE_KEY Then Call DrawText(TexthDC, X * PIC_X + 8, y * PIC_Y + 8, "K", QBColor(Grey))
                        If .Type = TILE_TYPE_KEYOPEN Then Call DrawText(TexthDC, X * PIC_X + 8, y * PIC_Y + 8, "O", QBColor(White))
                        If .Type = TILE_TYPE_SIGN Then Call DrawText(TexthDC, X * PIC_X + 8, y * PIC_Y + 8, "S", QBColor(Pink))
                        If .Type = TILE_TYPE_FULLHEAL Then Call DrawText(TexthDC, X * PIC_X + 8, y * PIC_Y + 8, "+", QBColor(Blue))
                        If .Type = TILE_TYPE_DEATH Then Call DrawText(TexthDC, X * PIC_X + 8, y * PIC_Y + 8, "D", QBColor(Red))
                        If .Type = TILE_TYPE_SFX Then Call DrawText(TexthDC, X * PIC_X + 8, y * PIC_Y + 8, "S", QBColor(Yellow))
                        If .Type = TILE_TYPE_QUEST Then Call DrawText(TexthDC, X * PIC_X + 8, y * PIC_Y + 8, "Q", QBColor(Magenta))
                        If .Type = TILE_TYPE_ARENA Then Call DrawText(TexthDC, X * PIC_X + 8, y * PIC_Y + 8, "A", QBColor(Magenta))
                        If .Type = TILE_TYPE_FISHING Then Call DrawText(TexthDC, X * PIC_X + 8, y * PIC_Y + 8, "F", QBColor(White))
                        If .Type = TILE_TYPE_MINING Then Call DrawText(TexthDC, X * PIC_X + 8, y * PIC_Y + 8, "M", QBColor(White))
                    End With
                Next X
            Next y
        End If
        
        ' Blit the text they are putting in
        'Call DrawText(TexthDC, 0, (MAX_MAPY + 1) * PIC_Y - 20, MyText, RGB(255, 255, 255))
        frmMirage.MyTextBox.Text = MyText
        If Len(MyText) > 4 Then
        frmMirage.MyTextBox.SelStart = Len(frmMirage.MyTextBox.Text) - 1
        End If
        ' Draw map name
        If Map.Moral = MAP_MORAL_NONE Then
            Call DrawText(TexthDC, Int((MAX_MAPX + 1) * PIC_X / 2) - (Int(Len(Trim(Map.name)) / 2) * 8), 1, Trim(Map.name), QBColor(BrightRed))
        Else
            Call DrawText(TexthDC, Int((MAX_MAPX + 1) * PIC_X / 2) - (Int(Len(Trim(Map.name)) / 2) * 8), 1, Trim(Map.name), QBColor(White))
        End If
        
        ' Check if we are getting a map, and if we are tell them so
        'If GettingMap = True Then
        '    Call DrawText(TexthDC, 50, 50, "Receiving Map...", QBColor(BrightCyan))
        'End If
                        
        ' Release DC
        Call DD_BackBuffer.ReleaseDC(TexthDC)
        
        ' Get the rect for the back buffer to blit from
        rec.top = 0
        rec.Bottom = (MAX_MAPY + 1) * PIC_Y
        rec.Left = 0
        rec.Right = (MAX_MAPX + 1) * PIC_X
        
        ' Get the rect to blit to
        Call DX.GetWindowRect(frmMirage.picScreen.hwnd, rec_pos)
        rec_pos.Bottom = rec_pos.top + ((MAX_MAPY + 1) * PIC_Y)
        rec_pos.Right = rec_pos.Left + ((MAX_MAPX + 1) * PIC_X)
        
        ' Blit the backbuffer
        Call DD_PrimarySurf.Blt(rec_pos, DD_BackBuffer, rec, DDBLT_WAIT)
        
        ' Check if player is trying to move
        Call CheckMovement
        
        ' Check to see if player is trying to attack
        Call CheckAttack
        
        ' Process player movements (actually move them)
        For i = 1 To MAX_PLAYERS
            If IsPlaying(i) Then
                Call ProcessMovement(i)
            End If
        Next i
        
        ' Process npc movements (actually move them)
        For i = 1 To MAX_MAP_NPCS
            If Map.Npc(i) > 0 Then
                Call ProcessNpcMovement(i)
            End If
        Next i
            
        ' Change map animation every 250 milliseconds
        If GetTickCount > MapAnimTimer + 250 Then
            If MapAnim = 0 Then
                MapAnim = 1
            Else
                MapAnim = 0
            End If
            MapAnimTimer = GetTickCount
        End If
                
        ' Lock fps
        Do While GetTickCount < Tick + 50
            DoEvents
        Loop
        
        ' Calculate fps
        If GetTickCount > TickFPS + 1000 Then
            GameFPS = FPS
            TickFPS = GetTickCount
            FPS = 0
        Else
            FPS = FPS + 1
        End If
        
        DoEvents
        
    Loop
    
    frmMirage.Visible = False
    frmSendGetData.Visible = True
    Call SetStatus("Destroying game data...")
    
    ' Shutdown the game
    Call GameDestroy
    
    ' Report disconnection if server disconnects
    If IsConnected = False Then
        Call MsgBox("Thank you for playing " & GAME_NAME & "!", vbOKOnly, GAME_NAME)
    End If
End Sub

Sub GameDestroy()
    Call DestroyDirectX
    FSOUND_Close
    'Call ChangeRes(SWidth, SHeight)
    End
End Sub

Sub BltTile(ByVal X As Long, ByVal y As Long)
Dim Ground As Long
Dim Anim1 As Long
Dim Anim2 As Long
Dim Mask2 As Long
Dim M2Anim As Long

    Ground = Map.Tile(X, y).Ground
    Anim1 = Map.Tile(X, y).Mask
    Anim2 = Map.Tile(X, y).Anim
    Mask2 = Map.Tile(X, y).Mask2
    M2Anim = Map.Tile(X, y).M2Anim
    
    ' Only used if ever want to switch to blt rather then bltfast
    With rec_pos
        .top = y * PIC_Y
        .Bottom = .top + PIC_Y
        .Left = X * PIC_X
        .Right = .Left + PIC_X
    End With
    
    rec.top = Int(Ground / 7) * PIC_Y
    rec.Bottom = rec.top + PIC_Y
    rec.Left = (Ground - Int(Ground / 7) * 7) * PIC_X
    rec.Right = rec.Left + PIC_X
     'Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf, rec, DDBLT_WAIT)
    Call DD_BackBuffer.BltFast(X * PIC_X, y * PIC_Y, DD_TileSurf, rec, DDBLTFAST_WAIT)
    
    If (MapAnim = 0) Or (Anim2 <= 0) Then
        ' Is there an animation tile to plot?
        If Anim1 > 0 And TempTile(X, y).DoorOpen = NO Then
            rec.top = Int(Anim1 / 7) * PIC_Y
            rec.Bottom = rec.top + PIC_Y
            rec.Left = (Anim1 - Int(Anim1 / 7) * 7) * PIC_X
            rec.Right = rec.Left + PIC_X
            'Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
            Call DD_BackBuffer.BltFast(X * PIC_X, y * PIC_Y, DD_TileSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    Else
        ' Is there a second animation tile to plot?
        If Anim2 > 0 Then
            rec.top = Int(Anim2 / 7) * PIC_Y
            rec.Bottom = rec.top + PIC_Y
            rec.Left = (Anim2 - Int(Anim2 / 7) * 7) * PIC_X
            rec.Right = rec.Left + PIC_X
            'Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
            Call DD_BackBuffer.BltFast(X * PIC_X, y * PIC_Y, DD_TileSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    End If
    
    If (MapAnim = 0) Or (M2Anim <= 0) Then
        ' Is there an animation tile to plot?
        If Mask2 > 0 Then
            rec.top = Int(Mask2 / 7) * PIC_Y
            rec.Bottom = rec.top + PIC_Y
            rec.Left = (Mask2 - Int(Mask2 / 7) * 7) * PIC_X
            rec.Right = rec.Left + PIC_X
            'Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
            Call DD_BackBuffer.BltFast(X * PIC_X, y * PIC_Y, DD_TileSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    Else
        ' Is there a second animation tile to plot?
        If M2Anim > 0 Then
            rec.top = Int(M2Anim / 7) * PIC_Y
            rec.Bottom = rec.top + PIC_Y
            rec.Left = (M2Anim - Int(M2Anim / 7) * 7) * PIC_X
            rec.Right = rec.Left + PIC_X
            'Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
            Call DD_BackBuffer.BltFast(X * PIC_X, y * PIC_Y, DD_TileSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    End If
End Sub

Sub BltItem(ByVal ItemNum As Long)
    ' Only used if ever want to switch to blt rather then bltfast
    With rec_pos
        .top = MapItem(ItemNum).y * PIC_Y
        .Bottom = .top + PIC_Y
        .Left = MapItem(ItemNum).X * PIC_X
        .Right = .Left + PIC_X
    End With

    rec.top = Item(MapItem(ItemNum).num).Pic * PIC_Y
    rec.Bottom = rec.top + PIC_Y
    rec.Left = 0
    rec.Right = rec.Left + PIC_X
    
    'Call DD_BackBuffer.Blt(rec_pos, DD_ItemSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
    Call DD_BackBuffer.BltFast(MapItem(ItemNum).X * PIC_X, MapItem(ItemNum).y * PIC_Y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub

Sub BltFringeTile(ByVal X As Long, ByVal y As Long)
Dim Fringe As Long
Dim FAnim As Long
Dim Fringe2 As Long
Dim F2Anim As Long
Dim Fringe3 As Long

    ' Only used if ever want to switch to blt rather then bltfast
    With rec_pos
        .top = y * PIC_Y
        .Bottom = .top + PIC_Y
        .Left = X * PIC_X
        .Right = .Left + PIC_X
    End With
    
    Fringe = Map.Tile(X, y).Fringe
    FAnim = Map.Tile(X, y).FAnim
    Fringe2 = Map.Tile(X, y).Fringe2
    F2Anim = Map.Tile(X, y).F2Anim
    If GameTime = TIME_NIGHT Then
    If Map.BootX > 0 Then
    Fringe3 = 0
    Else
    Fringe3 = 658 '<--- make this a shadow tile to fringe
    End If
    Else
    Fringe3 = 0 '<--- any empty black tile
    End If
        
    If (MapAnim = 0) Or (FAnim <= 0) Then
        ' Is there an animation tile to plot?
        
        If Fringe > 0 Then
        rec.top = Int(Fringe / 7) * PIC_Y
        rec.Bottom = rec.top + PIC_Y
        rec.Left = (Fringe - Int(Fringe / 7) * 7) * PIC_X
        rec.Right = rec.Left + PIC_X
        'Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
        Call DD_BackBuffer.BltFast(X * PIC_X, y * PIC_Y, DD_TileSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    
    Else
    
        If FAnim > 0 Then
        rec.top = Int(FAnim / 7) * PIC_Y
        rec.Bottom = rec.top + PIC_Y
        rec.Left = (FAnim - Int(FAnim / 7) * 7) * PIC_X
        rec.Right = rec.Left + PIC_X
        'Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
        Call DD_BackBuffer.BltFast(X * PIC_X, y * PIC_Y, DD_TileSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
        
    End If

    If (MapAnim = 0) Or (F2Anim <= 0) Then
        ' Is there an animation tile to plot?
        
        If Fringe2 > 0 Then
        rec.top = Int(Fringe2 / 7) * PIC_Y
        rec.Bottom = rec.top + PIC_Y
        rec.Left = (Fringe2 - Int(Fringe2 / 7) * 7) * PIC_X
        rec.Right = rec.Left + PIC_X
        'Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
        Call DD_BackBuffer.BltFast(X * PIC_X, y * PIC_Y, DD_TileSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    
    Else
    
        If F2Anim > 0 Then
        rec.top = Int(F2Anim / 7) * PIC_Y
        rec.Bottom = rec.top + PIC_Y
        rec.Left = (F2Anim - Int(F2Anim / 7) * 7) * PIC_X
        rec.Right = rec.Left + PIC_X
        'Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
        Call DD_BackBuffer.BltFast(X * PIC_X, y * PIC_Y, DD_TileSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
        
    End If
        If Fringe3 > 0 Then
        rec.top = Int(Fringe3 / 7) * PIC_Y
        rec.Bottom = rec.top + PIC_Y
        rec.Left = (Fringe3 - Int(Fringe3 / 7) * 7) * PIC_X
        rec.Right = rec.Left + PIC_X
        'Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
        Call DD_BackBuffer.BltFast(X * PIC_X, y * PIC_Y, DD_TileSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If

End Sub

Sub BltPlayer(ByVal index As Long)
Dim Anim As Byte
Dim X As Long, y As Long

    ' Only used if ever want to switch to blt rather then bltfast
    With rec_pos
        .top = GetPlayerY(index) * PIC_Y + Player(index).YOffset
        .Bottom = .top + PIC_Y
        .Left = GetPlayerX(index) * PIC_X + Player(index).XOffset
        .Right = .Left + PIC_X
    End With
    
    ' Check for animation
    Anim = 0
    If Player(index).Attacking = 0 Then
        Select Case GetPlayerDir(index)
            Case DIR_UP
                If (Player(index).YOffset < PIC_Y / 2) Then Anim = 1
            Case DIR_DOWN
                If (Player(index).YOffset < PIC_Y / 2 * -1) Then Anim = 1
            Case DIR_LEFT
                If (Player(index).XOffset < PIC_Y / 2) Then Anim = 1
            Case DIR_RIGHT
                If (Player(index).XOffset < PIC_Y / 2 * -1) Then Anim = 1
        End Select
    Else
        If Player(index).AttackTimer + 500 > GetTickCount Then
            Anim = 2
        End If
    End If
    
    ' Check to see if we want to stop making him attack
    If Player(index).AttackTimer + 1000 < GetTickCount Then
        Player(index).Attacking = 0
        Player(index).AttackTimer = 0
    End If
    
    
    Select Case GetPlayerDir(index)
            Case DIR_UP
                Call DRAWPLAYERUP(index, Anim, X, y)
            Case DIR_DOWN
                Call DRAWPLAYERdown(index, Anim, X, y)
            Case DIR_LEFT
                Call DRAWPLAYERleft(index, Anim, X, y)
            Case DIR_RIGHT
                Call DRAWPLAYERright(index, Anim, X, y)
        End Select
                Call DRAWspellgfxPLAYER(index, Anim, X, y)
                Call DRAWemotegfxPLAYER(index, Anim, X, y)

End Sub
Sub BltPlayerName(ByVal index As Long)
Dim TextX As Long
Dim TextY As Long
Dim Color As Long
Dim name As String
Dim guild As String
    Dim Anim As Byte
    Anim = 0
    ' Check access level
    If GetPlayerPK(index) = NO Then
        Select Case GetPlayerAccess(index)
            Case 0
                Color = QBColor(White)
            Case 1
                Color = QBColor(BrightCyan)
            Case 2
                Color = QBColor(BrightCyan)
            Case 3
                Color = QBColor(BrightCyan)
            Case 4
                Color = QBColor(BrightCyan)
            Case 5
                Color = QBColor(BrightCyan)
            Case 6
                Color = QBColor(BrightCyan)
            Case 7
                Color = QBColor(BrightCyan)
            Case 8
                Color = QBColor(BrightCyan)
            Case 9
                Color = QBColor(BrightCyan)
        End Select
     Else
        Color = QBColor(BrightRed)
    End If
    If GetPlayerAnonymous(index) > 0 Then
    If GetPlayerAccess(index) > 0 Then
     Color = QBColor(Cyan)
    Else
     Color = QBColor(Yellow)
    End If
    End If
    
    
    If GetPlayerFaction(index) > 1 Then
    If GetPlayerFaction(index) = 2 Then
    name = "(C)" & GetPlayerName(index)
    ElseIf GetPlayerFaction(index) = 3 Then
    name = "(B)" & GetPlayerName(index)
    Else
    name = GetPlayerName(index)
    End If
    Else
    name = GetPlayerName(index)
    End If
    If GetPlayerGuild(index) <> "" Then
    guild = "<" & GetPlayerGuild(index) & ">"
    Else
    guild = ""
    End If
    
    ' Draw name
    TextX = GetPlayerX(index) * PIC_X + Player(index).XOffset + Int(PIC_X / 2) - ((Len(GetPlayerName(index)) / 2) * 8)
    TextY = GetPlayerY(index) * PIC_Y + Player(index).YOffset - Int(PIC_Y / 2) - 28
    If GetPlayerAccess(index) < 10 Then Call DrawText(TexthDC, TextX, TextY, name, Color)
    TextX = GetPlayerX(index) * PIC_X + Player(index).XOffset + Int(PIC_X / 2) - ((Len(GetPlayerGuild(index)) / 2) * 8)
    TextY = GetPlayerY(index) * PIC_Y + Player(index).YOffset - Int(PIC_Y / 2) - 12
    If GetPlayerAccess(index) < 10 Then Call DrawText(TexthDC, TextX, TextY, guild, Color)
    
End Sub

Sub BltNpc(ByVal MapNpcNum As Long)
Dim Anim As Byte
Dim X As Long, y As Long

    ' Make sure that theres an npc there, and if not exit the sub
    If MapNpc(MapNpcNum).num <= 0 Then
        Exit Sub
    End If
    
    ' Only used if ever want to switch to blt rather then bltfast
    With rec_pos
        .top = MapNpc(MapNpcNum).y * PIC_Y + MapNpc(MapNpcNum).YOffset
        .Bottom = .top + PIC_Y
        .Left = MapNpc(MapNpcNum).X * PIC_X + MapNpc(MapNpcNum).XOffset
        .Right = .Left + PIC_X
    End With
    
    ' Check for animation
    Anim = 0
    If MapNpc(MapNpcNum).Attacking = 0 Then
        Select Case MapNpc(MapNpcNum).Dir
            Case DIR_UP
                If (MapNpc(MapNpcNum).YOffset < PIC_Y / 2) Then Anim = 1
            Case DIR_DOWN
                If (MapNpc(MapNpcNum).YOffset < PIC_Y / 2 * -1) Then Anim = 1
            Case DIR_LEFT
                If (MapNpc(MapNpcNum).XOffset < PIC_Y / 2) Then Anim = 1
            Case DIR_RIGHT
                If (MapNpc(MapNpcNum).XOffset < PIC_Y / 2 * -1) Then Anim = 1
        End Select
    Else
        If MapNpc(MapNpcNum).AttackTimer + 500 > GetTickCount Then
            Anim = 2
        End If
    End If
    
    ' Check to see if we want to stop making him attack
    If MapNpc(MapNpcNum).AttackTimer + 1000 < GetTickCount Then
        MapNpc(MapNpcNum).Attacking = 0
        MapNpc(MapNpcNum).AttackTimer = 0
    End If
    
    Select Case MapNpc(MapNpcNum).Dir
            Case DIR_UP
             Call DRAWNPCup(MapNpcNum, Anim, X, y)
            Case DIR_DOWN
             Call DRAWNPCdown(MapNpcNum, Anim, X, y)
            Case DIR_LEFT
             Call DRAWNPCleft(MapNpcNum, Anim, X, y)
            Case DIR_RIGHT
             Call DRAWNPCright(MapNpcNum, Anim, X, y)
            End Select
            Call DRAWspellgfxNPC(MapNpcNum, Anim, X, y)
            
    End Sub

Sub ProcessMovement(ByVal index As Long)
    ' Check if player is walking, and if so process moving them over
    If Player(index).Moving = MOVING_WALKING Then
        Select Case GetPlayerDir(index)
            Case DIR_UP
                Player(index).YOffset = Player(index).YOffset - WALK_SPEED
            Case DIR_DOWN
                Player(index).YOffset = Player(index).YOffset + WALK_SPEED
            Case DIR_LEFT
                Player(index).XOffset = Player(index).XOffset - WALK_SPEED
            Case DIR_RIGHT
                Player(index).XOffset = Player(index).XOffset + WALK_SPEED
        End Select
        
        ' Check if completed walking over to the next tile
        If (Player(index).XOffset = 0) And (Player(index).YOffset = 0) Then
            Player(index).Moving = 0
        End If
    End If

    ' Check if player is running, and if so process moving them over
    If Player(index).Moving = MOVING_RUNNING Then
        Select Case GetPlayerDir(index)
            Case DIR_UP
                Player(index).YOffset = Player(index).YOffset - RUN_SPEED
            Case DIR_DOWN
                Player(index).YOffset = Player(index).YOffset + RUN_SPEED
            Case DIR_LEFT
                Player(index).XOffset = Player(index).XOffset - RUN_SPEED
            Case DIR_RIGHT
                Player(index).XOffset = Player(index).XOffset + RUN_SPEED
        End Select
        
        ' Check if completed walking over to the next tile
        If (Player(index).XOffset = 0) And (Player(index).YOffset = 0) Then
            Player(index).Moving = 0
        End If
    End If
End Sub

Sub ProcessNpcMovement(ByVal MapNpcNum As Long)
    ' Check if player is walking, and if so process moving them over
    If MapNpc(MapNpcNum).Moving = MOVING_WALKING Then
        Select Case MapNpc(MapNpcNum).Dir
            Case DIR_UP
                MapNpc(MapNpcNum).YOffset = MapNpc(MapNpcNum).YOffset - WALK_SPEED
            Case DIR_DOWN
                MapNpc(MapNpcNum).YOffset = MapNpc(MapNpcNum).YOffset + WALK_SPEED
            Case DIR_LEFT
                MapNpc(MapNpcNum).XOffset = MapNpc(MapNpcNum).XOffset - WALK_SPEED
            Case DIR_RIGHT
                MapNpc(MapNpcNum).XOffset = MapNpc(MapNpcNum).XOffset + WALK_SPEED
        End Select
        
        ' Check if completed walking over to the next tile
        If (MapNpc(MapNpcNum).XOffset = 0) And (MapNpc(MapNpcNum).YOffset = 0) Then
            MapNpc(MapNpcNum).Moving = 0
        End If
    End If
End Sub

Sub HandleKeypresses(ByVal KeyAscii As Integer)
Dim ChatText As String
Dim name As String
Dim i As Long
Dim n As Long

    ' Handle when the player presses the return key
    If (KeyAscii = vbKeyReturn) Then
        ' Broadcast message
        If Mid(MyText, 1, 1) = "'" Then
            ChatText = Mid(MyText, 2, Len(MyText) - 1)
            If Len(Trim(ChatText)) > 0 Then
                Call BroadcastMsg(ChatText)
            End If
            MyText = ""
            Exit Sub
        End If
        
        ' Emote message
        If Mid(MyText, 1, 1) = "-" Then
            ChatText = Mid(MyText, 2, Len(MyText) - 1)
            If Len(Trim(ChatText)) > 0 Then
                Call EmoteMsg(ChatText)
            End If
            MyText = ""
            Exit Sub
        End If
        
        If Mid(MyText, 1, 6) = "/emote" Then
            ChatText = Mid(MyText, 7, Len(MyText) - 6)
            If Len(Trim(ChatText)) > 0 Then
                Call SendData("emote" & SEP_CHAR & ChatText & SEP_CHAR & END_CHAR)
            End If
            MyText = ""
            Exit Sub
        End If
        
        If LCase(Mid(MyText, 1, 5)) = "/bank" Then
            ' Make sure they are actually sending something
                Call SendData("bank" & SEP_CHAR & END_CHAR)
            MyText = ""
            Exit Sub
        End If
        
        If LCase(Mid(MyText, 1, 10)) = "/anonymous" Then
            ' Make sure they are actually sending something
               frmAnonymous.Show vbModal
            MyText = ""
            Exit Sub
        End If
        
        ' Player message
        If Mid(MyText, 1, 1) = "!" Then
            ChatText = Mid(MyText, 2, Len(MyText) - 1)
            name = ""
                    
            ' Get the desired player from the user text
            For i = 1 To Len(ChatText)
                If Mid(ChatText, i, 1) <> " " Then
                    name = name & Mid(ChatText, i, 1)
                Else
                    Exit For
                End If
            Next i
                    
            ' Make sure they are actually sending something
            If Len(ChatText) - i > 0 Then
                ChatText = Mid(ChatText, i + 1, Len(ChatText) - i)
                    
                ' Send the message to the player
                Call PlayerMsg(ChatText, name)
            Else
                Call AddText("Usage: !playername msghere", AlertColor)
            End If
            MyText = ""
            Exit Sub
        End If
            
        ' // Commands //
        ' Help
        If LCase(Mid(MyText, 1, 5)) = "/help" Then
            Call AddText("Social Commands:", HelpColor)
            Call AddText("'msghere = Broadcast Message", HelpColor)
            Call AddText("-msghere = Emote Message", HelpColor)
            Call AddText("!namehere msghere = Player Message", HelpColor)
            Call AddText("Available Commands: /help, /info, /who, /fps, /inv, /stats, /train, /trade, /party, /join, /leave", HelpColor)
            Call AddText("If you get stuck like frozen or you can't pick up an item type /stuck", HelpColor)
            MyText = ""
            Exit Sub
        End If
        
        If LCase(Mid(MyText, 1, 6)) = "/agree" Then
        Call SendData("agreetrade" & SEP_CHAR & END_CHAR)
        MyText = ""
        Exit Sub
        End If
        
        
        If LCase(Mid(MyText, 1, 6)) = "/stuck" Then
        Call SendData("stuck" & SEP_CHAR & END_CHAR)
        MyText = ""
        Exit Sub
        End If
        
        ' Verification User
        If LCase(Mid(MyText, 1, 5)) = "/info" Then
            ChatText = Mid(MyText, 6, Len(MyText) - 5)
            Call SendData("playerinforequest" & SEP_CHAR & ChatText & SEP_CHAR & END_CHAR)
            MyText = ""
            Exit Sub
        End If
        
        ' Whos Online
        If LCase(Mid(MyText, 1, 4)) = "/who" Then
            Call SendWhosOnline
            MyText = ""
            Exit Sub
        End If
                        
        ' Checking fps
        If LCase(Mid(MyText, 1, 4)) = "/fps" Then
            Call AddText("FPS: " & GameFPS, Pink)
            MyText = ""
            Exit Sub
        End If
                
        ' Show inventory
        If LCase(Mid(MyText, 1, 4)) = "/inv" Then
            Call UpdateInventory
            frmInventory.picInv.Visible = True
            MyText = ""
            Exit Sub
        End If
        
        ' Request stats
        If LCase(Mid(MyText, 1, 6)) = "/stats" Then
            Call SendData("getstats" & SEP_CHAR & END_CHAR)
            MyText = ""
            Exit Sub
        End If
    
        ' Show training
        If LCase(Mid(MyText, 1, 6)) = "/train" Then
            frmTraining.Show vbModal
            MyText = ""
            Exit Sub
        End If

        ' Request stats
        If LCase(Mid(MyText, 1, 6)) = "/trade" Then
            Call SendData("trade" & SEP_CHAR & END_CHAR)
            MyText = ""
            Exit Sub
        End If
        
        If LCase(Mid(MyText, 1, 11)) = "/startguild" Then
            ' Make sure they are actually sending something
            If Len(MyText) > 12 Then
                ChatText = Mid(MyText, 13, Len(MyText) - 12)
                Call SendData("startguild" & SEP_CHAR & ChatText & SEP_CHAR & END_CHAR)
            Else
                Call AddText("Usage: /startguild guildname here", AlertColor)
            End If
            MyText = ""
            Exit Sub
        End If
            If LCase(Mid(MyText, 1, 10)) = "/guildrank" Then
                ' Get access #
                i = Val(Mid(MyText, 12, 1))
                
                MyText = Mid(MyText, 14, Len(MyText) - 13)
                
                Call SendData("setguildrank" & SEP_CHAR & MyText & SEP_CHAR & i & SEP_CHAR & END_CHAR)
                MyText = ""
                Exit Sub
            End If
            
            If LCase(Mid(MyText, 1, 10)) = "/guildwho" Then
                Call SendData("guildwho" & SEP_CHAR & END_CHAR)
                MyText = ""
                Exit Sub
            End If

If LCase(Mid(MyText, 1, 11)) = "/leaveguild" Then
            ' Make sure they are actually sending something
                Call SendData("leaveguild" & SEP_CHAR & END_CHAR)
            MyText = ""
            Exit Sub
        End If

If LCase(Mid(MyText, 1, 12)) = "/guildinvite" Then
            ' Make sure they are actually sending something
            If Len(MyText) > 13 Then
                ChatText = Mid(MyText, 14, Len(MyText) - 13)
                Call SendData("guildinvite" & SEP_CHAR & ChatText & SEP_CHAR & END_CHAR)
            Else
                Call AddText("Usage: /guildinvite Player Name here", AlertColor)
            End If
            MyText = ""
            Exit Sub
        End If

If LCase(Mid(MyText, 1, 11)) = "/guildjoin" Then
            ' Make sure they are actually sending something
                Call SendData("guildjoin" & SEP_CHAR & END_CHAR)
            MyText = ""
            Exit Sub
        End If
        
        If LCase(Mid(MyText, 1, 12)) = "/joinfaction" Then
            ' Make sure they are actually sending something
                Call SendData("joinfaction" & SEP_CHAR & END_CHAR)
            MyText = ""
            Exit Sub
        End If
        
        If LCase(Mid(MyText, 1, 6)) = "/gchat" Then
            ' Make sure they are actually sending something
            If Len(MyText) > 7 Then
                ChatText = Mid(MyText, 8, Len(MyText) - 7)
                Call SendData("gchat" & SEP_CHAR & ChatText & SEP_CHAR & END_CHAR)
            Else
                Call AddText("Usage: /gchat message here", AlertColor)
            End If
            MyText = ""
            Exit Sub
        End If
        
        
        
        If LCase(Mid(MyText, 1, 9)) = "/tcswitch" Then
            ' Make sure they are actually sending something
            If Len(MyText) > 10 Then
                ChatText = Mid(MyText, 11, Len(MyText) - 10)
                Call SendData("tcswitch" & SEP_CHAR & ChatText & SEP_CHAR & END_CHAR)
            Else
                Call AddText("Usage: /tcswitch on/off", AlertColor)
            End If
            MyText = ""
            Exit Sub
        End If
        
        If LCase(Mid(MyText, 1, 4)) = "/ooc" Then
            ' Make sure they are actually sending something
            If Len(MyText) > 5 Then
                ChatText = Mid(MyText, 6, Len(MyText) - 5)
                Call SendData("ooc" & SEP_CHAR & ChatText & SEP_CHAR & END_CHAR)
            Else
                Call AddText("Usage: /ooc on/off", AlertColor)
            End If
            MyText = ""
            Exit Sub
        End If
        
        If LCase(Mid(MyText, 1, 3)) = "/tc" Then
            ' Make sure they are actually sending something
            If Len(MyText) > 4 Then
                ChatText = Mid(MyText, 5, Len(MyText) - 4)
                Call SendData("tcmsg" & SEP_CHAR & ChatText & SEP_CHAR & END_CHAR)
            Else
                Call AddText("Usage: /tc message here", AlertColor)
            End If
            MyText = ""
            Exit Sub
        End If
        
        
        If LCase(Mid(MyText, 1, 5)) = "/time" Then
            ' Make sure they are actually sending something
                Call SendData("time" & SEP_CHAR & END_CHAR)
            MyText = ""
            Exit Sub
        End If
        
        If LCase(Mid(MyText, 1, 5)) = "/bind" Then
            ' Make sure they are actually sending something
                Call SendData("bind" & SEP_CHAR & END_CHAR)
            MyText = ""
            Exit Sub
        End If
        
        If LCase(Mid(MyText, 1, 6)) = "/pinfo" Then
            ' Make sure they are actually sending something
                Call SendData("pinfo" & SEP_CHAR & END_CHAR)
            MyText = ""
            Exit Sub
        End If
        
        If LCase(Mid(MyText, 1, 5)) = "/pwho" Then
            ' Make sure they are actually sending something
                Call SendData("pwho" & SEP_CHAR & END_CHAR)
            MyText = ""
            Exit Sub
        End If
        
        If LCase(Mid(MyText, 1, 3)) = "/pc" Then
            ' Make sure they are actually sending something
            If Len(MyText) > 4 Then
                ChatText = Mid(MyText, 5, Len(MyText) - 4)
                Call SendData("partychat" & SEP_CHAR & ChatText & SEP_CHAR & END_CHAR)
            Else
                Call AddText("Usage: /pc message here", AlertColor)
            End If
            MyText = ""
            Exit Sub
        End If
        
        
        
        ' Party request
        If LCase(Mid(MyText, 1, 6)) = "/party" Then
            ' Make sure they are actually sending something
            If Len(MyText) > 7 Then
                ChatText = Mid(MyText, 8, Len(MyText) - 7)
                Call SendPartyRequest(ChatText)
            Else
                Call AddText("Usage: /party playernamehere", AlertColor)
            End If
            MyText = ""
            Exit Sub
        End If
        
        ' Join party
        If LCase(Mid(MyText, 1, 5)) = "/join" Then
            Call SendJoinParty
            MyText = ""
            Exit Sub
        End If
        
        ' Leave party
        If LCase(Mid(MyText, 1, 6)) = "/leave" Then
            Call SendLeaveParty
            MyText = ""
            Exit Sub
        End If
        
        ' // Moniter Admin Commands //
                    
                  ' Admin Help
            If LCase(Mid(MyText, 1, 6)) = "/admin" Then
            If InStr(1, Player(MyIndex).AdminCmds, LCase(Mid(MyText, 1, 6))) Then
                Call AddText("Social Commands:", HelpColor)
                Call AddText("""msghere = Global Admin Message", HelpColor)
                Call AddText("`msghere = Private Admin Message", HelpColor)
                Call AddText("Available Commands: /admin, /loc, /mapeditor, /warpmeto, /warptome, /warpto, /setsprite, /mapreport, /kick, /ban, /edititem, /respawn, /editnpc, /motd, /editshop, /ban, /editspell", HelpColor)
                MyText = ""
                Exit Sub
            End If
            End If
            'mute
            If LCase(Mid(MyText, 1, 5)) = "/mute" Then
            If InStr(1, Player(MyIndex).AdminCmds, LCase(Mid(MyText, 1, 5))) Then
                If Len(MyText) > 6 Then
                    MyText = Mid(MyText, 7, Len(MyText) - 6)
                    Call SendMute(MyText)
                End If
                MyText = ""
                Exit Sub
                End If
            End If
            'warpme to a player
            If LCase(Mid(MyText, 1, 9)) = "/warpmeto" Then
            If InStr(1, Player(MyIndex).AdminCmds, LCase(Mid(MyText, 1, 9))) Then
                If Len(MyText) > 10 Then
                    MyText = Mid(MyText, 10, Len(MyText) - 9)
                    Call WarpMeTo(MyText)
                End If
                MyText = ""
                Exit Sub
                End If
            End If
            
            'unmute
            If LCase(Mid(MyText, 1, 7)) = "/unmute" Then
            If InStr(1, Player(MyIndex).AdminCmds, LCase(Mid(MyText, 1, 7))) Then
                If Len(MyText) > 8 Then
                    MyText = Mid(MyText, 9, Len(MyText) - 8)
                    Call SendUnMute(MyText)
                End If
                MyText = ""
                Exit Sub
            End If
            End If
            
            
            ' Kicking a player
            If LCase(Mid(MyText, 1, 5)) = "/kick" Then
             If InStr(1, Player(MyIndex).AdminCmds, LCase(Mid(MyText, 1, 5))) Then
                If Len(MyText) > 6 Then
                    MyText = Mid(MyText, 7, Len(MyText) - 6)
                    Call SendKick(MyText)
                End If
                MyText = ""
                Exit Sub
            End If
            End If
            
            ' making a player feel stupid
            If LCase(Mid(MyText, 1, 5)) = "/dumb" Then
            If InStr(1, Player(MyIndex).AdminCmds, LCase(Mid(MyText, 1, 5))) Then
                If Len(MyText) > 6 Then
                    MyText = Mid(MyText, 7, Len(MyText) - 6)
                    Call SendDumb(MyText)
                End If
                MyText = ""
                Exit Sub
                End If
            End If
            
            ' making a player feel stupid
            If LCase(Mid(MyText, 1, 5)) = "/pban" Then
            If InStr(1, Player(MyIndex).AdminCmds, LCase(Mid(MyText, 1, 5))) Then
                If Len(MyText) > 6 Then
                    MyText = Mid(MyText, 7, Len(MyText) - 6)
                    Call SendPBan(MyText)
                End If
                MyText = ""
                Exit Sub
                End If
            End If
        
            ' Global Message
            
            If Mid(MyText, 1, 1) = """" Then
            If InStr(1, Player(MyIndex).AdminCmds, LCase(Mid(MyText, 1, 1))) Then
                ChatText = Mid(MyText, 2, Len(MyText) - 1)
                If Len(Trim(ChatText)) > 0 Then
                    Call GlobalMsg(ChatText)
                End If
                MyText = ""
                Exit Sub
                End If
            End If
        
            ' Admin Message
            If Mid(MyText, 1, 1) = "`" Then
            If InStr(1, Player(MyIndex).AdminCmds, LCase(Mid(MyText, 1, 1))) Then
                ChatText = Mid(MyText, 2, Len(MyText) - 1)
                If Len(Trim(ChatText)) > 0 Then
                    Call AdminMsg(ChatText)
                End If
                MyText = ""
                Exit Sub
                End If
            End If
                
        ' // Mapper Admin Commands //
          ' Location
            If LCase(Mid(MyText, 1, 4)) = "/loc" Then
            If InStr(1, Player(MyIndex).AdminCmds, LCase(Mid(MyText, 1, 4))) Then
                Call SendRequestLocation
                MyText = ""
                Exit Sub
                End If
            End If
            
            ' Map Editor
            If LCase(Mid(MyText, 1, 10)) = "/mapeditor" Then
            If InStr(1, Player(MyIndex).AdminCmds, LCase(Mid(MyText, 1, 10))) Then
                Call SendRequestEditMap
                MyText = ""
                Exit Sub
                End If
            End If
            
             'Warping to a player
            If LCase(Mid(MyText, 1, 9)) = "/warpmeto" Then
            If InStr(1, Player(MyIndex).AdminCmds, LCase(Mid(MyText, 1, 9))) Then
                If Len(MyText) > 10 Then
                   MyText = Mid(MyText, 10, Len(MyText) - 9)
                   Call WarpMeTo(MyText)
               End If
               MyText = ""
              Exit Sub
              End If
           End If
                                              
            ' Warping a player to you
            If LCase(Mid(MyText, 1, 9)) = "/warptome" Then
            If InStr(1, Player(MyIndex).AdminCmds, LCase(Mid(MyText, 1, 9))) Then
                If Len(MyText) > 10 Then
                    MyText = Mid(MyText, 10, Len(MyText) - 9)
                    Call WarpToMe(MyText)
                End If
                MyText = ""
                Exit Sub
                End If
            End If
                        
            ' Warping to a map
            If LCase(Mid(MyText, 1, 7)) = "/warpto" Then
            If InStr(1, Player(MyIndex).AdminCmds, LCase(Mid(MyText, 1, 7))) Then
                If Len(MyText) > 8 Then
                    MyText = Mid(MyText, 8, Len(MyText) - 7)
                    n = Val(MyText)
                
                    ' Check to make sure its a valid map #
                    If n > 0 And n <= MAX_MAPS Then
                        Call WarpTo(n)
                    Else
                        Call AddText("Invalid map number.", Red)
                    End If
                End If
                MyText = ""
                Exit Sub
                End If
            End If
            
            ' Setting sprite
            If LCase(Mid(MyText, 1, 10)) = "/setsprite" Then
            If InStr(1, Player(MyIndex).AdminCmds, LCase(Mid(MyText, 1, 10))) Then
                If Len(MyText) > 11 Then
                    ' Get sprite #
                    MyText = Mid(MyText, 12, Len(MyText) - 11)
                
                    Call SendSetSprite(Val(MyText))
                End If
                MyText = ""
                Exit Sub
                End If
            End If
            
            ' Map report
            If LCase(Mid(MyText, 1, 10)) = "/mapreport" Then
            If InStr(1, Player(MyIndex).AdminCmds, LCase(Mid(MyText, 1, 10))) Then
                Call SendData("mapreport" & SEP_CHAR & END_CHAR)
                MyText = ""
                Exit Sub
                End If
            End If
        
            ' Respawn request
            If Mid(MyText, 1, 8) = "/respawn" Then
            If InStr(1, Player(MyIndex).AdminCmds, LCase(Mid(MyText, 1, 8))) Then
                Call SendMapRespawn
                MyText = ""
                Exit Sub
                End If
            End If
        
            ' MOTD change
            If Mid(MyText, 1, 5) = "/motd" Then
            If InStr(1, Player(MyIndex).AdminCmds, LCase(Mid(MyText, 1, 5))) Then
                If Len(MyText) > 6 Then
                    MyText = Mid(MyText, 7, Len(MyText) - 6)
                    If Trim(MyText) <> "" Then
                        Call SendMOTDChange(MyText)
                    End If
                End If
                MyText = ""
                Exit Sub
                End If
            End If
            
            ' Check the ban list
            If Mid(MyText, 1, 8) = "/banlist" Then
            If InStr(1, Player(MyIndex).AdminCmds, LCase(Mid(MyText, 1, 8))) Then
                Call SendBanList
                MyText = ""
                Exit Sub
            End If
            End If
            
            ' Banning a player
            If LCase(Mid(MyText, 1, 4)) = "/ban" Then
            If InStr(1, Player(MyIndex).AdminCmds, LCase(Mid(MyText, 1, 4))) Then
                If Len(MyText) > 5 Then
                    MyText = Mid(MyText, 6, Len(MyText) - 5)
                    Call SendBan(MyText)
                    MyText = ""
                End If
                Exit Sub
                End If
            End If
                    
        ' // Developer Admin Commands //
             ' Editing item request
            If Mid(MyText, 1, 9) = "/edititem" Then
            If InStr(1, Player(MyIndex).AdminCmds, LCase(Mid(MyText, 1, 9))) Then
                Call SendRequestEditItem
                MyText = ""
                Exit Sub
                End If
            End If
            
            ' Editing npc request
            If Mid(MyText, 1, 8) = "/editnpc" Then
            If InStr(1, Player(MyIndex).AdminCmds, LCase(Mid(MyText, 1, 8))) Then
                Call SendRequestEditNpc
                MyText = ""
                Exit Sub
                End If
            End If
            
            ' Editing shop request
            If Mid(MyText, 1, 9) = "/editshop" Then
            If InStr(1, Player(MyIndex).AdminCmds, LCase(Mid(MyText, 1, 9))) Then
                Call SendRequestEditShop
                MyText = ""
                Exit Sub
            End If
            End If
        
            ' Editing spell request
            If Mid(MyText, 1, 10) = "/editspell" Then
            If InStr(1, Player(MyIndex).AdminCmds, LCase(Mid(MyText, 1, 10))) Then
                Call SendRequestEditSpell
                MyText = ""
                Exit Sub
            End If
        End If
        
        ' // Creator Admin Commands //
          ' Giving another player access
            If LCase(Mid(MyText, 1, 10)) = "/setaccess" Then
            If InStr(1, Player(MyIndex).AdminCmds, LCase(Mid(MyText, 1, 10))) Then
                ' Get access #
                i = Val(Mid(MyText, 12, 1))
                
                MyText = Mid(MyText, 14, Len(MyText) - 13)
                
                Call SendSetAccess(MyText, i)
                MyText = ""
                Exit Sub
            End If
            End If
            ' Ban destroy
             If LCase(Mid(MyText, 1, 15)) = "/destroybanlist" Then
             If InStr(1, Player(MyIndex).AdminCmds, LCase(Mid(MyText, 1, 15))) Then
                Call SendBanDestroy
                MyText = ""
                Exit Sub
            End If
        End If
        
        ' Say message
        If Len(Trim(MyText)) > 0 Then
            Call SayMsg(MyText)
        End If
        MyText = ""
        Exit Sub
    End If
    
    ' Handle when the user presses the backspace key
    If (KeyAscii = vbKeyBack) Then
        If Len(MyText) > 0 Then
            MyText = Mid(MyText, 1, Len(MyText) - 1)
        End If
    End If
    
    ' And if neither, then add the character to the user's text buffer
    If (KeyAscii <> vbKeyReturn) And (KeyAscii <> vbKeyBack) Then
        ' Make sure they just use standard keys, no gay shitty macro keys
        If KeyAscii >= 32 And KeyAscii <= 126 Then
            MyText = MyText & Chr(KeyAscii)
        End If
    End If
End Sub

Sub CheckMapGetItem()
    If GetTickCount > Player(MyIndex).MapGetTimer + 250 And Trim(MyText) = "" Then
        Player(MyIndex).MapGetTimer = GetTickCount
        Call SendData("mapgetitem" & SEP_CHAR & END_CHAR)
    End If
    'If
    
End Sub

Sub CheckAttack()
    If ControlDown = True And Player(MyIndex).AttackTimer + 1000 < GetTickCount And Player(MyIndex).Attacking = 0 Then
        Player(MyIndex).Attacking = 1
        Player(MyIndex).AttackTimer = GetTickCount
        Call SendData("attack" & SEP_CHAR & END_CHAR)
    End If
End Sub

Sub CheckInput2()
    If GettingMap = False Then
        If GetKeyState(VK_RETURN) < 0 Then
            Call CheckMapGetItem
        End If
        If GetKeyState(VK_CONTROL) < 0 Then
            ControlDown = True
        Else
            ControlDown = False
        End If
        If GetKeyState(VK_UP) < 0 Then
            DirUp = True
            DirDown = False
            DirLeft = False
            DirRight = False
        Else
            DirUp = False
        End If
        If GetKeyState(VK_DOWN) < 0 Then
            DirUp = False
            DirDown = True
            DirLeft = False
            DirRight = False
        Else
            DirDown = False
        End If
        If GetKeyState(VK_LEFT) < 0 Then
            DirUp = False
            DirDown = False
            DirLeft = True
            DirRight = False
        Else
            DirLeft = False
        End If
        If GetKeyState(VK_RIGHT) < 0 Then
            DirUp = False
            DirDown = False
            DirLeft = False
            DirRight = True
        Else
            DirRight = False
        End If
        If GetKeyState(VK_SHIFT) < 0 Then
            ShiftDown = True
        Else
            ShiftDown = False
        End If
    End If
End Sub

Sub CheckInput(ByVal KeyState As Byte, ByVal KeyCode As Integer, ByVal Shift As Integer)
    If GettingMap = False Then
        If KeyState = 1 Then
            If KeyCode = vbKeyReturn Then
                Call CheckMapGetItem
            End If
            If KeyCode = vbKeyControl Then
                ControlDown = True
            End If
            If KeyCode = vbKeyF1 Then
                Call CastSpell1
            End If
            If KeyCode = vbKeyF2 Then
                Call CastSpell2
            End If
            If KeyCode = vbKeyF3 Then
                Call CastSpell3
            End If
            If KeyCode = vbKeyF4 Then
                Call Screenshot
            End If
            If KeyCode = vbKeyUp Then
                DirUp = True
                DirDown = False
                DirLeft = False
                DirRight = False
            End If
            If KeyCode = vbKeyDown Then
                DirUp = False
                DirDown = True
                DirLeft = False
                DirRight = False
            End If
            If KeyCode = vbKeyLeft Then
                DirUp = False
                DirDown = False
                DirLeft = True
                DirRight = False
            End If
            If KeyCode = vbKeyRight Then
                DirUp = False
                DirDown = False
                DirLeft = False
                DirRight = True
            End If
            If KeyCode = vbKeyShift Then
                ShiftDown = True
            End If
        Else
            If KeyCode = vbKeyUp Then DirUp = False
            If KeyCode = vbKeyDown Then DirDown = False
            If KeyCode = vbKeyLeft Then DirLeft = False
            If KeyCode = vbKeyRight Then DirRight = False
            If KeyCode = vbKeyShift Then ShiftDown = False
            If KeyCode = vbKeyControl Then ControlDown = False
        End If
    End If
End Sub

Function IsTryingToMove() As Boolean
    If (DirUp = True) Or (DirDown = True) Or (DirLeft = True) Or (DirRight = True) Then
        IsTryingToMove = True
    Else
        IsTryingToMove = False
    End If
End Function

Function CanMove() As Boolean
Dim i As Long, d As Long

    CanMove = True
    
    ' Make sure they aren't trying to move when they are already moving
    If Player(MyIndex).Moving <> 0 Then
        CanMove = False
        Exit Function
    End If
    
    ' Make sure they haven't just casted a spell
    If Player(MyIndex).CastedSpell = YES Then
        If GetTickCount > Player(MyIndex).AttackTimer + 1000 Then
            Player(MyIndex).CastedSpell = NO
        Else
            CanMove = False
            Exit Function
        End If
    End If
    
    d = GetPlayerDir(MyIndex)
    If DirUp Then
        Call SetPlayerDir(MyIndex, DIR_UP)
        
        ' Check to see if they are trying to go out of bounds
        If GetPlayerY(MyIndex) > 0 Then
            ' Check to see if the map tile is blocked or not
            If Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).Type = TILE_TYPE_BLOCKED Then
                CanMove = False
                
                ' Set the new direction if they weren't facing that direction
                If d <> DIR_UP Then
                    Call SendPlayerDir
                End If
                Exit Function
            End If
                                        
            ' Check to see if the key door is open or not
            If Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).Type = TILE_TYPE_KEY Then
                ' This actually checks if its open or not
                If TempTile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).DoorOpen = NO Then
                    CanMove = False
                
                    ' Set the new direction if they weren't facing that direction
                    If d <> DIR_UP Then
                        Call SendPlayerDir
                    End If
                    Exit Function
                End If
            End If
            
            ' Check to see if a player is already on that tile
            For i = 1 To MAX_PLAYERS
                If IsPlaying(i) Then
                    If GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                        If (GetPlayerX(i) = GetPlayerX(MyIndex)) And (GetPlayerY(i) = GetPlayerY(MyIndex) - 1) Then
                            CanMove = False
                        
                            ' Set the new direction if they weren't facing that direction
                            If d <> DIR_UP Then
                                Call SendPlayerDir
                            End If
                            Exit Function
                        End If
                    End If
                End If
            Next i
        
            ' Check to see if a npc is already on that tile
            For i = 1 To MAX_MAP_NPCS
                If MapNpc(i).num > 0 Then
                    If (MapNpc(i).X = GetPlayerX(MyIndex)) And (MapNpc(i).y = GetPlayerY(MyIndex) - 1) Then
                        CanMove = False
                        
                        ' Set the new direction if they weren't facing that direction
                        If d <> DIR_UP Then
                            Call SendPlayerDir
                        End If
                        Exit Function
                    End If
                End If
            Next i
        Else
            ' Check if they can warp to a new map
            If Map.Up > 0 Then
                Call SendPlayerRequestNewMap
                GettingMap = True
            End If
            CanMove = False
            Exit Function
        End If
    End If
            
    If DirDown Then
        Call SetPlayerDir(MyIndex, DIR_DOWN)
        
        ' Check to see if they are trying to go out of bounds
        If GetPlayerY(MyIndex) < MAX_MAPY Then
            ' Check to see if the map tile is blocked or not
            If Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) + 1).Type = TILE_TYPE_BLOCKED Then
                CanMove = False
                
                ' Set the new direction if they weren't facing that direction
                If d <> DIR_DOWN Then
                    Call SendPlayerDir
                End If
                Exit Function
            End If
                                        
            ' Check to see if the key door is open or not
            If Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) + 1).Type = TILE_TYPE_KEY Then
                ' This actually checks if its open or not
                If TempTile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) + 1).DoorOpen = NO Then
                    CanMove = False
                
                    ' Set the new direction if they weren't facing that direction
                    If d <> DIR_DOWN Then
                        Call SendPlayerDir
                    End If
                    Exit Function
                End If
            End If
            
            ' Check to see if a player is already on that tile
            For i = 1 To MAX_PLAYERS
                If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                    If (GetPlayerX(i) = GetPlayerX(MyIndex)) And (GetPlayerY(i) = GetPlayerY(MyIndex) + 1) Then
                        CanMove = False
                        
                        ' Set the new direction if they weren't facing that direction
                        If d <> DIR_DOWN Then
                            Call SendPlayerDir
                        End If
                        Exit Function
                    End If
                End If
            Next i
            
            ' Check to see if a npc is already on that tile
            For i = 1 To MAX_MAP_NPCS
                If MapNpc(i).num > 0 Then
                    If (MapNpc(i).X = GetPlayerX(MyIndex)) And (MapNpc(i).y = GetPlayerY(MyIndex) + 1) Then
                        CanMove = False
                        
                        ' Set the new direction if they weren't facing that direction
                        If d <> DIR_DOWN Then
                            Call SendPlayerDir
                        End If
                        Exit Function
                    End If
                End If
            Next i
        Else
            ' Check if they can warp to a new map
            If Map.Down > 0 Then
                Call SendPlayerRequestNewMap
                GettingMap = True
            End If
            CanMove = False
            Exit Function
        End If
    End If
                
    If DirLeft Then
        Call SetPlayerDir(MyIndex, DIR_LEFT)
        
        ' Check to see if they are trying to go out of bounds
        If GetPlayerX(MyIndex) > 0 Then
            ' Check to see if the map tile is blocked or not
            If Map.Tile(GetPlayerX(MyIndex) - 1, GetPlayerY(MyIndex)).Type = TILE_TYPE_BLOCKED Then
                CanMove = False
                
                ' Set the new direction if they weren't facing that direction
                If d <> DIR_LEFT Then
                    Call SendPlayerDir
                End If
                Exit Function
            End If
                                        
            ' Check to see if the key door is open or not
            If Map.Tile(GetPlayerX(MyIndex) - 1, GetPlayerY(MyIndex)).Type = TILE_TYPE_KEY Then
                ' This actually checks if its open or not
                If TempTile(GetPlayerX(MyIndex) - 1, GetPlayerY(MyIndex)).DoorOpen = NO Then
                    CanMove = False
                    
                    ' Set the new direction if they weren't facing that direction
                    If d <> DIR_LEFT Then
                        Call SendPlayerDir
                    End If
                    Exit Function
                End If
            End If
            
            ' Check to see if a player is already on that tile
            For i = 1 To MAX_PLAYERS
                If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                    If (GetPlayerX(i) = GetPlayerX(MyIndex) - 1) And (GetPlayerY(i) = GetPlayerY(MyIndex)) Then
                        CanMove = False
                        
                        ' Set the new direction if they weren't facing that direction
                        If d <> DIR_LEFT Then
                            Call SendPlayerDir
                        End If
                        Exit Function
                    End If
                End If
            Next i
        
            ' Check to see if a npc is already on that tile
            For i = 1 To MAX_MAP_NPCS
                If MapNpc(i).num > 0 Then
                    If (MapNpc(i).X = GetPlayerX(MyIndex) - 1) And (MapNpc(i).y = GetPlayerY(MyIndex)) Then
                        CanMove = False
                        
                        ' Set the new direction if they weren't facing that direction
                        If d <> DIR_LEFT Then
                            Call SendPlayerDir
                        End If
                        Exit Function
                    End If
                End If
            Next i
        Else
            ' Check if they can warp to a new map
            If Map.Left > 0 Then
                Call SendPlayerRequestNewMap
                GettingMap = True
            End If
            CanMove = False
            Exit Function
        End If
    End If
        
    If DirRight Then
        Call SetPlayerDir(MyIndex, DIR_RIGHT)
        
        ' Check to see if they are trying to go out of bounds
        If GetPlayerX(MyIndex) < MAX_MAPX Then
            ' Check to see if the map tile is blocked or not
            If Map.Tile(GetPlayerX(MyIndex) + 1, GetPlayerY(MyIndex)).Type = TILE_TYPE_BLOCKED Then
                CanMove = False
                
                ' Set the new direction if they weren't facing that direction
                If d <> DIR_RIGHT Then
                    Call SendPlayerDir
                End If
                Exit Function
            End If
                                        
            ' Check to see if the key door is open or not
            If Map.Tile(GetPlayerX(MyIndex) + 1, GetPlayerY(MyIndex)).Type = TILE_TYPE_KEY Then
                ' This actually checks if its open or not
                If TempTile(GetPlayerX(MyIndex) + 1, GetPlayerY(MyIndex)).DoorOpen = NO Then
                    CanMove = False
                    
                    ' Set the new direction if they weren't facing that direction
                    If d <> DIR_RIGHT Then
                        Call SendPlayerDir
                    End If
                    Exit Function
                End If
            End If
            
            ' Check to see if a player is already on that tile
            For i = 1 To MAX_PLAYERS
                If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                    If (GetPlayerX(i) = GetPlayerX(MyIndex) + 1) And (GetPlayerY(i) = GetPlayerY(MyIndex)) Then
                        CanMove = False
                        
                        ' Set the new direction if they weren't facing that direction
                        If d <> DIR_RIGHT Then
                            Call SendPlayerDir
                        End If
                        Exit Function
                    End If
                End If
            Next i
        
            ' Check to see if a npc is already on that tile
            For i = 1 To MAX_MAP_NPCS
                If MapNpc(i).num > 0 Then
                    If (MapNpc(i).X = GetPlayerX(MyIndex) + 1) And (MapNpc(i).y = GetPlayerY(MyIndex)) Then
                        CanMove = False
                        
                        ' Set the new direction if they weren't facing that direction
                        If d <> DIR_RIGHT Then
                            Call SendPlayerDir
                        End If
                        Exit Function
                    End If
                End If
            Next i
        Else
            ' Check if they can warp to a new map
            If Map.Right > 0 Then
                Call SendPlayerRequestNewMap
                GettingMap = True
            End If
            CanMove = False
            Exit Function
        End If
    End If
End Function

Sub CheckMovement()
    If GettingMap = False Then
        If IsTryingToMove Then
            If CanMove Then
                ' Check if player has the shift key down for running
                If ShiftDown Then
                    Player(MyIndex).Moving = MOVING_RUNNING
                Else
                    Player(MyIndex).Moving = MOVING_WALKING
                End If
            
                Select Case GetPlayerDir(MyIndex)
                    Case DIR_UP
                        Call SendPlayerMove
                        Player(MyIndex).YOffset = PIC_Y
                        Call SetPlayerY(MyIndex, GetPlayerY(MyIndex) - 1)
                
                    Case DIR_DOWN
                        Call SendPlayerMove
                        Player(MyIndex).YOffset = PIC_Y * -1
                        Call SetPlayerY(MyIndex, GetPlayerY(MyIndex) + 1)
                
                    Case DIR_LEFT
                        Call SendPlayerMove
                        Player(MyIndex).XOffset = PIC_X
                        Call SetPlayerX(MyIndex, GetPlayerX(MyIndex) - 1)
                
                    Case DIR_RIGHT
                        Call SendPlayerMove
                        Player(MyIndex).XOffset = PIC_X * -1
                        Call SetPlayerX(MyIndex, GetPlayerX(MyIndex) + 1)
                End Select
            
                ' Gotta check :)
                If Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex)).Type = TILE_TYPE_WARP Then
                    GettingMap = True
                End If
            End If
        End If
    End If
End Sub

Function FindPlayer(ByVal name As String) As Long
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            ' Make sure we dont try to check a name thats to small
            If Len(GetPlayerName(i)) >= Len(Trim(name)) Then
                If UCase(Mid(GetPlayerName(i), 1, Len(Trim(name)))) = UCase(Trim(name)) Then
                    FindPlayer = i
                    Exit Function
                End If
            End If
        End If
    Next i
    
    FindPlayer = 0
End Function

Public Sub EditorInit()
    SaveMap = Map
    InEditor = True
    frmMirage.picMapEditor.Visible = True
    With frmMirage.picBackSelect
        .Width = 7 * PIC_X
        .Height = 500 * PIC_Y
        .Picture = LoadPicture(App.Path + "\tiles.bmp")
    End With
End Sub

Public Sub EditorMouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim x1, y1 As Long

    If InEditor Then
        x1 = Int(X / PIC_X)
        y1 = Int(y / PIC_Y)
        If (Button = 1) And (x1 >= 0) And (x1 <= MAX_MAPX) And (y1 >= 0) And (y1 <= MAX_MAPY) Then
            If frmMirage.optLayers.value = True Then
                With Map.Tile(x1, y1)
                frmMirage.infolbl.Caption = EditorTileY * 7 + EditorTileX
                    If frmMirage.optGround.value = True Then .Ground = EditorTileY * 7 + EditorTileX
                    If frmMirage.optMask.value = True Then .Mask = EditorTileY * 7 + EditorTileX
                    If frmMirage.optAnim.value = True Then .Anim = EditorTileY * 7 + EditorTileX
                    If frmMirage.optMask2.value = True Then .Mask2 = EditorTileY * 7 + EditorTileX
                    If frmMirage.optM2Anim.value = True Then .M2Anim = EditorTileY * 7 + EditorTileX
                    If frmMirage.optFringe.value = True Then .Fringe = EditorTileY * 7 + EditorTileX
                    If frmMirage.optFAnim.value = True Then .FAnim = EditorTileY * 7 + EditorTileX
                    If frmMirage.optFringe2.value = True Then .Fringe2 = EditorTileY * 7 + EditorTileX
                    If frmMirage.optF2Anim.value = True Then .F2Anim = EditorTileY * 7 + EditorTileX
                End With
            Else
                With Map.Tile(x1, y1)
                    If frmMirage.optBlocked.value = True Then .Type = TILE_TYPE_BLOCKED
                    If frmMirage.optWarp.value = True Then
                        .Type = TILE_TYPE_WARP
                        .Data1 = EditorWarpMap
                        .Data2 = EditorWarpX
                        .Data3 = EditorWarpY
                    End If
                    If frmMirage.optItem.value = True Then
                        .Type = TILE_TYPE_ITEM
                        .Data1 = ItemEditorNum
                        .Data2 = ItemEditorValue
                        .Data3 = 0
                    End If
                    If frmMirage.optNpcAvoid.value = True Then
                        .Type = TILE_TYPE_NPCAVOID
                        .Data1 = 0
                        .Data2 = 0
                        .Data3 = 0
                    End If
                    If frmMirage.optKey.value = True Then
                        .Type = TILE_TYPE_KEY
                        .Data1 = KeyEditorNum
                        .Data2 = KeyEditorTake
                        .Data3 = 0
                    End If
                    If frmMirage.optKeyOpen.value = True Then
                        .Type = TILE_TYPE_KEYOPEN
                        .Data1 = KeyOpenEditorX
                        .Data2 = KeyOpenEditorY
                        .Data3 = 0
                    End If
                    If frmMirage.Optsign.value = True Then
                        .Type = TILE_TYPE_SIGN
                        .Data1 = SignTitle
                        .Data2 = SignText
                        .Data3 = SoundFileS
                    End If
                    If frmMirage.fheal.value = True Then
                        .Type = TILE_TYPE_FULLHEAL
                        .Data1 = 0
                        .Data2 = 0
                        .Data3 = SoundFileH
                    End If
                    If frmMirage.deathopt.value = True Then
                        .Type = TILE_TYPE_DEATH
                        .Data1 = Deathsay
                        .Data2 = SoundFileD
                        .Data3 = 0
                    End If
                    If frmMirage.SFXopt.value = True Then
                        .Type = TILE_TYPE_SFX
                        .Data1 = SoundFile
                        .Data2 = 0
                        .Data3 = Deepness
                    End If
                    If frmMirage.optQuest.value = True Then
                        .Type = TILE_TYPE_QUEST
                        .Data1 = Quest1
                        .Data2 = Quest2
                        .Data3 = Quest3
                    End If
                    If frmMirage.optArena.value = True Then
                        .Type = TILE_TYPE_ARENA
                        .Data1 = 0
                        .Data2 = 0
                        .Data3 = 0
                    End If
                    
                    If frmMirage.optFishing.value = True Then
                        .Type = TILE_TYPE_FISHING
                        .Data1 = Fishing1
                        .Data2 = Fishing2
                        .Data3 = Fishing3
                    End If
                    
                    If frmMirage.optMining.value = True Then
                        .Type = TILE_TYPE_MINING
                        .Data1 = Mining1
                        .Data2 = Mining2
                        .Data3 = Mining3
                    End If
                        
                    
                End With
            End If
        End If
        
        If (Button = 2) And (x1 >= 0) And (x1 <= MAX_MAPX) And (y1 >= 0) And (y1 <= MAX_MAPY) Then
            If frmMirage.optLayers.value = True Then
                With Map.Tile(x1, y1)
                     If frmMirage.optGround.value = True Then .Ground = 0
                    If frmMirage.optMask.value = True Then .Mask = 0
                    If frmMirage.optAnim.value = True Then .Anim = 0
                    If frmMirage.optMask2.value = True Then .Mask2 = 0
                    If frmMirage.optM2Anim.value = True Then .M2Anim = 0
                    If frmMirage.optFringe.value = True Then .Fringe = 0
                    If frmMirage.optFAnim.value = True Then .FAnim = 0
                    If frmMirage.optFringe2.value = True Then .Fringe2 = 0
                    If frmMirage.optF2Anim.value = True Then .F2Anim = 0
                End With
            Else
                With Map.Tile(x1, y1)
                    .Type = 0
                    .Data1 = 0
                    .Data2 = 0
                    .Data3 = 0
                End With
            End If
        End If
    End If
End Sub

Public Sub EditorChooseTile(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button = 1 Then
        EditorTileX = Int(X / PIC_X)
        EditorTileY = Int(y / PIC_Y)
    End If
    Call BitBlt(frmMirage.picSelect.hdc, 0, 0, PIC_X, PIC_Y, frmMirage.picBackSelect.hdc, EditorTileX * PIC_X, EditorTileY * PIC_Y, SRCCOPY)
frmMirage.infolbl.Caption = EditorTileY * 7 + EditorTileX
End Sub

Public Sub EditorTileScroll()
    frmMirage.picBackSelect.top = (frmMirage.scrlPicture.value * PIC_Y) * -1
End Sub

Public Sub EditorSend()
    Call SendMap
    Call EditorCancel
End Sub

Public Sub EditorCancel()
    Map = SaveMap
    InEditor = False
    frmMirage.picMapEditor.Visible = False
End Sub

Public Sub EditorClearLayer()
Dim YesNo As Long, X As Long, y As Long
'EditorTileY * 7 + EditorTileX

    ' Ground layer
    If frmMirage.optGround.value = True Then
        YesNo = MsgBox("Are you sure you wish to clear the ground layer?", vbYesNo, GAME_NAME)
        
        If YesNo = vbYes Then
            For y = 0 To MAX_MAPY
                For X = 0 To MAX_MAPX
                    Map.Tile(X, y).Ground = 0
                Next X
            Next y
        End If
    End If

    ' Mask layer
    If frmMirage.optMask.value = True Then
        YesNo = MsgBox("Are you sure you wish to clear the mask layer?", vbYesNo, GAME_NAME)
        
        If YesNo = vbYes Then
            For y = 0 To MAX_MAPY
                For X = 0 To MAX_MAPX
                    Map.Tile(X, y).Mask = 0
                Next X
            Next y
        End If
    End If

    ' Mask Animation layer
    If frmMirage.optAnim.value = True Then
        YesNo = MsgBox("Are you sure you wish to clear the animation layer?", vbYesNo, GAME_NAME)
        
        If YesNo = vbYes Then
            For y = 0 To MAX_MAPY
                For X = 0 To MAX_MAPX
                    Map.Tile(X, y).Anim = 0
                Next X
            Next y
        End If
    End If

    ' Mask 2 layer
    If frmMirage.optMask2.value = True Then
        YesNo = MsgBox("Are you sure you wish to clear the mask 2 layer?", vbYesNo, GAME_NAME)
        
        If YesNo = vbYes Then
            For y = 0 To MAX_MAPY
                For X = 0 To MAX_MAPX
                    Map.Tile(X, y).Mask2 = 0
                Next X
            Next y
        End If
    End If

    ' Mask 2 Animation layer
    If frmMirage.optM2Anim.value = True Then
        YesNo = MsgBox("Are you sure you wish to clear the mask 2 animation layer?", vbYesNo, GAME_NAME)
        
        If YesNo = vbYes Then
            For y = 0 To MAX_MAPY
                For X = 0 To MAX_MAPX
                    Map.Tile(X, y).M2Anim = 0
                Next X
            Next y
        End If
    End If

    ' Fringe layer
    If frmMirage.optFringe.value = True Then
        YesNo = MsgBox("Are you sure you wish to clear the fringe layer?", vbYesNo, GAME_NAME)
        
        If YesNo = vbYes Then
            For y = 0 To MAX_MAPY
                For X = 0 To MAX_MAPX
                    Map.Tile(X, y).Fringe = 0
                Next X
            Next y
        End If
    End If

    ' Fringe Animation layer
    If frmMirage.optFAnim.value = True Then
        YesNo = MsgBox("Are you sure you wish to clear the fringe animation layer?", vbYesNo, GAME_NAME)
        
        If YesNo = vbYes Then
            For y = 0 To MAX_MAPY
                For X = 0 To MAX_MAPX
                    Map.Tile(X, y).FAnim = 0
                Next X
            Next y
        End If
    End If

    ' Fringe 2 layer
    If frmMirage.optFringe2.value = True Then
        YesNo = MsgBox("Are you sure you wish to clear the fringe 2 layer?", vbYesNo, GAME_NAME)
        
        If YesNo = vbYes Then
            For y = 0 To MAX_MAPY
                For X = 0 To MAX_MAPX
                    Map.Tile(X, y).Fringe2 = 0
                Next X
            Next y
        End If
    End If

    ' Fringe 2 Animation layer
    If frmMirage.optF2Anim.value = True Then
        YesNo = MsgBox("Are you sure you wish to clear the fringe 2 animation layer?", vbYesNo, GAME_NAME)
        
        If YesNo = vbYes Then
            For y = 0 To MAX_MAPY
                For X = 0 To MAX_MAPX
                    Map.Tile(X, y).F2Anim = 0
                Next X
            Next y
        End If
    End If
End Sub
Public Sub EditorFillLayer()
Dim YesNo As Long, X As Long, y As Long
'EditorTileY * 7 + EditorTileX

    ' Ground layer
    If frmMirage.optGround.value = True Then
        YesNo = MsgBox("Are you sure you wish to Fill the ground layer?", vbYesNo, GAME_NAME)
        
        If YesNo = vbYes Then
            For y = 0 To MAX_MAPY
                For X = 0 To MAX_MAPX
                    Map.Tile(X, y).Ground = EditorTileY * 7 + EditorTileX
                Next X
            Next y
        End If
    End If

    ' Mask layer
    If frmMirage.optMask.value = True Then
        YesNo = MsgBox("Are you sure you wish to Fill the mask layer?", vbYesNo, GAME_NAME)
        
        If YesNo = vbYes Then
            For y = 0 To MAX_MAPY
                For X = 0 To MAX_MAPX
                    Map.Tile(X, y).Mask = EditorTileY * 7 + EditorTileX
                Next X
            Next y
        End If
    End If

    ' Animation layer
    If frmMirage.optAnim.value = True Then
        YesNo = MsgBox("Are you sure you wish to Fill the animation layer?", vbYesNo, GAME_NAME)
        
        If YesNo = vbYes Then
            For y = 0 To MAX_MAPY
                For X = 0 To MAX_MAPX
                    Map.Tile(X, y).Anim = EditorTileY * 7 + EditorTileX
                Next X
            Next y
        End If
    End If

    ' Fringe layer
    If frmMirage.optFringe.value = True Then
        YesNo = MsgBox("Are you sure you wish to Fill the fringe layer?", vbYesNo, GAME_NAME)
        
        If YesNo = vbYes Then
            For y = 0 To MAX_MAPY
                For X = 0 To MAX_MAPX
                    Map.Tile(X, y).Fringe = EditorTileY * 7 + EditorTileX
                Next X
            Next y
        End If
    End If
End Sub

Public Sub EditorClearAttribs()
Dim YesNo As Long, X As Long, y As Long

    YesNo = MsgBox("Are you sure you wish to clear the attributes on this map?", vbYesNo, GAME_NAME)
    
    If YesNo = vbYes Then
        For y = 0 To MAX_MAPY
            For X = 0 To MAX_MAPX
                Map.Tile(X, y).Type = 0
            Next X
        Next y
    End If
End Sub
Public Sub EditorFillAttribs()
Dim YesNo As Long, X As Long, y As Long












    YesNo = MsgBox("Are you sure you wish to fill the attributes on this map?", vbYesNo, GAME_NAME)
    
    If YesNo = vbYes Then
        For y = 0 To MAX_MAPY
            For X = 0 To MAX_MAPX
                Map.Tile(X, y).Type = 0
            Next X
        Next y
    End If


End Sub



Public Sub ItemEditorInit()
On Error Resume Next
    
    frmItemEditor.picItems.Picture = LoadPicture(App.Path & "\items.bmp")
    frmItemEditor.picItems2.Picture = LoadPicture(App.Path & "\sprites.bmp")
    'frmMirage.picItems.Picture = LoadPicture(App.Path & "\items.bmp")
    
    frmItemEditor.txtName.Text = Trim(Item(EditorIndex).name)
    frmItemEditor.Text4.Text = Trim(Item(EditorIndex).SellValue)
    frmItemEditor.txtND.Text = Trim(Item(EditorIndex).NoDrop)
    frmItemEditor.cmbClass.ListIndex = Item(EditorIndex).Class
    frmItemEditor.scrlPic.value = Item(EditorIndex).Pic
    frmItemEditor.cmbType.ListIndex = Item(EditorIndex).Type
    frmItemEditor.Text3.Text = Item(EditorIndex).SPRITE
    If (frmItemEditor.cmbType.ListIndex = ITEM_TYPE_WEAPON) Then
    frmItemEditor.scrlsound = Item(EditorIndex).Data3
    End If
    
    
    If (frmItemEditor.cmbType.ListIndex >= ITEM_TYPE_WEAPON) And (frmItemEditor.cmbType.ListIndex <= ITEM_TYPE_SHIELD) Then
        frmItemEditor.fraEquipment.Visible = True
        frmItemEditor.scrlDurability.value = Item(EditorIndex).Data1
        frmItemEditor.scrlStrength.value = Item(EditorIndex).Data2
        
        frmItemEditor.STRMODER.value = Item(EditorIndex).strmod
        frmItemEditor.DEFMODER.value = Item(EditorIndex).defmod
        frmItemEditor.MAGIMODER.value = Item(EditorIndex).magimod
    Else
        frmItemEditor.fraEquipment.Visible = False
    End If
    
    If frmItemEditor.cmbType.ListIndex = ITEM_TYPE_WEAPON Then
    frmItemEditor.Label15.Visible = True
    frmItemEditor.scrlsound.Visible = True
    frmItemEditor.lblsound.Visible = True
    Else
    frmItemEditor.Label15.Visible = False
    frmItemEditor.scrlsound.Visible = False
    frmItemEditor.lblsound.Visible = False
    End If
    
    
    If (frmItemEditor.cmbType.ListIndex >= ITEM_TYPE_POTIONADDHP) And (frmItemEditor.cmbType.ListIndex <= ITEM_TYPE_POTIONSUBSP) Then
        frmItemEditor.fraVitals.Visible = True
        frmItemEditor.scrlVitalMod.value = Item(EditorIndex).Data1
    Else
        frmItemEditor.fraVitals.Visible = False
    End If
    
    If (frmItemEditor.cmbType.ListIndex = ITEM_TYPE_SPELL) Then
        frmItemEditor.fraSpell.Visible = True
        frmItemEditor.scrlSpell.value = Item(EditorIndex).Data1
    Else
        frmItemEditor.fraSpell.Visible = False
        
    End If
    
    If (frmItemEditor.cmbType.ListIndex = ITEM_TYPE_MASKEDIT) Then
        frmItemEditor.FRMmaskedit.Visible = True
        frmItemEditor.Text2.Text = Item(EditorIndex).Data1
        frmItemEditor.fringemask.value = Item(EditorIndex).Data1
    Else
        frmItemEditor.FRMmaskedit.Visible = False
        
    End If
    If (frmItemEditor.cmbType.ListIndex = ITEM_TYPE_SCHANGE) Or (frmItemEditor.cmbType.ListIndex = ITEM_TYPE_BARBER) Then
        frmItemEditor.SChanger.Visible = True
        frmItemEditor.HScroll1.value = Item(EditorIndex).Data1
    Else
        frmItemEditor.SChanger.Visible = False
        
    End If
    
    If (frmItemEditor.cmbType.ListIndex >= ITEM_TYPE_RANDITEM) Then
    frmItemEditor.randframe.Visible = True
         frmItemEditor.HScroll2.value = Item(EditorIndex).Data1
         frmItemEditor.HScroll3.value = Item(EditorIndex).Data2
         frmItemEditor.HScroll4.value = Item(EditorIndex).Data3
    Else
        frmItemEditor.randframe.Visible = False
    End If
    
    
    If (frmItemEditor.cmbType.ListIndex = ITEM_TYPE_BOOK) Then
        frmItemEditor.ringammy.Visible = True
        frmItemEditor.Text1.Text = Item(EditorIndex).Data2
    Else
        frmItemEditor.ringammy.Visible = False
        
    End If
    
    
    
    frmItemEditor.Show vbModal
End Sub

Public Sub ItemEditorOk()
    Item(EditorIndex).name = frmItemEditor.txtName.Text
    Item(EditorIndex).Pic = frmItemEditor.scrlPic.value
    Item(EditorIndex).Type = frmItemEditor.cmbType.ListIndex
    Item(EditorIndex).Class = frmItemEditor.cmbClass.ListIndex
    Item(EditorIndex).SellValue = Val(frmItemEditor.Text4.Text)
    Item(EditorIndex).NoDrop = Val(frmItemEditor.txtND.Text)

    If (frmItemEditor.cmbType.ListIndex >= ITEM_TYPE_WEAPON) And (frmItemEditor.cmbType.ListIndex <= ITEM_TYPE_SHIELD) Then
        Item(EditorIndex).Data1 = frmItemEditor.scrlDurability.value
        Item(EditorIndex).Data2 = frmItemEditor.scrlStrength.value
        Item(EditorIndex).Data3 = 0
        Item(EditorIndex).Data4 = frmItemEditor.scrlRange.value
        Item(EditorIndex).spdmod = frmItemEditor.SPDMODER.value
        Item(EditorIndex).strmod = frmItemEditor.STRMODER.value
        Item(EditorIndex).defmod = frmItemEditor.DEFMODER.value
        Item(EditorIndex).magimod = frmItemEditor.MAGIMODER.value
        If frmItemEditor.cmbType.ListIndex >= ITEM_TYPE_WEAPON Then Item(EditorIndex).Data3 = Val(frmItemEditor.lblsound.Caption)
        
    End If
    
    If (frmItemEditor.cmbType.ListIndex >= ITEM_TYPE_POTIONADDHP) And (frmItemEditor.cmbType.ListIndex <= ITEM_TYPE_POTIONSUBSP) Then
        Item(EditorIndex).Data1 = frmItemEditor.scrlVitalMod.value
        Item(EditorIndex).Data2 = 0
        Item(EditorIndex).Data3 = 0
    End If
    
    If (frmItemEditor.cmbType.ListIndex >= ITEM_TYPE_RANDITEM) Then
        Item(EditorIndex).Data1 = frmItemEditor.HScroll2.value
        Item(EditorIndex).Data2 = frmItemEditor.HScroll3.value
        Item(EditorIndex).Data3 = frmItemEditor.HScroll4.value
    End If
    
    If (frmItemEditor.cmbType.ListIndex = ITEM_TYPE_SPELL) Then
        Item(EditorIndex).Data1 = frmItemEditor.scrlSpell.value
        Item(EditorIndex).Data2 = 0
        Item(EditorIndex).Data3 = 0
    End If
    
    If (frmItemEditor.cmbType.ListIndex = ITEM_TYPE_MASKEDIT) Then
        Item(EditorIndex).Data1 = Val(frmItemEditor.Text2.Text)
        Item(EditorIndex).Data2 = frmItemEditor.fringemask.value
        Item(EditorIndex).Data3 = 0
    End If
    
    If (frmItemEditor.cmbType.ListIndex = ITEM_TYPE_SCHANGE) Or (frmItemEditor.cmbType.ListIndex = ITEM_TYPE_BARBER) Then
        Item(EditorIndex).Data1 = frmItemEditor.HScroll1.value
        Item(EditorIndex).Data2 = 0
        Item(EditorIndex).Data3 = 0
    End If
    If (frmItemEditor.cmbType.ListIndex = ITEM_TYPE_BOOK) Then
        Item(EditorIndex).Data1 = 0
        Item(EditorIndex).Data2 = frmItemEditor.Text1.Text
        Item(EditorIndex).Data3 = 0

    End If
    If (frmItemEditor.cmbType.ListIndex = ITEM_TYPE_PETEGG) Then
        Item(EditorIndex).Data1 = frmItemEditor.HScroll6.value
        Item(EditorIndex).Data2 = 0
        Item(EditorIndex).Data3 = 0

    End If
    Item(EditorIndex).SPRITE = frmItemEditor.Text3.Text

   
    Call SendSaveItem(EditorIndex)
    InItemsEditor = False
    Unload frmItemEditor
End Sub

Public Sub ItemEditorCancel()
    InItemsEditor = False
    Unload frmItemEditor
End Sub

Public Sub ItemEditorBltItem()
    Call BitBlt(frmItemEditor.picPic.hdc, 0, 0, PIC_X, PIC_Y, frmItemEditor.picItems.hdc, 0, frmItemEditor.scrlPic.value * PIC_Y, SRCCOPY)
    Call BitBlt(frmItemEditor.Picpic2.hdc, 0, 0, 3 * PIC_X, PIC_Y, frmItemEditor.picItems2.hdc, 0, Val(frmItemEditor.Text3.Text) * PIC_Y, SRCCOPY)

End Sub
Public Sub WeaponBltItem()
tEmp = 0
If GetPlayerWeaponSlot(MyIndex) > 0 Then
tEmp = Item(GetPlayerInvItemNum(MyIndex, (GetPlayerWeaponSlot(MyIndex)))).Pic
End If
    Call BitBlt(frmMirage.picPic.hdc, 0, 0, PIC_X, PIC_Y, frmMirage.picItems.hdc, 0, (tEmp) * PIC_Y, SRCCOPY)

End Sub

Public Sub BlitInventoryX()











End Sub
Public Sub ITEMBltItem()
tEmp = 0
If frmMirage.ITEM1.Tag > 0 Then
tEmp = frmMirage.ITEM1.Tag
End If
    Call BitBlt(frmMirage.ITEM1.hdc, 0, 0, PIC_X, PIC_Y, frmMirage.picItems.hdc, 0, (tEmp) * PIC_Y, SRCCOPY)
tEmp = 0
If frmMirage.ITEM2.Tag > 0 Then
tEmp = frmMirage.ITEM2.Tag
End If
    Call BitBlt(frmMirage.ITEM2.hdc, 0, 0, PIC_X, PIC_Y, frmMirage.picItems.hdc, 0, (tEmp) * PIC_Y, SRCCOPY)

tEmp = 0
If frmMirage.ITEM3.Tag > 0 Then
tEmp = frmMirage.ITEM3.Tag
End If
    Call BitBlt(frmMirage.ITEM3.hdc, 0, 0, PIC_X, PIC_Y, frmMirage.picItems.hdc, 0, (tEmp) * PIC_Y, SRCCOPY)

End Sub
Public Sub HelmBltItem()
tEmp = 0
If GetPlayerShieldSlot(MyIndex) > 0 Then
tEmp = Item(GetPlayerInvItemNum(MyIndex, (GetPlayerShieldSlot(MyIndex)))).Pic
End If
    Call BitBlt(frmMirage.picPic3.hdc, 0, 0, PIC_X, PIC_Y, frmMirage.picItems.hdc, 0, (tEmp) * PIC_Y, SRCCOPY)

End Sub

Public Sub ArmorBltItem()
tEmp = 0
If GetPlayerArmorSlot(MyIndex) > 0 Then
tEmp = Item(GetPlayerInvItemNum(MyIndex, (GetPlayerArmorSlot(MyIndex)))).Pic
End If
    Call BitBlt(frmMirage.Picpic2.hdc, 0, 0, PIC_X, PIC_Y, frmMirage.picItems.hdc, 0, (tEmp) * PIC_Y, SRCCOPY)
End Sub
Public Sub SetEquipShizzle()
'armor name dur
If GetPlayerArmorSlot(MyIndex) > 0 Then
TempStr = Item(GetPlayerInvItemNum(MyIndex, (GetPlayerArmorSlot(MyIndex)))).name
frmMirage.Label6.Caption = TempStr
'frmMirage.Label8.Caption = Item(GetPlayerInvItemNum(MyIndex, (GetPlayerArmorSlot(MyIndex)))).Data2
Else
frmMirage.Label6.Caption = "{Empty}"
End If

'weapon name dur
If GetPlayerWeaponSlot(MyIndex) > 0 Then
TempStr = Item(GetPlayerInvItemNum(MyIndex, (GetPlayerWeaponSlot(MyIndex)))).name
frmMirage.Label7.Caption = TempStr
Else
frmMirage.Label7.Caption = "{Empty}"
End If

'Acc. name dur
If GetPlayerShieldSlot(MyIndex) > 0 Then
TempStr = Item(GetPlayerInvItemNum(MyIndex, (GetPlayerShieldSlot(MyIndex)))).name
frmMirage.Label5.Caption = TempStr
Else
frmMirage.Label5.Caption = "{Empty}"
End If

End Sub

Public Sub NpcEditorInit()
On Error Resume Next
    
    frmNpcEditor.picSprites.Picture = LoadPicture(App.Path & "\sprites.bmp")
    
    frmNpcEditor.txtName.Text = Trim(Npc(EditorIndex).name)
    frmNpcEditor.txtAttackSay.Text = Trim(Npc(EditorIndex).AttackSay)
    frmNpcEditor.scrlSprite.value = Npc(EditorIndex).SPRITE
    frmNpcEditor.txtSpawnSecs.Text = STR(Npc(EditorIndex).SpawnSecs)
    frmNpcEditor.cmbBehavior.ListIndex = Npc(EditorIndex).Behavior
    frmNpcEditor.scrlRange.value = Npc(EditorIndex).range
    frmNpcEditor.txtChance.Text = STR(Npc(EditorIndex).DropChance)
    frmNpcEditor.scrlNum.value = Npc(EditorIndex).DropItem
    frmNpcEditor.scrlValue.value = Npc(EditorIndex).DropItemValue
    frmNpcEditor.scrlSTR.value = Npc(EditorIndex).STR
    frmNpcEditor.scrlDEF.value = Npc(EditorIndex).DEF
    frmNpcEditor.scrlSPEED.value = Npc(EditorIndex).speed
    frmNpcEditor.scrlMAGI.value = Npc(EditorIndex).MAGI
    
    frmNpcEditor.HScroll3.value = Npc(EditorIndex).SPRITE2
    frmNpcEditor.HScroll4.value = Npc(EditorIndex).SPRITE3
    frmNpcEditor.HScroll5.value = Npc(EditorIndex).SPRITE4
    
    frmNpcEditor.Text1.Text = STR(Npc(EditorIndex).DropChance2)
    frmNpcEditor.HScroll2.value = Npc(EditorIndex).DropItem2
    frmNpcEditor.HScroll1.value = Npc(EditorIndex).DropItemValue2
    
    frmNpcEditor.Show vbModal
End Sub

Public Sub NpcEditorOk()
    Npc(EditorIndex).name = frmNpcEditor.txtName.Text
    Npc(EditorIndex).AttackSay = frmNpcEditor.txtAttackSay.Text
    Npc(EditorIndex).SPRITE = frmNpcEditor.scrlSprite.value
    Npc(EditorIndex).SPRITE2 = frmNpcEditor.HScroll3.value
    Npc(EditorIndex).SPRITE3 = frmNpcEditor.HScroll4.value
    Npc(EditorIndex).SPRITE4 = frmNpcEditor.HScroll5.value
    Npc(EditorIndex).SpawnSecs = Val(frmNpcEditor.txtSpawnSecs.Text)
    Npc(EditorIndex).Behavior = frmNpcEditor.cmbBehavior.ListIndex
    Npc(EditorIndex).range = frmNpcEditor.scrlRange.value
    Npc(EditorIndex).DropChance = Val(frmNpcEditor.txtChance.Text)
    Npc(EditorIndex).DropItem = frmNpcEditor.scrlNum.value
    Npc(EditorIndex).DropItemValue = frmNpcEditor.scrlValue.value
    
    Npc(EditorIndex).DropChance2 = Val(frmNpcEditor.Text1.Text)
    Npc(EditorIndex).DropItem2 = frmNpcEditor.HScroll2.value
    Npc(EditorIndex).DropItemValue2 = frmNpcEditor.HScroll1.value
    
    Npc(EditorIndex).STR = frmNpcEditor.scrlSTR.value
    Npc(EditorIndex).DEF = frmNpcEditor.scrlDEF.value
    Npc(EditorIndex).speed = frmNpcEditor.scrlSPEED.value
    Npc(EditorIndex).MAGI = frmNpcEditor.scrlMAGI.value
    
    Call SendSaveNpc(EditorIndex)
    InNpcEditor = False
    Unload frmNpcEditor
End Sub

Public Sub NpcEditorCancel()
    InNpcEditor = False
    Unload frmNpcEditor
End Sub

Public Sub NpcEditorBltSprite()
    Call BitBlt(frmNpcEditor.picSprite.hdc, 0, 0, PIC_X, PIC_Y, frmNpcEditor.picSprites.hdc, 3 * PIC_X, frmNpcEditor.scrlSprite.value * PIC_Y, SRCCOPY)
    Call BitBlt(frmNpcEditor.Picture1.hdc, 0, 0, PIC_X, PIC_Y, frmNpcEditor.picSprites.hdc, 3 * PIC_X, frmNpcEditor.HScroll3.value * PIC_Y, SRCCOPY)
    Call BitBlt(frmNpcEditor.Picture2.hdc, 0, 0, PIC_X, PIC_Y, frmNpcEditor.picSprites.hdc, 3 * PIC_X, frmNpcEditor.HScroll4.value * PIC_Y, SRCCOPY)
    Call BitBlt(frmNpcEditor.Picture3.hdc, 0, 0, PIC_X, PIC_Y, frmNpcEditor.picSprites.hdc, 3 * PIC_X, frmNpcEditor.HScroll5.value * PIC_Y, SRCCOPY)


End Sub

Public Sub ShopEditorInit()
On Error Resume Next

Dim i As Long

    frmShopEditor.txtName.Text = Trim(Shop(EditorIndex).name)
    frmShopEditor.txtJoinSay.Text = Trim(Shop(EditorIndex).JoinSay)
    frmShopEditor.txtLeaveSay.Text = Trim(Shop(EditorIndex).LeaveSay)
    frmShopEditor.chkFixesItems.value = Shop(EditorIndex).FixesItems
    frmShopEditor.Check1.value = Shop(EditorIndex).OneSale
    
    frmShopEditor.cmbItemGive.Clear
    frmShopEditor.cmbItemGive.AddItem "None"
    frmShopEditor.cmbItemGet.Clear
    frmShopEditor.cmbItemGet.AddItem "None"
    For i = 1 To MAX_ITEMS
        frmShopEditor.cmbItemGive.AddItem i & ": " & Trim(Item(i).name)
        frmShopEditor.cmbItemGet.AddItem i & ": " & Trim(Item(i).name)
    Next i
    frmShopEditor.cmbItemGive.ListIndex = 0
    frmShopEditor.cmbItemGet.ListIndex = 0
    
    Call UpdateShopTrade
    
    frmShopEditor.Show vbModal
End Sub

Public Sub UpdateShopTrade()
Dim i As Long, GetItem As Long, GetValue As Long, GiveItem As Long, GiveValue As Long
    
    frmShopEditor.lstTradeItem.Clear
    For i = 1 To MAX_TRADES
        GetItem = Shop(EditorIndex).TradeItem(i).GetItem
        GetValue = Shop(EditorIndex).TradeItem(i).GetValue
        GiveItem = Shop(EditorIndex).TradeItem(i).GiveItem
        GiveValue = Shop(EditorIndex).TradeItem(i).GiveValue
        
        If GetItem > 0 And GiveItem > 0 Then
            frmShopEditor.lstTradeItem.AddItem i & ": " & GiveValue & " " & Trim(Item(GiveItem).name) & " for " & GetValue & " " & Trim(Item(GetItem).name)
        Else
            frmShopEditor.lstTradeItem.AddItem "Empty Trade Slot"
        End If
    Next i
    frmShopEditor.lstTradeItem.ListIndex = 0
End Sub

Public Sub ShopEditorOk()
    Shop(EditorIndex).name = frmShopEditor.txtName.Text
    Shop(EditorIndex).JoinSay = frmShopEditor.txtJoinSay.Text
    Shop(EditorIndex).LeaveSay = frmShopEditor.txtLeaveSay.Text
    Shop(EditorIndex).FixesItems = frmShopEditor.chkFixesItems.value
    Shop(EditorIndex).OneSale = frmShopEditor.Check1.value
    
    Call SendSaveShop(EditorIndex)
    InShopEditor = False
    Unload frmShopEditor
End Sub

Public Sub ShopEditorCancel()
    InShopEditor = False
    Unload frmShopEditor
End Sub

Public Sub SpellEditorInit()
On Error Resume Next

Dim i As Long

    frmSpellEditor.cmbClassReq.AddItem "All Classes"
    For i = 0 To Max_Classes
        frmSpellEditor.cmbClassReq.AddItem Trim(Class(i).name)
    Next i
    
    frmSpellEditor.txtName.Text = Trim(Spell(EditorIndex).name)
    frmSpellEditor.cmbClassReq.ListIndex = Spell(EditorIndex).ClassReq
    frmSpellEditor.scrlLevelReq.value = Spell(EditorIndex).LevelReq
    frmSpellEditor.ScrMPUSED.value = Spell(EditorIndex).MPused
    frmSpellEditor.Text1.Text = Spell(EditorIndex).Sfx
    frmSpellEditor.Text2.Text = Spell(EditorIndex).GFX
        
    frmSpellEditor.cmbType.ListIndex = Spell(EditorIndex).Type
    If Spell(EditorIndex).Type <> SPELL_TYPE_GIVEITEM Then
        frmSpellEditor.fraVitals.Visible = True
        frmSpellEditor.fraGiveItem.Visible = False
        frmSpellEditor.scrlVitalMod.value = Spell(EditorIndex).Data1
    Else
        frmSpellEditor.fraVitals.Visible = False
        frmSpellEditor.fraGiveItem.Visible = True
        frmSpellEditor.scrlItemNum.value = Spell(EditorIndex).Data1
        frmSpellEditor.scrlItemValue.value = Spell(EditorIndex).Data2
    End If
    If Spell(EditorIndex).Type = SPELL_TYPE_CRAFT Then
    frmSpellEditor.warpframe.Visible = True
    frmSpellEditor.HScroll1.value = Spell(EditorIndex).Data1
    frmSpellEditor.HScroll2.value = Spell(EditorIndex).Data2
    frmSpellEditor.HScroll3.value = Spell(EditorIndex).Data3
        Else
    frmSpellEditor.warpframe.Visible = False
    End If
    
        
    frmSpellEditor.Show vbModal
End Sub

Public Sub SpellEditorOk()
    Spell(EditorIndex).name = frmSpellEditor.txtName.Text
    Spell(EditorIndex).ClassReq = frmSpellEditor.cmbClassReq.ListIndex
    Spell(EditorIndex).LevelReq = frmSpellEditor.scrlLevelReq.value
    Spell(EditorIndex).Type = frmSpellEditor.cmbType.ListIndex
    Spell(EditorIndex).MPused = frmSpellEditor.ScrMPUSED.value
    Spell(EditorIndex).Sfx = frmSpellEditor.Text1.Text
    Spell(EditorIndex).GFX = Val(frmSpellEditor.Text2.Text)
    If Spell(EditorIndex).Type <> SPELL_TYPE_GIVEITEM Then
        Spell(EditorIndex).Data1 = frmSpellEditor.scrlVitalMod.value
    Else
        Spell(EditorIndex).Data1 = frmSpellEditor.scrlItemNum.value
        Spell(EditorIndex).Data2 = frmSpellEditor.scrlItemValue.value
        
    End If
    Spell(EditorIndex).Data3 = 0
    If Spell(EditorIndex).Type = SPELL_TYPE_CRAFT Then
    Spell(EditorIndex).Data1 = frmSpellEditor.HScroll1.value
    Spell(EditorIndex).Data2 = frmSpellEditor.HScroll2.value
    Spell(EditorIndex).Data3 = frmSpellEditor.HScroll3.value
    End If
    
    Call SendSaveSpell(EditorIndex)
    InSpellEditor = False
    Unload frmSpellEditor
End Sub

Public Sub SpellEditorCancel()
    InSpellEditor = False
    Unload frmSpellEditor
End Sub

Sub BltMapNPCName(ByVal index As Long)
    Dim TextX As Long
    Dim TextY As Long
    
    With Npc(MapNpc(index).num)
        'Draw name
        TextX = MapNpc(index).X * PIC_X + MapNpc(index).XOffset + CLng(PIC_X / 2) - ((Len(Trim$(.name)) / 2) * 8)
        TextY = MapNpc(index).y * PIC_Y + MapNpc(index).YOffset - CLng(PIC_Y / 2) - 4
        DrawText TexthDC, TextX, TextY, Trim$(.name), vbWhite
    End With
End Sub

Public Sub UpdateInventory()
Dim i As Long

    frmInventory.LSTINV.Clear
    
    ' Show the inventory
    For i = 1 To MAX_INV
        If GetPlayerInvItemNum(MyIndex, i) > 0 And GetPlayerInvItemNum(MyIndex, i) <= MAX_ITEMS Then
            If Item(GetPlayerInvItemNum(MyIndex, i)).Type = ITEM_TYPE_CURRENCY Then
                frmInventory.LSTINV.AddItem i & ": " & Trim(Item(GetPlayerInvItemNum(MyIndex, i)).name) & " (" & GetPlayerInvItemValue(MyIndex, i) & ")"
            Else
                ' Check if this item is being worn
                If GetPlayerWeaponSlot(MyIndex) = i Or GetPlayerArmorSlot(MyIndex) = i Or GetPlayerHelmetSlot(MyIndex) = i Or GetPlayerShieldSlot(MyIndex) = i Then
                    frmInventory.LSTINV.AddItem i & ": " & Trim(Item(GetPlayerInvItemNum(MyIndex, i)).name) & " (worn)"
                Else
                    frmInventory.LSTINV.AddItem i & ": " & Trim(Item(GetPlayerInvItemNum(MyIndex, i)).name)
                End If
            End If
        Else
            frmInventory.LSTINV.AddItem "<free inventory slot>"
        End If
    Next i
    
    frmInventory.LSTINV.ListIndex = 0
End Sub
Public Sub UpdateInventory2()
Dim i As Long

    frmBank.lstInventory.Clear
    
    ' Show the inventory
    For i = 1 To MAX_INV
        If GetPlayerInvItemNum(MyIndex, i) > 0 And GetPlayerInvItemNum(MyIndex, i) <= MAX_ITEMS Then
            If Item(GetPlayerInvItemNum(MyIndex, i)).Type = ITEM_TYPE_CURRENCY Then
               frmBank.lstInventory.AddItem i & ": " & Trim(Item(GetPlayerInvItemNum(MyIndex, i)).name) & " (" & GetPlayerInvItemValue(MyIndex, i) & ")"
            Else
                ' Check if this item is being worn
                If GetPlayerWeaponSlot(MyIndex) = i Or GetPlayerArmorSlot(MyIndex) = i Or GetPlayerHelmetSlot(MyIndex) = i Or GetPlayerShieldSlot(MyIndex) = i Then
                    frmBank.lstInventory.AddItem i & ": " & Trim(Item(GetPlayerInvItemNum(MyIndex, i)).name) & " (worn)"
                Else
                    frmBank.lstInventory.AddItem i & ": " & Trim(Item(GetPlayerInvItemNum(MyIndex, i)).name)
                End If
            End If
        Else
            frmBank.lstInventory.AddItem "<Empty>"
        End If
    Next i
    
    frmBank.lstInventory.ListIndex = 0
End Sub
Public Sub UpdateBank()
Dim i As Long

    frmBank.lstBank.Clear
    
    ' Show the inventory
    For i = 1 To MAX_BANK
        If GetPlayerBankItemNum(MyIndex, i) > 0 And GetPlayerBankItemNum(MyIndex, i) <= MAX_ITEMS Then
            If Item(GetPlayerBankItemNum(MyIndex, i)).Type = ITEM_TYPE_CURRENCY Then
                frmBank.lstBank.AddItem i & ": " & Trim(Item(GetPlayerBankItemNum(MyIndex, i)).name) & " (" & GetPlayerbankItemValue(MyIndex, i) & ")"
            Else
                ' Check if this item is being worn
                
                    frmBank.lstBank.AddItem i & ": " & Trim(Item(GetPlayerBankItemNum(MyIndex, i)).name)
            End If
        Else
            frmBank.lstBank.AddItem ""
        End If
    Next i
    
    frmBank.lstBank.ListIndex = 0
End Sub

Sub ResizeGUI()
    If frmMirage.WindowState <> vbMinimized Then
        'frmMirage.txtChat.Height = Int(frmMirage.Height / Screen.TwipsPerPixelY) - frmMirage.txtChat.top - 32
        'frmMirage.txtChat.Width = Int((frmMirage.Width / Screen.TwipsPerPixelX) - 8) - 92
    End If
End Sub

Sub PlayerSearch(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim x1 As Long, y1 As Long

    x1 = Int(X / PIC_X)
    y1 = Int(y / PIC_Y)
    
    If (x1 >= 0) And (x1 <= MAX_MAPX) And (y1 >= 0) And (y1 <= MAX_MAPY) Then
        Call SendData("search" & SEP_CHAR & x1 & SEP_CHAR & y1 & SEP_CHAR & END_CHAR)
    End If
End Sub

Sub quicksl()
tEmp = frmInventory.LSTINV.ListIndex + SELectINV + 1 + (frmInventory.invbar.value)
If (GetPlayerInvItemNum(MyIndex, tEmp) > 0) And (GetPlayerInvItemNum(MyIndex, tEmp) <= MAX_ITEMS) Then
tEmp = GetPlayerInvItemNum(MyIndex, tEmp)
If tEmp = 0 Then Exit Sub
tEmp = frmInventory.LSTINV.ListIndex + SELectINV + (frmInventory.invbar.value)
TEMPITEMNUM = Item(GetPlayerInvItemNum(MyIndex, tEmp + 1)).Pic
TEMPINVSLOT = tEmp
Form2.Label1.Caption = Trim(Item(GetPlayerInvItemNum(MyIndex, tEmp + 1)).name)
Form2.Show vbModal
Else

End If
End Sub
