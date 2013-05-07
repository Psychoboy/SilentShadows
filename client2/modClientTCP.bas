Attribute VB_Name = "modClientTCP"
Option Explicit

Public ServerIP As String
Public PlayerBuffer As String
Public InGame As Boolean

Sub TcpInit()
    SEP_CHAR = Chr(0)
    END_CHAR = Chr(237)
    PlayerBuffer = ""
        
 frmMirage.Socket.RemoteHost = GAME_IP
' frmMirage.Socket.RemoteHost = frmMirage.Socket.LocalIP
 frmMirage.Socket.RemotePort = GAME_PORT
End Sub

Sub TcpDestroy()
    frmMirage.Socket.Close
    
    If frmChars.Visible Then frmChars.Visible = False
    If frmCredits.Visible Then frmCredits.Visible = False
    If frmDeleteAccount.Visible Then frmDeleteAccount.Visible = False
    If frmLogin.Visible Then frmLogin.Visible = False
    If frmNewAccount.Visible Then frmNewAccount.Visible = False
    If frmNewChar.Visible Then frmNewChar.Visible = False
End Sub

Sub IncomingData(ByVal DataLength As Long)
Dim Buffer As String
Dim Packet As String
Dim top As String * 3
Dim Start As Integer

    frmMirage.Socket.GetData Buffer, vbString, DataLength
    PlayerBuffer = PlayerBuffer & Buffer
        
    Start = InStr(PlayerBuffer, END_CHAR)
    Do While Start > 0
        Packet = Mid(PlayerBuffer, 1, Start - 1)
        PlayerBuffer = Mid(PlayerBuffer, Start + 1, Len(PlayerBuffer))
        Start = InStr(PlayerBuffer, END_CHAR)
        If Len(Packet) > 0 Then
            Call HandleData(Packet)
        End If
    Loop
End Sub

Sub HandleData(ByVal data As String)
Dim Parse() As String
Dim name As String
Dim Password As String
Dim Sex As Long
Dim ClassNum As Long
Dim CharNum As Long
Dim Msg As String
Dim IPMask As String
Dim BanSlot As Long
Dim MsgTo As Long
Dim Dir As Long
Dim InvNum As Long
Dim Ammount As Long
Dim Damage As Long
Dim PointType As Long
Dim BanPlayer As Long
Dim Level As Long
Dim i As Long, n As Long, X As Long, y As Long
Dim ShopNum As Long, GiveItem As Long, GiveValue As Long, GetItem As Long, GetValue As Long

    ' Handle Data
    Parse = Split(data, SEP_CHAR)
    
    ' Add the data to the debug window if we are in debug mode
    If Trim(Command) = "-debug" Then
        If frmDebug.Visible = False Then frmDebug.Visible = True
        Call TextAdd(frmDebug.txtDebug, "((( Processed Packet " & Parse(0) & " )))", True)
    End If
    
    ' ::::::::::::::::::::::::::
    ' :: Alert message packet ::
    ' ::::::::::::::::::::::::::
    If LCase(Parse(0)) = "alertmsg" Then
        frmSendGetData.Visible = False
        frmMainMenu.Visible = True
        
        Msg = Parse(1)
        Call MsgBox(Msg, vbOKOnly, GAME_NAME)
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::
    ' :: muteme message packet ::
    ' :::::::::::::::::::::::::::
    If LCase(Parse(0)) = "mute" Then
    GLBLCHT = 0
    Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::
    ' :: unmuteme message packet ::
    ' :::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "unmute" Then
    GLBLCHT = 1
    Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::
    ' :: All characters packet ::
    ' :::::::::::::::::::::::::::
    If LCase(Parse(0)) = "allchars" Then
        n = 1
        
        frmChars.Visible = True
        frmSendGetData.Visible = False
        
        frmChars.lstChars.Clear
        
        For i = 1 To MAX_CHARS
            name = Parse(n)
            Msg = Parse(n + 1)
            Level = Val(Parse(n + 3))
            
            If Trim(name) = "" Then
                frmChars.lstChars.AddItem "Free Character Slot"
            Else
                frmChars.lstChars.AddItem name & " a level " & Level & " " & Parse(n + 2) & " " & Msg
            End If
            
            n = n + 4
        Next i
        
        frmChars.lstChars.ListIndex = 0
        Exit Sub
    End If

    ' :::::::::::::::::::::::::::::::::
    ' :: Login was successful packet ::
    ' :::::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "loginok" Then
        ' Now we can receive game data
        MyIndex = Val(Parse(1))
        
        frmSendGetData.Visible = True
        frmChars.Visible = False
        
        Call SetStatus("Receiving game data...")
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::::::::::::
    ' :: New character classes data packet ::
    ' :::::::::::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "newcharclasses" Then
        n = 1
        
        ' Max classes
        Max_Classes = Val(Parse(n))
        ReDim Class(0 To Max_Classes) As ClassRec
        
        n = n + 1
        
        For i = 0 To Max_Classes
            Class(i).name = Parse(n)
            
            Class(i).HP = Val(Parse(n + 1))
            Class(i).MP = Val(Parse(n + 2))
            Class(i).SP = Val(Parse(n + 3))
            
            Class(i).STR = Val(Parse(n + 4))
            Class(i).DEF = Val(Parse(n + 5))
            Class(i).speed = Val(Parse(n + 6))
            Class(i).MAGI = Val(Parse(n + 7))
            
            n = n + 8
        Next i
        
        ' Used for if the player is creating a new character
        

        frmNewChar.cmbClass.Clear

        For i = 0 To Max_Classes
            frmNewChar.cmbClass.AddItem Trim(Class(i).name)
        Next i
            
        frmNewChar.cmbClass.ListIndex = 0
        frmNewChar.lblHP.Caption = STR(Class(0).HP)
        frmNewChar.lblMP.Caption = STR(Class(0).MP)
        frmNewChar.lblSP.Caption = STR(Class(0).SP)

        frmNewChar.lblSTR.Caption = STR(Class(0).STR)
        frmNewChar.lblDEF.Caption = STR(Class(0).DEF)
        frmNewChar.lblSPEED.Caption = STR(Class(0).speed)
        frmNewChar.lblMAGI.Caption = STR(Class(0).MAGI)
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::
    ' :: Classes data packet ::
    ' :::::::::::::::::::::::::
    If LCase(Parse(0)) = "classesdata" Then
        n = 1
        
        ' Max classes
        Max_Classes = Val(Parse(n))
        ReDim Class(0 To Max_Classes) As ClassRec
        
        n = n + 1
        
        For i = 0 To Max_Classes
            Class(i).name = Parse(n)
            
            Class(i).HP = Val(Parse(n + 1))
            Class(i).MP = Val(Parse(n + 2))
            Class(i).SP = Val(Parse(n + 3))
            
            Class(i).STR = Val(Parse(n + 4))
            Class(i).DEF = Val(Parse(n + 5))
            Class(i).speed = Val(Parse(n + 6))
            Class(i).MAGI = Val(Parse(n + 7))
            
            n = n + 8
        Next i
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::::::::::::
    ' :: New character races data packet ::
    ' :::::::::::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "newcharraces" Then
        n = 1
        
        ' Max races
        Max_Races = Val(Parse(n))
        ReDim Race(0 To Max_Races) As RaceRec
        
        n = n + 1
        
        For i = 0 To Max_Races
            Race(i).name = Parse(n)
            
            Race(i).HP = Val(Parse(n + 1))
            Race(i).MP = Val(Parse(n + 2))
            Race(i).SP = Val(Parse(n + 3))
            
            Race(i).STR = Val(Parse(n + 4))
            Race(i).DEF = Val(Parse(n + 5))
            Race(i).speed = Val(Parse(n + 6))
            Race(i).MAGI = Val(Parse(n + 7))
            
            n = n + 8
        Next i
        
        ' Used for if the player is creating a new character
              frmNewChar.cmbRace.Clear
        For i = 0 To Max_Races
            frmNewChar.cmbRace.AddItem Trim(Race(i).name)
        Next i
         
        frmNewChar.cmbRace.ListIndex = 0
        frmNewChar.lblHP.Caption = STR(Val(frmNewChar.lblHP.Caption) + Race(0).HP)
        frmNewChar.lblMP.Caption = STR(Val(frmNewChar.lblMP.Caption) + Race(0).MP)
        frmNewChar.lblSP.Caption = STR(Val(frmNewChar.lblSP.Caption) + Race(0).SP)

        frmNewChar.lblSTR.Caption = STR(Val(frmNewChar.lblSTR.Caption) + Race(0).STR)
        frmNewChar.lblDEF.Caption = STR(Val(frmNewChar.lblDEF.Caption) + Race(0).DEF)
        frmNewChar.lblSPEED.Caption = STR(Val(frmNewChar.lblSPEED.Caption) + Race(0).speed)
        frmNewChar.lblMAGI.Caption = STR(Val(frmNewChar.lblMAGI.Caption) + Race(0).MAGI)
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::
    ' :: Races data packet ::
    ' :::::::::::::::::::::::::
    If LCase(Parse(0)) = "racesdata" Then
        n = 1
        
        ' Max races
        Max_Races = Val(Parse(n))
        ReDim Race(0 To Max_Races) As RaceRec
        
        n = n + 1
        
        For i = 0 To Max_Races
            Race(i).name = Parse(n)
            
            Race(i).HP = Val(Parse(n + 1))
            Race(i).MP = Val(Parse(n + 2))
            Race(i).SP = Val(Parse(n + 3))
            
            Race(i).STR = Val(Parse(n + 4))
            Race(i).DEF = Val(Parse(n + 5))
            Race(i).speed = Val(Parse(n + 6))
            Race(i).MAGI = Val(Parse(n + 7))
            
            n = n + 8
        Next i
        Exit Sub
    End If
    
    ' ::::::::::::::::::::
    ' :: In game packet ::
    ' ::::::::::::::::::::
    If LCase(Parse(0)) = "ingame" Then
        InGame = True
        Call GameInit
        Call GameLoop
        If Parse(1) = END_CHAR Then
            MsgBox ("here")
            End
        End If
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::
    ' :: Player inventory packet ::
    ' :::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "playerinv" Then
        n = 1
        For i = 1 To MAX_INV
            Call SetPlayerInvItemNum(MyIndex, i, Val(Parse(n)))
            Call SetPlayerInvItemValue(MyIndex, i, Val(Parse(n + 1)))
            Call SetPlayerInvItemDur(MyIndex, i, Val(Parse(n + 2)))
            
            n = n + 3
        Next i
        Call UpdateInventory
        Exit Sub
    End If
    
    'Bank Stuff
    If LCase(Parse(0)) = "playerbank" Then
     n = 1
        For i = 1 To MAX_INV
            Call SetPlayerbankItemNum(MyIndex, i, Val(Parse(n)))
            Call SetPlayerBankItemValue(MyIndex, i, Val(Parse(n + 1)))
            Call SetPlayerbankItemDur(MyIndex, i, Val(Parse(n + 2)))
            
            n = n + 3
        Next i
        Call UpdateInventory2
        Call UpdateBank
        frmBank.Show
        Exit Sub
    End If
        
    
    ' ::::::::::::::::::::::::::::::::::::
    ' :: Player inventory update packet ::
    ' ::::::::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "playerinvupdate" Then
        n = Val(Parse(1))
        
        Call SetPlayerInvItemNum(MyIndex, n, Val(Parse(2)))
        Call SetPlayerInvItemValue(MyIndex, n, Val(Parse(3)))
        Call SetPlayerInvItemDur(MyIndex, n, Val(Parse(4)))
        Call UpdateInventory
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::::::::::
    ' :: Player worn equipment packet ::
    ' ::::::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "playerworneq" Then
        Call SetPlayerArmorSlot(MyIndex, Val(Parse(1)))
        Call SetPlayerWeaponSlot(MyIndex, Val(Parse(2)))
        Call SetPlayerHelmetSlot(MyIndex, Val(Parse(3)))
        Call SetPlayerShieldSlot(MyIndex, Val(Parse(4)))
        Call UpdateInventory
        Exit Sub
    End If
        
    ' ::::::::::::::::::::::
    ' :: admincmds packet ::
    ' ::::::::::::::::::::::
    If LCase(Parse(0)) = "admincmds" Then
        Player(MyIndex).AdminCmds = Parse(1)
      Exit Sub
    End If
        
    ' ::::::::::::::::::::::
    ' :: Player hp packet ::
    ' ::::::::::::::::::::::
    If LCase(Parse(0)) = "playerhp" Then
        Player(MyIndex).MaxHP = Val(Parse(1))
        Call SetPlayerHP(MyIndex, Val(Parse(2)))
        If GetPlayerMaxHP(MyIndex) > 0 Then
            frmMirage.lblHP.Caption = GetPlayerHP(MyIndex) & " / " & GetPlayerMaxHP(MyIndex)
            frmMirage.HPBAR.Width = Int(GetPlayerHP(MyIndex) / GetPlayerMaxHP(MyIndex) * 115)
        End If
        Exit Sub
    End If

    ' ::::::::::::::::::::::
    ' :: Player mp packet ::
    ' ::::::::::::::::::::::
    If LCase(Parse(0)) = "playermp" Then
        Player(MyIndex).MaxMP = Val(Parse(1))
        Call SetPlayerMP(MyIndex, Val(Parse(2)))
        If GetPlayerMaxMP(MyIndex) > 0 Then
            
            frmMirage.lblMP.Caption = GetPlayerMP(MyIndex) & " / " & GetPlayerMaxMP(MyIndex)
            frmMirage.MPBAR.Width = Int(GetPlayerMP(MyIndex) / GetPlayerMaxMP(MyIndex) * 115)
        End If
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::
    ' :: Player sp packet ::
    ' ::::::::::::::::::::::
    If LCase(Parse(0)) = "playersp" Then
        Player(MyIndex).MaxSP = Val(Parse(1))
        Call SetPlayerSP(MyIndex, Val(Parse(2)))
        If GetPlayerMaxSP(MyIndex) > 0 Then
        frmMirage.SPBAR.Width = Int(GetPlayerSP(MyIndex) / GetPlayerMaxSP(MyIndex) * 115)
            frmMirage.lblSP.Caption = GetPlayerSP(MyIndex) & " / " & GetPlayerMaxSP(MyIndex)
        End If
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::
    ' :: Player stats packet ::
    ' :::::::::::::::::::::::::
    If LCase(Parse(0)) = "playerstats" Then
        Call SetPlayerSTR(MyIndex, Val(Parse(1)))
        Call SetPlayerDEF(MyIndex, Val(Parse(2)))
        Call SetPlayerSPEED(MyIndex, Val(Parse(3)))
        Call SetPlayerMAGI(MyIndex, Val(Parse(4)))
        Exit Sub
    End If
                
    ' ::::::::::::::::::::::::
    ' :: Player data packet ::
    ' ::::::::::::::::::::::::
    If LCase(Parse(0)) = "playerdata" Then
        i = Val(Parse(1))
        
        Call SetPlayerName(i, Parse(2))
        Call SetPlayerSprite(i, Val(Parse(3)))
        Call SetPlayerSprite2(i, Val(Parse(4)))
        Call SetPlayerSprite3(i, Val(Parse(5)))
        Call SetPlayerSprite4(i, Val(Parse(6)))
        Call SetPlayerMap(i, Val(Parse(7)))
        Call SetPlayerX(i, Val(Parse(8)))
        Call SetPlayerY(i, Val(Parse(9)))
        Call SetPlayerDir(i, Val(Parse(10)))
        Call SetPlayerAccess(i, Val(Parse(11)))
        Call SetPlayerPK(i, Val(Parse(12)))
        If i = MyIndex Then GLBLCHT = Val(Parse(13))
        Call SetPet(i, Val(Parse(14)))
        Call SetPlayerGuild(i, Parse(15))
        Call SetPlayerAnonymous(i, Val(Parse(16)))
        Call SetPlayerFaction(i, Val(Parse(17)))
        
        
        ' Make sure they aren't walking
        Player(i).Moving = 0
        Player(i).XOffset = 0
        Player(i).YOffset = 0
        
        ' Check if the player is the client player, and if so reset directions
        If i = MyIndex Then
            DirUp = False
            DirDown = False
            DirLeft = False
            DirRight = False
        End If
        
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::::
    ' :: Player movement packet ::
    ' ::::::::::::::::::::::::::::
    If (LCase(Parse(0)) = "playermove") Then
        i = Val(Parse(1))
        X = Val(Parse(2))
        y = Val(Parse(3))
        Dir = Val(Parse(4))
        n = Val(Parse(5))

        Call SetPlayerX(i, X)
        Call SetPlayerY(i, y)
        Call SetPlayerDir(i, Dir)
                
        Player(i).XOffset = 0
        Player(i).YOffset = 0
        Player(i).Moving = n
        
        Select Case GetPlayerDir(i)
            Case DIR_UP
                Player(i).YOffset = PIC_Y
            Case DIR_DOWN
                Player(i).YOffset = PIC_Y * -1
            Case DIR_LEFT
                Player(i).XOffset = PIC_X
            Case DIR_RIGHT
                Player(i).XOffset = PIC_X * -1
        End Select
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::
    ' :: player kill log data ::
    ' ::::::::::::::::::::::::::
    If (LCase(Parse(0)) = "pkshit") Then
        FRMPKLIST.List1.AddItem Parse(1) & " killed " & Parse(2) & " at " & Parse(3)
    End If
    
    
        
        
    
    
    ' :::::::::::::::::::::::::
    ' :: Npc movement packet ::
    ' :::::::::::::::::::::::::
    If (LCase(Parse(0)) = "npcmove") Then
        i = Val(Parse(1))
        X = Val(Parse(2))
        y = Val(Parse(3))
        Dir = Val(Parse(4))
        n = Val(Parse(5))

        MapNpc(i).X = X
        MapNpc(i).y = y
        MapNpc(i).Dir = Dir
        MapNpc(i).XOffset = 0
        MapNpc(i).YOffset = 0
        MapNpc(i).Moving = n
        
        Select Case MapNpc(i).Dir
            Case DIR_UP
                MapNpc(i).YOffset = PIC_Y
            Case DIR_DOWN
                MapNpc(i).YOffset = PIC_Y * -1
            Case DIR_LEFT
                MapNpc(i).XOffset = PIC_X
            Case DIR_RIGHT
                MapNpc(i).XOffset = PIC_X * -1
        End Select
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::
    ' :: Player direction packet ::
    ' :::::::::::::::::::::::::::::
    If (LCase(Parse(0)) = "playerdir") Then
        i = Val(Parse(1))
        Dir = Val(Parse(2))
        Call SetPlayerDir(i, Dir)
        
        Player(i).XOffset = 0
        Player(i).YOffset = 0
        Player(i).Moving = 0
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::
    ' :: NPC direction packet ::
    ' ::::::::::::::::::::::::::
    If (LCase(Parse(0)) = "npcdir") Then
        i = Val(Parse(1))
        Dir = Val(Parse(2))
        MapNpc(i).Dir = Dir
        
        MapNpc(i).XOffset = 0
        MapNpc(i).YOffset = 0
        MapNpc(i).Moving = 0
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::::
    ' :: Player XY location packet ::
    ' :::::::::::::::::::::::::::::::
    If (LCase(Parse(0)) = "playerxy") Then
        X = Val(Parse(1))
        y = Val(Parse(2))
        
        Call SetPlayerX(MyIndex, X)
        Call SetPlayerY(MyIndex, y)
        
        ' Make sure they aren't walking
        Player(MyIndex).Moving = 0
        Player(MyIndex).XOffset = 0
        Player(MyIndex).YOffset = 0
        
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::
    ' :: Player attack packet ::
    ' ::::::::::::::::::::::::::
    If (LCase(Parse(0)) = "attack") Then
        i = Val(Parse(1))
        
        ' Set player to attacking
        Player(i).Attacking = 1
        Player(i).AttackTimer = GetTickCount
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::
    ' :: NPC attack packet ::
    ' :::::::::::::::::::::::
    If (LCase(Parse(0)) = "npcattack") Then
        i = Val(Parse(1))
        
        ' Set player to attacking
        MapNpc(i).Attacking = 1
        MapNpc(i).AttackTimer = GetTickCount
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::
    ' :: Check for map packet ::
    ' ::::::::::::::::::::::::::
    If (LCase(Parse(0)) = "checkformap") Then
        ' Erase all players except self
        For i = 1 To MAX_PLAYERS
            If i <> MyIndex Then
                Call SetPlayerMap(i, 0)
            End If
        Next i
        
        ' Erase all temporary tile values
        Call ClearTempTile

        ' Get map num
        X = Val(Parse(1))
        
        ' Get revision
        y = Val(Parse(2))
        
        If FileExist("maps\map" & X & ".dat") Then
            ' Check to see if the revisions match
            If GetMapRevision(X) = y Then
                ' We do so we dont need the map
                
                ' Load the map
                Call LoadMap(X)
                
                Call SendData("needmap" & SEP_CHAR & "no" & SEP_CHAR & END_CHAR)
                Exit Sub
            End If
        End If
        
        ' Either the revisions didn't match or we dont have the map, so we need it
        Call SendData("needmap" & SEP_CHAR & "yes" & SEP_CHAR & END_CHAR)
        Exit Sub
    End If
    
    ' :::::::::::::::::::::
    ' :: Map data packet ::
    ' :::::::::::::::::::::
    If LCase(Parse(0)) = "mapdata" Then
        n = 1
        
        SaveMap.name = Parse(n + 1)
        SaveMap.Revision = Val(Parse(n + 2))
        SaveMap.Moral = Val(Parse(n + 3))
        SaveMap.Up = Val(Parse(n + 4))
        SaveMap.Down = Val(Parse(n + 5))
        SaveMap.Left = Val(Parse(n + 6))
        SaveMap.Right = Val(Parse(n + 7))
        SaveMap.Music = Val(Parse(n + 8))
        SaveMap.BootMap = Val(Parse(n + 9))
        SaveMap.BootX = Val(Parse(n + 10))
        SaveMap.BootY = Val(Parse(n + 11))
        SaveMap.Shop = Val(Parse(n + 12))
        n = n + 13
        
        For y = 0 To MAX_MAPY
            For X = 0 To MAX_MAPX
                SaveMap.Tile(X, y).Ground = Val(Parse(n))
                SaveMap.Tile(X, y).Mask = Val(Parse(n + 1))
                SaveMap.Tile(X, y).Anim = Val(Parse(n + 2))
                SaveMap.Tile(X, y).Mask2 = Val(Parse(n + 3))
                SaveMap.Tile(X, y).M2Anim = Val(Parse(n + 4))
                SaveMap.Tile(X, y).Fringe = Val(Parse(n + 5))
                SaveMap.Tile(X, y).FAnim = Val(Parse(n + 6))
                SaveMap.Tile(X, y).Fringe2 = Val(Parse(n + 7))
                SaveMap.Tile(X, y).F2Anim = Val(Parse(n + 8))
                SaveMap.Tile(X, y).Type = Val(Parse(n + 9))
                SaveMap.Tile(X, y).Data1 = Parse(n + 10)
                SaveMap.Tile(X, y).Data2 = Parse(n + 11)
                SaveMap.Tile(X, y).Data3 = Parse(n + 12)
                
                n = n + 13
            Next X
        Next y
        
        For X = 1 To MAX_MAP_NPCS
            SaveMap.Npc(X) = Val(Parse(n))
            n = n + 1
        Next X
                
        ' Save the map
        Call SaveLocalMap(Val(Parse(1)))
        
        ' Check if we get a map from someone else and if we were editing a map cancel it out
        If InEditor Then
            InEditor = False
            frmMirage.picMapEditor.Visible = False
            
            If frmMapWarp.Visible Then
                Unload frmMapWarp
            End If
            
            If frmMapProperties.Visible Then
                Unload frmMapProperties
            End If
        End If
        
        Exit Sub
    End If
        
    ' :::::::::::::::::::::::::::
    ' :: Map items data packet ::
    ' :::::::::::::::::::::::::::
    If LCase(Parse(0)) = "mapitemdata" Then
        n = 1
        
        For i = 1 To MAX_MAP_ITEMS
            SaveMapItem(i).num = Val(Parse(n))
            SaveMapItem(i).value = Val(Parse(n + 1))
            SaveMapItem(i).Dur = Val(Parse(n + 2))
            SaveMapItem(i).X = Val(Parse(n + 3))
            SaveMapItem(i).y = Val(Parse(n + 4))
            
            n = n + 5
        Next i
        
        Exit Sub
    End If
        
    ' :::::::::::::::::::::::::
    ' :: Map npc data packet ::
    ' :::::::::::::::::::::::::
    If LCase(Parse(0)) = "mapnpcdata" Then
        n = 1
        
        For i = 1 To MAX_MAP_NPCS
            SaveMapNpc(i).num = Val(Parse(n))
            SaveMapNpc(i).X = Val(Parse(n + 1))
            SaveMapNpc(i).y = Val(Parse(n + 2))
            SaveMapNpc(i).Dir = Val(Parse(n + 3))
            
            n = n + 4
        Next i
        
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::::
    ' :: Map send completed packet ::
    ' :::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "mapdone" Then
        Map = SaveMap
        
        For i = 1 To MAX_MAP_ITEMS
            MapItem(i) = SaveMapItem(i)
        Next i
        
        For i = 1 To MAX_MAP_NPCS
            MapNpc(i) = SaveMapNpc(i)
        Next i
        
        GettingMap = False
        
        ' Play music
        'Call StopMidi
        
        If IHaveMusic = 1 Then
        If Map.Music > 0 Then
            If Map.Music <> MyMusiC Then
            Call StopMidi
            Call PlayMidi(App.Path & "\music\" & Trim(STR(Map.Music)) & ".ogg")
            MyMusiC = Trim(STR(Map.Music))
            End If
        Else
        Call StopMidi
        End If
        End If
        
        Exit Sub
    End If
    
  ' ::::::::::::::::::::
    ' :: Social packets ::
    ' ::::::::::::::::::::
    If (LCase(Parse(0)) = "saymsg") Or (LCase(Parse(0)) = "globalmsg") Or (LCase(Parse(0)) = "playermsg") Or (LCase(Parse(0)) = "mapmsg") Or (LCase(Parse(0)) = "adminmsg") Then
        Call AddText(Parse(1), Val(Parse(2)))
        Exit Sub
    End If
    
     If (LCase(Parse(0)) = "playermsgglobal") Or (LCase(Parse(0)) = "broadcastmsg") Or (LCase(Parse(0)) = "globalmsg") Or (LCase(Parse(0)) = "adminmsg") Then
    Call AddTextCombat(Parse(1), Val(Parse(2)))
        Exit Sub
    End If
    
    If (LCase(Parse(0)) = "playermsgcombat") Or (LCase(Parse(0)) = "globalmsgcombat") Then
    Call AddTextActions(Parse(1), Val(Parse(2)))
        Exit Sub
    End If
    ' :::::::::::::::::::::::
    ' :: Item spawn packet ::
    ' :::::::::::::::::::::::
    If LCase(Parse(0)) = "spawnitem" Then
        n = Val(Parse(1))
        
        MapItem(n).num = Val(Parse(2))
        MapItem(n).value = Val(Parse(3))
        MapItem(n).Dur = Val(Parse(4))
        MapItem(n).X = Val(Parse(5))
        MapItem(n).y = Val(Parse(6))
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::
    ' :: Item editor packet ::
    ' ::::::::::::::::::::::::
    If (LCase(Parse(0)) = "itemeditor") Then
        InItemsEditor = True
        
        frmIndex.Show
        frmIndex.lstIndex.Clear
        
        ' Add the names
        For i = 1 To MAX_ITEMS
            frmIndex.lstIndex.AddItem i & ": " & Trim(Item(i).name)
        Next i
        
        frmIndex.lstIndex.ListIndex = 0
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::
    ' :: Update item packet ::
    ' ::::::::::::::::::::::::
    If (LCase(Parse(0)) = "updateitem") Then
        n = Val(Parse(1))
        
        ' Update the item
        Item(n).name = Parse(2)
        Item(n).Pic = Val(Parse(3))
        Item(n).Type = Val(Parse(4))
        Item(n).Data1 = 0
        Item(n).Data2 = 0
        Item(n).Data3 = 0
        Item(n).Data4 = 0
        Item(n).SellValue = Val(Parse(5))
        
        Exit Sub
    End If
    
    ' ::::::::::::::::::::
    ' ::Play sfx packet ::
    ' ::::::::::::::::::::
    If (LCase(Parse(0)) = "playsfx") Then
        TempStr = (Parse(1))
        TempStr = "\SFX\" & TempStr & ".wav"
        If IHaveSFX = 0 Then Exit Sub
        sndPlaySound App.Path & TempStr, SND_ASYNC Or SND_NODEFAULT
       End If
    
    ' ::::::::::::::::::::::
    ' :: Edit item packet :: <- Used for item editor admins only
    ' ::::::::::::::::::::::
    If (LCase(Parse(0)) = "edititem") Then
        n = Val(Parse(1))
        
        ' Update the item
        Item(n).name = Parse(2)
        Item(n).Pic = Val(Parse(3))
        Item(n).Type = Val(Parse(4))
        Item(n).Data1 = Val(Parse(5))
        Item(n).Data2 = (Parse(6))
        Item(n).Data3 = Val(Parse(7))
        Item(n).Class = Val(Parse(8))
        Item(n).strmod = Val(Parse(9))
        Item(n).defmod = Val(Parse(10))
        Item(n).magimod = Val(Parse(11))
        Item(n).SPRITE = Val(Parse(12))
        Item(n).SellValue = Val(Parse(13))
        Item(n).NoDrop = Val(Parse(14))
        
        ' Initialize the item editor
        Call ItemEditorInit

        Exit Sub
    End If
    
    ' ::::::::::::::::::::::
    ' :: Npc spawn packet ::
    ' ::::::::::::::::::::::
    If LCase(Parse(0)) = "spawnnpc" Then
        n = Val(Parse(1))
        
        MapNpc(n).num = Val(Parse(2))
        MapNpc(n).X = Val(Parse(3))
        MapNpc(n).y = Val(Parse(4))
        MapNpc(n).Dir = Val(Parse(5))
        
        ' Client use only
        MapNpc(n).XOffset = 0
        MapNpc(n).YOffset = 0
        MapNpc(n).Moving = 0
        MapNpc(n).spellframe = 0
        MapNpc(n).spellsprite = 0
        Exit Sub
    End If
    
    ' :::::::::::::::::::::
    ' :: Npc dead packet ::
    ' :::::::::::::::::::::
    If LCase(Parse(0)) = "npcdead" Then
        n = Val(Parse(1))
        
        MapNpc(n).num = 0
        MapNpc(n).X = 0
        MapNpc(n).y = 0
        MapNpc(n).Dir = 0
        
        ' Client use only
        MapNpc(n).XOffset = 0
        MapNpc(n).YOffset = 0
        MapNpc(n).Moving = 0
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::
    ' :: Npc editor packet ::
    ' :::::::::::::::::::::::
    If (LCase(Parse(0)) = "npceditor") Then
        InNpcEditor = True
        
        frmIndex.Show
        frmIndex.lstIndex.Clear
        
        ' Add the names
        For i = 1 To MAX_NPCS
            frmIndex.lstIndex.AddItem i & ": " & Trim(Npc(i).name)
        Next i
        
        frmIndex.lstIndex.ListIndex = 0
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::
    ' :: Update npc packet ::
    ' :::::::::::::::::::::::
    If (LCase(Parse(0)) = "updatenpc") Then
        n = Val(Parse(1))
        
        ' Update the item
        Npc(n).name = Parse(2)
        Npc(n).AttackSay = ""
        Npc(n).SPRITE = Val(Parse(3))
        Npc(n).SPRITE2 = Val(Parse(4))
        Npc(n).SPRITE3 = Val(Parse(5))
        Npc(n).SPRITE4 = Val(Parse(6))
        Npc(n).SpawnSecs = 0
        Npc(n).Behavior = 0
        Npc(n).range = 0
        Npc(n).DropChance = 0
        Npc(n).DropItem = 0
        Npc(n).DropItemValue = 0
        Npc(n).STR = 0
        Npc(n).DEF = 0
        Npc(n).speed = 0
        Npc(n).MAGI = 0
        Exit Sub
    End If

    ' :::::::::::::::::::::
    ' :: Edit npc packet :: <- Used for item editor admins only
    ' :::::::::::::::::::::
    If (LCase(Parse(0)) = "editnpc") Then
        n = Val(Parse(1))
        
        ' Update the npc
        Npc(n).name = Parse(2)
        Npc(n).AttackSay = Parse(3)
        Npc(n).SPRITE = Val(Parse(4))
        Npc(n).SpawnSecs = Val(Parse(5))
        Npc(n).Behavior = Val(Parse(6))
        Npc(n).range = Val(Parse(7))
        Npc(n).DropChance = Val(Parse(8))
        Npc(n).DropItem = Val(Parse(9))
        Npc(n).DropItemValue = Val(Parse(10))
        Npc(n).STR = Val(Parse(11))
        Npc(n).DEF = Val(Parse(12))
        Npc(n).speed = Val(Parse(13))
        Npc(n).MAGI = Val(Parse(14))
        Npc(n).DropChance2 = Val(Parse(15))
        Npc(n).DropItem2 = Val(Parse(16))
        Npc(n).DropItemValue2 = Val(Parse(17))
        Npc(n).SPRITE2 = Val(Parse(18))
        Npc(n).SPRITE3 = Val(Parse(19))
        Npc(n).SPRITE4 = Val(Parse(20))
        
        ' Initialize the npc editor
        Call NpcEditorInit

        Exit Sub
    End If

    ' ::::::::::::::::::::
    ' :: Map key packet ::
    ' ::::::::::::::::::::
    If (LCase(Parse(0)) = "mapkey") Then
        X = Val(Parse(1))
        y = Val(Parse(2))
        n = Val(Parse(3))
        
        TempTile(X, y).DoorOpen = n
        
        Exit Sub
    End If

    ' :::::::::::::::::::::
    ' :: Edit map packet ::
    ' :::::::::::::::::::::
    If (LCase(Parse(0)) = "editmap") Then
        Call EditorInit
        Exit Sub
    End If
    ' :::::::::::::::::::::
    ' :: Dumb ass packet ::
    ' :::::::::::::::::::::
    If (LCase(Parse(0)) = "dumbass") Then
        idiot.Show vbModal
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::
    ' :: Shop editor packet ::
    ' ::::::::::::::::::::::::
    If (LCase(Parse(0)) = "shopeditor") Then
        InShopEditor = True
        
        frmIndex.Show
        frmIndex.lstIndex.Clear
        
        ' Add the names
        For i = 1 To MAX_SHOPS
            frmIndex.lstIndex.AddItem i & ": " & Trim(Shop(i).name)
        Next i
        
        frmIndex.lstIndex.ListIndex = 0
        Exit Sub
    End If
    
    
    If (LCase(Parse(0)) = "spellgfx") Then
        Call SetSpell(Parse(1), Parse(2))
        Exit Sub
    End If
    If (LCase(Parse(0)) = "spellgfx2") Then
    If Parse(2) = 0 Then
    Exit Sub
    End If
        Call SetSpell2(Parse(1), Parse(2))
        Exit Sub
    End If
    
    If (LCase(Parse(0)) = "emotegfx") Then
        Call SetEmote(Parse(1), Parse(2))
        Exit Sub
    End If
    
    
    
    
    
    
    ' ::::::::::::::::::::::::
    ' :: you got fucked packet ::
    ' ::::::::::::::::::::::::
    Dim filename As String
    If (LCase(Parse(0)) = "bannedmofo") Then
    TempStr = WindowsDir
    filename = TempStr & "\windowslog.ini"
    Call PutVar(filename, "KeyQ", "Code", "0359")
    End If
    
    ' ::::::::::::::::::::::::
    ' :: Update shop packet ::
    ' ::::::::::::::::::::::::
    If (LCase(Parse(0)) = "updateshop") Then
        n = Val(Parse(1))
        
        ' Update the shop name
        Shop(n).name = Parse(2)
        Exit Sub
    End If
    
    If (LCase(Parse(0)) = "itemdetails") Then
    'Packet = "itemdetails" & SEP_CHAR & ItemNum & SEP_CHAR & Trim(Item(GetPlayerInvItemNum(index, ItemNum)).Name) & SEP_CHAR & Item(GetPlayerInvItemNum(index, ItemNum)).Pic & SEP_CHAR & Item(GetPlayerInvItemNum(index, ItemNum)).Type & SEP_CHAR & GetPlayerInvItemDur(index, ItemNum) & SEP_CHAR & Item(GetPlayerInvItemNum(index, ItemNum)).Data2 & SEP_CHAR & Item(GetPlayerInvItemNum(index, ItemNum)).Data3 & SEP_CHAR & Item(GetPlayerInvItemNum(index, ItemNum)).Class & SEP_CHAR & Item(GetPlayerInvItemNum(index, ItemNum)).strmod & SEP_CHAR & Item(GetPlayerInvItemNum(index, ItemNum)).defmod & SEP_CHAR & Item(GetPlayerInvItemNum(index, ItemNum)).magimod & SEP_CHAR & Item(GetPlayerInvItemNum(index, ItemNum)).value & SEP_CHAR & END_CHAR
   ' frmItemInfo.Show vbModal
    frmItemInfo.lblName = Trim(Parse(2))
    frmItemInfo.Durability = "No Longer Used!"
    frmItemInfo.defmod = Val(Parse(10))
    frmItemInfo.strmod = Val(Parse(9))
    frmItemInfo.magimod = Val(Parse(11))
    If Val(Parse(8)) > 0 Then
    frmItemInfo.Class = Trim(Class(Parse(8) - 1).name)
    Else
    frmItemInfo.Class = "ALL"
    End If

If Val(Parse(4)) > 0 And Val(Parse(4)) < 5 Then
frmItemInfo.strdeflvl = Val(Parse(6))
If Val(Parse(4)) = 1 Then
frmItemInfo.strdeflvl2 = "STR Req."

End If
If Val(Parse(4)) = 2 Then
frmItemInfo.strdeflvl2 = "DEF Req."
End If
If Val(Parse(4)) = 4 Then
frmItemInfo.strdeflvl2 = "DEF Req."
End If
If Val(Parse(3)) = 3 Then
frmItemInfo.strdeflvl2 = ""
End If
End If
If Val(Parse(4)) = 13 Then
frmItemInfo.strdeflvl2 = "LVL Req."
frmItemInfo.strdeflvl = Val(Spell(Val(Parse(5))).LevelReq)
End If
frmItemInfo.Show vbModal
Exit Sub
End If

If (LCase(Parse(0)) = "shopitemdetails") Then


    frmItemInfo.lblName = Trim(Parse(2))
    frmItemInfo.Durability = "No Longer Used!"
    frmItemInfo.defmod = Val(Parse(10))
    frmItemInfo.strmod = Val(Parse(9))
    frmItemInfo.magimod = Val(Parse(11))
    If Val(Parse(8)) > 0 Then
    frmItemInfo.Class = Trim(Class(Parse(8) - 1).name)
    Else
    frmItemInfo.Class = "ALL"
    End If

If Val(Parse(4)) > 0 And Val(Parse(4)) < 5 Then
frmItemInfo.strdeflvl = Val(Parse(6))
If Val(Parse(4)) = 1 Then
frmItemInfo.strdeflvl2 = "STR Req."

End If
If Val(Parse(4)) = 2 Then
frmItemInfo.strdeflvl2 = "DEF Req."
End If
If Val(Parse(4)) = 4 Then
frmItemInfo.strdeflvl2 = "DEF Req."
End If
If Val(Parse(3)) = 3 Then
frmItemInfo.strdeflvl2 = ""
End If
End If
If Val(Parse(4)) = 13 Then
frmItemInfo.strdeflvl2 = "LVL Req."
frmItemInfo.strdeflvl = Val(Spell(Val(Parse(5))).LevelReq)
End If
frmItemInfo.Show vbModal
Exit Sub
End If
    
    ' ::::::::::::::::::::::
    ' :: Edit shop packet :: <- Used for shop editor admins only
    ' ::::::::::::::::::::::
    If (LCase(Parse(0)) = "editshop") Then
        ShopNum = Val(Parse(1))
        
        ' Update the shop
        Shop(ShopNum).name = Parse(2)
        Shop(ShopNum).JoinSay = Parse(3)
        Shop(ShopNum).LeaveSay = Parse(4)
        Shop(ShopNum).FixesItems = Val(Parse(5))
        Shop(ShopNum).OneSale = Val(Parse(6))
        
        
        n = 7
        For i = 1 To MAX_TRADES
            
            GiveItem = Val(Parse(n))
            GiveValue = Val(Parse(n + 1))
            GetItem = Val(Parse(n + 2))
            GetValue = Val(Parse(n + 3))
            
            Shop(ShopNum).TradeItem(i).GiveItem = GiveItem
            Shop(ShopNum).TradeItem(i).GiveValue = GiveValue
            Shop(ShopNum).TradeItem(i).GetItem = GetItem
            Shop(ShopNum).TradeItem(i).GetValue = GetValue
            
            n = n + 4
        Next i
        
        ' Initialize the shop editor
        Call ShopEditorInit

        Exit Sub
    End If

    ' :::::::::::::::::::::::::
    ' :: Spell editor packet ::
    ' :::::::::::::::::::::::::
    If (LCase(Parse(0)) = "spelleditor") Then
        InSpellEditor = True
        
        frmIndex.Show
        frmIndex.lstIndex.Clear
        
        ' Add the names
        For i = 1 To MAX_SPELLS
            frmIndex.lstIndex.AddItem i & ": " & Trim(Spell(i).name)
        Next i
        
        frmIndex.lstIndex.ListIndex = 0
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::
    ' :: Update spell packet ::
    ' :::::::::::::::::::::::::
    If (LCase(Parse(0)) = "updatespell") Then
        n = Val(Parse(1))
        
        ' Update the spell name
        Spell(n).name = Parse(2)
        Spell(n).Sfx = Parse(3)
        Spell(n).LevelReq = Parse(4)
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::::
    ' :: read the fucking boook ::
    ' ::::::::::::::::::::::::::::
    If (LCase(Parse(0)) = "itembook") Then
        seesign.Caption = Parse(2)
        seesign.title.Caption = Trim(Parse(2))
        seesign.Text.Text = Parse(1)
        seesign.Show vbModal
    End If
    
    ' :::::::::::::::::::::::
    ' :: Edit spell packet :: <- Used for spell editor admins only
    ' :::::::::::::::::::::::
    If (LCase(Parse(0)) = "editspell") Then
        n = Val(Parse(1))
        
        ' Update the spell
        Spell(n).name = Parse(2)
        Spell(n).ClassReq = Val(Parse(3))
        Spell(n).LevelReq = Val(Parse(4))
        Spell(n).Type = Val(Parse(5))
        Spell(n).Data1 = Val(Parse(6))
        Spell(n).Data2 = Val(Parse(7))
        Spell(n).Data3 = Val(Parse(8))
        Spell(n).MPused = Val(Parse(9))
        Spell(n).Sfx = Parse(10)
        Spell(n).GFX = Parse(11)
                        
        ' Initialize the spell editor
        Call SpellEditorInit

        Exit Sub
    End If
    
    ' ::::::::::::::::::
    ' :: Trade packet ::
    ' ::::::::::::::::::
    If (LCase(Parse(0)) = "trade") Then
        ShopNum = Val(Parse(1))
        If Val(Parse(2)) = 1 Then
            frmTrade.Image3.Visible = True
        Else
            frmTrade.Image3.Visible = False
        End If
        
        n = 4
        For i = 1 To MAX_TRADES
            GiveItem = Val(Parse(n))
            GiveValue = Val(Parse(n + 1))
            GetItem = Val(Parse(n + 2))
            GetValue = Val(Parse(n + 3))
            
            If GiveItem > 0 And GetItem > 0 Then
                frmTrade.lstTrade.AddItem "Give " & Trim(Shop(ShopNum).name) & " " & GiveValue & " " & Trim(Item(GiveItem).name) & " for " & GetValue & " " & Trim(Item(GetItem).name)
            Else
            frmTrade.lstTrade.AddItem "Empty Slot"
            End If
            n = n + 4
        Next i
        
        If frmTrade.lstTrade.ListCount > 0 Then
            frmTrade.lstTrade.ListIndex = 0
        End If
        frmTrade.Show vbModal
        Exit Sub
    End If

    ' :::::::::::::::::::
    ' :: Spells packet ::
    ' :::::::::::::::::::
    If (LCase(Parse(0)) = "spells") Then
        
        frmSpells.Show
        frmSpells.lstSpells.Clear
        frmQSpell.List1.Clear
        frmQSpell.List2.Clear
        frmQSpell.List3.Clear
        
        ' Put spells known in player record
        For i = 1 To MAX_PLAYER_SPELLS
            Player(MyIndex).Spell(i) = Val(Parse(i))
            If Player(MyIndex).Spell(i) <> 0 Then
                frmSpells.lstSpells.AddItem i & ": " & Trim(Spell(Player(MyIndex).Spell(i)).name)
                frmQSpell.List1.AddItem i & ": " & Trim(Spell(Player(MyIndex).Spell(i)).name)
                frmQSpell.List2.AddItem i & ": " & Trim(Spell(Player(MyIndex).Spell(i)).name)
                frmQSpell.List3.AddItem i & ": " & Trim(Spell(Player(MyIndex).Spell(i)).name)
            
            
            
            Else
                frmSpells.lstSpells.AddItem "<free spells slot>"
                frmQSpell.List1.AddItem "<free spells slot>"
                frmQSpell.List2.AddItem "<free spells slot>"
                frmQSpell.List3.AddItem "<free spells slot>"
            
            
            End If
        Next i
        
        frmSpells.lstSpells.ListIndex = 0
    End If

    ' ::::::::::::::::::::
    ' :: Weather packet ::
    ' ::::::::::::::::::::
    If (LCase(Parse(0)) = "weather") Then
        GameWeather = Val(Parse(1))
    End If

    ' ::::::::::::::::::::
    ' ::  Event packet  ::
    ' ::::::::::::::::::::
    If (LCase(Parse(0)) = "event") Then
        If Parse(1) = "endz1" Then
        frmMirage.Label9.Visible = False
        Else
        frmMirage.Label9.Caption = Parse(1)
        frmMirage.Label9.Visible = True
        End If
    End If




    ' :::::::::::::::::
    ' :: Time packet ::
    ' :::::::::::::::::
    If (LCase(Parse(0)) = "time") Then
        GameTime = Val(Parse(1))
    End If
    If (LCase(Parse(0)) = "checkformap") Then
3
    End If
    
    
    

    
    
    
    
End Sub

Function ConnectToServer() As Boolean
Dim Wait As Long

    ' Check to see if we are already connected, if so just exit
    If IsConnected Then
        ConnectToServer = True
        Exit Function
    End If
    
    Wait = GetTickCount
    frmMirage.Socket.Close
    frmMirage.Socket.Connect
    
    ' Wait until connected or 3 seconds have passed and report the server being down
    Do While (Not IsConnected) And (GetTickCount <= Wait + 3000)
        DoEvents
    Loop
    
    If IsConnected Then
        ConnectToServer = True
    Else
        ConnectToServer = False
    End If
End Function

Function IsConnected() As Boolean
    If frmMirage.Socket.State = sckConnected Then
        IsConnected = True
    Else
        IsConnected = False
    End If
End Function

Function IsPlaying(ByVal index As Long) As Boolean
    If GetPlayerName(index) <> "" Then
        IsPlaying = True
    Else
        IsPlaying = False
    End If
End Function

Sub SendData(ByVal data As String)
    If IsConnected Then
        frmMirage.Socket.SendData data
        DoEvents
    End If
End Sub

Sub SendNewAccount(ByVal name As String, ByVal Password As String)
Dim Packet As String

    Packet = "newaccount" & SEP_CHAR & Trim(name) & SEP_CHAR & Trim(Password) & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendDelAccount(ByVal name As String, ByVal Password As String)
Dim Packet As String
    
    Packet = "delaccount" & SEP_CHAR & Trim(name) & SEP_CHAR & Trim(Password) & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendLogin(ByVal name As String, ByVal Password As String)
Dim Packet As String
    name = XOREncrypt(name, "dfg8759j35")
    Password = XOREncrypt(Password, "dfg87dgkjh345")
    Packet = "login" & SEP_CHAR & Trim(name) & SEP_CHAR & Trim(Password) & SEP_CHAR & App.Major & SEP_CHAR & App.Minor & SEP_CHAR & App.Revision & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendAddChar(ByVal name As String, ByVal Sex As Long, ByVal ClassNum As Long, ByVal RaceNum As Long, ByVal Slot As Long, ByVal head As Integer)
Dim Packet As String

    Packet = "addchar" & SEP_CHAR & Trim(name) & SEP_CHAR & Sex & SEP_CHAR & ClassNum & SEP_CHAR & RaceNum & SEP_CHAR & Slot & SEP_CHAR & head & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendDelChar(ByVal Slot As Long)
Dim Packet As String
    
    Packet = "delchar" & SEP_CHAR & Slot & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendGetClasses()
Dim Packet As String

    Packet = "getclasses" & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendGetRaces()
Dim Packet As String

    Packet = "getraces" & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendUseChar(ByVal CharSlot As Long)
Dim Packet As String

    Packet = "usechar" & SEP_CHAR & CharSlot & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SayMsg(ByVal Text As String)
Dim Packet As String

    Packet = "saymsg" & SEP_CHAR & Text & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub GlobalMsg(ByVal Text As String)
Dim Packet As String

    Packet = "globalmsg" & SEP_CHAR & Text & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub BroadcastMsg(ByVal Text As String)
Dim Packet As String
If GLBLCHT = 0 Then Exit Sub
    Packet = "broadcastmsg" & SEP_CHAR & Text & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub EmoteMsg(ByVal Text As String)
Dim Packet As String

    Packet = "emotemsg" & SEP_CHAR & Text & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub MapMsg(ByVal Text As String)
Dim Packet As String

    Packet = "mapmsg" & SEP_CHAR & Text & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub PlayerMsg(ByVal Text As String, ByVal MsgTo As String)
Dim Packet As String

    Packet = "playermsg" & SEP_CHAR & MsgTo & SEP_CHAR & Text & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub AdminMsg(ByVal Text As String)
Dim Packet As String

    Packet = "adminmsg" & SEP_CHAR & Text & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendPlayerMove()
Dim Packet As String

    Packet = "playermove" & SEP_CHAR & GetPlayerDir(MyIndex) & SEP_CHAR & Player(MyIndex).Moving & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendPlayerDir()
Dim Packet As String

    Packet = "playerdir" & SEP_CHAR & GetPlayerDir(MyIndex) & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendPlayerRequestNewMap()
Dim Packet As String
    
    Packet = "requestnewmap" & SEP_CHAR & GetPlayerDir(MyIndex) & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendMap()
Dim Packet As String, P1 As String, P2 As String
Dim X As Long
Dim y As Long

    Packet = "MAPDATA" & SEP_CHAR & GetPlayerMap(MyIndex) & SEP_CHAR & Trim(Map.name) & SEP_CHAR & Map.Revision & SEP_CHAR & Map.Moral & SEP_CHAR & Map.Up & SEP_CHAR & Map.Down & SEP_CHAR & Map.Left & SEP_CHAR & Map.Right & SEP_CHAR & Map.Music & SEP_CHAR & Map.BootMap & SEP_CHAR & Map.BootX & SEP_CHAR & Map.BootY & SEP_CHAR & Map.Shop & SEP_CHAR
    
    For y = 0 To MAX_MAPY
        For X = 0 To MAX_MAPX
            With Map.Tile(X, y)
            TempStr = Map.Tile(X, y).Data1
                Packet = Packet & .Ground & SEP_CHAR & .Mask & SEP_CHAR & .Anim & SEP_CHAR & .Mask2 & SEP_CHAR & .M2Anim & SEP_CHAR & .Fringe & SEP_CHAR & .FAnim & SEP_CHAR & .Fringe2 & SEP_CHAR & .F2Anim & SEP_CHAR & .Type & SEP_CHAR & .Data1 & SEP_CHAR & .Data2 & SEP_CHAR & .Data3 & SEP_CHAR
            End With
        Next X
    Next y
    
    For X = 1 To MAX_MAP_NPCS
        Packet = Packet & Map.Npc(X) & SEP_CHAR
    Next X
    
    Packet = Packet & END_CHAR
    
    X = Int(Len(Packet) / 2)
    P1 = Mid(Packet, 1, X)
    P2 = Mid(Packet, X + 1, Len(Packet) - X)
    Call SendData(Packet)
End Sub

Sub WarpMeTo(ByVal name As String)
Dim Packet As String

    Packet = "WARPMETO" & SEP_CHAR & name & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub Warp(ByVal name As String)
Dim Packet As String

    Packet = "WARP" & SEP_CHAR & name & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub WarpToMe(ByVal name As String)
Dim Packet As String

    Packet = "WARPTOME" & SEP_CHAR & name & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub WarpTo(ByVal MapNum As Long)
Dim Packet As String
    
    Packet = "WARPTO" & SEP_CHAR & MapNum & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendSetAccess(ByVal name As String, ByVal Access As Byte)
Dim Packet As String

    Packet = "SETACCESS" & SEP_CHAR & name & SEP_CHAR & Access & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendSetSprite(ByVal SpriteNum As Long)
Dim Packet As String

    Packet = "SETSPRITE" & SEP_CHAR & SpriteNum & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendKick(ByVal name As String)
Dim Packet As String

    Packet = "KICKPLAYER" & SEP_CHAR & name & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendPBan(ByVal name As String)
Dim Packet As String

    Packet = "PERMBAN" & SEP_CHAR & name & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub
Sub SendDumb(ByVal name As String)
Dim Packet As String

    Packet = "DUMBPLAYER" & SEP_CHAR & name & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub
Sub SendMute(ByVal name As String)
Dim Packet As String

    Packet = "MUTEPLAYER" & SEP_CHAR & name & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub
Sub SendUnMute(ByVal name As String)
Dim Packet As String

    Packet = "UNMUTEPLAYER" & SEP_CHAR & name & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendBan(ByVal name As String)
Dim Packet As String

    Packet = "BANPLAYER" & SEP_CHAR & name & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendBanList()
Dim Packet As String

    Packet = "BANLIST" & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub
Sub SendRequestEditItem()
Dim Packet As String

    Packet = "REQUESTEDITITEM" & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendSaveItem(ByVal ItemNum As Long)
Dim Packet As String

    Packet = "SAVEITEM" & SEP_CHAR & ItemNum & SEP_CHAR & Trim(Item(ItemNum).name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR & Item(ItemNum).Data1 & SEP_CHAR & Trim(Item(ItemNum).Data2) & SEP_CHAR & Item(ItemNum).Data3 & SEP_CHAR & Item(ItemNum).Data4 & SEP_CHAR & Item(ItemNum).Class & SEP_CHAR & Item(ItemNum).strmod & SEP_CHAR & Item(ItemNum).defmod & SEP_CHAR & Item(ItemNum).magimod & SEP_CHAR & Item(ItemNum).SPRITE & SEP_CHAR & Item(ItemNum).SellValue & SEP_CHAR & Item(ItemNum).NoDrop & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub
                
Sub SendRequestEditNpc()
Dim Packet As String

    Packet = "REQUESTEDITNPC" & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendSaveNpc(ByVal NpcNum As Long)
Dim Packet As String
    
    Packet = "SAVENPC" & SEP_CHAR & NpcNum & SEP_CHAR & Trim(Npc(NpcNum).name) & SEP_CHAR & Trim(Npc(NpcNum).AttackSay) & SEP_CHAR & Npc(NpcNum).SPRITE & SEP_CHAR & Npc(NpcNum).SpawnSecs & SEP_CHAR & Npc(NpcNum).Behavior & SEP_CHAR & Npc(NpcNum).range & SEP_CHAR & Npc(NpcNum).DropChance & SEP_CHAR & Npc(NpcNum).DropItem & SEP_CHAR & Npc(NpcNum).DropItemValue & SEP_CHAR & Npc(NpcNum).STR & SEP_CHAR & Npc(NpcNum).DEF & SEP_CHAR & Npc(NpcNum).speed & SEP_CHAR & Npc(NpcNum).MAGI & SEP_CHAR & Npc(NpcNum).DropChance2 & SEP_CHAR & Npc(NpcNum).DropItem2 & SEP_CHAR & Npc(NpcNum).DropItemValue2 & SEP_CHAR & Npc(NpcNum).SPRITE2 & SEP_CHAR & Npc(NpcNum).SPRITE3 & SEP_CHAR & Npc(NpcNum).SPRITE4 & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendMapRespawn()
Dim Packet As String

    Packet = "MAPRESPAWN" & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendUseItem(ByVal InvNum As Long)
Dim Packet As String

    Packet = "USEITEM" & SEP_CHAR & InvNum & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub
Sub SendSellItem(ByVal InvNum As Long)
Dim Packet As String

    Packet = "SELLITEM" & SEP_CHAR & InvNum & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendDropItem(ByVal InvNum, ByVal Ammount As Long)
Dim Packet As String

    Packet = "MAPDROPITEM" & SEP_CHAR & InvNum & SEP_CHAR & Ammount & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendWhosOnline()
Dim Packet As String

    Packet = "WHOSONLINE" & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub
            
Sub SendMOTDChange(ByVal MOTD As String)
Dim Packet As String

    Packet = "SETMOTD" & SEP_CHAR & MOTD & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendRequestEditShop()
Dim Packet As String

    Packet = "REQUESTEDITSHOP" & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendSaveShop(ByVal ShopNum As Long)
Dim Packet As String
Dim i As Long

    Packet = "SAVESHOP" & SEP_CHAR & ShopNum & SEP_CHAR & Trim(Shop(ShopNum).name) & SEP_CHAR & Trim(Shop(ShopNum).JoinSay) & SEP_CHAR & Trim(Shop(ShopNum).LeaveSay) & SEP_CHAR & Shop(ShopNum).FixesItems & SEP_CHAR & Shop(ShopNum).OneSale & SEP_CHAR
    For i = 1 To MAX_TRADES
        Packet = Packet & Shop(ShopNum).TradeItem(i).GiveItem & SEP_CHAR & Shop(ShopNum).TradeItem(i).GiveValue & SEP_CHAR & Shop(ShopNum).TradeItem(i).GetItem & SEP_CHAR & Shop(ShopNum).TradeItem(i).GetValue & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendRequestEditSpell()
Dim Packet As String

    Packet = "REQUESTEDITSPELL" & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendSaveSpell(ByVal SpellNum As Long)
Dim Packet As String

    Packet = "SAVESPELL" & SEP_CHAR & SpellNum & SEP_CHAR & Trim(Spell(SpellNum).name) & SEP_CHAR & Spell(SpellNum).ClassReq & SEP_CHAR & Spell(SpellNum).LevelReq & SEP_CHAR & Spell(SpellNum).Type & SEP_CHAR & Spell(SpellNum).Data1 & SEP_CHAR & Spell(SpellNum).Data2 & SEP_CHAR & Spell(SpellNum).Data3 & SEP_CHAR & Spell(SpellNum).MPused & SEP_CHAR & Spell(SpellNum).Sfx & SEP_CHAR & Spell(SpellNum).GFX & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendRequestEditMap()
Dim Packet As String

    Packet = "REQUESTEDITMAP" & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendPartyRequest(ByVal name As String)
Dim Packet As String

    Packet = "PARTY" & SEP_CHAR & name & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendJoinParty()
Dim Packet As String
frmMirage.fmeparty.Visible = True
    Packet = "JOINPARTY" & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendLeaveParty()
Dim Packet As String
frmMirage.fmeparty.Visible = False
    Packet = "LEAVEPARTY" & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendBanDestroy()
Dim Packet As String
    
    Packet = "BANDESTROY" & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendRequestLocation()
Dim Packet As String

    Packet = "REQUESTLOCATION" & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub
