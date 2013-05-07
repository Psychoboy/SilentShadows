Attribute VB_Name = "modGameLogic"
Option Explicit

Function GetPlayerDamage(ByVal index As Long) As Long
Dim WeaponSlot As Long

    GetPlayerDamage = 0
    
    ' Check for subscript out of range
    If IsPlaying(index) = False Or index <= 0 Or index > MAX_PLAYERS Then
        Exit Function
    End If
    
    GetPlayerDamage = Int(GetPlayerSTR(index) / 2)
    
    If GetPlayerDamage <= 0 Then
        GetPlayerDamage = 1
    End If
    
    If GetPlayerWeaponSlot(index) > 0 Then
        WeaponSlot = GetPlayerWeaponSlot(index)
        
        GetPlayerDamage = GetPlayerDamage + Item(GetPlayerInvItemNum(index, WeaponSlot)).Data2
        
        'Call SetPlayerInvItemDur(Index, WeaponSlot, GetPlayerInvItemDur(Index, WeaponSlot) - 1)
        
        If GetPlayerInvItemDur(index, WeaponSlot) <= 0 Then
            Call PlayerMsgCombat(index, "Your " & Trim(Item(GetPlayerInvItemNum(index, WeaponSlot)).Name) & " has broken.", Yellow)
            Temp = GetPlayerWeaponSlot(index)
            Call SetPlayerSTR(index, (GetPlayerSTR(index) - Item(GetPlayerInvItemNum(index, Temp)).STRmod))
                        Call SetPlayerDEF(index, (GetPlayerDEF(index) - Item(GetPlayerInvItemNum(index, Temp)).DEFmod))
                        Call SetPlayerMAGI(index, (GetPlayerMAGI(index) - Item(GetPlayerInvItemNum(index, Temp)).MAGImod))
                        Call SetPlayerSprite4(index, 2000)
                        Call SendPlayerData(index)
            Call TakeItem(index, GetPlayerInvItemNum(index, WeaponSlot), 0)
        Else
            If GetPlayerInvItemDur(index, WeaponSlot) <= 0 Then
                Call PlayerMsgCombat(index, "Your " & Trim(Item(GetPlayerInvItemNum(index, WeaponSlot)).Name) & " is about to break!", Yellow)
            End If
        End If
    End If
End Function

Function GetPlayerProtection(ByVal index As Long) As Long
Dim ArmorSlot As Long, HelmSlot As Long
    
    GetPlayerProtection = 0
    
    ' Check for subscript out of range
    If IsPlaying(index) = False Or index <= 0 Or index > MAX_PLAYERS Then
        Exit Function
    End If
    
    ArmorSlot = GetPlayerArmorSlot(index)
    HelmSlot = GetPlayerHelmetSlot(index)
    GetPlayerProtection = Int(GetPlayerDEF(index) / 5)

    If ArmorSlot > 0 Then
        GetPlayerProtection = GetPlayerProtection + Item(GetPlayerInvItemNum(index, ArmorSlot)).Data2
        'Call SetPlayerInvItemDur(Index, ArmorSlot, GetPlayerInvItemDur(Index, ArmorSlot) - 1)
        
        If GetPlayerInvItemDur(index, ArmorSlot) <= 0 Then
            Call PlayerMsgCombat(index, "Your " & Trim(Item(GetPlayerInvItemNum(index, ArmorSlot)).Name) & " has broken.", Yellow)
            Temp = GetPlayerArmorSlot(index)
            Call SetPlayerSTR(index, (GetPlayerSTR(index) - Item(GetPlayerInvItemNum(index, Temp)).STRmod))
                        Call SetPlayerDEF(index, (GetPlayerDEF(index) - Item(GetPlayerInvItemNum(index, Temp)).DEFmod))
                        Call SetPlayerMAGI(index, (GetPlayerMAGI(index) - Item(GetPlayerInvItemNum(index, Temp)).MAGImod))
                        If GetPlayerSex(index) = 0 Then
                        Call SetPlayerSprite(index, 0)
                        Else
                        Call SetPlayerSprite(index, 1)
                        End If
                        Call SendPlayerData(index)
            Call TakeItem(index, GetPlayerInvItemNum(index, ArmorSlot), 0)
            
        Else
            If GetPlayerInvItemDur(index, ArmorSlot) <= 0 Then
                Call PlayerMsgCombat(index, "Your " & Trim(Item(GetPlayerInvItemNum(index, ArmorSlot)).Name) & " is about to break!", Yellow)
            End If
        End If
    End If
    
    'If HelmSlot > 0 Then
    '    GetPlayerProtection = GetPlayerProtection + Item(GetPlayerInvItemNum(index, HelmSlot)).Data2
    '    'Call SetPlayerInvItemDur(index, HelmSlot, GetPlayerInvItemDur(index, HelmSlot) - 1)
    '
    '    If GetPlayerInvItemDur(index, HelmSlot) <= 0 Then
    '        Call PlayerMsgCombat(index, "Your " & Trim(Item(GetPlayerInvItemNum(index, HelmSlot)).Name) & " has broken.", Yellow)
    '        Call TakeItem(index, GetPlayerInvItemNum(index, HelmSlot), 0)
    '    Else
    '        If GetPlayerInvItemDur(index, HelmSlot) <= 5 Then
    '            'Call PlayerMsgCombat(index, "Your " & Trim(Item(GetPlayerInvItemNum(index, HelmSlot)).Name) & " is about to break!", Yellow)
    '        End If
    '    End If
    'End If
End Function

Function FindOpenPlayerSlot() As Long
Dim i As Long

    FindOpenPlayerSlot = 0
    
    For i = 1 To MAX_PLAYERS
        If Not IsConnected(i) Then
            FindOpenPlayerSlot = i
            Exit Function
        End If
    Next i
End Function


Function FindOpenBankSlot(ByVal index As Long, ByVal ItemNum As Long) As Long
Dim i As Long
    
    FindOpenBankSlot = 0
    
    ' Check for subscript out of range
    If IsPlaying(index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Function
    End If
    
    If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Then
        ' If currency then check to see if they already have an instance of the item and add it to that
        For i = 1 To MAX_BANK
            If GetPlayerBankItemNum(index, i) = ItemNum Then
                FindOpenBankSlot = i
                Exit Function
            End If
        Next i
    End If
    
    For i = 1 To MAX_BANK
        ' Try to find an open free slot
        If GetPlayerBankItemNum(index, i) = 0 Then
            FindOpenBankSlot = i
            Exit Function
        End If
    Next i
End Function
Sub TakeBankItem(ByVal index As Long, ByVal ItemNum As Long, ByVal ItemVal As Long)
Dim i As Long, n As Long
Dim TakeItem As Boolean

    TakeItem = False
    
    ' Check for subscript out of range
    If IsPlaying(index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Sub
    End If
    
    For i = 1 To MAX_INV
        ' Check to see if the player has the item
        If GetPlayerBankItemNum(index, i) = ItemNum Then
            If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Then
                ' Is what we are trying to take away more then what they have?  If so just set it to zero
                If ItemVal >= GetPlayerBankItemValue(index, i) Then
                    TakeItem = True
                Else
                    Call SetPlayerBankItemValue(index, i, GetPlayerBankItemValue(index, i) - ItemVal)
                    Call SendInventoryUpdate(index, i)
                End If
            Else
                                
                n = Item(GetPlayerBankItemNum(index, i)).Type
                ' Check if its not an equipable weapon, and if it isn't then take it away
               ' If (n <> ITEM_TYPE_WEAPON) And (n <> ITEM_TYPE_ARMOR) And (n <> ITEM_TYPE_HELMET) And (n <> ITEM_TYPE_SHIELD) Then
                    TakeItem = True
                'End If
            End If
                            
            If TakeItem = True Then
                Call SetPlayerBankItemNum(index, i, 0)
                Call SetPlayerBankItemValue(index, i, 0)
                Call SetPlayerBankItemDur(index, i, 0)
                
                ' Send the inventory update
                'Call SendBankUpdate(Index, i)
                Exit Sub
            End If
        End If
    Next i
End Sub
Sub GiveBankItem(ByVal index As Long, ByVal ItemNum As Long, ByVal ItemVal As Long)
Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Sub
    End If
    
    i = FindOpenBankSlot(index, ItemNum)
    
    ' Check to see if inventory is full
    If i <> 0 Then
        Call SetPlayerBankItemNum(index, i, ItemNum)
        Call SetPlayerBankItemValue(index, i, GetPlayerBankItemValue(index, i) + ItemVal)
        
        If (Item(ItemNum).Type = ITEM_TYPE_ARMOR) Or (Item(ItemNum).Type = ITEM_TYPE_WEAPON) Or (Item(ItemNum).Type = ITEM_TYPE_HELMET) Or (Item(ItemNum).Type = ITEM_TYPE_SHIELD) Then
            Call SetPlayerBankItemDur(index, i, Item(ItemNum).Data1)
        End If
        
        'Call SendBankUpdate(Index, i)
    Else
        Call PlayerMsgCombat(index, "Your inventory is full.", Yellow)
    End If
End Sub
Function FindOpenInvSlot(ByVal index As Long, ByVal ItemNum As Long) As Long


Dim i As Long
    
    FindOpenInvSlot = 0
    
    ' Check for subscript out of range
    If IsPlaying(index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Function
    End If
    
    If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Then
        ' If currency then check to see if they already have an instance of the item and add it to that
        For i = 1 To MAX_INV
            If GetPlayerInvItemNum(index, i) = ItemNum Then
                FindOpenInvSlot = i
                Exit Function
            End If
        Next i
    End If
    
    For i = 1 To MAX_INV
        ' Try to find an open free slot
        If GetPlayerInvItemNum(index, i) = 0 Then
            FindOpenInvSlot = i
            Exit Function
        End If
    Next i
End Function

Function FindOpenMapItemSlot(ByVal MapNum As Long) As Long


Dim i As Long

    FindOpenMapItemSlot = 0
    
    ' Check for subscript out of range
    If MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Function
    End If
    
    For i = 1 To MAX_MAP_ITEMS
        If MapItem(MapNum, i).Num = 0 Then
            FindOpenMapItemSlot = i
            Exit Function
        End If
    Next i
End Function

Function FindOpenSpellSlot(ByVal index As Long) As Long


Dim i As Long

    FindOpenSpellSlot = 0
    
    For i = 1 To MAX_PLAYER_SPELLS
        If GetPlayerSpell(index, i) = 0 Then
            FindOpenSpellSlot = i
            Exit Function
        End If
    Next i
End Function

Function HasSpell(ByVal index As Long, ByVal SpellNum As Long) As Boolean


Dim i As Long

    HasSpell = False
    
    For i = 1 To MAX_PLAYER_SPELLS
        If GetPlayerSpell(index, i) = SpellNum Then
            HasSpell = True
            Exit Function
        End If
    Next i
End Function

Function TotalOnlinePlayers() As Long


Dim i As Long

    TotalOnlinePlayers = 0
    
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            TotalOnlinePlayers = TotalOnlinePlayers + 1
        End If
    Next i
End Function

Function FindPlayer(ByVal Name As String) As Long


Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            ' Make sure we dont try to check a name thats to small
            If Len(GetPlayerName(i)) = Len(Trim(Name)) Then
                'If UCase(Mid(GetPlayerName(i), 1, Len(Trim(Name)))) = UCase(Trim(Name)) Then
                 If UCase(Trim(GetPlayerName(i))) = UCase(Trim(Name)) Then
                    FindPlayer = i
                    Exit Function
                End If
            End If
        End If
    Next i
    
    FindPlayer = 0
End Function
Function FindOpenParty() As Long


Dim i As Long

    For i = 1 To MAX_PARTIES
        If Parties(i).NumParty = 0 Then
        FindOpenParty = i
        Exit Function
        End If
    Next i

End Function

Function HasItem(ByVal index As Long, ByVal ItemNum As Long) As Long


Dim i As Long
    
    HasItem = 0
    
    ' Check for subscript out of range
    If IsPlaying(index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Function
    End If
    
    For i = 1 To MAX_INV
        ' Check to see if the player has the item
        If GetPlayerInvItemNum(index, i) = ItemNum Then
            If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Then
                HasItem = GetPlayerInvItemValue(index, i)
            Else
                HasItem = 1
            End If
            Exit Function
        End If
    Next i
End Function

Sub TakeItem(ByVal index As Long, ByVal ItemNum As Long, ByVal ItemVal As Long)


Dim i As Long, n As Long
Dim TakeItem As Boolean

    TakeItem = False
    
    ' Check for subscript out of range
    If IsPlaying(index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Sub
    End If
    
    For i = 1 To MAX_INV
        ' Check to see if the player has the item
        If GetPlayerInvItemNum(index, i) = ItemNum Then
            If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Then
                ' Is what we are trying to take away more then what they have?  If so just set it to zero
                If ItemVal >= GetPlayerInvItemValue(index, i) Then
                    TakeItem = True
                Else
                    Call SetPlayerInvItemValue(index, i, GetPlayerInvItemValue(index, i) - ItemVal)
                    Call SendInventoryUpdate(index, i)
                End If
            Else
                ' Check to see if its any sort of ArmorSlot/WeaponSlot
                Select Case Item(GetPlayerInvItemNum(index, i)).Type
                    Case ITEM_TYPE_WEAPON
                        If GetPlayerWeaponSlot(index) > 0 Then
                            If i = GetPlayerWeaponSlot(index) Then
                                Call SetPlayerWeaponSlot(index, 0)
                                Call SendWornEquipment(index)
                                TakeItem = True
                            Else
                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index)) Then
                                    TakeItem = True
                                End If
                            End If
                        Else
                            TakeItem = True
                        End If
                
                    Case ITEM_TYPE_ARMOR
                        If GetPlayerArmorSlot(index) > 0 Then
                            If i = GetPlayerArmorSlot(index) Then
                                Call SetPlayerArmorSlot(index, 0)
                                Call SendWornEquipment(index)
                                TakeItem = True
                            Else
                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerInvItemNum(index, GetPlayerArmorSlot(index)) Then
                                    TakeItem = True
                                End If
                            End If
                        Else
                            TakeItem = True
                        End If
                    
                    Case ITEM_TYPE_HELMET
                        If GetPlayerHelmetSlot(index) > 0 Then
                            If i = GetPlayerHelmetSlot(index) Then
                                Call SetPlayerHelmetSlot(index, 0)
                                Call SendWornEquipment(index)
                                TakeItem = True
                            Else
                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerInvItemNum(index, GetPlayerHelmetSlot(index)) Then
                                    TakeItem = True
                                End If
                            End If
                        Else
                            TakeItem = True
                        End If
                    
                    Case ITEM_TYPE_SHIELD
                        If GetPlayerShieldSlot(index) > 0 Then
                            If i = GetPlayerShieldSlot(index) Then
                                Call SetPlayerShieldSlot(index, 0)
                                Call SendWornEquipment(index)
                                TakeItem = True
                            Else
                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerInvItemNum(index, GetPlayerShieldSlot(index)) Then
                                    TakeItem = True
                                End If
                            End If
                        Else
                            TakeItem = True
                        End If
                End Select

                
                n = Item(GetPlayerInvItemNum(index, i)).Type
                ' Check if its not an equipable weapon, and if it isn't then take it away
                If (n <> ITEM_TYPE_WEAPON) And (n <> ITEM_TYPE_ARMOR) And (n <> ITEM_TYPE_HELMET) And (n <> ITEM_TYPE_SHIELD) Then
                    TakeItem = True
                End If
            End If
                            
            If TakeItem = True Then
                Call SetPlayerInvItemNum(index, i, 0)
                Call SetPlayerInvItemValue(index, i, 0)
                Call SetPlayerInvItemDur(index, i, 0)
                
                ' Send the inventory update
                Call SendInventoryUpdate(index, i)
                Exit Sub
            End If
        End If
    Next i
End Sub

Sub GiveItem(ByVal index As Long, ByVal ItemNum As Long, ByVal ItemVal As Long)


Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Sub
    End If
    
    i = FindOpenInvSlot(index, ItemNum)
    
    ' Check to see if inventory is full
    If i <> 0 Then
        Call SetPlayerInvItemNum(index, i, ItemNum)
        Call SetPlayerInvItemValue(index, i, GetPlayerInvItemValue(index, i) + ItemVal)
        
        If (Item(ItemNum).Type = ITEM_TYPE_ARMOR) Or (Item(ItemNum).Type = ITEM_TYPE_WEAPON) Or (Item(ItemNum).Type = ITEM_TYPE_HELMET) Or (Item(ItemNum).Type = ITEM_TYPE_SHIELD) Then
            Call SetPlayerInvItemDur(index, i, Item(ItemNum).Data1)
        End If
        
        Call SendInventoryUpdate(index, i)
    Else
        Call PlayerMsgCombat(index, "Your inventory is full.", BrightRed)
    End If
End Sub

Sub SpawnItem(ByVal ItemNum As Long, ByVal ItemVal As Long, ByVal MapNum As Long, ByVal X As Long, ByVal y As Long)


Dim i As Long

    ' Check for subscript out of range
    If ItemNum < 0 Or ItemNum > MAX_ITEMS Or MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If
    If ItemVal < 1 Then ItemVal = 1
    ' Find open map item slot
    i = FindOpenMapItemSlot(MapNum)
    
    Call SpawnItemSlot(i, ItemNum, ItemVal, Item(ItemNum).Data1, MapNum, X, y)
End Sub

Sub SpawnItemSlot(ByVal MapItemSlot As Long, ByVal ItemNum As Long, ByVal ItemVal As Long, ByVal ItemDur As Long, ByVal MapNum As Long, ByVal X As Long, ByVal y As Long)


Dim Packet As String
Dim i As Long
    
    ' Check for subscript out of range
    If MapItemSlot <= 0 Or MapItemSlot > MAX_MAP_ITEMS Or ItemNum < 0 Or ItemNum > MAX_ITEMS Or MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If
    
    i = MapItemSlot
    
    If i <> 0 And ItemNum >= 0 And ItemNum <= MAX_ITEMS Then
        MapItem(MapNum, i).Num = ItemNum
        MapItem(MapNum, i).Value = ItemVal
        
        If ItemNum <> 0 Then
            If (Item(ItemNum).Type >= ITEM_TYPE_WEAPON) And (Item(ItemNum).Type <= ITEM_TYPE_SHIELD) Then
                MapItem(MapNum, i).Dur = ItemDur
            Else
                MapItem(MapNum, i).Dur = 0
            End If
        Else
            MapItem(MapNum, i).Dur = 0
        End If
        
        MapItem(MapNum, i).X = X
        MapItem(MapNum, i).y = y
            
        Packet = "SPAWNITEM" & SEP_CHAR & i & SEP_CHAR & ItemNum & SEP_CHAR & ItemVal & SEP_CHAR & MapItem(MapNum, i).Dur & SEP_CHAR & X & SEP_CHAR & y & SEP_CHAR & END_CHAR
        Call SendDataToMap(MapNum, Packet)
    End If
End Sub

Sub SpawnAllMapsItems()


Dim i As Long
    
    For i = 1 To MAX_MAPS
        Call SpawnMapItems(i)
    Next i
End Sub

Sub SpawnMapItems(ByVal MapNum As Long)


Dim X As Long
Dim y As Long
Dim i As Long

    ' Check for subscript out of range
    If MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If
    
    ' Spawn what we have
    For y = 0 To MAX_MAPY
        For X = 0 To MAX_MAPX
            ' Check if the tile type is an item or a saved tile incase someone drops something
            If (MAP(MapNum).Tile(X, y).Type = TILE_TYPE_ITEM) Then
                ' Check to see if its a currency and if they set the value to 0 set it to 1 automatically
                If Item(MAP(MapNum).Tile(X, y).Data1).Type = ITEM_TYPE_CURRENCY And MAP(MapNum).Tile(X, y).Data2 <= 0 Then
                    Call SpawnItem(MAP(MapNum).Tile(X, y).Data1, 1, MapNum, X, y)
                Else
                    Call SpawnItem(MAP(MapNum).Tile(X, y).Data1, MAP(MapNum).Tile(X, y).Data2, MapNum, X, y)
                End If
            End If
        Next X
    Next y
End Sub

Sub PlayerMapGetItem(ByVal index As Long)


Dim i As Long
Dim n As Long
Dim MapNum As Long
Dim Msg As String

    If IsPlaying(index) = False Then
        Exit Sub
    End If
    
    MapNum = GetPlayerMap(index)
    
    For i = 1 To MAX_MAP_ITEMS
        ' See if theres even an item here
        If (MapItem(MapNum, i).Num > 0) And (MapItem(MapNum, i).Num <= MAX_ITEMS) Then
            ' Check if item is at the same location as the player
            If (MapItem(MapNum, i).X = GetPlayerX(index)) And (MapItem(MapNum, i).y = GetPlayerY(index)) Then
                ' Find open slot
                n = FindOpenInvSlot(index, MapItem(MapNum, i).Num)
                
                ' Open slot available?
                If n <> 0 Then
                    ' Set item in players inventor
                    Call SetPlayerInvItemNum(index, n, MapItem(MapNum, i).Num)
                    If Item(GetPlayerInvItemNum(index, n)).Type = ITEM_TYPE_CURRENCY Then
                        Call SetPlayerInvItemValue(index, n, GetPlayerInvItemValue(index, n) + MapItem(MapNum, i).Value)
                        Msg = "You picked up " & MapItem(MapNum, i).Value & " " & Trim(Item(GetPlayerInvItemNum(index, n)).Name) & "."
                    Else
                        Call SetPlayerInvItemValue(index, n, 0)
                        Msg = "You picked up a " & Trim(Item(GetPlayerInvItemNum(index, n)).Name) & "."
                    End If
                    Call SetPlayerInvItemDur(index, n, MapItem(MapNum, i).Dur)
                        
                    ' Erase item from the map
                    MapItem(MapNum, i).Num = 0
                    MapItem(MapNum, i).Value = 0
                    MapItem(MapNum, i).Dur = 0
                    MapItem(MapNum, i).X = 0
                    MapItem(MapNum, i).y = 0
                        
                    Call SendInventoryUpdate(index, n)
                    Call SpawnItemSlot(i, 0, 0, 0, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index))
                    Call PlayerMsgCombat(index, Msg, Yellow)
                    Exit Sub
                Else
                    Call PlayerMsgCombat(index, "Your inventory is full.", BrightRed)
                    Exit Sub
                End If
            End If
        End If
    Next i
End Sub

Sub PlayerMapDropItem(ByVal index As Long, ByVal InvNum As Long, ByVal Ammount As Long)


Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(index) = False Or InvNum <= 0 Or InvNum > MAX_INV Then
        Exit Sub
    End If
    
    If (GetPlayerInvItemNum(index, InvNum) > 0) And (GetPlayerInvItemNum(index, InvNum) <= MAX_ITEMS) Then
        i = FindOpenMapItemSlot(GetPlayerMap(index))
        
        If i <> 0 Then
            MapItem(GetPlayerMap(index), i).Dur = 0
            
            ' Check to see if its any sort of ArmorSlot/WeaponSlot
            Select Case Item(GetPlayerInvItemNum(index, InvNum)).Type
                Case ITEM_TYPE_ARMOR
                    If InvNum = GetPlayerArmorSlot(index) Then
                    Temp = GetPlayerArmorSlot(index)
                        Call SetPlayerSTR(index, (GetPlayerSTR(index) - Item(GetPlayerInvItemNum(index, Temp)).STRmod))
                        Call SetPlayerDEF(index, (GetPlayerDEF(index) - Item(GetPlayerInvItemNum(index, Temp)).DEFmod))
                        Call SetPlayerMAGI(index, (GetPlayerMAGI(index) - Item(GetPlayerInvItemNum(index, Temp)).MAGImod))
                        Call SetPlayerArmorSlot(index, 0)
                        Call SendWornEquipment(index)
                        If GetPlayerSex(index) = 0 Then
                        Call SetPlayerSprite(index, 0)
                        Else
                        Call SetPlayerSprite(index, 1)
                        End If
                        Call SendPlayerData(index)
                    End If
                    MapItem(GetPlayerMap(index), i).Dur = GetPlayerInvItemDur(index, InvNum)
                
                Case ITEM_TYPE_WEAPON
                    If InvNum = GetPlayerWeaponSlot(index) Then
                    Temp = GetPlayerWeaponSlot(index)
                        Call SetPlayerSTR(index, (GetPlayerSTR(index) - Item(GetPlayerInvItemNum(index, Temp)).STRmod))
                        Call SetPlayerDEF(index, (GetPlayerDEF(index) - Item(GetPlayerInvItemNum(index, Temp)).DEFmod))
                        Call SetPlayerMAGI(index, (GetPlayerMAGI(index) - Item(GetPlayerInvItemNum(index, Temp)).MAGImod))
                        Call SetPlayerWeaponSlot(index, 0)
                        Call SendWornEquipment(index)
                        Call SetPlayerSprite4(index, 1000)
                        Call SendPlayerData(index)
                    End If
                    MapItem(GetPlayerMap(index), i).Dur = GetPlayerInvItemDur(index, InvNum)
                    
                Case ITEM_TYPE_HELMET
                    If InvNum = GetPlayerHelmetSlot(index) Then
                        Temp = GetPlayerHelmetSlot(index)
                        Call SetPlayerSTR(index, (GetPlayerSTR(index) - Item(GetPlayerInvItemNum(index, Temp)).STRmod))
                        Call SetPlayerDEF(index, (GetPlayerDEF(index) - Item(GetPlayerInvItemNum(index, Temp)).DEFmod))
                        Call SetPlayerMAGI(index, (GetPlayerMAGI(index) - Item(GetPlayerInvItemNum(index, Temp)).MAGImod))
                        Call SetPlayerHelmetSlot(index, 0)
                        'Call SetPlayerSprite3(Index, 4000)
                        Call SendWornEquipment(index)
                        Call SendPlayerData(index)
                        
                    End If
                    MapItem(GetPlayerMap(index), i).Dur = GetPlayerInvItemDur(index, InvNum)
                                    
                Case ITEM_TYPE_SHIELD
                    If InvNum = GetPlayerShieldSlot(index) Then
                    Temp = GetPlayerShieldSlot(index)
                        Call SetPlayerSTR(index, (GetPlayerSTR(index) - Item(GetPlayerInvItemNum(index, Temp)).STRmod))
                        Call SetPlayerDEF(index, (GetPlayerDEF(index) - Item(GetPlayerInvItemNum(index, Temp)).DEFmod))
                        Call SetPlayerMAGI(index, (GetPlayerMAGI(index) - Item(GetPlayerInvItemNum(index, Temp)).MAGImod))
                        Call SetPlayerShieldSlot(index, 0)
                        Call SetPlayerSprite3(index, 4000)
                        Call SetPlayerShieldSlot(index, 0)
                        Call SendWornEquipment(index)
                        Call SendPlayerData(index)
                    End If
                    MapItem(GetPlayerMap(index), i).Dur = GetPlayerInvItemDur(index, InvNum)
                End Select
                                
            MapItem(GetPlayerMap(index), i).Num = GetPlayerInvItemNum(index, InvNum)
            MapItem(GetPlayerMap(index), i).X = GetPlayerX(index)
            MapItem(GetPlayerMap(index), i).y = GetPlayerY(index)
                        
            If Item(GetPlayerInvItemNum(index, InvNum)).Type = ITEM_TYPE_CURRENCY Then
                ' Check if its more then they have and if so drop it all
                If Ammount >= GetPlayerInvItemValue(index, InvNum) Then
                    MapItem(GetPlayerMap(index), i).Value = GetPlayerInvItemValue(index, InvNum)
                    Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " drops " & GetPlayerInvItemValue(index, InvNum) & " " & Trim(Item(GetPlayerInvItemNum(index, InvNum)).Name) & ".", Yellow)
                    Call SetPlayerInvItemNum(index, InvNum, 0)
                    Call SetPlayerInvItemValue(index, InvNum, 0)
                    Call SetPlayerInvItemDur(index, InvNum, 0)
                Else
                    MapItem(GetPlayerMap(index), i).Value = Ammount
                    Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " drops " & Ammount & " " & Trim(Item(GetPlayerInvItemNum(index, InvNum)).Name) & ".", Yellow)
                    Call SetPlayerInvItemValue(index, InvNum, GetPlayerInvItemValue(index, InvNum) - Ammount)
                End If
            Else
                ' Its not a currency object so this is easy
                MapItem(GetPlayerMap(index), i).Value = 0
                If Item(GetPlayerInvItemNum(index, InvNum)).Type >= ITEM_TYPE_WEAPON And Item(GetPlayerInvItemNum(index, InvNum)).Type <= ITEM_TYPE_SHIELD Then
                    Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " drops a " & Trim(Item(GetPlayerInvItemNum(index, InvNum)).Name) & ".", Yellow)
                Else
                    Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " drops a " & Trim(Item(GetPlayerInvItemNum(index, InvNum)).Name) & ".", Yellow)
                End If
                
                Call SetPlayerInvItemNum(index, InvNum, 0)
                Call SetPlayerInvItemValue(index, InvNum, 0)
                Call SetPlayerInvItemDur(index, InvNum, 0)
            End If
                                        
            ' Send inventory update
            Call SendInventoryUpdate(index, InvNum)
            ' Spawn the item before we set the num or we'll get a different free map item slot
            Call SpawnItemSlot(i, MapItem(GetPlayerMap(index), i).Num, Ammount, MapItem(GetPlayerMap(index), i).Dur, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index))
        Else
            Call PlayerMsgCombat(index, "To many items already on the ground.", BrightRed)
        End If
    End If
End Sub
Sub PlayerDeathDropItem(ByVal index As Long, ByVal InvNum As Long, ByVal Ammount As Long)


Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(index) = False Or InvNum <= 0 Or InvNum > MAX_INV Then
        Exit Sub
    End If
    
    If (GetPlayerInvItemNum(index, InvNum) > 0) And (GetPlayerInvItemNum(index, InvNum) <= MAX_ITEMS) Then
            
            ' Check to see if its any sort of ArmorSlot/WeaponSlot
            Select Case Item(GetPlayerInvItemNum(index, InvNum)).Type
                Case ITEM_TYPE_ARMOR
                    If InvNum = GetPlayerArmorSlot(index) Then
                    Temp = GetPlayerArmorSlot(index)
                        Call SetPlayerSTR(index, (GetPlayerSTR(index) - Item(GetPlayerInvItemNum(index, Temp)).STRmod))
                        Call SetPlayerDEF(index, (GetPlayerDEF(index) - Item(GetPlayerInvItemNum(index, Temp)).DEFmod))
                        Call SetPlayerMAGI(index, (GetPlayerMAGI(index) - Item(GetPlayerInvItemNum(index, Temp)).MAGImod))
                        Call SetPlayerArmorSlot(index, 0)
                        Call SendWornEquipment(index)
                        If GetPlayerSex(index) = 0 Then
                        Call SetPlayerSprite(index, 0)
                        Else
                        Call SetPlayerSprite(index, 1)
                        End If
                        Call SendPlayerData(index)
                    End If
                   
                Case ITEM_TYPE_WEAPON
                    If InvNum = GetPlayerWeaponSlot(index) Then
                    Temp = GetPlayerWeaponSlot(index)
                        Call SetPlayerSTR(index, (GetPlayerSTR(index) - Item(GetPlayerInvItemNum(index, Temp)).STRmod))
                        Call SetPlayerDEF(index, (GetPlayerDEF(index) - Item(GetPlayerInvItemNum(index, Temp)).DEFmod))
                        Call SetPlayerMAGI(index, (GetPlayerMAGI(index) - Item(GetPlayerInvItemNum(index, Temp)).MAGImod))
                        Call SetPlayerWeaponSlot(index, 0)
                        Call SendWornEquipment(index)
                        Call SetPlayerSprite4(index, 1000)
                        Call SendPlayerData(index)
                    End If
                    'MapItem(GetPlayerMap(Index), i).Dur = GetPlayerInvItemDur(Index, InvNum)
                    
                Case ITEM_TYPE_HELMET
                    If InvNum = GetPlayerHelmetSlot(index) Then
                        Temp = GetPlayerHelmetSlot(index)
                        Call SetPlayerSTR(index, (GetPlayerSTR(index) - Item(GetPlayerInvItemNum(index, Temp)).STRmod))
                        Call SetPlayerDEF(index, (GetPlayerDEF(index) - Item(GetPlayerInvItemNum(index, Temp)).DEFmod))
                        Call SetPlayerMAGI(index, (GetPlayerMAGI(index) - Item(GetPlayerInvItemNum(index, Temp)).MAGImod))
                        Call SetPlayerHelmetSlot(index, 0)
                        'Call SetPlayerSprite3(Index, 4000)
                        Call SendWornEquipment(index)
                        Call SendPlayerData(index)
                        
                    End If
                    'MapItem(GetPlayerMap(Index), i).Dur = GetPlayerInvItemDur(Index, InvNum)
                                    
                Case ITEM_TYPE_SHIELD
                    If InvNum = GetPlayerShieldSlot(index) Then
                    Temp = GetPlayerShieldSlot(index)
                        Call SetPlayerSTR(index, (GetPlayerSTR(index) - Item(GetPlayerInvItemNum(index, Temp)).STRmod))
                        Call SetPlayerDEF(index, (GetPlayerDEF(index) - Item(GetPlayerInvItemNum(index, Temp)).DEFmod))
                        Call SetPlayerMAGI(index, (GetPlayerMAGI(index) - Item(GetPlayerInvItemNum(index, Temp)).MAGImod))
                        Call SetPlayerShieldSlot(index, 0)
                        Call SetPlayerSprite3(index, 4000)
                        Call SetPlayerShieldSlot(index, 0)
                        Call SendWornEquipment(index)
                        Call SendPlayerData(index)
                    End If
                    'MapItem(GetPlayerMap(Index), i).Dur = GetPlayerInvItemDur(Index, InvNum)
                End Select
                                
            i = GetPlayerInvItemNum(index, InvNum)
            'MapItem(GetPlayerMap(Index), i).X = GetPlayerX(Index)
            'MapItem(GetPlayerMap(Index), i).y = GetPlayerY(Index)
                        
            If Item(GetPlayerInvItemNum(index, InvNum)).Type = ITEM_TYPE_CURRENCY Then
                ' Check if its more then they have and if so drop it all
                If Ammount >= GetPlayerInvItemValue(index, InvNum) Then
                   ' MapItem(GetPlayerMap(Index), i).Value = GetPlayerInvItemValue(Index, InvNum)
                  '  Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " drops " & GetPlayerInvItemValue(Index, InvNum) & " " & Trim(Item(GetPlayerInvItemNum(Index, InvNum)).Name) & ".", Yellow)
                    Call SetPlayerInvItemNum(index, InvNum, 0)
                    Call SetPlayerInvItemValue(index, InvNum, 0)
                    Call SetPlayerInvItemDur(index, InvNum, 0)
                Else
                  '  MapItem(GetPlayerMap(Index), i).Value = Ammount
                   ' Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " drops " & Ammount & " " & Trim(Item(GetPlayerInvItemNum(Index, InvNum)).Name) & ".", Yellow)
                    Call SetPlayerInvItemValue(index, InvNum, GetPlayerInvItemValue(index, InvNum) - Ammount)
                End If
            Else
                ' Its not a currency object so this is easy
               ' MapItem(GetPlayerMap(Index), i).Value = 0
                If Item(GetPlayerInvItemNum(index, InvNum)).Type >= ITEM_TYPE_WEAPON And Item(GetPlayerInvItemNum(index, InvNum)).Type <= ITEM_TYPE_SHIELD Then
                   ' Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " drops a " & Trim(Item(GetPlayerInvItemNum(Index, InvNum)).Name) & ".", Yellow)
                Else
                   ' Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " drops a " & Trim(Item(GetPlayerInvItemNum(Index, InvNum)).Name) & ".", Yellow)
                End If
                
                Call SetPlayerInvItemNum(index, InvNum, 0)
                Call SetPlayerInvItemValue(index, InvNum, 0)
                Call SetPlayerInvItemDur(index, InvNum, 0)
            End If
                                        
            ' Send inventory update
            Call SendInventoryUpdate(index, InvNum)
            Call GiveBankItem(index, i, 1)
            ' Spawn the item before we set the num or we'll get a different free map item slot
           ' Call SpawnItemSlot(i, MapItem(GetPlayerMap(Index), i).Num, Ammount, MapItem(GetPlayerMap(Index), i).Dur, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index))
        Else
            Call PlayerMsgCombat(index, "To many items in bank.", BrightRed)
        End If
End Sub

Sub SpawnNpc(ByVal MapNpcNum As Long, ByVal MapNum As Long)


Dim Packet As String
Dim NpcNum As Long
Dim i As Long, X As Long, y As Long
Dim Spawned As Boolean

    ' Check for subscript out of range
    If MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If
    
    Spawned = False
    
    NpcNum = MAP(MapNum).Npc(MapNpcNum)
    If NpcNum > 0 Then
        MapNpc(MapNum, MapNpcNum).Num = NpcNum
        MapNpc(MapNum, MapNpcNum).Target = 0
        
        MapNpc(MapNum, MapNpcNum).HP = GetNpcMaxHP(NpcNum)
        MapNpc(MapNum, MapNpcNum).MP = GetNpcMaxMP(NpcNum)
        MapNpc(MapNum, MapNpcNum).SP = GetNpcMaxSP(NpcNum)
                
        MapNpc(MapNum, MapNpcNum).Dir = Int(Rnd * 4)
        
        ' Well try 100 times to randomly place the sprite
        For i = 1 To 100
            X = Int(Rnd * MAX_MAPX)
            y = Int(Rnd * MAX_MAPY)
            
            ' Check if the tile is walkable
            If MAP(MapNum).Tile(X, y).Type = TILE_TYPE_WALKABLE Then
                MapNpc(MapNum, MapNpcNum).X = X
                MapNpc(MapNum, MapNpcNum).y = y
                Spawned = True
                Exit For
            End If
        Next i
        
        ' Didn't spawn, so now we'll just try to find a free tile
        If Not Spawned Then
            For y = 0 To MAX_MAPY
                For X = 0 To MAX_MAPX
                    If MAP(MapNum).Tile(X, y).Type = TILE_TYPE_WALKABLE Then
                        MapNpc(MapNum, MapNpcNum).X = X
                        MapNpc(MapNum, MapNpcNum).y = y
                        Spawned = True
                    End If
                Next X
            Next y
        End If
             
        ' If we suceeded in spawning then send it to everyone
        If Spawned Then
            Packet = "SPAWNNPC" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Num & SEP_CHAR & MapNpc(MapNum, MapNpcNum).X & SEP_CHAR & MapNpc(MapNum, MapNpcNum).y & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Dir & SEP_CHAR & END_CHAR
            Call SendDataToMap(MapNum, Packet)
        End If
    End If
End Sub

Sub SpawnMapNpcs(ByVal MapNum As Long)


Dim i As Long

    For i = 1 To MAX_MAP_NPCS
        Call SpawnNpc(i, MapNum)
    Next i
End Sub

Sub SpawnAllMapNpcs()
Dim i As Long

    For i = 1 To MAX_MAPS
        Call SpawnMapNpcs(i)
    Next i
End Sub

Function CanAttackPlayer(ByVal Attacker As Long, ByVal Victim As Long) As Boolean


    CanAttackPlayer = False
    
    ' Check for subscript out of range
    If IsPlaying(Attacker) = False Or IsPlaying(Victim) = False Then
        Exit Function
    End If

    ' Make sure they have more then 0 hp
    If GetPlayerHP(Victim) <= 0 Then
        Exit Function
    End If
    
    ' Make sure we dont attack the player if they are switching maps
    If player(Victim).GettingMap = YES Then
        Exit Function
    End If
    
    
    ' Make sure they are on the same map
    If (GetPlayerMap(Attacker) = GetPlayerMap(Victim)) And (GetTickCount > player(Attacker).AttackTimer + 1000) Then
       ' If GetPlayerLevel(Attacker) + 6 >= GetPlayerLevel(Victim) And GetPlayerLevel(Attacker) - 6 <= GetPlayerLevel(Victim) Then
                                    
                                
        ' Check if at same coordinates
        Select Case GetPlayerDir(Attacker)
            Case DIR_UP
            
                If MAP(GetPlayerMap(Attacker)).BootY = 0 And (GetPlayerY(Victim) + 1 = GetPlayerY(Attacker)) And (GetPlayerX(Victim) = GetPlayerX(Attacker)) Then
                    ' Check to make sure that they dont have access
                            ' Check if map is attackable
                            If MAP(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NONE Or GetPlayerPK(Victim) = YES Or (MAP(GetPlayerMap(Attacker)).Tile(GetPlayerX(Attacker), GetPlayerY(Attacker)).Type = TILE_TYPE_ARENA And MAP(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).Type = TILE_TYPE_ARENA) Or (GetPlayerFaction(Attacker) > 0 And GetPlayerFaction(Victim) > 0 And GetPlayerFaction(Attacker) <> GetPlayerFaction(Victim)) Or (GetPlayerAnonymous(Attacker) > 0 And GetPlayerAnonymous(Victim) > 0 And GetPlayerGuildRank(Attacker) > 0 And GetPlayerGuildRank(Victim) > 0 And GetPlayerGuild(Attacker) <> "" And GetPlayerGuild(Victim) <> "") Then
                                ' Make sure they are high enough level
                                If GetPlayerLevel(Attacker) < 10 Then
                                    Call PlayerMsgCombat(Attacker, "You are below level 10, you cannot attack another player yet!", BrightRed)
                                Else
                                    If GetPlayerLevel(Victim) < 10 Then
                                        Call PlayerMsgCombat(Attacker, GetPlayerName(Victim) & " is below level 10, you cannot attack this player yet!", BrightRed)
                                    Else
                                    
                                        CanAttackPlayer = True
                                         
                                    End If
                                    End If
                               
                            Else
                                Call PlayerMsgCombat(Attacker, "This is a safe zone!", BrightRed)
                        
                End If
        End If
            Case DIR_DOWN
            
                If MAP(GetPlayerMap(Attacker)).BootY = 0 And (GetPlayerY(Victim) - 1 = GetPlayerY(Attacker)) And (GetPlayerX(Victim) = GetPlayerX(Attacker)) Then
                    ' Check to make sure that they dont have access
                   
                            ' Check if map is attackable
                            If MAP(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NONE Or GetPlayerPK(Victim) = YES Or (MAP(GetPlayerMap(Attacker)).Tile(GetPlayerX(Attacker), GetPlayerY(Attacker)).Type = TILE_TYPE_ARENA And MAP(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).Type = TILE_TYPE_ARENA) Or (GetPlayerFaction(Attacker) > 0 And GetPlayerFaction(Victim) > 0 And GetPlayerFaction(Attacker) <> GetPlayerFaction(Victim)) Or (GetPlayerAnonymous(Attacker) > 0 And GetPlayerAnonymous(Victim) > 0 And GetPlayerGuildRank(Attacker) > 0 And GetPlayerGuildRank(Victim) > 0 And GetPlayerGuild(Attacker) <> "" And GetPlayerGuild(Victim) <> "") Then
                                ' Make sure they are high enough level
                                
                                If GetPlayerLevel(Attacker) < 10 Then
                                    Call PlayerMsgCombat(Attacker, "You are below level 10, you cannot attack another player yet!", BrightRed)
                                Else
                                    If GetPlayerLevel(Victim) < 10 Then
                                        Call PlayerMsgCombat(Attacker, GetPlayerName(Victim) & " is below level 10, you cannot attack this player yet!", BrightRed)
                                    Else
                                   
                                        CanAttackPlayer = True
                                      
                                    End If
                                    End If
                                     
                                
                            Else
                                Call PlayerMsgCombat(Attacker, "This is a safe zone!", BrightRed)
                            End If
                        
                End If
      
            Case DIR_LEFT
            
                If MAP(GetPlayerMap(Attacker)).BootY = 0 And (GetPlayerY(Victim) = GetPlayerY(Attacker)) And (GetPlayerX(Victim) + 1 = GetPlayerX(Attacker)) Then
                    ' Check to make sure that they dont have access
                    
                            ' Check if map is attackable
                            If MAP(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NONE Or GetPlayerPK(Victim) = YES Or (MAP(GetPlayerMap(Attacker)).Tile(GetPlayerX(Attacker), GetPlayerY(Attacker)).Type = TILE_TYPE_ARENA And MAP(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).Type = TILE_TYPE_ARENA) Or (GetPlayerFaction(Attacker) > 0 And GetPlayerFaction(Victim) > 0 And GetPlayerFaction(Attacker) <> GetPlayerFaction(Victim)) Or (GetPlayerAnonymous(Attacker) > 0 And GetPlayerAnonymous(Victim) > 0 And GetPlayerGuildRank(Attacker) > 0 And GetPlayerGuildRank(Victim) > 0 And GetPlayerGuild(Attacker) <> "" And GetPlayerGuild(Victim) <> "") Then
                                ' Make sure they are high enough level
                                
                                
                                If GetPlayerLevel(Attacker) < 10 Then
                                    Call PlayerMsgCombat(Attacker, "You are below level 10, you cannot attack another player yet!", BrightRed)
                                Else
                                    If GetPlayerLevel(Victim) < 10 Then
                                        Call PlayerMsgCombat(Attacker, GetPlayerName(Victim) & " is below level 10, you cannot attack this player yet!", BrightRed)
                                    Else
                                       CanAttackPlayer = True
                                        
                                    End If
                                End If
                                    
                            Else
                                Call PlayerMsgCombat(Attacker, "This is a safe zone!", BrightRed)
                            End If
                        
                End If

            Case DIR_RIGHT
            
                If MAP(GetPlayerMap(Attacker)).BootY = 0 And (GetPlayerY(Victim) = GetPlayerY(Attacker)) And (GetPlayerX(Victim) - 1 = GetPlayerX(Attacker)) Then
                    ' Check to make sure that they dont have access
                   
                            ' Check if map is attackable
                            If MAP(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NONE Or GetPlayerPK(Victim) = YES Or (MAP(GetPlayerMap(Attacker)).Tile(GetPlayerX(Attacker), GetPlayerY(Attacker)).Type = TILE_TYPE_ARENA And MAP(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).Type = TILE_TYPE_ARENA) Or (GetPlayerFaction(Attacker) > 0 And GetPlayerFaction(Victim) > 0 And GetPlayerFaction(Attacker) <> GetPlayerFaction(Victim)) Or (GetPlayerAnonymous(Attacker) > 0 And GetPlayerAnonymous(Victim) > 0 And GetPlayerGuildRank(Attacker) > 0 And GetPlayerGuildRank(Victim) > 0 And GetPlayerGuild(Attacker) <> "" And GetPlayerGuild(Victim) <> "") Then
                                ' Make sure they are high enough level
                                
                                
                                If GetPlayerLevel(Attacker) < 10 Then
                                    Call PlayerMsgCombat(Attacker, "You are below level 10, you cannot attack another player yet!", BrightRed)
                                Else
                                    If GetPlayerLevel(Victim) < 10 Then
                                        Call PlayerMsgCombat(Attacker, GetPlayerName(Victim) & " is below level 10, you cannot attack this player yet!", BrightRed)
                                    Else
                                        CanAttackPlayer = True
                                       
                                    End If
                                End If
                                
                            Else
                                Call PlayerMsgCombat(Attacker, "This is a safe zone!", BrightRed)
                            End If
                        
                End If
        End Select
       ' Else
                                  '  Call PlayerMsgCombat(Attacker, "To much of a level difference!", BrightRed)
                                'End If
                                
    End If
End Function

Function CanAttackNpc(ByVal Attacker As Long, ByVal MapNpcNum As Long) As Boolean


Dim MapNum As Long, NpcNum As Long, Range As Long, distance As Long
'Call TextAdd(frmServer.txtText, "CanAttack NPC", True)
    CanAttackNpc = False
    'Call TextAdd(frmServer.txtText, "CanAttack2 NPC", True)
    ' Check for subscript out of range
    If IsPlaying(Attacker) = False Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Then
        'Call TextAdd(frmServer.txtText, "CanAttack3 NPC", True)
        Exit Function
        
    End If
    
    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(Attacker), MapNpcNum).Num <= 0 Then
        Exit Function
    End If
    
    MapNum = GetPlayerMap(Attacker)
    NpcNum = MapNpc(MapNum, MapNpcNum).Num
    
    ' Make sure the npc isn't already dead
    If MapNpc(MapNum, MapNpcNum).HP <= 0 Then
        Exit Function
    End If
'Range = Item(player(Attacker).Char(player(Attacker).charnum).WeaponSlot).Data4
'MsgBox (STR(Range) + "   " + STR(Attacker) + "   " + STR(player(Attacker).charnum) + "   " + STR(player(Attacker).Char(player(Attacker).charnum).WeaponSlot))
'For distance = Range To 1 Step -1
'    ' Make sure they are on the same map
'    If IsPlaying(Attacker) Then
'        If NpcNum > 0 And GetTickCount > player(Attacker).AttackTimer + 950 Then
'            ' Check if at same coordinates
'            Select Case GetPlayerDir(Attacker)
'            Case DIR_UP
'              If MAP(GetPlayerMap(Attacker)).Tile(GetPlayerX(Attacker), GetPlayerY(Attacker) - distance).Type <> TILE_TYPE_BLOCKED Then
'               If MAP(GetPlayerMap(Attacker)).Tile(GetPlayerX(Attacker), GetPlayerY(Attacker) - distance).Type <> TILE_TYPE_KEY Then
'                If (MapNpc(MapNum, MapNpcNum).y + distance = GetPlayerY(Attacker)) And (MapNpc(MapNum, MapNpcNum).X = GetPlayerX(Attacker)) Then
'                        If Npc(NpcNum).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(NpcNum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then
'                            CanAttackNpc = True
'                        Else
'                            Call PlayerMsg(Attacker, "A " & Trim(Npc(NpcNum).Name) & " says, '" & Trim(Npc(NpcNum).AttackSay) & "' to you.", SayColor)
'                            'Call PlayerMsg(Attacker, "You cannot attack a " & Trim(Npc(NpcNum).Name) & "!", BrightBlue)
'                        End If
'                       End If
'                    End If
'            End If

'                Case DIR_DOWN
'                If MAP(GetPlayerMap(Attacker)).Tile(GetPlayerX(Attacker), GetPlayerY(Attacker) + distance).Type <> TILE_TYPE_BLOCKED Then
'                If MAP(GetPlayerMap(Attacker)).Tile(GetPlayerX(Attacker), GetPlayerY(Attacker) + distance).Type <> TILE_TYPE_KEY Then
'                    If (MapNpc(MapNum, MapNpcNum).y - distance = GetPlayerY(Attacker)) And (MapNpc(MapNum, MapNpcNum).X = GetPlayerX(Attacker)) Then
'                        If Npc(NpcNum).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(NpcNum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then
'                            CanAttackNpc = True
'                        Else
'                            Call PlayerMsg(Attacker, "A " & Trim(Npc(NpcNum).Name) & " says, '" & Trim(Npc(NpcNum).AttackSay) & "' to you.", SayColor)
'                            'Call PlayerMsg(Attacker, "You cannot attack a " & Trim(Npc(NpcNum).Name) & "!", BrightBlue)
'                        End If
'                    End If
'               End If
'               End If

'                Case DIR_LEFT
'                If MAP(GetPlayerMap(Attacker)).Tile(GetPlayerX(Attacker) - distance, GetPlayerY(Attacker)).Type <> TILE_TYPE_BLOCKED Then
'                If MAP(GetPlayerMap(Attacker)).Tile(GetPlayerX(Attacker) - distance, GetPlayerY(Attacker)).Type <> TILE_TYPE_KEY Then
'                   If (MapNpc(MapNum, MapNpcNum).y = GetPlayerY(Attacker)) And (MapNpc(MapNum, MapNpcNum).X + distance = GetPlayerX(Attacker)) Then
'                       If Npc(NpcNum).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(NpcNum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then
'                           CanAttackNpc = True
'                       Else
'                         Call PlayerMsg(Attacker, "A " & Trim(Npc(NpcNum).Name) & " says, '" & Trim(Npc(NpcNum).AttackSay) & "' to you.", SayColor)
'                         'Call PlayerMsg(Attacker, "You cannot attack a " & Trim(Npc(NpcNum).Name) & "!", BrightBlue)
'                       End If
'                    End If
'                    End If
'                End If
     
'               Case DIR_RIGHT
'
'                If MAP(GetPlayerMap(Attacker)).Tile(GetPlayerX(Attacker) + distance, GetPlayerY(Attacker)).Type <> TILE_TYPE_BLOCKED Then
'                If MAP(GetPlayerMap(Attacker)).Tile(GetPlayerX(Attacker) + distance, GetPlayerY(Attacker)).Type <> TILE_TYPE_KEY Then
'                    If (MapNpc(MapNum, MapNpcNum).y = GetPlayerY(Attacker)) And (MapNpc(MapNum, MapNpcNum).X - distance = GetPlayerX(Attacker)) Then
'                        If Npc(NpcNum).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(NpcNum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then
'                            CanAttackNpc = True
'                       Else
'                            Call PlayerMsg(Attacker, "A " & Trim(Npc(NpcNum).Name) & " says, '" & Trim(Npc(NpcNum).AttackSay) & "' to you.", SayColor)
'                            'Call PlayerMsg(Attacker, "You cannot attack a " & Trim(Npc(NpcNum).Name) & "!", BrightBlue)
'                        End If
'                    End If
'                    End If
'                    End If
'            End Select
'        End If
'    End If
'Next distance

    ' Make sure they are on the same map
        If IsPlaying(Attacker) Then
        If NpcNum > 0 And GetTickCount > player(Attacker).AttackTimer + 950 Then
            ' Check if at same coordinates
            Select Case GetPlayerDir(Attacker)
                Case DIR_UP
                    If (MapNpc(MapNum, MapNpcNum).y + 1 = GetPlayerY(Attacker)) And (MapNpc(MapNum, MapNpcNum).X = GetPlayerX(Attacker)) Then
                        If Npc(NpcNum).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(NpcNum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then
                            CanAttackNpc = True
                        Else
                            Call PlayerMsg(Attacker, "A " & Trim(Npc(NpcNum).Name) & " says, '" & Trim(Npc(NpcNum).AttackSay) & "' to you.", SayColor)
                            'Call PlayerMsg(Attacker, "You cannot attack a " & Trim(Npc(NpcNum).Name) & "!", BrightBlue)
                        End If
                    End If

                Case DIR_DOWN
                    If (MapNpc(MapNum, MapNpcNum).y - 1 = GetPlayerY(Attacker)) And (MapNpc(MapNum, MapNpcNum).X = GetPlayerX(Attacker)) Then
                        If Npc(NpcNum).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(NpcNum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then
                            CanAttackNpc = True
                        Else
                            Call PlayerMsg(Attacker, "A " & Trim(Npc(NpcNum).Name) & " says, '" & Trim(Npc(NpcNum).AttackSay) & "' to you.", SayColor)
                            'Call PlayerMsg(Attacker, "You cannot attack a " & Trim(Npc(NpcNum).Name) & "!", BrightBlue)
                        End If
                    End If

                Case DIR_LEFT
                    If (MapNpc(MapNum, MapNpcNum).y = GetPlayerY(Attacker)) And (MapNpc(MapNum, MapNpcNum).X + 1 = GetPlayerX(Attacker)) Then
                        If Npc(NpcNum).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(NpcNum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then
                            CanAttackNpc = True
                        Else
                            Call PlayerMsg(Attacker, "A " & Trim(Npc(NpcNum).Name) & " says, '" & Trim(Npc(NpcNum).AttackSay) & "' to you.", SayColor)
                            'Call PlayerMsg(Attacker, "You cannot attack a " & Trim(Npc(NpcNum).Name) & "!", BrightBlue)
                        End If
                    End If

                Case DIR_RIGHT
                    If (MapNpc(MapNum, MapNpcNum).y = GetPlayerY(Attacker)) And (MapNpc(MapNum, MapNpcNum).X - 1 = GetPlayerX(Attacker)) Then
                        If Npc(NpcNum).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(NpcNum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then
                            CanAttackNpc = True
                        Else
                            Call PlayerMsg(Attacker, "A " & Trim(Npc(NpcNum).Name) & " says, '" & Trim(Npc(NpcNum).AttackSay) & "' to you.", SayColor)
                            'Call PlayerMsg(Attacker, "You cannot attack a " & Trim(Npc(NpcNum).Name) & "!", BrightBlue)
                        End If
                    End If
            End Select
        End If
    End If

End Function
     


Function CanNpcAttackPlayer(ByVal MapNpcNum As Long, ByVal index As Long) As Boolean


Dim MapNum As Long, NpcNum As Long
    
    CanNpcAttackPlayer = False
    
    ' Check for subscript out of range
    If MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or IsPlaying(index) = False Then
        Exit Function
    End If
    
    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(index), MapNpcNum).Num <= 0 Then
        Exit Function
    End If
    
    MapNum = GetPlayerMap(index)
    NpcNum = MapNpc(MapNum, MapNpcNum).Num
    
    ' Make sure the npc isn't already dead
    If MapNpc(MapNum, MapNpcNum).HP <= 0 Then
        Exit Function
    End If
    
    ' Make sure npcs dont attack more then once a second
    If GetTickCount < MapNpc(MapNum, MapNpcNum).AttackTimer + 1000 Then
        Exit Function
    End If
    
    ' Make sure we dont attack the player if they are switching maps
    If player(index).GettingMap = YES Then
        Exit Function
    End If
    
    MapNpc(MapNum, MapNpcNum).AttackTimer = GetTickCount
    
    ' Make sure they are on the same map
    If IsPlaying(index) Then
        If NpcNum > 0 Then
            ' Check if at same coordinates
            If (GetPlayerY(index) + 1 = MapNpc(MapNum, MapNpcNum).y) And (GetPlayerX(index) = MapNpc(MapNum, MapNpcNum).X) Then
                CanNpcAttackPlayer = True
            Else
                If (GetPlayerY(index) - 1 = MapNpc(MapNum, MapNpcNum).y) And (GetPlayerX(index) = MapNpc(MapNum, MapNpcNum).X) Then
                    CanNpcAttackPlayer = True
                Else
                    If (GetPlayerY(index) = MapNpc(MapNum, MapNpcNum).y) And (GetPlayerX(index) + 1 = MapNpc(MapNum, MapNpcNum).X) Then
                        CanNpcAttackPlayer = True
                    Else
                        If (GetPlayerY(index) = MapNpc(MapNum, MapNpcNum).y) And (GetPlayerX(index) - 1 = MapNpc(MapNum, MapNpcNum).X) Then
                            CanNpcAttackPlayer = True
                        End If
                    End If
                End If
            End If

'            Select Case MapNpc(MapNum, MapNpcNum).Dir
'                Case DIR_UP
'                    If (GetPlayerY(Index) + 1 = MapNpc(MapNum, MapNpcNum).y) And (GetPlayerX(Index) = MapNpc(MapNum, MapNpcNum).x) Then
'                        CanNpcAttackPlayer = True
'                    End If
'
'                Case DIR_DOWN
'                    If (GetPlayerY(Index) - 1 = MapNpc(MapNum, MapNpcNum).y) And (GetPlayerX(Index) = MapNpc(MapNum, MapNpcNum).x) Then
'                        CanNpcAttackPlayer = True
'                    End If
'
'                Case DIR_LEFT
'                    If (GetPlayerY(Index) = MapNpc(MapNum, MapNpcNum).y) And (GetPlayerX(Index) + 1 = MapNpc(MapNum, MapNpcNum).x) Then
'                        CanNpcAttackPlayer = True
'                    End If
'
'                Case DIR_RIGHT
'                    If (GetPlayerY(Index) = MapNpc(MapNum, MapNpcNum).y) And (GetPlayerX(Index) - 1 = MapNpc(MapNum, MapNpcNum).x) Then
'                        CanNpcAttackPlayer = True
'                    End If
'            End Select
        End If
    End If
End Function

Sub AttackPlayer(ByVal Attacker As Long, ByVal Victim As Long, ByVal Damage As Long, ByVal spellnameR As String)


Dim exp As Long
Dim n As Long
Dim i As Long
Dim t As Long
Dim NameR As String

    ' Check for subscript out of range
    If IsPlaying(Attacker) = False Or IsPlaying(Victim) = False Or Damage < 0 Then
        Exit Sub
    End If
    
   ' If GetPlayerLevel(Attacker) + 6 >= GetPlayerLevel(Victim) And GetPlayerLevel(Attacker) - 6 <= GetPlayerLevel(Victim) Then
    
    
    ' Check for weapon
    If GetPlayerWeaponSlot(Attacker) > 0 And spellnameR = "z" Then
        n = GetPlayerInvItemNum(Attacker, GetPlayerWeaponSlot(Attacker))
        Call WeapSound(Attacker, n)
        Call WeapSound(Victim, n)
        
    Else
        n = 0
    End If
    If spellnameR <> "z" Then
    'Call SpellSound(Attacker, spell(spellnameR).Sfx)
    'Call SpellSound(Victim, spell(spellnameR).Sfx)
    End If
    
    
    ' Send this packet so they can see the person attacking
    Call SendDataToMapBut(Attacker, GetPlayerMap(Attacker), "ATTACK" & SEP_CHAR & Attacker & SEP_CHAR & END_CHAR)
    Call DeathSound(Attacker, 88, 88)
    If Damage >= GetPlayerHP(Victim) Then
        ' Set HP to nothing
        Call SetPlayerHP(Victim, 0)
        
        ' Check for a weapon and say damage
        NameR = Trim(Item(n).Name)
        If spellnameR <> "z" Then NameR = spellnameR
        If n = 0 Then
            Call PlayerMsgCombat(Attacker, "You hit " & GetPlayerName(Victim) & " for " & Damage & " hit points.", White)
            Call PlayerMsgCombat(Victim, GetPlayerName(Attacker) & " hit you for " & Damage & " hit points.", BrightRed)
        Else
            Call PlayerMsgCombat(Attacker, "You hit " & GetPlayerName(Victim) & " with a " & NameR & " for " & Damage & " hit points.", White)
            Call PlayerMsgCombat(Victim, GetPlayerName(Attacker) & " hit you with a " & NameR & " for " & Damage & " hit points.", BrightRed)
        End If
        
        ' Player is dead
        Call GlobalMsgCombat(GetPlayerName(Victim) & " has been killed by " & GetPlayerName(Attacker), BrightRed)
        Call PkData(Attacker, Victim)
        ' Drop all worn items by victim
        If GetPlayerWeaponSlot(Victim) > 0 Then
        'tEmp = GetPlayerWeaponSlot(Victim)
                        ' Call SetPlayerSTR(Victim, (GetPlayerSTR(Victim) - Item(GetPlayerInvItemNum(Victim, tEmp)).STRmod))
                        ' Call SetPlayerDEF(Victim, (GetPlayerDEF(Victim) - Item(GetPlayerInvItemNum(Victim, tEmp)).DEFmod))
                        ' Call SetPlayerMAGI(Victim, (GetPlayerMAGI(Victim) - Item(GetPlayerInvItemNum(Victim, tEmp)).MAGImod))
                        
            Call PlayerDeathDropItem(Victim, GetPlayerWeaponSlot(Victim), 0)
            'Call GiveBankItem(Index, GetPlayerWeaponSlot(Victim), 1)
            'Call TakeItem(Index, GetPlayerWeaponSlot(Victim), 1)
            
        End If
        If GetPlayerArmorSlot(Victim) > 0 Then
        'tEmp = GetPlayerArmorSlot(Victim)
                         'Call SetPlayerSTR(Victim, (GetPlayerSTR(Victim) - Item(GetPlayerInvItemNum(Victim, tEmp)).STRmod))
                        ' Call SetPlayerDEF(Victim, (GetPlayerDEF(Victim) - Item(GetPlayerInvItemNum(Victim, tEmp)).DEFmod))
                         'Call SetPlayerMAGI(Victim, (GetPlayerMAGI(Victim) - Item(GetPlayerInvItemNum(Victim, tEmp)).MAGImod))
                        
            Call PlayerDeathDropItem(Victim, GetPlayerArmorSlot(Victim), 0)
        End If
        If GetPlayerHelmetSlot(Victim) > 0 Then
            Call PlayerDeathDropItem(Victim, GetPlayerHelmetSlot(Victim), 0)
        End If
        If GetPlayerShieldSlot(Victim) > 0 Then
            Call PlayerDeathDropItem(Victim, GetPlayerShieldSlot(Victim), 0)
        End If
        
        For t = 1 To MAX_INV
        If GetPlayerInvItemNum(Victim, t) = 1 Then
        Call PlayerMapDropItem(Victim, t, GetPlayerInvItemValue(Victim, t) / 10)
        End If
        Next
        
        
        

        ' Calculate exp to give attacker
        exp = Int(GetPlayerExp(Victim) * 0.01)
        
        ' Make sure we dont get less then 0
        If exp < 0 Then
            exp = 0
        End If
        
        If exp = 0 Then
            Call PlayerMsgCombat(Victim, "You lost no experience points.", BrightRed)
            Call PlayerMsgCombat(Attacker, "You received no experience points from that weak insignificant player.", BrightBlue)
        Else
            Call SetPlayerExp(Victim, GetPlayerExp(Victim) - exp)
            Call PlayerMsgCombat(Victim, "You lost " & exp & " experience points.", BrightRed)
            Call SetPlayerExp(Attacker, GetPlayerExp(Attacker) + exp)
            Call PlayerMsgCombat(Attacker, "You got " & exp & " experience points for killing " & GetPlayerName(Victim) & ".", BrightBlue)
        End If
        
        'play death soudn for map
        Call DeathSound(GetPlayerMap(Victim), GetPlayerClass(Victim), GetPlayerSex(Victim))
                
        ' Warp player away
        'If MAP(GetPlayerMap(Victim)).BootMap > 0 Then
        'Call PlayerWarp(Victim, MAP(GetPlayerMap(Victim)).BootMap, MAP(GetPlayerMap(Victim)).BootX, MAP(GetPlayerMap(Victim)).BootY)
        'Else
        'Call PlayerWarp(Victim, START_MAP, START_X, START_Y)
        'End If
        Call PlayerWarp(Victim, GetPlayerBindMap(Victim), GetPlayerBindX(Victim), GetPlayerBindY(Victim))
        
        ' Restore vitals
        Call SetPlayerHP(Victim, GetPlayerMaxHP(Victim))
        Call SetPlayerMP(Victim, GetPlayerMaxMP(Victim))
        Call SetPlayerSP(Victim, GetPlayerMaxSP(Victim))
        Call SendHP(Victim)
        Call SendMP(Victim)
        Call SendSP(Victim)
                
        ' Check for a level up
        Call CheckPlayerLevelUp(Attacker)
        
        ' Check if target is player who died and if so set target to 0
        If player(Attacker).TargetType = TARGET_TYPE_PLAYER And player(Attacker).Target = Victim Then
            player(Attacker).Target = 0
            player(Attacker).TargetType = 0
        End If
        
        If GetPlayerPK(Victim) = NO Then
            If GetPlayerPK(Attacker) = NO Then
                If GetPlayerFaction(Attacker) > 1 And GetPlayerFaction(Victim) > 1 And GetPlayerFaction(Attacker) <> GetPlayerFaction(Victim) Then
                'do nothing
                Else
                Call SetPlayerPK(Attacker, YES)
                Call GlobalMsgCombat(GetPlayerName(Attacker) & " has been deemed a Player Killer!!!", BrightRed)
                End If
                Call SendPlayerData(Attacker)
                   End If
        Else
            Call SetPlayerPK(Victim, NO)
            Call SendPlayerData(Victim)
            Call GlobalMsgCombat(GetPlayerName(Victim) & " has paid the price for being a Player Killer!!!", BrightRed)
        End If
    Else
        ' Player not dead, just do the damage
        Call SetPlayerHP(Victim, GetPlayerHP(Victim) - Damage)
        Call SendHP(Victim)
        
        ' Check for a weapon and say damage
        If n = 0 Then
            Call PlayerMsgCombat(Attacker, "You hit " & GetPlayerName(Victim) & " for " & Damage & " hit points.", White)
            Call PlayerMsgCombat(Victim, GetPlayerName(Attacker) & " hit you for " & Damage & " hit points.", BrightRed)
        Else
            Call PlayerMsgCombat(Attacker, "You hit " & GetPlayerName(Victim) & " with a " & Trim(Item(n).Name) & " for " & Damage & " hit points.", White)
            Call PlayerMsgCombat(Victim, GetPlayerName(Attacker) & " hit you with a " & Trim(Item(n).Name) & " for " & Damage & " hit points.", BrightRed)
        End If
    End If
    
    ' Reset timer for attacking
    player(Attacker).AttackTimer = GetTickCount
  '  End If
End Sub

Sub NpcAttackPlayer(ByVal MapNpcNum As Long, ByVal Victim As Long, ByVal Damage As Long)


Dim Name As String
Dim exp As Long
Dim t As Long
Dim MapNum As Long

    ' Check for subscript out of range
    If MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or IsPlaying(Victim) = False Or Damage < 0 Then
        Exit Sub
    End If
    
    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(Victim), MapNpcNum).Num <= 0 Then
        Exit Sub
    End If
    
    ' Send this packet so they can see the person attacking
    Call SendDataToMap(GetPlayerMap(Victim), "NPCATTACK" & SEP_CHAR & MapNpcNum & SEP_CHAR & END_CHAR)
    
    MapNum = GetPlayerMap(Victim)
    Name = Trim(Npc(MapNpc(MapNum, MapNpcNum).Num).Name)
    
    If Damage >= GetPlayerHP(Victim) Then
        ' Say damage
        Call PlayerMsgCombat(Victim, "A " & Name & " hit you for " & Damage & " hit points.", BrightRed)
        
        ' Player is dead
        Call GlobalMsgCombat(GetPlayerName(Victim) & " has been killed by a " & Name, BrightRed)
        
        ' Drop all worn items by victim
        If GetPlayerWeaponSlot(Victim) > 0 Then
       ' tEmp = GetPlayerWeaponSlot(Victim)
                         'Call SetPlayerSTR(Victim, (GetPlayerSTR(Victim) - Item(GetPlayerInvItemNum(Victim, tEmp)).STRmod))
                         'Call SetPlayerDEF(Victim, (GetPlayerDEF(Victim) - Item(GetPlayerInvItemNum(Victim, tEmp)).DEFmod))
                         'Call SetPlayerMAGI(Victim, (GetPlayerMAGI(Victim) - Item(GetPlayerInvItemNum(Victim, tEmp)).MAGImod))
                        
            Call PlayerDeathDropItem(Victim, GetPlayerWeaponSlot(Victim), 0)
        End If
        If GetPlayerArmorSlot(Victim) > 0 Then
        'tEmp = GetPlayerArmorSlot(Victim)
                         'Call SetPlayerSTR(Victim, (GetPlayerSTR(Victim) - Item(GetPlayerInvItemNum(Victim, tEmp)).STRmod))
                         'Call SetPlayerDEF(Victim, (GetPlayerDEF(Victim) - Item(GetPlayerInvItemNum(Victim, tEmp)).DEFmod))
                         'Call SetPlayerMAGI(Victim, (GetPlayerMAGI(Victim) - Item(GetPlayerInvItemNum(Victim, tEmp)).MAGImod))
                        
            Call PlayerDeathDropItem(Victim, GetPlayerArmorSlot(Victim), 0)
        End If
        If GetPlayerHelmetSlot(Victim) > 0 Then
            Call PlayerDeathDropItem(Victim, GetPlayerHelmetSlot(Victim), 0)
        End If
        If GetPlayerShieldSlot(Victim) > 0 Then
            Call PlayerDeathDropItem(Victim, GetPlayerShieldSlot(Victim), 0)
        End If
        For t = 1 To MAX_INV
        If GetPlayerInvItemNum(Victim, t) = 1 Then
        Call PlayerMapDropItem(Victim, t, GetPlayerInvItemValue(Victim, t) * 0.1)
        End If
        Next
        
        ' Calculate exp to give attacker
        exp = Int(GetPlayerExp(Victim) / 25)
        
        ' Make sure we dont get less then 0
        If exp < 0 Then
            exp = 0
        End If
        
        If exp = 0 Then
            Call PlayerMsgCombat(Victim, "You lost no experience points.", BrightRed)
        Else
            Call SetPlayerExp(Victim, GetPlayerExp(Victim) - exp)
            Call PlayerMsgCombat(Victim, "You lost " & exp & " experience points.", BrightRed)
        End If
        
        'play death soudn for map
        Call DeathSound(GetPlayerMap(Victim), GetPlayerClass(Victim), GetPlayerSex(Victim))
                
        ' Warp player away
        'If MAP(GetPlayerMap(Victim)).BootMap > 0 Then
        'Call PlayerWarp(Victim, MAP(GetPlayerMap(Victim)).BootMap, MAP(GetPlayerMap(Victim)).BootX, MAP(GetPlayerMap(Victim)).BootY)
        'Else
        'Call PlayerWarp(Victim, START_MAP, START_X, START_Y)
        'End If
        Call PlayerWarp(Victim, GetPlayerBindMap(Victim), GetPlayerBindX(Victim), GetPlayerBindY(Victim))
        ' Restore vitals
        Call SetPlayerHP(Victim, GetPlayerMaxHP(Victim))
        Call SetPlayerMP(Victim, GetPlayerMaxMP(Victim))
        Call SetPlayerSP(Victim, GetPlayerMaxSP(Victim))
        Call SendHP(Victim)
        Call SendMP(Victim)
        Call SendSP(Victim)
        
        ' Set NPC target to 0
        MapNpc(MapNum, MapNpcNum).Target = 0
        
        ' If the player the attacker killed was a pk then take it away
        If GetPlayerPK(Victim) = YES Then
            Call SetPlayerPK(Victim, NO)
            Call SendPlayerData(Victim)
        End If
    Else
        ' Player not dead, just do the damage
        Call SetPlayerHP(Victim, GetPlayerHP(Victim) - Damage)
        Call SendHP(Victim)
        
        ' Say damage
        Call PlayerMsgCombat(Victim, "A " & Name & " hit you for " & Damage & " hit points.", BrightRed)
    End If
End Sub

Sub AttackNpc(ByVal Attacker As Long, ByVal MapNpcNum As Long, ByVal Damage As Long, ByVal spellnameR As String)


    Dim Name As String
Dim exp As Long
Dim n As Long, i As Long
Dim STR As Long, DEF As Long, MapNum As Long, NpcNum As Long
Dim NameR As String
Dim maxlevel As Long

    ' Check for subscript out of range
    If IsPlaying(Attacker) = False Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or Damage < 0 Then
        Exit Sub
    End If
    
    ' Check for weapon
    If GetPlayerWeaponSlot(Attacker) > 0 And spellnameR = "z" Then
        n = GetPlayerInvItemNum(Attacker, GetPlayerWeaponSlot(Attacker))
        Call WeapSound(Attacker, n)
    Else
        n = 0
    End If
    'If spellnameR <> "z" Then Call SpellSound(Attacker, spell(spellnameR).Sfx)
    
    ' Send this packet so they can see the person attacking
    Call SendDataToMapBut(Attacker, GetPlayerMap(Attacker), "ATTACK" & SEP_CHAR & Attacker & SEP_CHAR & END_CHAR)
    
    MapNum = GetPlayerMap(Attacker)
    NpcNum = MapNpc(MapNum, MapNpcNum).Num
    Name = Trim(Npc(NpcNum).Name)
        
    If Damage >= MapNpc(MapNum, MapNpcNum).HP Then
        ' Check for a weapon and say damage
        NameR = Trim(Item(n).Name)
        If spellnameR <> "z" Then NameR = spellnameR
        If n = 0 Then
            Call PlayerMsgCombat(Attacker, "You hit a " & Name & " for " & Damage & " hit points, killing it.", BrightRed)
        Else
            Call PlayerMsgCombat(Attacker, "You hit a " & Name & " with a " & NameR & " for " & Damage & " hit points, killing it.", BrightRed)
        End If
                        
        ' Calculate exp to give attacker
        STR = Npc(NpcNum).STR
        DEF = Npc(MapNpc(GetPlayerMap(Attacker), MapNpcNum).Num).DEF
        exp = STR * 6.1
        maxlevel = (DEF - (DEF * 0.25)) * 1.2
        If maxlevel < 5 Then
        maxlevel = 5
        End If
        Dim levdif As Long
        levdif = (maxlevel - GetPlayerLevel(Attacker)) + 2
        
        If levdif < 0 Then
        exp = exp / (GetPlayerLevel(Attacker) / 3)
        End If
        
        If exp < 0 Then exp = 0

        
        
        
        
        
        'If maxlevel < GetPlayerLevel(Attacker) Then
        'exp = exp / GetPlayerLevel(Attacker)
        'End If
        
        
        
        ' Make sure we dont get less then 0
        If exp < 0 Then
            exp = 0
        End If
        
        ' Check if in party, if so divide the exp up by 2
        If player(Attacker).InParty = NO Then
            Call SetPlayerExp(Attacker, GetPlayerExp(Attacker) + exp)
            If exp > 0 Then
            Call PlayerMsgCombat(Attacker, "You have gained " & exp & " experience points.", BrightBlue)
            Else
            Call PlayerMsgCombat(Attacker, "You gain NO Experiance for this fight.", BrightBlue)
            End If
            Call NPCSound(MapNum, NpcNum)
            Call SendDataToMap(GetPlayerMap(Attacker), "SPELLGFX2" & SEP_CHAR & NpcNum & SEP_CHAR & 0 & SEP_CHAR & END_CHAR)
        Else
            Call NPCSound(MapNum, NpcNum)
            
    If player(Attacker).Party > 0 Then
    If Parties(player(Attacker).Party).NumParty = 4 Then
    exp = (exp * 3.25) / 4
    ElseIf Parties(player(Attacker).Party).NumParty = 3 Then
    exp = (exp * 2.5) / 3
    ElseIf Parties(player(Attacker).Party).NumParty = 2 Then
    exp = (exp * 1.75) / 2
    End If
    End If
    
    
    
            
            
            
            
            If exp < 0 Then
                exp = 1
            End If
            If player(Attacker).Party > 0 Then
            n = Parties(player(Attacker).Party).Player1
            If n > 0 Then
            'If GetPlayerMap(Attacker) = GetPlayerMap(n) Then
                Call SetPlayerExp(n, GetPlayerExp(n) + exp)
                Call PlayerMsgCombat(n, "You have gained " & exp & " party experience points.", BrightBlue)
            'End If
            End If
            n = Parties(player(Attacker).Party).Player2
            If n > 0 Then
            'If GetPlayerMap(Attacker) = GetPlayerMap(n) Then
                Call SetPlayerExp(n, GetPlayerExp(n) + exp)
                Call PlayerMsgCombat(n, "You have gained " & exp & " party experience points.", BrightBlue)
            'End If
            End If
            n = Parties(player(Attacker).Party).Player3
            If n > 0 Then
            'If GetPlayerMap(Attacker) = GetPlayerMap(n) Then
                Call SetPlayerExp(n, GetPlayerExp(n) + exp)
                Call PlayerMsgCombat(n, "You have gained " & exp & " party experience points.", BrightBlue)
            'End If
            End If
             n = Parties(player(Attacker).Party).Player4
            If n > 0 Then
            'If GetPlayerMap(Attacker) = GetPlayerMap(n) Then
                Call SetPlayerExp(n, GetPlayerExp(n) + exp)
                Call PlayerMsgCombat(n, "You have gained " & exp & " party experience points.", BrightBlue)
            'End If
            End If
            End If
        End If
                                
        ' Drop the goods if they get it
        n = Int(Rnd * Npc(NpcNum).DropChance) + 1
        If n = 1 Then
            If Npc(NpcNum).DropItem > 0 Then
            Call SpawnItem(Npc(NpcNum).DropItem, Npc(NpcNum).DropItemValue, MapNum, MapNpc(MapNum, MapNpcNum).X, MapNpc(MapNum, MapNpcNum).y)
            End If
        End If
        If Npc(NpcNum).DropItem2 > 0 Then
        n = Int(Rnd * Npc(NpcNum).DropChance2) + 1
        If n = 1 Then
            Call SpawnItem(Npc(NpcNum).DropItem2, Npc(NpcNum).DropItemValue2, MapNum, MapNpc(MapNum, MapNpcNum).X, MapNpc(MapNum, MapNpcNum).y)
        End If
        End If
        
        ' Now set HP to 0 so we know to actually kill them in the server loop (this prevents subscript out of range)
        MapNpc(MapNum, MapNpcNum).Num = 0
        MapNpc(MapNum, MapNpcNum).SpawnWait = GetTickCount
        MapNpc(MapNum, MapNpcNum).HP = 0
        Call SendDataToMap(MapNum, "NPCDEAD" & SEP_CHAR & MapNpcNum & SEP_CHAR & END_CHAR)
        
        ' Check for level up
        Call CheckPlayerLevelUp(Attacker)
        
        ' Check for level up party member
       ' If player(Attacker).InParty = YES Then
       '     Call CheckPlayerLevelUp(player(Attacker).PartyPlayer)
       ' End If
    
        ' Check if target is npc that died and if so set target to 0
        If player(Attacker).TargetType = TARGET_TYPE_NPC And player(Attacker).Target = MapNpcNum Then
            player(Attacker).Target = 0
            player(Attacker).TargetType = 0
        End If
    Else
        ' NPC not dead, just do the damage
        MapNpc(MapNum, MapNpcNum).HP = MapNpc(MapNum, MapNpcNum).HP - Damage
        
        ' Check for a weapon and say damage
        If n = 0 Then
            Call PlayerMsgCombat(Attacker, "You hit a " & Name & " for " & Damage & " hit points.", White)
        Else
            Call PlayerMsgCombat(Attacker, "You hit a " & Name & " with a " & Trim(Item(n).Name) & " for " & Damage & " hit points.", White)
        End If
        
        ' Check if we should send a message
        If MapNpc(MapNum, MapNpcNum).Target = 0 And MapNpc(MapNum, MapNpcNum).Target <> Attacker Then
            If Trim(Npc(NpcNum).AttackSay) <> "" Then
                Call PlayerMsg(Attacker, Trim(Npc(NpcNum).Name) & " says, '" & Trim(Npc(NpcNum).AttackSay) & "'", SayColor)
            End If
        End If
        
        ' Set the NPC target to the player
        MapNpc(MapNum, MapNpcNum).Target = Attacker
        
        ' Now check for guard ai and if so have all onmap guards come after'm
        If Npc(MapNpc(MapNum, MapNpcNum).Num).Behavior = NPC_BEHAVIOR_GUARD Then
            For i = 1 To MAX_MAP_NPCS
                If MapNpc(MapNum, i).Num = MapNpc(MapNum, MapNpcNum).Num Then
                    MapNpc(MapNum, i).Target = Attacker
                End If
            Next i
        End If
    End If
    
    ' Reset attack timer
    player(Attacker).AttackTimer = GetTickCount
End Sub

Sub PlayerWarp(ByVal index As Long, ByVal MapNum As Long, ByVal X As Long, ByVal y As Long)


Dim Packet As String
Dim ShopNum As Long, OldMap As Long

    ' Check for subscript out of range
    If IsPlaying(index) = False Or MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If
    
    ' Check if there was an npc on the map the player is leaving, and if so say goodbye
    ShopNum = MAP(GetPlayerMap(index)).Shop
    If ShopNum > 0 Then
        If Trim(Shop(ShopNum).LeaveSay) <> "" Then
            Call PlayerMsg(index, Trim(Shop(ShopNum).Name) & " says, '" & Trim(Shop(ShopNum).LeaveSay) & "'", SayColor)
        End If
    End If
    
    ' Save old map to send erase player data to
    OldMap = GetPlayerMap(index)
    Call SendLeaveMap(index, OldMap)
    
    Call SetPlayerMap(index, MapNum)
    Call SetPlayerX(index, X)
    Call SetPlayerY(index, y)
    
    ' Check if there is an npc on the map and say hello if so
    ShopNum = MAP(GetPlayerMap(index)).Shop
    If ShopNum > 0 Then
        If Trim(Shop(ShopNum).JoinSay) <> "" Then
            Call PlayerMsg(index, Trim(Shop(ShopNum).Name) & " says, '" & Trim(Shop(ShopNum).JoinSay) & "'", SayColor)
        End If
    End If
            
    ' Now we check if there were any players left on the map the player just left, and if not stop processing npcs
    If GetTotalMapPlayers(OldMap) = 0 Then
        PlayersOnMap(OldMap) = NO
    End If
    
    ' Sets it so we know to process npcs on the map
    PlayersOnMap(MapNum) = YES
    
    player(index).GettingMap = YES
    Call SendDataTo(index, "CHECKFORMAP" & SEP_CHAR & MapNum & SEP_CHAR & MAP(MapNum).Revision & SEP_CHAR & END_CHAR)
End Sub

Sub PlayerMove(ByVal index As Long, ByVal Dir As Long, ByVal Movement As Long)


Dim Packet As String
Dim MapNum As Long
Dim X As Long
Dim y As Long
Dim i As Long
Dim Moved As Byte

    ' Check for subscript out of range
    If IsPlaying(index) = False Or Dir < DIR_UP Or Dir > DIR_RIGHT Or Movement < 1 Or Movement > 2 Then
        Exit Sub
    End If
    
    Call SetPlayerDir(index, Dir)
    
    Moved = NO
    
    Select Case Dir
        Case DIR_UP
            ' Check to make sure not outside of boundries
            If GetPlayerY(index) > 0 Then
                ' Check to make sure that the tile is walkable
                If MAP(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Type <> TILE_TYPE_BLOCKED Then
                    ' Check to see if the tile is a key and if it is check if its opened
                    If MAP(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Type <> TILE_TYPE_KEY Or (MAP(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Type = TILE_TYPE_KEY And TempTile(GetPlayerMap(index)).DoorOpen(GetPlayerX(index), GetPlayerY(index) - 1) = YES) Then
                        Call SetPlayerY(index, GetPlayerY(index) - 1)
                        
                        Packet = "PLAYERMOVE" & SEP_CHAR & index & SEP_CHAR & GetPlayerX(index) & SEP_CHAR & GetPlayerY(index) & SEP_CHAR & GetPlayerDir(index) & SEP_CHAR & Movement & SEP_CHAR & END_CHAR
                        Call SendDataToMapBut(index, GetPlayerMap(index), Packet)
                        Moved = YES
                    End If
                End If
            Else
                ' Check to see if we can move them to the another map
                If MAP(GetPlayerMap(index)).Up > 0 Then
                    Call PlayerWarp(index, MAP(GetPlayerMap(index)).Up, GetPlayerX(index), MAX_MAPY)
                    Moved = YES
                End If
            End If
                    
        Case DIR_DOWN
            ' Check to make sure not outside of boundries
            If GetPlayerY(index) < MAX_MAPY Then
                ' Check to make sure that the tile is walkable
                If MAP(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).Type <> TILE_TYPE_BLOCKED Then
                    ' Check to see if the tile is a key and if it is check if its opened
                    If MAP(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).Type <> TILE_TYPE_KEY Or (MAP(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).Type = TILE_TYPE_KEY And TempTile(GetPlayerMap(index)).DoorOpen(GetPlayerX(index), GetPlayerY(index) + 1) = YES) Then
                        Call SetPlayerY(index, GetPlayerY(index) + 1)
                        
                        Packet = "PLAYERMOVE" & SEP_CHAR & index & SEP_CHAR & GetPlayerX(index) & SEP_CHAR & GetPlayerY(index) & SEP_CHAR & GetPlayerDir(index) & SEP_CHAR & Movement & SEP_CHAR & END_CHAR
                        Call SendDataToMapBut(index, GetPlayerMap(index), Packet)
                        Moved = YES
                    End If
                End If
            Else
                ' Check to see if we can move them to the another map
                If MAP(GetPlayerMap(index)).Down > 0 Then
                    Call PlayerWarp(index, MAP(GetPlayerMap(index)).Down, GetPlayerX(index), 0)
                    Moved = YES
                End If
            End If
        
        Case DIR_LEFT
            ' Check to make sure not outside of boundries
            If GetPlayerX(index) > 0 Then
                ' Check to make sure that the tile is walkable
                If MAP(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).Type <> TILE_TYPE_BLOCKED Then
                    ' Check to see if the tile is a key and if it is check if its opened
                    If MAP(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).Type <> TILE_TYPE_KEY Or (MAP(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).Type = TILE_TYPE_KEY And TempTile(GetPlayerMap(index)).DoorOpen(GetPlayerX(index) - 1, GetPlayerY(index)) = YES) Then
                        Call SetPlayerX(index, GetPlayerX(index) - 1)
                        
                        Packet = "PLAYERMOVE" & SEP_CHAR & index & SEP_CHAR & GetPlayerX(index) & SEP_CHAR & GetPlayerY(index) & SEP_CHAR & GetPlayerDir(index) & SEP_CHAR & Movement & SEP_CHAR & END_CHAR
                        Call SendDataToMapBut(index, GetPlayerMap(index), Packet)
                        Moved = YES
                    End If
                End If
            Else
                ' Check to see if we can move them to the another map
                If MAP(GetPlayerMap(index)).Left > 0 Then
                    Call PlayerWarp(index, MAP(GetPlayerMap(index)).Left, MAX_MAPX, GetPlayerY(index))
                    Moved = YES
                End If
            End If
        
        Case DIR_RIGHT
            ' Check to make sure not outside of boundries
            If GetPlayerX(index) < MAX_MAPX Then
                ' Check to make sure that the tile is walkable
                If MAP(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).Type <> TILE_TYPE_BLOCKED Then
                    ' Check to see if the tile is a key and if it is check if its opened
                    If MAP(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).Type <> TILE_TYPE_KEY Or (MAP(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).Type = TILE_TYPE_KEY And TempTile(GetPlayerMap(index)).DoorOpen(GetPlayerX(index) + 1, GetPlayerY(index)) = YES) Then
                        Call SetPlayerX(index, GetPlayerX(index) + 1)
                        
                        Packet = "PLAYERMOVE" & SEP_CHAR & index & SEP_CHAR & GetPlayerX(index) & SEP_CHAR & GetPlayerY(index) & SEP_CHAR & GetPlayerDir(index) & SEP_CHAR & Movement & SEP_CHAR & END_CHAR
                        Call SendDataToMapBut(index, GetPlayerMap(index), Packet)
                        Moved = YES
                    End If
                End If
            Else
                ' Check to see if we can move them to the another map
                If MAP(GetPlayerMap(index)).Right > 0 Then
                    Call PlayerWarp(index, MAP(GetPlayerMap(index)).Right, 0, GetPlayerY(index))
                    Moved = YES
                End If
            End If
    End Select
    
    'check for sign
    If MAP(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Type = TILE_TYPE_SIGN Then
    Call PlayerMsgCombat(index, (MAP(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1), Magenta)
    Call PlayerMsgCombat(index, (MAP(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data2), Magenta)
        If MAP(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data3 <> "" Then
            Packet = "PLAYSFX" & SEP_CHAR & (MAP(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data3) & SEP_CHAR & END_CHAR
            Call SendDataTo(index, Packet)
        End If
    Moved = YES
    End If
    
    'check for full heal
    
    If MAP(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Type = TILE_TYPE_FULLHEAL Then
    Temp = GetPlayerMaxHP(index)
    Call SetPlayerHP(index, Temp)
    Call PlayerMsgCombat(index, "You have been fully healed", Magenta)
    Call SendHP(index)
    
    If MAP(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data3 <> "" Then
    Packet = "PLAYSFX" & SEP_CHAR & (MAP(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data3) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
    End If
    
    Moved = YES
    End If
    
    'check for death
    If MAP(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Type = TILE_TYPE_DEATH Then
    Call SetPlayerHP(index, 0)
    Call PlayerMsgCombat(index, MAP(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1, Magenta)
    Call SendHP(index)
    
    If MAP(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data2 <> "" Then
    Packet = "PLAYSFX" & SEP_CHAR & (MAP(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data2) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
    End If
    
     Moved = YES
     If GetPlayerWeaponSlot(index) > 0 Then
         Temp = GetPlayerWeaponSlot(index)
        ' Call SetPlayerSTR(index, (GetPlayerSTR(index) - Item(GetPlayerInvItemNum(index, tEmp)).STRmod))
        ' Call SetPlayerDEF(index, (GetPlayerDEF(index) - Item(GetPlayerInvItemNum(index, tEmp)).DEFmod))
        ' Call SetPlayerMAGI(index, (GetPlayerMAGI(index) - Item(GetPlayerInvItemNum(index, tEmp)).MAGImod))
         Call PlayerDeathDropItem(index, GetPlayerWeaponSlot(index), 0)
     End If
        If GetPlayerArmorSlot(index) > 0 Then
        Temp = GetPlayerArmorSlot(index)
                       ' Call SetPlayerSTR(index, (GetPlayerSTR(index) - Item(GetPlayerInvItemNum(index, tEmp)).STRmod))
                       ' Call SetPlayerDEF(index, (GetPlayerDEF(index) - Item(GetPlayerInvItemNum(index, tEmp)).DEFmod))
                       ' Call SetPlayerMAGI(index, (GetPlayerMAGI(index) - Item(GetPlayerInvItemNum(index, tEmp)).MAGImod))
                        
            Call PlayerDeathDropItem(index, GetPlayerArmorSlot(index), 0)
            End If
        If GetPlayerHelmetSlot(index) > 0 Then
            Call PlayerDeathDropItem(index, GetPlayerHelmetSlot(index), 0)
            End If
        If GetPlayerShieldSlot(index) > 0 Then
            Call PlayerDeathDropItem(index, GetPlayerShieldSlot(index), 0)
           End If
           ' Calculate exp to give attacker
          
        Temp = Int(GetPlayerExp(index) / 15)
        
        ' Make sure we dont get less then 0
        If Temp < 0 Then
            Temp = 0
        End If
        
        If Temp = 0 Then
            Call PlayerMsgCombat(index, "You lost no experience points.", BrightRed)
        Else
            Call SetPlayerExp(index, GetPlayerExp(index) - Temp)
            Call PlayerMsgCombat(index, "You lost " & Temp & " experience points.", BrightRed)
        End If
            ' Warp player away
            
        'If MAP(GetPlayerMap(index)).BootMap > 0 Then
        'Call PlayerWarp(index, MAP(GetPlayerMap(index)).BootMap, MAP(GetPlayerMap(index)).BootX, MAP(GetPlayerMap(index)).BootY)
        'Else
        'Call PlayerWarp(index, START_MAP, START_X, START_Y)
        'End If
        Call PlayerWarp(index, GetPlayerBindMap(index), GetPlayerBindX(index), GetPlayerBindY(index))
        Moved = YES
        'Call SendPlayerData(Index)
        ' Restore vitals
        Call SetPlayerHP(index, GetPlayerMaxHP(index))
        Call SetPlayerMP(index, GetPlayerMaxMP(index))
        Call SetPlayerSP(index, GetPlayerMaxSP(index))
        Call SendHP(index)
        Call SendMP(index)
        Call SendSP(index)
        'Call SendPlayerData(Index)
        End If
    'check for SFXTile
    If MAP(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Type = TILE_TYPE_SFX Then
    Packet = "PLAYSFX" & SEP_CHAR & (MAP(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
    Moved = YES
    End If
    
    'QUEST TILE W007!
    If MAP(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Type = TILE_TYPE_QUEST Then
        Temp = Val(MAP(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1)
        tEmp2 = Val(MAP(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data3)
        If HasItem(index, Temp) >= 1 Then
            Call TakeItem(index, Temp, 1)
            Call GiveItem(index, tEmp2, 1)
            Call PlayerMsgCombat(index, (MAP(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data2), Grey)
        End If
        Moved = YES
    End If
    
    
    
        ' Check to see if the tile is a warp tile, and if so warp them
    If MAP(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Type = TILE_TYPE_WARP Then
        MapNum = MAP(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1
        X = MAP(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data2
        y = MAP(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data3
                        
        Call PlayerWarp(index, MapNum, X, y)
        'Packet = "PLAYSFX" & SEP_CHAR & "teleport1" & SEP_CHAR & END_CHAR
        'Call SendDataTo(index, Packet)
        Moved = YES
    End If
    
    ' Check for key trigger open
    If MAP(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Type = TILE_TYPE_KEYOPEN Then
        X = MAP(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1
        y = MAP(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data2
        
        If MAP(GetPlayerMap(index)).Tile(X, y).Type = TILE_TYPE_KEY And TempTile(GetPlayerMap(index)).DoorOpen(X, y) = NO Then
            TempTile(GetPlayerMap(index)).DoorOpen(X, y) = YES
            TempTile(GetPlayerMap(index)).DoorTimer = GetTickCount
                            
            Call SendDataToMap(GetPlayerMap(index), "MAPKEY" & SEP_CHAR & X & SEP_CHAR & y & SEP_CHAR & 1 & SEP_CHAR & END_CHAR)
            Call MapMsg(GetPlayerMap(index), "A door has been unlocked.", White)
        End If
    End If
     Moved = YES
    ' They tried to hack
    If Moved = NO Then
        Call HackingAttempt(index, "Position Modification")
    End If
End Sub

Function CanNpcMove(ByVal MapNum As Long, ByVal MapNpcNum As Long, ByVal Dir) As Boolean


Dim i As Long, n As Long
Dim X As Long, y As Long

    CanNpcMove = False
    
    ' Check for subscript out of range
    If MapNum <= 0 Or MapNum > MAX_MAPS Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Function
    End If
    
    X = MapNpc(MapNum, MapNpcNum).X
    y = MapNpc(MapNum, MapNpcNum).y
    
    CanNpcMove = True
    
    Select Case Dir
        Case DIR_UP
            ' Check to make sure not outside of boundries
            If y > 0 Then
                n = MAP(MapNum).Tile(X, y - 1).Type
                
                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_MINING And n <> TILE_TYPE_FISHING And n <> TILE_TYPE_SIGN Then
                    CanNpcMove = False
                    Exit Function
                End If
                
                ' Check to make sure that there is not a player in the way
                For i = 1 To MAX_PLAYERS
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = MapNpc(MapNum, MapNpcNum).X) And (GetPlayerY(i) = MapNpc(MapNum, MapNpcNum).y - 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next i
                
                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> MapNpcNum) And (MapNpc(MapNum, i).Num > 0) And (MapNpc(MapNum, i).X = MapNpc(MapNum, MapNpcNum).X) And (MapNpc(MapNum, i).y = MapNpc(MapNum, MapNpcNum).y - 1) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next i
            Else
                CanNpcMove = False
            End If
                
        Case DIR_DOWN
            ' Check to make sure not outside of boundries
            If y < MAX_MAPY Then
                n = MAP(MapNum).Tile(X, y + 1).Type
                
                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_MINING And n <> TILE_TYPE_FISHING And n <> TILE_TYPE_SIGN Then
                     CanNpcMove = False
                    Exit Function
                End If
                
                ' Check to make sure that there is not a player in the way
                For i = 1 To MAX_PLAYERS
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = MapNpc(MapNum, MapNpcNum).X) And (GetPlayerY(i) = MapNpc(MapNum, MapNpcNum).y + 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next i
                
                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> MapNpcNum) And (MapNpc(MapNum, i).Num > 0) And (MapNpc(MapNum, i).X = MapNpc(MapNum, MapNpcNum).X) And (MapNpc(MapNum, i).y = MapNpc(MapNum, MapNpcNum).y + 1) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next i
            Else
                CanNpcMove = False
            End If
                
        Case DIR_LEFT
            ' Check to make sure not outside of boundries
            If X > 0 Then
                n = MAP(MapNum).Tile(X - 1, y).Type
                
                ' Check to make sure that the tile is walkable
               If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_MINING And n <> TILE_TYPE_FISHING And n <> TILE_TYPE_SIGN Then
                    CanNpcMove = False
                    Exit Function
                End If
                
                ' Check to make sure that there is not a player in the way
                For i = 1 To MAX_PLAYERS
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = MapNpc(MapNum, MapNpcNum).X - 1) And (GetPlayerY(i) = MapNpc(MapNum, MapNpcNum).y) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next i
                
                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> MapNpcNum) And (MapNpc(MapNum, i).Num > 0) And (MapNpc(MapNum, i).X = MapNpc(MapNum, MapNpcNum).X - 1) And (MapNpc(MapNum, i).y = MapNpc(MapNum, MapNpcNum).y) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next i
            Else
                CanNpcMove = False
            End If
                
        Case DIR_RIGHT
            ' Check to make sure not outside of boundries
            If X < MAX_MAPX Then
                n = MAP(MapNum).Tile(X + 1, y).Type
                
                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_MINING And n <> TILE_TYPE_FISHING And n <> TILE_TYPE_SIGN Then
                    CanNpcMove = False
                    Exit Function
                End If
                
                ' Check to make sure that there is not a player in the way
                For i = 1 To MAX_PLAYERS
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = MapNpc(MapNum, MapNpcNum).X + 1) And (GetPlayerY(i) = MapNpc(MapNum, MapNpcNum).y) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next i
                
                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> MapNpcNum) And (MapNpc(MapNum, i).Num > 0) And (MapNpc(MapNum, i).X = MapNpc(MapNum, MapNpcNum).X + 1) And (MapNpc(MapNum, i).y = MapNpc(MapNum, MapNpcNum).y) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next i
            Else
                CanNpcMove = False
            End If
    End Select
End Function

Sub NpcMove(ByVal MapNum As Long, ByVal MapNpcNum As Long, ByVal Dir As Long, ByVal Movement As Long)


Dim Packet As String
Dim X As Long
Dim y As Long
Dim i As Long

    ' Check for subscript out of range
    If MapNum <= 0 Or MapNum > MAX_MAPS Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or Dir < DIR_UP Or Dir > DIR_RIGHT Or Movement < 1 Or Movement > 2 Then
        Exit Sub
    End If
    
    MapNpc(MapNum, MapNpcNum).Dir = Dir
    
    Select Case Dir
        Case DIR_UP
            MapNpc(MapNum, MapNpcNum).y = MapNpc(MapNum, MapNpcNum).y - 1
            Packet = "NPCMOVE" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum, MapNpcNum).X & SEP_CHAR & MapNpc(MapNum, MapNpcNum).y & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Dir & SEP_CHAR & Movement & SEP_CHAR & END_CHAR
            Call SendDataToMap(MapNum, Packet)
    
        Case DIR_DOWN
            MapNpc(MapNum, MapNpcNum).y = MapNpc(MapNum, MapNpcNum).y + 1
            Packet = "NPCMOVE" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum, MapNpcNum).X & SEP_CHAR & MapNpc(MapNum, MapNpcNum).y & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Dir & SEP_CHAR & Movement & SEP_CHAR & END_CHAR
            Call SendDataToMap(MapNum, Packet)
    
        Case DIR_LEFT
            MapNpc(MapNum, MapNpcNum).X = MapNpc(MapNum, MapNpcNum).X - 1
            Packet = "NPCMOVE" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum, MapNpcNum).X & SEP_CHAR & MapNpc(MapNum, MapNpcNum).y & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Dir & SEP_CHAR & Movement & SEP_CHAR & END_CHAR
            Call SendDataToMap(MapNum, Packet)
    
        Case DIR_RIGHT
            MapNpc(MapNum, MapNpcNum).X = MapNpc(MapNum, MapNpcNum).X + 1
            Packet = "NPCMOVE" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum, MapNpcNum).X & SEP_CHAR & MapNpc(MapNum, MapNpcNum).y & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Dir & SEP_CHAR & Movement & SEP_CHAR & END_CHAR
            Call SendDataToMap(MapNum, Packet)
    End Select
End Sub

Sub NpcDir(ByVal MapNum As Long, ByVal MapNpcNum As Long, ByVal Dir As Long)
Dim Packet As String

    ' Check for subscript out of range
    If MapNum <= 0 Or MapNum > MAX_MAPS Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Sub
    End If
    
    MapNpc(MapNum, MapNpcNum).Dir = Dir
    Packet = "NPCDIR" & SEP_CHAR & MapNpcNum & SEP_CHAR & Dir & SEP_CHAR & END_CHAR
    Call SendDataToMap(MapNum, Packet)
End Sub

Sub JoinGame(ByVal index As Long)
Dim i As Long
    ' Set the flag so we know the person is in the game
    player(index).InGame = True
    ' Send a global message that he/she joined
   If GetPlayerAccess(index) <= 0 Then

        Call GetRank(index)
        Call GlobalMsgCombat(GetPlayerName(index) & " has joined " & GAME_NAME & "!", JoinLeftColor)
    Else
        Call GlobalMsgCombat(GetPlayerName(index) & " has joined " & GAME_NAME & "!", White)
    End If
        
    ' Send an ok to client to start receiving in game data
    Call SendDataTo(index, "LOGINOK" & SEP_CHAR & index & SEP_CHAR & END_CHAR)
    
    ' Send some more little goodies, no need to explain these
    Call CheckEquippedItems(index)
    Call SendClasses(index)
    Call SendRaces(index)
    Call SendItems(index)
    Call SendNpcs(index)
    Call SendShops(index)
    Call SendSpells(index)
    Call SendInventory(index)
    Call SendWornEquipment(index)
    Call SendHP(index)
    Call SendMP(index)
    Call SendSP(index)
    Call SendStats(index)
    Call SendWeatherTo(index)
    Call SendTimeTo(index)
    Call SendAdminCmds(index)
    
    
    ' Warp the player to his saved location
    Call PlayerWarp(index, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index))
            
    ' Send welcome messages
    Call SendWelcome(index)

    ' Send the flag so they know they can start doing stuff
    Call SendDataTo(index, "INGAME" & SEP_CHAR & END_CHAR)
End Sub

Sub LeftGame(ByVal index As Long)


Dim n As Long
Dim i As Long
Call SavePlayer(index)
    If player(index).InGame = True Then
        player(index).InGame = False
        
        ' Check if player was the only player on the map and stop npc processing if so
        If GetTotalMapPlayers(GetPlayerMap(index)) = 0 Then
            PlayersOnMap(GetPlayerMap(index)) = NO
        End If
        
        ' Check for boot map
       ' If MAP(GetPlayerMap(index)).BootMap > 0 Then
       '     Call SetPlayerX(index, MAP(GetPlayerMap(index)).BootX)
       '     Call SetPlayerY(index, MAP(GetPlayerMap(index)).BootY)
       '     Call SetPlayerMap(index, MAP(GetPlayerMap(index)).BootMap)
       ' End If
        
        ' Check if the player was in a party, and if so cancel it out so the other player doesn't continue to get half exp
        If player(index).InParty = YES Then
           If player(index).Party > 0 Then
           i = player(index).Party
           If Parties(i).NumParty >= 1 Then
               Parties(i).NumParty = Parties(i).NumParty - 1
               End If
               If Parties(i).Player1 = index Then
               Parties(i).Player1 = 0
               ElseIf Parties(i).Player2 = index Then
               Parties(i).Player2 = 0
                ElseIf Parties(i).Player3 = index Then
               Parties(i).Player3 = 0
                ElseIf Parties(i).Player4 = index Then
               Parties(i).Player4 = 0
               End If
               If Parties(i).NumParty >= 1 Then
               If Parties(i).PartyLeader = index Then
               If Parties(i).Player1 > 0 Then
               Parties(i).PartyLeader = Parties(i).Player1
               Call PlayerMsgCombat(Parties(i).Player1, "You are now the leader of the party!", BrightRed)
               ElseIf Parties(i).Player2 > 0 Then
               Parties(i).PartyLeader = Parties(i).Player2
               Call PlayerMsgCombat(Parties(i).Player2, "You are now the leader of the party!", BrightRed)
               ElseIf Parties(i).Player3 > 0 Then
               Parties(i).PartyLeader = Parties(i).Player3
               Call PlayerMsgCombat(Parties(i).Player3, "You are now the leader of the party!", BrightRed)
               ElseIf Parties(i).Player4 > 0 Then
               Parties(i).PartyLeader = Parties(i).Player4
               Call PlayerMsgCombat(Parties(i).Player4, "You are now the leader of the party!", BrightRed)
               End If
               End If
               End If
            
            'Call PlayerMsg(n, GetPlayerName(Index) & " has left " & GAME_NAME & ".", Pink)
            
        End If
        End If
        
    
        ' Send a global message that he/she left
        If GetPlayerAccess(index) <= ADMIN_MONITER Then
            Call GlobalMsgCombat(GetPlayerName(index) & " has left " & GAME_NAME & "!", JoinLeftColor)
        Else
            Call GlobalMsgCombat(GetPlayerName(index) & " has left " & GAME_NAME & "!", White)
        End If
        Call TextAdd(frmServer.txtText, GetPlayerName(index) & " has disconnected from " & GAME_NAME & ".", True)
        Call SendLeftGame(index)
    End If
    
    Call ClearPlayer(index)
End Sub

Function GetTotalMapPlayers(ByVal MapNum As Long) As Long


Dim i As Long, n As Long

    n = 0
    
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) And GetPlayerMap(i) = MapNum Then
            n = n + 1
        End If
    Next i
    
    GetTotalMapPlayers = n
End Function

Function GetNpcMaxHP(ByVal NpcNum As Long)


Dim X As Long, y As Long

    ' Prevent subscript out of range
    If NpcNum <= 0 Or NpcNum > MAX_NPCS Then
        GetNpcMaxHP = 0
        Exit Function
    End If
    
    X = Npc(NpcNum).STR
    y = Npc(NpcNum).DEF
    GetNpcMaxHP = ((X * y) * 0.4) * 1.1
End Function

Function GetNpcMaxMP(ByVal NpcNum As Long)


    ' Prevent subscript out of range
    If NpcNum <= 0 Or NpcNum > MAX_NPCS Then
        GetNpcMaxMP = 0
        Exit Function
    End If
        
    GetNpcMaxMP = Npc(NpcNum).MAGI
End Function

Function GetNpcMaxSP(ByVal NpcNum As Long)


    ' Prevent subscript out of range
    If NpcNum <= 0 Or NpcNum > MAX_NPCS Then
        GetNpcMaxSP = 0
        Exit Function
    End If
        
    GetNpcMaxSP = Npc(NpcNum).SPeed * 2
End Function

Function GetPlayerHPRegen(ByVal index As Long)


Dim i As Long

    ' Prevent subscript out of range
    If IsPlaying(index) = False Or index <= 0 Or index > MAX_PLAYERS Then
        GetPlayerHPRegen = 0
        Exit Function
    End If
    
    i = Int(GetPlayerSPEED(index))
    If i < 2 Then i = 2
    
    GetPlayerHPRegen = i / 1.4
End Function

Function GetPlayerMPRegen(ByVal index As Long)


Dim i As Long

    ' Prevent subscript out of range
    If IsPlaying(index) = False Or index <= 0 Or index > MAX_PLAYERS Then
        GetPlayerMPRegen = 0
        Exit Function
    End If
    
    i = Int(GetPlayerSPEED(index))
    If i < 2 Then i = 2
    
    GetPlayerMPRegen = i / 1.2
End Function

Function GetPlayerSPRegen(ByVal index As Long)


Dim i As Long

    ' Prevent subscript out of range
    If IsPlaying(index) = False Or index <= 0 Or index > MAX_PLAYERS Then
        GetPlayerSPRegen = 0
        Exit Function
    End If
    
    i = Int(GetPlayerSPEED(index))
    If i < 2 Then i = 2
    
    GetPlayerSPRegen = i / 1.8
End Function

Function GetNpcHPRegen(ByVal NpcNum As Long)


Dim i As Long

    'Prevent subscript out of range
    If NpcNum <= 0 Or NpcNum > MAX_NPCS Then
        GetNpcHPRegen = 0
        Exit Function
    End If
    
    i = Int(Npc(NpcNum).DEF / 3)
    If i < 1 Then i = 1
    
    GetNpcHPRegen = i
End Function

Sub CheckPlayerLevelUp(ByVal index As Long)


Dim i As Long

    ' Check if attacker got a level up
    If GetPlayerExp(index) >= GetPlayerNextLevel(index) Then
        Call SetPlayerLevel(index, GetPlayerLevel(index) + 1)
                    
        ' Get the ammount of skill points to add
        i = Int(GetPlayerLevel(index) / 10)
        i = i + 1
        If i > 6 Then i = 6
        Call SetPlayerPOINTS(index, GetPlayerPOINTS(index) + i)
        Call SetPlayerExp(index, 0)
        Call GlobalMsgCombat(GetPlayerName(index) & " has gained a level!", Brown)
        Call PlayerMsgCombat(index, "You have gained a level!  You now have " & GetPlayerPOINTS(index) & " stat points to distribute.", BrightBlue)
        Call DeathSound(index, 99, 99)
    End If
End Sub

Sub CastSpell(ByVal index As Long, ByVal SpellSlot As Long)


Dim SpellNum As Long, MPReq As Long, i As Long, n As Long, Damage As Long
Dim Casted As Boolean
Dim X As Long, y As Long

    Casted = False
    
    ' Prevent subscript out of range
    If SpellSlot <= 0 Or SpellSlot > MAX_PLAYER_SPELLS Then
        Exit Sub
    End If
    
    SpellNum = GetPlayerSpell(index, SpellSlot)
    
    ' Make sure player has the spell
    If Not HasSpell(index, SpellNum) Then
        Call PlayerMsgCombat(index, "You do not have this spell!", BrightRed)
        Exit Sub
    End If
    If spell(SpellNum).Type <> SPELL_TYPE_CRAFT Then
    i = GetSpellReqLevel(index, SpellNum)
    Else
    i = 1
    End If
    MPReq = spell(SpellNum).MPused
    If spell(SpellNum).Type <> SPELL_TYPE_CRAFT Then
    ' Check if they have enough MP
    If GetPlayerMP(index) < MPReq Then
        Call PlayerMsgCombat(index, "Not enough mana points!", BrightRed)
        Exit Sub
    End If
    End If
    ' Make sure they are the right level
    If i > GetPlayerLevel(index) Then
        Call PlayerMsgCombat(index, "You must be level " & i & " to cast this spell.", BrightRed)
        Exit Sub
    End If
    
    ' Check if timer is ok
    If GetTickCount < player(index).AttackTimer + 1000 Then
        Exit Sub
    End If
   
    
    
    ' Check if the spell is a give item and do that instead of a stat modification
    If spell(SpellNum).Type = SPELL_TYPE_GIVEITEM Then
        n = FindOpenInvSlot(index, spell(SpellNum).Data1)
        Call SpellSound(index, spell(SpellNum).Sfx)
        If n > 0 Then
            Call GiveItem(index, spell(SpellNum).Data1, spell(SpellNum).Data2)
            Call MapMsgCombat(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim(spell(SpellNum).Name) & ".", BrightBlue)
            
            ' Take away the mana points
            Call SetPlayerMP(index, GetPlayerMP(index) - MPReq)
            Call SendMP(index)
            Casted = True
        Else
            Call PlayerMsgCombat(index, "Your inventory is full!", BrightRed)
        End If
        
        Exit Sub
    End If
    ' Check if the spell is a craft item and do that instead of a stat modification
    If spell(SpellNum).Type = SPELL_TYPE_CRAFT Then
    'X = (GetSpellReqLevel(Index, SpellNum) - GetPlayerCraft(Index)) + 5
    'If X < 0 Then X = 0
    Randomize
    X = GetPlayerCraft(index)
    y = Int(Rnd * GetSpellReqLevel(index, SpellNum))
    If y <= X Then
        Temp = spell(SpellNum).Data1
        tEmp2 = spell(SpellNum).Data2
    'If y = 1 Or y = 2 Or y = 3 Or y = 4 Or y = 5 Then
        n = FindOpenInvSlot(index, spell(SpellNum).Data1)
    
        Temp = spell(SpellNum).Data1
        tEmp2 = spell(SpellNum).Data2
    
        
        If n > 0 Then
            If HasItem(index, Temp) >= Val(spell(SpellNum).Sfx) And HasItem(index, tEmp2) >= Val(spell(SpellNum).Gfx) Then
                Call TakeItem(index, Temp, Val(spell(SpellNum).Sfx))
                Call TakeItem(index, tEmp2, Val(spell(SpellNum).Gfx))
                Call GiveItem(index, spell(SpellNum).Data3, 1)
                Call PlayerMsgCombat(index, "You have crafted a " & Trim(Item(spell(SpellNum).Data3).Name) & ".", BrightRed)
                y = Int(Rnd * 150)
                If y = 1 Then
                Call SetPlayerCraft(index, GetPlayerCraft(index) + 1)
                Call PlayerMsgCombat(index, "You have gained a level in Crafting!", BrightRed)
                End If
            Else
            Tmpstr = "You need " & Val(spell(SpellNum).Sfx) & " " & Trim(Item(spell(SpellNum).Data1).Name) & " And " & Val(spell(SpellNum).Gfx) & " " & Trim(Item(spell(SpellNum).Data2).Name) & " to create a " & Trim(Item(spell(SpellNum).Data3).Name) & "."
                Call PlayerMsgCombat(index, Tmpstr, BrightRed)
            End If
        Else
            Call PlayerMsgCombat(index, "Your inventory is full!", BrightRed)
        End If
    Else
    Temp = spell(SpellNum).Data1
        tEmp2 = spell(SpellNum).Data2
    If HasItem(index, Temp) >= Val(spell(SpellNum).Sfx) And HasItem(index, tEmp2) >= Val(spell(SpellNum).Gfx) Then
                Call TakeItem(index, Temp, Val(spell(SpellNum).Sfx))
                Call TakeItem(index, tEmp2, Val(spell(SpellNum).Gfx))
                Call PlayerMsgCombat(index, "Your items have been destroyed in your failed attempt in crafting!", BrightRed)
                Else
                 Tmpstr = "You need " & Val(spell(SpellNum).Sfx) & " " & Trim(Item(spell(SpellNum).Data1).Name) & " And " & Val(spell(SpellNum).Gfx) & " " & Trim(Item(spell(SpellNum).Data2).Name) & " to create a " & Trim(Item(spell(SpellNum).Data3).Name) & "."
                Call PlayerMsgCombat(index, Tmpstr, BrightRed)
            End If
    End If
        
        Exit Sub
        
    End If
        
        
        
        
        
        
        
        
        
    n = player(index).Target
    
    If player(index).TargetType = TARGET_TYPE_PLAYER Then
        If IsPlaying(n) Then
       
            If GetPlayerHP(n) > 0 And GetPlayerMap(index) = GetPlayerMap(n) And GetPlayerLevel(index) >= 10 And GetPlayerLevel(n) >= 10 And spell(SpellNum).Type >= SPELL_TYPE_SUBHP And spell(SpellNum).Type <= SPELL_TYPE_SUBSP Then
If MAP(GetPlayerMap(index)).BootY = 0 And GetPlayerLevel(index) >= 10 And GetPlayerLevel(n) >= 10 And ((MAP(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Type = TILE_TYPE_ARENA And MAP(GetPlayerMap(n)).Tile(GetPlayerX(n), GetPlayerY(n)).Type = TILE_TYPE_ARENA) Or (GetPlayerFaction(index) > 0 And GetPlayerFaction(n) > 0) Or (GetPlayerAnonymous(index) > 0 And GetPlayerAnonymous(n) > 0 And GetPlayerGuildRank(index) > 0 And GetPlayerGuildRank(n) > 0 And GetPlayerGuild(index) <> "" And GetPlayerGuild(n) <> "") Or MAP(GetPlayerMap(index)).Moral = MAP_MORAL_NONE) Then
                'If GetPlayerLevel(n) + 6 >= GetPlayerLevel(Index) Then
                   ' If GetPlayerLevel(n) - 6 <= GetPlayerLevel(Index) Then
                   If index = n Then
                   Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " Trys to kill themselfs and fails poorly!", BrightBlue)
                   Exit Sub
                   End If
                         Call MapMsgCombat(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim(spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                         
                        Select Case spell(SpellNum).Type
                            Case SPELL_TYPE_SUBHP
                        
                                Damage = (Int(GetPlayerMAGI(index) / 4) + spell(SpellNum).Data1) - GetPlayerProtection(n)
                                If Damage > 0 Then
                                    Call AttackPlayer(index, n, Damage, SpellNum)
                                Else
                                    Call PlayerMsgCombat(index, "The spell was to weak to hurt " & GetPlayerName(n) & "!", BrightRed)
                                End If
                    
                            Case SPELL_TYPE_SUBMP
                                Call SetPlayerMP(n, GetPlayerMP(n) - spell(SpellNum).Data1)
                                Call SendMP(n)
                
                            Case SPELL_TYPE_SUBSP
                                Call SetPlayerSP(n, GetPlayerSP(n) - spell(SpellNum).Data1)
                                Call SendSP(n)
                        End Select
                        End If
                    'Else
                    ' Call PlayerMsgCombat(Index, GetPlayerName(n) & " is to weak to even bother with.", BrightBlue)
                
                    '   End If
               ' Else
                 '   Call PlayerMsgCombat(Index, GetPlayerName(n) & " is far to powerful to even consider attacking.", BrightBlue)
                    
                  '' End If
            '
                ' Take away the mana points
                Call SetPlayerMP(index, GetPlayerMP(index) - MPReq)
                Call SendMP(index)
                Casted = True
            Else
                If GetPlayerMap(index) = GetPlayerMap(n) And spell(SpellNum).Type >= SPELL_TYPE_ADDHP And spell(SpellNum).Type <= SPELL_TYPE_ADDSP Then
                    Call SpellSound(index, spell(SpellNum).Sfx)
                         Call SendDataToMap(GetPlayerMap(index), "SPELLGFX" & SEP_CHAR & n & SEP_CHAR & spell(SpellNum).Gfx & SEP_CHAR & END_CHAR)
                          
                    Select Case spell(SpellNum).Type
                   
                        Case SPELL_TYPE_ADDHP
                            Call MapMsgCombat(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim(spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                            Call SetPlayerHP(n, GetPlayerHP(n) + spell(SpellNum).Data1)
                            Call SendHP(n)
                                    
                        Case SPELL_TYPE_ADDMP
                            Call MapMsgCombat(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim(spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                            Call SetPlayerMP(n, GetPlayerMP(n) + spell(SpellNum).Data1)
                            Call SendMP(n)
                    
                        Case SPELL_TYPE_ADDSP
                            Call MapMsgCombat(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim(spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                            Call SetPlayerSP(n, GetPlayerSP(n) + spell(SpellNum).Data1)
                            Call SendSP(n)
                    End Select
    
                    ' Take away the mana points
                    Call SetPlayerMP(index, GetPlayerMP(index) - MPReq)
                    Call SendMP(index)
                    Casted = True
                Else
                    Call PlayerMsgCombat(index, "Could not cast spell!", BrightRed)
                End If
            End If
        Else
            Call PlayerMsgCombat(index, "Could not cast spell!", BrightRed)
        End If
    Else
        Call SpellSound(index, spell(SpellNum).Sfx)
        If Npc(MapNpc(GetPlayerMap(index), n).Num).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(MapNpc(GetPlayerMap(index), n).Num).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then
            Call SendDataToMap(GetPlayerMap(index), "SPELLGFX2" & SEP_CHAR & n & SEP_CHAR & spell(SpellNum).Gfx & SEP_CHAR & END_CHAR)
            Call MapMsgCombat(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim(spell(SpellNum).Name) & " on a " & Trim(Npc(MapNpc(GetPlayerMap(index), n).Num).Name) & ".", BrightBlue)
            
            Select Case spell(SpellNum).Type
                Case SPELL_TYPE_ADDHP
                    MapNpc(GetPlayerMap(index), n).HP = MapNpc(GetPlayerMap(index), n).HP + spell(SpellNum).Data1
                
                Case SPELL_TYPE_SUBHP
                    
                    Damage = (Int(GetPlayerMAGI(index) / 4) + spell(SpellNum).Data1) - Int(Npc(MapNpc(GetPlayerMap(index), n).Num).DEF / 2)
                    If Damage > 0 Then
                        Call AttackNpc(index, n, Damage, SpellNum)
                    Else
                        Call PlayerMsgCombat(index, "The spell was to weak to hurt " & Trim(Npc(MapNpc(GetPlayerMap(index), n).Num).Name) & "!", BrightRed)
                    End If
                    
                Case SPELL_TYPE_ADDMP
                    MapNpc(GetPlayerMap(index), n).MP = MapNpc(GetPlayerMap(index), n).MP + spell(SpellNum).Data1
                
                Case SPELL_TYPE_SUBMP
                    MapNpc(GetPlayerMap(index), n).MP = MapNpc(GetPlayerMap(index), n).MP - spell(SpellNum).Data1
            
                Case SPELL_TYPE_ADDSP
                    MapNpc(GetPlayerMap(index), n).SP = MapNpc(GetPlayerMap(index), n).SP + spell(SpellNum).Data1
                
                Case SPELL_TYPE_SUBSP
                    MapNpc(GetPlayerMap(index), n).SP = MapNpc(GetPlayerMap(index), n).SP - spell(SpellNum).Data1
            End Select
        
            ' Take away the mana points
            Call SetPlayerMP(index, GetPlayerMP(index) - MPReq)
            Call SendMP(index)
            Casted = True
        Else
            Call PlayerMsgCombat(index, "Could not cast spell!", BrightRed)
        End If
    End If

    If Casted = True Then
        player(index).AttackTimer = GetTickCount
        player(index).CastedSpell = YES
    End If
End Sub

Function GetSpellReqLevel(ByVal index As Long, ByVal SpellNum As Long)


    GetSpellReqLevel = spell(SpellNum).LevelReq
    
End Function

Function CanPlayerCriticalHit(ByVal index As Long) As Boolean


Dim i As Long, n As Long

    CanPlayerCriticalHit = False
    
    If GetPlayerWeaponSlot(index) > 0 Then
        n = Int(Rnd * 2)
        If n = 1 Then
            i = Int(GetPlayerSTR(index) / 2) + Int(GetPlayerLevel(index) / 2)
    
            n = Int(Rnd * 100) + 1
            If n <= i Then
                CanPlayerCriticalHit = True
            End If
        End If
    End If
End Function

Function CanPlayerBlockHit(ByVal index As Long) As Boolean


Dim i As Long, n As Long, ShieldSlot As Long

    CanPlayerBlockHit = False
    
    ShieldSlot = GetPlayerShieldSlot(index)
    
    If ShieldSlot > 0 Then
        n = Int(Rnd * 2)
        If n = 1 Then
            i = Int(GetPlayerDEF(index) / 2) + Int(GetPlayerLevel(index) / 2)
        
            n = Int(Rnd * 100) + 1
            If n <= i Then
                CanPlayerBlockHit = True
            End If
        End If
    End If
End Function

Sub CheckEquippedItems(ByVal index As Long)


Dim Slot As Long, ItemNum As Long

    ' We want to check incase an admin takes away an object but they had it equipped
    Slot = GetPlayerWeaponSlot(index)
    If Slot > 0 Then
        ItemNum = GetPlayerInvItemNum(index, Slot)
        
        If ItemNum > 0 Then
            If Item(ItemNum).Type <> ITEM_TYPE_WEAPON Then
                Call SetPlayerWeaponSlot(index, 0)
            End If
        Else
            Call SetPlayerWeaponSlot(index, 0)
        End If
    End If

    Slot = GetPlayerArmorSlot(index)
    If Slot > 0 Then
        ItemNum = GetPlayerInvItemNum(index, Slot)
        
        If ItemNum > 0 Then
            If Item(ItemNum).Type <> ITEM_TYPE_ARMOR Then
                Call SetPlayerArmorSlot(index, 0)
            End If
        Else
            Call SetPlayerArmorSlot(index, 0)
        End If
    End If

    Slot = GetPlayerHelmetSlot(index)
    If Slot > 0 Then
        ItemNum = GetPlayerInvItemNum(index, Slot)
        
        If ItemNum > 0 Then
            If Item(ItemNum).Type <> ITEM_TYPE_HELMET Then
                Call SetPlayerHelmetSlot(index, 0)
            End If
        Else
            Call SetPlayerHelmetSlot(index, 0)
        End If
    End If

    Slot = GetPlayerShieldSlot(index)
    If Slot > 0 Then
        ItemNum = GetPlayerInvItemNum(index, Slot)
        
        If ItemNum > 0 Then
            If Item(ItemNum).Type <> ITEM_TYPE_SHIELD Then
                Call SetPlayerShieldSlot(index, 0)
            End If
        Else
            Call SetPlayerShieldSlot(index, 0)
        End If
    End If
End Sub
