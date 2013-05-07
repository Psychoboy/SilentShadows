Attribute VB_Name = "modTypes"
Option Explicit

Public CanGlobal As Integer

Public PlayerI As Byte
Public Temp   As Integer
Public tEmp2  As Integer
Public tEmp3 As Integer
Public Tmpstr As String

' Winsock globals
Public Const GAME_PORT = 7555

Public RankName As String



' General constants
Public Const GAME_NAME = "Silent Shadows"
Public Const MAX_PLAYERS = 100

Public Const MAX_ITEMS = 1000
'Public Const MAX_ITEMS = 10

Public Const MAX_NPCS = 500
'Public Const MAX_NPCS = 10

Public Const MAX_INV = 50
Public Const MAX_BANK = 50
Public Const MAX_MAP_ITEMS = 20
Public Const MAX_MAP_NPCS = 15

Public Const MAX_SHOPS = 255
'Public Const MAX_SHOPS = 10

Public Const MAX_PLAYER_SPELLS = 20

Public Const MAX_SPELLS = 255
'Public Const MAX_SPELLS = 10

Public Const MAX_TRADES = 15
Public Const MAX_GUILDS = 20
Public Const MAX_GUILD_MEMBERS = 10
Public Const MAX_PARTIES = MAX_PLAYERS

Public Const NO = 0
Public Const YES = 1

' Account constants
Public Const NAME_LENGTH = 20
Public Const MAX_CHARS = 3

' Sex constants
Public Const SEX_MALE = 0
Public Const SEX_FEMALE = 1

' Map constants

Public Const MAX_MAPS = 10000
'Public Const MAX_MAPS = 1000

Public Const MAX_MAPX = 15
Public Const MAX_MAPY = 11
Public Const MAP_MORAL_NONE = 0
Public Const MAP_MORAL_SAFE = 1
Public Const MAP_MORAL_ARENA = 2
Public Const MAP_MORAL_HOUSE = 3

' Image constants
Public Const PIC_X = 32
Public Const PIC_Y = 32

' Tile consants
Public Const TILE_TYPE_WALKABLE = 0
Public Const TILE_TYPE_BLOCKED = 1
Public Const TILE_TYPE_WARP = 2
Public Const TILE_TYPE_ITEM = 3
Public Const TILE_TYPE_NPCAVOID = 4
Public Const TILE_TYPE_KEY = 5
Public Const TILE_TYPE_KEYOPEN = 6
Public Const TILE_TYPE_SIGN = 7
Public Const TILE_TYPE_FULLHEAL = 8
Public Const TILE_TYPE_DEATH = 9
Public Const TILE_TYPE_SFX = 10
Public Const TILE_TYPE_QUEST = 11
Public Const TILE_TYPE_ARENA = 12
Public Const TILE_TYPE_FISHING = 13
Public Const TILE_TYPE_MINING = 14


' Item constants
Public Const ITEM_TYPE_NONE = 0
Public Const ITEM_TYPE_WEAPON = 1
Public Const ITEM_TYPE_ARMOR = 2
Public Const ITEM_TYPE_HELMET = 3
Public Const ITEM_TYPE_SHIELD = 4
Public Const ITEM_TYPE_POTIONADDHP = 5
Public Const ITEM_TYPE_POTIONADDMP = 6
Public Const ITEM_TYPE_POTIONADDSP = 7
Public Const ITEM_TYPE_POTIONSUBHP = 8
Public Const ITEM_TYPE_POTIONSUBMP = 9
Public Const ITEM_TYPE_POTIONSUBSP = 10
Public Const ITEM_TYPE_KEY = 11
Public Const ITEM_TYPE_CURRENCY = 12
Public Const ITEM_TYPE_SPELL = 13
Public Const ITEM_TYPE_SCHANGE = 14
Public Const ITEM_TYPE_BOOK = 15
Public Const ITEM_TYPE_RANDITEM = 16
Public Const ITEM_TYPE_MASKEDIT = 17
Public Const ITEM_TYPE_PETEGG = 18
Public Const ITEM_TYPE_BARBER = 19


' Direction constants
Public Const DIR_UP = 0
Public Const DIR_DOWN = 1
Public Const DIR_LEFT = 2
Public Const DIR_RIGHT = 3

' Constants for player movement
Public Const MOVING_WALKING = 1
Public Const MOVING_RUNNING = 2

' Weather constants
Public Const WEATHER_NONE = 0
Public Const WEATHER_RAINING = 1
Public Const WEATHER_SNOWING = 2

' Time constants
Public Const TIME_DAY = 0
Public Const TIME_NIGHT = 1

' Admin constants
Public Const ADMIN_MONITER = 1
Public Const ADMIN_MAPPER = 2
Public Const ADMIN_DEVELOPER = 3
Public Const ADMIN_CREATOR = 4
Public Const ADMIN_LORD = 5

' NPC constants
Public Const NPC_BEHAVIOR_ATTACKONSIGHT = 0
Public Const NPC_BEHAVIOR_ATTACKWHENATTACKED = 1
Public Const NPC_BEHAVIOR_FRIENDLY = 2
Public Const NPC_BEHAVIOR_SHOPKEEPER = 3
Public Const NPC_BEHAVIOR_GUARD = 4

' Spell constants
Public Const SPELL_TYPE_ADDHP = 0
Public Const SPELL_TYPE_ADDMP = 1
Public Const SPELL_TYPE_ADDSP = 2
Public Const SPELL_TYPE_SUBHP = 3
Public Const SPELL_TYPE_SUBMP = 4
Public Const SPELL_TYPE_SUBSP = 5
Public Const SPELL_TYPE_GIVEITEM = 6
Public Const SPELL_TYPE_CRAFT = 7

' Target type constants
Public Const TARGET_TYPE_PLAYER = 0
Public Const TARGET_TYPE_NPC = 1

Type PlayerInvRec
    Num As Integer
    Value As Long
    Dur As Long
End Type
Type PlayerBankRec
    Num As Integer
    Value As Long
    Dur As Long
End Type

Type PartyRec
NumParty As Byte
Player1 As Long
Player2 As Long
Player3 As Long
Player4 As Long
PartyLeader As Long
End Type

Type PlayerRec
    ' General
    Name As String * NAME_LENGTH
    Sex As Byte
    Class As Byte
    Race As Byte
    SPRITE As Long
    SPRITE2 As Long
    SPRITE3 As Long
    SPRITE4 As Long
    spellsprite As Integer
    spellframe As Integer
    PET As Integer
    Level As Long
    exp As Long
    Access As Byte
    AdminCmds As String
    PK As Byte
    GlobalTemp As Byte
    Guild As String * NAME_LENGTH
    GuildRank As Byte
    faction As Byte
    
    ' Vitals
    HP As Long
    MP As Long
    SP As Long
    
    ' Stats
    STR As Long
    DEF As Long
    SPeed As Long
    MAGI As Long
    Fishing As Long
    Mining As Long
    Crafting As Long
    POINTS As Long
    ANONYMOUS As Byte
    
    ' Worn equipment
    ArmorSlot As Long
    WeaponSlot As Long
    HelmetSlot As Long
    ShieldSlot As Long
    
    ' Inventory
    Inv(1 To MAX_INV) As PlayerInvRec
    Bank(1 To MAX_BANK) As PlayerBankRec
    spell(1 To MAX_PLAYER_SPELLS) As Long
    
    ' Position
    MAP As Integer
    X As Byte
    y As Byte
    Dir As Byte
    
    BindMap As Integer
    BindX As Byte
    BindY As Byte
    
    'Options
    ooc As Byte
    tc As Byte
End Type
    
    


Type AccountRec
    ' Account
    Login As String * NAME_LENGTH
    Password As String * NAME_LENGTH
    GlobalPriv As Integer
       
    ' Characters (we use 0 to prevent a crash that still needs to be figured out)
    Char(0 To MAX_CHARS) As PlayerRec
    
    ' None saved local vars
    Buffer As String
    IncBuffer As String
    charnum As Byte
    InGame As Boolean
    AttackTimer As Long
    DataTimer As Long
    StuckTimer As Long
    FishingTimer As Long
    MiningTimer As Long
    GlobalChatTimer As Long
    TradeChatTimer As Long
    DataBytes As Long
    DataPackets As Long
    PartyLeader As Long
    PartyPlayer As Long
    PartyPlayer2 As Long
    PartyPlayer3 As Long
    nPartyPlayer As Long
    nPartyPlayer2 As Long
    nPartyPlayer3 As Long
    PartyNum As Byte
    InParty As Byte
    TargetType As Byte
    Target As Byte
    CastedSpell As Byte
    PartyStarter As Byte
    GettingMap As Byte
    InviteGuild As Long
    playertrade As Long
    playertradeitem As Long
    playertradeamount As Long
    playertradeagree As Long
    playertotradewith As Long
    Party As Long
End Type

Type TileRec
    Ground As Integer
    Mask As Integer
    Anim As Integer
    Mask2 As Integer
    M2Anim As Integer
    Fringe As Integer
    FAnim As Integer
    Fringe2 As Integer
    F2Anim As Integer
    Type As Byte
    Data1 As String
    Data2 As String
    Data3 As String
End Type

Type OldMapRec
    Name As String * NAME_LENGTH
    Revision As Long
    Moral As Byte
    Up As Integer
    Down As Integer
    Left As Integer
    Right As Integer
    Music As Integer
    BootMap As Integer
    BootX As Byte
    BootY As Byte
    Shop As Byte
    Indoors As Byte
    Tile(0 To MAX_MAPX, 0 To MAX_MAPY) As TileRec
    Npc(1 To MAX_MAP_NPCS) As Byte
End Type

Type MapRec
    Name As String * NAME_LENGTH
    Revision As Long
    Moral As Byte
    Up As Integer
    Down As Integer
    Left As Integer
    Right As Integer
    Music As Integer
    BootMap As Integer
    BootX As Byte
    BootY As Byte
    Shop As Byte
    Indoors As Byte
    Tile(0 To MAX_MAPX, 0 To MAX_MAPY) As TileRec
    Npc(1 To MAX_MAP_NPCS) As Integer
End Type

Type ClassRec
    Name As String * NAME_LENGTH
    
    SPRITE As Long
    SPRITE2 As Long
    SPRITE3 As Long
    SPRITE4 As Long
    
    STR As Byte
    DEF As Byte
    SPeed As Byte
    MAGI As Byte
End Type

Type RaceRec
    Name As String * NAME_LENGTH
    
    SPRITE As Long
    SPRITE2 As Long
    SPRITE3 As Long
    SPRITE4 As Long
    
    STR As Byte
    DEF As Byte
    SPeed As Byte
    MAGI As Byte
End Type

Type ItemRec
    Name As String * 25
    Class As Integer
    Pic As Integer
    Type As Byte
    Data1 As String
    Data2 As String
    Data3 As Integer
    Data4 As Integer
    STRmod As Integer
    DEFmod As Integer
    MAGImod As Integer
    SPRITE As Integer
    spdmod As Integer
    SellValue As Long
    NoDrop As Byte
End Type

Type MapItemRec
    Num As Integer
    Value As Long
    Dur As Integer
    
    X As Byte
    y As Byte
End Type

Type NpcRec
    Name As String * NAME_LENGTH
    AttackSay As String * 300
    
    SPRITE As Long
    SPRITE2 As Long
    SPRITE3 As Long
    SPRITE4 As Long
    spellsprite As Integer
    spellframe As Integer
    SpawnSecs As Long
    Behavior As Byte
    Range As Byte
    
    DropChance As Integer
    DropItem As Integer
    DropItemValue As Integer
    
    DropChance2 As Integer
    DropItem2 As Integer
    DropItemValue2 As Integer
    
    STR  As Integer
    DEF As Integer
    SPeed As Integer
    MAGI As Integer
End Type

Type MapNpcRec
    Num As Integer
    
    Target As Integer
    
    HP As Long
    MP As Long
    SP As Long
        
    X As Byte
    y As Byte
    Dir As Integer
    
    ' For server use only
    SpawnWait As Long
    AttackTimer As Long
End Type

Type TradeItemRec
    GiveItem As Long
    GiveValue As Long
    GetItem As Long
    GetValue As Long
End Type

Type ShopRec
    Name As String * NAME_LENGTH
    JoinSay As String * 255
    LeaveSay As String * 255
    FixesItems As Byte
    OneSale As Byte
    TradeItem(1 To MAX_TRADES) As TradeItemRec
End Type
    
Type SpellRec
    Name As String * NAME_LENGTH
    ClassReq As Byte
    LevelReq As Byte
    Type As Byte
    Data1 As Integer
    Data2 As Integer
    Data3 As Integer
    MPused As Integer
    Range As Integer
    Usedef As Boolean
    Sfx As String
    Gfx As Integer
End Type

Type TempTileRec
    DoorOpen(0 To MAX_MAPX, 0 To MAX_MAPY)  As Byte
    DoorTimer As Long
    MineralTimer(0 To MAX_MAPX, 0 To MAX_MAPY) As Long
    MineralCount(0 To MAX_MAPX, 0 To MAX_MAPY)  As Byte
End Type

Type GuildRec
    Name As String * NAME_LENGTH
    Founder As String * NAME_LENGTH
    Member(1 To MAX_GUILD_MEMBERS) As String * NAME_LENGTH
End Type

' Used for parsing
Public SEP_CHAR As String * 1
Public END_CHAR As String * 1

' Maximum classes
Public Max_Classes As Byte

' Maximun races
Public Max_Races As Byte

Public MAP(1 To MAX_MAPS) As MapRec
Public TempTile(1 To MAX_MAPS) As TempTileRec
Public PlayersOnMap(1 To MAX_MAPS) As Long
Public player(1 To MAX_PLAYERS) As AccountRec
Public Class() As ClassRec
Public Race() As RaceRec
Public Item(0 To MAX_ITEMS) As ItemRec
Public Npc(0 To MAX_NPCS) As NpcRec
Public MapItem(1 To MAX_MAPS, 1 To MAX_MAP_ITEMS) As MapItemRec
Public MapNpc(1 To MAX_MAPS, 1 To MAX_MAP_NPCS) As MapNpcRec
Public Shop(1 To MAX_SHOPS) As ShopRec
Public spell(1 To MAX_SPELLS) As SpellRec
Public Guild(1 To MAX_GUILDS) As GuildRec
Public Parties(1 To MAX_PARTIES) As PartyRec

Sub ClearTempTile()


Dim i As Long, y As Long, X As Long

    For i = 1 To MAX_MAPS
        TempTile(i).DoorTimer = 0
        
        For y = 0 To MAX_MAPY
            For X = 0 To MAX_MAPX
                TempTile(i).DoorOpen(X, y) = NO
                TempTile(i).MineralCount(X, y) = 0
                TempTile(i).MineralTimer(X, y) = 0
            Next X
        Next y
    Next i
End Sub

Sub ClearRaces()


Dim i As Long

    For i = 0 To Max_Races
        Race(i).Name = ""
        Race(i).STR = 0
        Race(i).DEF = 0
        Race(i).SPeed = 0
        Race(i).MAGI = 0
    Next i
End Sub

Sub ClearClasses()


Dim i As Long

    For i = 0 To Max_Classes
        Class(i).Name = ""
        Class(i).STR = 0
        Class(i).DEF = 0
        Class(i).SPeed = 0
        Class(i).MAGI = 0
    Next i
End Sub
Sub ClearParties()


Dim i As Long

For i = 1 To MAX_PARTIES
    Parties(i).NumParty = 0
    Parties(i).PartyLeader = 0
    Parties(i).Player1 = 0
    Parties(i).Player2 = 0
    Parties(i).Player3 = 0
    Parties(i).Player4 = 0
Next i


End Sub
Sub ClearPlayer(ByVal index As Long)


Dim i As Long
Dim n As Long

    player(index).Login = ""
    player(index).Password = ""
    
    For i = 1 To MAX_CHARS
        player(index).Char(i).Name = ""
        player(index).Char(i).Class = 0
        player(index).Char(i).Race = 0
        player(index).Char(i).Level = 0
        player(index).Char(i).SPRITE = 0
        player(index).Char(i).exp = 0
        player(index).Char(i).Access = 0
        player(index).Char(i).PK = NO
        player(index).Char(i).POINTS = 0
        player(index).Char(i).ANONYMOUS = 0
        player(index).Char(i).Guild = ""
        player(index).Char(i).GuildRank = 0
        player(index).Char(i).faction = 0
        player(index).InviteGuild = 0
        
        player(index).Char(i).HP = 0
        player(index).Char(i).MP = 0
        player(index).Char(i).SP = 0
        
        player(index).Char(i).STR = 0
        player(index).Char(i).DEF = 0
        player(index).Char(i).SPeed = 0
        player(index).Char(i).MAGI = 0
        player(index).Char(i).Crafting = 0
        player(index).Char(i).Mining = 0
        player(index).Char(i).Fishing = 0
        
        For n = 1 To MAX_INV
            player(index).Char(i).Inv(n).Num = 0
            player(index).Char(i).Inv(n).Value = 0
            player(index).Char(i).Inv(n).Dur = 0
        Next n
        
        For n = 1 To MAX_BANK
            player(index).Char(i).Bank(n).Num = 0
            player(index).Char(i).Bank(n).Value = 0
            player(index).Char(i).Bank(n).Dur = 0
        Next n
        
        For n = 1 To MAX_PLAYER_SPELLS
            player(index).Char(i).spell(n) = 0
        Next n
        
        player(index).Char(i).ArmorSlot = 0
        player(index).Char(i).WeaponSlot = 0
        player(index).Char(i).HelmetSlot = 0
        player(index).Char(i).ShieldSlot = 0
        
        player(index).Char(i).MAP = 0
        player(index).Char(i).X = 0
        player(index).Char(i).y = 0
        player(index).Char(i).Dir = 0
        
        player(index).Char(i).BindMap = 0
        player(index).Char(i).BindX = 0
        player(index).Char(i).BindY = 0
        
        player(index).Char(i).ooc = 0
        player(index).Char(i).tc = 0
        
        ' Temporary vars
        player(index).Buffer = ""
        player(index).IncBuffer = ""
        player(index).charnum = 0
        player(index).InGame = False
        player(index).AttackTimer = 0
        player(index).StuckTimer = 0
        player(index).GlobalChatTimer = 0
        player(index).TradeChatTimer = 0
        player(index).FishingTimer = 0
        player(index).MiningTimer = 0
        player(index).DataTimer = 0
        player(index).DataBytes = 0
        player(index).DataPackets = 0
        player(index).Party = 0
        player(index).PartyPlayer = 0
        player(index).PartyPlayer2 = 0
        player(index).PartyPlayer3 = 0
        player(index).nPartyPlayer = 0
        player(index).nPartyPlayer2 = 0
        player(index).nPartyPlayer3 = 0
        player(index).PartyNum = 0
        player(index).InParty = NO
        player(index).Target = 0
        player(index).TargetType = 0
        player(index).CastedSpell = NO
        player(index).PartyStarter = NO
        player(index).GettingMap = NO
        
    Next i
End Sub

Sub ClearChar(ByVal index As Long, ByVal charnum As Long)


Dim n As Long
    
    player(index).Char(charnum).Name = ""
    player(index).Char(charnum).Class = 0
    player(index).Char(charnum).Race = 0
    player(index).Char(charnum).SPRITE = 0
    player(index).Char(charnum).SPRITE2 = 5000
    player(index).Char(charnum).SPRITE3 = 5000
    player(index).Char(charnum).SPRITE4 = 5000
    player(index).Char(charnum).PET = 5000
    player(index).Char(charnum).Level = 0
    player(index).Char(charnum).exp = 0
    player(index).Char(charnum).Access = 0
    player(index).Char(charnum).PK = NO
    player(index).Char(charnum).POINTS = 0
    player(index).Char(charnum).ANONYMOUS = 0
    player(index).Char(charnum).Guild = ""
    player(index).Char(charnum).GuildRank = 0
    player(index).Char(charnum).faction = 0
    
    player(index).Char(charnum).HP = 0
    player(index).Char(charnum).MP = 0
    player(index).Char(charnum).SP = 0
    
    player(index).Char(charnum).STR = 0
    player(index).Char(charnum).DEF = 0
    player(index).Char(charnum).SPeed = 0
    player(index).Char(charnum).MAGI = 0
    player(index).Char(charnum).Crafting = 0
    player(index).Char(charnum).Fishing = 0
    player(index).Char(charnum).Mining = 0
    
    For n = 1 To MAX_INV
        player(index).Char(charnum).Inv(n).Num = 0
        player(index).Char(charnum).Inv(n).Value = 0
        player(index).Char(charnum).Inv(n).Dur = 0
    Next n
    
    For n = 1 To MAX_BANK
        player(index).Char(charnum).Bank(n).Num = 0
        player(index).Char(charnum).Bank(n).Value = 0
        player(index).Char(charnum).Bank(n).Dur = 0
    Next n
    
    For n = 1 To MAX_PLAYER_SPELLS
        player(index).Char(charnum).spell(n) = 0
    Next n
    
    player(index).Char(charnum).ArmorSlot = 0
    player(index).Char(charnum).WeaponSlot = 0
    player(index).Char(charnum).HelmetSlot = 0
    player(index).Char(charnum).ShieldSlot = 0
    
    player(index).Char(charnum).MAP = 0
    player(index).Char(charnum).X = 0
    player(index).Char(charnum).y = 0
    player(index).Char(charnum).Dir = 0
    
    player(index).Char(charnum).BindMap = 0
    player(index).Char(charnum).BindX = 0
    player(index).Char(charnum).BindY = 0
    player(index).Char(charnum).ooc = 0
    player(index).Char(charnum).tc = 0
End Sub
    
Sub ClearItem(ByVal index As Long)


    Item(index).Name = ""
    Item(index).Class = 0
    Item(index).Type = 0
    Item(index).Data1 = 0
    Item(index).Data2 = 0
    Item(index).Data3 = 0
    Item(index).Data4 = 0
    Item(index).Class = 0
    Item(index).STRmod = 0
    Item(index).DEFmod = 0
    Item(index).MAGImod = 0
End Sub

Sub ClearItems()


Dim i As Long

    For i = 1 To MAX_ITEMS
        Call ClearItem(i)
    Next i
End Sub

Sub ClearNpc(ByVal index As Long)


    Npc(index).Name = ""
    Npc(index).AttackSay = ""
    Npc(index).SPRITE = 0
    Npc(index).SpawnSecs = 0
    Npc(index).Behavior = 0
    Npc(index).Range = 0
    Npc(index).DropChance = 0
    Npc(index).DropItem = 0
    Npc(index).DropItemValue = 0
    Npc(index).STR = 0
    Npc(index).DEF = 0
    Npc(index).SPeed = 0
    Npc(index).MAGI = 0
End Sub

Sub ClearNpcs()


Dim i As Long

    For i = 1 To MAX_NPCS
        Call ClearNpc(i)
    Next i
End Sub

Sub ClearMapItem(ByVal index As Long, ByVal MapNum As Long)


    MapItem(MapNum, index).Num = 0
    MapItem(MapNum, index).Value = 0
    MapItem(MapNum, index).Dur = 0
    MapItem(MapNum, index).X = 0
    MapItem(MapNum, index).y = 0
End Sub

Sub ClearMapItems()


Dim X As Long
Dim y As Long

    For y = 1 To MAX_MAPS
        For X = 1 To MAX_MAP_ITEMS
            Call ClearMapItem(X, y)
        Next X
    Next y
End Sub

Sub ClearMapNpc(ByVal index As Long, ByVal MapNum As Long)


    MapNpc(MapNum, index).Num = 0
    MapNpc(MapNum, index).Target = 0
    MapNpc(MapNum, index).HP = 0
    MapNpc(MapNum, index).MP = 0
    MapNpc(MapNum, index).SP = 0
    MapNpc(MapNum, index).X = 0
    MapNpc(MapNum, index).y = 0
    MapNpc(MapNum, index).Dir = 0
    
    ' Server use only
    MapNpc(MapNum, index).SpawnWait = 0
    MapNpc(MapNum, index).AttackTimer = 0
End Sub

Sub ClearMapNpcs()


Dim X As Long
Dim y As Long

    For y = 1 To MAX_MAPS
        For X = 1 To MAX_MAP_NPCS
            Call ClearMapNpc(X, y)
        Next X
    Next y
End Sub

Sub ClearMap(ByVal MapNum As Long)


Dim i As Long
Dim X As Long
Dim y As Long

    MAP(MapNum).Name = ""
    MAP(MapNum).Revision = 0
    MAP(MapNum).Moral = 0
    MAP(MapNum).Up = 0
    MAP(MapNum).Down = 0
    MAP(MapNum).Left = 0
    MAP(MapNum).Right = 0
        
    For y = 0 To MAX_MAPY
        For X = 0 To MAX_MAPX
            MAP(MapNum).Tile(X, y).Ground = 0
            MAP(MapNum).Tile(X, y).Mask = 0
            MAP(MapNum).Tile(X, y).Anim = 0
            MAP(MapNum).Tile(X, y).Mask2 = 0
            MAP(MapNum).Tile(X, y).M2Anim = 0
            MAP(MapNum).Tile(X, y).Fringe = 0
            MAP(MapNum).Tile(X, y).FAnim = 0
            MAP(MapNum).Tile(X, y).Fringe2 = 0
            MAP(MapNum).Tile(X, y).F2Anim = 0
            MAP(MapNum).Tile(X, y).Type = 0
            MAP(MapNum).Tile(X, y).Data1 = 0
            MAP(MapNum).Tile(X, y).Data2 = 0
            MAP(MapNum).Tile(X, y).Data3 = 0
        Next X
    Next y
    
    ' Reset the values for if a player is on the map or not
    PlayersOnMap(MapNum) = NO
End Sub

Sub ClearMaps()


Dim i As Long

    For i = 1 To MAX_MAPS
        Call ClearMap(i)
    Next i
End Sub

Sub ClearShop(ByVal index As Long)


Dim i As Long

    Shop(index).Name = ""
    Shop(index).JoinSay = ""
    Shop(index).LeaveSay = ""
    
    For i = 1 To MAX_TRADES
        Shop(index).TradeItem(i).GiveItem = 0
        Shop(index).TradeItem(i).GiveValue = 0
        Shop(index).TradeItem(i).GetItem = 0
        Shop(index).TradeItem(i).GetValue = 0
    Next i
End Sub

Sub ClearShops()


Dim i As Long

    For i = 1 To MAX_SHOPS
        Call ClearShop(i)
    Next i
End Sub
Sub DeathSound(ByVal MapD As Integer, ByVal Class As Integer, ByVal SexD As Integer)


Dim Packet As String
If Class = 0 And SexD = 0 Then Packet = "PLAYSFX" & SEP_CHAR & "AncientMD" & SEP_CHAR & END_CHAR
If Class = 1 And SexD = 0 Then Packet = "PLAYSFX" & SEP_CHAR & "MedicMD" & SEP_CHAR & END_CHAR
If Class = 2 And SexD = 0 Then Packet = "PLAYSFX" & SEP_CHAR & "MarineMD" & SEP_CHAR & END_CHAR
If Class = 0 And SexD = 1 Then Packet = "PLAYSFX" & SEP_CHAR & "AncientFD" & SEP_CHAR & END_CHAR
If Class = 1 And SexD = 1 Then Packet = "PLAYSFX" & SEP_CHAR & "MedicFD" & SEP_CHAR & END_CHAR
If Class = 2 And SexD = 1 Then Packet = "PLAYSFX" & SEP_CHAR & "MarineFD" & SEP_CHAR & END_CHAR
If Class = 3 And SexD = 0 Then Packet = "PLAYSFX" & SEP_CHAR & "AncientMD" & SEP_CHAR & END_CHAR
If Class = 4 And SexD = 0 Then Packet = "PLAYSFX" & SEP_CHAR & "MedicMD" & SEP_CHAR & END_CHAR
If Class = 5 And SexD = 0 Then Packet = "PLAYSFX" & SEP_CHAR & "MarineMD" & SEP_CHAR & END_CHAR
If Class = 3 And SexD = 1 Then Packet = "PLAYSFX" & SEP_CHAR & "AncientFD" & SEP_CHAR & END_CHAR
If Class = 4 And SexD = 1 Then Packet = "PLAYSFX" & SEP_CHAR & "MedicFD" & SEP_CHAR & END_CHAR
If Class = 5 And SexD = 1 Then Packet = "PLAYSFX" & SEP_CHAR & "MarineFD" & SEP_CHAR & END_CHAR
If Class = 6 And SexD = 0 Then Packet = "PLAYSFX" & SEP_CHAR & "AncientMD" & SEP_CHAR & END_CHAR
If Class = 7 And SexD = 0 Then Packet = "PLAYSFX" & SEP_CHAR & "MedicMD" & SEP_CHAR & END_CHAR
If Class = 6 And SexD = 1 Then Packet = "PLAYSFX" & SEP_CHAR & "AncientFD" & SEP_CHAR & END_CHAR
If Class = 7 And SexD = 1 Then Packet = "PLAYSFX" & SEP_CHAR & "MedicFD" & SEP_CHAR & END_CHAR
If Class = 99 And SexD = 99 Then Packet = "PLAYSFX" & SEP_CHAR & "LvLUp" & SEP_CHAR & END_CHAR
Call SendDataToMap(MapD, Packet)
End Sub
Sub WeapSound(ByVal player As Integer, ByVal WeapItem As Integer)


Dim Packet As String
Packet = "PLAYSFX" & SEP_CHAR & "Weap" & Val((Item(WeapItem).Data3)) & SEP_CHAR & END_CHAR
Call SendDataTo(player, Packet)
End Sub
Sub SpellSound(ByVal player As Integer, ByVal Soundfile As String)


Dim Packet As String
Packet = "PLAYSFX" & SEP_CHAR & Soundfile & SEP_CHAR & END_CHAR
Call SendDataToMap(GetPlayerMap(player), Packet)
End Sub
Sub NPCSound(ByVal MAP As Integer, ByVal index As Integer)


Dim Packet As String
Packet = "PLAYSFX" & SEP_CHAR & "npc" & Val(Npc(index).SPeed) & SEP_CHAR & END_CHAR
Call SendDataToMap(MAP, Packet)
End Sub
Function GetPlayerSex(ByVal index As Integer)


GetPlayerSex = player(index).Char(player(index).charnum).Sex
End Function
Sub ClearSpell(ByVal index As Long)


    spell(index).Name = ""
    spell(index).ClassReq = 0
    spell(index).LevelReq = 0
    spell(index).Type = 0
    spell(index).Data1 = 0
    spell(index).Data2 = 0
    spell(index).Data3 = 0
End Sub

Sub ClearSpells()


Dim i As Long

    For i = 1 To MAX_SPELLS
        Call ClearSpell(i)
    Next i
End Sub




' //////////////////////
' // PLAYER FUNCTIONS //
' //////////////////////

Function GetPlayerLogin(ByVal index As Long) As String


    GetPlayerLogin = Trim(player(index).Login)
End Function

Sub SetPlayerLogin(ByVal index As Long, ByVal Login As String)


    player(index).Login = Login
End Sub

Function GetPlayerPassword(ByVal index As Long) As String


    GetPlayerPassword = Trim(player(index).Password)
End Function

Sub SetPlayerPassword(ByVal index As Long, ByVal Password As String)


    player(index).Password = Password
End Sub

Function GetPlayerName(ByVal index As Long) As String


    GetPlayerName = Trim(player(index).Char(player(index).charnum).Name)
End Function

Sub SetPlayerName(ByVal index As Long, ByVal Name As String)


    player(index).Char(player(index).charnum).Name = Name
End Sub

Function GetPlayerClass(ByVal index As Long) As Long


    GetPlayerClass = player(index).Char(player(index).charnum).Class
End Function

Sub SetPlayerClass(ByVal index As Long, ByVal ClassNum As Long)


    player(index).Char(player(index).charnum).Class = ClassNum
End Sub

Function GetPlayerRace(ByVal index As Long) As Long


    GetPlayerRace = player(index).Char(player(index).charnum).Race
End Function

Sub SetPlayerRace(ByVal index As Long, ByVal RaceNum As Long)


    player(index).Char(player(index).charnum).Race = RaceNum
End Sub

Function GetPlayerSprite(ByVal index As Long) As Long


    GetPlayerSprite = player(index).Char(player(index).charnum).SPRITE
End Function
Function GetPlayerSprite2(ByVal index As Long) As Long


    GetPlayerSprite2 = player(index).Char(player(index).charnum).SPRITE2
End Function
Function GetPlayerSprite3(ByVal index As Long) As Long


    GetPlayerSprite3 = player(index).Char(player(index).charnum).SPRITE3
End Function
Function GetPlayerSprite4(ByVal index As Long) As Long


    GetPlayerSprite4 = player(index).Char(player(index).charnum).SPRITE4
End Function

Function GetPet(ByVal index As Long) As Long


    GetPet = player(index).Char(player(index).charnum).PET
End Function

Sub SetPlayerSprite(ByVal index As Long, ByVal SPRITE As Long)


    player(index).Char(player(index).charnum).SPRITE = SPRITE
End Sub
Sub SetPlayerSprite2(ByVal index As Long, ByVal SPRITE As Long)


    player(index).Char(player(index).charnum).SPRITE2 = SPRITE
End Sub
Sub SetPlayerSprite3(ByVal index As Long, ByVal SPRITE As Long)


    player(index).Char(player(index).charnum).SPRITE3 = SPRITE
End Sub
Sub SetPlayerSprite4(ByVal index As Long, ByVal SPRITE As Long)


    player(index).Char(player(index).charnum).SPRITE4 = SPRITE
End Sub
Sub SetPet(ByVal index As Long, ByVal SPRITE As Long)


    player(index).Char(player(index).charnum).PET = SPRITE
End Sub

Function GetPlayerLevel(ByVal index As Long) As Long


    GetPlayerLevel = player(index).Char(player(index).charnum).Level
End Function

Sub SetPlayerLevel(ByVal index As Long, ByVal Level As Long)


    player(index).Char(player(index).charnum).Level = Level
End Sub

Function GetPlayerNextLevel(ByVal index As Long) As Long


Dim i As Long

Select Case GetPlayerClass(index)
' knight (201)
Case 0
GetPlayerNextLevel = (GetPlayerLevel(index) * (GetPlayerLevel(index) * 0.7)) * 201 * 1.25
' Pally (181)
Case 1
GetPlayerNextLevel = (GetPlayerLevel(index) * (GetPlayerLevel(index) * 0.7)) * 181 * 1.25
' Wizard (171)
Case 2
GetPlayerNextLevel = (GetPlayerLevel(index) * (GetPlayerLevel(index) * 0.7)) * 171 * 1.25
' Cleric (171)
Case 3
GetPlayerNextLevel = (GetPlayerLevel(index) * (GetPlayerLevel(index) * 0.7)) * 171 * 1.25
' Dragoon (201)
Case 4
GetPlayerNextLevel = (GetPlayerLevel(index) * (GetPlayerLevel(index) * 0.7)) * 201 * 1.25
' Assassin (181)
Case 5
GetPlayerNextLevel = (GetPlayerLevel(index) * (GetPlayerLevel(index) * 0.7)) * 181 * 1.25
' Necromancer
Case 6
GetPlayerNextLevel = (GetPlayerLevel(index) * (GetPlayerLevel(index) * 0.7)) * 171 * 1.25
' Druid
Case 7
GetPlayerNextLevel = (GetPlayerLevel(index) * (GetPlayerLevel(index) * 0.7)) * 171 * 1.25
End Select

'i = GetPlayerLevel(Index) / 0.75
'If i < 1 Then i = 1
'    GetPlayerNextLevel = (GetPlayerLevel(index) * (GetPlayerLevel(index) * 0.7)) * 163
End Function

Function GetPlayerExp(ByVal index As Long) As Long


    GetPlayerExp = player(index).Char(player(index).charnum).exp
End Function

Sub SetPlayerExp(ByVal index As Long, ByVal exp As Long)


    player(index).Char(player(index).charnum).exp = exp
    Call CheckPlayerLevelUp(index)
End Sub

Function GetPlayerAccess(ByVal index As Long) As Long


    GetPlayerAccess = player(index).Char(player(index).charnum).Access
End Function

Sub SetPlayerAccess(ByVal index As Long, ByVal Access As Long)


    player(index).Char(player(index).charnum).Access = Access
End Sub
Function GetNpcLevel(ByVal NpcNum As Long) As Long


Dim STR As Long, DEF As Long
        

        DEF = Npc(NpcNum).DEF
        
        GetNpcLevel = (DEF - (DEF * 0.25)) * 1.2
End Function

Function GetPlayerPK(ByVal index As Long) As Long


    GetPlayerPK = player(index).Char(player(index).charnum).PK
End Function

Sub SetPlayerPK(ByVal index As Long, ByVal PK As Long)


    player(index).Char(player(index).charnum).PK = PK
End Sub

Function GetPlayerHP(ByVal index As Long) As Long


    GetPlayerHP = player(index).Char(player(index).charnum).HP
End Function

Sub SetPlayerHP(ByVal index As Long, ByVal HP As Long)


    player(index).Char(player(index).charnum).HP = HP
    
    If GetPlayerHP(index) > GetPlayerMaxHP(index) Then
        player(index).Char(player(index).charnum).HP = GetPlayerMaxHP(index)
    End If
    If GetPlayerHP(index) < 0 Then
        player(index).Char(player(index).charnum).HP = 0
    End If
End Sub

Function GetPlayerMP(ByVal index As Long) As Long


    GetPlayerMP = player(index).Char(player(index).charnum).MP
End Function

Sub SetPlayerMP(ByVal index As Long, ByVal MP As Long)


    player(index).Char(player(index).charnum).MP = MP

    If GetPlayerMP(index) > GetPlayerMaxMP(index) Then
        player(index).Char(player(index).charnum).MP = GetPlayerMaxMP(index)
    End If
    If GetPlayerMP(index) < 0 Then
        player(index).Char(player(index).charnum).MP = 0
    End If
End Sub

Function GetPlayerSP(ByVal index As Long) As Long


    GetPlayerSP = player(index).Char(player(index).charnum).SP
End Function
Function GetPlayerOocSwitch(ByVal index As Long) As Long


    GetPlayerOocSwitch = player(index).Char(player(index).charnum).ooc
End Function
Function GetPlayertcSwitch(ByVal index As Long) As Long


    GetPlayertcSwitch = player(index).Char(player(index).charnum).tc
End Function
Sub SetPlayerOoc(ByVal index As Long, ByVal switch As String)


If switch = "on" Then
player(index).Char(player(index).charnum).ooc = 1
ElseIf switch = "off" Then
player(index).Char(player(index).charnum).ooc = 0
End If
End Sub
Sub SetPlayertc(ByVal index As Long, ByVal switch As String)


If switch = "on" Then
player(index).Char(player(index).charnum).tc = 1
ElseIf switch = "off" Then
player(index).Char(player(index).charnum).tc = 0
End If
End Sub

Sub SetPlayerSP(ByVal index As Long, ByVal SP As Long)


    player(index).Char(player(index).charnum).SP = SP

    If GetPlayerSP(index) > GetPlayerMaxSP(index) Then
        player(index).Char(player(index).charnum).SP = GetPlayerMaxSP(index)
    End If
    If GetPlayerSP(index) < 0 Then
        player(index).Char(player(index).charnum).SP = 0
    End If
End Sub


Function GetPlayerMaxHP(ByVal index As Long) As Long


Dim charnum As Long
Dim i As Long

    charnum = player(index).charnum
    GetPlayerMaxHP = (player(index).Char(charnum).Level + Int(GetPlayerDEF(index) / 2) + Class(player(index).Char(charnum).Class).DEF + Race(player(index).Char(charnum).Race).DEF) * 2
End Function

Function GetPlayerMaxMP(ByVal index As Long) As Long


Dim charnum As Long

    charnum = player(index).charnum
    GetPlayerMaxMP = (player(index).Char(charnum).Level + Int(GetPlayerMAGI(index) / 2) + Class(player(index).Char(charnum).Class).MAGI + Race(player(index).Char(charnum).Race).MAGI) * 2
End Function

Function GetPlayerMaxSP(ByVal index As Long) As Long


Dim charnum As Long

    charnum = player(index).charnum
    GetPlayerMaxSP = ((GetPlayerSTR(index) * 1.5) * (GetPlayerDEF(index) * 0.5)) / 2
End Function

Function GetClassName(ByVal ClassNum As Long) As String


    GetClassName = Trim(Class(ClassNum).Name)
End Function

Function GetClassMaxHP(ByVal ClassNum As Long) As Long


    GetClassMaxHP = (1 + Int(Class(ClassNum).STR / 2) + Class(ClassNum).STR) * 2
End Function

Function GetClassMaxMP(ByVal ClassNum As Long) As Long


    GetClassMaxMP = (1 + Int(Class(ClassNum).MAGI / 2) + Class(ClassNum).MAGI) * 2
End Function

Function GetClassMaxSP(ByVal ClassNum As Long) As Long


    GetClassMaxSP = (1 + Int(Class(ClassNum).SPeed / 2) + Class(ClassNum).SPeed) * 2
End Function

Function GetClassSTR(ByVal ClassNum As Long) As Long


    GetClassSTR = Class(ClassNum).STR
End Function

Function GetClassDEF(ByVal ClassNum As Long) As Long


    GetClassDEF = Class(ClassNum).DEF
End Function

Function GetClassSPEED(ByVal ClassNum As Long) As Long


    GetClassSPEED = Class(ClassNum).SPeed
End Function

Function GetClassMAGI(ByVal ClassNum As Long) As Long


    GetClassMAGI = Class(ClassNum).MAGI
End Function

Function GetRaceName(ByVal RaceNum As Long) As String


    GetRaceName = Trim(Race(RaceNum).Name)
End Function

Function GetRaceMaxHP(ByVal RaceNum As Long) As Long


    GetRaceMaxHP = (1 + Int(Race(RaceNum).STR / 2) + Race(RaceNum).STR) * 2
End Function

Function GetRaceMaxMP(ByVal RaceNum As Long) As Long


    GetRaceMaxMP = (1 + Int(Race(RaceNum).MAGI / 2) + Race(RaceNum).MAGI) * 2
End Function

Function GetRaceMaxSP(ByVal RaceNum As Long) As Long


    GetRaceMaxSP = (1 + Int(Race(RaceNum).SPeed / 2) + Race(RaceNum).SPeed) * 2
End Function

Function GetRaceSTR(ByVal RaceNum As Long) As Long


    GetRaceSTR = Race(RaceNum).STR
End Function

Function GetRaceDEF(ByVal RaceNum As Long) As Long


    GetRaceDEF = Race(RaceNum).DEF
End Function

Function GetRaceSPEED(ByVal RaceNum As Long) As Long


    GetRaceSPEED = Race(RaceNum).SPeed
End Function

Function GetRaceMAGI(ByVal RaceNum As Long) As Long


    GetRaceMAGI = Race(RaceNum).MAGI
End Function


Function GetPlayerSTR(ByVal index As Long) As Long


    GetPlayerSTR = player(index).Char(player(index).charnum).STR
End Function

Sub SetPlayerSTR(ByVal index As Long, ByVal STR As Long)


    player(index).Char(player(index).charnum).STR = STR
End Sub

Function GetPlayerGuild(ByVal index As Long) As String


    GetPlayerGuild = Trim(player(index).Char(player(index).charnum).Guild)
End Function
Function GetPlayerFaction(ByVal index As Long) As Byte


If player(index).Char(player(index).charnum).faction <= 1 Or player(index).Char(player(index).charnum).faction >= 4 Then
GetPlayerFaction = 0
player(index).Char(player(index).charnum).faction = 0
Exit Function
Else
GetPlayerFaction = player(index).Char(player(index).charnum).faction
Exit Function
End If
End Function
Function GetPlayerGuildRank(ByVal index As Long) As String


    GetPlayerGuildRank = player(index).Char(player(index).charnum).GuildRank
End Function
Sub SetPlayerGuild(ByVal index As Long, ByVal Guild As String)


    player(index).Char(player(index).charnum).Guild = Guild
End Sub
Sub SetPlayerGuildRank(ByVal index As Long, ByVal GuildRank As Long)


 If GuildRank < 0 Then GuildRank = 0
 If GuildRank > 10 Then GuildRank = 10
    player(index).Char(player(index).charnum).GuildRank = GuildRank
End Sub
Sub SetPlayerFaction(ByVal index As Long, ByVal faction As Long)


 If faction < 0 Then faction = 0
    player(index).Char(player(index).charnum).faction = faction
End Sub
Function GetPlayerDEF(ByVal index As Long) As Long


    GetPlayerDEF = player(index).Char(player(index).charnum).DEF
End Function

Sub SetPlayerDEF(ByVal index As Long, ByVal DEF As Long)


    player(index).Char(player(index).charnum).DEF = DEF
End Sub

Function GetPlayerSPEED(ByVal index As Long) As Long


    GetPlayerSPEED = player(index).Char(player(index).charnum).SPeed
End Function

Sub SetPlayerSPEED(ByVal index As Long, ByVal SPeed As Long)


    player(index).Char(player(index).charnum).SPeed = SPeed
End Sub

Function GetPlayerMAGI(ByVal index As Long) As Long


    GetPlayerMAGI = player(index).Char(player(index).charnum).MAGI
End Function

Sub SetPlayerMAGI(ByVal index As Long, ByVal MAGI As Long)


If MAGI < 0 Then MAGI = 0
    player(index).Char(player(index).charnum).MAGI = MAGI
End Sub
Function GetPlayerFishing(ByVal index As Long) As Long


    GetPlayerFishing = player(index).Char(player(index).charnum).Fishing
End Function

Sub SetPlayerFishing(ByVal index As Long, ByVal MAGI As Long)


If MAGI < 0 Then MAGI = 0
    If GetPlayerFishing(index) > 100 Then MAGI = GetPlayerFishing(index)
    player(index).Char(player(index).charnum).Fishing = MAGI
End Sub
Function GetPlayerMining(ByVal index As Long) As Long


    GetPlayerMining = player(index).Char(player(index).charnum).Mining
End Function

Sub SetPlayerMining(ByVal index As Long, ByVal MAGI As Long)


If MAGI < 0 Then MAGI = 0
If GetPlayerMining(index) > 100 Then MAGI = GetPlayerMining(index)
    player(index).Char(player(index).charnum).Mining = MAGI
End Sub
Function GetPlayerCraft(ByVal index As Long) As Long


    GetPlayerCraft = player(index).Char(player(index).charnum).Crafting
End Function

Sub SetPlayerCraft(ByVal index As Long, ByVal MAGI As Long)


If MAGI < 0 Then MAGI = 0
    player(index).Char(player(index).charnum).Crafting = MAGI
End Sub

Function GetPlayerPOINTS(ByVal index As Long) As Long


    GetPlayerPOINTS = player(index).Char(player(index).charnum).POINTS
End Function
Function GetPlayerAnonymous(ByVal index As Long) As Byte


    GetPlayerAnonymous = player(index).Char(player(index).charnum).ANONYMOUS
End Function

Sub SetPlayerPOINTS(ByVal index As Long, ByVal POINTS As Long)


    player(index).Char(player(index).charnum).POINTS = POINTS
End Sub
Sub SetPlayeranonymous(ByVal index As Long, ByVal POINTS As Long)


    player(index).Char(player(index).charnum).ANONYMOUS = POINTS
End Sub

Function GetPlayerMap(ByVal index As Long) As Long


    GetPlayerMap = player(index).Char(player(index).charnum).MAP
End Function

Sub SetPlayerMap(ByVal index As Long, ByVal MapNum As Long)


    If MapNum > 0 And MapNum <= MAX_MAPS Then
        player(index).Char(player(index).charnum).MAP = MapNum
    End If
End Sub

Function GetPlayerX(ByVal index As Long) As Long


    GetPlayerX = player(index).Char(player(index).charnum).X
End Function

Sub SetPlayerX(ByVal index As Long, ByVal X As Long)


    player(index).Char(player(index).charnum).X = X
End Sub

Function GetPlayerY(ByVal index As Long) As Long


    GetPlayerY = player(index).Char(player(index).charnum).y
End Function

Sub SetPlayerY(ByVal index As Long, ByVal y As Long)


    player(index).Char(player(index).charnum).y = y
End Sub

Function GetPlayerDir(ByVal index As Long) As Long


    GetPlayerDir = player(index).Char(player(index).charnum).Dir
End Function

Sub SetPlayerDir(ByVal index As Long, ByVal Dir As Long)


    player(index).Char(player(index).charnum).Dir = Dir
End Sub

Function GetPlayerIP(ByVal index As Long) As String


    GetPlayerIP = frmServer.Socket(index).RemoteHostIP
End Function

Function GetPlayerInvItemNum(ByVal index As Long, ByVal InvSlot As Long) As Long


    GetPlayerInvItemNum = player(index).Char(player(index).charnum).Inv(InvSlot).Num
End Function

Sub SetPlayerInvItemNum(ByVal index As Long, ByVal InvSlot As Long, ByVal ItemNum As Long)


    player(index).Char(player(index).charnum).Inv(InvSlot).Num = ItemNum
End Sub

Function GetPlayerInvItemValue(ByVal index As Long, ByVal InvSlot As Long) As Long


    GetPlayerInvItemValue = player(index).Char(player(index).charnum).Inv(InvSlot).Value
End Function

Sub SetPlayerInvItemValue(ByVal index As Long, ByVal InvSlot As Long, ByVal ItemValue As Long)


    player(index).Char(player(index).charnum).Inv(InvSlot).Value = ItemValue
End Sub

Function GetPlayerInvItemDur(ByVal index As Long, ByVal InvSlot As Long) As Long


    GetPlayerInvItemDur = player(index).Char(player(index).charnum).Inv(InvSlot).Dur
End Function

Sub SetPlayerInvItemDur(ByVal index As Long, ByVal InvSlot As Long, ByVal ItemDur As Long)


    player(index).Char(player(index).charnum).Inv(InvSlot).Dur = ItemDur
End Sub

Function GetPlayerSpell(ByVal index As Long, ByVal SpellSlot As Long) As Long


    GetPlayerSpell = player(index).Char(player(index).charnum).spell(SpellSlot)
End Function

Sub SetPlayerSpell(ByVal index As Long, ByVal SpellSlot As Long, ByVal SpellNum As Long)


    player(index).Char(player(index).charnum).spell(SpellSlot) = SpellNum
End Sub
Function GetPlayerBindMap(ByVal index As Long) As Long


    GetPlayerBindMap = player(index).Char(player(index).charnum).BindMap
End Function
Function GetPlayerBindX(ByVal index As Long) As Long


    GetPlayerBindX = player(index).Char(player(index).charnum).BindX
End Function
Function GetPlayerBindY(ByVal index As Long) As Long


    GetPlayerBindY = player(index).Char(player(index).charnum).BindY
End Function
Sub SetPlayerBindMap(ByVal index As Long, ByVal MapNum As Long)


    If MapNum > 0 And MapNum <= MAX_MAPS Then
        player(index).Char(player(index).charnum).BindMap = MapNum
    End If
End Sub
Sub SetPlayerBindX(ByVal index As Long, ByVal X As Long)


    player(index).Char(player(index).charnum).BindX = X
End Sub
Sub SetPlayerBindY(ByVal index As Long, ByVal y As Long)


    player(index).Char(player(index).charnum).BindY = y
End Sub
Function GetPlayerArmorSlot(ByVal index As Long) As Long


    GetPlayerArmorSlot = player(index).Char(player(index).charnum).ArmorSlot
End Function

Sub SetPlayerArmorSlot(ByVal index As Long, InvNum As Long)


    player(index).Char(player(index).charnum).ArmorSlot = InvNum
End Sub

Function GetPlayerWeaponSlot(ByVal index As Long) As Long


    GetPlayerWeaponSlot = player(index).Char(player(index).charnum).WeaponSlot
End Function

Sub SetPlayerWeaponSlot(ByVal index As Long, InvNum As Long)


    player(index).Char(player(index).charnum).WeaponSlot = InvNum
End Sub

Function GetPlayerHelmetSlot(ByVal index As Long) As Long


    GetPlayerHelmetSlot = player(index).Char(player(index).charnum).HelmetSlot
End Function

Sub SetPlayerHelmetSlot(ByVal index As Long, InvNum As Long)


    player(index).Char(player(index).charnum).HelmetSlot = InvNum
End Sub

Function GetPlayerShieldSlot(ByVal index As Long) As Long


    GetPlayerShieldSlot = player(index).Char(player(index).charnum).ShieldSlot
End Function

Sub SetPlayerShieldSlot(ByVal index As Long, InvNum As Long)


    player(index).Char(player(index).charnum).ShieldSlot = InvNum
End Sub
Function GetPlayerGlobal(ByVal index As Long) As Long


GetPlayerGlobal = player(index).GlobalPriv
GetPlayerGlobal = player(index).GlobalPriv
End Function

Sub PkData(ByVal PKer As Integer, ByVal Vict As Integer)


    Dim Packet As String
    Packet = "PKSHIT" & SEP_CHAR & GetPlayerName(PKer) & SEP_CHAR & GetPlayerName(Vict) & SEP_CHAR & Time & SEP_CHAR & END_CHAR
    Call SendDataToAll(Packet)
End Sub
Function GetPlayerBankItemNum(ByVal index As Long, ByVal InvSlot As Long) As Long


    GetPlayerBankItemNum = player(index).Char(player(index).charnum).Bank(InvSlot).Num
End Function
Sub SetPlayerBankItemNum(ByVal index As Long, ByVal InvSlot As Long, ByVal ItemNum As Long)


    player(index).Char(player(index).charnum).Bank(InvSlot).Num = ItemNum
End Sub
Function GetPlayerBankItemValue(ByVal index As Long, ByVal InvSlot As Long) As Long


    GetPlayerBankItemValue = player(index).Char(player(index).charnum).Bank(InvSlot).Value
End Function
Sub SetPlayerBankItemValue(ByVal index As Long, ByVal InvSlot As Long, ByVal ItemValue As Long)


    player(index).Char(player(index).charnum).Bank(InvSlot).Value = ItemValue
End Sub
Function GetPlayerBankItemDur(ByVal index As Long, ByVal InvSlot As Long) As Long


    GetPlayerBankItemDur = player(index).Char(player(index).charnum).Bank(InvSlot).Dur
End Function
Sub SetPlayerBankItemDur(ByVal index As Long, ByVal InvSlot As Long, ByVal ItemDur As Long)


    player(index).Char(player(index).charnum).Bank(InvSlot).Dur = ItemDur
End Sub
Public Sub GetRank(ByVal PlayerZ As Long)


Dim classx As Integer
Dim lvlx As Integer
classx = GetPlayerClass(PlayerZ)
lvlx = GetPlayerLevel(PlayerZ)
If lvlx = 1 Then RankName = "Noob"
If lvlx > 1 Then RankName = "Trainee"
If lvlx > 4 Then RankName = "Pk Bait"
If lvlx > 6 Then RankName = "Experianced"

If classx = 0 Then
If lvlx > 9 Then RankName = "Apprentice"
If lvlx > 15 Then RankName = "Adept"
If lvlx > 20 Then RankName = "Archadept"
If lvlx > 25 Then RankName = "Cantropy"
If lvlx > 30 Then RankName = "Seer"
If lvlx > 35 Then RankName = "Savant"
If lvlx > 40 Then RankName = "Sage"
If lvlx > 45 Then RankName = "Incanter"
If lvlx > 50 Then RankName = "Augur"
If lvlx > 55 Then RankName = "Invoker"
If lvlx > 60 Then RankName = "Advocate"
If lvlx > 70 Then RankName = "Occultist"
If lvlx > 80 Then RankName = "Neymina"
If lvlx > 90 Then RankName = "Occultist"
If lvlx > 100 Then RankName = "Conjurer"
If lvlx > 110 Then RankName = "Neophyte"
If lvlx > 120 Then RankName = "Necrolyte"
If lvlx > 130 Then RankName = "Heretic"
If lvlx > 140 Then RankName = "Maelstrom"
If lvlx > 160 Then RankName = "Rilphia's Ancient"
End If

If classx = 1 Then
If lvlx > 9 Then RankName = "Dislyan"
If lvlx > 15 Then RankName = "Acolyte"
If lvlx > 20 Then RankName = "Sigilist"
If lvlx > 25 Then RankName = "Hierophant"
If lvlx > 30 Then RankName = "Vigilist"
If lvlx > 35 Then RankName = "Guardian"
If lvlx > 40 Then RankName = "Pacafist"
If lvlx > 45 Then RankName = "Shaman"
If lvlx > 50 Then RankName = "Reverent"
If lvlx > 55 Then RankName = "Sunderer"
If lvlx > 60 Then RankName = "Solace"
If lvlx > 70 Then RankName = "Temptest"
If lvlx > 80 Then RankName = "Meysia"
If lvlx > 90 Then RankName = "Occultist"
If lvlx > 100 Then RankName = "Clairvoyant"
If lvlx > 110 Then RankName = "Oracle"
If lvlx > 120 Then RankName = "Luminary"
If lvlx > 130 Then RankName = "Spirit Weaver"
If lvlx > 140 Then RankName = "Enlightened"
If lvlx > 160 Then RankName = "Rilphia's Medic"
End If

If classx = 2 Then
If lvlx > 9 Then RankName = "Cadet"
If lvlx > 15 Then RankName = "Private"
If lvlx > 20 Then RankName = "Corporal"
If lvlx > 25 Then RankName = "Sergeant"
If lvlx > 30 Then RankName = "Staff Sergeant"
If lvlx > 35 Then RankName = "Master Sergeant"
If lvlx > 40 Then RankName = "Sergeant Major"
If lvlx > 45 Then RankName = "Second Lieutenant"
If lvlx > 50 Then RankName = "First Lietenant"
If lvlx > 55 Then RankName = "Captain"
If lvlx > 60 Then RankName = "Major"
If lvlx > 70 Then RankName = "Lieutenant Colonel"
If lvlx > 80 Then RankName = "Colonel"
If lvlx > 90 Then RankName = "Brigadier General"
If lvlx > 100 Then RankName = "Lieutenant General"
If lvlx > 110 Then RankName = "Major General"
If lvlx > 120 Then RankName = "General"
If lvlx > 130 Then RankName = "Field Marshal"
If lvlx > 140 Then RankName = "Wargod"
If lvlx > 160 Then RankName = "Rilphia's Marine"
End If
If lvlx > 190 Then RankName = "Demi God"
End Sub

