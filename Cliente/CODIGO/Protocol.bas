Attribute VB_Name = "Protocol"
Option Explicit

' TODO : /BANIP y /UNBANIP ya no trabajan con nicks. Esto lo puede mentir en forma local el cliente con un paquete a NickToIp

'When we have a list of strings, we use this to separate them and prevent
'having too many string lengths in the queue. Yes, each string is NULL-terminated :P
Private Const SEPARATOR As String * 1 = vbNullChar

Private Type tFont
    red As Byte
    green As Byte
    blue As Byte
    bold As Boolean
    italic As Boolean
End Type

Private Enum ServerPacketID
    Logged                  ' LOGGED
    RemoveDialogs           ' QTDL
    RemoveCharDialog        ' QDL
    NavigateToggle          ' NAVEG
    EquitandoToggle
    Disconnect              ' FINOK
    CommerceEnd             ' FINCOMOK
    BankEnd                 ' FINBANOK
    CommerceInit            ' INITCOM
    BankInit                ' INITBANCO
    UserCommerceInit        ' INITCOMUSU
    UserCommerceEnd         ' FINCOMUSUOK
    ShowBlacksmithForm      ' SFH
    ShowCarpenterForm       ' SFC
    NPCSwing                ' N1
    NPCKillUser             ' 6
    BlockedWithShieldUser   ' 7
    BlockedWithShieldOther  ' 8
    UserSwing               ' U1
    UpdateNeeded            ' REAU
    ResuscitationSafeOn
    ResuscitationSafeOff
    NobilityLost            ' PN
    CantUseWhileMeditating  ' M!
    UpdateSta               ' ASS
    UpdateMana              ' ASM
    UpdateHP                ' ASH
    UpdateGold              ' ASG
    UpdateExp               ' ASE
    ChangeMap               ' CM
    PosUpdate               ' PU
    NPCHitUser              ' N2
    UserHitNPC              ' U2
    UserAttackedSwing       ' U3
    UserHittedByUser        ' N4
    UserHittedUser          ' N5
    ChatOverHead            ' ||
    ChatOverHeadHit
    ConsoleMsg              ' || - Beware!! its the same as above, but it was properly splitted
    GuildChat               ' |+
    ShowMessageBox          ' !!
    UserIndexInServer       ' IU
    UserCharIndexInServer   ' IP
    CharacterCreate         ' CC
    CharacterRemove         ' BP
    CharacterChangeNick
    CharacterMove           ' MP, +, * and _ '
    ForceCharMove
    CharacterChange         ' CP
    ObjectCreate            ' HO
    ObjectDelete            ' BO
    BlockPosition           ' BQ
    PlayMIDI                ' TM
    PlayWave                ' TW
    guildList               ' GL
    AreaChanged             ' CA
    PauseToggle             ' BKW
    CreateFX                ' CFX
    particleCreate
    UpdateUserStats         ' EST
    WorkRequestTarget       ' T01
    ChangeInventorySlot     ' CSI
    ChangeBankSlot          ' SBO
    ChangeSpellSlot         ' SHS
    Atributes               ' ATR
    BlacksmithWeapons       ' LAH
    BlacksmithArmors        ' LAR
    CarpenterObjects        ' OBR
    RestOK                  ' DOK
    ErrorMsg                ' ERR
    Blind                   ' CEGU
    Dumb                    ' DUMB
    ShowSignal              ' MCAR
    ChangeNPCInventorySlot  ' NPCI
    UpdateHungerAndThirst   ' EHYS
    Fame                    ' FAMA
    MiniStats               ' MEST
    LevelUp                 ' SUNI
    AddForumMsg             ' FMSG
    ShowForumForm           ' MFOR
    SetInvisible            ' NOVER
    MeditateToggle          ' MEDOK
    BlindNoMore             ' NSEGUE
    DumbNoMore              ' NESTUP
    SendSkills              ' SKILLS
    TrainerCreatureList     ' LSTCRI
    guildNews               ' GUILDNE
    OfferDetails            ' PEACEDE & ALLIEDE
    AlianceProposalsList    ' ALLIEPR
    PeaceProposalsList      ' PEACEPR
    CharacterInfo           ' CHRINFO
    GuildLeaderInfo         ' LEADERI
    GuildDetails            ' CLANDET
    ShowGuildFundationForm  ' SHOWFUN
    ParalizeOK              ' PARADOK
    ShowUserRequest         ' PETICIO
    TradeOK                 ' TRANSOK
    BankOK                  ' BANCOOK
    ChangeUserTradeSlot     ' COMUSUINV
    Pong
    
    UpdateTagAndStatus
    
    'GM messages
    SpawnList               ' SPL
    ShowSOSForm             ' MSOS
    ShowMOTDEditionForm     ' ZMOTD
    ShowGMPanelForm         ' ABPANEL
    UserNameList            ' LISTUSU
    
    UpdateArmor
    UpdateEscu
    UpdateCasco
    UpdateHit
    
    UpdateStrenghtAndDexterity
    UpdateStrenght
    UpdateDexterity
    CuentaSV
    
    QuestDetails
    QuestListSend
    
    '**********Lorwik/Noche************
    noche
    '**********************************
    CanjesInit
    TimeInvi
    TimeParal
    ShowPartyForm
End Enum

Private Enum ClientPacketID
    LoginExistingChar       'OLOGIN
    LoginNewChar            'NLOGIN
    Talk                    ';
    Yell                    '-
    Whisper                 '\
    Walk                    'M
    RequestPositionUpdate   'RPU
    Attack                  'AT
    PickUp                  'AG
    CombatModeToggle        'TAB        - SHOULD BE HANLDED JUST BY THE CLIENT!!
    ResuscitationSafeToggle
    RequestGuildLeaderInfo  'GLINFO
    RequestAtributes        'ATR
    RequestFame             'FAMA
    RequestSkills           'ESKI
    RequestMiniStats        'FEST
    CommerceEnd             'FINCOM
    UserCommerceEnd         'FINCOMUSU
    BankEnd                 'FINBAN
    UserCommerceOk          'COMUSUOK
    UserCommerceReject      'COMUSUNO
    Drop                    'TI
    CastSpell               'LH
    LeftClick               'LC
    DoubleClick             'RC
    Work                    'UK
    UseSpellMacro           'UMH
    UseItem                 'USA
    CraftBlacksmith         'CNS
    CraftCarpenter          'CNC
    WorkLeftClick           'WLC
    CreateNewGuild          'CIG
    SpellInfo               'INFS
    EquipItem               'EQUI
    ChangeHeading           'CHEA
    ModifySkills            'SKSE
    Train                   'ENTR
    CommerceBuy             'COMP
    BankExtractItem         'RETI
    CommerceSell            'VEND
    BankDeposit             'DEPO
    ForumPost               'DEMSG
    MoveSpell               'DESPHE
    MoveBank
    ClanCodexUpdate         'DESCOD
    UserCommerceOffer       'OFRECER
    GuildAcceptPeace        'ACEPPEAT
    GuildRejectAlliance     'RECPALIA
    GuildRejectPeace        'RECPPEAT
    GuildAcceptAlliance     'ACEPALIA
    GuildOfferPeace         'PEACEOFF
    GuildOfferAlliance      'ALLIEOFF
    GuildAllianceDetails    'ALLIEDET
    GuildPeaceDetails       'PEACEDET
    GuildRequestJoinerInfo  'ENVCOMEN
    GuildAlliancePropList   'ENVALPRO
    GuildPeacePropList      'ENVPROPP
    GuildDeclareWar         'DECGUERR
    GuildNewWebsite         'NEWWEBSI
    GuildAcceptNewMember    'ACEPTARI
    GuildRejectNewMember    'RECHAZAR
    GuildKickMember         'ECHARCLA
    GuildUpdateNews         'ACTGNEWS
    GuildMemberInfo         '1HRINFO<
    GuildOpenElections      'ABREELEC
    GuildRequestMembership  'SOLICITUD
    GuildRequestDetails     'CLANDETAILS
    Online                  '/ONLINE
    Quit                    '/SALIR
    GuildLeave              '/SALIRCLAN
    RequestAccountState     '/BALANCE
    PetStand                '/QUIETO
    PetFollow               '/ACOMPAÑAR
    TrainList               '/ENTRENAR
    Rest                    '/DESCANSAR
    Meditate                '/MEDITAR
    Resucitate              '/RESUCITAR
    Heal                    '/CURAR
    CommerceStart           '/COMERCIAR
    BankStart               '/BOVEDA
    Enlist                  '/ENLISTAR
    Information             '/INFORMACION
    Reward                  '/RECOMPENSA
    RequestMOTD             '/MOTD
    PartyLeave              '/SALIRGRUPO
    PartyCreate             '/CREARGRUPO
    PartyJoin               '/GRUPO
    GuildMessage            '/CMSG
    PartyMessage            '/PMSG
    CentinelReport          '/CENTINELA
    GuildOnline             '/ONLINECLAN
    PartyOnline             '/ONLINEGRUPO
    CouncilMessage          '/BMSG
    RoleMasterRequest       '/ROL
    GMRequest               '/GM
    ChangeDescription       '/DESC
    GuildVote               '/VOTO
    Punishments             '/PENAS
    Gamble                  '/APOSTAR
    LeaveFaction            '/RETIRAR ( with no arguments )
    BankExtractGold         '/RETIRAR ( with arguments )
    BankDepositGold         '/DEPOSITAR
    Denounce                '/DENUNCIAR
    GuildFundate            '/FUNDARCLAN
    PartyKick               '/ECHARGRUPO
    PartySetLeader          '/GRUPOLIDER
    PartyAcceptMember       '/ACCEPTGRUPO
    Ping                    '/PING
    Rank                    '/RANK
    TopQuince               '/TOP5
    GlobalOn                '/GLOBALON
    GlobalOff               '/GLOBALOFF
    MsgGlobal               ';
    
    'GM messages
    GMMessage               '/GMSG
    showName                '/SHOWNAME
    OnlineRoyalArmy         '/ONLINEREAL
    OnlineChaosLegion       '/ONLINECAOS
    GoNearby                '/IRCERCA
    Comment                 '/REM
    serverTime              '/HORA
    Where                   '/DONDE
    CreaturesInMap          '/NENE
    WarpMeToTarget          '/TELEPLOC
    WarpChar                '/TELEP
    Silence                 '/SILENCIAR
    SOSShowList             '/SHOW SOS
    SOSRemove               'SOSDONE
    GoToChar                '/IRA
    invisible               '/INVISIBLE
    GMPanel                 '/PANELGM
    RequestUserList         'LISTUSU
    Working                 '/TRABAJANDO
    Hiding                  '/OCULTANDO
    Jail                    '/CARCEL
    KillNPC                 '/RMATA
    WarnUser                '/ADVERTENCIA
    EditChar                '/MOD
    RequestCharInfo         '/INFO
    RequestCharStats        '/STAT
    RequestCharGold         '/BAL
    RequestCharInventory    '/INV
    RequestCharBank         '/BOV
    RequestCharSkills       '/SKILLS
    ReviveChar              '/REVIVIR
    OnlineGM                '/ONLINEGM
    OnlineMap               '/ONLINEMAP
    Forgive                 '/PERDON
    Kick                    '/ECHAR
    Execute                 '/EJECUTAR
    BanChar                 '/BAN
    UnbanChar               '/UNBAN
    BanACC                  '/BANACC
    UnbanACC                '/UNBANACC
    NPCFollow               '/SEGUIR
    SummonChar              '/SUM
    SpawnListRequest        '/CC
    SpawnCreature           'SPA
    ResetNPCInventory       '/RESETINV
    CleanWorld              '/LIMPIAR
    ServerMessage           '/RMSG
    NickToIP                '/NICK2IP
    IPToNick                '/IP2NICK
    GuildOnlineMembers      '/ONCLAN
    TeleportCreate          '/CT
    TeleportDestroy         '/DT
    SetCharDescription      '/SETDESC
    ForceMIDIToMap          '/FORCEMIDIMAP
    ForceWAVEToMap          '/FORCEWAVMAP
    RoyalArmyMessage        '/REALMSG
    ChaosLegionMessage      '/CAOSMSG
    CitizenMessage          '/CIUMSG
    CriminalMessage         '/CRIMSG
    TalkAsNPC               '/TALKAS
    DestroyAllItemsInArea   '/MASSDEST
    AcceptRoyalCouncilMember '/ACEPTCONSE
    AcceptChaosCouncilMember '/ACEPTCONSECAOS
    ItemsInTheFloor         '/PISO
    MakeDumb                '/ESTUPIDO
    MakeDumbNoMore          '/NOESTUPIDO
    DumpIPTables            '/DUMPSECURITY
    CouncilKick             '/KICKCONSE
    SetTrigger              '/TRIGGER
    AskTrigger              '/TRIGGER with no arguments
    BannedIPList            '/BANIPLIST
    BannedIPReload          '/BANIPRELOAD
    GuildMemberList         '/MIEMBROSCLAN
    GuildBan                '/BANCLAN
    BanIP                   '/BANIP
    UnbanIP                 '/UNBANIP
    CreateItem              '/CI
    DestroyItems            '/DEST
    ChaosLegionKick         '/NOCAOS
    RoyalArmyKick           '/NOREAL
    ForceMIDIAll            '/FORCEMIDI
    ForceWAVEAll            '/FORCEWAV
    RemovePunishment        '/BORRARPENA
    TileBlockedToggle       '/BLOQ
    KillNPCNoRespawn        '/MATA
    KillAllNearbyNPCs       '/MASSKILL
    LastIP                  '/LASTIP
    ChangeMOTD              '/MOTDCAMBIA
    SetMOTD                 'ZMOTD
    SystemMessage           '/SMSG
    CreateNPC               '/ACC
    CreateNPCWithRespawn    '/RACC
    ImperialArmour          '/AI1 - 4
    ChaosArmour             '/AC1 - 4
    NavigateToggle          '/NAVE
    ServerOpenToUsersToggle '/HABILITAR
    TurnOffServer           '/APAGAR
    TurnCriminal            '/CONDEN
    ResetFactions           '/RAJAR
    RemoveCharFromGuild     '/RAJARCLAN
    RequestCharMail         '/LASTEMAIL
    AlterPassword           '/APASS
    AlterMail               '/AEMAIL
    AlterName               '/ANAME
    ToggleCentinelActivated '/CENTINELAACTIVADO
    DoBackUp                '/DOBACKUP
    ShowGuildMessages       '/SHOWCMSG
    SaveMap                 '/GUARDAMAPA
    ChangeMapInfoPK         '/MODMAPINFO PK
    ChangeMapInfoBackup     '/MODMAPINFO BACKUP
    ChangeMapInfoRestricted '/MODMAPINFO RESTRINGIR
    ChangeMapInfoNoMagic    '/MODMAPINFO MAGIASINEFECTO
    ChangeMapInfoNoInvi     '/MODMAPINFO INVISINEFECTO
    ChangeMapInfoNoResu     '/MODMAPINFO RESUSINEFECTO
    ChangeMapInfoLand       '/MODMAPINFO TERRENO
    ChangeMapInfoZone       '/MODMAPINFO ZONA
    SaveChars               '/GRABAR
    CleanSOS                '/BORRAR SOS
    ShowServerForm          '/SHOW INT
    KickAllChars            '/ECHARTODOSPJS
    ReloadNPCs              '/RELOADNPCS
    ReloadServerIni         '/RELOADSINI
    ReloadSpells            '/RELOADHECHIZOS
    ReloadObjects           '/RELOADOBJ
    Restart                 '/REINICIAR
    ChatColor               '/CHATCOLOR
    Ignored                 '/IGNORADO
    CheckSlot               '/SLOT
    SearchObj               '/BUSCAR [OBJETO]
    SetIniVar               '/SETINIVAR LLAVE CLAVE VALOR
    CuentaCL
    WinterClientPacket
End Enum

Private Enum WinterClientPacketID
    Quest                   '/QUEST
    QuestAccept
    QuestListRequest
    QuestDetailsRequest
    QuestAbandon
    
    Anochecer
    SwapObjects
    canjes
    Puntos
    DueloAutomatico
    CancelarDueloAutomatico
    ParticiparDueloAutomatico
    RequestPartyForm
    Invocar
    DesDenuncia
    HacerTorneo
    Torneo
    Cerrartorneo
    Transferencia
End Enum
    
Public Enum FontTypeNames
    FONTTYPE_TALK
    FONTTYPE_FIGHT
    FONTTYPE_WARNING
    FONTTYPE_INFO
    FONTTYPE_INFOBOLD
    FONTTYPE_EJECUCION
    FONTTYPE_PARTY
    FONTTYPE_VENENO
    FONTTYPE_GUILD
    FONTTYPE_SERVER
    FONTTYPE_GUILDMSG
    FONTTYPE_CONSEJO
    FONTTYPE_CONSEJOCAOS
    FONTTYPE_CONSEJOVesA
    FONTTYPE_CONSEJOCAOSVesA
    FONTTYPE_CENTINELA
    FONTTYPE_GMMSG
    FONTTYPE_GM
    FONTTYPE_CITIZEN
    FONTTYPE_CONSE
    FONTTYPE_DIOS
    FONTTYPE_GLOBAL
    FONTTYPE_GLD
    FONTTYPE_EXP
End Enum

Public FontTypes(23) As tFont

''
' Initializes the fonts array

Public Sub InitFonts()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With FontTypes(FontTypeNames.FONTTYPE_TALK)
        .red = 255
        .green = 255
        .blue = 255
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_FIGHT)
        .red = 255
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_WARNING)
        .red = 233
        .green = 22
        .blue = 28
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_INFO)
        .red = 65
        .green = 190
        .blue = 156
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_INFOBOLD)
        .red = 65
        .green = 190
        .blue = 156
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_EJECUCION)
        .red = 130
        .green = 130
        .blue = 130
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_PARTY)
        .red = 255
        .green = 180
        .blue = 250
    End With
    
    FontTypes(FontTypeNames.FONTTYPE_VENENO).green = 255
    
    With FontTypes(FontTypeNames.FONTTYPE_GUILD)
        .red = 255
        .green = 255
        .blue = 255
        .bold = 1
    End With
    
    FontTypes(FontTypeNames.FONTTYPE_SERVER).green = 185
    
    With FontTypes(FontTypeNames.FONTTYPE_GUILDMSG)
        .red = 228
        .green = 199
        .blue = 27
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CONSEJO)
        .red = 130
        .green = 130
        .blue = 255
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CONSEJOCAOS)
        .red = 255
        .green = 60
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CONSEJOVesA)
        .green = 200
        .blue = 255
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CONSEJOCAOSVesA)
        .red = 255
        .green = 50
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CENTINELA)
        .green = 255
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_GMMSG)
        .red = 255
        .green = 255
        .blue = 255
        .italic = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_GM)
        .red = 30
        .green = 255
        .blue = 30
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CITIZEN)
        .blue = 200
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CONSE)
        .red = 30
        .green = 150
        .blue = 30
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_DIOS)
        .red = 250
        .green = 250
        .blue = 150
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_GLOBAL)
        .red = 139
        .green = 248
        .blue = 244
        .italic = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_GLD)
        .red = 204
        .green = 193
        .blue = 115
        .italic = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_EXP)
        .red = 53
        .green = 183
        .blue = 247
        .bold = 1
    End With
End Sub

''
' Handles incoming data.

Public Sub HandleIncomingData()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
On Error Resume Next
    Select Case incomingData.PeekByte()
        Case ServerPacketID.Logged                  ' LOGGED
            Call HandleLogged
        
        Case ServerPacketID.RemoveDialogs           ' QTDL
            Call HandleRemoveDialogs
        
        Case ServerPacketID.RemoveCharDialog        ' QDL
            Call HandleRemoveCharDialog
        
        Case ServerPacketID.NavigateToggle          ' NAVEG
            Call HandleNavigateToggle
            
        Case ServerPacketID.EquitandoToggle
            Call HandleEquitandoToggle
        
        Case ServerPacketID.Disconnect              ' FINOK
            Call HandleDisconnect
        
        Case ServerPacketID.CommerceEnd             ' FINCOMOK
            Call HandleCommerceEnd
        
        Case ServerPacketID.BankEnd                 ' FINBANOK
            Call HandleBankEnd
        
        Case ServerPacketID.CommerceInit            ' INITCOM
            Call HandleCommerceInit
        
        Case ServerPacketID.BankInit                ' INITBANCO
            Call HandleBankInit
        
        Case ServerPacketID.UserCommerceInit        ' INITCOMUSU
            Call HandleUserCommerceInit
        
        Case ServerPacketID.UserCommerceEnd         ' FINCOMUSUOK
            Call HandleUserCommerceEnd
        
        Case ServerPacketID.ShowBlacksmithForm      ' SFH
            Call HandleShowBlacksmithForm
        
        Case ServerPacketID.ShowCarpenterForm       ' SFC
            Call HandleShowCarpenterForm
        
        Case ServerPacketID.NPCSwing                ' N1
            Call HandleNPCSwing
        
        Case ServerPacketID.NPCKillUser             ' 6
            Call HandleNPCKillUser
        
        Case ServerPacketID.BlockedWithShieldUser   ' 7
            Call HandleBlockedWithShieldUser
        
        Case ServerPacketID.BlockedWithShieldOther  ' 8
            Call HandleBlockedWithShieldOther
        
        Case ServerPacketID.UserSwing               ' U1
            Call HandleUserSwing
        
        Case ServerPacketID.UpdateNeeded            ' REAU
            Call HandleUpdateNeeded
            
        Case ServerPacketID.ResuscitationSafeOff
            Call HandleResuscitationSafeOff
        
        Case ServerPacketID.ResuscitationSafeOn
            Call HandleResuscitationSafeOn
        
        Case ServerPacketID.NobilityLost            ' PN
            Call HandleNobilityLost
        
        Case ServerPacketID.CantUseWhileMeditating  ' M!
            Call HandleCantUseWhileMeditating
        
        Case ServerPacketID.UpdateSta               ' ASS
            Call HandleUpdateSta
        
        Case ServerPacketID.UpdateMana              ' ASM
            Call HandleUpdateMana
        
        Case ServerPacketID.UpdateHP                ' ASH
            Call HandleUpdateHP
        
        Case ServerPacketID.UpdateGold              ' ASG
            Call HandleUpdateGold
        
        Case ServerPacketID.UpdateExp               ' ASE
            Call HandleUpdateExp
        
        Case ServerPacketID.ChangeMap               ' CM
            Call HandleChangeMap
        
        Case ServerPacketID.PosUpdate               ' PU
            Call HandlePosUpdate
        
        Case ServerPacketID.NPCHitUser              ' N2
            Call HandleNPCHitUser
        
        Case ServerPacketID.UserHitNPC              ' U2
            Call HandleUserHitNPC
        
        Case ServerPacketID.UserAttackedSwing       ' U3
            Call HandleUserAttackedSwing
        
        Case ServerPacketID.UserHittedByUser        ' N4
            Call HandleUserHittedByUser
        
        Case ServerPacketID.UserHittedUser          ' N5
            Call HandleUserHittedUser
        
        Case ServerPacketID.ChatOverHead            ' ||
            Call HandleChatOverHead
            
        Case ServerPacketID.ChatOverHeadHit
            Call HandleChatOverHeadhit
        
        Case ServerPacketID.ConsoleMsg              ' || - Beware!! its the same as above, but it was properly splitted
            Call HandleConsoleMessage
        
        Case ServerPacketID.GuildChat               ' |+
            Call HandleGuildChat
        
        Case ServerPacketID.ShowMessageBox          ' !!
            Call HandleShowMessageBox
        
        Case ServerPacketID.UserIndexInServer       ' IU
            Call HandleUserIndexInServer
        
        Case ServerPacketID.UserCharIndexInServer   ' IP
            Call HandleUserCharIndexInServer
        
        Case ServerPacketID.CharacterCreate         ' CC
            Call HandleCharacterCreate
        
        Case ServerPacketID.CharacterRemove         ' BP
            Call HandleCharacterRemove
        
        Case ServerPacketID.CharacterChangeNick
            Call HandleCharacterChangeNick
            
        Case ServerPacketID.CharacterMove           ' MP, +, * and _ '
            Call HandleCharacterMove
            
        Case ServerPacketID.ForceCharMove
            Call HandleForceCharMove
        
        Case ServerPacketID.CharacterChange         ' CP
            Call HandleCharacterChange
        
        Case ServerPacketID.ObjectCreate            ' HO
            Call HandleObjectCreate
        
        Case ServerPacketID.ObjectDelete            ' BO
            Call HandleObjectDelete
        
        Case ServerPacketID.BlockPosition           ' BQ
            Call HandleBlockPosition
        
        Case ServerPacketID.PlayMIDI                ' TM
            Call HandlePlayMIDI
        
        Case ServerPacketID.PlayWave                ' TW
            Call HandlePlayWave
        
        Case ServerPacketID.guildList               ' GL
            Call HandleGuildList
        
        Case ServerPacketID.AreaChanged             ' CA
            Call HandleAreaChanged
        
        Case ServerPacketID.PauseToggle             ' BKW
            Call HandlePauseToggle
        
        Case ServerPacketID.CreateFX                ' CFX
            Call HandleCreateFX
        
        Case ServerPacketID.particleCreate
            Call HandleParticleCreate
        
        Case ServerPacketID.UpdateUserStats         ' EST
            Call HandleUpdateUserStats
        
        Case ServerPacketID.WorkRequestTarget       ' T01
            Call HandleWorkRequestTarget
        
        Case ServerPacketID.ChangeInventorySlot     ' CSI
            Call HandleChangeInventorySlot
        
        Case ServerPacketID.ChangeBankSlot          ' SBO
            Call HandleChangeBankSlot
        
        Case ServerPacketID.ChangeSpellSlot         ' SHS
            Call HandleChangeSpellSlot
        
        Case ServerPacketID.Atributes               ' ATR
            Call HandleAtributes
        
        Case ServerPacketID.BlacksmithWeapons       ' LAH
            Call HandleBlacksmithWeapons
        
        Case ServerPacketID.BlacksmithArmors        ' LAR
            Call HandleBlacksmithArmors
        
        Case ServerPacketID.CarpenterObjects        ' OBR
            Call HandleCarpenterObjects
        
        Case ServerPacketID.RestOK                  ' DOK
            Call HandleRestOK
        
        Case ServerPacketID.ErrorMsg                ' ERR
            Call HandleErrorMessage
        
        Case ServerPacketID.Blind                   ' CEGU
            Call HandleBlind
        
        Case ServerPacketID.Dumb                    ' DUMB
            Call HandleDumb
        
        Case ServerPacketID.ShowSignal              ' MCAR
            Call HandleShowSignal
        
        Case ServerPacketID.ChangeNPCInventorySlot  ' NPCI
            Call HandleChangeNPCInventorySlot
        
        Case ServerPacketID.UpdateHungerAndThirst   ' EHYS
            Call HandleUpdateHungerAndThirst
        
        Case ServerPacketID.Fame                    ' FAMA
            Call HandleFame
        
        Case ServerPacketID.MiniStats               ' MEST
            Call HandleMiniStats
        
        Case ServerPacketID.LevelUp                 ' SUNI
            Call HandleLevelUp
        
        Case ServerPacketID.AddForumMsg             ' FMSG
            Call HandleAddForumMessage
        
        Case ServerPacketID.ShowForumForm           ' MFOR
            Call HandleShowForumForm
        
        Case ServerPacketID.SetInvisible            ' NOVER
            Call HandleSetInvisible
        
        Case ServerPacketID.MeditateToggle          ' MEDOK
            Call HandleMeditateToggle
        
        Case ServerPacketID.BlindNoMore             ' NSEGUE
            Call HandleBlindNoMore
        
        Case ServerPacketID.DumbNoMore              ' NESTUP
            Call HandleDumbNoMore
        
        Case ServerPacketID.SendSkills              ' SKILLS
            Call HandleSendSkills
        
        Case ServerPacketID.TrainerCreatureList     ' LSTCRI
            Call HandleTrainerCreatureList
        
        Case ServerPacketID.guildNews               ' GUILDNE
            Call HandleGuildNews
        
        Case ServerPacketID.OfferDetails            ' PEACEDE and ALLIEDE
            Call HandleOfferDetails
        
        Case ServerPacketID.AlianceProposalsList    ' ALLIEPR
            Call HandleAlianceProposalsList
        
        Case ServerPacketID.PeaceProposalsList      ' PEACEPR
            Call HandlePeaceProposalsList
        
        Case ServerPacketID.CharacterInfo           ' CHRINFO
            Call HandleCharacterInfo
        
        Case ServerPacketID.GuildLeaderInfo         ' LEADERI
            Call HandleGuildLeaderInfo
        
        Case ServerPacketID.GuildDetails            ' CLANDET
            Call HandleGuildDetails
        
        Case ServerPacketID.ShowGuildFundationForm  ' SHOWFUN
            Call HandleShowGuildFundationForm
        
        Case ServerPacketID.ParalizeOK              ' PARADOK
            Call HandleParalizeOK
        
        Case ServerPacketID.ShowUserRequest         ' PETICIO
            Call HandleShowUserRequest
        
        Case ServerPacketID.TradeOK                 ' TRANSOK
            Call HandleTradeOK
        
        Case ServerPacketID.BankOK                  ' BANCOOK
            Call HandleBankOK
        
        Case ServerPacketID.ChangeUserTradeSlot     ' COMUSUINV
            Call HandleChangeUserTradeSlot
        
        Case ServerPacketID.Pong
            Call HandlePong
        
        Case ServerPacketID.UpdateTagAndStatus
            Call HandleUpdateTagAndStatus

        
        '*******************
        'GM messages
        '*******************
        Case ServerPacketID.SpawnList               ' SPL
            Call HandleSpawnList
        
        Case ServerPacketID.ShowSOSForm             ' RSOS and MSOS
            Call HandleShowSOSForm
        
        Case ServerPacketID.ShowMOTDEditionForm     ' ZMOTD
            Call HandleShowMOTDEditionForm
        
        Case ServerPacketID.ShowGMPanelForm         ' ABPANEL
            Call HandleShowGMPanelForm
        
        Case ServerPacketID.UserNameList            ' LISTUSU
            Call HandleUserNameList
            
        Case ServerPacketID.UpdateArmor
            Call HandleUpdateArmor
        
        Case ServerPacketID.UpdateEscu
            Call HandleUpdateEscu
            
        Case ServerPacketID.UpdateCasco
            Call HandleUpdateCasco
            
        Case ServerPacketID.UpdateHit
            Call HandleUpdateHit
            
        Case ServerPacketID.UpdateStrenghtAndDexterity
            Call HandleUpdateStrenghtAndDexterity
            
        Case ServerPacketID.UpdateStrenght
            Call HandleUpdateStrenght
            
        Case ServerPacketID.UpdateDexterity
            Call HandleUpdateDexterity
            
        Case ServerPacketID.CuentaSV
            Call HandleCuentas
            
        Case ServerPacketID.QuestDetails
            Call HandleQuestDetails
        
        Case ServerPacketID.QuestListSend
            Call HandleQuestListSend
                      
        '*********************Lorwik/Noche*************
        Case ServerPacketID.noche
            Call handleNoche
        '***********************************************
        Case ServerPacketID.CanjesInit
            Call HandleCanjesInit
            
        Case ServerPacketID.TimeInvi
            Call HandleTimeInvi
            
        Case ServerPacketID.TimeParal
            Call HandleTimeParal
            
        Case ServerPacketID.ShowPartyForm
            Call HandleShowPartyForm
            
        Case Else
            'ERROR : Abort!
            Exit Sub
    End Select
    
    'Done with this packet, move on to next one
    If incomingData.Length > 0 And Err.number <> incomingData.NotEnoughDataErrCode Then
        Err.Clear
        Call HandleIncomingData
    End If
End Sub

''
' Handles the Logged message.

Private Sub HandleLogged()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    ' Variable initialization
    EngineRun = True
    Nombres = True
    
    'Set connected state
    Call SetConnected
    
    'Minimapa
    Call DibujarMiniMapa
    
    Inventario.UpdateInventory
    
End Sub

''
' Handles the RemoveDialogs message.

Private Sub HandleRemoveDialogs()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call Dialogos.RemoveAllDialogs
End Sub

''
' Handles the RemoveCharDialog message.

Private Sub HandleRemoveCharDialog()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Check if the packet is complete
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call Dialogos.RemoveDialog(incomingData.ReadInteger())
End Sub

''
' Handles the NavigateToggle message.

Private Sub HandleNavigateToggle()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserNavegando = Not UserNavegando
End Sub

''
' Handles the EquitandoToggle message.

Private Sub HandleEquitandoToggle()
'***************************************************
'Author: Lorwik
'Last Modification: 23/08/11
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserEquitando = Not UserEquitando
    
    Call SpeedCaballo
End Sub

''
' Handles the Disconnect message.

Private Sub HandleDisconnect()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    Dim i As Long
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    'Close connection
    If frmMain.Winsock1.State <> sckClosed Then _
        frmMain.Winsock1.Close
    
    Call SetMusicInfo("", "", "", "Games", , False)
    'Hide main form
    frmMain.Visible = False
    
    'Stop audio
    Call Audio.StopWave
    frmMain.IsPlaying = PlayLoop.plNone
    
    'Show connection form
    frmConnect.Visible = True
    frmConnect.Enabled = False
    General_Set_Song 2, True
    Call Relog
    
    'Reset global vars
    IScombate = False
    UserDescansar = False
    UserParalizado = False
    pausa = False
    UserCiego = False
    UserMeditar = False
    UserNavegando = False
    UserEquitando = False
    'Reseteo el Speed
    Call SpeedCaballo
    bFogata = False
    SkillPoints = 0
    
    'Delete all kind of dialogs
    Call CleanDialogs
    
    'Reset some char variables...
    For i = 1 To LastChar
        charlist(i).invisible = False
    Next i
    
    'Unload all forms except frmMain and frmConnect
    Dim frm As Form
    
    For Each frm In Forms
        If frm.Name <> frmMain.Name And frm.Name <> frmConnect.Name Then
            Unload frm
        End If
    Next
    
    For i = 1 To MAX_INVENTORY_SLOTS
        Call Inventario.SetItem(i, 0, 0, 0, 0, 0, 0, 0, 0, 0, "")
    Next i
End Sub

''
' Handles the CommerceEnd message.

Private Sub HandleCommerceEnd()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    Set InvComUsu = Nothing
    Set InvComNpc = Nothing
    
    'Hide form
    Unload frmComerciar
    
    'Reset vars
    Comerciando = False
End Sub

''
' Handles the BankEnd message.

Private Sub HandleBankEnd()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    Set InvBanco(0) = Nothing
    Set InvBanco(1) = Nothing
    
    Unload frmBancoObj
    Comerciando = False
End Sub


''
' Handles the CommerceInit message.

Private Sub HandleCommerceInit()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    Dim i As Long
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Set InvComUsu = New clsGrapchicalInventory
    Set InvComNpc = New clsGrapchicalInventory
    
    frmComerciar.Show , frmMain
    DoEvents
    
    ' Initialize commerce inventories
    Call InvComUsu.Initialize(DirectD3D8, frmComerciar.picInvUser, Inventario.MaxObjs)
    Call InvComNpc.Initialize(DirectD3D8, frmComerciar.picInvNpc, MAX_NPC_INVENTORY_SLOTS)

    'Fill user inventory
    For i = 1 To MAX_INVENTORY_SLOTS
        If Inventario.OBJIndex(i) <> 0 Then
            With Inventario
                Call InvComUsu.SetItem(i, .OBJIndex(i), _
                .Amount(i), .Equipped(i), .GrhIndex(i), _
                .OBJType(i), .MaxHit(i), .MinHit(i), .Def(i), _
                .Valor(i), .ItemName(i))
            End With
        End If
    Next i
    
    ' Fill Npc inventory
    For i = 1 To 50
        If NPCInventory(i).OBJIndex <> 0 Then
            With NPCInventory(i)
                Call InvComNpc.SetItem(i, .OBJIndex, _
                .Amount, 0, .GrhIndex, _
                .OBJType, .MaxHit, .MinHit, .Def, _
                .Valor, .Name)
            End With
        End If
    Next i
    
    'Set state and show form
    Comerciando = True
    
    'Reproducimos el saludo del comerciante
    Call General_Set_Wav("comerciante" & RandomNumber(1, 9) & ".wav")
End Sub

''
' Handles the BankInit message.

Private Sub HandleBankInit()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    Dim i As Long
    Dim BankGold As Long
    'Set state and show form
    Comerciando = True
    
    frmBancoObj.Show , frmMain
    DoEvents
    Sleep 10
    
    'Remove packet ID
    Call incomingData.ReadByte
    BankGold = incomingData.ReadLong
    frmBancoObj.lblUserGld.Caption = BankGold
    Call InvBanco(0).Initialize(DirectD3D8, frmBancoObj.PicBancoInv, MAX_BANCOINVENTORY_SLOTS)
    
    Call InvBanco(1).Initialize(DirectD3D8, frmBancoObj.PicInv, Inventario.MaxObjs)

    
    For i = 1 To Inventario.MaxObjs
        With Inventario
            Call InvBanco(1).SetItem(i, .OBJIndex(i), _
                .Amount(i), .Equipped(i), .GrhIndex(i), _
                .OBJType(i), .MaxHit(i), .MinHit(i), .Def(i), _
                .Valor(i), .ItemName(i))
        End With
    Next i
    
    For i = 1 To MAX_BANCOINVENTORY_SLOTS
        With UserBancoInventory(i)
            Call InvBanco(0).SetItem(i, .OBJIndex, _
                .Amount, .Equipped, .GrhIndex, _
                .OBJType, .MaxHit, .MinHit, .Def, _
                .Valor, .Name)
        End With
    Next i
    
    

End Sub

''
' Handles the UserCommerceInit message.

Private Sub HandleUserCommerceInit()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    Dim i As Long
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    'Clears lists if necessary
    If frmComerciarUsu.List1.ListCount > 0 Then frmComerciarUsu.List1.Clear
    If frmComerciarUsu.List2.ListCount > 0 Then frmComerciarUsu.List2.Clear
    
    'Fill inventory list
    For i = 1 To MAX_INVENTORY_SLOTS
        If Inventario.OBJIndex(i) <> 0 Then
            frmComerciarUsu.List1.AddItem Inventario.ItemName(i)
            frmComerciarUsu.List1.ItemData(frmComerciarUsu.List1.NewIndex) = Inventario.Amount(i)
        Else
            frmComerciarUsu.List1.AddItem ""
            frmComerciarUsu.List1.ItemData(frmComerciarUsu.List1.NewIndex) = 0
        End If
    Next i
    
    'Set state and show form
    Comerciando = True
    Call frmComerciarUsu.Show(vbModal, frmMain)
End Sub

''
' Handles the UserCommerceEnd message.

Private Sub HandleUserCommerceEnd()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    'Clear the lists
    frmComerciarUsu.List1.Clear
    frmComerciarUsu.List2.Clear
    
    'Destroy the form and reset the state
    Unload frmComerciarUsu
    Comerciando = False
End Sub

''
' Handles the ShowBlacksmithForm message.

Private Sub HandleShowBlacksmithForm()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte

        frmHerrero.Show , frmMain
End Sub

''
' Handles the ShowCarpenterForm message.

Private Sub HandleShowCarpenterForm()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
        frmCarp.Show , frmMain
End Sub

''
' Handles the NPCSwing message.

Private Sub HandleNPCSwing()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_CRIATURA_FALLA_GOLPE, 255, 0, 0, True, False, False)
End Sub

''
' Handles the NPCKillUser message.

Private Sub HandleNPCKillUser()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_CRIATURA_MATADO, 255, 0, 0, True, False, False)
End Sub

''
' Handles the BlockedWithShieldUser message.

Private Sub HandleBlockedWithShieldUser()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_RECHAZO_ATAQUE_ESCUDO, 255, 0, 0, True, False, False)
End Sub

''
' Handles the BlockedWithShieldOther message.

Private Sub HandleBlockedWithShieldOther()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_USUARIO_RECHAZO_ATAQUE_ESCUDO, 255, 0, 0, True, False, False)
End Sub

''
' Handles the UserSwing message.

Private Sub HandleUserSwing()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_FALLADO_GOLPE, 255, 0, 0, True, False, False)
End Sub

''
' Handles the UpdateNeeded message.

Private Sub HandleUpdateNeeded()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    'Call frmMain.DibujarSatelite
End Sub

''
' Handles the ResuscitationSafeOff message.

Private Sub HandleResuscitationSafeOff()
'***************************************************
'Author: Rapsodius
'Creation date: 10/10/07
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call frmMain.ControlSeguroResu(False)
    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_SEGURO_RESU_OFF, 255, 0, 0, True, False, False)
End Sub

''
' Handles the ResuscitationSafeOn message.

Private Sub HandleResuscitationSafeOn()
'***************************************************
'Author: Rapsodius
'Creation date: 10/10/07
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call frmMain.ControlSeguroResu(True)
    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_SEGURO_RESU_ON, 0, 255, 0, True, False, False)
End Sub

''
' Handles the NobilityLost message.

Private Sub HandleNobilityLost()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PIERDE_NOBLEZA, 255, 0, 0, False, False, False)
End Sub

''
' Handles the CantUseWhileMeditating message.

Private Sub HandleCantUseWhileMeditating()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_USAR_MEDITANDO, 255, 0, 0, False, False, False)
End Sub

''
' Handles the UpdateSta message.

Private Sub HandleUpdateSta()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Check packet is complete
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    'Get data and update form
    UserMinSTA = incomingData.ReadInteger()
    
    frmMain.lblEnergia = UserMinSTA & "/" & UserMaxSTA
    
    frmMain.STAShp.Width = (((UserMinSTA / 100) / (UserMaxSTA / 100)) * 92)
    
End Sub

''
' Handles the UpdateMana message.
Private Sub HandleUpdateMana()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Check packet is complete
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    'Get data and update form
    UserMinMAN = incomingData.ReadInteger()
    
    frmMain.lblMana = UserMinMAN & "/" & UserMaxMAN
    
    Dim bWidth As Byte
    
    If UserMaxMAN > 0 Then _
        
    frmMain.MANShp.Width = (((UserMinMAN / 100) / (UserMaxMAN / 100)) * 92)
    End If
End Sub
''
' Handles the UpdateHP message.

Private Sub HandleUpdateHP()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Check packet is complete
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    'Get data and update form
    UserMinHP = incomingData.ReadInteger()
    
    frmMain.lblVida = UserMinHP & "/" & UserMaxHP
    
    Dim bWidth As Byte
    
    
    frmMain.Hpshp.Width = (((UserMinHP / 100) / (UserMaxHP / 100)) * 92)
    
    
    'Is the user alive??
    If UserMinHP = 0 Then
        UserEstado = 1
        Call ClimaX
        UserEquitando = False
        'Reseteo el Speed
        Call SpeedCaballo
    Else
        UserEstado = 0
    End If
End Sub

''
' Handles the UpdateGold message.

Private Sub HandleUpdateGold()
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 08/14/07
'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
'- 08/14/07: Added GldLbl color variation depending on User Gold and Level
'***************************************************
    'Check packet is complete
    If incomingData.Length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    'Get data and update form
    UserGLD = incomingData.ReadLong()
    
    
    frmMain.GldLbl.Caption = UserGLD
End Sub

''
' Handles the UpdateExp message.

Private Sub HandleUpdateExp()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Check packet is complete
    If incomingData.Length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    'Get data and update form
    UserExp = incomingData.ReadLong()
    frmMain.exp.Caption = UserExp & " / " & UserPasarNivel
    frmMain.ExpShp.Width = (((UserExp / 100) / (UserPasarNivel / 100)) * 210)
    
    If Not frmMain.LvlLbl.Caption = 50 Then
        frmMain.lblPorcLvl.Caption = Round(CDbl(UserExp) * CDbl(100) / CDbl(UserPasarNivel), 2) & "%"
    Else
        frmMain.lblPorcLvl.Caption = "¡Nivel Máximo!"
    End If
End Sub
Private Sub HandleUpdateStrenghtAndDexterity()
'***************************************************
'Author: Budi
'Last Modification: 11/26/09
'***************************************************
    'Check packet is complete
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    'Get data and update form
    UserFuerza = incomingData.ReadByte
    UserAgilidad = incomingData.ReadByte
    frmMain.lblStrg.Caption = UserFuerza
    frmMain.lblDext.Caption = UserAgilidad
    
    frmMain.BarraAgilidad.Width = ((UserAgilidad / 25) * 27)
    frmMain.BarraFuerza.Width = ((UserFuerza / 25) * 27)
End Sub

' Handles the UpdateStrenghtAndDexterity message.

Private Sub HandleUpdateStrenght()
'***************************************************
'Author: Budi
'Last Modification: 11/26/09
'***************************************************
    'Check packet is complete
    If incomingData.Length < 2 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    'Get data and update form
    UserFuerza = incomingData.ReadByte
    frmMain.lblStrg.Caption = UserFuerza
    frmMain.BarraFuerza.Width = ((UserFuerza / 25) * 27)
End Sub

' Handles the UpdateStrenghtAndDexterity message.

Private Sub HandleUpdateDexterity()
'***************************************************
'Author: Budi
'Last Modification: 11/26/09
'***************************************************
    'Check packet is complete
    If incomingData.Length < 2 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    'Get data and update form
    UserAgilidad = incomingData.ReadByte
    frmMain.lblDext.Caption = UserAgilidad
    frmMain.BarraAgilidad.Width = ((UserAgilidad / 25) * 27)
End Sub
''
' Handles the ChangeMap message.
''
' Handles the ChangeMap message.

Private Sub HandleChangeMap()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 10 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    Dim file As String
    Dim tempint As Integer
    
    'Remove packet ID
    Call incomingData.ReadByte
    UserMap = incomingData.ReadInteger()
    Zona = incomingData.ReadASCIIString()
    file = Get_Extract(Map, "Mapa" & UserMap & ".map")
    
    
    If FileExist(file, vbNormal) Then
        Open file For Binary As #1
        Seek #1, 1
        Get #1, , tempint
        Close #1
        Call SwitchMap(UserMap, file)
        Delete_File file
    Else
        'no encontramos el mapa en el hdd
        MsgBox "Error en los mapas, algún archivo ha sido modificado o esta dañado."
        Call CloseClient
    End If
    
    'Dibujo el minimapa
    Call DibujarMiniMapa
    
    
    file = Get_Extract(Scripts, "WorldMapData.dat")
    'Nombre del mapa
    MapName = GetVar(file, "WorldData", UserMap)
    If Not MapName = "" Then
        frmMain.lblMapName.Caption = MapName
    Else
        frmMain.lblMapName.Caption = "Mapa Desconocido"
    End If
    
    Delete_File file
    'Sonidos Ambientales
    Call Ambient
End Sub

''
' Handles the PosUpdate message.

Private Sub HandlePosUpdate()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    'Remove char from old position
    If MapData(UserPos.X, UserPos.Y).CharIndex = UserCharIndex Then
        MapData(UserPos.X, UserPos.Y).CharIndex = 0
    End If
    
    'Set new pos
    UserPos.X = incomingData.ReadByte()
    UserPos.Y = incomingData.ReadByte()
    
    'Set char
    MapData(UserPos.X, UserPos.Y).CharIndex = UserCharIndex
    charlist(UserCharIndex).Pos = UserPos
    
    'Are we under a roof?
    bTecho = IIf(MapData(UserPos.X, UserPos.Y).Trigger = 1 Or _
            MapData(UserPos.X, UserPos.Y).Trigger = 2 Or _
            MapData(UserPos.X, UserPos.Y).Trigger = 4, True, False)
                
    'Update pos label
    frmMain.Coord.Caption = "[" & UserMap & ", " & UserPos.X & ", " & UserPos.Y & "]"
End Sub

''
' Handles the NPCHitUser message.

Private Sub HandleNPCHitUser()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 4 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Select Case incomingData.ReadByte()
        Case bCabeza
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_CABEZA & CStr(incomingData.ReadInteger()) & "!!", 255, 0, 0, True, False, False)
        Case bBrazoIzquierdo
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_BRAZO_IZQ & CStr(incomingData.ReadInteger()) & "!!", 255, 0, 0, True, False, False)
        Case bBrazoDerecho
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_BRAZO_DER & CStr(incomingData.ReadInteger()) & "!!", 255, 0, 0, True, False, False)
        Case bPiernaIzquierda
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_PIERNA_IZQ & CStr(incomingData.ReadInteger()) & "!!", 255, 0, 0, True, False, False)
        Case bPiernaDerecha
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_PIERNA_DER & CStr(incomingData.ReadInteger()) & "!!", 255, 0, 0, True, False, False)
        Case bTorso
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_TORSO & CStr(incomingData.ReadInteger() & "!!"), 255, 0, 0, True, False, False)
    End Select
End Sub

''
' Handles the UserHitNPC message.

Private Sub HandleUserHitNPC()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_CRIATURA_1 & CStr(incomingData.ReadLong()) & MENSAJE_2, 255, 0, 0, True, False, False)
End Sub

''
' Handles the UserAttackedSwing message.

Private Sub HandleUserAttackedSwing()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & charlist(incomingData.ReadInteger()).Nombre & MENSAJE_ATAQUE_FALLO, 255, 0, 0, True, False, False)
End Sub

''
' Handles the UserHittingByUser message.

Private Sub HandleUserHittedByUser()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 6 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim attacker As String
    
    attacker = charlist(incomingData.ReadInteger()).Nombre
    
    Select Case incomingData.ReadByte
        Case bCabeza
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & attacker & MENSAJE_RECIVE_IMPACTO_CABEZA & CStr(incomingData.ReadInteger() & MENSAJE_2), 255, 0, 0, True, False, False)
        Case bBrazoIzquierdo
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & attacker & MENSAJE_RECIVE_IMPACTO_BRAZO_IZQ & CStr(incomingData.ReadInteger() & MENSAJE_2), 255, 0, 0, True, False, False)
        Case bBrazoDerecho
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & attacker & MENSAJE_RECIVE_IMPACTO_BRAZO_DER & CStr(incomingData.ReadInteger() & MENSAJE_2), 255, 0, 0, True, False, False)
        Case bPiernaIzquierda
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & attacker & MENSAJE_RECIVE_IMPACTO_PIERNA_IZQ & CStr(incomingData.ReadInteger() & MENSAJE_2), 255, 0, 0, True, False, False)
        Case bPiernaDerecha
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & attacker & MENSAJE_RECIVE_IMPACTO_PIERNA_DER & CStr(incomingData.ReadInteger() & MENSAJE_2), 255, 0, 0, True, False, False)
        Case bTorso
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & attacker & MENSAJE_RECIVE_IMPACTO_TORSO & CStr(incomingData.ReadInteger() & MENSAJE_2), 255, 0, 0, True, False, False)
    End Select
End Sub

''
' Handles the UserHittedUser message.

Private Sub HandleUserHittedUser()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 6 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim victim As String
    
    victim = charlist(incomingData.ReadInteger()).Nombre
    
    Select Case incomingData.ReadByte
        Case bCabeza
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & victim & MENSAJE_PRODUCE_IMPACTO_CABEZA & CStr(incomingData.ReadInteger() & MENSAJE_2), 255, 0, 0, True, False, False)
        Case bBrazoIzquierdo
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & victim & MENSAJE_PRODUCE_IMPACTO_BRAZO_IZQ & CStr(incomingData.ReadInteger() & MENSAJE_2), 255, 0, 0, True, False, False)
        Case bBrazoDerecho
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & victim & MENSAJE_PRODUCE_IMPACTO_BRAZO_DER & CStr(incomingData.ReadInteger() & MENSAJE_2), 255, 0, 0, True, False, False)
        Case bPiernaIzquierda
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & victim & MENSAJE_PRODUCE_IMPACTO_PIERNA_IZQ & CStr(incomingData.ReadInteger() & MENSAJE_2), 255, 0, 0, True, False, False)
        Case bPiernaDerecha
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & victim & MENSAJE_PRODUCE_IMPACTO_PIERNA_DER & CStr(incomingData.ReadInteger() & MENSAJE_2), 255, 0, 0, True, False, False)
        Case bTorso
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & victim & MENSAJE_PRODUCE_IMPACTO_TORSO & CStr(incomingData.ReadInteger() & MENSAJE_2), 255, 0, 0, True, False, False)
    End Select
End Sub

''
' Handles the ChatOverHead message.

Private Sub HandleChatOverHead()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 8 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim chat As String
    Dim CharIndex As Integer
    Dim r As Byte
    Dim g As Byte
    Dim B As Byte
    Dim Namer As Byte
    Dim Nameg As Byte
    Dim Nameb As Byte
    chat = Buffer.ReadASCIIString()
    CharIndex = Buffer.ReadInteger()
    
    r = Buffer.ReadByte()
    g = Buffer.ReadByte()
    B = Buffer.ReadByte()
    
   With charlist(CharIndex)
        If .priv = 0 Then
            If .Atacable Then
                Namer = ColoresPJ(48).r
                Nameg = ColoresPJ(48).g
                Nameb = ColoresPJ(48).B
            Else
                If .Criminal Then
                    Namer = ColoresPJ(50).r
                    Nameg = ColoresPJ(50).g
                    Nameb = ColoresPJ(50).B
                Else
                    Namer = ColoresPJ(49).r
                    Nameg = ColoresPJ(49).g
                    Nameb = ColoresPJ(49).B
                End If
            End If
         Else
            Namer = ColoresPJ(.priv).r
            Nameg = ColoresPJ(.priv).g
            Nameb = ColoresPJ(.priv).B
        End If

    'Only add the chat if the character exists (a CharacterRemove may have been sent to the PC / NPC area before the buffer was flushed)
    If charlist(CharIndex).Active Then _
        Call Dialogos.CreateDialog(chat, CharIndex, RGB(r, g, B))
        
    Dim Name As String
    Dim Pos As Integer
    Pos = InStr(charlist(CharIndex).Nombre, "<")
    If Pos = 0 Then Pos = Len(charlist(CharIndex).Nombre) + 2
   
    Name = Left$(charlist(CharIndex).Nombre, Pos - 2)
   
       
        If .Nombre <> "" Then _
            Call AddtoRichTextBox(frmMain.RecTxt, Name & "> ", Namer, Nameg, Nameb, True, False, True)
    
            Call AddtoRichTextBox(frmMain.RecTxt, chat, r, g, B, False, False)
    End With
    Call incomingData.CopyBuffer(Buffer)

ErrHandler:
    Dim error As Long
    error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub
Private Sub HandleChatOverHeadhit()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 8 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim chat As String
    Dim CharIndex As Integer
    Dim r As Byte
    Dim g As Byte
    Dim B As Byte
    
    chat = Buffer.ReadASCIIString()
    CharIndex = Buffer.ReadInteger()
    
    r = Buffer.ReadByte()
    g = Buffer.ReadByte()
    B = Buffer.ReadByte()
    
    'Only add the chat if the character exists (a CharacterRemove may have been sent to the PC / NPC area before the buffer was flushed)
    If charlist(CharIndex).Active Then _
        Call Dialogos.CreateDialog("       " & chat, CharIndex, RGB(r, g, B))
        
    Dim Name As String
    Dim Pos As Integer
    Pos = InStr(charlist(CharIndex).Nombre, "<")
    If Pos = 0 Then Pos = Len(charlist(CharIndex).Nombre) + 2
   
    Name = Left$(charlist(CharIndex).Nombre, Pos - 2)
   
    Call incomingData.CopyBuffer(Buffer)

ErrHandler:
    Dim error As Long
    error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub
''
' Handles the ConsoleMessage message.

Private Sub HandleConsoleMessage()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 4 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim chat As String
    Dim fontIndex As Integer
    Dim str As String
    Dim r As Byte
    Dim g As Byte
    Dim B As Byte
    
    chat = Buffer.ReadASCIIString()
    fontIndex = Buffer.ReadByte()
    
    If InStr(1, chat, "~") Then
        str = ReadField(2, chat, 126)
            If Val(str) > 255 Then
                r = 255
            Else
                r = Val(str)
            End If
            
            str = ReadField(3, chat, 126)
            If Val(str) > 255 Then
                g = 255
            Else
                g = Val(str)
            End If
            
            str = ReadField(4, chat, 126)
            If Val(str) > 255 Then
                B = 255
            Else
                B = Val(str)
            End If
            
        Call AddtoRichTextBox(frmMain.RecTxt, Left$(chat, InStr(1, chat, "~") - 1), r, g, B, Val(ReadField(5, chat, 126)) <> 0, Val(ReadField(6, chat, 126)) <> 0)
    Else
        With FontTypes(fontIndex)
            Call AddtoRichTextBox(frmMain.RecTxt, chat, .red, .green, .blue, .bold, .italic)
        End With
        
        ' Para no perder el foco cuando chatea por party
        If fontIndex = FontTypeNames.FONTTYPE_PARTY Then
            If MirandoParty Then frmParty.SendTxt.SetFocus
        End If
    End If
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the GuildChat message.

Private Sub HandleGuildChat()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/07/08 (NicoNZ)
'
'***************************************************
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim chat As String
    Dim str As String
    Dim r As Byte
    Dim g As Byte
    Dim B As Byte
    Dim tmp As Integer
    Dim Cont As Integer
    
    
    chat = Buffer.ReadASCIIString()
    
    With FontTypes(FontTypeNames.FONTTYPE_GUILDMSG)
    Call AddtoRichTextBox(frmMain.RecTxt, chat, .red, .green, .blue, .bold, .italic)
    End With
   
    Call DialogosClanes.PushBackText(ReadField(1, chat, 126))
    
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the ShowMessageBox message.

Private Sub HandleShowMessageBox()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    frmMensaje.msg.Caption = Buffer.ReadASCIIString()
    frmMensaje.Show
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the UserIndexInServer message.

Private Sub HandleUserIndexInServer()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserIndex = incomingData.ReadInteger()
End Sub

''
' Handles the UserCharIndexInServer message.

Private Sub HandleUserCharIndexInServer()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserCharIndex = incomingData.ReadInteger()
    UserPos = charlist(UserCharIndex).Pos
    
    Call DibujarMiniMapa
    
    'Are we under a roof?
    bTecho = IIf(MapData(UserPos.X, UserPos.Y).Trigger = 1 Or _
            MapData(UserPos.X, UserPos.Y).Trigger = 2 Or _
            MapData(UserPos.X, UserPos.Y).Trigger = 4, True, False)

    frmMain.Coord.Caption = "[" & UserMap & ", " & UserPos.X & ", " & UserPos.Y & "]"
End Sub

''
' Handles the CharacterCreate message.

Private Sub HandleCharacterCreate()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 26 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim CharIndex As Integer
    Dim Body As Integer
    Dim Head As Integer
    Dim Heading As E_Heading
    Dim X As Byte
    Dim Y As Byte
    Dim Weapon As Integer
    Dim Shield As Integer
    Dim helmet As Integer
    Dim privs As Integer
    
    CharIndex = Buffer.ReadInteger()
    Body = Buffer.ReadInteger()
    Head = Buffer.ReadInteger()
    Heading = Buffer.ReadByte()
    X = Buffer.ReadByte()
    Y = Buffer.ReadByte()
    Weapon = Buffer.ReadInteger()
    Shield = Buffer.ReadInteger()
    helmet = Buffer.ReadInteger()
    
    With charlist(CharIndex)
        Call SetCharacterFx(CharIndex, Buffer.ReadInteger(), Buffer.ReadInteger())
        
        .Nombre = Buffer.ReadASCIIString()
        .Criminal = Buffer.ReadByte()
        
        privs = Buffer.ReadByte()
        
        If privs <> 0 Then
            'If the player belongs to a council AND is an admin, only whos as an admin
            If (privs And PlayerType.ChaosCouncil) <> 0 And (privs And PlayerType.User) = 0 Then
                privs = privs Xor PlayerType.ChaosCouncil
            End If
            
            If (privs And PlayerType.RoyalCouncil) <> 0 And (privs And PlayerType.User) = 0 Then
                privs = privs Xor PlayerType.RoyalCouncil
            End If
            
            'If the player is a RM, ignore other flags
            If privs And PlayerType.RoleMaster Then
                privs = PlayerType.RoleMaster
            End If
            
            'Log2 of the bit flags sent by the server gives our numbers ^^
            .priv = Log(privs) / Log(2)
        Else
            .priv = 0
        End If
        charlist(CharIndex).Aura_Index = Buffer.ReadInteger()
    End With
    
    Call MakeChar(CharIndex, Body, Head, Heading, X, Y, Weapon, Shield, helmet, charlist(CharIndex).Aura_Index)
    
    Call RefreshAllChars
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

Private Sub HandleCharacterChangeNick()
'***************************************************
'Author: Budi
'Last Modification: 07/23/09
'
'***************************************************
    If incomingData.Length < 6 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet id
    Call incomingData.ReadByte
    Dim CharIndex As Integer
    CharIndex = incomingData.ReadInteger
    charlist(CharIndex).Nombre = incomingData.ReadASCIIString
    
End Sub

''
' Handles the CharacterRemove message.

Private Sub HandleCharacterRemove()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim CharIndex As Integer
    
    CharIndex = incomingData.ReadInteger()
    
    Call EraseChar(CharIndex)
    Call RefreshAllChars
End Sub

''
' Handles the CharacterMove message.

Private Sub HandleCharacterMove()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim CharIndex As Integer
    Dim X As Byte
    Dim Y As Byte
    
    CharIndex = incomingData.ReadInteger()
    X = incomingData.ReadByte()
    Y = incomingData.ReadByte()
    
    With charlist(CharIndex)
    
        ' Play steps sounds if the user is not an admin of any kind
        If .priv <> 1 And .priv <> 2 And .priv <> 3 And .priv <> 5 And .priv <> 25 Then
            Call DoPasosFx(CharIndex)
        End If
    End With
    
    Call MoveCharbyPos(CharIndex, X, Y)
    
    Call RefreshAllChars
End Sub

''
' Handles the ForceCharMove message.

Private Sub HandleForceCharMove()
    
    If incomingData.Length < 2 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim Direccion As Byte
    
    Direccion = incomingData.ReadByte()

    Call MoveCharbyHead(UserCharIndex, Direccion)
    Call MoveScreen(Direccion)
    
    Call RefreshAllChars
End Sub

''
' Handles the CharacterChange message.

Private Sub HandleCharacterChange()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 25/08/2009
'25/08/2009: ZaMa - Changed a variable used incorrectly.
'***************************************************
    If incomingData.Length < 18 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim CharIndex As Integer
    Dim tempint As Integer
    Dim headIndex As Integer
    
    CharIndex = incomingData.ReadInteger()
    
    With charlist(CharIndex)
        tempint = incomingData.ReadInteger()
        
        If tempint < LBound(BodyData()) Or tempint > UBound(BodyData()) Then
            .Body = BodyData(0)
            .iBody = 0
        Else
            .Body = BodyData(tempint)
            .iBody = tempint
        End If
        
        
        headIndex = incomingData.ReadInteger()
        
        If headIndex < LBound(HeadData()) Or headIndex > UBound(HeadData()) Then
            .Head = HeadData(0)
            .iHead = 0
        Else
            .Head = HeadData(headIndex)
            .iHead = headIndex
        End If
        
        .Muerto = (headIndex = CASPER_HEAD)
        
        .Heading = incomingData.ReadByte()
        
        tempint = incomingData.ReadInteger()
        If tempint <> 0 Then .arma = WeaponAnimData(tempint)
        
        tempint = incomingData.ReadInteger()
        If tempint <> 0 Then .escudo = ShieldAnimData(tempint)
        
        tempint = incomingData.ReadInteger()
        If tempint <> 0 Then .casco = CascoAnimData(tempint)
        
        Call SetCharacterFx(CharIndex, incomingData.ReadInteger(), incomingData.ReadInteger())
        
        charlist(CharIndex).Aura_Index = incomingData.ReadInteger()
        Call InitGrh(charlist(CharIndex).Aura, charlist(CharIndex).Aura_Index)
    End With
    
    Call RefreshAllChars
End Sub

''
' Handles the ObjectCreate message.

Private Sub HandleObjectCreate()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim X As Byte
    Dim Y As Byte
    
    X = incomingData.ReadByte()
    Y = incomingData.ReadByte()
    
    MapData(X, Y).ObjGrh.GrhIndex = incomingData.ReadInteger()
    
    If MapData(X, Y).ObjGrh.GrhIndex = 1521 Then
        Call General_Particle_Create(32, X, Y)
        Call Light.Create_Light_To_Map(X, Y, 3, 255, 255, 255)
    End If
    
    Call InitGrh(MapData(X, Y).ObjGrh, MapData(X, Y).ObjGrh.GrhIndex)
End Sub

''
' Handles the ObjectDelete message.

Private Sub HandleObjectDelete()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim X As Byte
    Dim Y As Byte
    
    X = incomingData.ReadByte()
    Y = incomingData.ReadByte()
    
    If MapData(X, Y).particle_group_index > 0 Then
        If Particle_Get_Type(MapData(X, Y).particle_group_index) = 32 Then
            Call Particle_Group_Remove(MapData(X, Y).particle_group_index)
        End If
    End If
    Call Light.Delete_Light_To_Map(X, Y)
    
    MapData(X, Y).ObjGrh.GrhIndex = 0
End Sub

''
' Handles the BlockPosition message.

Private Sub HandleBlockPosition()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 4 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim X As Byte
    Dim Y As Byte
    
    X = incomingData.ReadByte()
    Y = incomingData.ReadByte()
    
    If incomingData.ReadBoolean() Then
        MapData(X, Y).blocked = 1
    Else
        MapData(X, Y).blocked = 0
    End If
End Sub

''
' Handles the PlayMIDI message.

Private Sub HandlePlayMIDI()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 4 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    Dim currentMP3 As Byte
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    currentMP3 = incomingData.ReadByte()
    
    If currentMP3 Then
        General_Set_Song CStr(currentMP3), True
        Call incomingData.ReadInteger
    Else
        'Remove the bytes to prevent errors
        Call incomingData.ReadInteger
    End If
End Sub

''
' Handles the PlayWave message.

Private Sub HandlePlayWave()
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 08/14/07
'Last Modified by: Rapsodius
'Added support for 3D Sounds.
'***************************************************
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
        
    Dim wave As Integer
    Dim srcX As Byte
    Dim srcY As Byte
    
    wave = incomingData.ReadInteger()
    srcX = incomingData.ReadByte()
    srcY = incomingData.ReadByte()
        
    Call General_Set_Wav(CStr(wave) & ".wav", srcX, srcY)
End Sub

''
' Handles the GuildList message.

Private Sub HandleGuildList()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    'Clear guild's list
    frmGuildAdm.guildslist.Clear
    
    Dim guilds() As String
    guilds = Split(Buffer.ReadASCIIString(), SEPARATOR)
    
    Dim i As Long
    For i = 0 To UBound(guilds())
        Call frmGuildAdm.guildslist.AddItem(guilds(i))
    Next i
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
    frmGuildAdm.Show vbModeless, frmMain
    
ErrHandler:
    Dim error As Long
    error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the AreaChanged message.

Private Sub HandleAreaChanged()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim X As Byte
    Dim Y As Byte
    
    X = incomingData.ReadByte()
    Y = incomingData.ReadByte()
        
    Call CambioDeArea(X, Y)
End Sub

''
' Handles the PauseToggle message.

Private Sub HandlePauseToggle()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    pausa = Not pausa
End Sub

''
' Handles the CreateFX message.

Private Sub HandleCreateFX()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 7 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim CharIndex As Integer
    Dim fX As Integer
    Dim Loops As Integer
    
    CharIndex = incomingData.ReadInteger()
    fX = incomingData.ReadInteger()
    Loops = incomingData.ReadInteger()
    
    If fX = 14 And Opciones.SangreAct Then
        With charlist(CharIndex).Pos
            MapData(.X, .Y).Blood.Active = 1
            InitGrh MapData(.X, .Y).Blood.Grh, 24470
            MapData(.X, .Y).Blood.LifeTime = FPS * 10
            MapData(.X, .Y).Blood.Alpha = 255
            MapData(.X, .Y).Blood.Head = charlist(CharIndex).Heading
        End With
    End If
        Call SetCharacterFx(CharIndex, fX, Loops)
End Sub
''
' Handles the ParticleCreate message.
 
Private Sub HandleParticleCreate()
'***************************************************
'Author: Lorwik
'Last Modification: 11/09/2011
'
'***************************************************
   
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim Life As Integer
    Dim CharIndex As Integer
    Dim SeBorra As Byte
    
    
    CharIndex = incomingData.ReadInteger()
    charlist(CharIndex).ParticulaIndex = incomingData.ReadInteger()
    SeBorra = incomingData.ReadByte()
    
    With charlist(CharIndex)
        If Not SeBorra = 1 Then
            If Not .ParticulaIndex < 1 And Not .ParticulaIndex > TotalStreams Then
                Call General_Char_Particle_Create(.ParticulaIndex, CharIndex)
            End If
        Else
            If Not .ParticulaIndex < 1 And Not .ParticulaIndex > TotalStreams Then
                Call Char_Particle_Group_Remove(CharIndex, .ParticulaIndex)
                .ParticulaIndex = 0
            End If
        End If
    End With
    
End Sub

''
' Handles the UpdateUserStats message.
Private Sub HandleUpdateUserStats()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 26 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserMaxHP = incomingData.ReadInteger()
    UserMinHP = incomingData.ReadInteger()
    UserMaxMAN = incomingData.ReadInteger()
    UserMinMAN = incomingData.ReadInteger()
    UserMaxSTA = incomingData.ReadInteger()
    UserMinSTA = incomingData.ReadInteger()
    UserGLD = incomingData.ReadLong()
    UserLvl = incomingData.ReadByte()
    UserPasarNivel = incomingData.ReadLong()
    UserExp = incomingData.ReadLong()
    
    frmMain.exp.Caption = UserExp & " / " & UserPasarNivel
    frmMain.ExpShp.Width = (((UserExp / 100) / (UserPasarNivel / 100)) * 210)
    
    If frmMain.LvlLbl.Caption = 50 Then
        frmMain.lblPorcLvl.Caption = "¡Nivel Máximo!"
    Else
        If UserPasarNivel > 0 Then
            frmMain.lblPorcLvl.Caption = Round(CDbl(UserExp) * CDbl(100) / CDbl(UserPasarNivel), 2) & "%"
        Else
            frmMain.lblPorcLvl.Caption = "[N/A]"
        End If
    End If
    
    frmMain.Hpshp.Width = (((UserMinHP / 100) / (UserMaxHP / 100)) * 94)
    
    If UserMaxMAN > 0 Then
        frmMain.MANShp.Width = (((UserMinMAN + 1 / 100) / (UserMaxMAN + 1 / 100)) * 94)
    Else
        frmMain.MANShp.Width = 0
    End If
    
    frmMain.STAShp.Width = (((UserMinSTA / 100) / (UserMaxSTA / 100)) * 94)

    frmMain.GldLbl.Caption = UserGLD
    frmMain.LvlLbl.Caption = UserLvl
    
            'Stats
    frmMain.lblMana = UserMinMAN & "/" & UserMaxMAN
    frmMain.lblVida = UserMinHP & "/" & UserMaxHP
    frmMain.lblEnergia = UserMinSTA & "/" & UserMaxSTA
    
    If UserMinHP = 0 Then
        UserEstado = 1
    Else
        UserEstado = 0
    End If
End Sub

''
' Handles the WorkRequestTarget message.

Private Sub HandleWorkRequestTarget()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 2 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    UsingSkill = incomingData.ReadByte()

    frmMain.MousePointer = 2
    
    Select Case UsingSkill
        Case Magia
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_MAGIA, 100, 100, 120, 0, 0)
        Case Pesca
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_PESCA, 100, 100, 120, 0, 0)
        Case Robar
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_ROBAR, 100, 100, 120, 0, 0)
        Case Talar
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_TALAR, 100, 100, 120, 0, 0)
        Case Mineria
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_MINERIA, 100, 100, 120, 0, 0)
        Case FundirMetal
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_FUNDIRMETAL, 100, 100, 120, 0, 0)
        Case Proyectiles
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_PROYECTILES, 100, 100, 120, 0, 0)
    End Select
End Sub

''
' Handles the ChangeInventorySlot message.

Private Sub HandleChangeInventorySlot()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 22 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim slot As Byte
    Dim OBJIndex As Integer
    Dim Name As String
    Dim Amount As Integer
    Dim Equipped As Boolean
    Dim GrhIndex As Integer
    Dim OBJType As Byte
    Dim MaxHit As Integer
    Dim MinHit As Integer
    Dim defense As Integer
    Dim value As Single
    Dim PuedeUsar As Byte
    
    slot = Buffer.ReadByte()
    OBJIndex = Buffer.ReadInteger()
    Name = Buffer.ReadASCIIString()
    Amount = Buffer.ReadInteger()
    Equipped = Buffer.ReadBoolean()
    GrhIndex = Buffer.ReadInteger()
    OBJType = Buffer.ReadByte()
    MaxHit = Buffer.ReadInteger()
    MinHit = Buffer.ReadInteger()
    defense = Buffer.ReadInteger()
    value = Buffer.ReadSingle()
    PuedeUsar = Buffer.ReadByte
    
    Call Inventario.SetItem(slot, OBJIndex, Amount, Equipped, GrhIndex, OBJType, MaxHit, MinHit, defense, value, Name, PuedeUsar)
    
    'Call Inventario.UpdateInventory
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the ChangeBankSlot message.

Private Sub HandleChangeBankSlot()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 21 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim slot As Byte
    slot = Buffer.ReadByte()
    
    With UserBancoInventory(slot)
        .OBJIndex = Buffer.ReadInteger()
        .Name = Buffer.ReadASCIIString()
        .Amount = Buffer.ReadInteger()
        .GrhIndex = Buffer.ReadInteger()
        .OBJType = Buffer.ReadByte()
        .MaxHit = Buffer.ReadInteger()
        .MinHit = Buffer.ReadInteger()
        .Def = Buffer.ReadInteger()
        .Valor = Buffer.ReadLong()
        .PuedeUsar = Buffer.ReadByte()
        
        If Comerciando Then
            Call InvBanco(0).SetItem(slot, .OBJIndex, .Amount, _
                .Equipped, .GrhIndex, .OBJType, .MaxHit, _
                .MinHit, .Def, .Valor, .Name, .PuedeUsar)
        End If
    End With
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the ChangeSpellSlot message.

Private Sub HandleChangeSpellSlot()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 6 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim slot As Byte
    slot = Buffer.ReadByte()
    
    UserHechizos(slot) = Buffer.ReadInteger()
    
    If slot <= frmMain.hlst.ListCount Then
        frmMain.hlst.List(slot - 1) = Buffer.ReadASCIIString()
    Else
        Call frmMain.hlst.AddItem(Buffer.ReadASCIIString())
    End If
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the Attributes message.

Private Sub HandleAtributes()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 1 + NUMATRIBUTES Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim i As Long
    
    For i = 1 To NUMATRIBUTES
        UserAtributos(i) = incomingData.ReadByte()
    Next i
    
    'Show them in character creation
    If EstadoLogin = E_MODO.Dados Then
        With frmCrearPersonaje
            If .Visible Then
                .lbFuerza.Caption = UserAtributos(1)
                .lbAgilidad.Caption = UserAtributos(2)
                .lbInteligencia.Caption = UserAtributos(3)
                .lbCarisma.Caption = UserAtributos(4)
                .lbConstitucion.Caption = UserAtributos(5)
            End If
        End With
    Else
        LlegaronAtrib = True
    End If
End Sub

''
' Handles the BlacksmithWeapons message.

Private Sub HandleBlacksmithWeapons()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim count As Integer
    Dim i As Long
    Dim tmp As String
    
    count = Buffer.ReadInteger()
    
    Call frmHerrero.lstArmas.Clear
    
    For i = 1 To count
        tmp = Buffer.ReadASCIIString() & " ("           'Get the object's name
        tmp = tmp & CStr(Buffer.ReadInteger()) & ","    'The iron needed
        tmp = tmp & CStr(Buffer.ReadInteger()) & ","    'The silver needed
        tmp = tmp & CStr(Buffer.ReadInteger()) & ")"    'The gold needed
        
        Call frmHerrero.lstArmas.AddItem(tmp)
        ArmasHerrero(i) = Buffer.ReadInteger()
    Next i
    
    For i = i To UBound(ArmasHerrero())
        ArmasHerrero(i) = 0
    Next i
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the BlacksmithArmors message.

Private Sub HandleBlacksmithArmors()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim count As Integer
    Dim i As Long
    Dim tmp As String
    
    count = Buffer.ReadInteger()
    
    Call frmHerrero.lstArmaduras.Clear
    
    For i = 1 To count
        tmp = Buffer.ReadASCIIString() & " ("           'Get the object's name
        tmp = tmp & CStr(Buffer.ReadInteger()) & ","    'The iron needed
        tmp = tmp & CStr(Buffer.ReadInteger()) & ","    'The silver needed
        tmp = tmp & CStr(Buffer.ReadInteger()) & ")"    'The gold needed
        
        Call frmHerrero.lstArmaduras.AddItem(tmp)
        ArmadurasHerrero(i) = Buffer.ReadInteger()
    Next i
    
    For i = i To UBound(ArmadurasHerrero())
        ArmadurasHerrero(i) = 0
    Next i
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the CarpenterObjects message.

Private Sub HandleCarpenterObjects()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim count As Integer
    Dim i As Long
    Dim tmp As String
    
    count = Buffer.ReadInteger()
    
    Call frmCarp.lstArmas.Clear
    
    For i = 1 To count
        tmp = Buffer.ReadASCIIString() & " ("          'Get the object's name
        tmp = tmp & CStr(Buffer.ReadInteger()) & ")"    'The wood needed
        
        Call frmCarp.lstArmas.AddItem(tmp)
        ObjCarpintero(i) = Buffer.ReadInteger()
    Next i
    
    For i = i To UBound(ObjCarpintero())
        ObjCarpintero(i) = 0
    Next i
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the RestOK message.

Private Sub HandleRestOK()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserDescansar = Not UserDescansar
    UserAvisado = False
End Sub

''
' Handles the ErrorMessage message.

Private Sub HandleErrorMessage()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    frmMensaje.msg.Caption = Buffer.ReadASCIIString()
    frmMensaje.Show vbModal
    
    frmCrearPersonaje.MousePointer = 1
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the Blind message.

Private Sub HandleBlind()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserCiego = True
End Sub

''
' Handles the Dumb message.

Private Sub HandleDumb()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserEstupido = True
End Sub

''
' Handles the ShowSignal message.

Private Sub HandleShowSignal()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim tmp As String
    tmp = Buffer.ReadASCIIString()
    
    Call InitCartel(tmp, Buffer.ReadInteger())
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the ChangeNPCInventorySlot message.

Private Sub HandleChangeNPCInventorySlot()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 21 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim slot As Byte
    slot = Buffer.ReadByte()
    
    With NPCInventory(slot)
        .Name = Buffer.ReadASCIIString()
        .Amount = Buffer.ReadInteger()
        .Valor = Buffer.ReadSingle()
        .GrhIndex = Buffer.ReadInteger()
        .OBJIndex = Buffer.ReadInteger()
        .OBJType = Buffer.ReadByte()
        .MaxHit = Buffer.ReadInteger()
        .MinHit = Buffer.ReadInteger()
        .Def = Buffer.ReadInteger()
    End With
        
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the UpdateHungerAndThirst message.

Private Sub HandleUpdateHungerAndThirst()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserMaxAGU = incomingData.ReadByte()
    UserMinAGU = incomingData.ReadByte()
    UserMaxHAM = incomingData.ReadByte()
    UserMinHAM = incomingData.ReadByte()
    frmMain.lblHambre = UserMinHAM & "/" & UserMaxHAM
    frmMain.lblSed = UserMinAGU & "/" & UserMaxAGU
    frmMain.AGUAsp.Width = (((UserMinAGU / 100) / (UserMaxAGU / 100)) * 94)
    frmMain.COMIDAsp.Width = (((UserMinHAM / 100) / (UserMaxHAM / 100)) * 94)
End Sub

''
' Handles the Fame message.

Private Sub HandleFame()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 29 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    With UserReputacion
        .AsesinoRep = incomingData.ReadLong()
        .BandidoRep = incomingData.ReadLong()
        .BurguesRep = incomingData.ReadLong()
        .LadronesRep = incomingData.ReadLong()
        .NobleRep = incomingData.ReadLong()
        .PlebeRep = incomingData.ReadLong()
        .promedio = incomingData.ReadLong()
    End With
    
    LlegoFama = True
End Sub

''
' Handles the MiniStats message.

Private Sub HandleMiniStats()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 20 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    With UserEstadisticas
        .CiudadanosMatados = incomingData.ReadLong()
        .CriminalesMatados = incomingData.ReadLong()
        .UsuariosMatados = incomingData.ReadLong()
        .NpcsMatados = incomingData.ReadInteger()
        .Clase = ListaClases(incomingData.ReadByte())
        .PenaCarcel = incomingData.ReadLong()
    End With
End Sub

''
' Handles the LevelUp message.

Private Sub HandleLevelUp()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    SkillPoints = SkillPoints + incomingData.ReadInteger()
End Sub

''
' Handles the AddForumMessage message.

Private Sub HandleAddForumMessage()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim Title As String
    Dim Message As String
    
    Title = Buffer.ReadASCIIString()
    Message = Buffer.ReadASCIIString()
    
    Call frmForo.List.AddItem(Title)
    frmForo.Text(frmForo.List.ListCount - 1).Text = Message
    Call Load(frmForo.Text(frmForo.List.ListCount))
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the ShowForumForm message.

Private Sub HandleShowForumForm()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    If Not frmForo.Visible Then
        frmForo.Show , frmMain
    End If
End Sub

''
' Handles the SetInvisible message.

Private Sub HandleSetInvisible()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 4 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim CharIndex As Integer
    
    CharIndex = incomingData.ReadInteger()
    charlist(CharIndex).invisible = incomingData.ReadBoolean()

End Sub

''
' Handles the MeditateToggle message.

Private Sub HandleMeditateToggle()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserMeditar = Not UserMeditar
    UserAvisado = False
End Sub

''
' Handles the BlindNoMore message.

Private Sub HandleBlindNoMore()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserCiego = False
End Sub

''
' Handles the DumbNoMore message.

Private Sub HandleDumbNoMore()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserEstupido = False
End Sub

''
' Handles the SendSkills message.

Private Sub HandleSendSkills()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 1 + NUMSKILLS Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim i As Long
    
    For i = 1 To NUMSKILLS
        UserSkills(i) = incomingData.ReadByte()
    Next i
    LlegaronSkills = True
End Sub

''
' Handles the TrainerCreatureList message.

Private Sub HandleTrainerCreatureList()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim creatures() As String
    Dim i As Long
    
    creatures = Split(Buffer.ReadASCIIString(), SEPARATOR)
    
    For i = 0 To UBound(creatures())
        Call frmEntrenador.lstCriaturas.AddItem(creatures(i))
    Next i
    frmEntrenador.Show , frmMain
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the GuildNews message.

Private Sub HandleGuildNews()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 7 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim guildList() As String
    Dim i As Long
    
    'Get news' string
    frmGuildNews.news = Buffer.ReadASCIIString()
    
    'Get Enemy guilds list
    guildList = Split(Buffer.ReadASCIIString(), SEPARATOR)
    
    For i = 0 To UBound(guildList)
        Call frmGuildNews.guerra.AddItem(guildList(i))
    Next i
    
    'Get Allied guilds list
    guildList = Split(Buffer.ReadASCIIString(), SEPARATOR)
    
    For i = 0 To UBound(guildList)
        Call frmGuildNews.aliados.AddItem(guildList(i))
    Next i
    
    frmGuildNews.Show vbModeless, frmMain
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the OfferDetails message.

Private Sub HandleOfferDetails()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Call frmUserRequest.recievePeticion(Buffer.ReadASCIIString())
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the AlianceProposalsList message.

Private Sub HandleAlianceProposalsList()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim guildList() As String
    Dim i As Long
    
    guildList = Split(Buffer.ReadASCIIString(), SEPARATOR)
    
    Call frmPeaceProp.lista.Clear
    For i = 0 To UBound(guildList())
        Call frmPeaceProp.lista.AddItem(guildList(i))
    Next i
    
    frmPeaceProp.ProposalType = TIPO_PROPUESTA.ALIANZA
    Call frmPeaceProp.Show(vbModeless, frmMain)
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the PeaceProposalsList message.

Private Sub HandlePeaceProposalsList()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim guildList() As String
    Dim i As Long
    
    guildList = Split(Buffer.ReadASCIIString(), SEPARATOR)
    
    Call frmPeaceProp.lista.Clear
    For i = 0 To UBound(guildList())
        Call frmPeaceProp.lista.AddItem(guildList(i))
    Next i
    
    frmPeaceProp.ProposalType = TIPO_PROPUESTA.PAZ
    Call frmPeaceProp.Show(vbModeless, frmMain)
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the CharacterInfo message.

Private Sub HandleCharacterInfo()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 35 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    With frmCharInfo
        If .frmType = CharInfoFrmType.frmMembers Then
            .Rechazar.Visible = False
            .Aceptar.Visible = False
            .Echar.Visible = True
            .desc.Visible = False
        Else
            .Rechazar.Visible = True
            .Aceptar.Visible = True
            .Echar.Visible = False
            .desc.Visible = True
        End If
        
        .Nombre.Caption = "Nombre: " & Buffer.ReadASCIIString()
        .raza.Caption = "Raza: " & ListaRazas(Buffer.ReadByte())
        .Clase.Caption = "Clase: " & ListaClases(Buffer.ReadByte())
        
        If Buffer.ReadByte() = 1 Then
            .Genero.Caption = "Genero: Hombre"
        Else
            .Genero.Caption = "Genero: Mujer"
        End If
        
        .Nivel.Caption = "Nivel: " & Buffer.ReadByte()
        .Oro.Caption = "Oro: " & Buffer.ReadLong()
        .Banco.Caption = "Banco: " & Buffer.ReadLong()
        
        Dim reputation As Long
        reputation = Buffer.ReadLong()
        
        .reputacion.Caption = "Reputación: " & reputation
        
        .txtPeticiones.Text = Buffer.ReadASCIIString()
        .guildactual.Caption = "Clan: " & Buffer.ReadASCIIString()
        .txtMiembro.Text = Buffer.ReadASCIIString()
        
        Dim armada As Boolean
        Dim caos As Boolean
        
        armada = Buffer.ReadBoolean()
        caos = Buffer.ReadBoolean()
        
        If armada Then
            .ejercito.Caption = "Ejército: Armada Real"
        ElseIf caos Then
            .ejercito.Caption = "Ejército: Legión Oscura"
        End If
        
        .Ciudadanos.Caption = "Ciudadanos asesinados: " & CStr(Buffer.ReadLong())
        .criminales.Caption = "Criminales asesinados: " & CStr(Buffer.ReadLong())
        
        If reputation > 0 Then
            .status.Caption = " (Ciudadano)"
            .status.ForeColor = vbBlue
        Else
            .status.Caption = " (Criminal)"
            .status.ForeColor = vbRed
        End If
        
        Call .Show(vbModeless, frmMain)
    End With
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the GuildLeaderInfo message.

Private Sub HandleGuildLeaderInfo()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 9 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim List() As String
    Dim i As Long
    
    With frmGuildLeader
        'Get list of existing guilds
        List = Split(Buffer.ReadASCIIString(), SEPARATOR)
        
        'Empty the list
        Call .guildslist.Clear
        
        For i = 0 To UBound(List())
            Call .guildslist.AddItem(List(i))
        Next i
        
        'Get list of guild's members
        List = Split(Buffer.ReadASCIIString(), SEPARATOR)
        .Miembros.Caption = "El clan cuenta con " & CStr(UBound(List()) + 1) & " miembros."
        
        'Empty the list
        Call .members.Clear
        
        For i = 0 To UBound(List())
            Call .members.AddItem(List(i))
        Next i
        
        .txtguildnews = Buffer.ReadASCIIString()
        
        'Get list of join requests
        List = Split(Buffer.ReadASCIIString(), SEPARATOR)
        
        'Empty the list
        Call .solicitudes.Clear
        
        For i = 0 To UBound(List())
            Call .solicitudes.AddItem(List(i))
        Next i
        
        .Show , frmMain
    End With

    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the GuildDetails message.

Private Sub HandleGuildDetails()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 26 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    With frmGuildBrief
        If Not .EsLeader Then
            .guerra.Visible = False
            .aliado.Visible = False
            .Command3.Visible = False
        Else
            .guerra.Visible = True
            .aliado.Visible = True
            .Command3.Visible = True
        End If
        
        .Nombre.Caption = Buffer.ReadASCIIString()
        .fundador.Caption = Buffer.ReadASCIIString()
        .creacion.Caption = Buffer.ReadASCIIString()
        .lider.Caption = Buffer.ReadASCIIString()
        .web.Caption = Buffer.ReadASCIIString()
        .Miembros.Caption = Buffer.ReadInteger()
        
        If Buffer.ReadBoolean() Then
            .eleccion.Caption = "ABIERTA"
        Else
            .eleccion.Caption = "CERRADA"
        End If
        
        .lblAlineacion.Caption = Buffer.ReadASCIIString()
        .Enemigos.Caption = Buffer.ReadInteger()
        .aliados.Caption = Buffer.ReadInteger()
        .antifaccion.Caption = Buffer.ReadASCIIString()
        
        Dim codexStr() As String
        Dim i As Long
        
        codexStr = Split(Buffer.ReadASCIIString(), SEPARATOR)
        
        For i = 0 To 7
            .Codex(i).Caption = codexStr(i)
        Next i
        
        .desc.Caption = Buffer.ReadASCIIString()
    End With
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
    frmGuildBrief.Show vbModeless, frmMain
    
ErrHandler:
    Dim error As Long
    error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the ShowGuildFundationForm message.

Private Sub HandleShowGuildFundationForm()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    CreandoClan = True
    frmGuildFoundation.Show , frmMain
End Sub

''
' Handles the ParalizeOK message.

Private Sub HandleParalizeOK()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserParalizado = Not UserParalizado
End Sub

''
' Handles the ShowUserRequest message.

Private Sub HandleShowUserRequest()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Call frmUserRequest.recievePeticion(Buffer.ReadASCIIString())
    Call frmUserRequest.Show(vbModeless, frmMain)
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the TradeOK message.

Private Sub HandleTradeOK()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    If frmComerciar.Visible Then
        Dim i As Long
        
        'Update user inventory
        For i = 1 To MAX_INVENTORY_SLOTS
            ' Agrego o quito un item en su totalidad
            If Inventario.OBJIndex(i) <> InvComUsu.OBJIndex(i) Then
                With Inventario
                    Call InvComUsu.SetItem(i, .OBJIndex(i), _
                    .Amount(i), .Equipped(i), .GrhIndex(i), _
                    .OBJType(i), .MaxHit(i), .MinHit(i), .Def(i), _
                    .Valor(i), .ItemName(i))
                End With
            ' Vendio o compro cierta cantidad de un item que ya tenia
            ElseIf Inventario.Amount(i) <> InvComUsu.Amount(i) Then
                Call InvComUsu.ChangeSlotItemAmount(i, Inventario.Amount(i))
            End If
        Next i
        
        ' Fill Npc inventory
        For i = 1 To 20
            ' Compraron la totalidad de un item, o vendieron un item que el npc no tenia
            If NPCInventory(i).OBJIndex <> InvComNpc.OBJIndex(i) Then
                With NPCInventory(i)
                    Call InvComNpc.SetItem(i, .OBJIndex, _
                    .Amount, 0, .GrhIndex, _
                    .OBJType, .MaxHit, .MinHit, .Def, _
                    .Valor, .Name)
                End With
            ' Compraron o vendieron cierta cantidad (no su totalidad)
            ElseIf NPCInventory(i).Amount <> InvComNpc.Amount(i) Then
                Call InvComNpc.ChangeSlotItemAmount(i, NPCInventory(i).Amount)
            End If
        Next i
    
    End If
End Sub

''
' Handles the BankOK message.

Private Sub HandleBankOK()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim i As Long
    
    If frmBancoObj.Visible Then
        
        For i = 1 To Inventario.MaxObjs
            With Inventario
                Call InvBanco(1).SetItem(i, .OBJIndex(i), .Amount(i), _
                    .Equipped(i), .GrhIndex(i), .OBJType(i), .MaxHit(i), _
                    .MinHit(i), .Def(i), .Valor(i), .ItemName(i))
            End With
        Next i
        
        'Alter order according to if we bought or sold so the labels and grh remain the same
        If frmBancoObj.LasActionBuy Then
            'frmBancoObj.List1(1).ListIndex = frmBancoObj.LastIndex2
            'frmBancoObj.List1(0).ListIndex = frmBancoObj.LastIndex1
        Else
            'frmBancoObj.List1(0).ListIndex = frmBancoObj.LastIndex1
            'frmBancoObj.List1(1).ListIndex = frmBancoObj.LastIndex2
        End If
        
        frmBancoObj.NoPuedeMover = False
    End If
       
End Sub
''
' Handles the ChangeUserTradeSlot message.

Private Sub HandleChangeUserTradeSlot()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 22 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    With OtroInventario(1)
        .OBJIndex = Buffer.ReadInteger()
        .Name = Buffer.ReadASCIIString()
        .Amount = Buffer.ReadLong()
        .GrhIndex = Buffer.ReadInteger()
        .OBJType = Buffer.ReadByte()
        .MaxHit = Buffer.ReadInteger()
        .MinHit = Buffer.ReadInteger()
        .Def = Buffer.ReadInteger()
        .Valor = Buffer.ReadLong()
        .PuedeUsar = Buffer.ReadByte()
        
        frmComerciarUsu.List2.Clear
        
        Call frmComerciarUsu.List2.AddItem(.Name)
        frmComerciarUsu.List2.ItemData(frmComerciarUsu.List2.NewIndex) = .Amount
        
        frmComerciarUsu.lblEstadoResp.Visible = False
    End With
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the SpawnList message.

Private Sub HandleSpawnList()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim creatureList() As String
    Dim i As Long
    
    creatureList = Split(Buffer.ReadASCIIString(), SEPARATOR)
    
    For i = 0 To UBound(creatureList())
        Call frmSpawnList.lstCriaturas.AddItem(creatureList(i))
    Next i
    frmSpawnList.Show , frmMain
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the ShowSOSForm message.

Private Sub HandleShowSOSForm()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Modificado: By Lorwik
'Last Modification: 05/17/06
'Lorwik> Le agrege un motivo de la consulta.
'***************************************************
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim sosList() As String
    'Dim NameUser() As String
    Dim i As Long
    
    sosList = Split(Buffer.ReadASCIIString(), SEPARATOR)
    
    For i = 0 To UBound(sosList())
        Call frmMSG.List1.AddItem(sosList(i))
    Next i
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
    frmMSG.Show , frmMain
    
ErrHandler:
    Dim error As Long
    error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the ShowMOTDEditionForm message.

Private Sub HandleShowMOTDEditionForm()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    frmCambiaMotd.txtMotd.Text = Buffer.ReadASCIIString()
    frmCambiaMotd.Show , frmMain
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the ShowGMPanelForm message.

Private Sub HandleShowGMPanelForm()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    frmPanelGm.Show vbModeless, frmMain
End Sub

''
' Handles the UserNameList message.

Private Sub HandleUserNameList()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim userList() As String
    Dim i As Long
    
    userList = Split(Buffer.ReadASCIIString(), SEPARATOR)
    
    If frmPanelGm.Visible Then
        frmPanelGm.cboListaUsus.Clear
        For i = 0 To UBound(userList())
            Call frmPanelGm.cboListaUsus.AddItem(userList(i))
        Next i
        If frmPanelGm.cboListaUsus.ListCount > 0 Then frmPanelGm.cboListaUsus.ListIndex = 0
    End If
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the Pong message.

Private Sub HandlePong()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    Call incomingData.ReadByte
    
    Call AddtoRichTextBox(frmMain.RecTxt, "El ping es " & (GetTickCount - pingTime) & " ms.", 255, 0, 0, True, False, False)
    
    pingTime = 0
End Sub

''
' Handles the UpdateTag message.

Private Sub HandleUpdateTagAndStatus()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 6 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim CharIndex As Integer
    Dim Criminal As Boolean
    Dim userTag As String
    
    CharIndex = Buffer.ReadInteger()
    Criminal = Buffer.ReadBoolean()
    userTag = Buffer.ReadASCIIString()
    
    'Update char status adn tag!
    With charlist(CharIndex)
        If Criminal Then
            .Criminal = 1
        Else
            .Criminal = 0
        End If
        
        .Nombre = userTag
    End With
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub


''
' Writes the "LoginExistingChar" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLoginExistingChar()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "LoginExistingChar" message to the outgoing data buffer
'***************************************************
    Dim i As Long
    
    With outgoingData
        Call .WriteByte(ClientPacketID.LoginExistingChar)
        
        Call .WriteASCIIString(Cuenta.Name)
        Call .WriteASCIIString(Cuenta.Pass)
        Call .WriteByte(PJClickeado)
        Call .WriteByte(App.Major)
        Call .WriteByte(App.Minor)
        Call .WriteByte(App.Revision)
    End With
End Sub

''
' Writes the "LoginNewChar" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLoginNewChar()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "LoginNewChar" message to the outgoing data buffer
'***************************************************
    Dim i As Integer
    
    With outgoingData
        Call .WriteByte(ClientPacketID.LoginNewChar)
        
        Call .WriteASCIIString(UserName)
        
        Call .WriteByte(App.Major)
        Call .WriteByte(App.Minor)
        Call .WriteByte(App.Revision)
        
        Call .WriteByte(UserRaza)
        Call .WriteByte(UserSexo)
        Call .WriteByte(UserClase)
        Call .WriteInteger(MiCabeza)
        
        Call .WriteASCIIString(Cuenta.Name)
        
        For i = 1 To 5
        Call .WriteByte(UserAtributos(i))
        Next i
        
    End With
End Sub

''
' Writes the "Talk" message to the outgoing data buffer.
'
' @param    chat The chat text to be sent.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTalk(ByVal chat As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Talk" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Talk)
        
        Call .WriteASCIIString(chat)
    End With
End Sub

''
' Writes the "Yell" message to the outgoing data buffer.
'
' @param    chat The chat text to be sent.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteYell(ByVal chat As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Yell" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Yell)
        
        Call .WriteASCIIString(chat)
    End With
End Sub

''
' Writes the "Whisper" message to the outgoing data buffer.
'
' @param    charIndex The index of the char to whom to whisper.
' @param    chat The chat text to be sent to the user.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWhisper(ByVal CharIndex As Integer, ByVal chat As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Whisper" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Whisper)
        
        Call .WriteInteger(CharIndex)
        
        Call .WriteASCIIString(chat)
    End With
End Sub

''
' Writes the "Walk" message to the outgoing data buffer.
'
' @param    heading The direction in wich the user is moving.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWalk(ByVal Heading As E_Heading)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Walk" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Walk)
        
        Call .WriteByte(Heading)
    End With
End Sub

''
' Writes the "RequestPositionUpdate" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestPositionUpdate()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestPositionUpdate" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.RequestPositionUpdate)
End Sub

''
' Writes the "Attack" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAttack()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Attack" message to the outgoing data buffer
'***************************************************
    IsAttacking = True
    
    With charlist(UserCharIndex)
        .arma.WeaponWalk(.Heading).Started = 1
        .escudo.ShieldWalk(.Heading).Started = 1
    End With
    
    Call outgoingData.WriteByte(ClientPacketID.Attack)
End Sub

''
' Writes the "PickUp" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePickUp()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PickUp" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.PickUp)
End Sub

''
' Writes the "CombatModeToggle" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCombatModeToggle()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CombatModeToggle" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.CombatModeToggle)
End Sub

''
' Writes the "ResuscitationSafeToggle" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteResuscitationToggle()
'**************************************************************
'Author: Rapsodius
'Creation Date: 10/10/07
'Writes the Resuscitation safe toggle packet to the outgoing data buffer.
'**************************************************************
    Call outgoingData.WriteByte(ClientPacketID.ResuscitationSafeToggle)
End Sub

''
' Writes the "RequestGuildLeaderInfo" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestGuildLeaderInfo()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestGuildLeaderInfo" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.RequestGuildLeaderInfo)
End Sub

''
' Writes the "RequestAtributes" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestAtributes()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestAtributes" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.RequestAtributes)
End Sub

''
' Writes the "RequestFame" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestFame()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestFame" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.RequestFame)
End Sub

''
' Writes the "RequestSkills" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestSkills()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestSkills" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.RequestSkills)
End Sub

''
' Writes the "RequestMiniStats" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestMiniStats()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestMiniStats" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.RequestMiniStats)
End Sub

''
' Writes the "CommerceEnd" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceEnd()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CommerceEnd" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.CommerceEnd)
End Sub

''
' Writes the "UserCommerceEnd" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceEnd()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UserCommerceEnd" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.UserCommerceEnd)
End Sub

''
' Writes the "BankEnd" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankEnd()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BankEnd" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.BankEnd)
End Sub

''
' Writes the "UserCommerceOk" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceOk()
'***************************************************
'Author: Fredy Horacio Treboux (liquid)
'Last Modification: 01/10/07
'Writes the "UserCommerceOk" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.UserCommerceOk)
End Sub

''
' Writes the "UserCommerceReject" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceReject()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UserCommerceReject" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.UserCommerceReject)
End Sub

''
' Writes the "Drop" message to the outgoing data buffer.
'
' @param    slot Inventory slot where the item to drop is.
' @param    amount Number of items to drop.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDrop(ByVal slot As Byte, ByVal Amount As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Drop" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Drop)
        
        Call .WriteByte(slot)
        Call .WriteInteger(Amount)
    End With
End Sub

''
' Writes the "CastSpell" message to the outgoing data buffer.
'
' @param    slot Spell List slot where the spell to cast is.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCastSpell(ByVal slot As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CastSpell" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.CastSpell)
        
        Call .WriteByte(slot)
    End With
End Sub

''
' Writes the "LeftClick" message to the outgoing data buffer.
'
' @param    x Tile coord in the x-axis in which the user clicked.
' @param    y Tile coord in the y-axis in which the user clicked.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLeftClick(ByVal X As Byte, ByVal Y As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "LeftClick" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.LeftClick)
        
        Call .WriteByte(X)
        Call .WriteByte(Y)
    End With
End Sub

''
' Writes the "DoubleClick" message to the outgoing data buffer.
'
' @param    x Tile coord in the x-axis in which the user clicked.
' @param    y Tile coord in the y-axis in which the user clicked.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDoubleClick(ByVal X As Byte, ByVal Y As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "DoubleClick" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.DoubleClick)
        
        Call .WriteByte(X)
        Call .WriteByte(Y)
    End With
End Sub

''
' Writes the "Work" message to the outgoing data buffer.
'
' @param    skill The skill which the user attempts to use.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWork(ByVal Skill As eSkill)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Work" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Work)
        
        Call .WriteByte(Skill)
    End With
End Sub

''
' Writes the "UseSpellMacro" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUseSpellMacro()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UseSpellMacro" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.UseSpellMacro)
End Sub

''
' Writes the "UseItem" message to the outgoing data buffer.
'
' @param    slot Invetory slot where the item to use is.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUseItem(ByVal slot As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UseItem" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.UseItem)
        
        Call .WriteByte(slot)
    End With
End Sub

''
' Writes the "CraftBlacksmith" message to the outgoing data buffer.
'
' @param    item Index of the item to craft in the list sent by the server.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCraftBlacksmith(ByVal Item As Integer, ByVal cantidad As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CraftBlacksmith" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.CraftBlacksmith)
        
        Call .WriteInteger(Item)
        Call .WriteInteger(cantidad)
    End With
End Sub

''
' Writes the "CraftCarpenter" message to the outgoing data buffer.
'
' @param    item Index of the item to craft in the list sent by the server.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCraftCarpenter(ByVal Item As Integer, ByVal cantidad As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CraftCarpenter" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.CraftCarpenter)
        
        Call .WriteInteger(Item)
        Call .WriteInteger(cantidad)
    End With
End Sub

''
' Writes the "WorkLeftClick" message to the outgoing data buffer.
'
' @param    x Tile coord in the x-axis in which the user clicked.
' @param    y Tile coord in the y-axis in which the user clicked.
' @param    skill The skill which the user attempts to use.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWorkLeftClick(ByVal X As Byte, ByVal Y As Byte, ByVal Skill As eSkill)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "WorkLeftClick" message to the outgoing data buffer
'***************************************************
    With outgoingData
        IsAttacking = True
        charlist(UserCharIndex).arma.WeaponWalk(charlist(UserCharIndex).Heading).Started = 1
        
        Call .WriteByte(ClientPacketID.WorkLeftClick)
        
        Call .WriteByte(X)
        Call .WriteByte(Y)
        
        Call .WriteByte(Skill)
    End With
End Sub

''
' Writes the "CreateNewGuild" message to the outgoing data buffer.
'
' @param    desc    The guild's description
' @param    name    The guild's name
' @param    site    The guild's website
' @param    codex   Array of all rules of the guild.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreateNewGuild(ByVal desc As String, ByVal Name As String, ByVal Site As String, ByRef Codex() As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CreateNewGuild" message to the outgoing data buffer
'***************************************************
    Dim temp As String
    Dim i As Long
    
    With outgoingData
        Call .WriteByte(ClientPacketID.CreateNewGuild)
        
        Call .WriteASCIIString(desc)
        Call .WriteASCIIString(Name)
        Call .WriteASCIIString(Site)
        
        For i = LBound(Codex()) To UBound(Codex())
            temp = temp & Codex(i) & SEPARATOR
        Next i
        
        If Len(temp) Then _
            temp = Left$(temp, Len(temp) - 1)
        
        Call .WriteASCIIString(temp)
    End With
End Sub

''
' Writes the "SpellInfo" message to the outgoing data buffer.
'
' @param    slot Spell List slot where the spell which's info is requested is.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSpellInfo(ByVal slot As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SpellInfo" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.SpellInfo)
        
        Call .WriteByte(slot)
    End With
End Sub

''
' Writes the "EquipItem" message to the outgoing data buffer.
'
' @param    slot Invetory slot where the item to equip is.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteEquipItem(ByVal slot As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "EquipItem" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.EquipItem)
        
        Call .WriteByte(slot)
    End With
End Sub

''
' Writes the "ChangeHeading" message to the outgoing data buffer.
'
' @param    heading The direction in wich the user is moving.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeHeading(ByVal Heading As E_Heading)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChangeHeading" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.ChangeHeading)
        
        Call .WriteByte(Heading)
    End With
End Sub

''
' Writes the "ModifySkills" message to the outgoing data buffer.
'
' @param    skillEdt a-based array containing for each skill the number of points to add to it.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteModifySkills(ByRef skillEdt() As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ModifySkills" message to the outgoing data buffer
'***************************************************
    Dim i As Long
    
    With outgoingData
        Call .WriteByte(ClientPacketID.ModifySkills)
        
        For i = 1 To NUMSKILLS
            Call .WriteByte(skillEdt(i))
        Next i
    End With
End Sub

''
' Writes the "Train" message to the outgoing data buffer.
'
' @param    creature Position within the list provided by the server of the creature to train against.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTrain(ByVal creature As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Train" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Train)
        
        Call .WriteByte(creature)
    End With
End Sub

''
' Writes the "CommerceBuy" message to the outgoing data buffer.
'
' @param    slot Position within the NPC's inventory in which the desired item is.
' @param    amount Number of items to buy.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceBuy(ByVal slot As Byte, ByVal Amount As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CommerceBuy" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.CommerceBuy)
        
        Call .WriteByte(slot)
        Call .WriteInteger(Amount)
    End With
End Sub

''
' Writes the "BankExtractItem" message to the outgoing data buffer.
'
' @param    slot Position within the bank in which the desired item is.
' @param    amount Number of items to extract.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankExtractItem(ByVal slot As Byte, ByVal Amount As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BankExtractItem" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.BankExtractItem)
        
        Call .WriteByte(slot)
        Call .WriteInteger(Amount)
    End With
End Sub

''
' Writes the "CommerceSell" message to the outgoing data buffer.
'
' @param    slot Position within user inventory in which the desired item is.
' @param    amount Number of items to sell.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceSell(ByVal slot As Byte, ByVal Amount As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CommerceSell" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.CommerceSell)
        
        Call .WriteByte(slot)
        Call .WriteInteger(Amount)
    End With
End Sub

''
' Writes the "BankDeposit" message to the outgoing data buffer.
'
' @param    slot Position within the user inventory in which the desired item is.
' @param    amount Number of items to deposit.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankDeposit(ByVal slot As Byte, ByVal Amount As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BankDeposit" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.BankDeposit)
        
        Call .WriteByte(slot)
        Call .WriteInteger(Amount)
    End With
End Sub

''
' Writes the "ForumPost" message to the outgoing data buffer.
'
' @param    title The message's title.
' @param    message The body of the message.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteForumPost(ByVal Title As String, ByVal Message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ForumPost" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.ForumPost)
        
        Call .WriteASCIIString(Title)
        Call .WriteASCIIString(Message)
    End With
End Sub

''
' Writes the "MoveSpell" message to the outgoing data buffer.
'
' @param    upwards True if the spell will be moved up in the list, False if it will be moved downwards.
' @param    slot Spell List slot where the spell which's info is requested is.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMoveSpell(ByVal upwards As Boolean, ByVal slot As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "MoveSpell" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.MoveSpell)
        
        Call .WriteBoolean(upwards)
        Call .WriteByte(slot)
    End With
End Sub

''
' Writes the "MoveBank" message to the outgoing data buffer.
'
' @param    upwards True if the item will be moved up in the list, False if it will be moved downwards.
' @param    slot Bank List slot where the item which's info is requested is.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMoveBank(ByVal upwards As Boolean, ByVal slot As Byte)
'***************************************************
'Author: Torres Patricio (Pato)
'Last Modification: 06/14/09
'Writes the "MoveBank" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.MoveBank)
        
        Call .WriteBoolean(upwards)
        Call .WriteByte(slot)
    End With
End Sub

''
' Writes the "ClanCodexUpdate" message to the outgoing data buffer.
'
' @param    desc New description of the clan.
' @param    codex New codex of the clan.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteClanCodexUpdate(ByVal desc As String, ByRef Codex() As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ClanCodexUpdate" message to the outgoing data buffer
'***************************************************
    Dim temp As String
    Dim i As Long
    
    With outgoingData
        Call .WriteByte(ClientPacketID.ClanCodexUpdate)
        
        Call .WriteASCIIString(desc)
        
        For i = LBound(Codex()) To UBound(Codex())
            temp = temp & Codex(i) & SEPARATOR
        Next i
        
        If Len(temp) Then _
            temp = Left$(temp, Len(temp) - 1)
        
        Call .WriteASCIIString(temp)
    End With
End Sub

''
' Writes the "UserCommerceOffer" message to the outgoing data buffer.
'
' @param    slot Position within user inventory in which the desired item is.
' @param    amount Number of items to offer.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceOffer(ByVal slot As Byte, ByVal Amount As Long)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UserCommerceOffer" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.UserCommerceOffer)
        
        Call .WriteByte(slot)
        Call .WriteLong(Amount)
    End With
End Sub

''
' Writes the "GuildAcceptPeace" message to the outgoing data buffer.
'
' @param    guild The guild whose peace offer is accepted.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildAcceptPeace(ByVal guild As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildAcceptPeace" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildAcceptPeace)
        
        Call .WriteASCIIString(guild)
    End With
End Sub

''
' Writes the "GuildRejectAlliance" message to the outgoing data buffer.
'
' @param    guild The guild whose aliance offer is rejected.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildRejectAlliance(ByVal guild As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildRejectAlliance" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildRejectAlliance)
        
        Call .WriteASCIIString(guild)
    End With
End Sub

''
' Writes the "GuildRejectPeace" message to the outgoing data buffer.
'
' @param    guild The guild whose peace offer is rejected.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildRejectPeace(ByVal guild As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildRejectPeace" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildRejectPeace)
        
        Call .WriteASCIIString(guild)
    End With
End Sub

''
' Writes the "GuildAcceptAlliance" message to the outgoing data buffer.
'
' @param    guild The guild whose aliance offer is accepted.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildAcceptAlliance(ByVal guild As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildAcceptAlliance" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildAcceptAlliance)
        
        Call .WriteASCIIString(guild)
    End With
End Sub

''
' Writes the "GuildOfferPeace" message to the outgoing data buffer.
'
' @param    guild The guild to whom peace is offered.
' @param    proposal The text to send with the proposal.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildOfferPeace(ByVal guild As String, ByVal proposal As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildOfferPeace" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildOfferPeace)
        
        Call .WriteASCIIString(guild)
        Call .WriteASCIIString(proposal)
    End With
End Sub

''
' Writes the "GuildOfferAlliance" message to the outgoing data buffer.
'
' @param    guild The guild to whom an aliance is offered.
' @param    proposal The text to send with the proposal.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildOfferAlliance(ByVal guild As String, ByVal proposal As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildOfferAlliance" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildOfferAlliance)
        
        Call .WriteASCIIString(guild)
        Call .WriteASCIIString(proposal)
    End With
End Sub

''
' Writes the "GuildAllianceDetails" message to the outgoing data buffer.
'
' @param    guild The guild whose aliance proposal's details are requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildAllianceDetails(ByVal guild As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildAllianceDetails" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildAllianceDetails)
        
        Call .WriteASCIIString(guild)
    End With
End Sub

''
' Writes the "GuildPeaceDetails" message to the outgoing data buffer.
'
' @param    guild The guild whose peace proposal's details are requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildPeaceDetails(ByVal guild As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildPeaceDetails" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildPeaceDetails)
        
        Call .WriteASCIIString(guild)
    End With
End Sub

''
' Writes the "GuildRequestJoinerInfo" message to the outgoing data buffer.
'
' @param    username The user who wants to join the guild whose info is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildRequestJoinerInfo(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildRequestJoinerInfo" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildRequestJoinerInfo)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "GuildAlliancePropList" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildAlliancePropList()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildAlliancePropList" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GuildAlliancePropList)
End Sub

''
' Writes the "GuildPeacePropList" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildPeacePropList()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildPeacePropList" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GuildPeacePropList)
End Sub

''
' Writes the "GuildDeclareWar" message to the outgoing data buffer.
'
' @param    guild The guild to which to declare war.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildDeclareWar(ByVal guild As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildDeclareWar" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildDeclareWar)
        
        Call .WriteASCIIString(guild)
    End With
End Sub

''
' Writes the "GuildNewWebsite" message to the outgoing data buffer.
'
' @param    url The guild's new website's URL.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildNewWebsite(ByVal url As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildNewWebsite" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildNewWebsite)
        
        Call .WriteASCIIString(url)
    End With
End Sub

''
' Writes the "GuildAcceptNewMember" message to the outgoing data buffer.
'
' @param    username The name of the accepted player.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildAcceptNewMember(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildAcceptNewMember" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildAcceptNewMember)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "GuildRejectNewMember" message to the outgoing data buffer.
'
' @param    username The name of the rejected player.
' @param    reason The reason for which the player was rejected.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildRejectNewMember(ByVal UserName As String, ByVal reason As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildRejectNewMember" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildRejectNewMember)
        
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(reason)
    End With
End Sub

''
' Writes the "GuildKickMember" message to the outgoing data buffer.
'
' @param    username The name of the kicked player.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildKickMember(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildKickMember" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildKickMember)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "GuildUpdateNews" message to the outgoing data buffer.
'
' @param    news The news to be posted.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildUpdateNews(ByVal news As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildUpdateNews" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildUpdateNews)
        
        Call .WriteASCIIString(news)
    End With
End Sub

''
' Writes the "GuildMemberInfo" message to the outgoing data buffer.
'
' @param    username The user whose info is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildMemberInfo(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildMemberInfo" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildMemberInfo)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "GuildOpenElections" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildOpenElections()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildOpenElections" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GuildOpenElections)
End Sub

''
' Writes the "GuildRequestMembership" message to the outgoing data buffer.
'
' @param    guild The guild to which to request membership.
' @param    application The user's application sheet.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildRequestMembership(ByVal guild As String, ByVal Application As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildRequestMembership" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildRequestMembership)
        
        Call .WriteASCIIString(guild)
        Call .WriteASCIIString(Application)
    End With
End Sub

''
' Writes the "GuildRequestDetails" message to the outgoing data buffer.
'
' @param    guild The guild whose details are requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildRequestDetails(ByVal guild As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildRequestDetails" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildRequestDetails)
        
        Call .WriteASCIIString(guild)
    End With
End Sub

''
' Writes the "Online" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteOnline()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Online" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.Online)
End Sub

''
' Writes the "Quit" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteQuit()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 08/16/08
'Writes the "Quit" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.Quit)
End Sub

''
' Writes the "GuildLeave" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildLeave()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildLeave" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GuildLeave)
End Sub

''
' Writes the "RequestAccountState" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestAccountState()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestAccountState" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.RequestAccountState)
End Sub

''
' Writes the "PetStand" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePetStand()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PetStand" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.PetStand)
End Sub

''
' Writes the "PetFollow" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePetFollow()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PetFollow" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.PetFollow)
End Sub

''
' Writes the "TrainList" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTrainList()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "TrainList" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.TrainList)
End Sub

''
' Writes the "Rest" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRest()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Rest" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.Rest)
End Sub

''
' Writes the "Meditate" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMeditate()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Meditate" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.Meditate)
End Sub

''
' Writes the "Resucitate" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteResucitate()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Resucitate" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.Resucitate)
End Sub

''
' Writes the "Heal" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteHeal()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Heal" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.Heal)
End Sub
''
' Writes the "CommerceStart" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceStart()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CommerceStart" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.CommerceStart)
End Sub

''
' Writes the "BankStart" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankStart()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BankStart" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.BankStart)
End Sub

''
' Writes the "Enlist" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteEnlist()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Enlist" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.Enlist)
End Sub

''
' Writes the "Information" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteInformation()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Information" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.Information)
End Sub

''
' Writes the "Reward" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteReward()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Reward" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.Reward)
End Sub

''
' Writes the "RequestMOTD" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestMOTD()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestMOTD" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.RequestMOTD)
End Sub

''
' Writes the "PartyLeave" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePartyLeave()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PartyLeave" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.PartyLeave)
End Sub

''
' Writes the "PartyCreate" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePartyCreate()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PartyCreate" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.PartyCreate)
End Sub

''
' Writes the "PartyJoin" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePartyJoin()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PartyJoin" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.PartyJoin)
End Sub

''
' Writes the "GuildMessage" message to the outgoing data buffer.
'
' @param    message The message to send to the guild.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildMessage(ByVal Message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildRequestDetails" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildMessage)
        
        Call .WriteASCIIString(Message)
    End With
End Sub

''
' Writes the "PartyMessage" message to the outgoing data buffer.
'
' @param    message The message to send to the party.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePartyMessage(ByVal Message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PartyMessage" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.PartyMessage)
        
        Call .WriteASCIIString(Message)
    End With
End Sub

''
' Writes the "CentinelReport" message to the outgoing data buffer.
'
' @param    number The number to report to the centinel.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCentinelReport(ByVal number As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CentinelReport" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.CentinelReport)
        
        Call .WriteInteger(number)
    End With
End Sub

''
' Writes the "GuildOnline" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildOnline()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildOnline" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GuildOnline)
End Sub

''
' Writes the "PartyOnline" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePartyOnline()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PartyOnline" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.PartyOnline)
End Sub

''
' Writes the "CouncilMessage" message to the outgoing data buffer.
'
' @param    message The message to send to the other council members.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCouncilMessage(ByVal Message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CouncilMessage" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.CouncilMessage)
        
        Call .WriteASCIIString(Message)
    End With
End Sub

''
' Writes the "RoleMasterRequest" message to the outgoing data buffer.
'
' @param    message The message to send to the role masters.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRoleMasterRequest(ByVal Message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RoleMasterRequest" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.RoleMasterRequest)
        
        Call .WriteASCIIString(Message)
    End With
End Sub

''
' Writes the "GMRequest" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGMRequest(ByVal TIPO As Byte, ByVal Message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GMRequest" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMRequest)
        
        Call .WriteByte(TIPO)
        Call .WriteASCIIString(Message)
    End With
End Sub

''
' Writes the "ChangeDescription" message to the outgoing data buffer.
'
' @param    desc The new description of the user's character.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeDescription(ByVal desc As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChangeDescription" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.ChangeDescription)
        
        Call .WriteASCIIString(desc)
    End With
End Sub

''
' Writes the "GuildVote" message to the outgoing data buffer.
'
' @param    username The user to vote for clan leader.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildVote(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildVote" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildVote)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "Punishments" message to the outgoing data buffer.
'
' @param    username The user whose's  punishments are requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePunishments(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Punishments" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Punishments)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "Gamble" message to the outgoing data buffer.
'
' @param    amount The amount to gamble.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGamble(ByVal Amount As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Gamble" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Gamble)
        
        Call .WriteInteger(Amount)
    End With
End Sub

''
' Writes the "LeaveFaction" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLeaveFaction()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "LeaveFaction" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.LeaveFaction)
End Sub

''
' Writes the "BankExtractGold" message to the outgoing data buffer.
'
' @param    amount The amount of money to extract from the bank.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankExtractGold(ByVal Amount As Long)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BankExtractGold" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.BankExtractGold)
        
        Call .WriteLong(Amount)
    End With
End Sub

''
' Writes the "BankDepositGold" message to the outgoing data buffer.
'
' @param    amount The amount of money to deposit in the bank.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankDepositGold(ByVal Amount As Long)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BankDepositGold" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.BankDepositGold)
        
        Call .WriteLong(Amount)
    End With
End Sub

''
' Writes the "Denounce" message to the outgoing data buffer.
'
' @param    message The message to send with the denounce.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDenounce(ByVal Message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Denounce" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Denounce)
        
        Call .WriteASCIIString(Message)
    End With
End Sub

''
' Writes the "GuildFundate" message to the outgoing data buffer.
'
' @param    clanType The alignment of the clan to be founded.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildFundate(ByVal clanType As eClanType)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildFundate" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildFundate)
        
        Call .WriteByte(clanType)
    End With
End Sub

''
' Writes the "PartyKick" message to the outgoing data buffer.
'
' @param    username The user to kick fro mthe party.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePartyKick(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PartyKick" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.PartyKick)
            
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "PartySetLeader" message to the outgoing data buffer.
'
' @param    username The user to set as the party's leader.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePartySetLeader(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PartySetLeader" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.PartySetLeader)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "PartyAcceptMember" message to the outgoing data buffer.
'
' @param    username The user to accept into the party.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePartyAcceptMember(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PartyAcceptMember" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.PartyAcceptMember)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "GuildMemberList" message to the outgoing data buffer.
'
' @param    guild The guild whose member list is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildMemberList(ByVal guild As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildMemberList" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildMemberList)
        
        Call .WriteASCIIString(guild)
    End With
End Sub

''
' Writes the "GMMessage" message to the outgoing data buffer.
'
' @param    message The message to be sent to the other GMs online.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGMMessage(ByVal Message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GMMessage" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMMessage)
        
        Call .WriteASCIIString(Message)
    End With
End Sub

''
' Writes the "ShowName" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowName()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowName" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.showName)
End Sub

''
' Writes the "OnlineRoyalArmy" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteOnlineRoyalArmy()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "OnlineRoyalArmy" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.OnlineRoyalArmy)
End Sub

''
' Writes the "OnlineChaosLegion" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteOnlineChaosLegion()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "OnlineChaosLegion" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.OnlineChaosLegion)
End Sub

''
' Writes the "GoNearby" message to the outgoing data buffer.
'
' @param    username The suer to approach.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGoNearby(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GoNearby" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GoNearby)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "Comment" message to the outgoing data buffer.
'
' @param    message The message to leave in the log as a comment.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteComment(ByVal Message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Comment" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Comment)
        
        Call .WriteASCIIString(Message)
    End With
End Sub

''
' Writes the "ServerTime" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteServerTime()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ServerTime" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.serverTime)
End Sub

''
' Writes the "Where" message to the outgoing data buffer.
'
' @param    username The user whose position is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWhere(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Where" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Where)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "CreaturesInMap" message to the outgoing data buffer.
'
' @param    map The map in which to check for the existing creatures.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreaturesInMap(ByVal Map As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CreaturesInMap" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.CreaturesInMap)
        
        Call .WriteInteger(Map)
    End With
End Sub

''
' Writes the "WarpMeToTarget" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWarpMeToTarget()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "WarpMeToTarget" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.WarpMeToTarget)
End Sub

''
' Writes the "WarpChar" message to the outgoing data buffer.
'
' @param    username The user to be warped. "YO" represent's the user's char.
' @param    map The map to which to warp the character.
' @param    x The x position in the map to which to waro the character.
' @param    y The y position in the map to which to waro the character.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWarpChar(ByVal UserName As String, ByVal Map As Integer, ByVal X As Byte, ByVal Y As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "WarpChar" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.WarpChar)
        
        Call .WriteASCIIString(UserName)
        
        Call .WriteInteger(Map)
        
        Call .WriteByte(X)
        Call .WriteByte(Y)
    End With
End Sub

''
' Writes the "Silence" message to the outgoing data buffer.
'
' @param    username The user to silence.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSilence(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Silence" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Silence)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "SOSShowList" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSOSShowList()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SOSShowList" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.SOSShowList)
End Sub

''
' Writes the "SOSRemove" message to the outgoing data buffer.
'
' @param    username The user whose SOS call has been already attended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSOSRemove(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SOSRemove" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.SOSRemove)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "GoToChar" message to the outgoing data buffer.
'
' @param    username The user to be approached.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGoToChar(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GoToChar" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GoToChar)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "invisible" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteInvisible()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "invisible" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.invisible)
End Sub

''
' Writes the "GMPanel" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGMPanel()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GMPanel" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMPanel)
End Sub

''
' Writes the "RequestUserList" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestUserList()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestUserList" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.RequestUserList)
End Sub

''
' Writes the "Working" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWorking()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Working" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.Working)
End Sub

''
' Writes the "Hiding" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteHiding()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Hiding" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.Hiding)
End Sub

''
' Writes the "Jail" message to the outgoing data buffer.
'
' @param    username The user to be sent to jail.
' @param    reason The reason for which to send him to jail.
' @param    time The time (in minutes) the user will have to spend there.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteJail(ByVal UserName As String, ByVal reason As String, ByVal Time As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Jail" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Jail)
        
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(reason)
        
        Call .WriteByte(Time)
    End With
End Sub

''
' Writes the "KillNPC" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteKillNPC()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "KillNPC" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.KillNPC)
End Sub

''
' Writes the "WarnUser" message to the outgoing data buffer.
'
' @param    username The user to be warned.
' @param    reason Reason for the warning.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWarnUser(ByVal UserName As String, ByVal reason As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "WarnUser" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.WarnUser)
        
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(reason)
    End With
End Sub

''
' Writes the "EditChar" message to the outgoing data buffer.
'
' @param    UserName    The user to be edited.
' @param    editOption  Indicates what to edit in the char.
' @param    arg1        Additional argument 1. Contents depend on editoption.
' @param    arg2        Additional argument 2. Contents depend on editoption.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteEditChar(ByVal UserName As String, ByVal EditOption As eEditOptions, ByVal arg1 As String, ByVal arg2 As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "EditChar" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.EditChar)
        
        Call .WriteASCIIString(UserName)
        
        Call .WriteByte(EditOption)
        
        Call .WriteASCIIString(arg1)
        Call .WriteASCIIString(arg2)
    End With
End Sub

''
' Writes the "RequestCharInfo" message to the outgoing data buffer.
'
' @param    username The user whose information is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestCharInfo(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestCharInfo" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.RequestCharInfo)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "RequestCharStats" message to the outgoing data buffer.
'
' @param    username The user whose stats are requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestCharStats(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestCharStats" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.RequestCharStats)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "RequestCharGold" message to the outgoing data buffer.
'
' @param    username The user whose gold is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestCharGold(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestCharGold" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.RequestCharGold)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub
    
''
' Writes the "RequestCharInventory" message to the outgoing data buffer.
'
' @param    username The user whose inventory is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestCharInventory(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestCharInventory" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.RequestCharInventory)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "RequestCharBank" message to the outgoing data buffer.
'
' @param    username The user whose banking information is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestCharBank(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestCharBank" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.RequestCharBank)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "RequestCharSkills" message to the outgoing data buffer.
'
' @param    username The user whose skills are requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestCharSkills(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestCharSkills" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.RequestCharSkills)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "ReviveChar" message to the outgoing data buffer.
'
' @param    username The user to eb revived.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteReviveChar(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ReviveChar" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.ReviveChar)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "OnlineGM" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteOnlineGM()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "OnlineGM" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.OnlineGM)
End Sub

''
' Writes the "OnlineMap" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteOnlineMap(ByVal Map As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 26/03/2009
'Writes the "OnlineMap" message to the outgoing data buffer
'26/03/2009: Now you don't need to be in the map to use the comand, so you send the map to server
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.OnlineMap)
        
        Call .WriteInteger(Map)
    End With
End Sub

''
' Writes the "Forgive" message to the outgoing data buffer.
'
' @param    username The user to be forgiven.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteForgive(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Forgive" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Forgive)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "Kick" message to the outgoing data buffer.
'
' @param    username The user to be kicked.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteKick(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Kick" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Kick)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "Execute" message to the outgoing data buffer.
'
' @param    username The user to be executed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteExecute(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Execute" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Execute)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "BanChar" message to the outgoing data buffer.
'
' @param    username The user to be banned.
' @param    reason The reson for which the user is to be banned.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBanChar(ByVal UserName As String, ByVal reason As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BanChar" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.BanChar)
        
        Call .WriteASCIIString(UserName)
        
        Call .WriteASCIIString(reason)
    End With
End Sub

''
' Writes the "UnbanChar" message to the outgoing data buffer.
'
' @param    username The user to be unbanned.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUnbanChar(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UnbanChar" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.UnbanChar)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub
''
' Writes the "BanACC" message to the outgoing data buffer.
'
' @param    username The user to be banned.
' @param    reason The reson for which the user is to be banned.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBanACC(ByVal UserName As String, ByVal reason As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BanChar" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.BanACC)
        
        Call .WriteASCIIString(UserName)
        
        Call .WriteASCIIString(reason)
    End With
End Sub

''
' Writes the "UnbanACC" message to the outgoing data buffer.
'
' @param    username The user to be unbanned.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUnbanACC(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UnbanChar" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.UnbanACC)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "NPCFollow" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNPCFollow()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "NPCFollow" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.NPCFollow)
End Sub

''
' Writes the "SummonChar" message to the outgoing data buffer.
'
' @param    username The user to be summoned.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSummonChar(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SummonChar" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.SummonChar)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "SpawnListRequest" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSpawnListRequest()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SpawnListRequest" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.SpawnListRequest)
End Sub

''
' Writes the "SpawnCreature" message to the outgoing data buffer.
'
' @param    creatureIndex The index of the creature in the spawn list to be spawned.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSpawnCreature(ByVal creatureIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SpawnCreature" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.SpawnCreature)
        
        Call .WriteInteger(creatureIndex)
    End With
End Sub

''
' Writes the "ResetNPCInventory" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteResetNPCInventory()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ResetNPCInventory" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.ResetNPCInventory)
End Sub

''
' Writes the "CleanWorld" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCleanWorld()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CleanWorld" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.CleanWorld)
End Sub
Public Sub WriteSearchObj(ByVal restrict As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.SearchObj)
       
        Call .WriteASCIIString(restrict)
    End With
End Sub

''
' Writes the "ServerMessage" message to the outgoing data buffer.
'
' @param    message The message to be sent to players.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteServerMessage(ByVal Message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ServerMessage" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.ServerMessage)
        
        Call .WriteASCIIString(Message)
    End With
End Sub

''
' Writes the "NickToIP" message to the outgoing data buffer.
'
' @param    username The user whose IP is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNickToIP(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "NickToIP" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.NickToIP)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "IPToNick" message to the outgoing data buffer.
'
' @param    IP The IP for which to search for players. Must be an array of 4 elements with the 4 components of the IP.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteIPToNick(ByRef Ip() As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "IPToNick" message to the outgoing data buffer
'***************************************************
    If UBound(Ip()) - LBound(Ip()) + 1 <> 4 Then Exit Sub   'Invalid IP
    
    Dim i As Long
    
    With outgoingData
        Call .WriteByte(ClientPacketID.IPToNick)
        
        For i = LBound(Ip()) To UBound(Ip())
            Call .WriteByte(Ip(i))
        Next i
    End With
End Sub

''
' Writes the "GuildOnlineMembers" message to the outgoing data buffer.
'
' @param    guild The guild whose online player list is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildOnlineMembers(ByVal guild As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildOnlineMembers" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildOnlineMembers)
        
        Call .WriteASCIIString(guild)
    End With
End Sub

''
' Writes the "TeleportCreate" message to the outgoing data buffer.
'
' @param    map the map to which the teleport will lead.
' @param    x The position in the x axis to which the teleport will lead.
' @param    y The position in the y axis to which the teleport will lead.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTeleportCreate(ByVal Map As Integer, ByVal X As Byte, ByVal Y As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "TeleportCreate" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.TeleportCreate)
        
        Call .WriteInteger(Map)
        
        Call .WriteByte(X)
        Call .WriteByte(Y)
    End With
End Sub

''
' Writes the "TeleportDestroy" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTeleportDestroy()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "TeleportDestroy" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.TeleportDestroy)
End Sub

''
' Writes the "SetCharDescription" message to the outgoing data buffer.
'
' @param    desc The description to set to players.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSetCharDescription(ByVal desc As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SetCharDescription" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.SetCharDescription)
        
        Call .WriteASCIIString(desc)
    End With
End Sub

''
' Writes the "ForceMIDIToMap" message to the outgoing data buffer.
'
' @param    midiID The ID of the midi file to play.
' @param    map The map in which to play the given midi.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteForceMIDIToMap(ByVal midiID As Byte, ByVal Map As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ForceMIDIToMap" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.ForceMIDIToMap)
        
        Call .WriteByte(midiID)
        
        Call .WriteInteger(Map)
    End With
End Sub

''
' Writes the "ForceWAVEToMap" message to the outgoing data buffer.
'
' @param    waveID  The ID of the wave file to play.
' @param    Map     The map into which to play the given wave.
' @param    x       The position in the x axis in which to play the given wave.
' @param    y       The position in the y axis in which to play the given wave.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteForceWAVEToMap(ByVal waveID As Byte, ByVal Map As Integer, ByVal X As Byte, ByVal Y As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ForceWAVEToMap" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.ForceWAVEToMap)
        
        Call .WriteByte(waveID)
        
        Call .WriteInteger(Map)
        
        Call .WriteByte(X)
        Call .WriteByte(Y)
    End With
End Sub

''
' Writes the "RoyalArmyMessage" message to the outgoing data buffer.
'
' @param    message The message to send to the royal army members.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRoyalArmyMessage(ByVal Message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RoyalArmyMessage" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.RoyalArmyMessage)
        
        Call .WriteASCIIString(Message)
    End With
End Sub

''
' Writes the "ChaosLegionMessage" message to the outgoing data buffer.
'
' @param    message The message to send to the chaos legion member.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChaosLegionMessage(ByVal Message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChaosLegionMessage" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.ChaosLegionMessage)
        
        Call .WriteASCIIString(Message)
    End With
End Sub

''
' Writes the "CitizenMessage" message to the outgoing data buffer.
'
' @param    message The message to send to citizens.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCitizenMessage(ByVal Message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CitizenMessage" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.CitizenMessage)
        
        Call .WriteASCIIString(Message)
    End With
End Sub

''
' Writes the "CriminalMessage" message to the outgoing data buffer.
'
' @param    message The message to send to criminals.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCriminalMessage(ByVal Message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CriminalMessage" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.CriminalMessage)
        
        Call .WriteASCIIString(Message)
    End With
End Sub

''
' Writes the "TalkAsNPC" message to the outgoing data buffer.
'
' @param    message The message to send to the royal army members.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTalkAsNPC(ByVal Message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "TalkAsNPC" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.TalkAsNPC)
        
        Call .WriteASCIIString(Message)
    End With
End Sub

''
' Writes the "DestroyAllItemsInArea" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDestroyAllItemsInArea()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "DestroyAllItemsInArea" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.DestroyAllItemsInArea)
End Sub

''
' Writes the "AcceptRoyalCouncilMember" message to the outgoing data buffer.
'
' @param    username The name of the user to be accepted into the royal army council.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAcceptRoyalCouncilMember(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "AcceptRoyalCouncilMember" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.AcceptRoyalCouncilMember)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "AcceptChaosCouncilMember" message to the outgoing data buffer.
'
' @param    username The name of the user to be accepted as a chaos council member.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAcceptChaosCouncilMember(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "AcceptChaosCouncilMember" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.AcceptChaosCouncilMember)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "ItemsInTheFloor" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteItemsInTheFloor()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ItemsInTheFloor" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.ItemsInTheFloor)
End Sub

''
' Writes the "MakeDumb" message to the outgoing data buffer.
'
' @param    username The name of the user to be made dumb.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMakeDumb(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "MakeDumb" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.MakeDumb)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "MakeDumbNoMore" message to the outgoing data buffer.
'
' @param    username The name of the user who will no longer be dumb.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMakeDumbNoMore(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "MakeDumbNoMore" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.MakeDumbNoMore)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "DumpIPTables" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDumpIPTables()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "DumpIPTables" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.DumpIPTables)
End Sub

''
' Writes the "CouncilKick" message to the outgoing data buffer.
'
' @param    username The name of the user to be kicked from the council.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCouncilKick(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CouncilKick" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.CouncilKick)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "SetTrigger" message to the outgoing data buffer.
'
' @param    trigger The type of trigger to be set to the tile.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSetTrigger(ByVal Trigger As eTrigger)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SetTrigger" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.SetTrigger)
        
        Call .WriteByte(Trigger)
    End With
End Sub

''
' Writes the "AskTrigger" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAskTrigger()
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 04/13/07
'Writes the "AskTrigger" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.AskTrigger)
End Sub

''
' Writes the "BannedIPList" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBannedIPList()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BannedIPList" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.BannedIPList)
End Sub

''
' Writes the "BannedIPReload" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBannedIPReload()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BannedIPReload" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.BannedIPReload)
End Sub

''
' Writes the "GuildBan" message to the outgoing data buffer.
'
' @param    guild The guild whose members will be banned.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildBan(ByVal guild As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildBan" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildBan)
        
        Call .WriteASCIIString(guild)
    End With
End Sub

''
' Writes the "BanIP" message to the outgoing data buffer.
'
' @param    byIp    If set to true, we are banning by IP, otherwise the ip of a given character.
' @param    IP      The IP for which to search for players. Must be an array of 4 elements with the 4 components of the IP.
' @param    nick    The nick of the player whose ip will be banned.
' @param    reason  The reason for the ban.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBanIP(ByVal byIp As Boolean, ByRef Ip() As Byte, ByVal Nick As String, ByVal reason As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BanIP" message to the outgoing data buffer
'***************************************************
    If byIp And UBound(Ip()) - LBound(Ip()) + 1 <> 4 Then Exit Sub   'Invalid IP
    
    Dim i As Long
    
    With outgoingData
        Call .WriteByte(ClientPacketID.BanIP)
        
        Call .WriteBoolean(byIp)
        
        If byIp Then
            For i = LBound(Ip()) To UBound(Ip())
                Call .WriteByte(Ip(i))
            Next i
        Else
            Call .WriteASCIIString(Nick)
        End If
        
        Call .WriteASCIIString(reason)
    End With
End Sub

''
' Writes the "UnbanIP" message to the outgoing data buffer.
'
' @param    IP The IP for which to search for players. Must be an array of 4 elements with the 4 components of the IP.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUnbanIP(ByRef Ip() As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UnbanIP" message to the outgoing data buffer
'***************************************************
    If UBound(Ip()) - LBound(Ip()) + 1 <> 4 Then Exit Sub   'Invalid IP
    
    Dim i As Long
    
    With outgoingData
        Call .WriteByte(ClientPacketID.UnbanIP)
        
        For i = LBound(Ip()) To UBound(Ip())
            Call .WriteByte(Ip(i))
        Next i
    End With
End Sub

''
' Writes the "CreateItem" message to the outgoing data buffer.
'
' @param    itemIndex The index of the item to be created.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreateItem(ByVal itemIndex As Long)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CreateItem" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.CreateItem)
        
        Call .WriteInteger(itemIndex)
    End With
End Sub

''
' Writes the "DestroyItems" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDestroyItems()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "DestroyItems" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.DestroyItems)
    
End Sub

''
' Writes the "ChaosLegionKick" message to the outgoing data buffer.
'
' @param    username The name of the user to be kicked from the Chaos Legion.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChaosLegionKick(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChaosLegionKick" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.ChaosLegionKick)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "RoyalArmyKick" message to the outgoing data buffer.
'
' @param    username The name of the user to be kicked from the Royal Army.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRoyalArmyKick(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RoyalArmyKick" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.RoyalArmyKick)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "ForceMIDIAll" message to the outgoing data buffer.
'
' @param    midiID The id of the midi file to play.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteForceMIDIAll(ByVal midiID As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ForceMIDIAll" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.ForceMIDIAll)
        
        Call .WriteByte(midiID)
    End With
End Sub

''
' Writes the "ForceWAVEAll" message to the outgoing data buffer.
'
' @param    waveID The id of the wave file to play.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteForceWAVEAll(ByVal waveID As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ForceWAVEAll" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.ForceWAVEAll)
        
        Call .WriteByte(waveID)
    End With
End Sub

''
' Writes the "RemovePunishment" message to the outgoing data buffer.
'
' @param    username The user whose punishments will be altered.
' @param    punishment The id of the punishment to be removed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRemovePunishment(ByVal UserName As String, ByVal punishment As Byte, ByVal NewText As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RemovePunishment" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.RemovePunishment)
        
        Call .WriteASCIIString(UserName)
        Call .WriteByte(punishment)
        Call .WriteASCIIString(NewText)
    End With
End Sub

''
' Writes the "TileBlockedToggle" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTileBlockedToggle()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "TileBlockedToggle" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.TileBlockedToggle)
End Sub

''
' Writes the "KillNPCNoRespawn" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteKillNPCNoRespawn()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "KillNPCNoRespawn" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.KillNPCNoRespawn)
End Sub

''
' Writes the "KillAllNearbyNPCs" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteKillAllNearbyNPCs()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "KillAllNearbyNPCs" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.KillAllNearbyNPCs)
End Sub

''
' Writes the "LastIP" message to the outgoing data buffer.
'
' @param    username The user whose last IPs are requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLastIP(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "LastIP" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.LastIP)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "ChangeMOTD" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMOTD()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChangeMOTD" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.ChangeMOTD)
End Sub

''
' Writes the "SetMOTD" message to the outgoing data buffer.
'
' @param    message The message to be set as the new MOTD.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSetMOTD(ByVal Message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SetMOTD" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.SetMOTD)
        
        Call .WriteASCIIString(Message)
    End With
End Sub

''
' Writes the "SystemMessage" message to the outgoing data buffer.
'
' @param    message The message to be sent to all players.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSystemMessage(ByVal Message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SystemMessage" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.SystemMessage)
        
        Call .WriteASCIIString(Message)
    End With
End Sub

''
' Writes the "CreateNPC" message to the outgoing data buffer.
'
' @param    npcIndex The index of the NPC to be created.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreateNPC(ByVal NPCIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CreateNPC" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.CreateNPC)
        
        Call .WriteInteger(NPCIndex)
    End With
End Sub

''
' Writes the "CreateNPCWithRespawn" message to the outgoing data buffer.
'
' @param    npcIndex The index of the NPC to be created.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreateNPCWithRespawn(ByVal NPCIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CreateNPCWithRespawn" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.CreateNPCWithRespawn)
        
        Call .WriteInteger(NPCIndex)
    End With
End Sub

''
' Writes the "ImperialArmour" message to the outgoing data buffer.
'
' @param    armourIndex The index of imperial armour to be altered.
' @param    objectIndex The index of the new object to be set as the imperial armour.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteImperialArmour(ByVal armourIndex As Byte, ByVal objectIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ImperialArmour" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.ImperialArmour)
        
        Call .WriteByte(armourIndex)
        
        Call .WriteInteger(objectIndex)
    End With
End Sub

''
' Writes the "ChaosArmour" message to the outgoing data buffer.
'
' @param    armourIndex The index of chaos armour to be altered.
' @param    objectIndex The index of the new object to be set as the chaos armour.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChaosArmour(ByVal armourIndex As Byte, ByVal objectIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChaosArmour" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.ChaosArmour)
        
        Call .WriteByte(armourIndex)
        
        Call .WriteInteger(objectIndex)
    End With
End Sub

''
' Writes the "NavigateToggle" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNavigateToggle()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "NavigateToggle" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.NavigateToggle)
End Sub

''
' Writes the "ServerOpenToUsersToggle" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteServerOpenToUsersToggle()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ServerOpenToUsersToggle" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.ServerOpenToUsersToggle)
End Sub

''
' Writes the "TurnOffServer" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTurnOffServer()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "TurnOffServer" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.TurnOffServer)
End Sub

''
' Writes the "TurnCriminal" message to the outgoing data buffer.
'
' @param    username The name of the user to turn into criminal.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTurnCriminal(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "TurnCriminal" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.TurnCriminal)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "ResetFactions" message to the outgoing data buffer.
'
' @param    username The name of the user who will be removed from any faction.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteResetFactions(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ResetFactions" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.ResetFactions)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "RemoveCharFromGuild" message to the outgoing data buffer.
'
' @param    username The name of the user who will be removed from any guild.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRemoveCharFromGuild(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RemoveCharFromGuild" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.RemoveCharFromGuild)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "RequestCharMail" message to the outgoing data buffer.
'
' @param    username The name of the user whose mail is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestCharMail(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestCharMail" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.RequestCharMail)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "AlterPassword" message to the outgoing data buffer.
'
' @param    username The name of the user whose mail is requested.
' @param    copyFrom The name of the user from which to copy the password.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAlterPassword(ByVal UserName As String, ByVal CopyFrom As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "AlterPassword" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.AlterPassword)
        
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(CopyFrom)
    End With
End Sub

''
' Writes the "AlterMail" message to the outgoing data buffer.
'
' @param    username The name of the user whose mail is requested.
' @param    newMail The new email of the player.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAlterMail(ByVal UserName As String, ByVal newMail As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "AlterMail" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.AlterMail)
        
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(newMail)
    End With
End Sub

''
' Writes the "AlterName" message to the outgoing data buffer.
'
' @param    username The name of the user whose mail is requested.
' @param    newName The new user name.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAlterName(ByVal UserName As String, ByVal newName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "AlterName" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.AlterName)
        
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(newName)
    End With
End Sub

''
' Writes the "ToggleCentinelActivated" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteToggleCentinelActivated()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ToggleCentinelActivated" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.ToggleCentinelActivated)
End Sub

''
' Writes the "DoBackup" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDoBackup()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "DoBackup" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.DoBackUp)
End Sub

''
' Writes the "ShowGuildMessages" message to the outgoing data buffer.
'
' @param    guild The guild to listen to.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowGuildMessages(ByVal guild As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowGuildMessages" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.ShowGuildMessages)
        
        Call .WriteASCIIString(guild)
    End With
End Sub

''
' Writes the "SaveMap" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSaveMap()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SaveMap" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.SaveMap)
End Sub

''
' Writes the "ChangeMapInfoPK" message to the outgoing data buffer.
'
' @param    isPK True if the map is PK, False otherwise.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoPK(ByVal isPK As Boolean)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChangeMapInfoPK" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.ChangeMapInfoPK)
        
        Call .WriteBoolean(isPK)
    End With
End Sub

''
' Writes the "ChangeMapInfoBackup" message to the outgoing data buffer.
'
' @param    backup True if the map is to be backuped, False otherwise.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoBackup(ByVal backup As Boolean)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChangeMapInfoBackup" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.ChangeMapInfoBackup)
        
        Call .WriteBoolean(backup)
    End With
End Sub

''
' Writes the "ChangeMapInfoRestricted" message to the outgoing data buffer.
'
' @param    restrict NEWBIES (only newbies), NO (everyone), ARMADA (just Armadas), CAOS (just caos) or FACCION (Armadas & caos only)
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoRestricted(ByVal restrict As String)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'Writes the "ChangeMapInfoRestricted" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.ChangeMapInfoRestricted)
        
        Call .WriteASCIIString(restrict)
    End With
End Sub

''
' Writes the "ChangeMapInfoNoMagic" message to the outgoing data buffer.
'
' @param    nomagic TRUE if no magic is to be allowed in the map.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoNoMagic(ByVal nomagic As Boolean)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'Writes the "ChangeMapInfoNoMagic" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.ChangeMapInfoNoMagic)
        
        Call .WriteBoolean(nomagic)
    End With
End Sub

''
' Writes the "ChangeMapInfoNoInvi" message to the outgoing data buffer.
'
' @param    noinvi TRUE if invisibility is not to be allowed in the map.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoNoInvi(ByVal noinvi As Boolean)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'Writes the "ChangeMapInfoNoInvi" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.ChangeMapInfoNoInvi)
        
        Call .WriteBoolean(noinvi)
    End With
End Sub
                            
''
' Writes the "ChangeMapInfoNoResu" message to the outgoing data buffer.
'
' @param    noresu TRUE if resurection is not to be allowed in the map.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoNoResu(ByVal noresu As Boolean)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'Writes the "ChangeMapInfoNoResu" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.ChangeMapInfoNoResu)
        
        Call .WriteBoolean(noresu)
    End With
End Sub
                        
''
' Writes the "ChangeMapInfoLand" message to the outgoing data buffer.
'
' @param    land options: "BOSQUE", "NIEVE", "DESIERTO", "CIUDAD", "CAMPO", "DUNGEON".
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoLand(ByVal land As String)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'Writes the "ChangeMapInfoLand" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.ChangeMapInfoLand)
        
        Call .WriteASCIIString(land)
    End With
End Sub
                        
''
' Writes the "ChangeMapInfoZone" message to the outgoing data buffer.
'
' @param    zone options: "BOSQUE", "NIEVE", "DESIERTO", "CIUDAD", "CAMPO", "DUNGEON".
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoZone(ByVal zone As String)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'Writes the "ChangeMapInfoZone" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.ChangeMapInfoZone)
        
        Call .WriteASCIIString(zone)
    End With
End Sub

''
' Writes the "SaveChars" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSaveChars()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SaveChars" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.SaveChars)
End Sub

''
' Writes the "CleanSOS" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCleanSOS()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CleanSOS" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.CleanSOS)
End Sub

''
' Writes the "ShowServerForm" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowServerForm()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowServerForm" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.ShowServerForm)
End Sub

''
' Writes the "KickAllChars" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteKickAllChars()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "KickAllChars" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.KickAllChars)
End Sub

''
' Writes the "ReloadNPCs" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteReloadNPCs()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ReloadNPCs" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.ReloadNPCs)
End Sub

''
' Writes the "ReloadServerIni" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteReloadServerIni()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ReloadServerIni" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.ReloadServerIni)
End Sub

''
' Writes the "ReloadSpells" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteReloadSpells()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ReloadSpells" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.ReloadSpells)
End Sub

''
' Writes the "ReloadObjects" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteReloadObjects()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ReloadObjects" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.ReloadObjects)
End Sub

''
' Writes the "Restart" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRestart()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Restart" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.Restart)
End Sub

''
' Writes the "ChatColor" message to the outgoing data buffer.
'
' @param    r The red component of the new chat color.
' @param    g The green component of the new chat color.
' @param    b The blue component of the new chat color.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChatColor(ByVal r As Byte, ByVal g As Byte, ByVal B As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChatColor" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.ChatColor)
        
        Call .WriteByte(r)
        Call .WriteByte(g)
        Call .WriteByte(B)
    End With
End Sub

''
' Writes the "Ignored" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteIgnored()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Ignored" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.Ignored)
End Sub

''
' Writes the "CheckSlot" message to the outgoing data buffer.
'
' @param    UserName    The name of the char whose slot will be checked.
' @param    slot        The slot to be checked.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCheckSlot(ByVal UserName As String, ByVal slot As Byte)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'Writes the "CheckSlot" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.CheckSlot)
        Call .WriteASCIIString(UserName)
        Call .WriteByte(slot)
    End With
End Sub

''
' Writes the "Ping" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePing()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 26/01/2007
'Writes the "Ping" message to the outgoing data buffer
'***************************************************
    'Prevent the timer from being cut
    If pingTime <> 0 Then Exit Sub
    
    Call outgoingData.WriteByte(ClientPacketID.Ping)
    
    ' Avoid computing errors due to frame rate
    Call FlushBuffer
    DoEvents
    
    pingTime = GetTickCount
End Sub

''
' Writes the "Rank" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRank()
'***************************************************
'Author: Manuel (Lorwik)
'Last Modification: 29/11/2011
'***************************************************
    
    Call outgoingData.WriteByte(ClientPacketID.Rank)
    
    ' Avoid computing errors due to frame rate
    Call FlushBuffer
End Sub

''
' Writes the "TopQuince" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTopQuince()
'***************************************************
'Author: Manuel (Lorwik)
'Last Modification: 29/11/2011
'***************************************************
    
    Call outgoingData.WriteByte(ClientPacketID.TopQuince)
    
    ' Avoid computing errors due to frame rate
    Call FlushBuffer
End Sub

''
' Writes the "SetIniVar" message to the outgoing data buffer.
'
' @param    sLlave the name of the key which contains the value to edit
' @param    sClave the name of the value to edit
' @param    sValor the new value to set to sClave
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSetIniVar(ByRef sLlave As String, ByRef sClave As String, ByRef sValor As String)
'***************************************************
'Author: Brian Chaia (BrianPr)
'Last Modification: 21/06/2009
'Writes the "SetIniVar" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.SetIniVar)
        
        Call .WriteASCIIString(sLlave)
        Call .WriteASCIIString(sClave)
        Call .WriteASCIIString(sValor)
    End With
End Sub

''
' Flushes the outgoing data buffer of the user.
'
' @param    UserIndex User whose outgoing data buffer will be flushed.

Public Sub FlushBuffer()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Sends all data existing in the buffer
'***************************************************
    Dim sndData As String
    
    With outgoingData
        If .Length = 0 Then _
            Exit Sub
        
        sndData = .ReadASCIIStringFixed(.Length)
        
        Call SendData(sndData)
    End With
End Sub

''
' Sends the data using the socket controls in the MainForm.
'
' @param    sdData  The data to be sent to the server.

Private Sub SendData(ByRef sdData As String)
    
    'No enviamos nada si no estamos conectados
    If frmMain.Winsock1.State <> sckConnected Then Exit Sub
    
    'Send data!
    Call frmMain.Winsock1.SendData(sdData)

End Sub
Public Sub HandleUpdateArmor()
    
    If incomingData.Length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserArmourMaxDef = incomingData.ReadInteger
    UserArmourMinDef = incomingData.ReadInteger
    
    frmMain.lblArmor.Caption = UserArmourMaxDef & "/" & UserArmourMinDef
End Sub
Public Sub HandleUpdateEscu()
    If incomingData.Length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserEscuMaxDef = incomingData.ReadInteger
    UserEscuMinDef = incomingData.ReadInteger
    
    frmMain.LblShield.Caption = UserEscuMaxDef & "/" & UserEscuMinDef
End Sub
Public Sub HandleUpdateCasco()
    If incomingData.Length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserHelmetMaxDef = incomingData.ReadInteger
    UserHelmetMinDef = incomingData.ReadInteger
    
    frmMain.lblHelm.Caption = UserHelmetMaxDef & "/" & UserHelmetMinDef
End Sub
Public Sub HandleUpdateHit()
    If incomingData.Length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserMaxHit = incomingData.ReadInteger
    UserMinHit = incomingData.ReadInteger
    
    frmMain.lblWeapon.Caption = UserMaxHit & "/" & UserMinHit
End Sub
Private Sub HandleCuentas()
Dim Command As String
Dim i As Byte
Dim j As Byte
Dim Estado As Byte
Dim Total As Byte

    'Remove PacketID
    Call incomingData.ReadByte
    Command = incomingData.ReadASCIIString
   
    Select Case Command
        Case "LoginOK"
            Estado = incomingData.ReadByte
           
            If Estado = 1 Then
                Unload frmConnect
                Unload frmCrearAccount
            End If
           
            With Cuenta
                .CantPJ = incomingData.ReadByte
                 If Not .CantPJ = 0 Then
                    .pjs(.CantPJ).NamePJ = incomingData.ReadASCIIString
                    .pjs(.CantPJ).ClasePJ = incomingData.ReadByte
                    .pjs(.CantPJ).LvlPJ = incomingData.ReadByte
                    .pjs(.CantPJ).promedio = incomingData.ReadLong
                    .pjs(.CantPJ).Acuerpo = incomingData.ReadInteger
                    .pjs(.CantPJ).rcvHead = incomingData.ReadInteger
                    .pjs(.CantPJ).rcvCasco = incomingData.ReadByte
                    .pjs(.CantPJ).rcvShield = incomingData.ReadByte
                    .pjs(.CantPJ).rcvWeapon = incomingData.ReadByte
                    .pjs(.CantPJ).rcvBaned = incomingData.ReadByte
                    .pjs(.CantPJ).rcvMuerto = incomingData.ReadByte
                    .pjs(.CantPJ).rcvRaza = incomingData.ReadByte
                    .pjs(.CantPJ).PJLogged = incomingData.ReadByte
                
                    If .pjs(.CantPJ).NamePJ = "" Then
                        frmCuenta.Nombre(.CantPJ).Caption = "Nada"
                        frmCuenta.Label1(.CantPJ).Caption = "-"
                        frmCuenta.Label2(.CantPJ).Caption = "-"
                    Else
                        Call DibujarTodo(.CantPJ, .pjs(.CantPJ).Acuerpo, .pjs(.CantPJ).rcvHead, .pjs(.CantPJ).rcvCasco, .pjs(.CantPJ).rcvShield, .pjs(.CantPJ).rcvWeapon, .pjs(.CantPJ).rcvBaned, .pjs(.CantPJ).NamePJ, .pjs(.CantPJ).LvlPJ, ListaClases(.pjs(.CantPJ).ClasePJ), .pjs(.CantPJ).rcvMuerto, ListaRazas(.pjs(.CantPJ).rcvRaza), .pjs(.CantPJ).PJLogged)
                    End If
                End If
                
                Total = incomingData.ReadByte

                If Not .CantPJ < Total Then
                    If Estado = 1 Then frmCuenta.Show
                End If
            End With
            
            Exit Sub
    End Select
End Sub

Public Sub WriteLoginCuenta()
    With outgoingData
        Call .WriteByte(ClientPacketID.CuentaCL)
        Call .WriteASCIIString("LoginAcc")
        Call .WriteASCIIString(Cuenta.Name)
        Call .WriteASCIIString(Cuenta.Pass)
    End With
End Sub
Public Sub WriteCrearCuenta()
    With outgoingData
        Call .WriteByte(ClientPacketID.CuentaCL)
        Call .WriteASCIIString("NewAcc")
        Call .WriteASCIIString(Cuenta.Name)
        Call .WriteASCIIString(Cuenta.Pass)
        Call .WriteASCIIString(Cuenta.Email)
        Call .WriteASCIIString(Cuenta.preg)
        Call .WriteASCIIString(Cuenta.resp)
    End With
    
    Debug.Print "MANDO EL CREAR"
End Sub
Public Sub WriteBorrPJCuenta()
    With outgoingData
        Call .WriteByte(ClientPacketID.CuentaCL)
        Call .WriteASCIIString("DltPJAcc")
        Call .WriteASCIIString(Cuenta.Name)
        Call .WriteASCIIString(Cuenta.Pass)
        Call .WriteByte(PJClickeado)
    End With
End Sub
Public Sub WriteCambiarPass()
With outgoingData
        Call .WriteByte(ClientPacketID.CuentaCL)
        Call .WriteASCIIString("Cmbdpss")
        Call .WriteASCIIString(Cuenta.Name)
        Call .WriteASCIIString(PsswdAnte)
        Call .WriteASCIIString(PasswdNew)
End With
End Sub
Public Sub writeRecuperarAccount()
With outgoingData
        Call .WriteByte(ClientPacketID.CuentaCL)
        Call .WriteASCIIString("RcprAcc")
        Call .WriteASCIIString(NameAccount)
        Call .WriteASCIIString(NWPasswd)
        Call .WriteASCIIString(PRGScrta)
        Call .WriteASCIIString(Repstscrta)
End With
End Sub

Public Sub WriteQuest()
'Escribe el paquete Quest al servidor.
    Call outgoingData.WriteByte(ClientPacketID.WinterClientPacket)
    Call outgoingData.WriteByte(WinterClientPacketID.Quest)
End Sub

Public Sub WriteQuestDetailsRequest(ByVal QuestSlot As Byte)
'Escribe el paquete QuestDetailsRequest al servidor.
    Call outgoingData.WriteByte(ClientPacketID.WinterClientPacket)
    Call outgoingData.WriteByte(WinterClientPacketID.QuestDetailsRequest)
    
    Call outgoingData.WriteByte(QuestSlot)
End Sub

Public Sub WriteQuestAccept()
'Escribe el paquete QuestAccept al servidor.
    Call outgoingData.WriteByte(ClientPacketID.WinterClientPacket)
    Call outgoingData.WriteByte(WinterClientPacketID.QuestAccept)
End Sub

Private Sub HandleQuestDetails()
'Recibe y maneja el paquete QuestDetails del servidor.
    
On Error GoTo ErrHandler
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    Dim tmpStr As String
    Dim tmpByte As Byte
    Dim QuestEmpezada As Boolean
    Dim i As Integer
    
    With Buffer
        'Leemos el id del paquete
        Call .ReadByte
        
        'Nos fijamos si se trata de una quest empezada, para poder leer los NPCs que se han matado.
        QuestEmpezada = IIf(.ReadByte, True, False)
        
        tmpStr = "Misión: " & .ReadASCIIString & vbCrLf
        tmpStr = tmpStr & "Detalles: " & .ReadASCIIString & vbCrLf
        tmpStr = tmpStr & "Nivel requerido: " & .ReadByte & vbCrLf
        
        tmpStr = tmpStr & vbCrLf & "OBJETIVOS" & vbCrLf
        
        tmpByte = .ReadByte
        If tmpByte Then 'Hay NPCs
            For i = 1 To tmpByte
                tmpStr = tmpStr & "*) Matar " & .ReadInteger & " " & .ReadASCIIString & "."
                If QuestEmpezada Then
                    tmpStr = tmpStr & " (Has matado " & .ReadInteger & ")" & vbCrLf
                Else
                    tmpStr = tmpStr & vbCrLf
                End If
            Next i
        End If
        
        tmpByte = .ReadByte
        If tmpByte Then 'Hay OBJs
            For i = 1 To tmpByte
                tmpStr = tmpStr & "*) Conseguir " & .ReadInteger & " " & .ReadASCIIString & "." & vbCrLf
            Next i
        End If

        tmpStr = tmpStr & vbCrLf & "RECOMPENSAS" & vbCrLf
        tmpStr = tmpStr & "*) Oro: " & .ReadLong & " monedas de oro." & vbCrLf
        tmpStr = tmpStr & "*) Experiencia: " & .ReadLong & " puntos de experiencia." & vbCrLf
        
        tmpByte = .ReadByte
        If tmpByte Then
            For i = 1 To tmpByte
                tmpStr = tmpStr & "*) " & .ReadInteger & " " & .ReadASCIIString & vbCrLf
            Next i
        End If
    End With
    
    'Determinamos que formulario se muestra, según si recibimos la información y la quest está empezada o no.
    If QuestEmpezada Then
        frmQuests.txtInfo.Caption = tmpStr
    Else
        frmQuestInfo.txtInfo.Caption = tmpStr
        frmQuestInfo.Show vbModeless, frmMain
    End If
    
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

Public Sub HandleQuestListSend()
'Recibe y maneja el paquete QuestListSend del servidor.
    If incomingData.Length < 1 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    Dim i As Integer
    Dim tmpByte As Byte
    Dim tmpStr As String
    
    'Leemos el id del paquete
    Call Buffer.ReadByte
     
    'Leemos la cantidad de quests que tiene el usuario
    tmpByte = Buffer.ReadByte
    
    'Limpiamos el ListBox y el TextBox del formulario
    frmQuests.lstQuests.Clear
    frmQuests.txtInfo.Caption = vbNullString
        
    'Si el usuario tiene quests entonces hacemos el handle
    If tmpByte Then
        'Leemos el string
        tmpStr = Buffer.ReadASCIIString
        
        'Agregamos los items
        For i = 1 To tmpByte
            frmQuests.lstQuests.AddItem ReadField(i, tmpStr, 45)
            frmQuests.QuestNum.Caption = tmpByte & "/15"
        Next i
    End If
    
    'Mostramos el formulario
    frmQuests.Show vbModeless, frmMain
    
    'Pedimos la información de la primer quest (si la hay)
    If tmpByte Then Call Protocol.WriteQuestDetailsRequest(1)
    
    'Copiamos de vuelta el buffer
    Call incomingData.CopyBuffer(Buffer)

ErrHandler:
    Dim error As Long
    error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

Public Sub WriteQuestListRequest()
'Escribe el paquete QuestListRequest al servidor.
    Call outgoingData.WriteByte(ClientPacketID.WinterClientPacket)
    Call outgoingData.WriteByte(WinterClientPacketID.QuestListRequest)
End Sub

Public Sub WriteQuestAbandon(ByVal QuestSlot As Byte)
'Escribe el paquete QuestAbandon al servidor.
    'Escribe el ID del paquete.
    Call outgoingData.WriteByte(ClientPacketID.WinterClientPacket)
    Call outgoingData.WriteByte(WinterClientPacketID.QuestAbandon)
    
    'Escribe el Slot de Quest.
    Call outgoingData.WriteByte(QuestSlot)
End Sub
Public Sub WriteNoche(ByVal Clima As Byte)
'******Lorwik/Noche**********
'*****www.lwk-foros.net******
With outgoingData
Call .WriteByte(ClientPacketID.WinterClientPacket)
Call .WriteByte(WinterClientPacketID.Anochecer)
Call .WriteByte(Clima)
End With
End Sub
Private Sub handleNoche()
'******Lorwik/Noche**********
'*****www.lwk-foros.net******
    'Remove packet ID
    Call incomingData.ReadByte
   
    Anocheceria = incomingData.ReadByte()
    
    If Anocheceria = 0 Then
        frmMain.Clima.Picture = General_Load_Picture_From_Resource("75.gif")
    ElseIf Anocheceria = 1 Then
        frmMain.Clima.Picture = General_Load_Picture_From_Resource("74.gif")
    ElseIf Anocheceria = 2 Then
        frmMain.Clima.Picture = General_Load_Picture_From_Resource("77.gif")
    ElseIf Anocheceria = 3 Then
        frmMain.Clima.Picture = General_Load_Picture_From_Resource("76.gif")
    End If
End Sub

Public Sub WriteSwapObjects(ByVal ObjSlot1 As Byte, ByVal ObjSlot2 As Byte)
    With outgoingData
        Call .WriteByte(ClientPacketID.WinterClientPacket)
        Call .WriteByte(WinterClientPacketID.SwapObjects)
   
        Call .WriteByte(ObjSlot1)
        Call .WriteByte(ObjSlot2)
    End With
End Sub

Public Sub WriteMsgGlobal(ByVal Message As String)
With outgoingData
    Call .WriteByte(ClientPacketID.MsgGlobal)
    Call .WriteASCIIString(Message)
End With
End Sub
Public Sub WriteGlobalOff()
With outgoingData
    Call .WriteByte(ClientPacketID.GlobalOff)
End With
End Sub
Public Sub WriteGlobalOn()
With outgoingData
    Call .WriteByte(ClientPacketID.GlobalOn)
End With
End Sub

Public Sub WriteCanjes(ByVal Canje As Byte)
    Call outgoingData.WriteByte(ClientPacketID.WinterClientPacket)
    Call outgoingData.WriteByte(WinterClientPacketID.canjes)
    Call outgoingData.WriteByte(Canje)
End Sub
Public Sub WritePuntos(ByVal UserName As String, ByVal Puntos As Long)
With outgoingData
        Call .WriteByte(ClientPacketID.WinterClientPacket)
        Call .WriteByte(WinterClientPacketID.Puntos)
        Call .WriteASCIIString(UserName)
        Call .WriteLong(Puntos)
End With
End Sub

''
' Handles the CommerceInit message.

Private Sub HandleCanjesInit()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    Dim i As Long
    
    'Remove packet ID
    Call incomingData.ReadByte
    frmCanjes.Puntos.Caption = incomingData.ReadLong
    
    frmCanjes.Show , frmMain
    DoEvents
    
    'Set state and show form
    Comerciando = True
    
    'Reproducimos el saludo del comerciante
    Call General_Set_Wav("comerciante" & RandomNumber(1, 9) & ".wav")
End Sub
Private Sub HandleTimeInvi()
    
    'Remove packet ID
    Call incomingData.ReadByte
    CartelInvisibilidad = incomingData.ReadByte
End Sub
Private Sub HandleTimeParal()
    
    'Remove packet ID
    Call incomingData.ReadByte
    CartelParalisis = incomingData.ReadByte
End Sub

Public Sub WriteDueloAutomatico(Rondas As Integer)
    Call outgoingData.WriteByte(ClientPacketID.WinterClientPacket)
    Call outgoingData.WriteByte(WinterClientPacketID.DueloAutomatico)
    Call outgoingData.WriteInteger(Rondas)
 End Sub
Public Sub WriteCancelarDueloAutomatico()
    Call outgoingData.WriteByte(ClientPacketID.WinterClientPacket)
    Call outgoingData.WriteByte(WinterClientPacketID.CancelarDueloAutomatico)
End Sub
Public Sub WriteParticiparDueloAutomatico()
    Call outgoingData.WriteByte(ClientPacketID.WinterClientPacket)
    Call outgoingData.WriteByte(WinterClientPacketID.ParticiparDueloAutomatico)
End Sub

Private Sub HandleShowPartyForm()
'***************************************************
'Author: Budi
'Last Modification: 11/26/09
'
'***************************************************
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim members() As String
    Dim i As Long
    
    EsPartyLeader = CBool(Buffer.ReadByte())
       
    members = Split(Buffer.ReadASCIIString(), SEPARATOR)
    For i = 0 To UBound(members())
        Call frmParty.lstMembers.AddItem(members(i))
    Next i
    
    frmParty.lblTotalExp.Caption = Buffer.ReadLong
    frmParty.Show , frmMain
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

Public Sub WriteRequestPartyForm()
'***************************************************
'Author: Budi
'Last Modification: 11/26/09
'Writes the "RequestPartyForm" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.WinterClientPacket)
    Call outgoingData.WriteByte(WinterClientPacketID.RequestPartyForm)

End Sub

''
' Writes the "Resucitate" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteInvocar()
'***************************************************
'Author: Lorwik
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.WinterClientPacket)
    Call outgoingData.WriteByte(WinterClientPacketID.Invocar)
End Sub

Public Sub WriteDesDenuncia(ByVal Activador As Byte)
'***************************************************
'Author: Lorwik
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.WinterClientPacket)
    Call outgoingData.WriteByte(WinterClientPacketID.DesDenuncia)
    Call outgoingData.WriteByte(Activador)
End Sub

Public Sub WriteHacerTorneo(ByVal MinLevel As Byte, ByVal Maxlevel As Byte, ByVal Cupos As Byte, ByVal AutoSum As Byte, Optional ByVal Mapa As Integer, Optional X As Byte, Optional Y As Byte)
    With outgoingData
        Call .WriteByte(ClientPacketID.WinterClientPacket)
        Call .WriteByte(WinterClientPacketID.HacerTorneo)
        Call .WriteByte(MinLevel)
        Call .WriteByte(Maxlevel)
        Call .WriteByte(Cupos)
        Call .WriteByte(AutoSum)
        
        If AutoSum = 1 Then
            Call .WriteInteger(Mapa)
            Call .WriteByte(X)
            Call .WriteByte(Y)
        End If
    End With
End Sub
 
Public Sub Writetorneo()
With outgoingData
    Call .WriteByte(ClientPacketID.WinterClientPacket)
    Call .WriteByte(WinterClientPacketID.Torneo)
    End With
End Sub
 
Public Sub WriteCerrartorneo()
With outgoingData
    Call .WriteByte(ClientPacketID.WinterClientPacket)
    Call .WriteByte(WinterClientPacketID.Cerrartorneo)
End With
End Sub

Public Sub WriteTransferencia(ByVal Name As String, ByVal cantidad As Long)
'******************************
'Autor: Manuel (Lorwik)
'******************************
With outgoingData
    Call .WriteByte(ClientPacketID.WinterClientPacket)
    Call .WriteByte(WinterClientPacketID.Transferencia)
    Call .WriteASCIIString(Name)
    Call .WriteLong(cantidad)
End With
End Sub
'**********************************************************************************
'******************************MODULO TCP******************************************
'**********************************************************************************
Sub Login()
    If EstadoLogin = E_MODO.Normal Then
        Call WriteLoginExistingChar
    ElseIf EstadoLogin = E_MODO.CrearNuevoPj Then
        Call WriteLoginNewChar
    ElseIf EstadoLogin = E_MODO.LoginCuenta Then
        Call WriteLoginCuenta
    ElseIf EstadoLogin = E_MODO.CrearCuenta Then
        Call WriteCrearCuenta
    ElseIf EstadoLogin = E_MODO.BorrarPJ Then
        Call WriteBorrPJCuenta
    ElseIf EstadoLogin = E_MODO.CambiarPass Then
        Call WriteCambiarPass
    ElseIf EstadoLogin = E_MODO.RecuperarAccount Then
        Call writeRecuperarAccount
    End If
    DoEvents
    Call FlushBuffer
End Sub
'**********************************************************************************
'**********************************************************************************

