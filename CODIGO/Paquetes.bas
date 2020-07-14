Attribute VB_Name = "Paquetes"
Option Explicit

Public Enum ClientPacketIDGuild
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
    ShowGuildNews
End Enum

Public Enum ClientPacketIDGM
    GMMessage               '/GMSG
    showName                '/SHOWNAME
    OnlineRoyalArmy         '/ONLINEREAL
    onlineChaosLegion       '/ONLINECAOS
    GoNearby                '/IRCERCA
    comment                 '/REM
    serverTime              '/HORA
    Where                   '/DONDE
    CreaturesInMap          '/NENE
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
    RainToggle              '/LLUVIA
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
    AskTrigger              '/TRIGGER with no args
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
    SaveChars               '/GRABAR
    CleanSOS                '/BORRAR SOS
    ShowServerForm          '/SHOW INT
    KickAllChars            '/ECHARTODOSPJS
    ReloadNPCs              '/RELOADNPCS
    ReloadServerIni         '/RELOADSINI
    ReloadSpells            '/RELOADHECHIZOS
    ReloadObjects           '/RELOADOBJ
    Restart                 '/REINICIAR
    ResetAutoUpdate         '/AUTOUPDATE
    ChatColor               '/CHATCOLOR
    Ignored                 '/IGNORADO
    CheckSlot               '/SLOT
    SetIniVar               '/SETINIVAR LLAVE CLAVE VALOR
    AddGM
End Enum

Public Enum ServerPacketID
    OpenAccount
    logged                  ' LOGGED
    ChangeHour
    RemoveDialogs           ' QTDL
    RemoveCharDialog        ' QDL
    NavigateToggle          ' NAVEG
    Disconnect              ' FINOK
    CommerceEnd             ' FINCOMOK
    BankEnd                 ' FINBANOK
    CommerceInit            ' INITCOM
    BankInit                ' INITBANCO
    UserCommerceInit        ' INITCOMUSU
    UserCommerceEnd         ' FINCOMUSUOK
    UserOfferConfirm
    CancelOfferItem
    ShowBlacksmithForm      ' SFH
    ShowCarpenterForm       ' SFC
    NPCSwing                ' N1
    NPCKillUser             ' 6
    BlockedWithShieldUser   ' 7
    BlockedWithShieldOther  ' 8
    UserSwing               ' U1
    SafeModeOn              ' SEGON
    SafeModeOff             ' SEGOFF
    ResuscitationSafeOn
    ResuscitationSafeOff
    NobilityLost            ' PN
    CantUseWhileMeditating  ' M!
    Ataca                   ' USER ATACA
    UpdateSta               ' ASS
    UpdateMana              ' ASM
    UpdateHP                ' ASH
    UpdateGold              ' ASG
    UpdateBankGold
    UpdateExp               ' ASE
    ChangeMap               ' CM
    PosUpdate               ' PU
    NPCHitUser              ' N2
    UserHitNPC              ' U2
    UserAttackedSwing       ' U3
    UserHittedByUser        ' N4
    UserHittedUser          ' N5
    ChatOverHead            ' ||
    ConsoleMsg              ' || - Beware!! its the same as above, but it was properly splitted
    GuildChat               ' |+
    ShowMessageBox          ' !!
    ShowMessageScroll
    UserIndexInServer       ' IU
    UserCharIndexInServer   ' IP
    CharacterCreate         ' CC
    CharacterRemove         ' BP
    CharacterMove           ' MP, +, * and _ '
    ForceCharMove
    CharacterChange         ' CP
    ObjectCreate            ' HO
    ObjectDelete            ' BO
    BlockPosition           ' BQ
    PlayWave                ' TW
    guildList               ' GL
    AreaChanged             ' CA
    PauseToggle             ' BKW
    RainToggle              ' LLU
    CreateFX                ' CFX
    CreateEfecto
    UpdateUserStats         ' EST
    WorkRequestTarget       ' T01
    ChangeInventorySlot     ' CSI
    ChangeBankSlot          ' SBO
    ChangeSpellSlot         ' SHS
    atributes               ' ATR
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
    GuildMemberInfo
    GuildDetails            ' CLANDET
    ShowGuildFundationForm  ' SHOWFUN
    ParalizeOK              ' PARADOK
    ShowUserRequest         ' PETICIO
    TradeOK                 ' TRANSOK
    BankOK                  ' BANCOOK
    ChangeUserTradeSlot     ' COMUSUINV
    Pong
    UpdateTagAndStatus
    UsersOnline
    CommerceChat
    ShowPartyForm
    StopWorking
    RetosAbre
    RetosRespuesta
    
    'GM messages
    SpawnList               ' SPL
    ShowSOSForm             ' MSOS
    ShowMOTDEditionForm     ' ZMOTD
    ShowGMPanelForm         ' ABPANEL
    UserNameList            ' LISTUSU
    ShowBarco
    AgregarPasajero
    QuitarPasajero
    QuitarBarco
    GoHome
    GotHome
    Tooltip
End Enum

Public Enum ClientPacketID
    OpenAccount             'ALOGIN
    LoginExistingChar       'OLOGIN
    LoginNewChar            'NLOGIN
    Talk                    ';
    Yell                    '-
    Whisper                 '\
    Walk                    'M
    RequestPositionUpdate   'RPU
    Attack                  'AT
    PickUp                  'AG
    SafeToggle              '/SEG & SEG  (SEG's behaviour has to be coded in the client)
    ResuscitationSafeToggle
    RequestGuildLeaderInfo  'GLINFO
    RequestAtributes        'ATR
    RequestFame             'FAMA
    RequestSkills           'ESKI
    RequestMiniStats        'FEST
    CommerceEnd             'FINCOM
    UserCommerceEnd         'FINCOMUSU
    UserCommerceConfirm
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
    MoveSpell               'DESPHE
    MoveBank
    ClanCodexUpdate         'DESCOD
    UserCommerceOffer       'OFRECER
    CommerceChat
    guild
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
    Help                    '/AYUDA
    RequestStats            '/EST
    CommerceStart           '/COMERCIAR
    BankStart               '/BOVEDA
    ComandosVarios                  '/ENLISTAR
    Information             '/INFORMACION
    Reward                  '/RECOMPENSA
    RequestMOTD             '/MOTD
    UpTime                  '/UPTIME
    PartyLeave              '/SALIRPARTY
    PartyCreate             '/CREARPARTY
    PartyJoin               '/PARTY
    Inquiry                 '/ENCUESTA ( with no params )
    GuildMessage            '/CMSG
    PartyMessage            '/PMSG
    CentinelReport          '/CENTINELA
    GuildOnline             '/ONLINECLAN
    PartyOnline             '/ONLINEPARTY
    CouncilMessage          '/BMSG
    RoleMasterRequest       '/ROL
    GMRequest               '/GM
    bugReport               '/_BUG
    ChangeDescription       '/DESC
    GuildVote               '/VOTO
    Punishments             '/PENAS
    ChangePassword          '/CONTRASEÑA
    Gamble                  '/APOSTAR
    InquiryVote             '/ENCUESTA ( with parameters )
    LeaveFaction            '/RETIRAR ( with no arguments )
    BankExtractGold         '/RETIRAR ( with arguments )
    BankDepositGold         '/DEPOSITAR
    Denounce                '/DENUNCIAR
    GuildFundate            '/FUNDARCLAN
    PartyKick               '/ECHARPARTY
    PartySetLeader          '/PARTYLIDER
    PartyAcceptMember       '/ACCEPTPARTY
    Ping                    '/PING
    RequestPartyForm
    ItemUpgrade
    InitCrafting
    RetosAbrir              '/Retos
    RetosCrear
    RetosDecide
    IntercambiarInv
    
    'GM messages
    gm
    WarpMeToTarget          '/TELEPLOC
    Home
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
    FONTTYPE_RETOS
    FONTTYPE_EXP
End Enum

Public PaqueteName(0 To 255) As String

Public Sub InitDebug()
PaqueteName(0) = "OpenAccount"
PaqueteName(1) = "logged"
PaqueteName(2) = "ChangeHour"
PaqueteName(3) = "RemoveDialogs"
PaqueteName(4) = "RemoveCharDialog"
PaqueteName(5) = "NavigateToggle"
PaqueteName(6) = "Disconnect"
PaqueteName(7) = "CommerceEnd"
PaqueteName(8) = "BankEnd"
PaqueteName(9) = "CommerceInit"
PaqueteName(10) = "BankInit"
PaqueteName(11) = "UserCommerceInit"
PaqueteName(12) = "UserCommerceEnd"
PaqueteName(13) = "UserOfferConfirm"
PaqueteName(14) = "CancelOfferItem"
PaqueteName(15) = "ShowBlacksmithForm"
PaqueteName(16) = "ShowCarpenterForm"
PaqueteName(17) = "NPCSwing"
PaqueteName(18) = "NPCKillUser"
PaqueteName(19) = "BlockedWithShieldUser"
PaqueteName(20) = "BlockedWithShieldOther"
PaqueteName(21) = "UserSwing"
PaqueteName(22) = "SafeModeOn"
PaqueteName(23) = "SafeModeOff"
PaqueteName(24) = "ResuscitationSafeOn"
PaqueteName(25) = "ResuscitationSafeOff"
PaqueteName(26) = "NobilityLost"
PaqueteName(27) = "CantUseWhileMeditating"
PaqueteName(28) = "Ataca"
PaqueteName(29) = "UpdateSta"
PaqueteName(30) = "UpdateMana"
PaqueteName(31) = "UpdateHP"
PaqueteName(32) = "UpdateGold"
PaqueteName(33) = "UpdateBankGold"
PaqueteName(34) = "UpdateExp"
PaqueteName(35) = "ChangeMap"
PaqueteName(36) = "PosUpdate"
PaqueteName(37) = "NPCHitUser"
PaqueteName(38) = "UserHitNPC"
PaqueteName(39) = "UserAttackedSwing"
PaqueteName(40) = "UserHittedByUser"
PaqueteName(41) = "UserHittedUser"
PaqueteName(42) = "ChatOverHead"
PaqueteName(43) = "ConsoleMsg"
PaqueteName(44) = "GuildChat"
PaqueteName(45) = "ShowMessageBox"
PaqueteName(46) = "ShowMessageScroll"
PaqueteName(47) = "UserIndexInServer"
PaqueteName(48) = "UserCharIndexInServer"
PaqueteName(49) = "CharacterCreate"
PaqueteName(50) = "CharacterRemove"
PaqueteName(51) = "CharacterMove"
PaqueteName(52) = "ForceCharMove"
PaqueteName(53) = "CharacterChange"
PaqueteName(54) = "ObjectCreate"
PaqueteName(55) = "ObjectDelete"
PaqueteName(56) = "BlockPosition"
PaqueteName(57) = "PlayWave"
PaqueteName(58) = "guildList"
PaqueteName(59) = "AreaChanged"
PaqueteName(60) = "PauseToggle"
PaqueteName(61) = "RainToggle"
PaqueteName(62) = "CreateFX"
PaqueteName(63) = "CreateEfecto"
PaqueteName(64) = "UpdateUserStats"
PaqueteName(65) = "WorkRequestTarget"
PaqueteName(66) = "ChangeInventorySlot"
PaqueteName(67) = "ChangeBankSlot"
PaqueteName(68) = "ChangeSpellSlot"
PaqueteName(69) = "atributes"
PaqueteName(70) = "BlacksmithWeapons"
PaqueteName(71) = "BlacksmithArmors"
PaqueteName(72) = "CarpenterObjects"
PaqueteName(73) = "RestOK"
PaqueteName(74) = "ErrorMsg"
PaqueteName(75) = "Blind"
PaqueteName(76) = "Dumb"
PaqueteName(77) = "ShowSignal"
PaqueteName(78) = "ChangeNPCInventorySlot"
PaqueteName(79) = "UpdateHungerAndThirst"
PaqueteName(80) = "Fame"
PaqueteName(81) = "MiniStats"
PaqueteName(82) = "LevelUp"
PaqueteName(83) = "SetInvisible"
PaqueteName(84) = "MeditateToggle"
PaqueteName(85) = "BlindNoMore"
PaqueteName(86) = "DumbNoMore"
PaqueteName(87) = "SendSkills"
PaqueteName(88) = "TrainerCreatureList"
PaqueteName(89) = "guildNews"
PaqueteName(90) = "OfferDetails"
PaqueteName(91) = "AlianceProposalsList"
PaqueteName(92) = "PeaceProposalsList"
PaqueteName(93) = "CharacterInfo"
PaqueteName(94) = "GuildLeaderInfo"
PaqueteName(95) = "GuildMemberInfo"
PaqueteName(96) = "GuildDetails"
PaqueteName(97) = "ShowGuildFundationForm"
PaqueteName(98) = "ParalizeOK"
PaqueteName(99) = "ShowUserRequest"
PaqueteName(100) = "TradeOK"
PaqueteName(101) = "BankOK"
PaqueteName(102) = "ChangeUserTradeSlot"
PaqueteName(103) = "Pong"
PaqueteName(104) = "UpdateTagAndStatus"
PaqueteName(105) = "UsersOnline"
PaqueteName(106) = "CommerceChat"
PaqueteName(107) = "ShowPartyForm"
PaqueteName(108) = "StopWorking"
PaqueteName(109) = "RetosAbre"
PaqueteName(110) = "RetosRespuesta"
PaqueteName(111) = "SpawnList"
PaqueteName(112) = "ShowSOSForm"
PaqueteName(113) = "ShowMOTDEditionForm"
PaqueteName(114) = "ShowGMPanelForm"
PaqueteName(115) = "UserNameList"
PaqueteName(116) = "ShowBarco"
PaqueteName(117) = "AgregarPasajero"
PaqueteName(118) = "QuitarPasajero"
PaqueteName(119) = "QuitarBarco"
PaqueteName(120) = "GoHome"
PaqueteName(121) = "GotHome"
PaqueteName(122) = "Tooltip"
End Sub
