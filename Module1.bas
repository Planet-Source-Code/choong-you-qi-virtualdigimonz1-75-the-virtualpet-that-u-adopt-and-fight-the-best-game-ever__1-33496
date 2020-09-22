Attribute VB_Name = "Module1"
Declare Function ShellExecute _
   Lib "shell32.dll" Alias "ShellExecuteA" _
   (ByVal hWnd As Long, _
   ByVal lpOperation As String, _
   ByVal lpFile As String, _
   ByVal lpParameters As String, _
   ByVal lpDirectory As String, _
   ByVal nShowCmd As Long) As Long
Public Const SW_SHOWNORMAL = 1

Public PlayerNameTournament
Public TempNameTournament
Public OpponentIP
Public OpponentName
Public YourPOST
Public TEMPstrdata
Public TEMPdata

Public Form31Show As Integer
Public Form37Show As Integer
Public Form27Show As Integer
Public FormBattleShow As Integer

Public AnimHealStart As Integer
Public AnimHealEnd As Integer

Public WinLoseCode As String
Public m_cN As cNeoCaption
Public m_cN2 As cNeoCaption
Public m_cN3 As cNeoCaption
Public m_cN4 As cNeoCaption
Public TournamentBotSkill As Long
Public TournamentBotName As String
Public TournamentAvailable As String
Public GraphUpdate As Integer
Public Top8OR1 As Integer
Public Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long

Public Function GenerateRandom(minVal As Long, maxVal As Long) As Long
    intr = -1
    maxVal = maxVal + 1
    If maxVal > 0 Then
        If minVal >= maxVal Then
            minVal = 0
        End If
    Else
        minVal = 0
        maxVal = 10
    End If
    Randomize (DatePart("s", Now) + DatePart("m", Now))
    Do While (intr < minVal Or intr = maxVal)
        intr = CLng(Rnd() * maxVal)
    Loop
    GenerateRandom = intr
End Function
Public Function CheckDatabaseActive(f As Form)
f.Hide
Form16.Show
Form16.Caption = "Checking Database Activation... Please Wait A Moment..."
MousePointer = vbHourglass
Checkdatabase = GetUrlSource("http://www2.domaindlx.com/choongyouqi/check.asp")
Do Until Checkdatabase = ""

If left(Checkdatabase, 18) = "ActivationComplete" Then
CheckDatabaseActive = "Activation"
MousePointer = vbDefault
f.Show
Unload Form16
Exit Function
End If

Checkdatabase = Mid(Checkdatabase, 2)
Loop
Unload Form16
f.Show
MsgBox "Server is currently not activated. Please try again later." & vbCrLf & "If you continuing getting this message, Please contact the author."
CheckDatabaseActive = "NotActivation"
MousePointer = vbDefault
End Function
Public Function Reset_Program()
DeleteSetting "Digimon"
End Function
Public Function CheckRegister()
If Not GetSetting("Digimon", "Application", "Register") = "1" Then
MsgBox "This Function only works after you register."
CheckRegister = "ExitSub"
Exit Function
End If
End Function
Public Function Food_Event_Caption()
Select Case Val(Form2.Label5.Tag)
Case Is > 20
Form2.Label5.Caption = "Full: Ultra Full"
Case 18 To 20
Form2.Label5.Caption = "Full: Ultra Full"
Case 15 To 17
Form2.Label5.Caption = "Full: Very Full"
Case 12 To 14
Form2.Label5.Caption = "Full: Full"
Case 10 To 11
Form2.Label5.Caption = "Full: Normal"
Case 7 To 9
Form2.Label5.Caption = "Full: Hungry"
Case 5 To 6
Form2.Label5.Caption = "Full: Very Hungry"
Case 3 To 4
Form2.Label5.Caption = "Full: Ultra Hungry"
Case 0 To 2
MsgBox "WARNING! YOUR DIGIMON NEAR DIE!"
Form2.Label5.Caption = "Full: Near Die!"
End Select
Form2.Label5.ToolTipText = Form2.Label5.Tag & "/" & GetSetting("Digimon", "Profile", "MaxFood")
End Function
Public Function ManageIP(WillHow As String, FirstIPvalue)
IP1value = GetSetting("Digimon", "Setting", "IP1")
IP2value = GetSetting("Digimon", "Setting", "IP2")
IP3value = GetSetting("Digimon", "Setting", "IP3")
IP4value = GetSetting("Digimon", "Setting", "IP4")
Select Case WillHow
Case "GET"
Form5.Combo1.Clear
If Not IP1value = "" Then Form5.Combo1.AddItem IP1value
If Not IP2value = "" Then Form5.Combo1.AddItem IP2value
If Not IP3value = "" Then Form5.Combo1.AddItem IP3value
If Not IP4value = "" Then Form5.Combo1.AddItem IP4value
Case "SAVE"
SaveSetting "Digimon", "Setting", "IP1", FirstIPvalue
SaveSetting "Digimon", "Setting", "IP2", IP1value
SaveSetting "Digimon", "Setting", "IP3", IP2value
SaveSetting "Digimon", "Setting", "IP4", IP3value
End Select
End Function
Public Function Form2_Update()
DigimonType
Form2.Label1.Caption = left(Form2.Label1.Caption, 8) & GetSetting("Digimon", "Digimon", "CurrentHealth") & "/" & GetSetting("Digimon", "Digimon", "health")
Form2.Label3.Caption = left(Form2.Label3.Caption, 7) & GetSetting("Digimon", "Digimon", "Power")
Form2.Label4.Caption = left(Form2.Label4.Caption, 9) & GetSetting("Digimon", "Digimon", "defence")
End Function
Public Function Animation_Heal(HealStart)
AnimHealStart = 0
AnimHealStart = HealStart
End Function
Public Function LotteryTicketList()
If Form37Show = "1" Then
Form37.List1.Clear
If Not GetSetting("Digimon", "Lottery", "Lottery1") = "" Then Form37.List1.AddItem GetSetting("Digimon", "Lottery", "Lottery1")
If Not GetSetting("Digimon", "Lottery", "Lottery2") = "" Then Form37.List1.AddItem GetSetting("Digimon", "Lottery", "Lottery2")
If Not GetSetting("Digimon", "Lottery", "Lottery3") = "" Then Form37.List1.AddItem GetSetting("Digimon", "Lottery", "Lottery3")
If Not GetSetting("Digimon", "Lottery", "Lottery4") = "" Then Form37.List1.AddItem GetSetting("Digimon", "Lottery", "Lottery4")
If Not GetSetting("Digimon", "Lottery", "Lottery5") = "" Then Form37.List1.AddItem GetSetting("Digimon", "Lottery", "Lottery5")
End If
End Function


Public Function Personality()
If GetSetting("Digimon", "Profile", "Hour") = "59" Then
Plushourtimer = 0
WhichMsgbox = ""
Select Case GetSetting("Digimon", "Profile", "Picture")
Case "1"
'increase attack 35
'decrease health 5
'decrease defence 10
WillAttack "PLUS", 35
WillHealth "MINUS", 5
WillDefence "MINUS", 10
Case "2"
'increase health 25
'decrease defence 2
'decrease attack 2
WillHealth "PLUS", 25
WillDefence "MINUS", 2
WillAttack "MINUS", 2
Case "3"
'increase defence 30
'decrease attack 9
'decrease health 1
WillDefence "PLUS", 30
WillAttack "MINUS", 9
WillHealth "MINUS", 1
Case "4"
'increase money 5,000
'decrease exp. point 5
WillMoney "PLUS", 5000
WillPoint "MINUS", 5
Case "5"
'increase money 10,000
'decrease health 10
'decrease attack 10
'decrease defence 10
WillMoney "PLUS", 10000
WillHealth "MINUS", 10
WillAttack "MINUS", 10
WillDefence "MINUS", 10
Case "6"
'maxfood + 5
'increase money 100
WillMaxFood "PLUS", 5
WillMoney "PLUS", 100
Case "7"
'chance to win double money
SaveSetting "Digimon", "Profile", "CasinoDouble", "1"
SaveSetting "Digimon", "Profile", "DoubleType", "money"
Form2.CasinoDouble.Enabled = True
WhichMsgbox = "DoubleCasino"
Case "8"
'chance to win double exp. point
SaveSetting "Digimon", "Profile", "CasinoDouble", "1"
SaveSetting "Digimon", "Profile", "DoubleType", "point"
Form2.CasinoDouble.Enabled = True
WhichMsgbox = "DoubleCasino"
End Select
'code'
Select Case WhichMsgbox
Case ""
If Not FormBattleShow = "1" Then MsgBox "Personality." & vbCrLf & "Ops! Something happen to your pet!", vbInformation
Case "DoubleCasino"
If Not FormBattleShow = "1" Then MsgBox "Personality." & vbCrLf & "The 'CasinoDouble!' open...", vbInformation
End Select
Form2_Update
SaveSetting "Digimon", "Profile", "Hour", Plushourtimer
Else
Plushourtimer = GetSetting("Digimon", "Profile", "Hour")
Plushourtimer = Val(Plushourtimer) + Val(1)
End If
SaveSetting "Digimon", "Profile", "Hour", Plushourtimer
End Function
Public Function MinusMoney(Money As Long)
moneyvalue = GetSetting("Digimon", "Profile", "Money")
moneyvalue = Val(moneyvalue) - Val(Money)
If left(moneyvalue, 1) = "-" Then moneyvalue = 0
SaveSetting "Digimon", "Profile", "Money", moneyvalue
End Function
Public Function AddEvent(whatEvent As String)
Form36.List2.AddItem whatEvent
Form36.List2.ListIndex = Form36.List2.ListCount - 1
End Function

Public Function ArrangeOpponentData(DataArrange As String)
TEMPdata = ""
YourPOST = left(DataArrange, 1)
TEMPstrdata = Mid(DataArrange, 2)

Do Until TEMPstrdata = ""
If left(TEMPstrdata, 1) = Chr(1) Then: OpponentIP = TEMPdata: TEMPdata = "": TEMPstrdata = Mid(TEMPstrdata, 2)
If left(TEMPstrdata, 1) = Chr(2) Then: OpponentName = TEMPdata: TEMPdata = "": TEMPstrdata = Mid(TEMPstrdata, 2)

TEMPdata = TEMPdata & left(TEMPstrdata, 1)
TEMPstrdata = Mid(TEMPstrdata, 2)
Loop

End Function
Public Function CheckVersusBot()
If left(OpponentName, 11) = "A.I. Skill(" Then
TournamentBotName = OpponentName
TournamentBotSkill = left(Mid(OpponentName, 12), Len(Mid(OpponentName, 12)) - Val(1))
AddEvent "You VS Bot Skill: " & TournamentBotSkill
TournamentAvailable = "1"
MsgBox "Tournament Begin!" & vbCrLf & "OpponentName:" & OpponentName & vbCrLf & "OpponentIP:" & OpponentIP
Form36.FlatButton1.Tag = "Form13"
Form36.FlatButton1.Visible = True
Form36.FlatButton1.Enabled = True
CheckVersusBot = "True"
End If
End Function

Public Function LotteryTimer()
Dim NewestLotteryNum As Integer
Dim A1LotteryNum
Dim A2LotteryNum
Dim A3LotteryNum
Dim A4LotteryNum
Dim A5LotteryNum
Dim BlankLottery

YOUWINLOTTERY = "NO"
BlankLottery = "0"

LotteryTime = GetSetting("Digimon", "Lottery", "LotteryTime")
LotteryTime = Val(LotteryTime) - Val(1)
If LotteryTime = "0" Then
LotteryTime = "10"
NewestLotteryNum = GenerateRandom(1, 200)
SaveSetting "Digimon", "Lottery", "LotteryTime", LotteryTime

A1LotteryNum = GetSetting("Digimon", "Lottery", "Lottery1")
A2LotteryNum = GetSetting("Digimon", "Lottery", "Lottery2")
A3LotteryNum = GetSetting("Digimon", "Lottery", "Lottery3")
A4LotteryNum = GetSetting("Digimon", "Lottery", "Lottery4")
A5LotteryNum = GetSetting("Digimon", "Lottery", "Lottery5")

If A1LotteryNum = "" Then If A2LotteryNum = "" Then If A3LotteryNum = "" Then If A4LotteryNum = "" Then If A5LotteryNum = "" Then BlankLottery = "1"

If A1LotteryNum = NewestLotteryNum Then YOUWINLOTTERY = "YES"
If A2LotteryNum = NewestLotteryNum Then YOUWINLOTTERY = "YES"
If A3LotteryNum = NewestLotteryNum Then YOUWINLOTTERY = "YES"
If A4LotteryNum = NewestLotteryNum Then YOUWINLOTTERY = "YES"
If A5LotteryNum = NewestLotteryNum Then YOUWINLOTTERY = "YES"
'MsgBox NewestLotteryNum & ": " & A1LotteryNum & ", " & A2LotteryNum & ", " & A3LotteryNum & ", " & A4LotteryNum & ", " & A5LotteryNum
'MsgBox YOUWINLOTTERY

Dim LastLotteryResultAndTicket As String
LastLotteryResultAndTicket = ""
If Not A1LotteryNum = "" Then LastLotteryResultAndTicket = LastLotteryResultAndTicket & A1LotteryNum
If Not A2LotteryNum = "" Then LastLotteryResultAndTicket = LastLotteryResultAndTicket & ", " & A2LotteryNum
If Not A3LotteryNum = "" Then LastLotteryResultAndTicket = LastLotteryResultAndTicket & ", " & A3LotteryNum
If Not A4LotteryNum = "" Then LastLotteryResultAndTicket = LastLotteryResultAndTicket & ", " & A4LotteryNum
If Not A5LotteryNum = "" Then LastLotteryResultAndTicket = LastLotteryResultAndTicket & ", " & A5LotteryNum
SaveSetting "Digimon", "Lottery", "LotteryLastNum", NewestLotteryNum & ": " & LastLotteryResultAndTicket


SaveSetting "Digimon", "Lottery", "Lottery1", ""
SaveSetting "Digimon", "Lottery", "Lottery2", ""
SaveSetting "Digimon", "Lottery", "Lottery3", ""
SaveSetting "Digimon", "Lottery", "Lottery4", ""
SaveSetting "Digimon", "Lottery", "Lottery5", ""

LotteryTicketList

If Form37Show = 1 Then
Form37.Label2.Caption = "Next OpenTime: " & GetSetting("Digimon", "Lottery", "LotteryTime") & "min"
Form37.Label1.Caption = GetSetting("Digimon", "Lottery", "LotteryLastNum")
End If
'MsgBox BlankLottery
If BlankLottery = "1" Then Exit Function

Select Case YOUWINLOTTERY
Case "YES"
WillBank "PLUS", 500000
If Form27Show = "1" Then Form27.Label1.Caption = "Bank cash: " & Label1.Tag
If Not FormBattleShow = "1" Then MsgBox "Lottery Result Out!" & vbCrLf & "YOU WIN THE LOTTERY." & vbCrLf & "$500,000 has been transfer to your bank."
Case "NO"
If Not FormBattleShow = "1" Then MsgBox "Lottery Result Out!" & vbCrLf & "Sorry, you never hit the lottery number."
End Select

Exit Function
End If


SaveSetting "Digimon", "Lottery", "LotteryTime", LotteryTime
If Form37Show = 1 Then
Form37.Label2.Caption = "Next OpenTime: " & GetSetting("Digimon", "Lottery", "LotteryTime") & "min"
Form37.Label1.Caption = GetSetting("Digimon", "Lottery", "LotteryLastNum")
End If

End Function
Public Function ShitTimer()
Select Case GetSetting("Digimon", "Digimon", "Shit")
Case "1"
If GenerateRandom(1, 2) = 2 Then
Select Case GenerateRandom(1, 3)
Case "1"
WillAttack "MINUS", GenerateRandom(1, 5)
Case "2"
WillDefence "MINUS", GenerateRandom(1, 5)
Case "3"
WillHealth "MINUS", GenerateRandom(1, 5)
End Select
MsgBox "Ops! Something happen to your pet!", vbInformation
End If
Case "0"
Select Case GenerateRandom(1, 10)
Case 5
SaveSetting "Digimon", "Digimon", "Shit", "1"
Form2.Image37.Visible = True
End Select
End Select
End Function
Public Function Form5_Event()
Select Case Form5.Tag
Case "battle"
Form5.Caption = "TCP/IP Battle"
Case "reset"
Form5.Caption = "TCP/IP Reset Together"
End Select
End Function
Public Function UnloadSubForm()
Unload Form1
Unload Form3
Unload Form4
Unload Form5
Unload Form6
Unload Form7
Unload Form8
Unload Form9
Unload Form10
Unload Form11
Unload Form12
Unload Form13
Unload Form14
Unload Form15
Unload Form16
Unload Form17
Unload Form18
Unload Form19
Unload Form20
Unload Form21
Unload Form22
Unload Form23
Unload Form24
Unload Form25
Unload Form26
Unload Form27
Unload Form28
Unload Form29
Unload Form30
Unload Form31
Unload Form32
Unload Form33
Unload Form34
Unload Form35
Unload Form36
Form2.Hide
End Function
Public Function ShareMarketValue()
ValuePrice1 = GetSetting("Digimon", "Share", "Price1")
ValuePrice2 = GetSetting("Digimon", "Share", "Price2")
ValuePrice3 = GetSetting("Digimon", "Share", "Price3")
plusminus1 = GenerateRandom(1, 3)
plusminus2 = GenerateRandom(1, 3)
plusminus3 = GenerateRandom(1, 3)
Raise1 = GenerateRandom(1, 20)
Raise2 = GenerateRandom(1, 30)
Raise3 = GenerateRandom(1, 25)
Drop1 = GenerateRandom(1, 20)
Drop2 = GenerateRandom(1, 30)
Drop3 = GenerateRandom(1, 25)

Select Case plusminus1
Case 1
ValuePrice1 = Val(ValuePrice1) + Val(Raise1)
Case 2
ValuePrice1 = Val(ValuePrice1) - Val(Drop1)
End Select

Select Case plusminus2
Case 1
ValuePrice2 = Val(ValuePrice2) + Val(Raise2)
Case 2
ValuePrice2 = Val(ValuePrice2) - Val(Drop2)
End Select

Select Case plusminus3
Case 1
ValuePrice3 = Val(ValuePrice3) + Val(Raise3)
Case 2
ValuePrice3 = Val(ValuePrice3) - Val(Drop3)
End Select

If Val(ValuePrice1) < 1 Then
SaveSetting "Digimon", "Share", "Own1", "0"
ValuePrice1 = GenerateRandom(10, 100)
If Not FormBattleShow = "1" Then MsgBox "Digi Drink Co. Ltd. Company has been Bankrup!"
End If
If Val(ValuePrice2) < 1 Then
SaveSetting "Digimon", "Share", "Own2", "0"
ValuePrice2 = GenerateRandom(100, 300)
If Not FormBattleShow = "1" Then MsgBox "Monster Equitment Corp. Company has been Bankrup!"
End If
If Val(ValuePrice3) < 1 Then
SaveSetting "Digimon", "Share", "Own3", "0"
ValuePrice3 = GenerateRandom(50, 200)
If Not FormBattleShow = "1" Then MsgBox "Truso Private Hospital Company has been Bankrup!"
End If

SaveSetting "Digimon", "Share", "Price1", ValuePrice1
SaveSetting "Digimon", "Share", "Price2", ValuePrice2
SaveSetting "Digimon", "Share", "Price3", ValuePrice3

If Form31Show = 1 Then Form31.Timer1.Enabled = True
End Function
Public Function ItemSub()
Item1value = GetSetting("Digimon", "Item", "1")
Item2value = GetSetting("Digimon", "Item", "2")
Item3value = GetSetting("Digimon", "Item", "3")
Item4value = GetSetting("Digimon", "Item", "4")
Item5value = GetSetting("Digimon", "Item", "5")
ItemAllSub = Val(Item1value) + Val(Item2value) + Val(Item3value) + Val(Item4value) + Val(Item5value)
SaveSetting "Digimon", "Item", "allsub", ItemAllSub
End Function

Public Function CheckMidiPlay()
Select Case GetSetting("Digimon", "Digimon", "Setting2")
Case 0
Form39.Command2.Value = True
Form39.Timer1.Enabled = False
Case 1
Form39.Command1.Value = True
Form39.Timer1.Enabled = False
Form39.Timer1.Interval = "0"
Form39.Timer1.Interval = "49000"
Form39.Timer1.Enabled = True
End Select
End Function

Public Function BankInterest()
Dim FinalInterest As Long
Dim LastTimeBank
Dim TimeNowBank
Dim BankMoney
Dim InterestTime
Dim BankLastInterest

LastTimeBank = GetSetting("Digimon", "Profile", "LastTimeBank")
TimeNowBank = GetSetting("Digimon", "Profile", "Played")
InterestTime = Val(TimeNowBank) - Val(LastTimeBank)
BankMoney = GetSetting("Digimon", "Profile", "Bank")
FinalInterest = Val(InterestTime) * Val(Val(1 / 100) * BankMoney)
If FinalInterest = "0" Then Exit Function
BankLastInterest = GetSetting("Digimon", "Profile", "BankInterest")
SaveSetting "Digimon", "Profile", "BankInterest", Val(BankLastInterest) + Val(FinalInterest)

FinalInterest = Val(FinalInterest) + Val(BankMoney)
SaveSetting "Digimon", "Profile", "Bank", FinalInterest
SaveSetting "Digimon", "Profile", "LastTimeBank", GetSetting("Digimon", "Profile", "Played")
If Form27Show = "1" Then Form27.Update.Enabled = True
End Function
Public Function Skin(f As Form, cN As cNeoCaption)
      cN.ActiveCaptionColor = &HFFFFFF
      cN.InActiveCaptionColor = &HCCCCCC
      cN.ActiveMenuColor = &HCCCCCC
      cN.ActiveMenuColorOver = &HFFFFFF
      cN.InActiveMenuColor = &H808080
      cN.MenuBackgroundColor = &H0&
      cN.CaptionFont = f.Font
      cN.MenuFont = f.Font
      cN.Attach f, Form7.Picture1.Picture, Form7.Picture2.Picture, 13, 14, 90, 142, 162, 162
      f.BackColor = &H0&
End Function
Public Function Skin2(f As Form, cN As cNeoCaption)
      cN.ActiveCaptionColor = &HFFFFFF
      cN.InActiveCaptionColor = &HCCCCCC
      cN.ActiveMenuColor = &HCCCCCC
      cN.ActiveMenuColorOver = &HFFFFFF
      cN.InActiveMenuColor = &H808080
      cN.MenuBackgroundColor = &H0&
      cN.CaptionFont = f.Font
      cN.MenuFont = f.Font
      cN.Attach f, Form7.Picture3.Picture, Form7.Picture2.Picture, 13, 14, 94, 150, 162, 0
      f.BackColor = &H0&
End Function
Public Function FileExist(asPath As String) As Boolean
    If UCase(Dir(asPath)) = UCase(TrimPath(asPath)) Then
      FileExist = True
    Else
      FileExist = False
    End If
End Function
Public Function TrimPath(ByVal asPath As String) As String
    If Len(asPath) = 0 Then Exit Function
    Dim X As Integer
    Do
        X = InStr(asPath, "\")
        If X = 0 Then Exit Do
        asPath = right(asPath, Len(asPath) - X)
    Loop
    TrimPath = asPath
End Function
Public Function MakeGraph(makegraphdata)
Select Case GraphUpdate
Case 1
Form36.Label3.Caption = makegraphdata
Case 2
Form36.Label4.Caption = makegraphdata
Case 3
Form36.Label5.Caption = makegraphdata
Case 4
Form36.Label6.Caption = makegraphdata
Case 5
Form36.Label7.Caption = makegraphdata
Case 6
Form36.Label8.Caption = makegraphdata
Case 7
Form36.Label9.Caption = makegraphdata
Case 8
Form36.Label10.Caption = makegraphdata
Case 9
Form36.Label11.Caption = makegraphdata
Case 10
Form36.Label12.Caption = makegraphdata
Case 11
Form36.Label13.Caption = makegraphdata
Case 12
Form36.Label14.Caption = makegraphdata
Case 13
Form36.Label15.Caption = makegraphdata
Case 14
Form36.Label16.Caption = makegraphdata
Case 15
Form36.Label17.Caption = makegraphdata
End Select
GraphUpdate = GraphUpdate + 1
End Function
Public Function AllMenu(WillHow As Boolean)
Form2.OwnerProfile.Enabled = WillHow
Form2.BattleStatistic.Enabled = WillHow
Form2.setting.Enabled = WillHow
Form2.Casino.Enabled = WillHow
Form2.WizardShop.Enabled = WillHow
Form2.ItemStore.Enabled = WillHow
Form2.bank.Enabled = WillHow
Form2.TraningCenter.Enabled = WillHow
Form2.NewsCenter.Enabled = WillHow
Form2.ShareMarket.Enabled = WillHow
Form2.TCPIP.Enabled = WillHow
Form2.LevelGainer.Enabled = WillHow
Form2.OnlineTournament = WillHow
Form2.feed.Enabled = WillHow
Form2.clean.Enabled = WillHow
Form2.rest.Enabled = WillHow
Form2.helptopic.Enabled = WillHow
Form2.errorfixed.Enabled = WillHow
Form2.About.Enabled = WillHow
Form2.CheckUpdate.Enabled = WillHow
Form2.Carry.Enabled = WillHow
Form2.reset.Enabled = WillHow
Form2.resettogether = WillHow
End Function

Sub Main()
Dim newlook
If left(Command, 7) = "change=" Then
Select Case Mid(Command, 8)
Case "Shitmon"
newlook = "Shitmon"
Case "Nidoran"
newlook = "Nidoran"
Case "Hakuryu"
newlook = "Hakuryu"
Case "Sandslash"
newlook = "Sandslash"
Case "Golduck"
newlook = "Golduck"
Case "Growlithe"
newlook = "Growlithe"
Case "Haunter"
newlook = "Haunter"
Case "Lapras"
newlook = "Lapras"
Case "Farfetch'd"
newlook = "Farfetch'd"
Case "Scyther"
newlook = "Scyther"
End Select
If Not newlook = "" Then
moneyvalue = GetSetting("Digimon", "Profile", "Money")
moneyvalue = Val(moneyvalue) - Val(50000)
If Not left(moneyvalue, 1) = "-" Then
SaveSetting "Digimon", "Profile", "Money", moneyvalue
SaveSetting "Digimon", "Digimon", "Type", newlook
End If
End If
End If

If App.PrevInstance Then
MsgBox "Virtual Digimonz already running."
End
Exit Sub
End If

If App.EXEName = "Project1" Then
Form1.Show
Exit Sub
End If
If Not App.EXEName & ".EXE" = "VIRTUAL DIGIMONZ.EXE" Then
If Not App.EXEName & ".EXE" = "Virtual Digimonz.EXE" Then
MsgBox "Wrong Filename."
End
Exit Sub
End If
End If

If Not App.CompanyName = "Choong You Qi" Then
MsgBox "File Corrupted."
End
Exit Sub
End If

If GetSetting("Digimon", "Profile", "Money") = "" Then
If GetSetting("Digimon", "Application", "PathAndFile") = "" Then
Form1.Show
Exit Sub
End If
End If

If GetSetting("Digimon", "Application", "PathAndFile") = "" Then
Reset_Program
Form4.Show
Exit Sub
End If

ThisEXEFileSetting = App.Path & "\" & "Virtual Digimonz.exe"
If Not GetSetting("Digimon", "Application", "PathAndFile") = ThisEXEFileSetting Then
MsgBox "This file only may run at:(Your Setup Profile Directory)" & vbCrLf & GetSetting("Digimon", "Application", "PathAndFile")
End
Exit Sub
End If
Form1.Show
End Sub
Public Function WillEnergy(WillHow As String, EnergyValue As Long)
Select Case WillHow
Case "PLUS"
If Not Form2.ProgressBar1.Value + EnergyValue > "100" Then Form2.ProgressBar1.Value = Val(Form2.ProgressBar1.Value) + Val(EnergyValue)
Case "MINUS"
If Not Form2.ProgressBar1.Value - EnergyValue < "0" Then Form2.ProgressBar1.Value = Val(Form2.ProgressBar1.Value) - Val(EnergyValue)
End Select
Form2.ProgressBar1.ToolTipText = Form2.ProgressBar1.Value & "/100"
Form2.Label9.Caption = Form2.ProgressBar1.Value & "/100"
SaveSetting "Digimon", "Digimon", "Energy", Form2.ProgressBar1.Value
End Function
Public Function WillMoney(WillHow As String, Money As Long)
moneyvalue = GetSetting("Digimon", "Profile", "Money")
Select Case WillHow
Case "PLUS"
moneyvalue = Val(moneyvalue) + Val(Money)
Case "MINUS"
moneyvalue = Val(moneyvalue) - Val(Money)
If left(moneyvalue, 1) = "-" Then
MsgBox "You don't have enough money"
WillMoney = "ExitSub"
Exit Function
End If
End Select
SaveSetting "Digimon", "Profile", "Money", moneyvalue
End Function
Public Function WillPoint(WillHow As String, Point As Long)
pointvalue = GetSetting("Digimon", "Profile", "Score")
Select Case WillHow
Case "PLUS"
pointvalue = Val(pointvalue) + Val(Point)
Case "MINUS"
pointvalue = Val(pointvalue) - Val(Point)
If left(pointvalue, 1) = "-" Then
MsgBox "You don't have enough Exp.Points."
WillPoint = "ExitSub"
Exit Function
End If
End Select
SaveSetting "Digimon", "Profile", "Score", pointvalue
End Function
Public Function WillAttack(WillHow As String, Attackpoint As Long)
attackvalue = GetSetting("Digimon", "Digimon", "Power")
Select Case WillHow
Case "PLUS"
attackvalue = Val(attackvalue) + Val(Attackpoint)
Case "MINUS"
attackvalue = Val(attackvalue) - Val(Attackpoint)
If left(attackvalue, 1) = "-" Then attackvalue = "1"
If attackvalue = "0" Then attackvalue = "1"
End Select
SaveSetting "Digimon", "Digimon", "Power", attackvalue
End Function
Public Function WillDefence(WillHow As String, Defencepoint As Long)
defencevalue = GetSetting("Digimon", "Digimon", "Defence")
Select Case WillHow
Case "PLUS"
defencevalue = Val(defencevalue) + Val(Defencepoint)
Case "MINUS"
defencevalue = Val(defencevalue) - Val(Defencepoint)
If left(defencevalue, 1) = "-" Then defencevalue = "1"
If defencevalue = "0" Then defencevalue = "1"
End Select
SaveSetting "Digimon", "Digimon", "Defence", defencevalue
End Function
Public Function WillHealth(WillHow As String, Healthpoint As Long)
healthvalue = GetSetting("Digimon", "Digimon", "Health")
currenthealthvalue = GetSetting("Digimon", "Digimon", "CurrentHealth")
Select Case WillHow
Case "PLUS"
healthvalue = Val(healthvalue) + Val(Healthpoint)
currenthealthvalue = Val(currenthealthvalue) + Val(Healthpoint)
Case "MINUS"
healthvalue = Val(healthvalue) - Val(Healthpoint)
currenthealthvalue = Val(currenthealthvalue) - Val(Healthpoint)
If left(healthvalue, 1) = "-" Then healthvalue = "1"
If healthvalue = "0" Then healthvalue = "1"
If left(currenthealthvalue, 1) = "-" Then currenthealthvalue = "0"
End Select
SaveSetting "Digimon", "Digimon", "Health", healthvalue
SaveSetting "Digimon", "Digimon", "CurrentHealth", currenthealthvalue
End Function
Public Function WillMaxFood(WillHow As String, MaxFoodpoint As Long)
maxfoodvalue = GetSetting("Digimon", "Profile", "MaxFood")
Select Case WillHow
Case "PLUS"
maxfoodvalue = Val(maxfoodvalue) + Val(MaxFoodpoint)
Case "MINUS"
maxfoodvalue = Val(maxfoodvalue) - Val(MaxFoodpoint)
If left(maxfoodvalue, 1) = "-" Then maxfoodvalue = "1"
If maxfoodvalue = "0" Then maxfoodvalue = "1"
End Select
SaveSetting "Digimon", "Profile", "MaxFood", maxfoodvalue
End Function
Public Function WillFood(WillHow As String, Foodpoint As Long)
foodvalue = GetSetting("Digimon", "Digimon", "Fooded")
maxfoodvalue = GetSetting("Digimon", "Profile", "MaxFood")
Select Case WillHow
Case "PLUS"
If Val(foodvalue) + Val(Foodpoint) > maxfoodvalue Then
foodvalue = maxfoodvalue
Else
foodvalue = Val(foodvalue) + Val(Foodpoint)
End If
Case "MINUS"
foodvalue = Val(foodvalue) - Val(Foodpoint)
If left(foodvalue, 1) = "-" Then
Reset_Program
MsgBox "Your Digimon Die. Cause it's too hungry."
End
Exit Function
End If
End Select
SaveSetting "Digimon", "Digimon", "Fooded", foodvalue
Form2.Label5.Tag = GetSetting("Digimon", "Digimon", "Fooded")
Food_Event_Caption
End Function
Public Function WillBank(WillHow As String, BankMoney As Long)
bankvalue = GetSetting("Digimon", "Profile", "Bank")
Select Case WillHow
Case "PLUS"
bankvalue = Val(bankvalue) + Val(BankMoney)
Case "MINUS"
bankvalue = Val(bankvalue) - Val(BankMoney)
If left(bankvalue, 1) = "-" Then
MsgBox "You don't have enough money in your bank."
WillBank = "ExitSub"
Exit Function
End If
End Select
SaveSetting "Digimon", "Profile", "Bank", bankvalue
End Function

Public Function Encode(Txt)
On Local Error Resume Next
Dim Letter As Variant
Dim Total As String

    For Counter = 1 To Len(Txt)
        Letter = Mid$(Txt, Counter, 1)
        Letter = Asc(Letter)
        Letter = Letter + 2
        Letter = Chr(Letter)
        Total = Total & Letter
    Next Counter
    Encode = Total
End Function
Public Function Decode(Txt)
On Local Error Resume Next
Dim Letter As Variant
Dim Total As String

    For Counter = 1 To Len(Txt)
        Letter = Mid$(Txt, Counter, 1)
        Letter = Asc(Letter)
        Letter = Letter - 2
        Letter = Chr(Letter)
        Total = Total & Letter
    Next Counter
    Decode = Total
End Function
