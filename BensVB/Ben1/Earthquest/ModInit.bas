Attribute VB_Name = "Module1"
Option Explicit
Global ActiveWeapon
Global Unit() As UnitType
Global Weapon(0 To 10) As Weapon
Global ActiveUnit As Integer
Global Turn As Integer, Teams As Integer
Global ViewX As Single
Global ViewY As Single
Global Units As Integer
Global UData(1 To 1) As UnitData
Global UnitDC(1 To 1, 1 To 8), MaskDC(1 To 1, 1 To 8) 'type, position
Global ActiveSlot
'define types for unit data to be used in manipulation of user interface and storing data for units
Type UnitType
    Name As String 'a personalized name for the unit
    Owner As Integer 'the player that owns the unit
    GraphicsDC As Integer 'a value used to determine which sprites will be blitted to the unit's position on the screen
    Position As Integer
    Weapon(0 To 3) As Integer 'a number that refers to a weapon type in slots 1 to 4
    Skill(0 To 3) As Double  ' skill with weapons 1 to 4
    Melee As Double 'multiplier for melee attacks
    Ranged As Double 'multiplier for ranged attacks
    Sight As Double
    Dexterity As Double
    Strength As Double
    Toughness As Double 'defense against attacks
    Health As Integer 'the unit's current health
    MaxHP As Integer ' the unit's maximum HP - stays constant unless the unit is upgraded
    X As Single 'coordinates on screen
    Y As Single
    TargetX As Single 'coordinates the unit is trying to reach
    TargetY As Single
    Movement As Integer ' the maximum distance a unit can move in one turn, in pixels
    Moved As Boolean 'whether the unit moved this turn
    Attacked As Boolean 'whether the unit attacked this turn
    Dead As Boolean
    Size As Integer 'the radius of the unit, in pixels
    level As Integer 'Unit's overall experience level
    levelThreshold As Integer 'xp necessary to level up
End Type
Type UnitData
    Cost As Integer 'cost in points
    GraphicsDC As Integer 'a value used to determine which sprites will be blitted to the unit's position on the screen
    Weapon(0 To 3) As Integer 'a number that refers to a weapon type in slots 1 to 4 - >= 0 is locked with a certain weapon _
    or can't put a weapon on that hardpoint. -1 = empty, usable hardpoint
    Skill(0 To 3) As Long  ' skill with weapons 1 to 4
    Melee As Double 'multiplier for melee attacks
    Ranged As Double 'multiplier for ranged attacks
    Sight As Long
    Dexterity As Long
    Strength As Long
    'Armor(1 To 4) As Integer 'pierce/slash/crush/burn
    MaxHP As Integer ' the unit's maximum HP - stays constant unless the unit is upgraded
    Movement As Integer ' the maximum distance a unit can move in one turn, in pixels
    Size As Integer
    level As Integer
    levelThreshold As Integer
End Type
Type Weapon
    Name As String 'the name of the weapon type
    Cost As Integer 'cost in points
    Damage As Integer 'damage the weapon does
    MaxRange As Integer 'maximum range
    MinRange As Integer 'minimum range
    AttackType As Integer '1, 2, 3, 4 = pierce, slash, crush, burn
    RateOfAttack As Integer
    SightX As Integer '1 = damage, 2 = range, 3 = rate of attack
    DexterityX As Integer 'ditto
    StrengthX As Integer 'ditto
    Accuracy As Double 'probability % for attack to hit - altered by skill of unit
    Defense(1 To 4) As Double 'pierce, slash, crush, burn
End Type
Type ObjectData
    Name As String
    GoldCost As Integer 'the amount of gold necessary to make the object
    FoodCost As Integer 'the amount of food necessary
    WoodCost As Integer 'the amount of wood necessary
    StoneCost As Integer 'the amount of stone necessary
    ManaCost  As Integer 'the amount of mana nessary
    GraphicsDC As Integer 'the graphics dc index for the object
    GraphicsSize As Integer 'the width and height of the graphics image
    BaseRadius As Integer 'the radius of the object's base -- determines impassable area
    Selectable As Boolean 'whether the object can be selected by user
    Building As Boolean ' whether the object is a building
    Unit As Boolean ' whether the object is a unit
    MoveRate As Integer ' how many pixels the object can be moved in a turn
    Spawn(1 To 15) As Integer ' units that can be created by this unit or building
    Research(1 To 15) As Integer ' techs or upgrades that this unit or building can research
    SpawnRate As Integer ' how many pp's (production points) this object produces per turn
    ResearchRate As Integer ' how many rp's (research points) this object produces per turn
    HP As Integer 'hitpoints
    Damage(1 To 2, 1 To 4) As Integer 'melee/ranged; pierce/slash/crush/fire
    Armor(1 To 2, 1 To 4) As Double '% of damage to block -- same array setup
    Range(1 To 2) As Integer 'maximum range for ranged attack/range for melee
    ProjectileDC As Integer 'what projectile graphics to use
    Mutates As Boolean 'another object this object can change into
    Mutate As Integer 'this number refers to the data index for the object this object can change into
    MutateOnDeath As Boolean ' mutates when destroyed -- i.e. comes back as a skeleton etc.
    MutateOnAttack As Boolean ' mutates when attacking -- i.e. packed trebuchet > unpacked
    MutateOnDefense As Boolean ' mutates when attacked
    MutateOnMovement As Boolean ' mutates when moved -- i.e. unpacked trebuchet > packed
    MutateAfterTime As Boolean ' mutates a certain number of turns after creation or last mutation -- i.e. phoenix respawns from egg
    MutateTime As Integer 'if above variable set to true, determines # of turns before mutation
    ManaPerDestroyedHP As Double 'the amount of mana player gets for each hp destroyed by this unit
    ManaOnDeath As Integer 'the amount of mana player gets when this unit is destroyed
    ManaOnBuilt As Integer 'the amount of mana player gets when this unit is built
    RegenRate As Integer 'the amount of hp unit regains per turn
    Magic As Integer 'the amount of magic points the unit has
    MagicRegen As Integer 'the amount of magic points the unit regains per turn
    CanMine As Boolean 'whether the unit or building can execute the mining order
    CanLog As Boolean 'whether the unit or building can gather wood
    CanGather As Boolean 'whether the unit or building can gather food
    MineRate As Integer 'how much the unit mines each turn
    LogRate As Integer 'how much wood the unit cuts each turn
    GatherRate As Integer 'how much food the unit gathers each turn
    Support As Integer 'how many population points the building supports
    Population As Integer 'how many population points the unit costs
    DecayTime As Integer 'how long it takes for a dead unit to decay
    Ressurectable As Boolean 'whether the necromany ability can be used on this unit
    NecromancyUnit As Integer 'the unit produced when necromancy ability is used
    NecroCost As Integer 'the amount of magic necessary to ressurect the unit
    Necromancy As Integer 'whether this unit has the necromany ability
    NecroRange As Integer 'the distance at which a unit can ressurect units
End Type
Public Function CheckRange(X1 As Single, Y1 As Single, X2 As Single, Y2 As Single, Range As Single) As Boolean
Dim Distance, XVar, YVar
XVar = X2 - X1
YVar = Y2 - Y1
Distance = Sqr(XVar ^ 2 + YVar ^ 2)
If Distance > Range Then
CheckRange = False
Else
CheckRange = True
End If
End Function

Public Sub Attack(Attacker, Defender)
Dim ADamage, ARate, AProb, ADefense, i, X, Hit As Boolean, TotalDamage, TotalXP
AProb = 1 - Weapon(ActiveWeapon).Accuracy
AProb = AProb / Unit(ActiveUnit).Skill(ActiveSlot)
AProb = 1 - AProb
AProb = Weapon(ActiveWeapon).Accuracy
Select Case Weapon(ActiveWeapon).SightX
    Case Is = 1 'Damage
    ADamage = Weapon(ActiveWeapon).Damage * Unit(Attacker).Sight
    Case Is = 3 'Rate of Attack
    ARate = Int(Weapon(ActiveWeapon).RateOfAttack * Unit(Attacker).Sight)
End Select
Select Case Weapon(ActiveWeapon).StrengthX
    Case Is = 1 'Damage
    ADamage = Weapon(ActiveWeapon).Damage * Unit(Attacker).Strength
    Case Is = 3 'Rate of Attack
    ARate = Int(Weapon(ActiveWeapon).RateOfAttack * Unit(Attacker).Strength)
End Select
Select Case Weapon(ActiveWeapon).DexterityX
    Case Is = 1 'Damage
    ADamage = Weapon(ActiveWeapon).Damage * Unit(Attacker).Dexterity
    Case Is = 3 'Rate of Attack
    ARate = Int(Weapon(ActiveWeapon).RateOfAttack)
End Select
ADamage = Weapon(ActiveWeapon).Damage * Unit(ActiveUnit).Skill(ActiveSlot)
Unit(ActiveUnit).Skill(ActiveSlot) = Unit(ActiveUnit).Skill(ActiveSlot) + 0.05
ADefense = 1
For i = 0 To 3
If Unit(Defender).Weapon(i) > 0 Then ADefense = ADefense * Weapon(Unit(Defender).Weapon(i)).Defense(Weapon(ActiveWeapon).AttackType) * Unit(Defender).Skill(i)
Next i
ADamage = ADamage / ADefense
For i = 1 To ARate
'Reduce Defender's health if attack succeeds
If Rnd < AProb Then
TotalDamage = TotalDamage + ADamage
Hit = True
End If
Next i
Unit(Defender).Health = Unit(Defender).Health - TotalDamage
frmInterface.lblMessages.Caption = Unit(Attacker).Name & " ATTACKED " & Unit(Defender).Name & " FOR " & Int(TotalDamage) & "."
If Hit = True Then Unit(ActiveUnit).Skill(ActiveSlot) = Unit(ActiveUnit).Skill(ActiveSlot) + 0.05
'Check to see if opponent is dead
If Unit(Defender).Health <= 0 Then
Unit(Defender).Dead = True
Unit(ActiveUnit).Skill(ActiveSlot) = Unit(ActiveUnit).Skill(ActiveSlot) + 0.1
frmInterface.lblMessages.Caption = Unit(Attacker).Name & " KILLED " & Unit(Defender).Name & "."
End If
'Check to see if unit levels up
For i = 0 To 3
If Unit(ActiveUnit).Skill(i) * 100 > 100 Then TotalXP = TotalXP + Unit(ActiveUnit).Skill(i) * 100 - 100
Next i
If TotalXP >= Unit(ActiveUnit).levelThreshold Then
    With Unit(ActiveUnit)
    .level = Unit(ActiveUnit).level + 1
    .levelThreshold = Unit(ActiveUnit).levelThreshold * 2.1
    .MaxHP = Unit(ActiveUnit).MaxHP * 1.25
    .Health = Unit(ActiveUnit).Health * 1.25
    .Movement = Unit(ActiveUnit).Movement * 1.5
    End With
End If
End Sub

Public Function GenerateName()
Dim Start(1 To 20), Middle(1 To 20), Ending(1 To 20), Syllables, i
Randomize
Start(1) = "NA"
Start(2) = "GA"
Start(3) = "GIL"
Start(4) = "CEL"
Start(5) = "UN"
Start(6) = "LE"
Start(7) = "TI"
Start(8) = "EL"
Start(9) = "LA"
Start(10) = "AN"
Start(11) = "DEN"
Start(12) = "REN"
Start(13) = "MITH"
Start(14) = "TAS"
Start(15) = "FAN"
Start(16) = "LIN"
Start(17) = "VAL"
Start(18) = "CIR"
Start(19) = "TA"
Start(20) = "GLOR"
Middle(1) = "RUN"
Middle(2) = "LAD"
Middle(3) = "GA"
Middle(4) = "EB"
Middle(5) = "DOM"
Middle(6) = "GO"
Middle(7) = "BE"
Middle(8) = "BER"
Middle(9) = "NA"
Middle(10) = "THON"
Ending(1) = "NIEL"
Ending(2) = "RIEL"
Ending(3) = "WYN"
Ending(4) = "LAD"
Ending(5) = "NORN"
Ending(6) = "REN"
Ending(7) = "LAS"
Ending(8) = "LON"
Ending(9) = "NATH"
Ending(10) = "NIE"
Syllables = Int(Rnd * 2)
GenerateName = Start(Int(Rnd * 20 + 1))
If Syllables > 0 Then
For i = 1 To Syllables
GenerateName = GenerateName & Middle(Int(Rnd * 10 + 1))
Next i
End If
GenerateName = GenerateName & Ending(Int(Rnd * 10 + 1))
End Function
Public Sub InitiateUnits()
With UData(1)
.Cost = 100
.Dexterity = 1
.GraphicsDC = 1
.MaxHP = 100
.Movement = 300
.Sight = 1.2
.Size = 30
.Skill(0) = 1.3
.Skill(1) = 1
.Skill(2) = 0.9
.Skill(3) = 0.8
.Strength = 1.2
.Weapon(0) = -1
.Weapon(1) = -1
.Weapon(2) = -1
.Weapon(3) = -1
.level = 1
.levelThreshold = 60
End With
With Weapon(1) 'longsword
.Name = "Longsword"
.Accuracy = 0.9
.AttackType = 2
.Cost = 45
.Damage = 60
.Defense(1) = 1.4
.Defense(2) = 1.9
.Defense(3) = 1
.Defense(4) = 1
.DexterityX = 3
.MaxRange = 50
.MinRange = 0
.RateOfAttack = 1
.SightX = 0
.StrengthX = 1
End With
With Weapon(2) 'bow
.Name = "Bow"
.Accuracy = 0.8
.AttackType = 1
.Cost = 35
.Damage = 20
.Defense(1) = 1
.Defense(2) = 1
.Defense(3) = 1
.Defense(4) = 1
.DexterityX = 3
.MaxRange = 800
.MinRange = 0
.RateOfAttack = 2
.SightX = 0
.StrengthX = 1
End With
With Weapon(3) 'short sword
.Name = "Short Sword"
.Accuracy = 0.7
.AttackType = 2
.Cost = 20
.Damage = 12
.Defense(1) = 1.2
.Defense(2) = 1.5
.Defense(3) = 1.1
.Defense(4) = 1
.DexterityX = 3
.MaxRange = 40
.MinRange = 0
.RateOfAttack = 4
.SightX = 0
.StrengthX = 1
End With
With Weapon(4) 'Chain Mail
.Name = "Chain Mail"
.Accuracy = 0
.AttackType = 0
.Cost = 20
.Damage = 0
.Defense(1) = 1.5
.Defense(2) = 1.3
.Defense(3) = 1.3
.Defense(4) = 1.2
.DexterityX = 3
.MaxRange = 40
.MinRange = 0
.RateOfAttack = 4
.SightX = 0
.StrengthX = 1
End With
With Weapon(5) 'Spear
.Name = "Spear"
.Accuracy = 95
.AttackType = 1
.Cost = 25
.Damage = 20
.Defense(1) = 2
.Defense(2) = 2
.Defense(3) = 1
.Defense(4) = 1
.DexterityX = 3
.MaxRange = 110
.MinRange = 0
.RateOfAttack = 1
.SightX = 0
.StrengthX = 1
End With
With Weapon(6) 'Flail
.Name = "Flail"
.Accuracy = 0.55
.AttackType = 3
.Cost = 30
.Damage = 5
.Defense(1) = 1
.Defense(2) = 1
.Defense(3) = 1
.Defense(4) = 1
.DexterityX = 3
.MaxRange = 50
.MinRange = 0
.RateOfAttack = 20
.SightX = 0
.StrengthX = 1
End With
With Weapon(7) 'Crossbow
.Name = "Crossbow"
.Accuracy = 0.8
.AttackType = 1
.Cost = 50
.Damage = 40
.Defense(1) = 1
.Defense(2) = 1
.Defense(3) = 1
.Defense(4) = 1
.DexterityX = 3
.MaxRange = 600
.MinRange = 0
.RateOfAttack = 1
.SightX = 0
.StrengthX = 1
End With
With Weapon(8) 'Leather Armor
.Name = "Leather Armor"
.Accuracy = 1
.AttackType = 1
.Cost = 30
.Damage = 0
.Defense(1) = 1.1
.Defense(2) = 1.4
.Defense(3) = 2.3
.Defense(4) = 1
.DexterityX = 3
.MaxRange = 0
.MinRange = 0
.RateOfAttack = 1
.SightX = 0
.StrengthX = 1
End With
With Weapon(9) 'Flaming Sword Of Doom
.Name = "Flaming Sword"
.Accuracy = 0.9
.AttackType = 4
.Cost = 100
.Damage = 40
.Defense(1) = 1
.Defense(2) = 2
.Defense(3) = 1
.Defense(4) = 1
.DexterityX = 3
.MaxRange = 70
.MinRange = 40
.RateOfAttack = 1
.SightX = 0
.StrengthX = 1
End With
End Sub
Public Sub LoadUnit(DataIndex, X, Y, Owner, Weapon1, Weapon2, Weapon3, Weapon4, Name)
Dim i, n
Units = Units + 1
ReDim Preserve Unit(1 To Units)
'For i = 1 To 2
'For n = 1 To 4
'Unit(Units).Attack(i, n) = UData(DataIndex).Attack(i, n)
'Unit(Units).Armor(i, n) = UData(DataIndex).Armor(i, n)
'Next n
'Next i
For i = 0 To 3
Unit(Units).Skill(i) = UData(DataIndex).Skill(i)
Next i
With Unit(Units)
.Name = Name
.Owner = Owner
.GraphicsDC = UData(DataIndex).GraphicsDC
.MaxHP = UData(DataIndex).MaxHP
.Health = UData(DataIndex).MaxHP
.Movement = UData(DataIndex).Movement
.Sight = UData(DataIndex).Sight
.Strength = UData(DataIndex).Strength
.Dexterity = UData(DataIndex).Dexterity
.X = X
.Y = Y
.Size = UData(DataIndex).Size
.Position = 1
.Weapon(0) = Weapon1
.Weapon(1) = Weapon2
.Weapon(2) = Weapon3
.Weapon(3) = Weapon4
.level = UData(DataIndex).level
.levelThreshold = UData(DataIndex).levelThreshold
End With
For i = 0 To 3
If Unit(Units).Weapon(i) = -1 Then Unit(Units).Weapon(i) = 0
Next i
End Sub

