Attribute VB_Name = "Module1"
Option Explicit
Type ObjectData
    Name As String
    GoldCost As Integer 'the amount of gold necessary to make the object
    FoodCost As Integer 'the amount of food necessary
    WoodCost As Integer 'the amount of wood necessary
    StoneCost As Integer 'the amount of stone necessary
    ManaCost  As Integer 'the amount of mana nessary
    ProductionCost As Integer 'the amount of pp's necessary to build this
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
    damage(1 To 2, 1 To 4) As Integer 'melee/ranged; pierce/slash/crush/fire
    Armor(1 To 2, 1 To 4) As Double '% of damage to block -- same array setup
    Range(1 To 2) As Integer 'maximum range for ranged attack/range for melee
    RetaliateArmor(1 To 2, 1 To 4) As Double 'armor during retaliation (<1)
    RetaliateDamage(1 To 2, 1 To 4) As Integer 'amount of damage done when retaliating
    ProjectileDC As Integer 'what projectile graphics to use
    Mutates As Boolean 'another object this object can change into
    Mutate As Integer 'this number refers to the data index for the object this object can change into
    MutateOnDeath As Boolean ' mutates when destroyed -- i.e. comes back as a skeleton etc.
    MutateOnAttack As Boolean ' mutates when attacking -- i.e. packed trebuchet > unpacked
    MutateOnDefense As Boolean ' mutates when attacked
    MutateOnMovement As Boolean ' mutates when moved -- i.e. unpacked trebuchet > packed
    MutateAfterTime As Boolean ' mutates a certain number of turns after creation or last mutation -- i.e. phoenix respawns from egg
    MutateTime As Integer 'if above variable set to true, determines # of turns before mutation
    MutateOnCommand As Boolean ' mutates if a button is pressed
    ManaPerDestroyedHP As Double 'the amount of mana player gets for each hp destroyed by this unit
    ManaOnDeath As Integer 'the amount of mana player gets when this unit is destroyed
    ManaOnBuilt As Integer 'the amount of mana player gets when this unit is built
    RegenRate As Integer 'the amount of hp unit regains per turn
    Magic As Integer 'the amount of magic points the unit has
    MagicRegen As Integer 'the amount of magic points the unit regains per turn
    CanMine As Boolean 'whether the unit or building can execute the mining order
    CanLog As Boolean 'whether the unit or building can gather wood
    CanGather As Boolean 'whether the unit or building can gather food
    CanBuild As Boolean 'wheter the unit can add to buildings
    CanAttack As Boolean
    MineRate As Integer 'how much the unit mines each turn
    LogRate As Integer 'how much wood the unit cuts each turn
    GatherRate As Integer 'how much food the unit gathers each turn
    BuildRate As Integer 'how fast the unit builds
    Support As Integer 'how many population points the building supports
    Population As Integer 'how many population points the unit costs
    DecayTime As Integer 'how long it takes for a dead unit to decay
    Ressurectable As Boolean 'whether the necromany ability can be used on this unit
    NecromancyUnit As Integer 'the unit produced when necromancy ability is used
    NecroCost As Integer 'the amount of magic necessary to ressurect the unit
    Necromancy As Integer 'whether this unit has the necromany ability
    NecroRange As Integer 'the distance at which a unit can ressurect units
    ImpassableWhenDead As Boolean
    Impassable As Boolean
    Food As Integer 'the amount of food left in a farm
    Wood As Integer
    Gold As Integer
    Stone As Integer
End Type
Type Objectstats
    Name As String
    x As Integer 'x-coordinate on screen
    y As Integer 'y-coordinate on screen
    Position As Integer 'the rotation position to blit to screen
    GraphicsDC As Integer 'the graphics dc index for the object
    Mode As Integer '1 = stand; 2 = attack; 3 = defend; 4 = die; 5 = dead; 6 = construct
    CurrFrame As Integer 'current frame of animation
    GraphicsSize As Integer 'the width and height of the graphics image
    BaseRadius As Integer 'the radius of the object's base -- determines impassable area
    Selectable As Boolean 'whether the object can be selected by user
    Building As Boolean ' whether the object is a building
    Unit As Boolean ' whether the object is a unit
    MoveRate As Integer ' how many pixels the object can be moved in a turn
    Spawn(1 To 15) As Integer ' units that can be created by this unit or building
    Research(1 To 15) As Integer ' techs or upgrades that this unit or building can research
    Researching As Integer 'current tech being researched
    NeededRPs As Integer 'how many total rp's are needed to research current project
    CurrentRPs As Integer 'the current number of rp's in production
    SpawnRate As Integer ' how many pp's (production points) this object produces per turn
    ResearchRate As Integer ' how many rp's (research points) this object produces per turn
    HP As Integer 'hitpoints
    MaxHP As Integer
    damage(1 To 2, 1 To 4) As Integer 'melee/ranged; pierce/slash/crush/fire
    Armor(1 To 2, 1 To 4) As Double '% of damage to block -- same array setup
    Range(1 To 2) As Integer 'maximum range for ranged attack/range for melee
    RetaliateArmor(1 To 2, 1 To 4) As Double 'armor during retaliation (<1)
    RetaliateDamage(1 To 2, 1 To 4) As Integer 'amount of damage done when retaliating
    ProjectileDC As Integer 'what projectile graphics to use
    Mutates As Boolean 'another object this object can change into
    Mutate As Integer 'this number refers to the data index for the object this object can change into
    MutateOnDeath As Boolean ' mutates when destroyed -- i.e. comes back as a skeleton etc.
    MutateOnAttack As Boolean ' mutates when attacking -- i.e. packed trebuchet > unpacked
    MutateOnDefense As Boolean ' mutates when attacked
    MutateOnMovement As Boolean ' mutates when moved -- i.e. unpacked trebuchet > packed
    MutateAfterTime As Boolean ' mutates a certain number of turns after creation or last mutation -- i.e. phoenix respawns from egg
    MutateTime As Integer 'if above variable set to true, determines # of turns before mutation
    MutateOnCommand As Boolean ' mutates if a button is pressed
    ManaPerDestroyedHP As Double 'the amount of mana player gets for each hp destroyed by this unit
    ManaOnDeath As Integer 'the amount of mana player gets when this unit is destroyed
    ManaOnBuilt As Integer 'the amount of mana player gets when this unit is built
    RegenRate As Integer 'the amount of hp unit regains per turn
    Magic As Integer 'the amount of magic points the unit has
    MagicRegen As Integer 'the amount of magic points the unit regains per turn
    CanMine As Boolean 'whether the unit or building can execute the mining order
    CanLog As Boolean 'whether the unit or building can gather wood
    CanGather As Boolean 'whether the unit or building can gather food
    CanBuild As Boolean 'whether the unit can add to buildings
    CanAttack As Boolean
    MineRate As Integer 'how much the unit mines each turn
    LogRate As Integer 'how much wood the unit cuts each turn
    GatherRate As Integer 'how much food the unit gathers each turn
    BuildRate As Integer 'how fast (HP/turn) the unit builds
    Support As Integer 'how many population points the building supports
    Population As Integer 'how many population points the unit costs
    DecayTime As Integer 'how long it takes for a dead unit to decay
    Ressurectable As Boolean 'whether the necromany ability can be used on this unit
    NecromancyUnit As Integer 'the unit produced when necromancy ability is used
    NecroCost As Integer 'the amount of magic necessary to ressurect the unit
    Necromancy As Integer 'whether this unit has the necromany ability
    NecroRange As Integer 'the distance at which a unit can ressurect units
    Dead As Boolean
    Deleted As Boolean
    Stance As Integer '1 = defensive, 2 = offensive
    ImpassableWhenDead As Boolean
    Impassable As Boolean
    Food As Integer 'the amount of food left in a farm
    Wood As Integer 'ditto for wood
    Gold As Integer 'ditto
    Stone As Integer 'ditto
End Type
Type Player
    Gold As Integer
    Food As Integer
    Wood As Integer
    Stone As Integer
    Mana As Integer
    Population As Integer
    Units As Integer
    Color As Integer
    TechDiscovered(1 To 20) As Integer
    CreateObjectEnabled(1 To 20) As Integer
    ResearchEnabled(1 To 20) As Integer
    ActionEnabled(1 To 10) As Integer
End Type
Type Technology
    GoldCost As Integer
    FoodCost As Integer
    WoodCost As Integer
    StoneCost As Integer
    ManaCost As Integer
    RPCost As Integer
    PictureDC As Integer
    EnableTech(1 To 5) As Integer
    EnableObject(1 To 5) As Integer
    EnableAction(1 To 3) As Integer
    AttackBonus As Double
    DefenseBonus As Double
    RangeBonus(1 To 2) As Integer
    MoveBonus As Integer
    GiveGather As Boolean
    GiveMine As Boolean
    GiveLog As Boolean
    GiveAttack As Boolean
    GatherAdd As Boolean
    MineAdd As Boolean
    LogAdd As Boolean
    HPBonus As Double
    AffectsObject(1 To 5) As Integer
    AffectsInfantry As Boolean
    AffectsSiege As Boolean
    AffectsRanged As Boolean
    AffectsBuiding As Boolean
    AffectsUnit As Boolean
    AffectsAir As Boolean
End Type
Type Action
    Name As String
    Enabled As Boolean
    Button As Integer
End Type
Type GraphicsDC
    Frames(1 To 6) As Integer 'for different modes: stand, attack, defend, die, dead, construction
    DC(1 To 6, 1 To 5, 1 To 8) As Long 'mode, frame of animation, rotation
    Path(1 To 6, 1 To 5, 1 To 8) As String 'the path to the frame on disk
End Type
Dim Objectstats() As Objectstats, Player() As Player, ObjectData() As ObjectData
Dim Objects As Integer, GraphicsDC As GraphicsDC
Dim Tool As Integer '0 = move/select/attack/build/mine, -1 = necromancy, else = place object(x) on map
Public Sub Attack(Attacker As Integer, Defender As Integer, Ranged As Integer)
Dim damage(1 To 2, 1 To 4), i, j, RetaliateDamage(1 To 2, 1 To 4) As Integer
For j = 1 To 4
If Objectstats(Defender).Stance = 1 Then 'defensive stance
damage(Ranged, j) = Objectstats(Attacker).damage(i, j) - Objectstats(Attacker).damage(i, j) * Objectstats(Defender).Armor(i, j)
Else 'offensive stance
damage(Ranged, j) = Objectstats(Attacker).damage(i, j) - Objectstats(Attacker).damage(i, j) * Objectstats(Defender).RetaliateArmor(i, j)
If Distance(Objectstats(Attacker).x, Objectstats(Attacker).y, Objectstats(Defender).x, Objectstats(Defender).y) <= Objectstats(Defender).Range(Ranged) Then RetaliateDamage(i, j) = Objectstats(Defender).RetaliateDamage(Ranged, j) - Objectstats(Defender).RetaliateDamage(Ranged, j) * Objectstats(Attacker).Armor(Ranged, j)
End If
Next j
damage(1, 1) = damage(1, 1) + damage(1, 2) + damage(1, 3) + damage(1, 4) + damage(2, 1) + damage(2, 2) + damage(2, 3) + damage(2, 4)
RetaliateDamage(1, 1) = RetaliateDamage(1, 1) + RetaliateDamage(1, 2) + RetaliateDamage(1, 3) + RetaliateDamage(1, 4) + RetaliateDamage(2, 1) + RetaliateDamage(2, 2) + RetaliateDamage(2, 3) + RetaliateDamage(2, 4)
Objectstats(Attacker).HP = Objectstats(Attacker).HP - RetaliateDamage
Objectstats(Defender).HP = Objectstats(Defender).HP - damage
If Objectstats(Attacker).HP <= 0 Then
Objectstats(Attacker).Dead = True
Objectstats(Attacker).Mode = 4
Objectstats(Attacker).CurrFrame = 1
End If
If Objectstats(Defender).HP <= 0 Then
Objectstats(Defender).Dead = True
Objectstats(Attacker).Mode = 4
End Sub
