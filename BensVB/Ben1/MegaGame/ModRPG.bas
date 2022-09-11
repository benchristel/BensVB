Attribute VB_Name = "Module1"
Type Player
    Location As Integer 'index of the point on the matrix where the player currently is
    Destination As Integer 'index of the point where the player is currently headed
    X As Integer ' x-coordinate of player on the map
    Position As Integer ' whether player is facing left or right
    'inventory
    Inventory(0 To 9) As Integer 'index of item in inventory
    ItemCount(0 To 9) As Integer 'number of items in that slot
    CargoVolume As Integer 'total volume of all items in inventory
    CargoMax As Integer 'maximum cargo volume
    TotalWeight As Integer
    ' equipment slots
    WeaponHardpoint(1 To 3) As Itemstats  '1 = wingtip 2= wingmid 3 = nose
    Booster As Itemstats
    Armor As Itemstats
    Bay As Itemstats
    FuelTank As Itemstats
    Fuel As Integer
    MaxFuel As Integer
    CurrentWeapon As Integer
End Type
Type Location
    Name As String ' name of the place
    Link(1 To 10) As Integer ' indexes of places you can go from this point
    Buy(1 To 6) As Double ' percentage of average price people here will buy stuff for
    '1 = swords, 2 = bows, 3 = shields, 4 = armor, 5 = magic, 6 = other stuff
    Sell(1 To 6) As Double ' percentage of average price people here will sell stuff for
    CreatureType(1 To 5) As Integer 'indexes of creature types you will encounter in this place
    CreatureNumber(1 To 5) As Integer 'number of creatures of types referred to by creaturetypes 1 to 5
End Type

Type ItemData
    Name As String
    AttackBonus As Integer
    DefenseBonus As Integer
    Arrow As Boolean
    MiningLvl As Integer
    WoodcuttingLvl As Integer
    
    Type As Integer
        '1 = swords, 2 = bows, 3 = shields, 4 = armor, 5 = magic, 6 = other stuff
    ReloadTime As Integer 'frames to reload weapon
    Ranged As Boolean
    Cost As Integer
End Type

Type Creature
    Name As String
    Treasure As Integer 'amount of gold you get from this creature
    HP As Integer
    Movement As Integer
End Type
