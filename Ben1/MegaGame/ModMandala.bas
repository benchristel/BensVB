Attribute VB_Name = "Module2"
Option Explicit
    Type Terrain
    Name As String
    Passable As Boolean
    DCx As Integer
    DCy As Integer
    End Type
    '
    Type EnemyData
    Name As String
    Range As Integer
    CDT As Integer
    CDTRemaining As Integer
    DC As Integer
    Dead As Boolean
    HP As Integer
    Damage As Integer
    ShotDC As Integer
    ShotSpeed As Integer
    AreaEffectRadius As Integer
    Hostile As Boolean
    XPGiven As Integer
    End Type
    '
    Type Shot
    Range As Integer
    DC As Integer
    Damage As Integer
    Speed As Integer
    Friendly As Boolean
    End Type
    '
    Type Store
    Inventory(1 To 20) As Integer
    SellPrice(1 To 20) As Integer
    BuyPrice(1 To 20) As Integer
    End Type
    '
    Type ItemData
    Name As String
    Hardpoint As Integer '0 = nose, 1 = sides, 2 = turret, 3 = front, _
    4 = rim armor 5 = pod armor, 6 = turret armor, 7 - 11 = inventory
    UseMode As Integer '0 = cannot be used, 1 = repair (health kit), _
    2 = empty ammo into inventory, 3 = equip as ammo
    Weapon As Boolean
    Ammo As Boolean
    LoadToWeapon As Integer 'if this item is ammunition, load to this weapon index
    Weight As Integer 'per-unit weight in kg's
    AddCaps As Integer 'number of capacity pts to add when equipped
    AddSpeed As Integer 'speed to add when equipped
    End Type
    '
    Type Item
    Name As String
    x As Integer
    y As Integer
    Hardpoint As Integer '0 = nose, 1 = sides, 2 = turret, 3 = front, _
    4 = rim armor 5 = pod armor, 6 = turret armor, 7 - 11 = inventory
    UseMode As Integer '0 = cannot be used, 1 = repair (health kit), _
    2 = empty ammo into inventory, 3 = equip as ammo
    Weapon As Boolean
    Ammo As Boolean
    LoadToWeapon As Integer 'if this item is ammunition, load to this weapon index
    Weight As Integer 'per-unit weight in kg's
    AddCaps As Integer 'number of capacity pts to add when equipped
    AddSpeed As Integer 'speed to add when equipped
End Type
