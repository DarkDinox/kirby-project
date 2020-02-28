Attribute VB_Name = "modTypes"
Public Enemy(1 To MAX_ENEMYS) As EnemyRec
Public Projectil(1 To 1) As ProjectilRec
Public Players(1 To 1) As PlayersRec


Private Type EnemyRec
    Dir As Byte
    Movetimer As Long
    Anim As Long
    HP As Long
    Target As Boolean
End Type

Private Type ProjectilRec
    Dir As Byte
    Active As Boolean
End Type

Private Type PlayersRec
    Char As Long
    Dir As Byte
End Type
