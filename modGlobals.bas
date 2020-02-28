Attribute VB_Name = "modGlobals"
'Variables de las Flechas
Public DirUp As Boolean
Public DirDown As Boolean
Public DirRight As Boolean
Public DirLeft As Boolean

Public AntiFly As Boolean

Public Const MAX_ENEMYS = 1

Public TimerExplosion As Long
Public ProjectilExplosion As Boolean


Public HP As Long

Public InMove As Boolean
Public ShiftDown As Boolean
Public AttackDown As Boolean
Public inGround As Boolean

Public ProjectilLaunch As Boolean

Public Anim As Long

Public PosX As Long
Public PosY As Long
Public ProX As Long
Public ProY As Long

Public MsgY As Long
Public MsgX As Long

Public EnmY As Long
Public EnmX As Long

Public Jumping As Boolean
Public Gravity As Long
Public InAir As Boolean

Public CanMove As Boolean
Public Puntos As Long
Public GameState As Byte

Public TileX As Long
Public Lado As Boolean

Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Public PointTimer As Long
Public ShiftUp As Boolean

Public Const WalkSpeed As Byte = 45
Public Const RunSpeed As Byte = 65

'Direcciones del Personaje
Public PlayerDirection As Byte

'Public Const DIR_DOWN As Byte = 0
'Public Const DIR_UP As Byte = 1
Public Const DIR_RIGHT As Byte = 1
Public Const DIR_LEFT As Byte = 2

Public Const PWIDTH As Long = 16
Public Const PHEIGHT As Long = 16

Public Enum MENUType
    Main_Menu = 1
    Main_Game
    Main_Play
    Main_Result
    Main_Count 'El contador de Menus
End Enum
