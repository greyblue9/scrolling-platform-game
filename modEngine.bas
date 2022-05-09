Attribute VB_Name = "modEngine"


Public Declare Function BitBlt Lib "gdi32" _
(ByVal hDestDC As Long, ByVal X As Long, _
ByVal Y As Long, ByVal nWidth As Long, _
ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc _
As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

'API for playing WAV sound effects.
Public Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwflags As Long) As Long
Public Result As Long


Public Const PINK As Long = 16711935
Public Const WHITE As Long = 16777215


Public Tiles(100, 3) 'in the tileset. (includes regular and sprites)
'(index, property)
'properties:
'0 = type (1[tile] or 2[sprite])
'1 = tile number
'2 = tilename
'3 = tiledesc
Public totalTiles As Integer 'in the tileset. (includes regular AND sprites)
Public firstSpriteIndex As Integer 'first tile in tileset that is a sprite.
Public tileIndex(32767) As Integer 'this stores tile indexes when you know the tile number(2000 -> 6,etc.)


Public Sprites(100, 7) As Integer
'0 = location (row)
'1 = location (col)
'2 = tile number
'3-7 = attributes
Public SpriteX(100) As Integer
Public SpriteY(100) As Integer

Public numOfSprites As Integer



Public Map(1000, 1000) As Integer
Public lvlWidth As Integer
Public lvlHeight As Integer












'------- Joystick -----------

Public myJoy As JOYINFOEX

Declare Function joyGetPosEx Lib "winmm.dll" (ByVal uJoyID As Long, pji As JOYINFOEX) As Long
Declare Function joyReleaseCapture Lib "winmm.dll" (ByVal id As Long) As Long
Declare Function joySetCapture Lib "winmm.dll" (ByVal hWnd As Long, ByVal uID As Long, ByVal uPeriod As Long, ByVal bChanged As Long) As Long
Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

Public Const JOYSTICKID1 = 0
Public Const JOYSTICKID2 = 1
Public Const JOY_POVCENTERED = -1
Public Const JOY_POVFORWARD = 0
Public Const JOY_POVRIGHT = 9000
Public Const JOY_POVLEFT = 27000
Public Const JOY_RETURNX = &H1&
Public Const JOY_RETURNY = &H2&
Public Const JOY_RETURNZ = &H4&
Public Const JOY_RETURNR = &H8&
Public Const JOY_RETURNU = &H10
Public Const JOY_RETURNV = &H20
Public Const JOY_RETURNPOV = &H40&
Public Const JOY_RETURNBUTTONS = &H80&
Public Const JOY_RETURNRAWDATA = &H100&
Public Const JOY_RETURNPOVCTS = &H200&
Public Const JOY_RETURNCENTERED = &H400&
Public Const JOY_USEDEADZONE = &H800&
Public Const JOY_RETURNALL = (JOY_RETURNX Or JOY_RETURNY Or JOY_RETURNZ Or JOY_RETURNR Or JOY_RETURNU Or JOY_RETURNV Or JOY_RETURNPOV Or JOY_RETURNBUTTONS)
Public Const JOY_CAL_READALWAYS = &H10000
Public Const JOY_CAL_READRONLY = &H2000000
Public Const JOY_CAL_READ3 = &H40000
Public Const JOY_CAL_READ4 = &H80000
Public Const JOY_CAL_READXONLY = &H100000
Public Const JOY_CAL_READYONLY = &H200000
Public Const JOY_CAL_READ5 = &H400000
Public Const JOY_CAL_READ6 = &H800000
Public Const JOY_CAL_READZONLY = &H1000000
Public Const JOY_CAL_READUONLY = &H4000000
Public Const JOY_CAL_READVONLY = &H8000000
Type JOYINFOEX
        dwSize As Long                 '  size of structure
        dwflags As Long                 '  flags to indicate what to return
        dwXpos As Long                '  x position
        dwYpos As Long                '  y position
        dwZpos As Long                '  z position
        dwRpos As Long                 '  rudder/4th axis position
        dwUpos As Long                 '  5th axis position
        dwVpos As Long                 '  6th axis position
        dwButtons As Long             '  button states
        dwButtonNumber As Long        '  current button number pressed
        dwPOV As Long                 '  point of view state
        dwReserved1 As Long                 '  reserved for communication between winmm driver
        dwReserved2 As Long                 '  reserved for future expansion
End Type

'-----------------------------
