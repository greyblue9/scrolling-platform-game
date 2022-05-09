VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmGame 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "SuperScroll Engine"
   ClientHeight    =   11565
   ClientLeft      =   120
   ClientTop       =   360
   ClientWidth     =   15750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11565
   ScaleWidth      =   15750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picBG 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   4320
      ScaleHeight     =   121
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   273
      TabIndex        =   6
      Top             =   5280
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.PictureBox picEndOfLevel 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   10080
      Picture         =   "frmGame.frx":0000
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   5
      Top             =   9480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picSel 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FF00FF&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   10080
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   3
      Top             =   8640
      Width           =   495
   End
   Begin VB.PictureBox picRealTiles 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   657
      TabIndex        =   2
      Top             =   8640
      Width           =   9855
   End
   Begin VB.PictureBox picRealMasks 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   657
      TabIndex        =   1
      Top             =   9480
      Width           =   9855
   End
   Begin VB.PictureBox picDisp 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   5055
      Left            =   0
      ScaleHeight     =   337
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   641
      TabIndex        =   0
      Top             =   0
      Width           =   9615
   End
   Begin RichTextLib.RichTextBox rtfMain 
      Height          =   1815
      Left            =   10800
      TabIndex        =   4
      Top             =   8280
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   3201
      _Version        =   393217
      Enabled         =   0   'False
      TextRTF         =   $"frmGame.frx":0C42
   End
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents Timer1 As ccrpTimer
Attribute Timer1.VB_VarHelpID = -1

Dim keyState(300) As Boolean

Private Const eTileWidth = 20
Private Const eTileHeight = 15
Private Const ev_JUMP = 5

Private Const act_DOWNPIPE = 10
Private Const act_DIE = 9

Private doingAction As Integer
Dim actI As Integer
Dim actJ As Integer
Dim actK As Integer
Dim actS As String

Private hOffset As Integer
Private vOffset As Integer

Dim sprite_hOffset(100) As Integer 'this is the sprite's X offset from its defined position in PIXELS.
Dim sprite_vOffset(100) As Integer 'this is the sprite's Y offset from its defined position in PIXELS.

'** Input **
Dim joyX As Integer 'from -1 to 1
Dim joyY As Integer 'from -1 to 1
'-- --- -- ---

'  _______________________________
' / New Camera Movement Variables \________________________
Dim absoluteX As Integer 'the absolute X pixel position of the screen starting from the left side of the level.
Dim absoluteY As Integer 'the absolute Y pixel position of the screen starting from the top of the level.
Dim absoluteTileX As Integer 'the absolute X, nearest-tile position of the screen.
Dim absoluteTileY As Integer 'the absolute y, nearest-tile position of the screen.
Dim offsetX As Integer 'the X offset of the screen in pixels from its nearest-tile position.
Dim offsetY As Integer 'the Y offset of the screen in pixels from its nearest-tile position.
Dim bufferSize As Integer 'the size of the buffer around the edges of the screen, in tiles.

'  ___________________
' / Display Variables \_______________
Dim animI As Integer '0 to 8
Dim animJ As Integer '0 or 1 (alternates with animI cycles.)
Dim lastDirn As Integer 'for determining what direction player should face

Dim oldAbsoluteX As Integer 'for determining which direction, if any, the screen moved.
Dim oldAbsoluteY As Integer
Dim BGXOffset As Integer

'  ___________________________
' / Player Movement Variables \____________________
Dim usingJoystick As Boolean
Dim velX As Single
Dim velY As Single
Dim accelY As Single
Dim decelX As Single
'variables for the player sprite
Dim spriteTileX
Dim spriteTileY
Dim collidedTILE_ID As Integer
Dim collidedTILE_Row As Integer
Dim collidedTILE_Col As Integer
Dim isRunning As Boolean

' ___________________
'/ Game Variables    \________
Dim Coins As Integer
Dim CoinsInLevel As Integer
Dim LevelNum As Integer
Dim YouWinTime As Integer




Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
SetVelocityOfPlayer
End Sub

Private Sub Form_Load()

Set Timer1 = New ccrpTimer
Timer1.EventType = TimerPeriodic
Timer1.Interval = 1000 / 60
Timer1.Stats.Frequency = 1
Timer1.Enabled = True


'------- Joystick -----------
Dim r&
Dim hWnd&
'Static TheX As Long
'Static TheY As Long
r& = joySetCapture(hWnd, JOYSTICKID1, 1, 0)
r& = joyReleaseCapture(JOYSTICKID1)
r& = joyGetPosEx(JOYSTICKID1, myJoy)
'----------------------------

numOfSprites = 0
GetCrossRef

LevelNum = 1
LoadLevel LevelNum


absoluteX = 0
absoluteY = 0
bufferSize = 2


picDisp.Width = eTileWidth * 32 * Screen.TwipsPerPixelX
picDisp.Height = eTileHeight * 32 * Screen.TwipsPerPixelY

decelX = 1
accelY = 0.5

usingJoystick = True


doingAction = 0

Me.Width = 640 * Screen.TwipsPerPixelX
Me.Height = 480 * Screen.TwipsPerPixelY
Me.Left = 0
Me.Top = 0


End Sub

Private Sub picDisp_KeyDown(KeyCode As Integer, Shift As Integer)

If usingJoystick = True Then usingJoystick = False

If Not keyState(KeyCode) = True Then
    UpdateKey KeyCode, True
End If

End Sub

Private Sub picDisp_KeyUp(KeyCode As Integer, Shift As Integer)
If Not keyState(KeyCode) = False Then
    UpdateKey KeyCode, False
End If
End Sub

Private Function UpdateKey(newKeyCode As Integer, newState As Boolean)

    keyState(newKeyCode) = newState
    
    If newKeyCode = vbKeyShift And newState = True Then
        JumpPlayer
    End If
    
    If newKeyCode = vbKeyDown And newState = True Then
        DownPlayer
    End If
    
    If newKeyCode = vbKeyControl Then
        isRunning = newState
    End If
    
    If newKeyCode = vbKeyEscape Then
        'ChangeRes 1680, 1050
        Unload Me
        End
    End If
    

End Function


Private Function GetCrossRef()

' **** define blank tile ***
Tiles(0, 0) = 1
Tiles(0, 1) = 0
totalTiles = 1


' **** load up cross-ref file and chop it up into lines
Dim tLine As Variant
rtfMain.LoadFile App.Path & "\res\crossref.txt", rtfText
tLine = Split(rtfMain.Text, vbCrLf)

Dim tSpaced As Variant
Dim lineNum As Integer


' **** first, we work on normal tiles. ****
Dim curTileType As Integer '1 for tiles, 2 for sprites
curTileType = 0 'not there yet


' **** process every line in file
For lineNum = 0 To UBound(tLine)
    'split current line by spaces
    tSpaced = Split(tLine(lineNum), " ")
    
    If UBound(tSpaced) >= 0 Then 'only if there IS a space in the line
        If tSpaced(0) = "'" Then GoTo WasComment
        
        If tSpaced(0) = "#" Then 'command
        
            Select Case tSpaced(1)
                Case "tiles"
                    curTileType = 1
                Case "sprites"
                    curTileType = 2
                    firstSpriteIndex = totalTiles
            End Select
            
        Else 'tile description line (not a command)
        
            If Not curTileType = 0 Then
            
                ' *** store tile information ***
                Tiles(totalTiles, 0) = curTileType
                Tiles(totalTiles, 1) = tSpaced(0)
                tileIndex(tSpaced(0)) = totalTiles
                
                Tiles(totalTiles, 2) = tSpaced(1)
                
                ' *** load tile image and copy to real tiles box ***
                picSel.Picture = LoadPicture(App.Path & "\res\tiles\" & Tiles(totalTiles, 2) & ".bmp")
                picSel.Refresh
                BitBlt picRealTiles.hDC, totalTiles * 32, 0, 32, 32, picSel.hDC, 0, 0, vbSrcCopy
                
                ' *** bump up total tiles and expand real tile boxes
                totalTiles = totalTiles + 1
                picRealTiles.Width = totalTiles * 64 * Screen.TwipsPerPixelX
                picRealMasks.Width = totalTiles * 64 * Screen.TwipsPerPixelX
                
            End If
            
        End If
    End If
    
WasComment:
    
Next lineNum
picRealTiles.Refresh


' **** make masks for transparency
For Y = 0 To 31
    For X = 0 To (totalTiles) * 32
        If picRealTiles.Point(X, Y) = PINK Then
            picRealMasks.PSet (X, Y), WHITE
            picRealTiles.PSet (X, Y), 0
        End If
    Next X
    
    For X = 0 To 31
        picRealMasks.PSet (X, Y), WHITE
    Next X
Next Y

picRealMasks.Refresh



End Function






Private Function LoadLevel(levelNumToLoad As Integer)
numOfSprites = 1 'first sprite is player!
CoinsInLevel = 0

rtfMain.LoadFile App.Path & "\res\levels\" & levelNumToLoad & ".txt", rtfText
Dim tLine As Variant
tLine = Split(rtfMain.Text, vbCrLf)

Dim tSpaced As Variant
Dim lineNum As Integer
Dim startOnLine As Integer
Dim curRow As Integer
Dim curCol As Integer
startOnLine = 0
curRow = 0
curCol = 0

Dim readMode As Integer
'0 = reading preliminary stuff
'1 = reading main map
'2 = reading sprites
readMode = 0

'process every line in file
For lineNum = 0 To UBound(tLine)
    'split current line by spaces
    tSpaced = Split(tLine(lineNum), " ")
    
    If UBound(tSpaced) >= 0 Then 'only if there IS a space in the line
        If tSpaced(0) = "'" Then GoTo WasComment
        
        If tSpaced(0) = "#" Then 'command
        
            Select Case tSpaced(1)
                Case "lvlwidth"
                    lvlWidth = tSpaced(2)
                Case "lvlheight"
                    lvlHeight = tSpaced(2)
                    
                Case "bgcolor"
                    picDisp.BackColor = tSpaced(2)
                Case "startleveldata"
                    startOnLine = lineNum + 1
                    readMode = 1
                Case "sprites"
                    readMode = 2
                    
            End Select
            
        Else 'tile description line (not a command)
        
                Select Case readMode
                
                    Case 1 'reading main map
                        
                        If Not tSpaced(0) = "#" Then
                            For i = 0 To UBound(tSpaced) - 1
                                Map(curRow, i) = tSpaced(i)
                                
                                
                                
                                If Map(curRow, i) = 9000 And Not doingAction = act_DOWNPIPE Then 'starting position for player
                                    Sprites(0, 0) = curRow
                                    Sprites(0, 1) = i
                                    Sprites(0, 2) = 1000
                                    Map(curRow, i) = 0
                                End If
                                
                                If Map(curRow, i) = 5 Then 'a coin
                                    CoinsInLevel = CoinsInLevel + 1
                                End If
                                
                                
                                
                            Next i
                            curRow = curRow + 1
                        Else
                            MsgBox tSpaced(1)
                        End If
                    
                    Case 2 'reading sprites
                    
                        For i = 0 To 7
                            If Not tSpaced(i + 1) = vbNullString Then
                                Sprites(numOfSprites, i) = tSpaced(i + 1)
                            End If
                        Next i
                        numOfSprites = numOfSprites + 1
                    
                End Select
            
        End If
    End If
    
WasComment:
    
Next lineNum


For i = 0 To numOfSprites - 1
    If i = 0 And doingAction = act_DOWNPIPE Then GoTo 10
    SpriteX(i) = Sprites(i, 1) * 32
    SpriteY(i) = Sprites(i, 0) * 32
10 Next i

picBG.Picture = LoadPicture(App.Path & "\res\backgrounds\1.bmp")
oldAbsoluteX = absoluteX
oldAbsoluteY = absoluteY
BGXOffset = 0

velX = 0
velY = 0

animI = 0
animJ = 0
lastDirn = 0

Coins = 0
YouWinTime = -1

End Function






Public Function DrawScreen()




''  _______________________________
'' / New Camera Movement Variables \________________________
'Dim absoluteX As Long 'the absolute X pixel position of the screen starting from the left side of the level.
'Dim absoluteY As Long 'the absolute Y pixel position of the screen starting from the top of the level.
'Dim absoluteTileX As Long 'the absolute X, nearest-tile position of the screen.
'Dim absoluteTileY As Long 'the absolute y, nearest-tile position of the screen.
'Dim offsetX As Long 'the X offset of the screen in pixels from its nearest-tile position.
'Dim offsetY As Long 'the Y offset of the screen in pixels from its nearest-tile position.


absoluteTileX = Int(absoluteX / 32)
absoluteTileY = Int(absoluteY / 32)
offsetX = (absoluteTileX * 32) - absoluteX
offsetY = (absoluteTileY * 32) - absoluteY


picDisp.Cls
'BitBlt picDisp.hDC, 0, 0, picDisp.ScaleWidth, picDisp.ScaleHeight, picBG.hDC, BGXOffset, 0, vbSrcCopy
'If BGXOffset > 640 Then
'    BitBlt picDisp.hDC, picDisp.ScaleWidth - (BGXOffset - 640), 0, 640, picDisp.ScaleHeight, picBG.hDC, 0, 0, vbSrcCopy
'ElseIf BGXOffset < 0 Then
'    BitBlt picDisp.hDC, 0, 0, -BGXOffset, picDisp.ScaleHeight, picBG.hDC, 1280 + BGXOffset, 0, vbSrcCopy
'End If



If Not doingAction = act_DOWNPIPE Then
    For row = 0 To eTileHeight
    For col = 0 To eTileWidth
    
        On Error Resume Next
            BitBlt picDisp.hDC, col * 32 + offsetX, row * 32 + offsetY, 32, 32, picRealMasks.hDC, tileIndex(Map(row + absoluteTileY, col + absoluteTileX)) * 32, 0, vbSrcAnd
            BitBlt picDisp.hDC, col * 32 + offsetX, row * 32 + offsetY, 32, 32, picRealTiles.hDC, tileIndex(Map(row + absoluteTileY, col + absoluteTileX)) * 32, 0, vbSrcPaint
        
    Next col
    Next row
End If

For k = 0 To numOfSprites - 1
    If SpriteX(k) >= absoluteX - 32 And SpriteX(k) <= absoluteX + (eTileWidth * 32) _
    And SpriteY(k) >= absoluteY - 32 And SpriteY(k) <= absoluteY + (eTileHeight * 32) Then
        'oh snap, there's a sprite onscreen!
        
    If Not k = 0 Then 'sprite other than the player
        If Not Sprites(k, 2) >= 9000 Then
            BitBlt picDisp.hDC, SpriteX(k) - absoluteX, SpriteY(k) - absoluteY, 32, 32, picRealMasks.hDC, tileIndex(Sprites(k, 2)) * 32, 0, vbSrcAnd
            BitBlt picDisp.hDC, SpriteX(k) - absoluteX, SpriteY(k) - absoluteY, 32, 32, picRealTiles.hDC, tileIndex(Sprites(k, 2)) * 32, 0, vbSrcPaint
        End If
    Else 'the player
    
        Dim IDtoUSE As Integer
        '1001 player-r1 [the player]
        '1002 player-r2 [the player]
        '1003 player-l1 [the player]
        '1004 player-l2 [the player]
        '1005 player-jumpr [the player]
        '1006 player-jumpl [the player]
        '1007 player-stand [the player]
        If velY = 0 Then
            If velX > 0 Then
                IDtoUSE = 1001 + animJ
            ElseIf velX < 0 Then
                IDtoUSE = 1003 + animJ
            Else
                IDtoUSE = 1007 + lastDirn
            End If
        Else
            If velX > 0 Then
                IDtoUSE = 1005
            ElseIf velX < 0 Then
                IDtoUSE = 1006
            Else
                IDtoUSE = 1005 + lastDirn
            End If
        End If
        
        BitBlt picDisp.hDC, SpriteX(k) - absoluteX, SpriteY(k) - absoluteY, 32, 32, picRealMasks.hDC, tileIndex(IDtoUSE) * 32, 0, vbSrcAnd
        BitBlt picDisp.hDC, SpriteX(k) - absoluteX, SpriteY(k) - absoluteY, 32, 32, picRealTiles.hDC, tileIndex(IDtoUSE) * 32, 0, vbSrcPaint
        
    End If
    
    End If
Next k

If doingAction = act_DOWNPIPE Then
    For row = 0 To eTileHeight
    For col = 0 To eTileWidth
    
        On Error Resume Next
            BitBlt picDisp.hDC, col * 32 + offsetX, row * 32 + offsetY, 32, 32, picRealMasks.hDC, tileIndex(Map(row + absoluteTileY, col + absoluteTileX)) * 32, 0, vbSrcAnd
            BitBlt picDisp.hDC, col * 32 + offsetX, row * 32 + offsetY, 32, 32, picRealTiles.hDC, tileIndex(Map(row + absoluteTileY, col + absoluteTileX)) * 32, 0, vbSrcPaint
        
    Next col
    Next row
End If


picDisp.Refresh


End Function

Private Sub Timer1_Timer(ByVal Milliseconds As Long)


If Not doingAction = 0 Then
'something's going on, like an animation ... process it and ignore the rest of this function for now.
    
    Select Case doingAction
    
        Case act_DOWNPIPE
            Select Case actJ 'stage in the pipe-decending process
                Case 0
                    actI = actI + 1
                    SpriteY(0) = SpriteY(0) + 1
                    If actI = 32 Then
                        actJ = 1
                    End If
                    
                Case 1
                    SpriteX(0) = Sprites(actK, 5) * 32
                    SpriteY(0) = Sprites(actK, 4) * 32
                    
                    If Not Sprites(actK, 3) = 0 Then 'exit to this same level.
                        LevelNum = Sprites(actK, 3)
                        LoadLevel Sprites(actK, 3)
                    End If
                    
                    doingAction = 0
            End Select
        
        Case act_DIE
            actI = actI + 1
                If actI = 220 Then
                    doingAction = 0
                    
                    LoadLevel LevelNum
                End If
        
    End Select
    
    DrawScreen
    
    Exit Sub
End If




ReturnJoystickDir

If Not YouWinTime = -1 Then
    YouWinTime = YouWinTime - 1
    
    If YouWinTime = -1 Then
        LevelNum = LevelNum + 1
        LoadLevel CStr(LevelNum)
    End If
    
    Exit Sub
End If

If velX > 0 Then
    velX = velX - decelX
ElseIf velX < 0 Then
    velX = velX + decelX
End If


animI = animI + 1
If animI = 8 + isRunning * 2 Then
    animI = 0
    animJ = 1 - animJ
ElseIf animI > 8 + isRunning * 2 Then animI = 0
End If

'set velocity of player according to keys pressed
SetVelocityOfPlayer

'move player with current velocities
SpriteX(0) = SpriteX(0) + velX
SpriteY(0) = SpriteY(0) + velY

If Not velX = 0 Then
    If velX > 0 Then
        lastDirn = 0
    Else
        lastDirn = 1
    End If
End If

Me.Caption = SpriteY(0) & " of " & (lvlHeight * 32 + 32)
If SpriteY(0) > lvlHeight * 32 + 32 Then
    'the player has died
    PlaySFX "die"
    
    doingAction = act_DIE
    actI = 0
    actJ = 0
End If


'simulated gravity
'check if player is not standing on the ground first.
'  do this by first moving the player down a pixel then checking for a collision.
SpriteY(0) = SpriteY(0) + 1
If Not CheckCOLLISION(6) = 1 _
And Not CheckCOLLISION(7) = 1 _
And Not CheckCOLLISION(8) = 1 Then
    If Not velY > 30 Then
    velY = velY + accelY
    End If
'  then move him back up.
End If
SpriteY(0) = SpriteY(0) - 1


' _____________________
'/ collision detection \___________
' 0 1 2
' 3 4 5
' 6 7 8

collidedTileID = 0 'this gets set in the CheckCOLLISION function.

Dim a As Byte
For a = 0 To 8 'look for collision with special tiles like coins, etc.
    If CheckCOLLISION(a) = 2 Then
        Select Case collidedTILE_ID
            Case 5, 6, 7
                Map(collidedTILE_Row, collidedTILE_Col) = 0
                
                If collidedTILE_ID = 7 Then
                    PlaySFX "redcoin"
                Else
                    PlaySFX "coin"
                End If
                Coins = Coins + 1
                
                If Coins = CoinsInLevel Then
                    YouWinTime = 200
                    
                    For row = 0 To lvlHeight - 1
                    For col = 0 To lvlWidth - 1
                        Map(row, col) = 0
                    Next col
                    Next row
                End If
        End Select
    End If
Next a

'normal wall-style collisions
If velX > 0 Then 'if moving to the right
    If CheckCOLLISION(2) = 1 _
    Or CheckCOLLISION(5) = 1 _
    Or CheckCOLLISION(8) = 1 Then 'if we collided with an wall to the right.
        'stop moving horizontally
        velX = 0
        'lock to correct position (clear of wall)
        SpriteX(0) = Round(SpriteX(0) / 32) * 32
    End If
    
ElseIf velX < 0 Then 'if moving to the left
    If CheckCOLLISION(0) = 1 _
    Or CheckCOLLISION(3) = 1 _
    Or CheckCOLLISION(6) = 1 Then 'if we collided with an wall to the left.
        'stop moving horizontally
        velX = 0
        'lock to correct position (clear of wall)
        SpriteX(0) = Round(SpriteX(0) / 32) * 32
    End If
End If

If velY > 0 Then 'if moving down
    If CheckCOLLISION(6) = 1 _
    Or CheckCOLLISION(7) = 1 _
    Or CheckCOLLISION(8) = 1 Then 'if we collided with an wall below us.
        'stop moving vertically
        velY = 0
        'lock to correct position (clear of wall)
        SpriteY(0) = Round(SpriteY(0) / 32) * 32
    End If
    If velY > 10 Then
    
        If CheckCOLLISION(3) = 1 _
        Or CheckCOLLISION(4) = 1 _
        Or CheckCOLLISION(5) = 1 Then 'if we collided with an wall on us! (due to high speed)
            'stop moving vertically
            velY = 0
            'lock to correct position (clear of wall)
            SpriteY(0) = Round(SpriteY(0) / 32) * 32 - 32
        End If
    
    End If
    
ElseIf velY < 0 Then 'if moving up
    If CheckCOLLISION(0) = 1 _
    Or CheckCOLLISION(1) = 1 _
    Or CheckCOLLISION(2) = 1 Then 'if we collided with an wall above us.
        'stop moving vertically!
        velY = 0
        'lock to correct position (clear of wall)
        SpriteY(0) = Round(SpriteY(0) / 32) * 32
        
        Select Case collidedTILE_ID
            Case 17
                Map(collidedTILE_Row, collidedTILE_Col) = 0
                PlaySFX "smash"
            Case 20
                Map(collidedTILE_Row, collidedTILE_Col) = 21
                PlaySFX "coinbump"
                Coins = Coins + 1
            Case Else
                PlaySFX "bump"
        End Select
    End If
End If

'___________________________________/

'move camera position to follow player
If SpriteX(0) - absoluteX >= (32 * 10) Then
    absoluteX = SpriteX(0) - (32 * 10)
    
    'rightmost level boundary
    If absoluteX + 32 * eTileWidth >= lvlWidth * 32 Then
        absoluteX = (lvlWidth - eTileWidth) * 32
    End If
    
ElseIf SpriteX(0) - absoluteX <= (32 * 8) And absoluteX > 0 Then
    absoluteX = SpriteX(0) - (32 * 8)

    If absoluteX < 0 Then
        absoluteX = 0
    End If

End If

If SpriteY(0) - absoluteY >= (32 * 10) Then
    absoluteY = SpriteY(0) - (32 * 10)
    
    'bottom level boundary
    If absoluteY + 32 * eTileHeight >= lvlHeight * 32 Then
        absoluteY = (lvlHeight - eTileHeight) * 32
    End If
    
ElseIf SpriteY(0) - absoluteY <= (32 * 4) Then
    absoluteY = SpriteY(0) - (32 * 4)
    
    If absoluteY < 0 Then
        absoluteY = 0
    End If
    
End If

If absoluteX > oldAbsoluteX Then
    BGXOffset = BGXOffset + 2
ElseIf absoluteX < oldAbsoluteX Then
    BGXOffset = BGXOffset - 2
End If

If BGXOffset = 1280 Then BGXOffset = 0
If BGXOffset = -640 Then BGXOffset = 640

oldAbsoluteX = absoluteX
oldAbsoluteY = absoluteY

DrawScreen

End Sub

Private Function PlaySFX(fileTitle As String)
Result = PlaySound(App.Path & "\res\sound\" & fileTitle & ".wav", 1, 1)
End Function

Private Function CheckCOLLISION(tilePosition As Byte) As Integer

' this function checks whether the player is colliding with a tile in the specified "player-surroundings" position.
' the positions are as follows:
' 0 1 2
' 3 4 5
' 6 7 8
' 4 being the player's nearest, closest tile.

'translate the given tile position into -1 to 1 tile offsets.
Dim posXOffset As Integer
Dim posYOffset As Integer

Select Case tilePosition
    Case 0
        posXOffset = -1
        posYOffset = -1
    Case 1
        posXOffset = 0
        posYOffset = -1
    Case 2
        posXOffset = 1
        posYOffset = -1
    Case 3
        posXOffset = -1
        posYOffset = 0
    Case 4
        posXOffset = 0
        posYOffset = 0
    Case 5
        posXOffset = 1
        posYOffset = 0
    Case 6
        posXOffset = -1
        posYOffset = 1
    Case 7
        posXOffset = 0
        posYOffset = 1
    Case 8
        posXOffset = 1
        posYOffset = 1
End Select

'get current, closest tile (tile in 4 position).
Dim playerTileX As Integer
Dim playerTileY As Integer
playerTileX = Round(SpriteX(0) / 32, 0)
playerTileY = Round(SpriteY(0) / 32, 0)
'SpriteX(1) = playerTileX * 32
'SpriteY(1) = playerTileY * 32

If (playerTileX + posXOffset) * 32 + 32 - SpriteX(0) < 64 _
And (playerTileX + posXOffset) * 32 + 32 - SpriteX(0) > 0 _
And (playerTileY + posYOffset) * 32 + 32 - SpriteY(0) < 64 _
And (playerTileY + posYOffset) * 32 + 32 - SpriteY(0) > 0 Then

    On Error Resume Next
    If Not Map(playerTileY + posYOffset, playerTileX + posXOffset) = 0 Then
    
        collidedTILE_ID = Map(playerTileY + posYOffset, playerTileX + posXOffset)
        
        If collidedTILE_ID < 10 Then
            'colliding with special tile
            CheckCOLLISION = 2
        ElseIf (collidedTILE_ID >= 10 And collidedTILE_ID <= 299) Or collidedTILE_ID = 542 Or collidedTILE_ID = 543 Then
            'colliding with a wall.
            CheckCOLLISION = 1
        End If
        
        collidedTILE_Row = playerTileY + posYOffset
        collidedTILE_Col = playerTileX + posXOffset
        
    End If

    
End If

End Function





Private Function ReturnJoystickDir() As Byte

If usingJoystick = True Then

    myJoy.dwSize = 64
    myJoy.dwflags = JOY_RETURNALL
    r& = joyGetPosEx(JOYSTICKID1, myJoy)
    
    Select Case myJoy.dwXpos
    
        Case 255
            keyState(vbKeyLeft) = True
            keyState(vbKeyRight) = False
            
        Case 32767
            keyState(vbKeyLeft) = False
            keyState(vbKeyRight) = False
            
        Case 65274
            keyState(vbKeyLeft) = False
            keyState(vbKeyRight) = True
    
    End Select
    
    Select Case myJoy.dwYpos
    
        Case 255
            keyState(vbKeyUp) = True
            keyState(vbKeyDown) = False
            
        Case 32767
            keyState(vbKeyUp) = False
            keyState(vbKeyDown) = False
            
        Case 65274
            keyState(vbKeyUp) = False
            If Not keyState(vbKeyDown) = True Then
                keyState(vbKeyDown) = True
                DownPlayer
            End If
    
    End Select
    
    If Not (myJoy.dwButtons And 2) = 0 Then
        JumpPlayer
    End If
    
    If Not (myJoy.dwButtons And 1) = 0 Then
        isRunning = True
    Else
        isRunning = False
    End If

End If

End Function

Private Function SetVelocityOfPlayer()


If isRunning = False Then

    If keyState(vbKeyRight) = True Then
        velX = 4
    ElseIf keyState(vbKeyLeft) = True Then
        velX = -4
    Else
        velX = 0
    End If
    
Else

    If keyState(vbKeyRight) = True Then
        velX = 8
    ElseIf keyState(vbKeyLeft) = True Then
        velX = -8
    End If

End If

End Function

Private Function JumpPlayer()

'player is trying to jump
'check if he is standing on the ground first,
'  do this by first moving the player down a pixel then checking for a collision.
SpriteY(0) = SpriteY(0) + 1
If CheckCOLLISION(6) = 1 _
Or CheckCOLLISION(7) = 1 _
Or CheckCOLLISION(8) = 1 Then
    velY = -12
    PlaySFX "jump"
End If
'  then move him back up.
SpriteY(0) = SpriteY(0) - 1

End Function

Private Function DownPlayer()

'player is trying to go down
'check if he is standing on a possible exit object like a pipe first,
'  do this by first moving the player down a pixel then checking for a collision.
SpriteY(0) = SpriteY(0) + 1
If CheckCOLLISION(7) = 1 Then

    Select Case collidedTILE_ID
        Case 110 '"down" pipe
            'check for pipe exit.
            
            For i = 0 To UBound(Sprites)
                If Sprites(i, 0) = Round(SpriteY(0) / 32, 0) + 1 _
                And Sprites(i, 1) = Round(SpriteX(0) / 32, 0) _
                And Sprites(i, 2) = 9500 Then
                
                    'there's an exit on this pipe, below where I'm standing.
                
                    PlaySFX "pipe"
                    
                    doingAction = act_DOWNPIPE
                    SpriteX(0) = Round(SpriteX(0) / 32, 0) * 32
                    actI = 0
                    actJ = 0
                    actK = i
                    
                End If
            Next i
            

    End Select
End If

'  then move him back up.
SpriteY(0) = SpriteY(0) - 1

End Function
