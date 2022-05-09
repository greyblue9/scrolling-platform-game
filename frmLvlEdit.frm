VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmLvlEdit 
   AutoRedraw      =   -1  'True
   Caption         =   "Level Editor"
   ClientHeight    =   7815
   ClientLeft      =   165
   ClientTop       =   795
   ClientWidth     =   12375
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLvlEdit.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7815
   ScaleWidth      =   12375
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar scrH 
      Height          =   255
      Left            =   2880
      Max             =   20
      TabIndex        =   12
      Top             =   7320
      Width           =   9255
   End
   Begin VB.VScrollBar scrV 
      Height          =   7335
      Left            =   12120
      Max             =   20
      TabIndex        =   11
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox picSel 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FF00FF&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   120
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   7
      Top             =   120
      Width           =   495
   End
   Begin VB.VScrollBar scrTiles 
      Height          =   5055
      Left            =   2520
      TabIndex        =   5
      Top             =   720
      Width           =   255
   End
   Begin RichTextLib.RichTextBox rtfMain 
      Height          =   1815
      Left            =   4200
      TabIndex        =   2
      Top             =   1320
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   3201
      _Version        =   393217
      Enabled         =   0   'False
      TextRTF         =   $"frmLvlEdit.frx":0CCA
   End
   Begin VB.PictureBox picEdit 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   7335
      Left            =   2880
      MousePointer    =   2  'Cross
      ScaleHeight     =   489
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   617
      TabIndex        =   0
      Top             =   0
      Width           =   9255
      Begin VB.PictureBox picRealMasks 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   120
         ScaleHeight     =   41
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   657
         TabIndex        =   13
         Top             =   5520
         Visible         =   0   'False
         Width           =   9855
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   120
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.PictureBox picEndOfLevel 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   3120
         Picture         =   "frmLvlEdit.frx":0D45
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   10
         Top             =   3720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.PictureBox picRealTiles 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   120
         ScaleHeight     =   41
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   657
         TabIndex        =   9
         Top             =   4680
         Visible         =   0   'False
         Width           =   9855
      End
      Begin VB.Shape shTile 
         BorderColor     =   &H0000FF00&
         FillColor       =   &H00C00000&
         Height          =   495
         Left            =   2520
         Top             =   3720
         Width           =   495
      End
   End
   Begin VB.PictureBox picTiles 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   120
      MousePointer    =   10  'Up Arrow
      ScaleHeight     =   337
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   161
      TabIndex        =   1
      Top             =   720
      Width           =   2415
   End
   Begin VB.Label lblSel 
      BackColor       =   &H8000000E&
      Height          =   495
      Left            =   2040
      TabIndex        =   8
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblStatus 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000002&
      Caption         =   "Ready."
      ForeColor       =   &H80000009&
      Height          =   255
      Left            =   2880
      TabIndex        =   6
      Top             =   7560
      Width           =   9525
   End
   Begin VB.Label lblTileDesc 
      Caption         =   "Description. (0)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label lblTileName 
      Caption         =   "tilename"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label shapeBorder 
      Height          =   7815
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   12375
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu itemOpen 
         Caption         =   "&Open..."
      End
      Begin VB.Menu itmSaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu spacer 
         Caption         =   "-"
      End
      Begin VB.Menu itmExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuLevel 
      Caption         =   "&Level"
      Begin VB.Menu itmProperties 
         Caption         =   "P&roperties..."
      End
      Begin VB.Menu itmSpriteList 
         Caption         =   "&Sprite List"
      End
      Begin VB.Menu spacer2 
         Caption         =   "-"
      End
      Begin VB.Menu itmClear 
         Caption         =   "&Clear level"
      End
      Begin VB.Menu itmFill 
         Caption         =   "&Fill with tile"
      End
   End
   Begin VB.Menu mnuSprite 
      Caption         =   "Sprite"
      Visible         =   0   'False
      Begin VB.Menu itmSpriteName 
         Caption         =   "name of sprite (x, x)"
         Enabled         =   0   'False
      End
      Begin VB.Menu spacer3 
         Caption         =   "-"
      End
      Begin VB.Menu itmSpriteDelete 
         Caption         =   "&Delete"
      End
      Begin VB.Menu itmSpriteProperties 
         Caption         =   "&Properties..."
      End
   End
End
Attribute VB_Name = "frmLvlEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'For Resizing.
Dim hOffset As Integer
Dim vOffset As Integer

Dim borderHeight As Integer
Dim borderWidth As Integer

Dim tileBoxOffsetY As Integer


'TILESET
Dim tbTileHeight As Integer 'tile box is x # of tiles high
Dim tbTileWidth As Integer 'tile box is x # of tiles wide
Dim oldTBTW As Integer
Dim oldTBTH As Integer

Dim selTileCol As Byte
Dim selTileRow As Byte
Dim selTileIndex As Integer

Dim placementType As Integer '1 for regular tile, 2 for sprite


'EDITOR AREA
Dim eTileHeight As Integer 'editor is x # of tiles high
Dim eTileWidth As Integer 'editor is x # of tiles wide
Dim oldETH As Integer
Dim oldETW As Integer

Dim eTileCol As Byte 'editor tile col (from edge of screen)
Dim eTileRow As Byte 'editor tile row (from edge of screen)








Private Sub Form_Load()

Me.Refresh


lvlWidth = 20
lvlHeight = 15
hOffset = 0
vOffset = 0

For i = 0 To lvlWidth
For j = 0 To lvlHeight
    Map(i, j) = 0
Next j
Next i

numOfSprites = 0
selSpriteIndex = -1

borderWidth = Me.Width - shapeBorder.Width
borderHeight = Me.Height - shapeBorder.Height

Form_Resize 'do this to get tile dimensions



GetCrossRef



picEdit.BackColor = 16436640
DrawScreen

End Sub


Public Function DrawScreen()


picEdit.Cls

For row = 0 To eTileHeight
For col = 0 To eTileWidth
    
    If Not (row + vOffset) > lvlHeight - 1 And Not (col + hOffset) > lvlWidth - 1 Then
        If Not Map(row + vOffset, col + hOffset) = 0 Then
        BitBlt picEdit.hDC, col * 32, row * 32, 32, 32, picRealMasks.hDC, tileIndex(Map(row + vOffset, col + hOffset)) * 32, 0, vbSrcAnd
        BitBlt picEdit.hDC, col * 32, row * 32, 32, 32, picRealTiles.hDC, tileIndex(Map(row + vOffset, col + hOffset)) * 32, 0, vbSrcPaint
        End If
    Else
        BitBlt picEdit.hDC, col * 32, row * 32, 32, 32, picEndOfLevel.hDC, 0, 0, vbSrcCopy
    End If

Next col
Next row


For k = 0 To numOfSprites - 1
    If Sprites(k, 0) >= vOffset And Sprites(k, 0) <= vOffset + eTileHeight _
    And Sprites(k, 1) >= hOffset And Sprites(k, 1) <= hOffset + eTileWidth Then
        'oh snap, there's a sprite onscreen!

    BitBlt picEdit.hDC, (Sprites(k, 1) - hOffset) * 32, (Sprites(k, 0) - vOffset) * 32, 32, 32, picRealMasks.hDC, tileIndex(Sprites(k, 2)) * 32, 0, vbSrcAnd
    BitBlt picEdit.hDC, (Sprites(k, 1) - hOffset) * 32, (Sprites(k, 0) - vOffset) * 32, 32, 32, picRealTiles.hDC, tileIndex(Sprites(k, 2)) * 32, 0, vbSrcPaint
    End If
Next k


picEdit.Refresh


End Function

Public Function GetCrossRef()

totalTiles = 1

Tiles(0, 0) = 1
Tiles(0, 1) = 0
Tiles(0, 2) = "nothing"
Tiles(0, 3) = "empty space"

Dim tLine As Variant

'load reference file & split by lines
rtfMain.LoadFile App.Path & "\res\crossref.txt", rtfText
tLine = Split(rtfMain.Text, vbCrLf)


Dim tSpaced As Variant
Dim lineNum As Integer

Dim curTileType As Integer '1 for tiles, 2 for sprites
curTileType = 0 'not there yet

Dim tileDesc As String 'temporary variable used to build tile description


'process every line in file
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
                'MsgBox tSpaced(1) & " is number " & tSpaced(0) & ", of type " & curTileType & "."
                
                Tiles(totalTiles, 0) = curTileType
                Tiles(totalTiles, 1) = tSpaced(0)
                tileIndex(tSpaced(0)) = totalTiles
                
                Tiles(totalTiles, 2) = tSpaced(1)
                
                tileDesc = vbNullString
                For i = 2 To UBound(tSpaced)
                    tileDesc = tileDesc & tSpaced(i) & " "
                Next i
                Tiles(totalTiles, 3) = tileDesc
                
                'If curTileType = 1 Then
                    picSel.Picture = LoadPicture(App.Path & "\res\tiles\" & Tiles(totalTiles, 2) & ".bmp")
                    picSel.Refresh
                    BitBlt picRealTiles.hDC, totalTiles * 32, 0, 32, 32, picSel.hDC, 0, 0, vbSrcCopy
                'ElseIf curtiletype=2 then
                'If curTileType = 2 Then 'added
                '    lstSprites.AddItem Tiles(totalTiles, 2) & " (#" & Tiles(totalTiles, 1) & ")"
                'End If
                
                totalTiles = totalTiles + 1
                picRealTiles.Width = totalTiles * 64
                picRealMasks.Width = totalTiles * 64
                
            End If
            
        End If
    End If
    
WasComment:
    
Next lineNum
picRealTiles.Refresh



'make masks for transparency
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


Private Function PopulateTileBox()

picTiles.Cls

Dim thisRow As Integer
Dim thisCol As Integer
Dim thisIndex As Integer
thisRow = 0
thisCol = 0
thisIndex = 0

thisIndex = tileBoxOffsetY * tbTileWidth
Me.Caption = thisIndex

picRealTiles.Refresh


For i = 0 To totalTiles - 1
On Error Resume Next
    
    
    BitBlt picTiles.hDC, thisCol * 32, thisRow * 32, 32, 32, picRealTiles.hDC, thisIndex * 32, 0, vbSrcCopy
    
    
    thisCol = thisCol + 1
    thisIndex = thisIndex + 1
    If thisCol = tbTileWidth Then
        thisCol = 0
        thisRow = thisRow + 1
    End If



    
   ' MsgBox "OK!"
Next i

picTiles.Refresh

End Function


Private Sub Form_Resize()
On Error Resume Next

'getting border sizes and resizing controls
picEdit.Width = Me.Width - borderWidth - picEdit.Left - scrV.Width
picEdit.Height = Me.Height - borderHeight - picEdit.Top - lblStatus.Height - scrH.Height

scrV.Left = picEdit.Left + picEdit.Width
scrV.Height = picEdit.Height

scrH.Top = picEdit.Top + picEdit.Height
scrH.Width = picEdit.Width

lblStatus.Width = picEdit.Width + scrV.Width
lblStatus.Top = Me.Height - borderHeight - lblStatus.Height

picTiles.Height = Me.Height - borderHeight - 855 - 540
scrTiles.Height = picTiles.Height

'getting tile dimensions
tbTileWidth = CInt(picTiles.ScaleWidth \ 32)
tbTileHeight = CInt(picTiles.ScaleHeight \ 32)

eTileWidth = CInt(picEdit.ScaleWidth \ 32)
eTileHeight = CInt(picEdit.ScaleHeight \ 32)

If tbTileWidth < 1 Then tbTileWidth = 1
If tbTileHeight < 1 Then tbTileHeight = 1
If eTileWidth < 1 Then eTileWidth = 1
If eTileHeight < 1 Then eTileHeight = 1

UpdateStatusBar

If Not totalTiles = 0 Then
    If Not tbTileWidth = oldTBTW _
    Or Not tbTileHeight = oldTBTH Then
        PopulateTileBox
        oldTBTW = tbTileWidth
        oldTBTH = tbTileHeight
    End If
End If

ResizeScrollbars
DrawScreen

End Sub

Public Function ResizeScrollbars()
If (lvlWidth - eTileWidth) >= 0 Then
    scrH.Max = (lvlWidth - eTileWidth)
Else
    scrH.Max = 0
End If

If (lvlHeight - eTileHeight) >= 0 Then
    scrV.Max = (lvlHeight - eTileHeight)
Else
    scrV.Max = 0
End If

scrTiles.Max = Round((totalTiles / tbTileWidth) - tbTileHeight, 0) + 1
End Function

Private Function UpdateStatusBar()
    lblStatus.Caption = "Tile box is " & tbTileWidth & " x " & tbTileHeight _
    & " (" & tbTileWidth * tbTileHeight & " tiles).   " _
    & "Editor is " & eTileWidth & " x " & eTileHeight & " (" & eTileWidth * eTileHeight & " tiles)."
End Function




Private Sub Form_Unload(Cancel As Integer)
End

End Sub

Private Sub itemOpen_Click()
CommonDialog1.FileName = App.Path & "\res\levels\0"
CommonDialog1.ShowOpen

Dim tResult
tResult = MsgBox("Are you sure you want to open this level:" & vbCrLf & CommonDialog1.FileName, vbYesNoCancel, Me.Caption)

If Not tResult = 6 Then
    Exit Sub
End If

numOfSprites = 0

rtfMain.LoadFile CommonDialog1.FileName, rtfText
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
                    picEdit.BackColor = tSpaced(2)
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

ResizeScrollbars
DrawScreen



End Sub

Private Sub itmClear_Click()
Dim tResult
tResult = MsgBox("Are you sure you want to clear the map?", vbYesNoCancel, Me.Caption)

If Not tResult = 6 Then
    Exit Sub
End If

For row = 0 To lvlHeight
For col = 0 To lvlWidth
    Map(row, col) = 0
Next col
Next row
numOfSprites = 0

picEdit.BackColor = 16436640
DrawScreen
End Sub

Private Sub itmExit_Click()
Unload Me
End
End Sub

Private Sub itmFill_Click()
Dim tResult
tResult = MsgBox("Are you sure you want to fill the map with the selected tile?" & vbCrLf & vbCrLf & Tiles(selTileIndex, 3), vbYesNoCancel, Me.Caption)

If Not tResult = 6 Then
    Exit Sub
End If

For row = 0 To lvlHeight
For col = 0 To lvlWidth
    Map(row, col) = Tiles(selTileIndex, 1)
Next col
Next row
DrawScreen
End Sub

Private Sub itmProperties_Click()
frmProperties.Show
End Sub

Private Sub itmSaveAs_Click()
CommonDialog1.FileName = App.Path & "\res\levels\0"
CommonDialog1.ShowSave

If Not Mid(CommonDialog1.FileName, Len(CommonDialog1.FileName) - 3, 4) = ".txt" Then
    CommonDialog1.FileName = CommonDialog1.FileName & ".txt"
End If

Dim tResult
tResult = MsgBox("Are you sure you want to save this level to:" & vbCrLf & CommonDialog1.FileName, vbYesNoCancel, Me.Caption)

If Not tResult = 6 Then
    MsgBox "Did not save.", , Me.Caption
    Exit Sub
End If

Dim lvlData As String
lvlData = vbNullString

lvlData = lvlData & "# lvlwidth " & lvlWidth
lvlData = lvlData & vbCrLf & "# lvlheight " & lvlHeight
lvlData = lvlData & vbCrLf & "# bgcolor " & picEdit.BackColor
lvlData = lvlData & vbCrLf
lvlData = lvlData & vbCrLf & "# startleveldata"

lvlData = lvlData & vbCrLf

For row = 0 To lvlHeight - 1

    lblStatus.Caption = "Building level data: " & CInt((row / (lvlHeight - 1)) * 100) & "% done."
    lblStatus.Refresh

    For col = 0 To lvlWidth - 1
        lvlData = lvlData & Map(row, col) & " "
    Next col
    
    lvlData = lvlData & vbCrLf
    
Next row

lvlData = lvlData & "# endleveldata"
lvlData = lvlData & vbCrLf
lvlData = lvlData & vbCrLf & "# sprites"

lvlData = lvlData & vbCrLf

If numOfSprites > 0 Then
    For i = 0 To numOfSprites - 1
        lvlData = lvlData & Tiles(tileIndex(Sprites(i, 2)), 2) & ": "
        For j = 0 To 7
            lvlData = lvlData & Sprites(i, j) & " "
        Next j
        lvlData = lvlData & vbCrLf
    Next i
End If

    lblStatus.Caption = "Saving file..."
    lblStatus.Refresh

rtfMain.Text = lvlData
rtfMain.SaveFile CommonDialog1.FileName, rtfText

    lblStatus.Caption = "Successfully saved to " & CommonDialog1.FileName & "."
    lblStatus.Refresh

End Sub

Private Sub itmSpriteDelete_Click()
    DeleteSelectedSprite

    DrawScreen
End Sub



Private Sub itmSpriteList_Click()

If Not numOfSprites = 0 Then
    frmSpriteList.Show
End If

End Sub

Private Sub itmSpriteProperties_Click()

    frmSpriteList.Show
    
    frmSpriteList.lstSprites.ListIndex = selSpriteIndex

End Sub

Private Sub picEdit_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next

Select Case KeyCode

    Case vbKeyRight
        scrH.Value = scrH.Value + 1
        
    Case vbKeyLeft
        scrH.Value = scrH.Value - 1
        
    Case vbKeyUp
        scrV.Value = scrV.Value - 1

    Case vbKeyDown
        scrV.Value = scrV.Value + 1

End Select
End Sub


Private Sub picEdit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

        eTileCol = CByte((X - 16) / 32)
        eTileRow = CByte((Y - 16) / 32)


If placementType = 2 Then
    If Button = 1 Then
        'plopped a sprite!
        
        If X < 0 Or Y < 0 Then Exit Sub
        
        Sprites(numOfSprites, 0) = eTileRow + vOffset
        Sprites(numOfSprites, 1) = eTileCol + hOffset
        Sprites(numOfSprites, 2) = Tiles(selTileIndex, 1)
        
        'MsgBox Sprites(3, 2)
        
        numOfSprites = numOfSprites + 1
        DrawScreen
    
    ElseIf Button = 2 Then
        'right-click on sprite: spawn context menu
        
        'check for sprite here
        Dim foundSpriteHere As Boolean
        foundSpriteHere = False
        
        For i = 0 To numOfSprites - 1
            If Sprites(i, 0) = eTileRow + vOffset _
            And Sprites(i, 1) = eTileCol + hOffset Then
                'yes, there actually WAS a sprite here
                foundSpriteHere = True
                Exit For
            End If
        Next i
        
        If foundSpriteHere = True Then
            selSpriteIndex = i
            
            itmSpriteName.Caption = Tiles(tileIndex(Sprites(i, 2)), 2) & " at (" & eTileRow + vOffset & ", " & eTileCol + hOffset & ")"
            Call Me.PopupMenu(mnuSprite, , picEdit.Left + X * Screen.TwipsPerPixelX, picEdit.Top + Y * Screen.TwipsPerPixelY, itmSpriteProperties)
            
        Else
            selSpriteIndex = -1 'no sprite here
        End If
        
    End If
    
    
    Exit Sub
End If

picEdit_MouseMove Button, Shift, X, Y
End Sub


Private Sub picEdit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If X < 0 Or Y < 0 Then Exit Sub

eTileCol = CByte((X - 16) / 32)
eTileRow = CByte((Y - 16) / 32)

shTile.Left = eTileCol * 32
shTile.Top = eTileRow * 32


If Button = 1 Then

    If placementType = 1 Then
        'painted a tile.
        Map(eTileRow + vOffset, eTileCol + hOffset) = Tiles(selTileIndex, 1)
        DrawScreen
    End If
End If


If Button = 2 Then
    
    Map(eTileRow + vOffset, eTileCol + hOffset) = 0
    DrawScreen
End If


lblStatus.Caption = "[" & Tiles(tileIndex(Map(eTileRow + vOffset, eTileCol + hOffset)), 2) & "] at (" & eTileRow + vOffset & ", " & eTileCol + hOffset & ")"


End Sub

Private Sub picTiles_KeyDown(KeyCode As Integer, Shift As Integer)
picEdit_KeyDown KeyCode, Shift
End Sub

Private Sub picTiles_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

selTileCol = CByte((X - 16) / 32)
selTileRow = CByte((Y - 16) / 32)
selTileIndex = (selTileRow + tileBoxOffsetY) * tbTileWidth + selTileCol

If selTileIndex < firstSpriteIndex Then
    placementType = 1
    shTile.BorderColor = RGB(0, 255, 0)
Else
    placementType = 2
    shTile.BorderColor = RGB(255, 255, 255)
End If



ShowSelectedTile

End Sub

Private Function ShowSelectedTile()

If Not selTileIndex = 0 Then
    picSel.Picture = LoadPicture(App.Path & "\res\tiles\" & Tiles(selTileIndex, 2) & ".bmp")
Else
    picSel.Picture = Nothing
End If

lblSel = "Tile " & selTileRow & ", " & selTileCol & " (index: " & selTileIndex & ")"

lblTileName.Caption = Tiles(selTileIndex, 2)
lblTileDesc.Caption = Tiles(selTileIndex, 3) & " (#" & Tiles(selTileIndex, 1) & ")"
lblTileDesc.Caption = UCase(Mid(lblTileDesc.Caption, 1, 1)) & Mid(lblTileDesc.Caption, 2, Len(lblTileDesc.Caption) - 1) & "."

End Function


Private Sub scrH_Change()
hOffset = scrH.Value
DrawScreen
End Sub

Private Sub scrH_Scroll()
scrH_Change
End Sub

Private Sub scrTiles_Change()
tileBoxOffsetY = scrTiles.Value
PopulateTileBox
End Sub

Private Sub scrTiles_Scroll()
scrTiles_Change
End Sub

Private Sub scrV_Change()
vOffset = scrV.Value
DrawScreen
End Sub

Private Sub scrV_Scroll()
scrV_Change
End Sub
