VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cascade 2003"
   ClientHeight    =   2235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5925
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   149
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   395
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picResetM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4020
      Picture         =   "Form1.frx":1042
      ScaleHeight     =   19
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   17
      TabIndex        =   12
      Top             =   540
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picReset 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3720
      Picture         =   "Form1.frx":1460
      ScaleHeight     =   19
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   17
      TabIndex        =   11
      Top             =   540
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picBonusM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1065
      Left            =   180
      Picture         =   "Form1.frx":187E
      ScaleHeight     =   71
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   71
      TabIndex        =   10
      Top             =   180
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.PictureBox picBonus 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1065
      Left            =   120
      Picture         =   "Form1.frx":54A8
      ScaleHeight     =   71
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   71
      TabIndex        =   9
      Top             =   120
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.TextBox txtName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1740
      TabIndex        =   8
      Top             =   1920
      Width           =   2235
   End
   Begin VB.Timer timHiScores 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   5700
      Top             =   1260
   End
   Begin VB.PictureBox picLogoMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1125
      Left            =   900
      Picture         =   "Form1.frx":90D2
      ScaleHeight     =   75
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   237
      TabIndex        =   7
      Top             =   3300
      Visible         =   0   'False
      Width           =   3555
   End
   Begin VB.PictureBox picLogo 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1125
      Left            =   900
      Picture         =   "Form1.frx":161AC
      ScaleHeight     =   75
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   237
      TabIndex        =   6
      Top             =   4680
      Visible         =   0   'False
      Width           =   3555
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   2100
      Top             =   1020
   End
   Begin VB.PictureBox picBlank 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1155
      Left            =   2520
      ScaleHeight     =   77
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   57
      TabIndex        =   5
      Top             =   900
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox picBuffer 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1155
      Left            =   1020
      ScaleHeight     =   77
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   145
      TabIndex        =   4
      Top             =   420
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.PictureBox picBaseCol 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   3
      Left            =   4740
      Picture         =   "Form1.frx":23286
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   3
      Top             =   60
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox picBaseCol 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   2
      Left            =   4500
      Picture         =   "Form1.frx":234D0
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   2
      Top             =   60
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox picBaseCol 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00008080&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   1
      Left            =   4260
      Picture         =   "Form1.frx":2371A
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   1
      Top             =   60
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox picBaseCol 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000080&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   0
      Left            =   4020
      Picture         =   "Form1.frx":23964
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   0
      Top             =   60
      Visible         =   0   'False
      Width           =   195
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private GameGrid(9, 25) As Long
Private bInGame As Boolean
Private CW As Long
Private CH As Long
Private Const xOffset As Long = 5
Private Const yOffset As Long = 33
Private lCurrentScore As Long
Private bProcessing As Boolean
Dim bEnteringHiScore As Boolean
Private lDisplayingHiScoreXPos As Integer
Private UserPressedReset As Boolean
Private sCurrentHI As String
Private lTheta As Long
Private lAlt As Long
Private lSize As Long
Private lperspective As Long
Private Const bStars As Boolean = False
Sub ShuffleFillGrid()
Dim X As Long, y As Long
    For y = 0 To 8
        For X = 0 To 24
            GameGrid(y, X) = Int(Rnd * 4) + 1
        Next
    Next
    'for x=0 to 8:for y=0 to 24:? gamegrid(x,y);:next:? vbcrlf:next 'debug.print
End Sub
Function EstablishCellFromMousePos(xIn As Long, yIn As Long) As Long
'use mod and div when extracting xy co'ords
Dim xPos As Long
Dim yPos As Long
    If xIn < xOffset Or xIn >= (picBuffer.Width - xOffset) _
    Or yIn < yOffset Or yIn >= (yOffset + ((CH + 1) * 9)) Then
        EstablishCellFromMousePos = 0
        Exit Function
    End If
    xPos = 1 + ((xIn - xOffset) \ (CW + 1))
    yPos = 1 + ((yIn - yOffset) \ (CW + 1))
    
    EstablishCellFromMousePos = xPos + (100 * yPos)

End Function
Sub DrawGridToBuffer()
Dim X As Long, y As Long


    CW = picBaseCol(0).Width
    CH = picBaseCol(0).Height
    
    For y = 0 To 8
        For X = 0 To 24
            If GameGrid(y, X) > 0 Then
                BitBlt picBuffer.hdc, _
                       (X * (CW + 1)) + xOffset, (y * (CH + 1)) + yOffset, _
                       CW, CH, _
                       picBaseCol(GameGrid(y, X) - 1).hdc, 0, 0, vbSrcCopy
            End If
        Next
    Next
End Sub



Private Sub Form_Load()
    Randomize Timer
    
    CheckForHiScores
    
    Me.Width = ((3 * xOffset) + ((picBaseCol(0).Width + 1) * 25)) * Screen.TwipsPerPixelX
    Me.Height = ((2 * yOffset) + ((picBaseCol(0).Height + 1) * 9) - 3) * Screen.TwipsPerPixelX
    picBuffer.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
    picBlank.Move 0, 0, picBuffer.Width, picBuffer.Height
    
    SetupStarField
    
    FreshDisplay
    
End Sub
Sub FreshDisplay()
    ShuffleFillGrid
    
    DrawGridToBuffer
    
    ApplyBlend
    
    DrawLogoToBuffer
    
    'pTextOut "Hi Score Line", "Arial", 10, True, 124, picBuffer.Height - 60, vbWhite
    txtName.Visible = False
    bEnteringHiScore = False
    ResetScoreDisplay
    
    BufferToScreen
    
    bInGame = False
    
    lTheta = 0
    lAlt = 0
    lSize = 1000
    lperspective = 10000
    
    timHiScores.Enabled = True
End Sub
Sub ResetScoreDisplay()
    lDisplayingHiScoreXPos = picBuffer.Width + 5
End Sub
Sub CheckForHiScores()
Dim X As Integer
    If Not FileExists(HiScoreFile) Then
        For X = 1 To 10
            writeINI "SETTINGS", "HISCORE" & X, 1000, HiScoreFile
            writeINI "SETTINGS", "HINAME" & X, "Cascade 2003", HiScoreFile
        Next
    End If
End Sub
Function IsTheGameOver() As Boolean
Dim X As Long
Dim y As Long

    IsTheGameOver = True
    For y = 1 To 9
    For X = 1 To 25
        If GameGrid(y - 1, X - 1) > 0 And _
           HowManySurroundingSameColor(X + (100 * y)) > 0 Then
            IsTheGameOver = False
            Exit Function
        End If
    Next
    Next
End Function

Sub NewGame()
    ShuffleFillGrid
    
    DrawGridToBuffer
        
    BufferToScreen
    
    lCurrentScore = 0
    
    bInGame = True
    bProcessing = False
    UserPressedReset = False
    
    sCurrentHI = "High Score: " & ReadINI("SETTINGS", "HINAME1", HiScoreFile) & _
                    " - " & ReadINI("SETTINGS", "HISCORE1", HiScoreFile)
    
    Timer1.Enabled = True
    
End Sub
Sub DrawLogoToBuffer()
    BitBlt picBuffer.hdc, (picBuffer.Width \ 2) - (picLogo.Width \ 2), (picBuffer.Height \ 2) - (picLogo.Height \ 2) - 40, picLogo.Width, picLogo.Height, picLogoMask.hdc, 0, 0, vbSrcAnd
    BitBlt picBuffer.hdc, (picBuffer.Width \ 2) - (picLogo.Width \ 2), (picBuffer.Height \ 2) - (picLogo.Height \ 2) - 40, picLogo.Width, picLogo.Height, picLogo.hdc, 0, 0, vbSrcPaint
End Sub
Sub DrawResetButtonToBuffer()
    BitBlt picBuffer.hdc, picBuffer.Width - picReset.Width - 2, 2, picReset.Width, picReset.Height, picResetM.hdc, 0, 0, vbSrcAnd
    BitBlt picBuffer.hdc, picBuffer.Width - picReset.Width - 2, 2, picReset.Width, picReset.Height, picReset.hdc, 0, 0, vbSrcPaint
End Sub
Sub DrawBonusToBuffer()
    BitBlt picBuffer.hdc, 5, 5, picBonus.Width, picBonus.Height, picBonusM.hdc, 0, 0, vbSrcAnd
    BitBlt picBuffer.hdc, 5, 5, picBonus.Width, picBonus.Height, picBonus.hdc, 0, 0, vbSrcPaint
End Sub
Sub ClearBuffer()
    BitBlt picBuffer.hdc, 0, 0, picBuffer.Width, picBuffer.Height, picBlank.hdc, 0, 0, vbSrcCopy
End Sub

Sub BufferToScreen()
    BitBlt Me.hdc, 0, 0, picBuffer.Width, picBuffer.Height, picBuffer.hdc, 0, 0, vbSrcCopy
    Me.Refresh
End Sub
Sub pTextOut(sIn As String, sFont As String, iFontSize As Integer, bFontBold As Boolean, xPos As Integer, yPos As Integer, lColor As Long)
    
    SetTextColor picBuffer.hdc, 0
    picBuffer.Font = sFont
    picBuffer.FontSize = iFontSize
    picBuffer.FontBold = bFontBold
    
    TextOut picBuffer.hdc, xPos + 1, yPos + 1, sIn, Len(sIn)
    TextOut picBuffer.hdc, xPos - 1, yPos - 1, sIn, Len(sIn)
    TextOut picBuffer.hdc, xPos - 1, yPos + 1, sIn, Len(sIn)
    TextOut picBuffer.hdc, xPos + 1, yPos - 1, sIn, Len(sIn)
    TextOut picBuffer.hdc, xPos - 1, yPos, sIn, Len(sIn)
    TextOut picBuffer.hdc, xPos + 1, yPos, sIn, Len(sIn)
    
    SetTextColor picBuffer.hdc, lColor
    TextOut picBuffer.hdc, xPos, yPos, sIn, Len(sIn)

End Sub
Function ScoreForBalls(iIn As Long) As Long
'2 * 1.5=3
'3 * 2=6
'4 * 2.5=10

    ScoreForBalls = iIn * ((iIn + 1) * 0.5)

End Function

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim lCellPos As Long
Dim xPos As Long
Dim yPos As Long
Dim Surr As Long

    If Button = 2 Then
        Me.WindowState = 1
        Exit Sub
    End If

    If Not bInGame Then
        If Not bEnteringHiScore Then
            timHiScores.Enabled = False
            ResetScoreDisplay
            NewGame
        End If
        Exit Sub
    End If
    
    If bProcessing Or UserPressedReset Then Exit Sub
    
    If X >= (picBuffer.Width - picReset.Width - 2) And _
       y >= 2 And _
       X <= ((picBuffer.Width - picReset.Width - 2) + picReset.Width) And _
       y <= (2 + picReset.Width) Then
        UserPressedReset = True
        Exit Sub
    End If
    
    bProcessing = True
    lCellPos = EstablishCellFromMousePos(CLng(X), CLng(y))
    
    If (lCellPos Mod 100 = 0) Or (lCellPos \ 100 = 0) Then
        GoTo ResetProcessFlag
    End If
    
    xPos = (lCellPos Mod 100) - 1
    yPos = (lCellPos \ 100) - 1
    
    If GameGrid(yPos, xPos) = 0 Then GoTo ResetProcessFlag
    
    If HowManySurroundingSameColor(lCellPos) Then
        Surr = GetCoordinatesOfTouchingBalls(lCellPos)
        lCurrentScore = lCurrentScore + ScoreForBalls(Surr)
        AnimateGrid
    End If
ResetProcessFlag:
    bProcessing = False
End Sub
Sub ApplyBlend()
Dim Blend As BLENDFUNCTION
Dim BlendPtr As Long
    Blend.SourceConstantAlpha = 130
    
    CopyMemory BlendPtr, Blend, 4
    
    AlphaBlend picBuffer.hdc, 0, 0, picBuffer.Width, picBuffer.Height, picBlank.hdc, 0, 0, picBuffer.Width, picBuffer.Height, BlendPtr
End Sub
Sub AnimateGrid()
Dim X As Long
Dim y As Long
Dim z As Long
Dim l As Long
Dim m As Long
Dim bColumnEmpty As Boolean

    For m = 1 To 25 'an unnecessary amount of loop, but I can't be arsed to flag when all blocks are moved properly!
        For X = 0 To 24
            For l = 1 To 8
                For y = 7 To 0 Step -1
                    'ignore if empty space
                    If GameGrid(y, X) > 0 Then 'there's a tile there. is there a space underneath?
                        If GameGrid(y + 1, X) = 0 Then 'need to shuffle this tile down until it hits the bottom or next tile
                            z = y + 1
                            Do
                                GameGrid(z, X) = GameGrid(z - 1, X)
                                GameGrid(z - 1, X) = 0
                                z = z + 1
                            Loop Until z >= 8 Or GameGrid(z, X) > 0
                            DoEvents

                        End If
                    End If
                Next
            Next
            If X < 24 Then
                bColumnEmpty = True
                If GameGrid(8, X) = 0 Then
                    For l = X + 1 To 24
                        For y = 0 To 9
                            GameGrid(y, l - 1) = GameGrid(y, l)
                            GameGrid(y, l) = 0
                        Next
                        
                        DoEvents
                    Next
                End If
            End If
        Next
    Next
End Sub
Function HowManySurroundingSameColor(lPosition As Long) As Long
Dim xPos As Long
Dim yPos As Long
Dim TargetCellColor As Long
    xPos = (lPosition Mod 100) - 1
    yPos = (lPosition \ 100) - 1
    TargetCellColor = GameGrid(yPos, xPos)
    
    HowManySurroundingSameColor = 0
    
    'left
    If xPos > 0 Then
        If GameGrid(yPos, xPos - 1) = TargetCellColor Then
            HowManySurroundingSameColor = HowManySurroundingSameColor + 1
        End If
    End If
    
    'top
    If yPos > 0 Then
        If GameGrid(yPos - 1, xPos) = TargetCellColor Then
            HowManySurroundingSameColor = HowManySurroundingSameColor + 1
        End If
    End If
    
    'right
    If xPos < 24 Then
        If GameGrid(yPos, xPos + 1) = TargetCellColor Then
            HowManySurroundingSameColor = HowManySurroundingSameColor + 1
        End If
    End If
    
    'bottom
    If yPos < 8 Then
        If GameGrid(yPos + 1, xPos) = TargetCellColor Then
            HowManySurroundingSameColor = HowManySurroundingSameColor + 1
        End If
    End If

End Function
Function GetCoordinatesOfSurroundingBalls(lPosition As Long) As Long()
Dim Coords() As Long
Dim xPos As Long
Dim yPos As Long
Dim TargetCellColor As Long

    xPos = (lPosition Mod 100) - 1
    yPos = (lPosition \ 100) - 1


    TargetCellColor = GameGrid(yPos, xPos)

    ReDim Coords(2, 0)
    
    'left
    If xPos > 0 Then
        If GameGrid(yPos, xPos - 1) = TargetCellColor Then
            ReDim Preserve Coords(2, UBound(Coords, 2) + 1)
            Coords(0, UBound(Coords, 2) - 1) = xPos - 1
            Coords(1, UBound(Coords, 2) - 1) = yPos
        End If
    End If
    
    'top
    If yPos > 0 Then
        If GameGrid(yPos - 1, xPos) = TargetCellColor Then
            ReDim Preserve Coords(2, UBound(Coords, 2) + 1)
            Coords(0, UBound(Coords, 2) - 1) = xPos
            Coords(1, UBound(Coords, 2) - 1) = yPos - 1
        End If
    End If
    
    'right
    If xPos < 24 Then
        If GameGrid(yPos, xPos + 1) = TargetCellColor Then
            ReDim Preserve Coords(2, UBound(Coords, 2) + 1)
            Coords(0, UBound(Coords, 2) - 1) = xPos + 1
            Coords(1, UBound(Coords, 2) - 1) = yPos
        End If
    End If
    
    'bottom
    If yPos < 8 Then
        If GameGrid(yPos + 1, xPos) = TargetCellColor Then
            ReDim Preserve Coords(2, UBound(Coords, 2) + 1)
            Coords(0, UBound(Coords, 2) - 1) = xPos
            Coords(1, UBound(Coords, 2) - 1) = yPos + 1
        End If
    End If
    If UBound(Coords, 2) > 0 Then
        ReDim Preserve Coords(2, UBound(Coords, 2) - 1)
    End If
    GetCoordinatesOfSurroundingBalls = Coords
End Function
Function AreAllTheTilesGone() As Boolean
Dim X As Long
Dim y As Long
    AreAllTheTilesGone = True
    For X = 0 To 24
        For y = 0 To 8
            If GameGrid(y, X) > 0 Then
                AreAllTheTilesGone = False
                Exit Function
            End If
        Next
    Next
End Function
Function GetCoordinatesOfTouchingBalls(lPosition As Long) As Long
Dim Coords() As Long
Dim MainCoords() As Long
Dim xPos As Long
Dim yPos As Long
Dim X As Long
Dim y As Long
Dim lWorkingOn As Long
Dim xCurrent As Long
Dim yCurrent As Long
Dim CurrCellPos As Long
Dim bNewPos As Boolean
    xPos = (lPosition Mod 100) - 1
    yPos = (lPosition \ 100) - 1

    
    Coords = GetCoordinatesOfSurroundingBalls(lPosition)
    ReDim MainCoords(2, UBound(Coords, 2) + 1)
    
    For X = 0 To UBound(Coords, 2) 'first step, initial touching balls to MainCoords, then loop
        MainCoords(0, X) = Coords(0, X) 'xpos
        MainCoords(1, X) = Coords(1, X) 'ypos
    Next
    
    lWorkingOn = 0
    
    Do
        xCurrent = MainCoords(0, lWorkingOn)
        yCurrent = MainCoords(1, lWorkingOn)
        CurrCellPos = (xCurrent + 1) + (100 * (yCurrent + 1))
        
        If HowManySurroundingSameColor(CurrCellPos) > 0 Then
            Coords = GetCoordinatesOfSurroundingBalls(CurrCellPos)
                    
            For X = 0 To UBound(Coords, 2)
                bNewPos = True
                For y = 0 To UBound(MainCoords, 2) - 1
                    If Coords(0, X) = MainCoords(0, y) And Coords(1, X) = MainCoords(1, y) Then 'already there!
                        bNewPos = False
                        Exit For
                    End If
                Next
                If bNewPos Then
                    ReDim Preserve MainCoords(2, UBound(MainCoords, 2) + 1)
                    MainCoords(0, UBound(MainCoords, 2) - 1) = Coords(0, X) 'xpos
                    MainCoords(1, UBound(MainCoords, 2) - 1) = Coords(1, X) 'ypos
                End If
            Next X
        End If
        GameGrid(yCurrent, xCurrent) = 0
        lWorkingOn = lWorkingOn + 1
        DoEvents
    Loop Until lWorkingOn >= UBound(MainCoords, 2)
    GameGrid(yPos, xPos) = 0
    GetCoordinatesOfTouchingBalls = lWorkingOn
End Function
Sub WriteTextToBuffer()
Dim sScore As String
    sScore = "Score: " & lCurrentScore
    pTextOut sScore, "Arial", 8, True, 3, 3, vbWhite
    pTextOut sCurrentHI, "Arial", 8, True, picBuffer.TextWidth(sScore) + 15, 3, RGB(100, 200, 100)
    If bProcessing Then
        pTextOut "Processing", "Arial", 8, False, 3, 16, vbWhite
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Timer1.Enabled = False
End Sub

Private Sub Timer1_Timer()
Dim sPos As String
Dim lCellPos As Long
Dim xPos As Long
Dim yPos As Long

    ClearBuffer
    DrawGridToBuffer
    WriteTextToBuffer
    If IsTheGameOver Or UserPressedReset Then
        Timer1.Enabled = False
        bInGame = False
        If AreAllTheTilesGone Then lCurrentScore = lCurrentScore + 500
        ApplyBlend
        DrawLogoToBuffer
        ApplyBlend
        ApplyBlend
        'check high scores
        If lCurrentScore > CLng(ReadINI("SETTINGS", "HISCORE10", HiScoreFile)) Then
            If AreAllTheTilesGone Then DrawBonusToBuffer
            pTextOut "Your score of", "Arial", 12, True, 128, 10, RGB(255, 255, 200)
            pTextOut CStr(lCurrentScore), "Arial", 16, True, 156, 35, vbWhite
            pTextOut "Made the Hall of Fame!", "Arial", 12, True, 95, 70, RGB(255, 255, 200)
            pTextOut "Enter you nickname and press enter", "Arial", 10, True, 65, 105, RGB(200, 255, 255)
            txtName.Visible = True
            bEnteringHiScore = True
            BufferToScreen
        Else
            FreshDisplay
        End If
    Else
        DrawResetButtonToBuffer
        BufferToScreen
    End If
    
End Sub
Private Sub timHiScores_Timer()
Dim HiScoreString As String
Dim X As Integer
Dim sText As String
Dim xPos As Long
Dim yPos As Long
Dim zPos As Long
Dim xOffset As Long
Dim yOffset As Long
Dim zOffset As Long
Dim Col(3) As Long
Dim RGBCol As Long
    
    DrawGridToBuffer
    
    ApplyBlend
    If bStars Then 'go to the top of the form code to change this boolean
    
        CreateFrame
    Else
        For zPos = -2 To 2 Step 2
            zOffset = zPos * 40
            For yPos = -2 To 2 Step 2
                yOffset = yPos * 40
                For xPos = -2 To 2 Step 2
                    Col(0) = Int(Rnd * 256): Col(1) = Int(Rnd * 256): Col(2) = Int(Rnd * 256)
                    RGBCol = RGB(190, 200, 250) 'RGB(Col(0), Col(1), Col(2))
                    xOffset = xPos * 40
                    'front plane
                    PlotLine xOffset - 16, yOffset - 16, zOffset + 15, xOffset + 16, yOffset - 16, zOffset + 15, lTheta, lAlt, lSize, lperspective, RGBCol
                    PlotLine xOffset + 16, yOffset - 16, zOffset + 15, xOffset + 16, yOffset + 16, zOffset + 15, lTheta, lAlt, lSize, lperspective, RGBCol
                    PlotLine xOffset - 16, yOffset + 16, zOffset + 15, xOffset + 16, yOffset + 16, zOffset + 15, lTheta, lAlt, lSize, lperspective, RGBCol
                    PlotLine xOffset - 16, yOffset - 16, zOffset + 15, xOffset - 16, yOffset + 16, zOffset + 15, lTheta, lAlt, lSize, lperspective, RGBCol
                    
                    'rear plane
                    PlotLine xOffset - 16, yOffset - 16, zOffset - 15, xOffset + 16, yOffset - 16, zOffset - 15, lTheta, lAlt, lSize, lperspective, RGBCol
                    PlotLine xOffset + 16, yOffset - 16, zOffset - 15, xOffset + 16, yOffset + 16, zOffset - 15, lTheta, lAlt, lSize, lperspective, RGBCol
                    PlotLine xOffset - 16, yOffset + 16, zOffset - 15, xOffset + 16, yOffset + 16, zOffset - 15, lTheta, lAlt, lSize, lperspective, RGBCol
                    PlotLine xOffset - 16, yOffset - 16, zOffset - 15, xOffset - 16, yOffset + 16, zOffset - 15, lTheta, lAlt, lSize, lperspective, RGBCol
                    
                    'join corners
                    PlotLine xOffset - 16, yOffset - 16, zOffset + 15, xOffset - 16, yOffset - 16, zOffset - 15, lTheta, lAlt, lSize, lperspective, RGBCol
                    PlotLine xOffset + 16, yOffset - 16, zOffset + 15, xOffset + 16, yOffset - 16, zOffset - 15, lTheta, lAlt, lSize, lperspective, RGBCol
                    PlotLine xOffset - 16, yOffset + 16, zOffset + 15, xOffset - 16, yOffset + 16, zOffset - 15, lTheta, lAlt, lSize, lperspective, RGBCol
                    PlotLine xOffset + 16, yOffset + 16, zOffset + 15, xOffset + 16, yOffset + 16, zOffset - 15, lTheta, lAlt, lSize, lperspective, RGBCol
                Next
            Next
        Next
        
        lAlt = lAlt + 2
        If lAlt >= 360 Then lAlt = lAlt - 360
        
        lTheta = lTheta + 4
        If lTheta >= 360 Then lTheta = lTheta - 360
        
        If lSize < 11000 Then lSize = lSize + 50
    End If
    DrawLogoToBuffer
    
    pTextOut "Click to play", "Arial", 10, True, 140, 82, vbWhite
    
    'hi score line
    sCurrentHI = 1

    For X = 1 To 10
        HiScoreString = HiScoreString & X & ". " & ReadINI("SETTINGS", "HINAME" & X, HiScoreFile) & _
                        " - " & ReadINI("SETTINGS", "HISCORE" & X, HiScoreFile) & "         "
    Next X
    pTextOut HiScoreString, "Century Gothic", 12, True, lDisplayingHiScoreXPos, picBuffer.Height - 55, RGB(255, 255, 200)
    
    If lCurrentScore > 0 Then
        sText = "Last score was " & lCurrentScore
        pTextOut sText, "Arial", 10, True, (picBuffer.Width \ 2) - (picBuffer.TextWidth(sText) \ 2) + 10, picBuffer.Height - 25, RGB(220, 220, 190)
    End If
    
    lDisplayingHiScoreXPos = lDisplayingHiScoreXPos - 4
    
    If (lDisplayingHiScoreXPos + picBuffer.TextWidth(HiScoreString)) < 0 Then
        lDisplayingHiScoreXPos = picBuffer.Width + 5
    End If
    BufferToScreen

End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
Dim X As Integer
Dim y As Integer
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtName.Visible = False
        'add name
        If lCurrentScore > CLng(ReadINI("SETTINGS", "HISCORE10", HiScoreFile)) Then
            For X = 1 To 10
                If lCurrentScore > CLng(ReadINI("SETTINGS", "HISCORE" & X, HiScoreFile)) Then
                    If X < 10 Then
                        For y = 10 To X + 1 Step -1
                            writeINI "SETTINGS", "HISCORE" & y, ReadINI("SETTINGS", "HISCORE" & y - 1, HiScoreFile), HiScoreFile
                            writeINI "SETTINGS", "HINAME" & y, ReadINI("SETTINGS", "HINAME" & y - 1, HiScoreFile), HiScoreFile
                        Next
                    End If
                    writeINI "SETTINGS", "HISCORE" & X, CStr(lCurrentScore), HiScoreFile
                    writeINI "SETTINGS", "HINAME" & X, txtName, HiScoreFile
                    Exit For
                End If
            Next
        End If
        bEnteringHiScore = False
        FreshDisplay
    End If
End Sub
'star field stuff
Sub PutStarsToBuffer()
Dim lX As Long
Dim lPosX As Long
Dim lPosY As Long

    For lX = 0 To lNumStars - 1
        lPosX = GimmeX(Stars(lX).Angle, Stars(lX).Len) + lCentreOfPicture(0)
        lPosY = GimmeY(Stars(lX).Angle, Stars(lX).Len) + lCentreOfPicture(1)
        SetPixel picBuffer.hdc, lPosX, lPosY, Stars(lX).Color
    Next lX
    
End Sub
Sub SetCentreOfPicture()
    lCentreOfPicture(0) = picBuffer.ScaleWidth \ 2
    lCentreOfPicture(1) = picBuffer.ScaleHeight \ 2
End Sub
Sub SetupStarField()
Dim lX As Long

    ReDim Stars(lNumStars)
    
    SetCentreOfPicture
    
    lMaxLength = IIf(lCentreOfPicture(0) < lCentreOfPicture(1), lCentreOfPicture(0), lCentreOfPicture(1))
    
    For lX = 0 To lNumStars - 1
        With Stars(lX)
            .Angle = Rnd * 360
            .Speed = Int(Rnd * 20) + 1
            .Len = 10 + (Rnd * (lMaxLength - 10))
            .Color = RGB(50, 50, 50)
        End With
    Next lX
End Sub
Sub MoveStars()
Dim lX As Long
Dim lPosX As Long
Dim lPosY As Long

    For lX = 0 To lNumStars - 1
        Stars(lX).Len = Stars(lX).Len + Stars(lX).Speed
        lPosX = GimmeX(Stars(lX).Angle, Stars(lX).Len) + lCentreOfPicture(0)
        lPosY = GimmeY(Stars(lX).Angle, Stars(lX).Len) + lCentreOfPicture(1)
        Stars(lX).Color = AdjustBrightness(Stars(lX).Color, 8, True)
        If (lPosX < 0 Or lPosX > picBuffer.ScaleWidth) _
        Or (lPosY < 0 Or lPosY > picBuffer.ScaleHeight) Then
            'if a star goes off screen, place it back in the middle
            With Stars(lX)
                .Angle = Rnd * 360
                .Speed = Int(Rnd * 20) + 1
                .Len = 10 + (Rnd * (lMaxLength - 10))
                .Color = RGB(50, 50, 50)
                
            End With
        End If
    Next lX

End Sub
Sub CreateFrame()
Dim X As Long
Dim hRPen As Long
Dim Point As POINTAPI
Dim yPos As Long
Dim xPos As Long
Dim sOut As String
Dim lCol As Long

    PutStarsToBuffer
    MoveStars
    
End Sub
'cube-----------------------------------------
Sub PlotLine(x1 As Long, y1 As Long, z1 As Long, x2 As Long, y2 As Long, z2 As Long, Theta As Long, Alt As Long, Size As Long, Perspective As Long, lColor As Long)
Dim cX As Single, cY As Single
Dim vX As Single, vY As Single, vZ As Single
Dim pX1 As Single, pY1 As Single
Dim pX2 As Single, pY2 As Single
Dim Phi As Single
Dim Sin_Theta As Single, Cos_Theta As Single, Sin_Phi   As Single, Cos_Phi   As Single
    
    Phi = 90 - Alt
    cX = (picBuffer.Width) / 2: cY = (picBuffer.Height) / 2
    Sin_Theta = Sine(Theta): Cos_Theta = Cosine(Theta): Sin_Phi = Sine(Phi): Cos_Phi = Cosine(Phi)
    
    vX = -x1 * Sin_Theta + y1 * Cos_Theta
    vY = -x1 * Cos_Theta * Cos_Phi - y1 * Sin_Theta * Cos_Phi + z1 * Sin_Phi
    vZ = -x1 * Cos_Theta * Sin_Phi - y1 * Sin_Theta * Sin_Phi - z1 * Cos_Phi + Perspective
    pX1 = cX + Size * vX / vZ: pY1 = cY - Size * vY / vZ
    
    vX = -x2 * Sin_Theta + y2 * Cos_Theta
    vY = -x2 * Cos_Theta * Cos_Phi - y2 * Sin_Theta * Cos_Phi + z2 * Sin_Phi
    vZ = -x2 * Cos_Theta * Sin_Phi - y2 * Sin_Theta * Sin_Phi - z2 * Cos_Phi + Perspective
    pX2 = cX + Size * vX / vZ: pY2 = cY - Size * vY / vZ

    DrawLineToBuffer CLng(pX1), CLng(pY1), CLng(pX2), CLng(pY2), lColor

End Sub

Sub DrawLineToBuffer(x1 As Long, y1 As Long, x2 As Long, y2 As Long, lColor As Long)
Dim hRPen As Long
Dim Point As POINTAPI
    picBuffer.ForeColor = lColor
    Point.X = x1: Point.y = y1
    MoveToEx picBuffer.hdc, x1, y1, Point
    LineTo picBuffer.hdc, x2, y2
End Sub

