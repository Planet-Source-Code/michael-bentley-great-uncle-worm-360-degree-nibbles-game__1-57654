VERSION 5.00
Begin VB.Form frmWorm 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Great Uncle Worm"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4500
   Icon            =   "frmWorm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   200
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   300
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picAppleM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   3
      Left            =   1920
      Picture         =   "frmWorm.frx":08CA
      ScaleHeight     =   14
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   14
      Top             =   1320
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picAppleM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   150
      Index           =   2
      Left            =   1800
      Picture         =   "frmWorm.frx":0C36
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   12
      TabIndex        =   13
      Top             =   1080
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.PictureBox picAppleM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   120
      Index           =   1
      Left            =   1800
      Picture         =   "frmWorm.frx":0F8D
      ScaleHeight     =   8
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   12
      Top             =   840
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox picApple 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   3
      Left            =   1560
      Picture         =   "frmWorm.frx":12CA
      ScaleHeight     =   14
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   11
      Top             =   1320
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picApple 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   150
      Index           =   2
      Left            =   1560
      Picture         =   "frmWorm.frx":1636
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   12
      TabIndex        =   10
      Top             =   1080
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.PictureBox picApple 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   120
      Index           =   1
      Left            =   1560
      Picture         =   "frmWorm.frx":1990
      ScaleHeight     =   8
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   9
      Top             =   840
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox picSword 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   150
      Index           =   1
      Left            =   1800
      Picture         =   "frmWorm.frx":1CDC
      ScaleHeight     =   150
      ScaleWidth      =   150
      TabIndex        =   7
      Top             =   600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox picSword 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   150
      Index           =   0
      Left            =   1560
      Picture         =   "frmWorm.frx":2027
      ScaleHeight     =   150
      ScaleWidth      =   150
      TabIndex        =   6
      Top             =   600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox picAppleM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   105
      Index           =   0
      Left            =   1800
      Picture         =   "frmWorm.frx":2375
      ScaleHeight     =   7
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   8
      TabIndex        =   5
      Top             =   360
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox picApple 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   105
      Index           =   0
      Left            =   1560
      Picture         =   "frmWorm.frx":26BC
      ScaleHeight     =   7
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   8
      TabIndex        =   4
      Top             =   360
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox picCircles 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4125
      Left            =   3120
      Picture         =   "frmWorm.frx":2A03
      ScaleHeight     =   4125
      ScaleWidth      =   5250
      TabIndex        =   3
      Top             =   1920
      Visible         =   0   'False
      Width           =   5250
   End
   Begin VB.PictureBox picHorLine 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   75
      Left            =   0
      Picture         =   "frmWorm.frx":7CC7
      ScaleHeight     =   75
      ScaleWidth      =   6750
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   6750
   End
   Begin VB.PictureBox picVertLine 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3750
      Left            =   0
      Picture         =   "frmWorm.frx":803C
      ScaleHeight     =   3750
      ScaleWidth      =   75
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.PictureBox picDot 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   75
      Left            =   720
      Picture         =   "frmWorm.frx":839E
      ScaleHeight     =   5
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   5
      TabIndex        =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Label lblScore 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Current Score: 0"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Visible         =   0   'False
      Width           =   1815
   End
End
Attribute VB_Name = "frmWorm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Code by Michael Bentley (ikillkenny@comcast.net)
'Visit my site, www.doc-ent.com for more freeware games

Const pi As Variant = 3.14159265358979

'Declarations for printing text
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
'Declarations for gameloop
Private Declare Function GetTickCount Lib "kernel32" () As Long
'Gameloop variables
Dim LastTick, CurrentTick As Long
Dim TickDif As Long

Dim drawAll As Boolean 'Whether or not to draw all of the dots in this loop

Dim intDiv As Integer

Dim initialGrowLength As Integer 'How much you grow in standard mode
Dim intGrowLength As Integer 'How much you grow when the level starts

Dim curLevel As Long 'Current level that the player is on
Dim bWin As Boolean 'Whether or not the player has eaten enough dots to open the exit door
Dim intShrinkDot As Long

Dim curMidi As String
Dim midiNum As Integer

Dim intShrinkWorm As Long
Dim intShrinkRate As Long 'How much the worm is shrinking over time

Dim DotsEaten As Long 'Number of dots eaten

Dim intScore As Long 'Current score

Dim intMaxLength As Long 'Maximum length that the worm can be

Dim bShake As Long 'Whether or not the form is "shaking" or not

Dim Daggar(1 To 6) As Dot
Dim intDaggarDrop As Long
Dim intDaggarMax As Long 'Number of daggars to drop

Dim bRunning As Boolean 'Whether or not the loop is being cycled
Dim WormLength As Long 'Current length of Uncle Worm
'If you're pressing the following buttons:
Dim bLeft As Boolean
Dim bRight As Boolean

'Current angle of movement for the worm
Dim curXAngle As Long

'Where the next x and y will be located
Dim intNextX As Variant
Dim intNextY As Variant

Dim curDot As Long 'Which dot in the array is being drawn
Dim bGrow As Boolean 'If uncle worm is growing or not
Dim Apple As typeApple 'Current Apple
Dim Worm() As Dot 'Array of dots representing uncle worm

Private Sub Form_Activate()
If bLoaded = False Then
    Call Form_Load
End If
drawAll = True
bRunning = True
Call GameLoop
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

'Detect keys
If KeyCode = vbKeyLeft Then
    bLeft = True
    bRight = False
End If
If KeyCode = vbKeyRight Then
    bRight = True
    bLeft = False
End If

If KeyCode = vbKeyP Then 'Pause/unpause the game
    If bRunning = True Then
        bRunning = False
        lblScore.Caption = "Current Score: " & intScore
        lblScore.Visible = True
    Else
        bRunning = True
        lblScore.Visible = False
        Call GameLoop
    End If
End If
If KeyCode = vbKeyF1 Then
    drawAll = True
    MsgBox Apple.Left & " " & Apple.Top & " form: " & Me.ScaleWidth & " " & Me.ScaleHeight
End If

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

'Set keys to being not pressed
If KeyCode = vbKeyLeft Then
    bLeft = False
End If
If KeyCode = vbKeyRight Then
    bRight = False
End If

End Sub

Private Sub Form_Load()
On Error Resume Next
'Make sure the game has not already been loaded
If bLoaded = True Then Exit Sub
bLoaded = True

bShake = 0 'Currently not shaking the form


'Play the background music
Randomize
Dim intRand As Integer
intRand = Int(Rnd * 4) + 1
midiNum = intRand
SetMidi
PlayMidi (curMidi)

'Set font information
Me.Font.Size = 8
Me.Font.Bold = True
Me.ForeColor = RGB(255, 0, 0)
Me.Font = "Arial"


If GameType = 1 Then 'Endless
    ReDim Worm(1 To 10000) As Dot
    intMaxLength = 10000
ElseIf GameType = 4 Then 'Dagger Drop
    ReDim Worm(1 To 11000) As Dot
    intMaxLength = 11000
    For i = 1 To 6
        Daggar(i).Visible = False
    Next 'i
Else 'Any other mode
    ReDim Worm(1 To 10000) As Dot
    intMaxLength = 10000
End If

intScore = 0 'Reset score
curDot = 1 'Current dot is the first dot

'Set the worm to move up when started
Call MoveWorm
curXAngle = intDiv / -2 + (intDiv / 10) - 1


For i = 1 To intMaxLength 'Clear uncle worm
    Worm(i).Visible = False
Next 'i


Speed = Speed + 1 'Increase speed so that the worm actually moves in slow mode
intGrowLength = 40 'Initial grow value of uncle worm
DotsEaten = 0 'No dots eaten yet

'Set screen size
If Size = 0 Then
    Me.Width = 220 * Screen.TwipsPerPixelX
    Me.Height = 175 * Screen.TwipsPerPixelY
ElseIf Size = 1 Then
    Me.Width = 260 * Screen.TwipsPerPixelX
    Me.Height = 210 * Screen.TwipsPerPixelY
Else
    Me.Width = 350 * Screen.TwipsPerPixelX
    Me.Height = 275 * Screen.TwipsPerPixelY
End If

If GameType <> 5 Then 'Make an apple unless you're in Survival Mode
    Call CreateApple
Else
    Apple.Visible = False
End If

'Not pressing any keys
bRight = False
bLeft = False

'Set worm length to 0 and state to growing
WormLength = 1
bGrow = True

TickDif = 20 'Gamespeed



If GameType = 0 Then
    initialGrowLength = 40 'Grow more as the game goes on
End If

'Set the rate at which you're shrinking if in shrink mode
If GameType = 3 Then
    If Difficulty = 0 Then
        intShrinkRate = 15
    ElseIf Difficulty = 1 Then
        intShrinkRate = 11
    ElseIf Difficulty = 2 Then
        intShrinkRate = 7
    End If
End If

'Set initial grow length in endless mode
If GameType = 1 Then
    If Difficulty = 0 Then
        intGrowLength = 100
    ElseIf Difficulty = 1 Then
        intGrowLength = 200
    Else
        intGrowLength = 300
    End If
End If

'Set number of daggers in dagger drop
If GameType = 4 Then
    intDaggarDrop = 0
    If Difficulty = 0 Then
        intDaggarMax = 100
    ElseIf Difficulty = 1 Then
        intDaggarMax = 90
    Else
        intDaggarMax = 80
    End If
End If

'Set grow length to huge in survival mode
If GameType = 5 Then
    intGrowLength = 15000
End If

'Create new dots for uncle worm
For i = 1 To intMaxLength
    Worm(i).Visible = False
    Worm(i).Left = Me.ScaleWidth / 2
    Worm(i).Top = Me.ScaleHeight - 10
Next 'i
bWin = False 'You haven't got all of the dots left
curLevel = 1
End Sub
Sub GameLoop()
On Error Resume Next
Do While bRunning
CurrentTick = GetTickCount()
    If CurrentTick - LastTick > TickDif Then 'Loop only every TickDif interval
        LastTick = CurrentTick 'Update the last tick to the current one
        If drawAll = True Then
            Me.Cls 'Clear the screen
        End If
        
        Call MoveWorm
        If GameType = 2 Then
            Call ShrinkForm
        End If
        Call DrawWorm
        
        If bShake > 0 Then 'Shake the screen
            Call Shake
        End If
        
        
        Me.Refresh
        DoEvents
    Else
        DoEvents
    End If
Loop
End Sub
Sub MoveWorm()
On Error Resume Next

'Change x angle if left or right are pressed
If bRight = True Then
    curXAngle = curXAngle + 1
End If
If bLeft = True Then
    curXAngle = curXAngle - 1
End If


'Determine at what angle Uncle Worm turns
If Control = 0 Then
    If Speed = 1 Then
        intDiv = 40
    ElseIf Speed = 2 Then
        intDiv = 36
    ElseIf Speed = 3 Then
        intDiv = 32
    Else
        intDiv = 28
    End If
ElseIf Control = 1 Then
    If Speed = 1 Then
        intDiv = 30
    ElseIf Speed = 2 Then
        intDiv = 26
    ElseIf Speed = 3 Then
        intDiv = 22
    Else
        intDiv = 18
    End If
ElseIf Control = 2 Then
    If Speed = 1 Then
        intDiv = 20
    ElseIf Speed = 2 Then
        intDiv = 18
    ElseIf Speed = 3 Then
        intDiv = 16
    Else
        intDiv = 14
    End If
Else
    If Speed = 1 Then
        intDiv = 16
    ElseIf Speed = 2 Then
        intDiv = 15
    ElseIf Speed = 3 Then
        intDiv = 14
    Else
        intDiv = 13
    End If
End If

'Change x and y of worm
intNextX = Sin(curXAngle * pi / intDiv)
intNextY = Cos(curXAngle * pi / intDiv)

End Sub
Sub DrawWorm()
On Error Resume Next
Dim tempDrawAll As Boolean
Dim LastDot As Long 'The last dot drawn

tempDrawAll = drawAll

If drawAll = True Then
    Call BitBlt(Me.hdc, 0, 0, Me.ScaleWidth, Me.ScaleHeight, picCircles.hdc, 0, 0, vbSrcCopy)
End If

If intGrowLength > 0 Then 'If uncle worm is still growing
    WormLength = WormLength + Speed 'Increase the worm length by the speed
    intGrowLength = intGrowLength - Speed
    If intGrowLength < 0 Then intGrowLength = 0
End If

If WormLength >= UBound(Worm) Then
    ReDim Worm(1 To WormLength + 10) As Dot 'Update the array to a new length if the worm has grown
End If


If GameType = 3 Then 'The worm is shrinking right now
    If WormLength = 0 Then 'The worm has shrunk to nothing, game over
        Call GameOver
        Exit Sub
    End If
    intShrinkWorm = intShrinkWorm + 1
    If intShrinkWorm >= intShrinkRate Then 'It's time to shrink the worm
        For i = 1 To Speed 'Shrink the worm by speed
            intShrinkWorm = 1
            WormLength = WormLength - 1 'Decrease the worm length
            LastDot = curDot - WormLength
            If LastDot < 0 Then
                LastDot = intMaxLength + LastDot
            End If
            Worm(LastDot).Visible = False 'Get rid of the last dot
            'Hide the last dot
            Call BitBlt(Me.hdc, Worm(LastDot).Left, Worm(LastDot).Top, 1, 1, picCircles.hdc, Worm(LastDot).Left, Worm(LastDot).Top, vbSrcCopy)
            If WormLength < 0 Then 'Uncle worm has shrunk to nothing
                Call GameOver
                Exit Sub
            End If
        Next 'i
    End If
    
End If

For i = 1 To Speed
    If GameType = 5 Then 'Increase score in Survival mode
        intScore = intScore + 1
    End If
    
    curDot = curDot + 1
    If curDot > intMaxLength Then curDot = 1
    
    If intGrowLength <= 0 Then 'The worm is not growing
    
        'Get rid of the last dot and make a new one in the front
        LastDot = curDot - WormLength
        If LastDot <= 0 Then
            LastDot = intMaxLength + LastDot
        End If
    
    
        Worm(LastDot).Visible = False 'Hide the last dot
        If drawAll = False Then 'Get rid of the last dot by drawing the background at its location
            Call BitBlt(Me.hdc, Worm(LastDot).Left, Worm(LastDot).Top, 1, 1, picCircles.hdc, Worm(LastDot).Left, Worm(LastDot).Top, vbSrcCopy)
        End If
        
    End If

    Worm(curDot).Visible = True 'Draw the new dot

    If curDot = 1 Then
        LastDot = intMaxLength
    Else
        LastDot = curDot - 1
    End If
    
    
    'Position the new dot
    Worm(curDot).Top = intNextX + Worm(LastDot).Top
    Worm(curDot).Left = intNextY + Worm(LastDot).Left


    Dim bcollide As Boolean 'Is the worm hitting the apple?
    bcollide = DetectCollide(Worm(curDot).Left, Worm(curDot).Top, 1, 1, Apple.Left, Apple.Top, Apple.Width, Apple.Height)
    If bcollide = True And bWin = False Then 'Worm has hit the apple, draw a new apple
        PlySound ("gulp")
        If GameType <> 4 Then 'If the game isn't dagger drop
            intGrowLength = intGrowLength + 35 + (Apple.Pic * 5) 'Increase worm length
        Else
            intDaggarMax = intDaggarMax - 2 'Increase speed at which daggers drop
            If intDaggarMax < 20 Then
                intDaggarMax = 20
            End If
        End If
        DotsEaten = DotsEaten + 1 'Increase number of dots eaten
        intScore = intScore + (6 + Apple.Pic) 'Increase score
        If GameType = 3 Then 'Shrinking mode
            intShrinkRate = intShrinkRate - 1
            If intShrinkRate < Speed + 2 Then intShrinkRate = Speed + 2
        End If
        If DotsEaten < 10 Then 'Make a new apple
            Call CreateApple
        Else
            If GameType = 0 Or GameType = 2 Then 'Normal or Shrinking mode, open gate
                bWin = True
                drawAll = True
                bShake = 1
            Else
                Call CreateApple 'Create a new apple
            End If
        End If
    End If

    'Detect if uncle worm has hit a border
    If Worm(curDot).Left <= 5 Or Worm(curDot).Left >= Me.ScaleWidth - 5 Or Worm(curDot).Top <= 5 Or Worm(curDot).Top >= Me.ScaleHeight - 5 Then
        If Worm(curDot).Top > 5 Then 'Not at the top wall,  means you definitely hit a wall, game over
            Call GameOver
            Exit Sub
        Else
            If bWin = False Then 'The gate isn't open, game over
                Call GameOver
                Exit Sub
            Else
                'You've gone through the exit gate, load the next level
                If Worm(curDot).Left >= Me.ScaleWidth / 2 - 10 And Worm(curDot).Left <= Me.ScaleWidth / 2 + 10 Then
                    Call LoadNextLevel
                    Exit Sub
                Else 'You missed the exit gate, game over
                    Call GameOver
                    Exit Sub
                End If
            End If
        End If
    End If

    Dim bCollide2 As Boolean 'Temporary boolean
    bcollide = False
    bCollide2 = wormCollide(Worm(curDot).Left, Worm(curDot).Top) 'Check if it's possible the worm has collided by searching for drawn black pixels
    If bCollide2 = True Then
        For q = 1 To intMaxLength 'Check to see if you have run into another part of you
            If Worm(q).Visible = True And q <> curDot And q < LastDot - 10 Then
                bCollide2 = DetectCollide(Worm(curDot).Left, Worm(curDot).Top, 1, 1, Worm(q).Left, Worm(q).Top, 1, 1)
                If bCollide2 = True Then 'You have hit yourslef
                    bcollide = True
                    Exit For
                End If
            End If
        Next 'i
    End If
    If bcollide = True Then 'Game Over
        Call GameOver
        Exit Sub
    End If
    'Draw the new dot
    Call BitBlt(Me.hdc, Worm(curDot).Left, Worm(curDot).Top, 1, 1, picDot.hdc, 0, 0, vbSrcCopy)
Next 'i



If drawAll = True Then
    For i = 1 To intMaxLength
        If Worm(i).Visible = True Then 'Draw all visible dots of uncle worm
            Call BitBlt(Me.hdc, Worm(i).Left, Worm(i).Top, 1, 1, picDot.hdc, 0, 0, vbSrcCopy)
        End If
    Next 'i
End If

If bWin = False Then 'Draw the apple
    Call BitBlt(Me.hdc, Apple.Left, Apple.Top, Apple.Width, Apple.Height, picAppleM(Apple.Pic).hdc, 0, 0, vbSrcAnd)
    Call BitBlt(Me.hdc, Apple.Left, Apple.Top, Apple.Width, Apple.Height, picApple(Apple.Pic).hdc, 0, 0, vbSrcPaint)
End If

'Draw the level borders
Call BitBlt(Me.hdc, 0, Me.ScaleHeight - 5, 450, 5, picHorLine.hdc, 0, 0, vbSrcCopy)
Call BitBlt(Me.hdc, 0, 0, 5, 300, picVertLine.hdc, 0, 0, vbSrcCopy)
Call BitBlt(Me.hdc, Me.ScaleWidth - 5, 0, 5, 300, picVertLine.hdc, 0, 0, vbSrcCopy)
If bWin = False Then 'Draw the entire top line
    Call BitBlt(Me.hdc, 0, 0, 450, 5, picHorLine.hdc, 0, 0, vbSrcCopy)
Else 'Open the exit door
    Call BitBlt(Me.hdc, 0, 0, Me.ScaleWidth / 2 - 10, 5, picHorLine.hdc, 0, 0, vbSrcCopy)
    Call BitBlt(Me.hdc, Me.ScaleWidth / 2 + 10, 0, 450, Me.ScaleWidth / 2, picHorLine.hdc, 0, 0, vbSrcCopy)
End If

If GameType = 4 Then 'Dagger Drop
    drawAll = True 'Always redraw the form in this mode
    
    For q = 1 To 6
        If Daggar(q).Visible = True Then 'If the daggers are visible
            Daggar(q).Top = Daggar(q).Top + 1 'Make the dagger go down the screen
            If Daggar(q).Top >= Me.ScaleHeight - 10 Then 'Dagger has reached the bottom of the screen
                Daggar(q).Visible = False
            End If
            For i = 1 To intMaxLength
                If Worm(i).Visible = True Then
                    bCollide2 = DetectCollide(Worm(i).Left, Worm(i).Top, 1, 1, Daggar(q).Left, Daggar(q).Top, 10, 10)
                    If bCollide2 = True Then 'Dagger has hit uncle worm, make him grow
                        Daggar(q).Visible = False
                        intGrowLength = intGrowLength + 40
                        bShake = 1
                    End If
                End If
            Next 'i
        End If
    Next 'q
    
    intDaggarDrop = intDaggarDrop + 1
    If intDaggarDrop >= intDaggarMax Then 'Interval of time for next dagger to drop has passed, make a new dagger
        For i = 1 To 6
            If Daggar(i).Visible = False Then 'Find a dagger not being used
                intDaggarDrop = i
            End If
        Next 'i
        'Reset dagger values
        Daggar(intDaggarDrop).Visible = True
        Daggar(intDaggarDrop).Top = 0
        Dim intRand As Long
        intRand = Int(Rnd * (Me.ScaleWidth - 30)) + 15 'Random x position for dagger
        Daggar(intDaggarDrop).Left = intRand
        intDaggarDrop = 0
    End If

    For i = 1 To 6
        If Daggar(i).Visible = True Then
            Call BitBlt(Me.hdc, Daggar(i).Left, Daggar(i).Top, 10, 10, picSword(1).hdc, 0, 0, vbSrcAnd)
            Call BitBlt(Me.hdc, Daggar(i).Left, Daggar(i).Top, 10, 10, picSword(0).hdc, 0, 0, vbSrcPaint)
        End If
    Next 'i
End If
    


If GameType = 2 Then 'Shrinking mode
    If Apple.Left >= Me.ScaleWidth - Apple.Width - 10 Or Apple.Top >= Me.ScaleHeight - Apple.Height - 10 Then 'Check if the apple was drawn off of the screen
        Call CreateApple
    End If
End If

If tempDrawAll = True And drawAll = True Then
    If GameType <> 4 Then 'The mdoe isn't shrinking
        drawAll = False 'Don't draw the entire form
    End If
End If

'Draw the score
Call TextOut(Me.hdc, 5, Me.ScaleHeight - 10, "Score: " & CStr(intScore * Multiplier), Len(CStr(intScore * Multiplier)) + 7)
        

End Sub

Private Sub Form_Unload(Cancel As Integer)

'Clear worm array
Erase Worm()
intScore = 0 'Reset score
bRunning = False 'Stop the game loop

Me.Hide
frmIntro.Show

End Sub
Sub CreateApple()
drawAll = True 'Redraw the entire form
'Randomize the x and y coordinates of the apple so long as they're on screen
Randomize
Apple.Left = Int(Rnd * (Me.ScaleWidth - 50)) + 30
Apple.Top = Int(Rnd * (Me.ScaleHeight - 50)) + 30
Apple.Pic = Int(Rnd * 4)
Apple.Width = picApple(Apple.Pic).ScaleWidth
Apple.Height = picApple(Apple.Pic).ScaleHeight
'Make sure it is not drawn on black (worm) pixels
For X = Apple.Left To Apple.Left + Apple.Width
    For Y = Apple.Top To Apple.Top + Apple.Height
        If GetPixel(Me.hdc, X, Y) = &H0& Then 'If the apple intersects a black pixel
            Call CreateApple
            Exit Sub
        End If
    Next 'y
Next 'x
Apple.Visible = True
End Sub
Function DetectCollide(intX1 As Double, intY1 As Double, intWidth1 As Long, intHeight1 As Long, intX2 As Variant, intY2 As Variant, intWidth2 As Long, intHeight2 As Long) As Boolean
'Check if the rectangles intersect each other
If intX1 <= intX2 + intWidth2 And intX1 + intWidth1 >= intX2 And intY1 <= intY2 + intWidth2 And intY1 + intWidth1 >= intY2 Then
    DetectCollide = True
Else
    DetectCollide = False
End If
End Function
Sub LoadNextLevel()
drawAll = True
bRunning = False 'Pause the game
intScore = intScore + 16 'Increase your score
StopMidi (curMidi)
MsgBox "Level complete! Bonus: " & (16 * Multiplier) & " current score: " & (Multiplier * intScore)

'Change the midi
Randomize
Dim intRand As Integer
intRand = 0
'Load a new random song that is not the same song that just played
Do
    intRand = Int(Rnd * 4) + 1
Loop Until (intRand <> midiNum)
midiNum = intRand

SetMidi


PlayMidi (curMidi)

'Reset worm length
curDot = 1
curXAngle = 0
intGrowLength = initialGrowLength
DotsEaten = 0

If GameType = 0 Then 'Standard
    Me.Width = Me.Width - (40 * Screen.TwipsPerPixelX)
    If Me.Width < (117 * Screen.TwipsPerPixelX) Then 'Create a minimum width of 117 pixels
        Me.Width = (117 * Screen.TwipsPerPixelX)
        initialGrowLength = initialGrowLength + 50
    End If
ElseIf GameType = 2 Then 'Shrinking
    Me.Width = Me.Width + (65 * Screen.TwipsPerPixelX)
    Me.Height = Me.Height + (50 * Screen.TwipsPerPixelY)
End If



For i = 1 To intMaxLength 'Hide uncle worm
    Worm(i).Visible = False
    Worm(i).Left = Me.ScaleWidth / 2
    Worm(i).Top = Me.ScaleWidth / 2
Next 'i

'Draw a new apple
Call CreateApple

'Depress buttons
bRight = False
bLeft = False
WormLength = 1
bGrow = True

'Reset worm angle
curXAngle = intDiv / -2 + (intDiv / 10) - 1



'Set value of initial dot
Worm(1).Top = Me.ScaleHeight - 15
Worm(1).Left = Me.ScaleWidth / 2

bWin = False 'Haven't collected 10 apples yet
curLevel = curLevel + 1 'Increase level

'Unpause game
bRunning = True
Call GameLoop
End Sub
Sub ShrinkForm()
'Decrease the form size
Dim intShrink As Long 'how much the form shrinks by each time
If Difficulty = 0 Then
    intShrink = 1
ElseIf Difficulty = 1 Then
    intShrink = 2
Else
    intShrink = 3
End If
Me.Width = Me.Width - intShrink
Me.Height = Me.Height - intShrink
End Sub
Sub GameOver()
'End the game in a loss
If bRunning = True Then
    bRunning = False 'Pause the gameloop
    intScore = intScore * Multiplier 'Calculate score
    MsgBox "You lose!  Your score was " & intScore
    StopMidi (curMidi)
    If intScore > HighScore(GameType) Then 'New high score
        HighScore(GameType) = intScore
        MsgBox "You got a new highscore!"
        Call Encode(CStr(HighScore(GameType)), CStr(GameType), CStr(GameType) & "L", App.Path & "\scores.ini")
    Else
        MsgBox "You were unable to reach a high score of " & HighScore(GameType)
    End If
    Unload Me
End If
End Sub
Sub Shake()
'Move the form back and form and play a sound
If bShake = 1 Then
    Call PlySound("explosion")
End If
If bShake Mod 2 = 0 Then
    Me.Top = Me.Top + 50
    Me.Left = Me.Left + 50
Else
    Me.Top = Me.Top - 50
    Me.Left = Me.Left - 50
End If
bShake = bShake + 1
If bShake = 6 Then
    bShake = 0
End If

End Sub
Sub SetMidi()
'Sets curMidi to the correct value
If Speed = 1 Then
    curMidi = "slow" & midiNum
ElseIf Speed = 2 Then
    curMidi = "med" & midiNum
ElseIf Speed = 3 Then
    curMidi = "fast" & midiNum
Else
    curMidi = "vfast" & midiNum
End If
End Sub
Function wormCollide(X As Double, Y As Double) As Boolean
If GetPixel(Me.hdc, X, Y) = &H0& Then
    wormCollide = True
Else
    wormCollide = False
End If
End Function
