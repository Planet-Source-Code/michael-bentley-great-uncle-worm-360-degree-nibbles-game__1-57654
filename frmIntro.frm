VERSION 5.00
Begin VB.Form frmIntro 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Great Uncle Worm"
   ClientHeight    =   5970
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6615
   Icon            =   "frmIntro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   398
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   441
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdHelp 
      Caption         =   "&Help"
      Height          =   495
      Left            =   3240
      TabIndex        =   33
      Top             =   5280
      Width           =   975
   End
   Begin VB.Timer timeMidi 
      Interval        =   35
      Left            =   120
      Top             =   240
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "&About"
      Height          =   495
      Left            =   4320
      TabIndex        =   29
      Top             =   5280
      Width           =   975
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   5400
      TabIndex        =   30
      Top             =   5280
      Width           =   975
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "&Start"
      Height          =   495
      Left            =   2160
      TabIndex        =   28
      Top             =   5280
      Width           =   975
   End
   Begin VB.Frame framControl 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Control Style"
      Height          =   615
      Left            =   0
      TabIndex        =   20
      Top             =   4560
      Width           =   6615
      Begin VB.OptionButton opControl 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Super Glued"
         Height          =   255
         Index           =   3
         Left            =   2640
         TabIndex        =   34
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton opControl 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tight"
         Height          =   255
         Index           =   2
         Left            =   1920
         TabIndex        =   23
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton opControl 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Medium"
         Height          =   255
         Index           =   1
         Left            =   960
         TabIndex        =   22
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton opControl 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Loose"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame framDifficulty 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Difficulty (Depends on Game Type)"
      Height          =   615
      Left            =   0
      TabIndex        =   15
      Top             =   3720
      Width           =   6615
      Begin VB.OptionButton opDif 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Hard"
         Height          =   255
         Index           =   2
         Left            =   1920
         TabIndex        =   18
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton opDif 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Medium"
         Height          =   255
         Index           =   1
         Left            =   960
         TabIndex        =   17
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton opDif 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Easy"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame framSize 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Arena Starting Size"
      Height          =   615
      Left            =   0
      TabIndex        =   11
      Top             =   2880
      Width           =   6615
      Begin VB.OptionButton opSize 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Large"
         Height          =   255
         Index           =   2
         Left            =   1920
         TabIndex        =   14
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton opSize 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Medium"
         Height          =   255
         Index           =   1
         Left            =   960
         TabIndex        =   13
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton opSize 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Small"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame framSpeed 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Speed"
      Height          =   615
      Left            =   0
      TabIndex        =   6
      Top             =   2040
      Width           =   6615
      Begin VB.OptionButton opSpeed 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Very Fast"
         Height          =   255
         Index           =   3
         Left            =   2880
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton opSpeed 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Fast"
         Height          =   255
         Index           =   2
         Left            =   2040
         TabIndex        =   9
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton opSpeed 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Medium"
         Height          =   255
         Index           =   1
         Left            =   960
         TabIndex        =   8
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton opSpeed 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Slow"
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   7
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame framGameType 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Game Type"
      Height          =   1095
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   6615
      Begin VB.OptionButton opGameType 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Survival"
         Height          =   255
         Index           =   5
         Left            =   0
         TabIndex        =   35
         Top             =   600
         Width           =   1095
      End
      Begin VB.OptionButton opGameType 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Daggar Drop"
         Height          =   255
         Index           =   4
         Left            =   4920
         TabIndex        =   19
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton opGameType 
         BackColor       =   &H00FFFFFF&
         Caption         =   "The Diet of Worms"
         Height          =   255
         Index           =   3
         Left            =   3000
         TabIndex        =   5
         Top             =   240
         Width           =   1695
      End
      Begin VB.OptionButton opGameType 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Shrinking"
         Height          =   255
         Index           =   2
         Left            =   2040
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton opGameType 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Endless"
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton opGameType 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Standard"
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   2
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.Label lblGen 
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.41"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   5520
      TabIndex        =   32
      Top             =   5760
      Width           =   1095
   End
   Begin VB.Label lblDocEnt 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Doc Entertainment's"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   1200
      TabIndex        =   31
      Top             =   0
      Width           =   3975
   End
   Begin VB.Label lblMult 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   27
      Top             =   5520
      Width           =   495
   End
   Begin VB.Label lblHighScore 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   26
      Top             =   5280
      Width           =   855
   End
   Begin VB.Label lblGen 
      BackStyle       =   0  'Transparent
      Caption         =   "High Score Multiplier:"
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   25
      Top             =   5520
      Width           =   1695
   End
   Begin VB.Label lblGen 
      BackStyle       =   0  'Transparent
      Caption         =   "Current High Score:"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   24
      Top             =   5280
      Width           =   1455
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Great Uncle Worm"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6375
   End
End
Attribute VB_Name = "frmIntro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAbout_Click()
MsgBox "Great Uncle Worm Version 1.41" & vbNewLine & "Programmed By Mike Bentley" & vbNewLine & "Tested By Tacvek" & vbNewLine & "Inspired By the TI-83 Game" & vbNewLine & "Developed By Doc Entertainment" & vbNewLine & "Visit Our Site: http://www.doc-ent.com", vbInformation, "About"

End Sub

Private Sub cmdExit_Click()
End
End Sub

Private Sub cmdHelp_Click()
frmHelp.Show
End Sub

Private Sub cmdStart_Click()
If Multiplier >= 1 Then 'Don't let the user play with a multiplier less than 1
    Dim strSave As String
    'Save user settings
    strSave = App.Path & "\scores.ini"
    Call WriteIni("GEN", "DIF", CStr(Difficulty), strSave)
    Call WriteIni("GEN", "GAMETYPE", CStr(GameType), strSave)
    Call WriteIni("GEN", "SPEED", CStr(Speed), strSave)
    Call WriteIni("GEN", "SIZE", CStr(Size), strSave)
    Call WriteIni("GEN", "CONTROL", CStr(Control), strSave)
    bLoaded = False 'The game has not loaded yet
    frmWorm.Show
    Me.Hide
Else
    MsgBox "Settings too easy, please make the multiplier atleast equal to 1."
End If
End Sub

Private Sub Form_Activate()
For i = 0 To 3
    If opSpeed(i).Value = True Then
        Speed = i
    End If
Next 'i
For i = 0 To 4
    If opGameType(i).Value = True Then
        lblHighScore.Caption = HighScore(i)
        GameType = i
    End If
Next 'i
For i = 0 To 2
    If opDif(i).Value = True Then
        Difficulty = i
    End If
Next 'i
For i = 0 To 2
    If opSize(i).Value = True Then
        Size = i
    End If
Next 'i
For i = 0 To 2
    If opControl(i).Value = True Then
        Control = i
    End If
Next 'i
Call CalculateMult

End Sub

Private Sub Form_Load()
On Error Resume Next
Dim strHighScore As String
'Load high scores for each game mode
For i = 0 To opGameType.UBound
    strHighScore = DecryptString(App.Path & "\scores.ini", CStr(i), CStr(i) & "L")
    If strHighScore = "" Then
        strHighScore = "0"
    End If
    HighScore(i) = CInt(strHighScore)
Next 'i
'Load last game settings from the high score file
Dim strTemp As String
strTemp = GetFromIni("GEN", "DIF", App.Path & "\scores.ini")
If strTemp = "" Then strTemp = "1"
i = CInt(strTemp)
opDif(i).Value = True
Difficulty = i

strTemp = GetFromIni("GEN", "SPEED", App.Path & "\scores.ini")
If strTemp = "" Then strTemp = "1"
i = CInt(strTemp)
opSpeed(i).Value = True
Speed = i

strTemp = GetFromIni("GEN", "CONTROL", App.Path & "\scores.ini")
If strTemp = "" Then strTemp = "1"
i = CInt(strTemp)
opControl(i).Value = True
Control = i

strTemp = GetFromIni("GEN", "SIZE", App.Path & "\scores.ini")
If strTemp = "" Then strTemp = "1"
opSize(i).Value = True
i = CInt(strTemp)
Size = i

strTemp = GetFromIni("GEN", "GAMETYPE", App.Path & "\scores.ini")
If strTemp = "" Then strTemp = "0"
i = CInt(strTemp)
opGameType(i).Value = True
GameType = i

End Sub

Private Sub opControl_Click(Index As Integer)
Control = Index
Call CalculateMult
End Sub

Private Sub opDif_Click(Index As Integer)
Difficulty = Index
Call CalculateMult
End Sub

Private Sub opGameType_Click(Index As Integer)
GameType = Index
Call CalculateMult
End Sub

Private Sub opSize_Click(Index As Integer)
Size = Index
Call CalculateMult
End Sub

Private Sub opSpeed_Click(Index As Integer)
Speed = Index
Call CalculateMult
End Sub
Sub CalculateMult()
'Calculate the multiplier based on speed, level size and difficulty
If GameType = 6 Then 'Survival
    Multiplier = 1 + Speed - Size
ElseIf GameType = 3 Then 'Diet of Worms
    Multiplier = Difficulty - (2 - Size) + Speed
Else
    Multiplier = Difficulty - Size + Speed
End If

lblMult.Caption = Multiplier
lblHighScore.Caption = HighScore(GameType)
End Sub

