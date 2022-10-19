VERSION 5.00
Begin VB.Form mainform 
   BackColor       =   &H00808000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tower Of Hanoi - Created by Shaun Howe"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   6855
   Icon            =   "mainform.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   6855
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer timebartim 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   6120
      Top             =   2760
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Press F5,F6,F7 or F8 to play"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3015
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   6615
   End
   Begin VB.Label letter 
      BackStyle       =   0  'Transparent
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   1935
      Index           =   3
      Left            =   4800
      TabIndex        =   8
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label letter 
      BackStyle       =   0  'Transparent
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   1935
      Index           =   2
      Left            =   2520
      TabIndex        =   7
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label timebar 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   5520
      TabIndex        =   6
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Label selectbutt 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
      Caption         =   "Select Source"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   3
      Left            =   4800
      TabIndex        =   3
      Top             =   2400
      Width           =   1815
      WordWrap        =   -1  'True
   End
   Begin VB.Label selectbutt 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
      Caption         =   "Select Source"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   2
      Left            =   2520
      TabIndex        =   2
      Top             =   2400
      Width           =   1815
      WordWrap        =   -1  'True
   End
   Begin VB.Label selectbutt 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
      Caption         =   "Select Source"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   2400
      Width           =   1815
      WordWrap        =   -1  'True
   End
   Begin VB.Label statusbar 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   3240
      Width           =   5535
   End
   Begin VB.Shape slab 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   255
      Index           =   8
      Left            =   240
      Shape           =   2  'Oval
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Shape slab 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   255
      Index           =   7
      Left            =   360
      Shape           =   2  'Oval
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Shape slab 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   255
      Index           =   6
      Left            =   480
      Shape           =   2  'Oval
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Shape slab 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   255
      Index           =   5
      Left            =   600
      Shape           =   2  'Oval
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Shape slab 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   255
      Index           =   4
      Left            =   720
      Shape           =   2  'Oval
      Top             =   1080
      Width           =   855
   End
   Begin VB.Shape slab 
      BackColor       =   &H00C0C000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   255
      Index           =   3
      Left            =   840
      Shape           =   2  'Oval
      Top             =   840
      Width           =   615
   End
   Begin VB.Shape slab 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   255
      Index           =   2
      Left            =   960
      Shape           =   2  'Oval
      Top             =   600
      Width           =   375
   End
   Begin VB.Shape slab 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   255
      Index           =   1
      Left            =   1080
      Shape           =   2  'Oval
      Top             =   360
      Width           =   135
   End
   Begin VB.Label letter 
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   1935
      Index           =   1
      Left            =   240
      TabIndex        =   4
      Top             =   360
      Width           =   1815
   End
   Begin VB.Menu menu_file 
      Caption         =   "&File"
      Begin VB.Menu menu_file_new 
         Caption         =   "&New"
         Begin VB.Menu menu_file_new_ 
            Caption         =   "&5 Slab Game (Easy)"
            Index           =   5
            Shortcut        =   {F5}
         End
         Begin VB.Menu menu_file_new_ 
            Caption         =   "&6 Slab Game (Medium)"
            Index           =   6
            Shortcut        =   {F6}
         End
         Begin VB.Menu menu_file_new_ 
            Caption         =   "&7 Slab Game (Hard)"
            Index           =   7
            Shortcut        =   {F7}
         End
         Begin VB.Menu menu_file_new_ 
            Caption         =   "&8 Slab Game (Expert)"
            Index           =   8
            Shortcut        =   {F8}
         End
      End
      Begin VB.Menu menu_file_exit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu menu_highscore 
      Caption         =   "&Highscore"
      Begin VB.Menu menu_highscore_show 
         Caption         =   "&Show Highscores"
         Shortcut        =   ^S
      End
      Begin VB.Menu menu_highscore_rankings 
         Caption         =   "Show &Rankings"
         Shortcut        =   ^R
      End
   End
End
Attribute VB_Name = "mainform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case 49, 50, 51: If letter(KeyAscii - 48).Enabled = True Then letter_Click KeyAscii - 48
Case Else
End Select
End Sub

Private Sub Form_Load()
IniScores
End Sub

Private Sub letter_Click(Index As Integer)
selectbutt_Click Index
End Sub

Private Sub menu_file_exit_Click()
End
End Sub

Private Sub menu_file_new__Click(Index As Integer)
mainform.Label1.Visible = False
NewGame Index
Update_Form
End Sub

Private Sub menu_highscore_rankings_Click()
Dim sco As New board
Dim numnames As Integer
Dim names(300) As String
Dim scos(300) As Integer
Dim curname As String
Dim mat As Boolean
Dim stsco(5 To 8) As Integer
numnames = 0
sco.NewHighScoreData False
stsco(5) = 900
stsco(6) = (stsco(5) * 2) + (stsco(5) / 31)
stsco(7) = (stsco(6) * 2) + (stsco(6) / 63)
stsco(8) = (stsco(7) * 2) + (stsco(7) / 127)
For a = 5 To 8
    HighScores(a).SaveHighScoreData
    Condense a
    HighScores(a).LoadHighScoreData
    For b = 1 To HighScores(a).NumberOfEntrys
        curname = HighScores(a).Name(b)
        mat = False
        For c = 1 To numnames
            If LSmooth(curname) = LSmooth(names(c)) Then
                mat = True
                scos(c) = scos(c) + (stsco(a) - HighScores(a).Score(b))
            End If
        Next c
        If mat = False Then
            numnames = numnames + 1
            names(numnames) = curname
            scos(numnames) = scos(numnames) + (stsco(a) - HighScores(a).Score(b))
        End If
    Next b
Next a
For a = 1 To numnames
    sco.AddEntry names(a), scos(a)
Next a
lmax = sco.NumberOfEntrys
If lmax > 30 Then lmax = 30
For a = 1 To lmax
    hig$ = hig$ & a & "   " & sco.Name(a) & "      " & sco.Score(a) & vbCrLf
Next a
MsgBox hig$, vbInformation, "Player Rankings"
End Sub

Private Sub menu_highscore_show_Click()
lmax = HighScores(5).NumberOfEntrys
If lmax > 10 Then lmax = 10
hig$ = hig$ & "5 slab game (easy)..." & vbCrLf
For a = 1 To lmax
    hig$ = hig$ & a & "   " & HighScores(5).Name(a) & "      " & HighScores(5).Score(a) & vbCrLf
Next a
lmax = HighScores(6).NumberOfEntrys
If lmax > 10 Then lmax = 10
hig$ = hig$ & vbCrLf & "6 slab game (medium)..." & vbCrLf
For a = 1 To lmax
    hig$ = hig$ & a & "   " & HighScores(6).Name(a) & "      " & HighScores(6).Score(a) & vbCrLf
Next a
lmax = HighScores(7).NumberOfEntrys
If lmax > 10 Then lmax = 10
hig$ = hig$ & vbCrLf & "7 slab game (hard)..." & vbCrLf
For a = 1 To lmax
    hig$ = hig$ & a & "   " & HighScores(7).Name(a) & "      " & HighScores(7).Score(a) & vbCrLf
Next a
lmax = HighScores(8).NumberOfEntrys
If lmax > 10 Then lmax = 10
hig$ = hig$ & vbCrLf & "8 slab game (expert)..." & vbCrLf
For a = 1 To lmax
    hig$ = hig$ & a & "   " & HighScores(8).Name(a) & "      " & HighScores(8).Score(a) & vbCrLf
Next a
MsgBox hig$, vbInformation, "High Scores"
End Sub

Private Sub selectbutt_Click(Index As Integer)
Select Case Current_Game.Status
Case Status_ChooseSrc
    Current_Game.Src = Index
    Current_Game.CurSlab = Current_Game.Pile(Index).slab(Current_Game.Pile(Index).Slabs)
    For a = 1 To 3
        If a <> Index Then mainform.selectbutt(a).Enabled = False: mainform.letter(a).Enabled = False
        If (Current_Game.Pile(a).slab(Current_Game.Pile(a).Slabs) > Current_Game.CurSlab) Or (Current_Game.Pile(a).Slabs = 0) Then
            mainform.selectbutt(a).Caption = "Place Here"
            mainform.selectbutt(a).Enabled = True
            mainform.letter(a).Enabled = True
        End If
    Next a
    Current_Game.Status = Status_ChooseDes
Case Status_ChooseDes
    Current_Game.Des = Index
    
    If Current_Game.Src <> Current_Game.Des Then
        With Current_Game
            .Pile(.Src).Slabs = .Pile(.Src).Slabs - 1
            .Pile(.Des).Slabs = .Pile(.Des).Slabs + 1
            .Pile(.Des).slab(.Pile(.Des).Slabs) = .CurSlab
        End With
        Current_Game.Moves = Current_Game.Moves + 1
    End If
    For a = 1 To 3
        mainform.selectbutt(a).Caption = "Select Source"
        mainform.selectbutt(a).Enabled = True
        mainform.letter(a).Enabled = True
    Next a
    Current_Game.Status = Status_ChooseSrc
    Update_Form
    If (Current_Game.Pile(2).Slabs = Current_Game.GameSize) Or (Current_Game.Pile(3).Slabs = Current_Game.GameSize) Then
        GameTime = Timer - Current_Game.StartTime
        timebartim.Enabled = False
        Score = CalculateScore(Current_Game.GameSize, Current_Game.Moves, GameTime)
        NName = InputBox("You have a score of " & Score & ", What is your Name?", "Name Entry", "NoName")
        HighScores(Current_Game.GameSize).AddEntry NName, Score
        HighScores(Current_Game.GameSize).SaveHighScoreData
        lmax = HighScores(Current_Game.GameSize).NumberOfEntrys
        If lmax > 20 Then lmax = 20
        For a = 1 To lmax
            hig$ = hig$ & a & "   " & HighScores(Current_Game.GameSize).Name(a) & "      " & HighScores(Current_Game.GameSize).Score(a) & vbCrLf
        Next a
        mainform.timebartim.Enabled = False
        MsgBox hig$, vbInformation, "Top 20 High Scores for " & Current_Game.GameSize & " slabs"
        NewGame Current_Game.GameSize
        Update_Form
    End If
End Select
End Sub

Private Sub timebartim_Timer()
timebar.Caption = "Time: " & Int(Timer - Current_Game.StartTime)
End Sub
