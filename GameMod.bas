Attribute VB_Name = "GameMod"
Public Const Status_ChooseSrc As Byte = 0
Public Const Status_ChooseDes As Byte = 1

Public Type PileType
    Slabs As Integer
    slab(8) As Integer
End Type

Public Type GameType
    GameSize As Integer
    Moves As Integer
    Startslab As Integer
    Pile(3) As PileType
    Status As Byte
    CurSlab As Integer
    Src As Integer
    Des As Integer
    StartTime As Single
End Type
Public Current_Game As GameType

Public Sub NewGame(ByVal GameSize As Integer)
With Current_Game
    .GameSize = GameSize
    .Moves = 0
    .Startslab = 1
    For a = 2 To 3
    For b = 1 To 8
        .Pile(a).Slabs = 0
    Next b
    Next a
    .Pile(1).Slabs = GameSize
    For a = 1 To GameSize
        .Pile(1).slab(a) = ((GameSize + 1) - a)
    Next a
    .Status = Status_ChooseSrc
    .StartTime = Timer
    mainform.selectbutt(1).Enabled = True
    mainform.selectbutt(2).Enabled = False
    mainform.selectbutt(3).Enabled = False
    mainform.letter(1).Enabled = True
    mainform.letter(2).Enabled = False
    mainform.letter(3).Enabled = False
    mainform.timebartim.Enabled = True
End With
End Sub

Public Sub Update_Form()
On Error Resume Next
For a = 1 To 8
    With mainform.slab(a)
    .Visible = False
    End With
Next a
If Current_Game.Pile(1).Slabs > 0 Then
    For a = 1 To Current_Game.Pile(1).Slabs
        CurSlab = Current_Game.Pile(1).slab(a)
        mainform.slab(CurSlab).Left = (1158 - (mainform.slab(CurSlab).Width / 2))
        mainform.slab(CurSlab).Top = 2040 - (mainform.slab(CurSlab).Height * (a - 1))
        mainform.slab(CurSlab).Visible = True
    Next a
Else
    mainform.selectbutt(1).Enabled = False
    mainform.letter(1).Enabled = False
End If
If Current_Game.Pile(2).Slabs > 0 Then
    For a = 1 To Current_Game.Pile(2).Slabs
        CurSlab = Current_Game.Pile(2).slab(a)
        mainform.slab(CurSlab).Left = (3473 - (mainform.slab(CurSlab).Width / 2))
        mainform.slab(CurSlab).Top = 2040 - (mainform.slab(CurSlab).Height * (a - 1))
        mainform.slab(CurSlab).Visible = True
    Next a
Else
    mainform.selectbutt(2).Enabled = False
    mainform.letter(2).Enabled = False
End If
If Current_Game.Pile(3).Slabs > 0 Then
    For a = 1 To Current_Game.Pile(3).Slabs
        CurSlab = Current_Game.Pile(3).slab(a)
        mainform.slab(CurSlab).Left = (5788 - (mainform.slab(CurSlab).Width / 2))
        mainform.slab(CurSlab).Top = 2040 - (mainform.slab(CurSlab).Height * (a - 1))
        mainform.slab(CurSlab).Visible = True
    Next a
Else
    mainform.selectbutt(3).Enabled = False
    mainform.letter(3).Enabled = False
End If
mainform.statusbar.Caption = "Moves Taken " & Current_Game.Moves & ", Game Possible in " & LeastMoves(Current_Game.GameSize) & " Moves."
End Sub

Public Sub Condense(ByVal game As Integer)
Dim Score As New board
Dim Newscore As New board
Dim Done(31) As String
Dim numnames As Integer
Dim nxt As Boolean
Score.HighScoreFile = "C:\WINDOWS\tohgame" & Trim(Str(game)) & ".dat"
Score.LoadHighScoreData
Newscore.NewHighScoreData True
Newscore.BestLow = True
If Score.NumberOfEntrys = 0 Then Exit Sub
numnames = 0
For a = 1 To 30
    Done(a) = ""
Next a
Do
    b = b + 1
    curname = Score.Name(b)
    nxt = False
    For a = 1 To 30
        If LSmooth(curname) = LSmooth(Done(a)) Then nxt = True: Exit For
    Next a
    If nxt = False Then
        numnames = numnames + 1
        Done(numnames) = curname
        Newscore.AddEntry curname, Score.Score(b)
    End If
Loop Until b = Score.NumberOfEntrys
Newscore.HighScoreFile = Score.HighScoreFile
Newscore.SaveHighScoreData
End Sub
