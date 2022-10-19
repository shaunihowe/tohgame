Attribute VB_Name = "DataMod"
Public HighScores(5 To 8) As New board
Public LeastMoves(5 To 8) As Integer

Public Sub IniScores()
LeastMoves(5) = 31
LeastMoves(6) = 63
LeastMoves(7) = 127
LeastMoves(8) = 255
For a = 5 To 8
    HighScores(a).HighScoreFile = "c:\windows\tohgame" & Trim(Str(a)) & ".dat"
    If FileExist(HighScores(a).HighScoreFile) = True Then
        HighScores(a).LoadHighScoreData
    Else
        HighScores(a).NewHighScoreData True
        HighScores(a).SaveHighScoreData
    End If
Next a
End Sub

Public Function CalculateScore(ByVal GameSize As Integer, ByVal Moves As Integer, ByVal SecondsTaken As Integer) As Integer
CalculateScore = ((SecondsTaken * 13) / LeastMoves(GameSize)) * Moves
End Function
