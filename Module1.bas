Attribute VB_Name = "Module1"
Option Explicit

Public Type SCard
    iOnes As Integer
    iTwos As Integer
    iThrees As Integer
    iFours As Integer
    iFives As Integer
    iSixes As Integer
    iBonus As Integer
    iTOAK As Integer
    iFOAK As Integer
    iLowStrait As Integer
    iHiStrait As Integer
    iFullHouse As Integer
    iChance As Integer
    iNumYartcees As Integer
End Type
Public CurrentSC As SCard
Public iDice(5) As Integer
Public iCurrentRoll As Integer
Public HiScores(2, 10) As String
Function ResetHiScores()
Dim x As Integer
    For x = 1 To 10
        SaveSetting App.Title, "HiScores", "NAME" & Format(x, "00"), "Yartcee"
        SaveSetting App.Title, "HiScores", "SCORE" & Format(x, "00"), CStr((11 - x) * 33)
    Next x
End Function
Sub GetHiScores()
Dim x As Integer
    If GetSetting(App.Title, "HiScores", "NAME01") = "" Then
        ResetHiScores
    End If
    For x = 1 To 10
        HiScores(0, x - 1) = GetSetting(App.Title, "HiScores", "NAME" & Format(x, "00"))
        HiScores(1, x - 1) = GetSetting(App.Title, "HiScores", "SCORE" & Format(x, "00"))
    Next x
End Sub
Function HowMany(iIn As Integer) As Integer
Dim x As Integer
    HowMany = 0
    For x = 0 To 4
    If iDice(x) = iIn Then
        HowMany = HowMany + 1
        End If
    Next x
End Function
Function IsThereA(iIn As Integer) As Boolean
Dim x As Integer
    IsThereA = False
    For x = 0 To 4
        If iDice(x) = iIn Then
            IsThereA = True
            Exit Function
        End If
    Next x
End Function
Function IsItAHiStrait() As Boolean
    If (IsThereA(1) And IsThereA(2) And IsThereA(3) And IsThereA(4) And IsThereA(5)) Or (IsThereA(2) And IsThereA(3) And IsThereA(4) And IsThereA(5) And IsThereA(6)) Then
        IsItAHiStrait = True
    Else
        IsItAHiStrait = False
    End If
End Function
Function AddEmUp() As Integer
    AddEmUp = HowMany(1) + (HowMany(2) * 2) + (HowMany(3) * 3) + (HowMany(4) * 4) + _
            (HowMany(5) * 5) + (HowMany(6) * 6)
End Function
Function OfAKind(iIn As Integer) As Boolean
Dim x As Integer
    OfAKind = False
    For x = 1 To 6
        If HowMany(x) >= iIn Then
            OfAKind = True
            Exit Function
        End If
    Next x
End Function
Function IsItALoStrait() As Boolean
    If (IsThereA(1) And IsThereA(2) And IsThereA(3) And IsThereA(4)) Or (IsThereA(2) And IsThereA(3) And IsThereA(4) And IsThereA(5)) Or (IsThereA(3) And IsThereA(4) And IsThereA(5) And IsThereA(6)) Then
        IsItALoStrait = True
    Else
        IsItALoStrait = False
    End If
End Function
Function IsThereAYartcee() As Boolean
Dim x As Integer
    IsThereAYartcee = False
    For x = 1 To 6
        If HowMany(x) = 5 Then
            IsThereAYartcee = True
            Exit Function
        End If
    Next x
End Function
Function IsThereAFullFouse() As Boolean
Dim ThreeOAK As Boolean
Dim ThreeOAKnum As Boolean
Dim x As Integer
    IsThereAFullFouse = False
    For x = 1 To 5
        If HowMany(x) = 3 Then
            ThreeOAK = True
            ThreeOAKnum = x
            Exit For
        End If
    Next x
    If Not ThreeOAK Then Exit Function
    For x = 1 To 5
        If HowMany(x) = 2 And ThreeOAKnum <> x Then
            IsThereAFullFouse = True
            Exit Function
        End If
    Next x
End Function
Function Yartceescore(iIn As Integer) As Integer
    Yartceescore = (iIn ^ 2) * 50
End Function
Sub ResetScorecard()
    With CurrentSC
        .iOnes = 0
        .iTwos = 0
        .iThrees = 0
        .iFours = 0
        .iFives = 0
        .iSixes = 0
        .iBonus = 0
        .iTOAK = 0
        .iFOAK = 0
        .iLowStrait = 0
        .iHiStrait = 0
        .iFullHouse = 0
        .iChance = 0
        .iNumYartcees = 0
    End With
End Sub

