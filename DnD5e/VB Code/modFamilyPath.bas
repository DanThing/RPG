Attribute VB_Name = "modFamilyPath"
Option Explicit
Option Compare Text

Dim lifestyle As String
Dim father As Variant
Dim mother As Variant
Dim rMod As Long

Sub RollForFamily()

lifestyle = getLifestyle & vbNewLine

pc.History =

End Sub

Private Function getChildhood() As String
Dim r As Long
r = Roll(20)
Select Case r
Case Is < 4: getChildhood = "You are still haunted by your childhood, where you were treated badly by your peers."
Case Is < 6: getChildhood = "You spent most of your childhood alone, with no close friends."
Case Is < 9: getChildhood = "Others saw you as different or strange, and so you had few companions."
Case Is < 13: getChildhood = "You had a few close friends and lived an ordinary childhood."
Case Is < 16: getChildhood = "You had several friends, and your childhood was generally a happy one."
Case Is < 18: getChildhood = "You always found it easy to make friends and you loved being around people."
Case Else: getChildhood = "Everyone knew who you were, and you had friends everywhere you went."
End Select
End Function

Private Function getStatusof() As String
Dim r As Long
r = Roll(20)
Select Case r
Case Is < 4: getStatusof = "is dead. Cause of death " & CauseofDeath
Case Is < 6: getStatusof = "is missing or unknown."
Case Is < 9: getStatusof = "is alive, but doing poorly due to injury, financial trouble, or relationship difficulties."
Case Is < 13: getStatusof = "is alive and well."
Case Is < 16: getStatusof = "alive and quite successful."
Case Is < 18: getStatusof = "alive and infamous."
Case Is > 18: getStatusof = "alive and famous."
End Select
End Function

Private Function getRelationship() As String
Dim r As Long
r = Roll(12)
Select Case r
Case Is < 5: getRelationship = "You are hostile towards each other."
Case Is < 11: getRelationship = "You get on with each other."
Case Is > 11: getRelationship = "You are indifferent to each other."
End Select
End Function

Private Function getOccpation() As String
Dim r As Long
r = Roll(100) + rMod
Select Case r
    Case Is < 6:     getOccpation = "is an exile, hermit or refugee"
    Case Is < 11:    getOccpation = "is a criminal"
    Case Is < 12:    getOccpation = "works as a laborer"
    Case Is < 27:    getOccpation = "is an explorer or wanderer"
    Case Is < 32:    getOccpation = "works as a hunter or trapper"
    Case Is < 37:    getOccpation = "works as a farmer or herder"
    Case Is < 39:    getOccpation = "works as an adventuring  " & getClass
    Case Is < 44:    getOccpation = "works as an entertainer"
    Case Is < 56:    getOccpation = "works as a priest"
    Case Is < 61:    getOccpation = "works as a sailor"
    Case Is < 76:    getOccpation = "serves as a soldier"
    Case Is < 81:    getOccpation = "works as an academic"
    Case Is < 86:    getOccpation = "works as a merchant"
    Case Is < 91:    getOccpation = "works as an artisan or guild member"
    Case Is < 96:    getOccpation = "works as a politician or bureaucrat"
    Case Else:       getOccpation = "is an aristocrat"
End Select
End Function

Private Function getRace(who_of As String) As String
Dim r As Long
Select Case True
    Case PC.Race Like "half*"
      r = Roll(4)
      Select Case True
        Case PC.Race Like "*elf*"
            Select Case r
                Case Is < 3: getRace = "an Elf"
                Case Is < 5: getRace = "a Human"
            End Select
        Case PC.Race Like "*orc*"
            Select Case r
                Case Is < 3: getRace = "an Orc"
                Case Is < 5: getRace = "a Human"
            End Select
        End Select
    Case Else
        getRace = "your race"
End Select
End Function

Private Function getSibling() As String
Dim r As Long, d As Long

r = Roll(4)
Select Case r
    Case 1: getSibling = "older male"
    Case 2: getSibling = "older female"
    Case 3: getSibling = "younger male"
    Case 4: getSibling = "younger female"
    Case Else
        Err.Raise CustomError.Err9("Get Siblings", "Unable to get siblings.")
End Function

Private Function isAbsent() As String
Dim r As Long
r = Roll(4)
Select Case r
    Case 1: isAbsent = "is dead."
    Case 2: isAbsent = "was imprisoned, enslaved, or otherwise taken away."
    Case 3: isAbsent = "abandoned you."
    Case 4: isAbsent = "disappeared to an unknown fate."
End Select
End Function

Private Function getRaisedBy() As String
Dim r As Long
    r = Roll(100) + rMod
      Select Case r
        Case r < 2
            getRaisedBy = "nobody"
            r = Roll(4)
            Select Case r
                Case r < 3: mother = "Your mother " & isAbsent
                Case r < 5: father = "Your father " & isAbsent
            End Select
        Case r < 3
            getRaisedBy = "an institution, such as an asylum"
            r = Roll(4)
            Select Case r
                Case r < 3: mother = "Your mother " & isAbsent
                Case r < 5: father = "Your father " & isAbsent
            End Select
        Case r < 4
            getRaisedBy = "a temple"
            r = Roll(4)
            Select Case r
                Case r < 3: mother = "Your mother " & isAbsent
                Case r < 5: father = "Your father " & isAbsent
            End Select
        Case r < 6: getRaisedBy = "an orphanage"
            r = Roll(4)
            Select Case r
                Case r < 3: mother = "Your mother " & isAbsent
                Case r < 5: father = "Your father " & isAbsent
            End Select
        Case r < 8:
            getRaisedBy = "a guardian"
            r = Roll(4)
            Select Case r
                Case r < 3: mother = "Your mother " & isAbsent
                Case r < 5: father = "Your father " & isAbsent
            End Select
        Case r < 16
            getRaisedBy = "your paternal or maternal aunt, uncle, or both; or extended family such as a tribe or clan"
            r = Roll(4)
            Select Case r
                Case r < 3: mother = "Your mother " & isAbsent
                Case r < 5: father = "Your father " & isAbsent
            End Select
        Case r < 26
            getRaisedBy = "your paternal or maternal grandparent(s)"
                        r = Roll(4)
            Select Case r
                Case r < 3: mother = "Your mother " & isAbsent
                Case r < 5: father = "Your father " & isAbsent
            End Select
        Case r < 36
            getRaisedBy = "your adoptive family (same or different race)"
            r = Roll(4)
            Select Case r
                Case r < 3: mother = "Your mother " & isAbsent
                Case r < 5: father = "Your father " & isAbsent
            End Select
        Case r < 56
            getRaisedBy = "your single father or step father"
            Case r < 3: mother = "Your mother " & isAbsent
        Case r < 76
            getRaisedBy = "your single mother or step mother"
            Case r < 5: father = "Your father " & isAbsent
        Case r > 76
            getRaisedBy = "your mother and father"
        Case Else
            Err.Raise CustomError.Err9("Get Raisedby", "Unable to get raised by.")
    End Select
End Function

Private Function getLifestyle() As String
Dim d As Long
    d = Roll(6, 3)
    Select Case d
        Case Is < 4
            getLifestyle = "Wretched"
            rMod = -40
        Case Is < 6
            getLifestyle = "Sqaulid"
            rMod = -20
        Case Is < 9
            getLifestyle = "Poor"
            rMod = -10
        Case Is < 13
            getLifestyle = "Modest"
            rMod = 0
        Case Is < 16
            getLifestyle = "Comfortable"
            rMod = 10
        Case Is < 18
            getLifestyle = "Wealthy"
            rMod = 20
        Case Is < 19
            getLifestyle = "Aristocratic"
            rMod = 40
        Case Else
            Err.Raise CustomError.Err9("Get Lifestyle", "Unable to get lifestyle.")
    End Select
End Function

Private Function getBorn() As String
Dim d As Long
    d = Roll(100) + rMod
    Select Case d

        Case Else
            Err.Raise CustomError.Err9("Get Birth Place", "Unable to get birth place.")
    End Select
End Function

Private Function getHome() As String
Dim d As Long
    d = Roll(100) + rMod
    Select Case d
        Case Is < 11: getHome = " on the streets"
        Case Is < 21: getHome = " in a rundown shack"
        Case Is < 31: getHome = " in lots of places, never staying still for long"
        Case Is < 41: getHome = " in an encampment or village in the wilderness"
        Case Is < 51: getHome = " in an apartment in a rundown neighbourhood"
        Case Is < 71: getHome = " in a small house"
        Case Is < 81: getHome = " in a large house"
        Case Is < 91: getHome = " in a mansion"
        Case Is > 91: getHome = " in a palace or castle"
        Case Else
            Err.Raise CustomError.Err9("Get Home", "Unable to get home.")
    End Select
End Function


