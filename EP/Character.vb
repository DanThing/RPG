Public Class Character

    Public Name As String

    Public Background As String

    Public GenderId As String

    Public ActualAge As Long

    Public RezPoints As Long

    Public pp As Long

    Public Motivations As Array

    Public Aptitudes As IDictionary

    Public Stats As IDictionary

    Public Sub New()

        pp = 10

        Aptitudes.Add("cog", 0)
        Aptitudes.Add("coo", 0)
        Aptitudes.Add("int", 0)
        Aptitudes.Add("ref", 0)
        Aptitudes.Add("sav", 0)
        Aptitudes.Add("som", 0)
        Aptitudes.Add("wil", 0)

        Stats.Add("tt", 0)
        Stats.Add("luc", 0)
        Stats.Add("ir", 0)
        Stats.Add("wt", 0)
        Stats.Add("dur", 0)
        Stats.Add("dr", 0)
        Stats.Add("init", 0)
        Stats.Add("spd", 0)
        Stats.Add("db", 0)

    End Sub
End Class
