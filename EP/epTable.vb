Public Class epTable

    Friend tableType As String

    Friend values As Array

    Friend contents As Array

    Friend nextAction As Array

    Friend cost As Array

    Friend Sub Add(ByVal tableObj As IDictionary)

        Dim itm As DictionaryEntry

        For Each itm In tableObj

            Select Case itm.Key
                Case "type"
                    tableType = tableObj(itm)
                Case "values"
                    values = tableObj(itm)
                Case "table"
                    contents = tableObj(itm)
                Case "action"
                    nextAction = tableObj(itm)
                Case "cost"
                    cost = tableObj(itm)
            End Select

        Next

    End Sub


End Class
