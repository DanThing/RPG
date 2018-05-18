Attribute VB_Name = "modLogger"
Option Explicit

Public Sub Log(loggerString As String)
If EnableLogger = True Then
    Debug.Print loggerString
End If
End Sub
