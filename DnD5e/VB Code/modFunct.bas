Attribute VB_Name = "modFunct"
Option Explicit

Function Roll(dSize As Long, Optional dQty As Long = 1) As Long
Dim rCount As Long
    Roll = 0
    rCount = 1
    Do
        Randomize
        Roll = Roll + Int((dSize - 1 + 1) * Rnd + 1)
        rCount = rCount + 1
    Loop While Not rCount > dQty
End Function

Function rollAttribute() As Long
Dim arr As Object
    Set arr = CreateObject("System.Collections.ArrayList")
    arr.Add Roll(6)
    arr.Add Roll(6)
    arr.Add Roll(6)
    arr.Add Roll(6)
    arr.Sort
    rollAttribute = arr(1) + arr(2) + arr(3)
End Function


Function getData(filename As String) As String
Dim fileNo As Integer
    fileNo = FreeFile
    If Not Right(filename, 5) = ".json" Then
        filename = filename & ".json"
    End If
    Open DefPath & "Data\" & filename For Input As #fileNo
    getData = Input(LOF(fileNo), fileNo)
    Close #fileNo
End Function


