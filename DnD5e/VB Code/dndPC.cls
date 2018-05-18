VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "dndPC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Public Player As String
Public Name As String
Public Race As String
Public Background As String
Public Alignment As String
Public History As String
Public Speed As Long
Public ProficiencyBonus As Long
Public Languages As Collection

Private pcTools As Collection
Private HP As Dictionary
Private pcWpns As Dictionary
Private pcTraits As Dictionary
Private pcClasses As Dictionary
Private attributes As Dictionary
Private pcSkills As Dictionary

Friend Property Get Attrib(attribName As String) As Long
    If Len(attribName) > 3 Then
        attribName = Left(attribName, 3)
    End If
    Attrib = attributes(attribName)("score")
End Property
Friend Property Get AttribMod(attribName As String) As Long
    If Len(attribName) > 3 Then
        attribName = Left(attribName, 3)
    End If
    
    AttribMod = attributes(attribName)("mod")
End Property
Friend Property Get AttribSave(attribName As String) As Long
    If Len(attribName) > 3 Then
        attribName = Left(attribName, 3)
    End If
    
    If attributes(attribName)("isprof") = True Then
        AttribSave = attributes(attribName)("mod") + ProficiencyBonus
        Exit Property
    Else
        AttribSave = attributes(attribName)("mod")
        Exit Property
    End If
    Err.Raise CustomError.Err1, "dndPC", "Unable to get save value for " & attribName
End Property

Friend Sub addAttrib(attribName As String, result As Long)
If Len(attribName) > 3 Then
    attribName = Left(attribName, 3)
End If
Debug.Print attribName; ":"; result
Dim tempDic As New Dictionary
    tempDic.Add "Score", result
    tempDic.Add "isProf", False
    tempDic.Add "mod", _
        Application.WorksheetFunction.RoundUp( _
            result / 2 - 5, 0)
    attributes.Add attribName, tempDic
    Set tempDic = Nothing
End Sub

Private Sub Class_Initialize()
    Set attributes = New Dictionary
    Set Languages = New Collection
    Set pcTools = New Collection
    Set HP = New Dictionary
    Set pcWpns = New Dictionary
    Set pcTraits = New Dictionary
    Set pcClasses = New Dictionary
    Set pcSkills = New Dictionary
End Sub
