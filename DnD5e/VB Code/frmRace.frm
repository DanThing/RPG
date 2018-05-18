VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRace 
   Caption         =   "Select Race"
   ClientHeight    =   6015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6735
   OleObjectBlob   =   "frmRace.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmRace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCancel_Click()

End Sub

Private Sub btnOk_Click()

End Sub

Private Sub ComboBoxRace_Change()
Dim tempObj As Object, tempstring As String, itm As Long, itm2 As Variant
If Not Me.ComboBoxRace.Text = "---------" Then
    Set tempObj = ParseJson(getData("\Races\" & Me.ComboBoxRace.Text))

    descText = tempObj("desc")
    traitsText = "Size: " & tempObj("size") & vbNewLine
    traitsText = traitsText & "Speed: " & tempObj("speed") & vbNewLine
    traitsText = traitsText & "Ability Scores: " & tempObj("ability") & vbNewLine
    
    For itm = 1 To tempObj("trait").Count
        traitsText = traitsText & tempObj("trait")(itm)("name") & vbNewLine
        For Each itm2 In tempObj("trait")(itm)("text")
            traitsText = traitsText & Space(4) & itm2 & vbNewLine
        Next itm2
    Next itm
End If

End Sub

Private Sub UserForm_Initialize()
With Me.ComboBoxRace
         .AddItem "Human"
         .AddItem "Dragonborn"
         .AddItem "Hill Dwarf"
         .AddItem "Mountain Dwarf"
         .AddItem "Eladrin"
         .AddItem "High Elf"
         .AddItem "Wood Elf"
         .AddItem "Half-Elf"
         .AddItem "Forest Gnome"
         .AddItem "Rock Gnome"
         .AddItem "Halfling"
         .AddItem "Half-Orc"
         .AddItem "Tiefling"
         .AddItem "---------"
         .AddItem "Aarakocra"
         .AddItem "Aasimar"
         .AddItem "Duergar"
         .AddItem "Svirfneblin"
         .AddItem "Drow"
         .AddItem "Goliath"
         .AddItem "Air Genasi"
         .AddItem "Earth Genasi"
         .AddItem "Fire Genasi"
         .AddItem "Water Genasi"
    End With
End Sub



