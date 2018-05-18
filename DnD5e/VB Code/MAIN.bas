Attribute VB_Name = "MAIN"
Option Explicit
Option Compare Text

Enum SourceBook
    AdvCOS ': "Curse of Strahd",
    DMG ': "Dungeon Master's Guide",
    AdvEE ': "Elemental Evil",
    AdvHDQ ': "Tyranny of Dragons: Hoard of the Dragon Queen",
    AdvOOA ': "Out of the Abyss",
    PHB ': "Player Handbook",
    AdvPOA ': "Princes of the Apocalypse",
    AdvROT ': "Tyranny of Dragons: The Rise of Tiamat",
    SCG ': "Sword Cost Adventurer's Guide",
    AdvSKT ': "Storm King's Thunder",
    UA ': "Unearthed Arcana",
    VGM ': "Volo's Guide to Monsters",
    XGE ': "Xanathar's Guide to Everything"
End Enum

Public DefPath As String

Public PC As dndPC

Public jObj As Object

Public EnableLogger As Boolean

Public progBar As ProgressBar

Sub MAIN()

    EnableLogger = True
    
    DefPath = ThisWorkbook.Path & Application.PathSeparator
    
    Set PC = New dndPC
    
    'getAttributes
  
    getRace

    jObj = Nothing

End Sub

Private Sub getRace()
Dim thisForm As New frmRace
    thisForm.Show
End Sub

Private Sub getAttributes()
Dim thisForm As New frmAttributes
    thisForm.Show
End Sub
