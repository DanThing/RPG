VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAttributes 
   Caption         =   "Assign Attributes"
   ClientHeight    =   4710
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "frmAttributes.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAttributes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCancel_Click()
    Me.Hide
    End
End Sub

Private Sub btnOk_Click()
    assignRolls
End Sub

Private Sub btnReroll_Click()
    rollAll
End Sub

Private Sub assignRolls()
On Error GoTo errHandler

Select Case ComboBox1.Text
    Case "Strength"
        PC.addAttrib "str", Label1.Caption
    Case "Dexterity"
        PC.addAttrib "dex", Label1.Caption
    Case "Constitution"
        PC.addAttrib "con", Label1.Caption
    Case "Intelligence"
        PC.addAttrib "int", Label1.Caption
    Case "Wisdom"
        PC.addAttrib "wis", Label1.Caption
    Case "Charisma"
        PC.addAttrib "cha", Label1.Caption
    Case Else
        Err.Raise CustomError.Err2, "Assigning Ability Scores"
End Select

Select Case ComboBox2.Text
    Case "Strength"
        PC.addAttrib "str", Label2.Caption
    Case "Dexterity"
        PC.addAttrib "dex", Label2.Caption
    Case "Constitution"
        PC.addAttrib "con", Label2.Caption
    Case "Intelligence"
        PC.addAttrib "int", Label2.Caption
    Case "Wisdom"
        PC.addAttrib "wis", Label2.Caption
    Case "Charisma"
        PC.addAttrib "cha", Label2.Caption
    Case Else
        Err.Raise CustomError.Err2, "Assigning Ability Scores"
End Select

Select Case ComboBox3.Text
    Case "Strength"
        PC.addAttrib "str", Label3.Caption
    Case "Dexterity"
        PC.addAttrib "dex", Label3.Caption
    Case "Constitution"
        PC.addAttrib "con", Label3.Caption
    Case "Intelligence"
        PC.addAttrib "int", Label3.Caption
    Case "Wisdom"
        PC.addAttrib "wis", Label3.Caption
    Case "Charisma"
        PC.addAttrib "cha", Label3.Caption
    Case Else
        Err.Raise CustomError.Err2, "Assigning Ability Scores"
End Select

Select Case ComboBox4.Text
    Case "Strength"
        PC.addAttrib "str", Label4.Caption
    Case "Dexterity"
        PC.addAttrib "dex", Label4.Caption
    Case "Constitution"
        PC.addAttrib "con", Label4.Caption
    Case "Intelligence"
        PC.addAttrib "int", Label4.Caption
    Case "Wisdom"
        PC.addAttrib "wis", Label4.Caption
    Case "Charisma"
        PC.addAttrib "cha", Label4.Caption
    Case Else
        Err.Raise CustomError.Err2, "Assigning Ability Scores"
End Select

Select Case ComboBox5.Text
    Case "Strength"
        PC.addAttrib "str", Label5.Caption
    Case "Dexterity"
        PC.addAttrib "dex", Label5.Caption
    Case "Constitution"
        PC.addAttrib "con", Label5.Caption
    Case "Intelligence"
        PC.addAttrib "int", Label5.Caption
    Case "Wisdom"
        PC.addAttrib "wis", Label5.Caption
    Case "Charisma"
        PC.addAttrib "cha", Label5.Caption
    Case Else
        Err.Raise CustomError.Err2, "Assigning Ability Scores"
End Select

Select Case ComboBox6.Text
    Case "Strength"
        PC.addAttrib "str", Label6.Caption
    Case "Dexterity"
        PC.addAttrib "dex", Label6.Caption
    Case "Constitution"
        PC.addAttrib "con", Label6.Caption
    Case "Intelligence"
        PC.addAttrib "int", Label6.Caption
    Case "Wisdom"
        PC.addAttrib "wis", Label6.Caption
    Case "Charisma"
        PC.addAttrib "cha", Label6.Caption
    Case Else
        Err.Raise CustomError.Err2, "Assigning Ability Scores"
End Select

cleanExit:
    Unload Me
    Exit Sub

errHandler:
    errHandler Err
    Stop
End Sub

Private Sub rollAll()
    Label1.Caption = rollAttribute
    Label2.Caption = rollAttribute
    Label3.Caption = rollAttribute
    Label4.Caption = rollAttribute
    Label5.Caption = rollAttribute
    Label6.Caption = rollAttribute
End Sub

Private Sub UserForm_Initialize()

    With Me.ComboBox1
        .AddItem "Strength"
        .AddItem "Dexterity"
        .AddItem "Constitution"
        .AddItem "Intelligence"
        .AddItem "Wisdom"
        .AddItem "Charisma"
        .Text = .List(0)
    End With
    
    With Me.ComboBox2
        .AddItem "Strength"
        .AddItem "Dexterity"
        .AddItem "Constitution"
        .AddItem "Intelligence"
        .AddItem "Wisdom"
        .AddItem "Charisma"
        .Text = .List(1)
    End With
    
    With Me.ComboBox3
        .AddItem "Strength"
        .AddItem "Dexterity"
        .AddItem "Constitution"
        .AddItem "Intelligence"
        .AddItem "Wisdom"
        .AddItem "Charisma"
        .Text = .List(2)
    End With
    
    With Me.ComboBox4
        .AddItem "Strength"
        .AddItem "Dexterity"
        .AddItem "Constitution"
        .AddItem "Intelligence"
        .AddItem "Wisdom"
        .AddItem "Charisma"
        .Text = .List(3)
    End With
        
    With Me.ComboBox5
        .AddItem "Strength"
        .AddItem "Dexterity"
        .AddItem "Constitution"
        .AddItem "Intelligence"
        .AddItem "Wisdom"
        .AddItem "Charisma"
        .Text = .List(4)
    End With
    
    With Me.ComboBox6
        .AddItem "Strength"
        .AddItem "Dexterity"
        .AddItem "Constitution"
        .AddItem "Intelligence"
        .AddItem "Wisdom"
        .AddItem "Charisma"
        .Text = .List(5)
    End With
    
    rollAll
    
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Me.Hide
    End
End Sub
