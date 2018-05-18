Attribute VB_Name = "modErrHandler"

Option Explicit

'******************************************************
' Workbook Name: Komatsu Report Base Tools & Data.xlsm
'   Module: modErrHandler
' Code by Daniel Boyce
' Version: 1.2.20180430
'
'Sub Example()
'    'Example of how to use the custom handler
'    Application.ScreenUpdating = False
'    On Error GoTo ErrorExit
'
'    Dim x As Long
'    x = 1
'    Dim y As Long
'    y = 2
'
'    If x = y Then Err.Raise CustomError.CustomError1'
'
'EOM:
'    Application.ScreenUpdating = True
'    Exit Sub
'
'ErrorExit:
'    ErrHandler Err
'    Resume EOM
'End Sub


Public Enum CustomError
    Err1 = vbObjectError + 2000
    Err2 = vbObjectError + 3000
    Err3 = vbObjectError + 4000
    Err4 = vbObjectError + 5000
    Err5 = vbObjectError + 6000
    Err6 = vbObjectError + 7000
    Err7 = vbObjectError + 8000
    Err8 = vbObjectError + 9000
    Err9 = vbObjectError + 10000
    Err10 = vbObjectError + 11000
End Enum

Public Sub errHandler(Err As Object)
    Select Case Err.Number
        Case CustomError.Err1
            MsgBox "Unable to get value.", vbExclamation
        
        Case CustomError.Err2
            MsgBox "Value already exists.", vbExclamation
        
        Case CustomError.Err3
            MsgBox "Custom Error Message 3", vbExclamation
        
        Case CustomError.Err4
            MsgBox "Custom Error Message 4", vbExclamation
        
        Case CustomError.Err5
            MsgBox "Custom Error Message 5", vbExclamation
            
        Case CustomError.Err6
            MsgBox "Custom Error Message 6", vbExclamation
        
        Case CustomError.Err7
            MsgBox "Custom Error Message 7", vbExclamation
        
        Case CustomError.Err8
            MsgBox "Custom Error Message 8", vbExclamation
        
        Case CustomError.Err9
            MsgBox "Error rolling on table.", vbExclamation
            
        Case CustomError.Err10
            MsgBox "Custom Error Message 10", vbExclamation
            
        Case Else
            MsgBox "Unexpected Error: " & Err.Number & "- " & Err.Description, vbCritical
    End Select
End Sub


