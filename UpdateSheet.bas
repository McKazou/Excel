Attribute VB_Name = "UpdateSheet"
Option Compare Text 'Ingore lower and upper case : https://stackoverflow.com/questions/17035660/in-vba-get-rid-of-the-case-sensitivity-when-comparing-words
'Beginning of the macro
Sub UpdateSheet()
'
' UpdateSheet
' And say "input validated" or "Invalidated"

    'Catching error by ignoring them
    On Error Resume Next
    
    'Stoppping screen update to gain time
    Application.ScreenUpdating = False
    
    'Update the workbook = Calculate now in the formulas tab
    Calculate
    
    'Reactivate screen update
    Application.ScreenUpdating = True
    
    'Create current date
    Dim nowDate As Date
    nowDate = Day(Now) & "/" & Month(Now) & "/" & Year(Now)
    
    'Getting current user https://forum.hardware.fr/hfr/Programmation/VB-VBA-VBS/recuperer-username-windows-sujet_119930_1.htm
    user = Environ$("Username")
    
    'List of messages for the button
    Dim tabMessage(4) As String
    tabMessage(0) = "INPUTS NOT VALIDATED" 'Default value for INPUTS buttons
    tabMessage(1) = "Inputs validated on " & nowDate & " by " & user
    tabMessage(2) = "OUTPUTS NOT VALIDATED" 'Default value for OUTPUTS buttons
    tabMessage(3) = "Outputs validated on " & nowDate & " by " & user
    
    'Get the button caller (where we clicked) https://www.mrexcel.com/forum/excel-questions/96466-how-get-caption-button-just-clicked-vba.html
    With ActiveSheet.Buttons(Application.Caller)
    
    'This part is looking for the name of the button
    'If button is *Input* then look for this current text and compare it to the list
    'https://www.exceltrainingvideos.com/tag/using-application-caller-with-form-button/
    'features LIKE : https://stackoverflow.com/questions/15585058/check-if-a-string-contains-another-string
    If (.Name Like "*input*") Then
        'If the message is the one by default then change it to the second
        If (.Caption = tabMessage(0)) Then
            .Caption = tabMessage(1)
        Else
            'if not put the default text
            .Caption = tabMessage(0)
        End If
    Else
    If (.Name Like "*output*") Then
        'If the message is the one by default then change it to the second
        If (.Caption = tabMessage(2)) Then
            .Caption = tabMessage(3)
        Else
            'if not put the default text
            .Caption = tabMessage(2)
        End If
    End If
    End If
    
    End With
    
    'ActiveSheet.Buttons(Application.Caller).Caption = buttonText
    

    MsgBox ("Update Done") 'Message box supplementaire in case the message isn't clear
    
    'Go to margin chek
    'Application.Goto (ActiveWorkbook.Sheets("RU").Range("A60")), scroll:=True
    
    
    
    
    
    
End Sub


