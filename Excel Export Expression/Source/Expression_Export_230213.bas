Attribute VB_Name = "Expression_Export"
'From German Macro
'Macro to export a file into C:/Temp/ that will be read in NX automatically
'It will go into the "Transfer" sheet
'Read all value in it
'Create a file with failed value and good value in two differents files
'Last changes :
'15/02/2019 : First use => Operationnal
'30/04/2019 : Found some ignore values for no reason; working on it; solv, "Filter in other sheet must be set properlly to avoid error"
'29/05/2019 : Date on the progress bar & User MsgBox disable for dev
Public DEFAULT_FOLDER As String 'This is the default Folder to put the export file in
Public DEFAULT_ERROR_LOG_FILE_NAME As String  'Default Error log file
Public DEFAULT_FILE_NAME As String

Sub Expression_Export_Dev()
        Call Init
        Call Expression_Export(False, False)
End Sub


Private Sub Expression_Export(Display_MsgBox As Boolean, OUTPUT_TO_FILE As Boolean)

'On Error GoTo ErrorHandler

    Dim FileName As Range 'Cell where to find the NX file name to export in
    Dim LastLine As Integer 'Number of values to check in the sheet "Transfer"
    Dim OutputText() As String 'Array with all string to put in the exp file for exporting
    Dim OutputError() As String 'Array with variables still in Error
    Dim Target As Range 'This is a temporary value during the loop
    Dim CountError As Integer 'Count how many error their is
    
    
    Dim nowDate As Date
    nowDate = Format(Now, "mm/dd/yyyy hh:mm")
    
'----------------------Creating the list of parameters to run---------------------
    With Sheets("Transfer") 'Forced to read in the Transfer sheet
        'Get the Nx file name on wich to put variables; this a safe guard to make sure variables are set in the intented file
        Set FileName = .Range("ControlFileName")
    
        Application.StatusBar = "Parameters verification"
    
        LastLine = .Cells(ActiveSheet.Rows.Count, "A").End(xlUp).Row
        
        ReDim OutputError(0) 'Initialisation for the output Error
        ReDim OutputText(0) 'Initialisation for the output text
        CountError = 0
        
        For i = 1 To LastLine
            Set Target = .Range("A" & i)
            'Target.Select
            ParametersName = Target.Value 'Name of the parameter
            NumRow = Target.Offset(0, 4).Value 'Row at wich we can find it in the relative sheet
            ParameterValue = Target.Offset(0, 1).Value

            '-----------------------------Error detected in the Name of the parameters---------------------- This will cause NX to stop import
            If IsError(ParametersName) Then
                If NumRow <> CVErr(xlErrNA) And NumRow <> CVErr(xlErrRef) Then
                    'This is looking for the row num to verify is this is bug, if it is we can ignore the line
                    Target.Select 'Developper only
                    Err.Raise 10000, "NameIsInError"
                Else
                     'We ignore value like #REF and #NA if the row line has error otherwise the name itself is bug
                End If
            '-----------------------------The value it self is in Error so we add it to the list---------------------- This will prevent errors puts into the 3D and break things
            'Rajouter un filtre pour les , sur les PCs qui ne sont pas correctement paramétré
            ElseIf IsError(ParameterValue) Then 'The value it self is in Error so we add it to the list
                If IsNumeric(NumRow) Then
                    'This is looking for the row num to verify is this is bug, if it is we can ignore the line
                    'Target.Select 'Developper only
                    OutputError(CountError) = "VALUE IN ERROR : " & ParametersName 'Adding the value to the list
                    CountError = CountError + 1 'Counting one more error
                    ReDim Preserve OutputError(CountError)  'adding size to the list to make sure we have more available place than error values
                    Application.StatusBar = "VALUE IN ERROR :" & ParametersName 'Message to show what happen to the user
                Else
                     'We ignore value like #REF and #NA if the row line has error otherwise the name itself is bug
                End If
            '-----------------------------Looking for "," because NX dont like it--------------------------------
            ElseIf InStr(ParameterValue, ",") > 0 Then
            'Their is a comma in the value, this value will not work in NX
                'Target.Select 'Developper only
                OutputError(CountError) = "COMMA : " & ParametersName & "=" & ParameterValue 'Adding the value to the list
                CountError = CountError + 1 'Counting one more error
                ReDim Preserve OutputError(CountError)  'adding size to the list to make sure we have more available place than error values
                Application.StatusBar = "COMMA : " & ParametersName & "=" & ParameterValue 'Message to show what happen to the user
            
            '-----------------------------Add it to the export list----------------------
            Else
                'No error has been detected in the name or the value
                'So we add it to the list to export
                OutputText(OutputLimit) = Target.Offset(0, 6).Value
                OutputLimit = UBound(OutputText)
                ReDim Preserve OutputText(OutputLimit + 1)
                Application.StatusBar = Target.Value
            End If
        Next
    End With
'----------------------Writing the Error file-------------------------------------
    'This function will create an output file with all values in error,
    'to make the process quicker, it has been removed from the macro
    If (OUTPUT_TO_FILE) Then
        Call FillFileWith(OutputError, DEFAULT_ERROR_LOG_FILE_NAME, Display_MsgBox)
        
        If UBound(OutputError) >= 3 Then
            If MsgBox("Their is more than 3 error in the selected list of parameters, you should review those value" & vbNewLine & _
                "An error log file has been created in " & DEFAULT_FOLDER & "Would you like to open it ?", _
                vbYesNo, "Multiple parameters in error has been detected") Then
                Call OpenFile("C:\WINDOWS\system32\notepad.exe", DEFAULT_FOLDER & "\" & DEFAULT_ERROR_LOG_FILE_NAME)
            End If
        Else
                If UBound(OutputError) + 1 > 1 Then
                For Each s In OutputError 'Telling the user wich value are in error
                    If Not s = "" Then MsgBox "Still in error: " & s
                Next
            End If
        End If
    End If
'----------------------Writing the file-------------------------------------
        
        Call FillFileWith(OutputText, DEFAULT_FILE_NAME, Display_MsgBox)
        
'----------------------Ending the macro------------------------------------
        Dim OutputMessage As String
                    
        If UBound(OutputError) + 1 > 1 And Not OutError = Empty Then
            OutputMessage = (UBound(OutError) + 1) & " defective values exported the " & nowDate
        Else
            OutputMessage = (UBound(OutputText) + 1) & " values exported the " & nowDate
        End If
        Application.StatusBar = OutputMessage
        If Display_MsgBox Then MsgBox "File generated sucessfully; as always"
Exit Sub
'ERROR HANDLER -------------------------------------------------------------
ErrorHandler:
Select Case Err.Number

        Case 1004
            GoTo MissingControlFileName
        Case 10000
            GoTo NameIsInError
        Case Else
            MsgBox "Error " & Err.Number & ": " & Err.Description & " at line " & Erl, vbCritical, "Not handle error"
            Stop
        End Select
Exit Sub

'ERROR HANDLER ---------------MissingControlFileName---------------------
MissingControlFileName:
    Application.StatusBar = "Cell with Control File name is missing"
    MsgBox "Cell with Control File name is missing, please verify the name of the NX part on where you want to export; this value is in the Excel."
    ApplicationUpdate (True)
    Err.Clear
    Exit Sub
    
'ERROR HANDLER ---------------NameIsInError-----------------------------
NameIsInError:
    Application.StatusBar = "A parameter's name is in Error, please call a developper to tchek in the list of Value to transfert into NX"
    MsgBox "A parameter's Name is in Error, please call a developper to tchek in the list of value to transfert into NX", _
    vbExclamation
    ApplicationUpdate (True)
    Err.Clear
    Resume Next
    
End Sub

'#Objectif: Open a file at a path provided
Private Sub OpenFile(ExecutablePath As String, FileToOpenPath As String)
    'https://stackoverflow.com/questions/15951837/wait-for-shell-command-to-complete
    Dim wsh As Object
    Set wsh = VBA.CreateObject("WScript.Shell")
    Dim waitOnReturn As Boolean: waitOnReturn = False

    If CheckForFile(ExecutablePath) Then
        If CheckForFile(FileToOpenPath) Then
            'Call Shell(ExecutablePath & " " & FileToOpenPath, vbNormalFocus) 'Open but wait user to close it
            wsh.Run ExecutablePath & " " & FileToOpenPath, 1, waitOnReturn
        Else
            MsgBox "Error logs file not found at : " & FileToOpenPath
        End If
    Else
        MsgBox "Executable not found at : " & ExecutablePath
    End If
End Sub
        
'#Objectif: This function will output the file using the current
Private Sub FillFileWith(ByRef OutputText() As String, ByVal FileName As String, Display_MsgBox As Boolean)

On Error GoTo ErrorHandler
        Dim FileFullPath As String
        FileFullPath = DEFAULT_FOLDER & "\" & FileName 'Concatenation of the defautl folder and the name
        
        Dim fso As Object 'Writer; later called as "Scripting.FileSystemObject"
        Dim oFile As Object 'File for the Wirter; later called as fs.CreateTextFile(path)
        
        If Not CheckForFolder(DEFAULT_FOLDER) Then 'Looking for the folder if not create it
            If Display_MsgBox Then MsgBox "Default folder not found creating it : " & DEFAULT_FOLDER
            With CreateObject("Scripting.FileSystemObject")
                If Not .FolderExists(DEFAULT_FOLDER) Then .CreateFolder DEFAULT_FOLDER
            End With
        End If
        If Not CheckForFile(FileFullPath) Then 'Looking for the file if not create it
        'https://stackoverflow.com/questions/11503174/how-to-create-and-write-to-a-txt-file-using-vba
            If Display_MsgBox Then MsgBox "File not found creating it: " & FileFullPath
            Set fso = CreateObject("Scripting.FileSystemObject")
            Set oFile = fso.CreateTextFile(FileFullPath)
            oFile.WriteLine ""
            oFile.Close
            Set fso = Nothing
            Set oFile = Nothing
        End If
        If UBound(OutputText) > 0 Then 'Their is text to put in
            Application.StatusBar = "Filling File"
            Set fso = CreateObject("Scripting.FileSystemObject")
            Set oFile = fso.CreateTextFile(FileFullPath)
            i = 1
            For Each s In OutputText 'Filling text
                Application.StatusBar = "Filling File:" & i & "/" & (UBound(OutputText) + 1)
                oFile.WriteLine s
                i = i + 1
            Next
            oFile.Close
            Set fso = Nothing
            Set oFile = Nothing
        End If
Exit Sub
'ERROR HANDLER -------------------------------------------------------------
ErrorHandler:
Select Case Err.Number
        Case Else
            MsgBox "Error " & Err.Number & ": " & Err.Description & " at line " & Err, vbCritical, "Not handle error"
            Stop
        End Select
Exit Sub
End Sub

'#Objectifs: Active or desactive calculation, screen update and events or activate them
Private Sub ApplicationUpdate(UpdateActive As Boolean)
        On Error GoTo UpdateErrorHandler
        With Application
            If UpdateActive = False Then
            .Calculation = xlCalculationManual
            Else
            .Calculation = xlCalculationSemiautomatic
            End If
            .ScreenUpdating = UpdateActive
            .EnableEvents = UpdateActive
        End With
    Exit Sub
    
UpdateErrorHandler:
    MsgBox "Somethings went wrong while trying to put the excel in Automatical calculation or manual calculation, please refer to a developper to help"
End Sub

'#Objectif: Create folder if isn't present
Private Sub FileUpdate(path As String)
        Dim FileFullPath As String
        FileFullPath = DEFAULT_FOLDER & "\" & FileName 'Concatenation of the defautl folder and the name
        
        With CreateObject("Scripting.FileSystemObject")
            If Not .FolderExists(DEFAULT_FOLDER) Then .CreateFolder DEFAULT_FOLDER
        End With
End Sub

'#Objectif: Test for the imported function
Private Sub TestIt() 'Fonction test for CheckForFolder and CheckForFile
Dim s As String
Dim b As Boolean
    s = "C:\TEMP": b = CheckForFolder(s): Debug.Print s & "|" & b
    s = "C:\TEMP\": b = CheckForFolder(s): Debug.Print s & "|" & b
    s = "C:\TEEMP": b = CheckForFolder(s): Debug.Print s & "|" & b
    s = "C:\Temp\TA-K-00H000_R0-1_SK_export.exp": b = CheckForFile(s): Debug.Print s & "|" & b
    s = "C:\TEMP\Parametric.exp": b = CheckForFile(s): Debug.Print s & "|" & b
End Sub
'------------------------------------------------------------------
Function CheckForFolder(ByVal strPathToFolder As String) As Boolean
'To use this fn you must set a reference for Scripting Runtime
'In this case we instance it using "Create Object" to avoid to have to get the Scripting Runtime object to import manually
'--------------------------------------------------
'1.  In the VBE window, Choose Tools | References
'2.   Check the box for Microsoft Scripting Runtime
'--------------------------------------------------
Dim fso As Object
Dim blnFolderExists As Boolean

    'Create object
    Set fso = CreateObject("Scripting.FileSystemObject")
    CheckForFolder = fso.FolderExists(strPathToFolder)
    If Not fso Is Nothing Then Set fso = Nothing

End Function
'------------------------------------------------------------------
Function CheckForFile(ByVal strPathToFile As String) As Boolean
'To use this fn you must set a reference for Scripting Runtime
'In this case we instance it using "Create Object" to avoid to have to get the Scripting Runtime object to import manually
'--------------------------------------------------
'1.  In the VBE window, Choose Tools | References
'2.   Check the box for Microsoft Scripting Runtime
'--------------------------------------------------
Dim fso As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    CheckForFile = fso.FileExists(strPathToFile)
    If Not fso Is Nothing Then Set fso = Nothing

End Function

'#Objectif: This function will be call by a button se on the "general data sheet" doing exactly the same as the macro but looking for the number of error
'In the name of the button we put the cell were to find the number of value in error,
'If their is more than one error, user can't use the macro
'Otherwise, run the macro to export
'This allow the user to not export errors and developpers to be able to export even with error
Sub UserExportButton()
    Call ApplicationUpdate(False)
    
    'Get the button caller (where we clicked) https://www.mrexcel.com/forum/excel-questions/96466-how-get-caption-button-just-clicked-vba.html
    With ActiveSheet.Buttons(Application.Caller)
        ButtonNameLength = Len(.Name) 'Length of the name of the button
        FoundSplitter = InStr(1, .Name, "_") 'Finding the splitter to get the cell where the nb or erros is
        NbErrosPosition = Mid(.Name, FoundSplitter + 1, ButtonNameLength) 'Getting the position in the name of the button
        
        NbErrors = ActiveSheet.Range(NbErrosPosition).Value
        
        Dim nowDate As Date
        nowDate = Format(Now, "mm/dd/yyyy hh:mm")

        If NbErrors > 0 Then
            Application.StatusBar = "Export failed because of " & NbErrors & " defective values the " & nowDate
            MsgBox "Their is atleast 1 errors to export, please contact a skilled users to solve it"
        Else
            Call Expression_Export_User
        End If
    
        Call ApplicationUpdate(True)
    End With
End Sub

Private Sub Expression_Export_User()
        Call Init
        Call Expression_Export(True, True)
End Sub


Sub Init()
    DEFAULT_FOLDER = "C:\Temp"
    DEFAULT_FILE_NAME = "Parametric.exp"
    DEFAULT_ERROR_LOG_FILE_NAME = "Parameters Error Logs.txt"
End Sub






