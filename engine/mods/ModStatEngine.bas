Attribute VB_Name = "ModStatEngine"
' This is the main mod for the start up events for the engine.

Public Sub RunCode(lpScript As String)
    GetLineCount lpScript 'Get the number of lines from the script
    
    If LineCount <= 0 Then
        'Abort if no lines were found in the script
        Abort 0, 0
        Exit Sub
    Else
        GetCodeLines lpScript 'Load in the code lines
        InitConsole 'Start the console
        cSetTitle "MyBasic-Script" 'Set the console title
        pngParser 'Start the executeing of the code
    End If
    
End Sub

Public Sub pngParser()
    'What this sub does is loop though all the code lines in LineHolder
    ' We then call TokenKeywords that returns keywords from the current
    ' executeing line eg PRINT,CLS,BEEP etc
    ' we then call Parser this then executes the current keywords
    
    Do While CurrentLine < LineCount 'Loop until we hit the max number of lines in the script
        CurrentLine = CurrentLine + 1 'Keep a count on our CurrentLine
        TokenKeywords 'call TokenKeywords
        Parser 'Call Parser
        DoEvents ' Allow other things to process
    Loop
    
End Sub

Public Sub Parser()
Dim thisToken As String
    thisToken = ""
    
    thisToken = TokenKeywords 'Get the current token or the current line been processed
    
    If Len(thisToken) <> 0 Then
        Select Case thisToken
            Case "REM"
                thisToken = "" 'Comments here we do nothing
            Case "DIM"
                DoDim ProcessLine
            Case "BEEP"
                cBeep 'Make the console Beep
            Case "CLS"
                cCls 'Clear the console
            Case "INPUT"
                Call DoInput
            Case "LET"
                DoAssign1 ProcessLine 'Assignment
            Case "LOCATE"
                Call Locate 'Locate used to position the text in the console
            Case "PRINT"
                Call PrintA ' Print Statement
            Case "PAUSE"
                Call cPause ' pause the console
            Case "END"
                'End Program and clean up
                cFree ' free the console
                'Reset ' reset
                thisToken = "" ' clear current token
            Case Else
                'an unkown keyword has been found
                Abort 1, CurrentLine, thisToken
        End Select
    End If
    
End Sub

Public Function TokenKeywords() As String
Dim x_pos As Integer, h_pos As Integer, sLine As String

    ' this function is used to process the current line and find any tokens
    ' the function works by looking for a white space in the current lines
    ' ex
    ' <KeyWord> |Space| <KeyWord data> ex PRINT "HELLO WORLD"
    
    sLine = Trim(LineHolder(CurrentLine)) 'Trim down the line
    x_pos = InStr(1, sLine, Chr$(32), vbBinaryCompare) 'Locate the space chr(32)
    
    If x_pos > 0 Then ' Yes we have we found a space
        TokenKeywords = UCase(Mid(sLine, 1, x_pos - 1)) 'Get and return the keyword
        ProcessLine = Mid(sLine, x_pos + 1, Len(sLine)) ' Get the keywords data
    Else ' OK no space was found so we must asume this is a keyword with no data eg BEEP,CLS etc
        If Len(sLine) <> 0 Then
            ProcessLine = vbNullChar 'Clear the process line as we have no data for this keyword
            TokenKeywords = UCase(sLine) ' Get and return the Keyword
        End If
    End If
    
    'Clean up used vars
    x_pos = 0
    sLine = ""
    
End Function

Public Sub AddCode(lpCodeScript As String)
    Reset 'Call Global Reset
    RunCode lpCodeScript 'Add the script code to be run
    cPause 'Put a pause on the console so it does not flash of
    cFree  'Clsoe the console we script has finsihed
    End ' close the engine
End Sub

Sub Main()
Dim sCode As String
Dim lzCommand As String
    lzCommand = Trim(Command$) 'Get the command lines argv
    
    If Len(lzCommand) = 0 Then
        'No command argv were found so we inform the user and exit
        MsgBox "Incorrect command line arguments" _
        & vbCrLf & "USE: " & App.EXEName & ".exe" & " ProgramName.myb", vbCritical, "Error"
        End
    Else
        If Left(lzCommand, 1) = Chr(34) And Right(lzCommand, 1) = Chr(34) Then
            'Fix the filename by removeing any quotes
            lzCommand = Right(lzCommand, Len(lzCommand) - 1): lzCommand = Left(lzCommand, Len(lzCommand) - 1)
        End If
        'Check the that the file does exsist
        If Not IsFileHere(lzCommand) Then
            'No file was found so we just exit
            MsgBox "File not found:" & vbCrLf & lzCommand, vbCritical, "Error"
            lzCommand = ""
            End
        Else
            'load in the script file
            sCode = OpenFile(lzCommand)
            'Now that we have the code we can now run the script
            AddCode sCode
            lzCommand = "" 'Clear the command line
        End If
    End If
End Sub
