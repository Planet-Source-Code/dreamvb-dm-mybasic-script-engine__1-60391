Attribute VB_Name = "modOutput"
Sub PrintA(Optional LineFreed As Boolean = False)
Dim StrA As String
    'Print to the console
    
    StrA = Eval(ProcessLine) 'Get the String to be places on the console
    If isEmptyLine(StrA) Then
        'No expression was found so we abort
        Abort 8, CurrentLine, "PRINT", " = Expression"
        Exit Sub
    End If
    
    cWriteLine StrA
    StrA = ""

End Sub

Sub Locate()
Dim e_pos As Integer, A As Integer, B As Integer
On Error GoTo LocateErr:

    e_pos = CharPos(ProcessLine, ",") 'Find the position of the parm
    
    If isEmptyLine(ProcessLine) Or e_pos = 0 Then
        'No expression was found so we abort
        Abort 8, CurrentLine, "LOCATE", "Expression,Expression"
        iTmp = ""
        Exit Sub
    End If
    
    'Get both parms for the function
    A = CInt(Eval(Mid(ProcessLine, 1, e_pos - 1)))
    B = CInt(Eval(Mid(ProcessLine, e_pos + 1, Len(ProcessLine))))
    cSetCursorPosition A, B 'Position the text on the console
    A = 0: B = 0: e_pos = 0 ' Clean up
    
    Exit Sub
LocateErr:
    A = 0: B = 0: e_pos = 0
    If Err Then Abort 2, CurrentLine, Err.Description & " LOCATE " & ProcessLine
End Sub
