Attribute VB_Name = "modInput"
Sub DoDim(lpStr As String)
Dim StrA As String, e_pos As Integer, n_pos As Integer, StrVarName As String, nVarType As VarType

    If isEmptyLine(lpStr) Then Abort 4, CurrentLine
    
    e_pos = CharPos(lpStr, Chr(32))
    If e_pos = 0 Then
        StrVarName = Trim(lpStr) 'Get variable name
        
        'check if the variable is not already in the stack
        If VariableIndex(StrVarName) <> -1 Then
            'Variable was already found
            e_pos = 0
            Abort 5, CurrentLine, StrVarName
        Else
            'Add the variable to the variables stack
            AddVariable StrVarName, nVar, , , ""
            e_pos = 0
            Exit Sub
        End If
    Else
        'Get the variable name
        StrVarName = Trim(Mid(lpStr, 1, e_pos - 1))
        
        n_pos = InStr(e_pos + 1, lpStr, Chr(32), vbBinaryCompare) 'Check for for next space
        If n_pos = 0 Then
            Abort 2, CurrentLine, "Required DataType"
            e_pos = 0: n_pos = 0: StrVarName = ""
            Exit Sub
        End If
        
        'Make sure that we have AS in the expression
        StrA = UCase(Trim(Mid(lpStr, e_pos + 1, n_pos - e_pos - 1)))
        If StrA <> "AS" Then
            StrA = "": e_pos = 0: StrVarName = ""
            Abort 2, CurrentLine, "Required AS"
        End If
        
        StrA = Trim(Mid(lpStr, n_pos + 1, Len(lpStr))) 'Extract the variables datatype
        nVarType = GetVarTypeFromStr(StrA) 'Store the varibales datatype
        
        If nVarType = NoKnownErr Then Abort 7, CurrentLine, StrA 'Invauld datatype was found
        StrA = ""
        
        'ok so we have our variable and the variable data type we can now add
        'them to the variable stack.But first we need to check if there alread here
         If VariableIndex(StrVarName) <> -1 Then
            e_pos = 0: n_pos = 0
            Abort 5, CurrentLine, StrVarName
        Else 'We can no Safely add the new variable
            AddVariable StrVarName, nVarType, False, , SetVarDefault(nVarType)
            e_pos = 0: n_pos = 0: StrA = "": StrVarName = ""
        End If
    End If
    
End Sub

Sub DoAssign1(lpExpr As String)
Dim e_pos As Integer, StrVarName As String, AssignData As Variant
Dim iTemp As Variant

    If isEmptyLine(lpExpr) Then Abort 8, CurrentLine, "LET", " = Expression"
    
    'Check for the assign pos
    e_pos = CharPos(lpExpr, "=") 'Get location of the assignment sign
    If e_pos = 0 Then Abort 2, CurrentLine, "Required '='" 'check for the assignment sign
    
    StrVarName = Trim(Mid(lpExpr, 1, e_pos - 1)) 'Extract the variable name
    
    'check that the variable name above is in the variable stack
    If VariableIndex(StrVarName) = -1 Then
        StrVarName = ""
        Abort 6, CurrentLine, StrVarName
    Else
        AssignData = Trim(Mid(lpExpr, e_pos + 1, Len(lpExpr))) 'Extract the expression
        
        If AssignData = "" Then
            'No expression was found
            StrVarName = "": AssignData = ""
            Abort 2, CurrentLine, "Required expression"
        Else
            iTemp = Eval(AssignData) 'eval the assign data
            SetVariableData StrVarName, SetVarDataType(GetVarType(StrVarName), iTemp)
        End If
    End If
End Sub

Sub DoInput()
Dim lpVarName As String, Str_Tmp As String
    e_pos = CharPos(ProcessLine, ",") 'Find the position of the parm seprator ,
    
    If isEmptyLine(ProcessLine) Or e_pos = 0 Then
        'No expression was found so we abort
        Abort 8, CurrentLine, "LOCATE", "Expression,Expression"
        Exit Sub
    End If
    
    'Extract the variable name
    lpVarName = Trim(Mid(ProcessLine, e_pos + 1, Len(ProcessLine)))
    'Check that the variable is in the variables stack
    If VariableIndex(lpVarName) = -1 Then
        'Variable was not found so we abort
        Abort 5, CurrentLine, lpVarName
        lpVarName = ""
        Exit Sub
    Else
        Str_Tmp = Eval(Mid(ProcessLine, 1, e_pos - 1)) 'Extract the propmt message
        'Now we need to print the propmting message to the user
        cWriteLine Str_Tmp
        Str_Tmp = ""
        'Now we will use the console read command to get input form the user
        Str_Tmp = cReadConsole()
        'Stote the user input data into the variable ->lpVarName
        SetVariableData lpVarName, Str_Tmp
        'Clean up used varaibles
        Str_Tmp = ""
        lpVarName = ""
        e_pos = 0
    End If
    
End Sub
