Attribute VB_Name = "modEval"
Function Token(Expression, pos) As Variant
Dim s_FuncName As String, e_pos As Long
Dim Value As Variant
    
    inSideQuotes = False
    Dim ch As String
    Dim pl As Integer, es As Integer
    
    Do
        ch = Mid(Expression, pos, 1)
            
        If isOperator(ch) Then
            Exit Do
        ElseIf ch = "(" Then
            pos = pos + 1
            pl = 1
            e_pos = pos
            'Look for the brackets in the expression
            Do
                ch = Mid(Expression, pos, 1)
                If ch = "(" Then pl = pl + 1
                If ch = ")" Then pl = pl - 1
                pos = pos + 1
            Loop Until pl = 0 Or pos > Len(Expression)
                
            Value = Mid(Expression, e_pos, pos - e_pos - 1) 'Get value of the function name
            s_FuncName = LCase(Trim(Token)) 'Get the function name
                
            If s_FuncName = "" Then
                'if no function name was found we just return the token
                Token = Eval(Value)
            End If
            
            'Preform built in functions
            Select Case s_FuncName
                Case "chr": Token = Chr(Eval(Value))
                Case "asc": Token = Asc(Eval(Value))
                Case "str": Token = Str(Eval(Value))
            End Select
            
            ElseIf ch = Chr(34) Then
                    inSideQuotes = True
                    'This allows us to know if we are in a sting with quotes
                    pl = 1
                    pos = pos + 1
                    Do
                        ch = Mid(Expression, pos, 1)
                        pos = pos + 1
                        If ch = Chr(34) Then
                            If Mid(Expression, pos, 1) = Chr(34) Then
                                Value = Value & Chr(34)
                                pos = pos + 1
                            Else
                                Exit Do
                            End If
                        Else
                            Value = Value & ch
                        End If
                    Loop Until pl = 0 Or pos > Len(Expression)
                    Token = Value
            Else
                'Keep on building the token
                Token = Token & ch
                pos = pos + 1
            End If

   Loop Until pos > Len(Expression) 'Loop until we reach the end of the expression
   
   Token = ReturnData(CStr(Token))
   
    If IsNumeric(Token) Then
        Token = Val(Token)
    Else
        Token = CStr(Token)
    End If
    
End Function

Function isOperator(StrExp As String) As Boolean
    isOp = False
    If StrExp = "+" Or StrExp = "-" Or StrExp = "*" Or StrExp = "\" Or StrExp = "/" _
    Or StrExp = "&" Or StrExp = "^" Or StrExp = "=" Or StrExp = "<" Or StrExp = ">" _
    Or StrExp = "%" Then isOperator = True
End Function

Public Function Eval(Expression As Variant)
Dim iCounter As Integer, sOperator As String, Value As Variant
Dim iTmp As Variant, ch As String
On Error Resume Next

    iCounter = 1
    
    Do While iCounter <= Len(Expression)
        ch = Mid(Expression, iCounter, 1)
        
        If isOperator(ch) Then
            sOperator = ch
            iCounter = iCounter + 1
        End If

        Select Case sOperator
            Case ""
                Value = Token(Expression, iCounter)
            Case "+"
                Value = Value + Token(Expression, iCounter)
            Case "-"
                Value = Value - Token(Expression, iCounter)
            Case "*"
                Value = Value * Token(Expression, iCounter)
            Case "/"
                Value = Value / Token(Expression, iCounter)
            Case "&"
                Value = Value & Token(Expression, iCounter)
                Value = Trim(Value)
            Case "="
                iTmp = Token(Expression, iCounter)
                Value = Value = iTmp
                Value = Abs(Value)
            Case Else
                Value = ""
        End Select
    Loop
    
    ch = ""
    
    If IsNumeric(Value) Then
        Eval = Val(Value)
    Else
        Eval = CStr(Value)
    End If
    
    Eval = Value
    
    Value = ""
    
    If Err Then
        Abort 2, CurrentLine, Err.Description
    End If
    
    
End Function


