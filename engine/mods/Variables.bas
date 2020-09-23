Attribute VB_Name = "Variables"
'OK this is the main mod were functions and Subs are kept/
' that are used for dealing with the variables

'Variable datatypes
Enum VarType
    NoKnownErr = 0
    nString = 1
    nInteger = 2
    nVar = 3
End Enum

'Variable stack
Private Type VarStack
    VariableName As String ' The Variables name
    VariableType As VarType ' Type of variable
    VarData As Variant ' Variables data
    isGlobal As Boolean 'Not used in this version
    CanChange As Boolean
End Type

Public MaxVars As Long ' Holds the current number of all variables
Public mVarStack() As VarStack

Public Function VariableIndex(lpVarName As String) As Integer
Dim idx As Integer, x As Integer
    
    'Locate an variables index based on it's name
    ' If lpVarName does nopt match mVarStack(x).VariableName then
    ' the default error index is returned -1
    idx = -1
    If MaxVars = -1 Then VariableIndex = idx: Exit Function
    
    For x = 0 To UBound(mVarStack) 'Lopp tho the variable stack
        If LCase(lpVarName) = mVarStack(x).VariableName Then
            'We have a match so store the variables index
            idx = x '--< store this
            Exit For 'Get out of this loop
        End If
    Next
    
    VariableIndex = idx 'Return good result index
    
End Function

Public Sub AddVariable(lVarName As String, lVarType As VarType, Optional isGlobalEx As Boolean = True, _
Optional isReadOnly As Boolean = False, Optional lpVarData As Variant)
    
    MaxVars = MaxVars + 1 'Keep a count of the total variables we have
    ReDim Preserve mVarStack(MaxVars) 'Resize the variable stack
    'Now we fill in the information for the current variable been added
    mVarStack(MaxVars).VariableName = LCase(lVarName)
    mVarStack(MaxVars).CanChange = isReadOnly ' Means if it can be changed by the user
    mVarStack(MaxVars).isGlobal = isGlobalEx ' Means can this variable be accessed outside of it's scope
    mVarStack(MaxVars).VarData = SetVarDataType(lVarType, lpVarData)  ' Get and set the variables data
    mVarStack(MaxVars).VariableType = lVarType ' Get the varibales datatype
End Sub

Public Sub SetVariableData(lVarName As String, Optional VarData As Variant)
Dim idx As Integer
    ' this function is used to set the variables data
    idx = VariableIndex(lVarName)
    mVarStack(idx).VarData = VarData
End Sub

Public Function GetVar(lpName As String) As Variant
Dim idx As Integer
    ' Function that is used to return the data from a Given variable
    ' if idx is returned -1 then a nullstring is sent back.
    idx = VariableIndex(lpName)
    
    If idx <> -1 Then
        GetVar = mVarStack(idx).VarData
        ' Return the variables data
        Exit Function
    End If
    
End Function

Function GetVarType(lpName As String) As VarType
    Dim idx As Integer
    'Returns a variables data type
    idx = VariableIndex(lpName)
    If idx <> -1 Then
        GetVarType = mVarStack(idx).VariableType
        Exit Function
    Else
        GetVarType = NoKnownErr
    End If
End Function

Function SetVarDataType(lpVarType As VarType, lpVarData As Variant) As Variant
On Error GoTo SetDataErr:
    ' This function is used for seting the proper datatypes with there data
    ' it also a good way to test for error such as overflows or incorrect datatypes
    Select Case lpVarType
        Case nInteger: SetVarDataType = CInt(lpVarData): Exit Function
        Case nString: SetVarDataType = CStr(lpVarData): Exit Function
        Case nVar: SetVarDataType = CVar(lpVarData): Exit Function
        Case Else: Abort 3, CurrentLine, GetVarDataTypeFromInt(lpVarType)
    End Select
    
    Exit Function
SetDataErr:
    Abort 2, CurrentLine, Err.Description
    
End Function

Function SetVarDefault(nType As VarType) As Variant
    'Set the default data of a new variable
    Select Case nType
        Case nInteger: SetVarDefault = 0
        Case nString: SetVarDefault = ""
        Case nVar: SetVarDefault = ""
    End Select
    
End Function

Function GetVarTypeFromStr(lpVarType As String) As VarType
    'This function is used to return the numric value of a variables datatype
    ' the function works by checking the string vartype and returning the value
    ' also see ENUM VarType
    
    Select Case UCase(lpVarType)
        Case "STRING": GetVarTypeFromStr = nString ' String variable
        Case "INTEGER": GetVarTypeFromStr = nInteger ' Numberic variable
        Case "VARIANT": GetVarTypeFromStr = nVar 'Variant datatype
        Case Else: GetVarTypeFromStr = NoKnownErr ' Unkown tpe or not supported yet
    End Select
    
End Function

Function GetVarDataTypeFromInt(ipVarType As VarType) As String
    ' Function is used to return the string name of the variables data type.
    Select Case UCase(lpVarType)
        Case nString: GetVarDataTypeFromInt = "STRING" ' String variable
        Case nInteger: GetVarDataTypeFromInt = "INTEGER" ' Numberic variable"
    End Select
End Function

