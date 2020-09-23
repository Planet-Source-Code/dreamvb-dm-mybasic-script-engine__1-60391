Attribute VB_Name = "modUtils"
'Any tools we use for the scripting engine will be placed in here.
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function GlobalAddAtom Lib "kernel32" Alias "GlobalAddAtomA" (ByVal lpString As String) As Integer

Public Sub GetLineCount(lpScript As String)
    ' All the sub does is store the number of lines the script has into LineCount
    LineCount = UBound(Split(lpScript, vbCrLf)) - 1
End Sub

Function GetIde() As Long
    'We use this to get the hangle of the ide window
    GetIde = FindWindow(vbNullString, "DM MyBasic-Script")
End Function

Function CharPos(lpStr As String, nChr As String) As Integer
Dim x As Integer, idx As Integer
    idx = 0
    'Function used to find the position of nChr in lpStr
    'Ex CharPos("hello world","r") returns 9
    
    For x = 1 To Len(lpStr) 'Loop tho lpStr
        If Mid(lpStr, x, 1) = nChr Then 'check if we have a match
            idx = x ' yes we have so store it's index
            Exit For ' get out of this loop
        End If
    Next
    
    x = 0
    CharPos = idx ' Return the index
    
End Function

Function isEmptyLine(expLine As String) As Boolean
    'Checks if the current executeing line is a nullchar
    isEmptyLine = (expLine = vbNullChar)
End Function

Function FixPath(lzPath As String) As String
    'Appends a \ to a given path
   If Right(lzPath, 1) = "\" Then FixPath = lzPath Else FixPath = lzPath & "\"
End Function

Public Function IsFileHere(lzFileName As String) As Boolean
    'Used to checking if a file exsits
    If Dir(lzFileName) = "" Then IsFileHere = False: Exit Function Else IsFileHere = True
End Function

Public Function OpenFile(Filename As String) As String
Dim iFile As Long
Dim mByte() As Byte 'Byte array to hold the contents of the file

    'Opens a given file
    iFile = FreeFile 'Pointer to a free file
    Open Filename For Binary As #iFile 'Open file in binary mode
        'Resize the array to hold the data based on the length of the file
        If LOF(iFile) = 0 Then
            ReDim Preserve mByte(0 To LOF(iFile))
        Else
            ReDim Preserve mByte(0 To LOF(iFile) - 1)
        End If
        Get #iFile, , mByte 'Stote the data into the byte array
    Close #iFile
    
    OpenFile = StrConv(mByte, vbUnicode) 'Convert the array to a VB string and return
    
    Erase mByte 'Erase the array conents
    
End Function

