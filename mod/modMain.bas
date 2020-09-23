Attribute VB_Name = "modMain"
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function GetShortPathName Lib "kernel32.dll" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Private Declare Function GlobalGetAtomName Lib "kernel32" Alias "GlobalGetAtomNameA" (ByVal nAtom As Integer, ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public lzEngine_File As String, lzScript_File As String

Public clsTextBox As New txtClass
Public Const dlg_filter As String = "MyBasic Script(*.myb)|*.myb|Text Files(*.txt)|*.txt|"
Public Str_Error As String 'Error holder

Public Function GetAtom(AtomIdx As Integer) As String
Dim iBuff As String * 256
Dim iRet As Long
    iRet = GlobalGetAtomName(AtomIdx, iBuff, Len(iBuff))
    GetAtom = Left(iBuff, iRet)
    iBuff = "": iRet = 0
    
End Function

Function GetShPath(lpLongPath As String) As String
Dim iRet As Long
Dim sBuff As String * 256
    iRet = GetShortPathName(lpLongPath, sBuff, Len(sBuff))
    
    GetShPath = Left(sBuff, iRet)
    sBuff = ""
    
End Function

Function RunFile(lpFile As String, nHwnd As Long, lParm As String)
    ShellExecute nHwnd, "open", lpFile, lParm, vbNullString, 1
End Function

Function FixPath(lzPath As String) As String
   If Right(lzPath, 1) = "\" Then FixPath = lzPath Else FixPath = lzPath & "\"
End Function

Public Function IsFileHere(lzFileName As String) As Boolean
    If Dir(lzFileName) = "" Then IsFileHere = False: Exit Function Else IsFileHere = True
End Function

Function FindDir(lzPath As String) As Boolean
    If Not Dir(lzPath, vbDirectory) = "." Then
        FindDir = False
        Exit Function
    Else
        FindDir = True
    End If
End Function

Function SaveFile(lzFile As String, FileData As String)
Dim iFile As Long
    iFile = FreeFile
    Open lzFile For Output As #iFile
        Print #iFile, FileData;
    Close #iFile
    
    lzFile = ""
    FileData = ""
End Function


Public Function OpenFile(Filename As String) As String
Dim iFile As Long
Dim mByte() As Byte

    iFile = FreeFile
    Open Filename For Binary As #iFile
        ReDim Preserve mByte(0 To LOF(iFile))
        Get #iFile, , mByte
    Close #iFile
    
    OpenFile = StrConv(mByte, vbUnicode)
    
    Erase mByte
    
End Function

Function GetAbsPath(lpPath As String) As String
Dim x As Integer, e_pos As Integer
    For x = 1 To Len(lpPath)
        If Mid(lpPath, x, 1) = "\" Then e_pos = x
    Next
    
    If e_pos <> 0 Then GetAbsPath = Mid(lpPath, 1, e_pos)

End Function
