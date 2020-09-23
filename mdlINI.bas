Attribute VB_Name = "mdlINI"
Option Explicit

Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileSectionNames Lib "kernel32" Alias "GetPrivateProfileSectionNamesA" (ByVal lpszReturnBuffer As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Public lvitem As ListItem
Public nodX As Node
Public INIFile As String
       
Public Sub INISetValue(ByVal Path$, ByVal Sect$, ByVal Key$, ByVal Value$)
Dim Result&
Result = WritePrivateProfileString(Sect, Key, Value, Path)
End Sub

Public Function INIGetValue(ByVal Path$, ByVal Sect$, ByVal Key$) As String
Dim Result&, Buffer$
Buffer = Space$(32)
Result = GetPrivateProfileString(Sect, Key, vbNullString, Buffer, Len(Buffer), Path)
INIGetValue = Left$(Buffer, Result)
End Function

Public Function INISetArray(ByVal Path$, ByVal Sect$, xArray() As String)
Dim X%, Buffer$, Result&
For X = LBound(xArray) To UBound(xArray)
    Buffer = Buffer & xArray(X) & Chr$(0)
Next X
Buffer = Left$(Buffer, Len(Buffer) - 1)
Result = WritePrivateProfileSection(Sect, Buffer, Path)
End Function

Public Sub INIGetArray(ByVal Path$, ByVal Sect$, xArray() As String)
Dim Result&, Buffer$
Dim l%, p%, z%
Buffer = Space(32767)
Result = GetPrivateProfileSection(Sect, Buffer, Len(Buffer), Path)
Buffer = Left$(Buffer, Result)
If Buffer <> "" Then
    l = 1
    ReDim xArray(0)
    Do While l < Result
        p = InStr(l, Buffer, Chr$(0))
        If p = 0 Then Exit Do
        xArray(z) = Mid$(Buffer, l, p - l)
        z = z + 1
        ReDim Preserve xArray(0 To z)
        l = p + 1
    Loop
End If
End Sub

Public Sub INIDeleteKey(ByVal Path$, ByVal Sect$, ByVal Key$)
Call WritePrivateProfileString(Sect, Key, 0&, Path)
End Sub
 
Public Sub INIDeleteSection(ByVal Path$, ByVal Sect$)
Call WritePrivateProfileString(Sect, 0&, 0&, Path)
End Sub

