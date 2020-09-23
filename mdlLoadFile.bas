Attribute VB_Name = "mdlLoadFile"
Public Sub LoadINIFile(ByVal FileName As String)
Dim Result&, Buffer$, X&, y&, pos&, i%
INIFile = FileName
frmMain.List1.Clear
X = 256
Do
    Buffer = String$(X, Chr$(0))
    Result = GetPrivateProfileSectionNames(Buffer, X, INIFile)
    X = X + 256
Loop While (Result + 2 = X - 256)
If Result Then
    pos = 1
    y = 1
    Do While y < Result
        pos = InStr(y, Buffer, Chr$(0))
        frmMain.List1.AddItem Mid$(Buffer, y, pos - y)
        y = pos + 1
    Loop
End If
frmMain.trvINI.Nodes.Clear
Set nodX = frmMain.trvINI.Nodes.Add(, , "Root", FileName, 1)
For i% = 0 To frmMain.List1.ListCount - 1
    Set nodX = frmMain.trvINI.Nodes.Add("Root", tvwChild, frmMain.List1.List(i%), frmMain.List1.List(i%), 2)
Next i%
End Sub
