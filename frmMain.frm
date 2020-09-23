VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "INIEdit (c) 2002 Thomas Sturm - No File loaded"
   ClientHeight    =   8070
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   12465
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8070
   ScaleWidth      =   12465
   StartUpPosition =   2  'Bildschirmmitte
   Begin MSComctlLib.StatusBar stbStatus 
      Align           =   2  'Unten ausrichten
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   7695
      Width           =   12465
      _ExtentX        =   21987
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   21934
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cdlgINI 
      Left            =   11400
      Top             =   6840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgTool 
      Left            =   5000
      Top             =   2760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":089C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1576
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2250
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2F2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3244
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3F1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4238
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbBar 
      Align           =   1  'Oben ausrichten
      Height          =   810
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   12465
      _ExtentX        =   21987
      _ExtentY        =   1429
      ButtonWidth     =   2196
      ButtonHeight    =   1376
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "imgTool"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Open"
            Key             =   "Open"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "&Add Section"
            Key             =   "AddSection"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "&Delete Section"
            Key             =   "DelSection"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Add &Key"
            Key             =   "AddKey"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "D&elete Key"
            Key             =   "DelKey"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Exit"
            Key             =   "Exit"
            ImageIndex      =   4
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.ImageList imgTrv 
      Left            =   12600
      Top             =   2160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4F12
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":59DC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView trvINI 
      Height          =   6855
      Left            =   0
      TabIndex        =   3
      Top             =   855
      Width           =   5355
      _ExtentX        =   9446
      _ExtentY        =   12091
      _Version        =   393217
      HideSelection   =   0   'False
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "imgTrv"
      Appearance      =   1
   End
   Begin MSComctlLib.ListView lvwINI 
      Height          =   6885
      Left            =   5400
      TabIndex        =   2
      Top             =   840
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   12144
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Key"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Value"
         Object.Width           =   8819
      EndProperty
   End
   Begin VB.ListBox List2 
      Height          =   2790
      Left            =   5640
      TabIndex        =   1
      Top             =   2520
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.ListBox List1 
      Height          =   2010
      Left            =   8760
      TabIndex        =   0
      Top             =   8640
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu X1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuSections 
      Caption         =   "&Sections"
      Enabled         =   0   'False
      Begin VB.Menu mnuAddSection 
         Caption         =   "&Add New Section"
      End
      Begin VB.Menu mnuDelSection 
         Caption         =   "&Delete selected Section"
      End
      Begin VB.Menu X2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCountSections 
         Caption         =   "&Count Sections in INI File"
      End
   End
   Begin VB.Menu mnuValues 
      Caption         =   "&Values"
      Enabled         =   0   'False
      Begin VB.Menu mnuAddValue 
         Caption         =   "&Add New Key"
      End
      Begin VB.Menu mnuDelKey 
         Caption         =   "&Delete Selected Key"
      End
      Begin VB.Menu X3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCountKeys 
         Caption         =   "&Count Keys in Section"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&?"
      Begin VB.Menu mnuAboutProg 
         Caption         =   "&About INIEdit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Set nodX = trvINI.Nodes.Add(, , "Root", "Please open an INI-File", 1)
stbStatus.Panels(1).Text = "No File loaded."
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim retVal As Integer
retVal = MsgBox("Do you want to exit ?", vbQuestion + vbYesNo, "Exit Program ?")
If retVal = vbNo Then
    Cancel = -1
End If
End Sub

Private Sub lvwINI_DblClick()
Load frmEdit
frmEdit.lblCurSection.Caption = trvINI.SelectedItem.Text
frmEdit.lblKey.Caption = lvwINI.SelectedItem.Text
frmEdit.Show vbModal, Me
End Sub

Private Sub mnuAboutProg_Click()
frmAbout.Show vbModal, Me
End Sub

Private Sub mnuCountKeys_Click()
Dim KeyCount As Long
KeyCount = List2.ListCount
MsgBox "There are " & KeyCount & " Key(s) in this Section.", vbInformation, "Keys counted."
End Sub

Private Sub mnuCountSections_Click()
Dim SectionCount As Long
SectionCount = trvINI.Nodes.Count - 1
MsgBox "There are " & SectionCount & " Sections in this INI-File.", vbInformation, "Sections counted."
End Sub

Private Sub mnuDelKey_Click()
INIDeleteKey INIFile, trvINI.SelectedItem.Key, lvwINI.SelectedItem.Text
lvwINI.ListItems.Remove (lvwINI.SelectedItem.Index)
End Sub

Private Sub mnuDelSection_Click()
INIDeleteSection INIFile, trvINI.SelectedItem.Text
trvINI.Nodes.Remove (trvINI.SelectedItem.Index)
lvwINI.ListItems.Clear
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub mnuOpen_Click()
With cdlgINI
    .Filter = "INI Files (*.ini)|*.ini|Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
    .FilterIndex = 1
    .DialogTitle = "Select INI File to load"
    .ShowOpen
End With
If cdlgINI.FileName <> "" Then
    LoadINIFile cdlgINI.FileName
    Me.Caption = "INIEdit (c) 2002 Thomas Sturm - " & INIFile
    mnuSections.Enabled = True
    stbStatus.Panels(1).Text = INIFile
    trvINI.Nodes(1).Expanded = True
End If
End Sub

Private Sub tlbBar_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
Case "Open"
    With cdlgINI
        .Filter = "INI Files (*.ini)|*.ini|Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
        .FilterIndex = 1
        .DialogTitle = "Select INI File to load"
        .ShowOpen
    End With
    If cdlgINI.FileName <> "" Then
        LoadINIFile cdlgINI.FileName
        Me.Caption = "INIEdit (c) 2002 Thomas Sturm - " & INIFile
        mnuSections.Enabled = True
        stbStatus.Panels(1).Text = INIFile
        trvINI.Nodes(1).Expanded = True
    End If
Case "AddSection"
    Load frmAddKeySection
    frmAddKeySection.optSection.Value = True
    frmAddKeySection.Show vbModal, Me
    LoadINIFile INIFile
    frmMain.trvINI.Nodes(frmMain.trvINI.Nodes.Count).Selected = True
    LoadINIArray trvINI.SelectedItem.Key
Case "DelSection"
    INIDeleteSection INIFile, trvINI.SelectedItem.Text
    trvINI.Nodes.Remove (trvINI.SelectedItem.Index)
    lvwINI.ListItems.Clear
Case "AddKey"
    Load frmAddKeySection
    frmAddKeySection.optKey.Value = True
    frmAddKeySection.lblCurSection.Caption = trvINI.SelectedItem.Text
    frmAddKeySection.Show vbModal, Me
    LoadINIArray trvINI.SelectedItem.Key
Case "DelKey"
    INIDeleteKey INIFile, trvINI.SelectedItem.Key, lvwINI.SelectedItem.Text
    lvwINI.ListItems.Remove (lvwINI.SelectedItem.Index)
Case "Exit"
    Unload Me
End Select
End Sub

Private Sub trvINI_NodeClick(ByVal Node As MSComctlLib.Node)
LoadINIArray Node.Key
End Sub

Private Sub LoadINIArray(ByVal CurNode As String)
Dim X%, xArray$()
Dim Key$, Value$, KeyLine() As String
ReDim xArray(0)
Call INIGetArray(INIFile, CurNode, xArray)
List2.Clear
lvwINI.ListItems.Clear
If CurNode = "Root" Then
    tlbBar.Buttons(3).Enabled = False
    tlbBar.Buttons(4).Enabled = False
    tlbBar.Buttons(6).Enabled = False
    tlbBar.Buttons(7).Enabled = False
    mnuSections.Enabled = False
    mnuValues.Enabled = False
    Exit Sub
Else
    tlbBar.Buttons(3).Enabled = True
    tlbBar.Buttons(4).Enabled = True
    tlbBar.Buttons(6).Enabled = True
    tlbBar.Buttons(7).Enabled = True
    mnuSections.Enabled = True
    mnuValues.Enabled = True
End If
If UBound(xArray) > 0 Then
    For X = 0 To UBound(xArray) - 1
        List2.AddItem xArray(X)
    Next X
Else
    Exit Sub
End If
If List2.ListCount = 0 Then
    tlbBar.Buttons(7).Enabled = False
    mnuValues.Enabled = False
End If
Dim i%
For i% = 0 To List2.ListCount - 1
    ReDim KeyLine(0) As String
    KeyLine() = Split(List2.List(i%), "=", 2)
    If UBound(KeyLine) > 0 Then
        Key = KeyLine(0)
        Value = KeyLine(1)
    Else
        Key = KeyLine(0)
        Value = ""
    End If
    Set lvitem = lvwINI.ListItems.Add(i% + 1, "Key" & i% + 1, Key)
    lvitem.SubItems(1) = Value
Next i%
stbStatus.Panels(1).Text = INIFile & "\" & CurNode
End Sub
