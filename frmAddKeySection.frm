VERSION 5.00
Begin VB.Form frmAddKeySection 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Add Key or Section"
   ClientHeight    =   2760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4695
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   4695
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3360
      TabIndex        =   12
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2160
      TabIndex        =   11
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Frame fraSection 
      Height          =   1455
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   4455
      Begin VB.TextBox txtSection 
         Height          =   285
         Left            =   720
         TabIndex        =   10
         Top             =   600
         Width           =   3615
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Name :"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   615
         Width           =   510
      End
   End
   Begin VB.Frame fraKey 
      Height          =   1455
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   4455
      Begin VB.TextBox txtKeyValue 
         Height          =   285
         Left            =   720
         TabIndex        =   7
         Top             =   930
         Width           =   3615
      End
      Begin VB.TextBox txtKeyName 
         Height          =   285
         Left            =   720
         TabIndex        =   6
         Top             =   570
         Width           =   3615
      End
      Begin VB.Label lblCurSection 
         BorderStyle     =   1  'Fest Einfach
         Height          =   255
         Left            =   1440
         TabIndex        =   14
         Top             =   240
         Width           =   2865
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Current Section :"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1185
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Value :"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Name :"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   510
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      Begin VB.OptionButton optKey 
         Caption         =   "Add Key"
         Height          =   255
         Left            =   2400
         TabIndex        =   2
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton optSection 
         Caption         =   "Add Section"
         Height          =   255
         Left            =   960
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmAddKeySection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
If optSection Then
    If Not SectionOK Then
        MsgBox "Section-Name not entered !", vbCritical, "Error"
        Exit Sub
    End If
    If SectionExists(txtSection.Text) Then
        MsgBox "Section already exists !", vbCritical, "Error"
        Exit Sub
    End If
    Open INIFile For Append As #1
    Print #1, "[" & txtSection.Text & "]"
    Close #1
ElseIf optKey Then
    If Not KeysOK Then
        MsgBox "Key-Name or Value not entered !", vbCritical, "Error"
        Exit Sub
    End If
    If KeyExists(txtKeyName.Text) Then
        MsgBox "Key already exists !", vbCritical, "Error"
        Exit Sub
    End If
    INISetValue INIFile, lblCurSection.Caption, txtKeyName.Text, txtKeyValue.Text
End If
Unload Me
End Sub

Private Sub optKey_Click()
fraSection.Visible = False
fraKey.Visible = True
End Sub

Private Sub optSection_Click()
fraSection.Visible = True
fraKey.Visible = False
End Sub

Private Function KeysOK() As Boolean
KeysOK = True
If txtKeyName.Text = "" Or txtKeyValue.Text = "" Then KeysOK = False
End Function

Private Function SectionOK() As Boolean
SectionOK = True
If txtSection.Text = "" Then SectionOK = False
End Function

Private Function SectionExists(ByVal Section As String) As Boolean
SectionExists = False
For i% = 1 To frmMain.trvINI.Nodes.Count
    If frmMain.trvINI.Nodes(i%).Text = Section Then
        SectionExists = True
        Exit For
    End If
Next i%
End Function

Private Function KeyExists(ByVal Key As String) As Boolean
KeyExists = False
For i% = 1 To frmMain.lvwINI.ListItems.Count
    If frmMain.lvwINI.ListItems(i%).Text = Key Then
        KeyExists = True
        Exit For
    End If
Next i%
End Function
