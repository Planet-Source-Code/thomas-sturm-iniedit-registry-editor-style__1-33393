VERSION 5.00
Begin VB.Form frmEdit 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Edit Key/Value"
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox txtValue 
      Height          =   285
      Left            =   720
      TabIndex        =   2
      Top             =   840
      Width           =   3855
   End
   Begin VB.Label lblKey 
      BorderStyle     =   1  'Fest Einfach
      Height          =   255
      Left            =   720
      TabIndex        =   7
      Top             =   480
      Width           =   3855
   End
   Begin VB.Label lblCurSection 
      BorderStyle     =   1  'Fest Einfach
      Height          =   255
      Left            =   1440
      TabIndex        =   6
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Current Section :"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   1185
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   4560
      Y1              =   1335
      Y2              =   1335
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   4560
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Value :"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Key :"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   360
   End
End
Attribute VB_Name = "frmEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
If txtValue.Text = "" Then
    MsgBox "No Value entered !", vbCritical, "Error"
    Exit Sub
End If
INISetValue INIFile, lblCurSection.Caption, lblKey.Caption, txtValue.Text
Unload Me
End Sub
