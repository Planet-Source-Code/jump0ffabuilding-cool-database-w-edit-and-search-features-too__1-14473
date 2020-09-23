VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search for stupid voculary words"
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5235
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3750
   ScaleWidth      =   5235
   Begin VB.ComboBox CbCategory 
      Height          =   315
      ItemData        =   "Form3.frx":030A
      Left            =   502
      List            =   "Form3.frx":0314
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   345
      Width           =   2085
   End
   Begin VB.ComboBox CbOperator 
      Height          =   315
      ItemData        =   "Form3.frx":032A
      Left            =   2647
      List            =   "Form3.frx":033D
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   345
      Width           =   2085
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   577
      TabIndex        =   5
      Text            =   "String to search"
      Top             =   15
      Width           =   4095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Search"
      Height          =   240
      Left            =   517
      TabIndex        =   4
      Top             =   690
      Width           =   4185
   End
   Begin VB.TextBox txtWord 
      DataField       =   "Word"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   45
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1290
      Width           =   5145
   End
   Begin VB.TextBox txtDef 
      DataField       =   "Definition"
      DataSource      =   "Data1"
      Height          =   1320
      Left            =   15
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1935
      Width           =   5205
   End
   Begin VB.Data Data1 
      Caption         =   "Voculary is a waste of time"
      Connect         =   "Access"
      DatabaseName    =   "Voculary.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   465
      Left            =   37
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Voculary"
      Top             =   3300
      Width           =   5160
   End
   Begin VB.Label Label1 
      Caption         =   "Word:"
      Height          =   255
      Left            =   2385
      TabIndex        =   3
      Top             =   1065
      Width           =   450
   End
   Begin VB.Label Label2 
      Caption         =   "Definition:"
      Height          =   225
      Left            =   2265
      TabIndex        =   2
      Top             =   1725
      Width           =   690
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo die
Dim Teeth As String
Teeth = CbCategory.Text & " " & CbOperator.Text & " '" & Text1.Text & "'"
Data1.RecordSource = "SELECT * FROM Voculary WHERE " & Teeth
Data1.Refresh
Data1.Recordset.MoveLast: Data1.Recordset.MoveFirst

MsgBox Data1.Recordset.RecordCount & " matches found."

Exit Sub
die:
MsgBox "No results y'all shjick"
End Sub

Private Sub Form_Load()
CbOperator.ListIndex = 0
CbCategory.ListIndex = 0
End Sub

