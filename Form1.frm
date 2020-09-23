VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Voculary is stupid"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5205
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3420
   ScaleWidth      =   5205
   Begin VB.CommandButton Command4 
      Caption         =   "Search"
      Height          =   405
      Left            =   3862
      TabIndex        =   7
      Top             =   2265
      Width           =   945
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Edit"
      Height          =   405
      Left            =   2707
      TabIndex        =   6
      Top             =   2265
      Width           =   945
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Delete"
      Height          =   405
      Left            =   1552
      TabIndex        =   5
      Top             =   2265
      Width           =   945
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      Height          =   405
      Left            =   397
      TabIndex        =   4
      Top             =   2265
      Width           =   945
   End
   Begin VB.TextBox txtDef 
      DataField       =   "Definition"
      DataSource      =   "Data1"
      Height          =   1320
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   930
      Width           =   5205
   End
   Begin VB.TextBox txtWord 
      DataField       =   "Word"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   30
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   285
      Width           =   5145
   End
   Begin VB.Data Data1 
      Caption         =   "I hate vocabulary. I am too education for it."
      Connect         =   "Access"
      DatabaseName    =   "Voculary.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   750
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Voculary"
      Top             =   2670
      Width           =   5205
   End
   Begin VB.Label Label2 
      Caption         =   "Definition:"
      Height          =   225
      Left            =   2257
      TabIndex        =   3
      Top             =   720
      Width           =   690
   End
   Begin VB.Label Label1 
      Caption         =   "Word:"
      Height          =   255
      Left            =   2377
      TabIndex        =   2
      Top             =   60
      Width           =   450
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Data1.Recordset.AddNew
Form2.Show
Form2.txtWord.SetFocus
End Sub

Private Sub Command2_Click()
pot = MsgBox("Are you sure you want to delete this record?? You won't be able to get it back after you delete it.", vbYesNo, "I am cool")
If pot = vbYes Then
Data1.Recordset.Delete
Data1.Recordset.MoveNext
End If
End Sub

Private Sub Command3_Click()
Form2.Show
Form2.txtWord.SetFocus
End Sub

Private Sub Command4_Click()
Form3.Show
End Sub


