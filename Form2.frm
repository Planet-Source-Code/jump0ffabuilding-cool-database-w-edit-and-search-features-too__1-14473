VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Your Voculary crap"
   ClientHeight    =   2685
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5205
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2685
   ScaleWidth      =   5205
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   465
      Left            =   2850
      TabIndex        =   3
      Top             =   2220
      Width           =   1965
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   465
      Left            =   390
      TabIndex        =   2
      Top             =   2220
      Width           =   1965
   End
   Begin VB.TextBox txtWord 
      DataField       =   "Word"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   30
      TabIndex        =   0
      Top             =   225
      Width           =   5145
   End
   Begin VB.TextBox txtDef 
      DataField       =   "Definition"
      DataSource      =   "Data1"
      Height          =   1320
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   870
      Width           =   5205
   End
   Begin VB.Label Label1 
      Caption         =   "Word:"
      Height          =   255
      Left            =   2377
      TabIndex        =   5
      Top             =   0
      Width           =   450
   End
   Begin VB.Label Label2 
      Caption         =   "Definition:"
      Height          =   225
      Left            =   2257
      TabIndex        =   4
      Top             =   660
      Width           =   690
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.txtWord = txtWord
Form1.txtDef = txtDef
Me.Hide
End Sub

Private Sub Command2_Click()
Me.Hide
End Sub

Private Sub Form_Activate()
txtWord = Form1.txtWord
txtDef = Form1.txtDef
End Sub

