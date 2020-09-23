VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Voculary is stupid"
   ClientHeight    =   870
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3015
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   870
   ScaleWidth      =   3015
   Begin VB.CommandButton Command1 
      Caption         =   "Click to close"
      Height          =   750
      Left            =   105
      TabIndex        =   0
      Top             =   75
      Width           =   2805
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim pot As Form4
Set pot = New Form4
pot.Refresh
pot.Show

End Sub
