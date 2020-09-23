VERSION 5.00
Object = "{60CC5D62-2D08-11D0-BDBE-00AA00575603}#1.0#0"; "SYSGARBAGE.OCX"
Begin VB.MDIForm MainWindow 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "Voculary is a waste of time"
   ClientHeight    =   5325
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7200
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin SysTrayCtl.cSysTray cSysTray1 
      Left            =   750
      Top             =   2490
      _ExtentX        =   900
      _ExtentY        =   900
      InTray          =   -1  'True
      TrayIcon        =   "MDIForm1.frx":030A
      TrayTip         =   "Voculary Crip"
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      Begin VB.Menu mnuShowMain 
         Caption         =   "Show main window"
      End
      Begin VB.Menu mnuBored 
         Caption         =   "Click if you are bored"
      End
   End
   Begin VB.Menu mnuSyscrip 
      Caption         =   "I am cool"
      Visible         =   0   'False
      Begin VB.Menu mnuShowfjdisf 
         Caption         =   "Show main window"
      End
      Begin VB.Menu mnuExithaioijh 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "MainWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cSysTray1_MouseDblClick(Button As Integer, Id As Long)
Me.Show
End Sub

Private Sub cSysTray1_MouseDown(Button As Integer, Id As Long)
If vbRightButton Then PopupMenu mnuSyscrip
End Sub

Private Sub MDIForm_Load()
Form1.Show

End Sub

Private Sub mnuBored_Click()
Form4.Show
End Sub

Private Sub mnuExit_Click()
End
End Sub

Private Sub mnuExithaioijh_Click()
End
End Sub

Private Sub mnuShowfjdisf_Click()
Me.Show
End Sub

Private Sub mnuShowMain_Click()
Form1.Show
End Sub
