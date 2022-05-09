VERSION 5.00
Begin VB.Form Black 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   2655
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3015
   LinkTopic       =   "Form1"
   MouseIcon       =   "Black.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   2655
   ScaleWidth      =   3015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   600
      Top             =   1080
   End
End
Attribute VB_Name = "Black"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
Unload Me
End
End Sub
Private Sub Timer1_Timer()
Timer1.Enabled = False
Intro1.Show
End Sub
