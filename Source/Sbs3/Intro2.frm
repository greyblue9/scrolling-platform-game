VERSION 5.00
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Begin VB.Form Intro2 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   ClientHeight    =   7695
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9615
   LinkTopic       =   "Form1"
   ScaleHeight     =   7695
   ScaleWidth      =   9615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "EXIT"
      Height          =   375
      Left            =   4080
      TabIndex        =   0
      Top             =   5040
      Width           =   1335
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   0
      Top             =   0
   End
   Begin PicClip.PictureClip PictureClip1 
      Left            =   0
      Top             =   49995
      _ExtentX        =   16960
      _ExtentY        =   13573
      _Version        =   393216
      Picture         =   "Intro2.frx":0000
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "Intro2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public I
Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long

Private Sub Command1_Click()
Result = PlaySound(GamePath & "\Sounds\Screen Filling In.wav", 1, 1)
Timer1.Enabled = True
End Sub

Private Sub Form_Load()
Dim Result As Long
End Sub
Private Sub Timer1_Timer()
Screen.ActiveForm.Width = Screen.ActiveForm.Width - 480
Screen.ActiveForm.Left = Screen.ActiveForm.Left + 240
Screen.ActiveForm.Height = Screen.ActiveForm.Height - 480
Screen.ActiveForm.Top = Screen.ActiveForm.Top + 240
If Intro2.Height = 15 Then
Timer1.Enabled = False
Timer2.Enabled = True
Screen.ActiveForm.Visible = False
End If
End Sub
Private Sub Timer2_Timer()
Timer2.Enabled = False
Unload Me
End
End Sub
