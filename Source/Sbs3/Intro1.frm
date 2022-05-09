VERSION 5.00
Begin VB.Form Intro1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   1935
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7695
   LinkTopic       =   "Form1"
   ScaleHeight     =   1935
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   2400
      Top             =   840
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1920
      Top             =   840
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Presents"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   1440
      Width           =   7695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "-Davidsoft-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7695
   End
End
Attribute VB_Name = "Intro1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public GamePath

Public I


Private Sub Form_Load()
GamePath = "E:\David\Files\Programming\Sbs3"

Label1.ForeColor = 0
Label2.ForeColor = 0
I = 0
End Sub

Private Sub Timer1_Timer()
If I = 0 Then
If Label1.ForeColor < 250 Then
Label1.ForeColor = Label1.ForeColor + 25
Label2.ForeColor = Label1.ForeColor
End If
End If
If Label1.ForeColor = 250 Then
I = I + 1
End If
If I = 25 Then
If Label1.ForeColor > 0 Then
Label1.ForeColor = Label1.ForeColor - 25
Label2.ForeColor = Label1.ForeColor
End If
If Label1.ForeColor = 0 Then
Timer1.Enabled = False
Timer2.Enabled = True
End If
End If
End Sub
Private Sub Timer2_Timer()
Timer2.Enabled = False
Intro1.Visible = False
Intro2.Show
End Sub
