VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form4 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   11775
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   11775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   2640
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   873
      _Version        =   393216
      Appearance      =   0
      Max             =   80
      Orientation     =   1
      Scrolling       =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   0
      Top             =   50000
   End
   Begin VB.FileListBox File1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3330
      Left            =   0
      Pattern         =   "*.gif"
      TabIndex        =   0
      Top             =   50000
      Width           =   6015
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/81"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   2280
      Width           =   11775
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CURRENTLY LOADING MAP SCREEN TILES"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   1920
      Width           =   11775
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "super BJ sisters 3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1335
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   11775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "MAP EDITOR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   975
      Left            =   0
      TabIndex        =   2
      Top             =   1320
      Width           =   11775
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public I
Public Z
Private Sub Form_Load()
File1.Path = "C:\David\Files\Programming\SBS3\Images\Map\"
I = -1
Z = 1
ProgressBar1.Max = File1.ListCount - 1
End Sub
Private Sub Timer1_Timer()
For j = 0 To 3
I = I + 1
Form2.Image1(I).Picture = LoadPicture("C:\David\Files\Programming\SBS3\Images\Map\" & File1.List(I))
ProgressBar1.Value = I
Label5.Caption = I + 1 & "/" & File1.ListCount
If I = File1.ListCount - 1 Then
Timer1.Enabled = False
Form4.Visible = False
Form1.Show
Exit Sub
End If
Next j
End Sub

