VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Begin VB.Form Form4 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6495
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      TabIndex        =   1
      Top             =   50000
      Width           =   6015
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   0
      Top             =   50000
   End
   Begin PicClip.PictureClip PictureClip1 
      Left            =   0
      Top             =   4320
      _ExtentX        =   6773
      _ExtentY        =   5927
      _Version        =   393216
      Picture         =   "Form4.frx":0000
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   3360
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      Max             =   80
      Scrolling       =   1
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "LEVEL EDITOR"
      BeginProperty Font 
         Name            =   "Ventilate"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   975
      Left            =   -120
      TabIndex        =   6
      Top             =   2040
      Width           =   6615
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "LEVEL EDITOR"
      BeginProperty Font 
         Name            =   "Ventilate"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   975
      Left            =   360
      TabIndex        =   5
      Top             =   2040
      Width           =   6495
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Super BJ Sisters 3"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   495
      Left            =   0
      TabIndex        =   4
      Top             =   1800
      Width           =   6495
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Loading map screen tiles..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   255
      Left            =   2160
      TabIndex        =   3
      Top             =   3000
      Width           =   3015
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0/81"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   495
      Left            =   5040
      TabIndex        =   2
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   3375
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6495
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public I
Public X
Private Sub Form_Load()
File1.Path = "C:\David\SBS3\Images\Levels"
I = -1
X = 0
ProgressBar1.Max = File1.ListCount - 1
End Sub

Private Sub Timer1_Timer()
For j = 0 To 3
I = I + 1
Form2.Image1(I).Picture = LoadPicture("C:\David\SBS3\Images\Levels\" & File1.List(I))
ProgressBar1.Value = I
Label5.Caption = I + 1 & "/" & File1.ListCount
If I = File1.ListCount - 1 Then
Timer1.Enabled = False
Form4.Visible = False
Form1.Show
Exit Sub
End If
Next j
X = X + 1
PictureClip1.ClipX = X * 2
PictureClip1.ClipY = X * 8
PictureClip1.ClipWidth = 224
PictureClip1.ClipHeight = 112
Image1.Picture = PictureClip1.Clip
End Sub

