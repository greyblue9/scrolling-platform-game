VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mandatory Map Editing Information"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7335
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Terminal"
      Size            =   6
      Charset         =   255
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   7335
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   2760
      TabIndex        =   5
      Top             =   4920
      Width           =   1815
   End
   Begin VB.Image Image18 
      Height          =   495
      Left            =   2280
      Picture         =   "Form3.frx":0000
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   495
   End
   Begin VB.Image Image17 
      Height          =   495
      Left            =   1800
      Picture         =   "Form3.frx":08E6
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   495
   End
   Begin VB.Image Image16 
      Height          =   495
      Left            =   1320
      Picture         =   "Form3.frx":11CC
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   495
   End
   Begin VB.Image Image15 
      Height          =   495
      Left            =   1320
      Picture         =   "Form3.frx":1AB2
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   495
   End
   Begin VB.Image Image14 
      Height          =   495
      Left            =   1320
      Picture         =   "Form3.frx":2398
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   495
   End
   Begin VB.Image Image13 
      Height          =   495
      Left            =   2280
      Picture         =   "Form3.frx":2C7E
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   495
   End
   Begin VB.Image Image12 
      Height          =   495
      Left            =   2760
      Picture         =   "Form3.frx":3564
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   495
   End
   Begin VB.Image Image11 
      Height          =   495
      Left            =   2280
      Picture         =   "Form3.frx":3E4A
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   495
   End
   Begin VB.Image Image10 
      Height          =   495
      Left            =   1800
      Picture         =   "Form3.frx":4730
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   495
   End
   Begin VB.Image Image9 
      Height          =   495
      Left            =   1800
      Picture         =   "Form3.frx":5016
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   495
   End
   Begin VB.Image Image8 
      Height          =   495
      Left            =   1800
      Picture         =   "Form3.frx":58FC
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label Label8 
      Caption         =   "*"
      Height          =   135
      Left            =   1080
      TabIndex        =   8
      Top             =   3000
      Width           =   135
   End
   Begin VB.Label Label7 
      Caption         =   "**"
      Height          =   135
      Left            =   3960
      TabIndex        =   7
      Top             =   2880
      Width           =   255
   End
   Begin VB.Label Label6 
      Caption         =   $"Form3.frx":61E2
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   720
      Width           =   7335
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "[ Rest of island ]"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4200
      TabIndex        =   4
      Top             =   3960
      Width           =   975
   End
   Begin VB.Line Line6 
      X1              =   6240
      X2              =   5760
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line Line5 
      X1              =   5760
      X2              =   5760
      Y1              =   4320
      Y2              =   3840
   End
   Begin VB.Line Line4 
      X1              =   5760
      X2              =   5280
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Line Line3 
      X1              =   5280
      X2              =   5280
      Y1              =   3840
      Y2              =   3360
   End
   Begin VB.Line Line2 
      X1              =   4800
      X2              =   5280
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Line Line1 
      X1              =   4800
      X2              =   4800
      Y1              =   2880
      Y2              =   3360
   End
   Begin VB.Image Image7 
      Height          =   495
      Left            =   5760
      Picture         =   "Form3.frx":6282
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   495
   End
   Begin VB.Image Image6 
      Height          =   495
      Left            =   5760
      Picture         =   "Form3.frx":6B68
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   495
   End
   Begin VB.Image Image5 
      Height          =   495
      Left            =   4320
      Picture         =   "Form3.frx":744E
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   495
   End
   Begin VB.Image Image4 
      Height          =   495
      Left            =   5280
      Picture         =   "Form3.frx":7D34
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   495
   End
   Begin VB.Image Image3 
      Height          =   495
      Left            =   5280
      Picture         =   "Form3.frx":861A
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   4800
      Picture         =   "Form3.frx":8F00
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   4800
      Picture         =   "Form3.frx":97E6
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   $"Form3.frx":A0CC
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   7335
   End
   Begin VB.Label Label3 
      Caption         =   $"Form3.frx":A226
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   0
      TabIndex        =   2
      Top             =   1800
      Width           =   7335
   End
   Begin VB.Label Label2 
      Caption         =   "NEVER use more than 1 minifortress! The game only knows to look for 1, so you would go to the same level twice!"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   1440
      Width           =   7335
   End
   Begin VB.Label Label1 
      Caption         =   $"Form3.frx":A3AD
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   7335
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form3.Visible = False
Form1.Show
End Sub

