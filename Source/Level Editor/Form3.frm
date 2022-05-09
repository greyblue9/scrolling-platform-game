VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9975
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   9975
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image1 
      Height          =   1935
      Index           =   8
      Left            =   120
      Picture         =   "Form3.frx":000C
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   1935
      Index           =   7
      Left            =   7320
      Picture         =   "Form3.frx":F432
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   1935
      Index           =   6
      Left            =   4920
      Picture         =   "Form3.frx":14B18
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   1935
      Index           =   5
      Left            =   2520
      Picture         =   "Form3.frx":23F3E
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   1935
      Index           =   4
      Left            =   120
      Picture         =   "Form3.frx":33364
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   1935
      Index           =   3
      Left            =   7320
      Picture         =   "Form3.frx":4278A
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   1935
      Index           =   2
      Left            =   4920
      Picture         =   "Form3.frx":51BB0
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   1935
      Index           =   1
      Left            =   2520
      Picture         =   "Form3.frx":60FD6
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   1935
      Index           =   0
      Left            =   120
      Picture         =   "Form3.frx":703FC
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image1_Click(Index As Integer)
Form1.Image2.Picture = Form3.Image1(Index).Picture
Form1.RichTextBox1.SelStart = Form1.Image1.UBound * 2 + 7
Form1.RichTextBox1.SelLength = 100
Form1.RichTextBox1.SelText = Index
Form3.Visible = False
End Sub
