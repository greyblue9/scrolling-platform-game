VERSION 5.00
Begin VB.Form frmProperties 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Level Properties"
   ClientHeight    =   5505
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   6975
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmProperties.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Level appearance"
      Height          =   1215
      Left            =   240
      TabIndex        =   8
      Top             =   2280
      Width           =   2895
      Begin VB.CommandButton cmdChgColor 
         Caption         =   "&Change..."
         Height          =   375
         Left            =   1680
         TabIndex        =   11
         Top             =   480
         Width           =   975
      End
      Begin VB.PictureBox picColor 
         BackColor       =   &H00FFC0C0&
         Height          =   255
         Left            =   360
         ScaleHeight     =   195
         ScaleWidth      =   795
         TabIndex        =   10
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Background color:"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Underwater level"
      Enabled         =   0   'False
      Height          =   255
      Left            =   4560
      TabIndex        =   7
      Top             =   4320
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dimensions"
      Height          =   1815
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   6495
      Begin VB.HScrollBar scrWidth 
         Height          =   255
         Left            =   240
         Max             =   1000
         Min             =   5
         TabIndex        =   4
         Top             =   600
         Value           =   5
         Width           =   5895
      End
      Begin VB.HScrollBar scrHeight 
         Height          =   255
         Left            =   240
         Max             =   1000
         Min             =   5
         TabIndex        =   3
         Top             =   1320
         Value           =   5
         Width           =   5895
      End
      Begin VB.Label lblWidth 
         Caption         =   "Width: x"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label lblHeight 
         Caption         =   "Height: y"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1080
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5520
      TabIndex        =   1
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   4200
      TabIndex        =   0
      Top             =   4920
      Width           =   1215
   End
End
Attribute VB_Name = "frmProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdChgColor_Click()
frmLvlEdit.CommonDialog1.ShowColor
picColor.BackColor = frmLvlEdit.CommonDialog1.Color
frmLvlEdit.picEdit.BackColor = frmLvlEdit.CommonDialog1.Color
frmLvlEdit.DrawScreen
End Sub

Private Sub cmdOk_click()

If scrWidth.Value + scrHeight.Value > 1100 Then

    Dim tResult
    tResult = MsgBox("Warning: The dimensions you specified result in an enormous map size, which will result in a large file size and long delays in saving and loading! Are you sure you want to make this adjustment?", vbYesNoCancel, Me.Caption)
    
    If Not tResult = 6 Then
        Exit Sub
    End If

End If

lvlWidth = scrWidth.Value
lvlHeight = scrHeight.Value

frmLvlEdit.DrawScreen
frmLvlEdit.ResizeScrollbars

Unload Me
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub Form_Load()
scrWidth.Value = lvlWidth
scrHeight.Value = lvlHeight
picColor.BackColor = frmLvlEdit.picEdit.BackColor
End Sub


Private Sub scrHeight_Scroll()
scrHeight_Change
End Sub

Private Sub scrWidth_Change()
lblWidth.Caption = scrWidth.Value
End Sub

Private Sub scrHeight_Change()
lblHeight.Caption = scrHeight.Value
End Sub

Private Sub scrWidth_Scroll()
scrWidth_Change
End Sub
