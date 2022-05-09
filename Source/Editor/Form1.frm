VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SBS3 Level Editor"
   ClientHeight    =   8295
   ClientLeft      =   4350
   ClientTop       =   1680
   ClientWidth     =   9615
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Palette         =   "Form1.frx":0442
   PaletteMode     =   2  'Custom
   ScaleHeight     =   8295
   ScaleWidth      =   9615
   Begin VB.CommandButton Command2 
      Caption         =   "Load"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8640
      TabIndex        =   2
      Top             =   7800
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7680
      TabIndex        =   0
      Top             =   7800
      Width           =   975
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Clear all tiles in map to..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      TabIndex        =   5
      Top             =   7800
      Width           =   855
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   7800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FileName        =   "Map.txt"
      Filter          =   "*.txt"
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   7800
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   873
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"Form1.frx":0D28
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Color"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   3
      Top             =   7800
      Width           =   1095
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   154
      Left            =   6720
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   319
      Left            =   9120
      Stretch         =   -1  'True
      Top             =   7200
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   318
      Left            =   8640
      Stretch         =   -1  'True
      Top             =   7200
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   317
      Left            =   8160
      Stretch         =   -1  'True
      Top             =   7200
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   316
      Left            =   7680
      Stretch         =   -1  'True
      Top             =   7200
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   315
      Left            =   7200
      Stretch         =   -1  'True
      Top             =   7200
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   314
      Left            =   6720
      Stretch         =   -1  'True
      Top             =   7200
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   313
      Left            =   6240
      Stretch         =   -1  'True
      Top             =   7200
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   312
      Left            =   5760
      Stretch         =   -1  'True
      Top             =   7200
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   311
      Left            =   5280
      Stretch         =   -1  'True
      Top             =   7200
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   310
      Left            =   4800
      Stretch         =   -1  'True
      Top             =   7200
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   309
      Left            =   4320
      Stretch         =   -1  'True
      Top             =   7200
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   308
      Left            =   3840
      Stretch         =   -1  'True
      Top             =   7200
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   307
      Left            =   3360
      Stretch         =   -1  'True
      Top             =   7200
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   306
      Left            =   2880
      Stretch         =   -1  'True
      Top             =   7200
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   305
      Left            =   2400
      Stretch         =   -1  'True
      Top             =   7200
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   304
      Left            =   1920
      Stretch         =   -1  'True
      Top             =   7200
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   303
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   7200
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   302
      Left            =   960
      Stretch         =   -1  'True
      Top             =   7200
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   301
      Left            =   480
      Stretch         =   -1  'True
      Top             =   7200
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   300
      Left            =   0
      Stretch         =   -1  'True
      Top             =   7200
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   299
      Left            =   9120
      Stretch         =   -1  'True
      Top             =   6720
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   298
      Left            =   8640
      Stretch         =   -1  'True
      Top             =   6720
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   297
      Left            =   8160
      Stretch         =   -1  'True
      Top             =   6720
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   296
      Left            =   7680
      Stretch         =   -1  'True
      Top             =   6720
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   295
      Left            =   7200
      Stretch         =   -1  'True
      Top             =   6720
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   294
      Left            =   6720
      Stretch         =   -1  'True
      Top             =   6720
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   293
      Left            =   6240
      Stretch         =   -1  'True
      Top             =   6720
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   292
      Left            =   5760
      Stretch         =   -1  'True
      Top             =   6720
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   291
      Left            =   5280
      Stretch         =   -1  'True
      Top             =   6720
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   290
      Left            =   4800
      Stretch         =   -1  'True
      Top             =   6720
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   289
      Left            =   4320
      Stretch         =   -1  'True
      Top             =   6720
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   288
      Left            =   3840
      Stretch         =   -1  'True
      Top             =   6720
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   287
      Left            =   3360
      Stretch         =   -1  'True
      Top             =   6720
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   286
      Left            =   2880
      Stretch         =   -1  'True
      Top             =   6720
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   285
      Left            =   2400
      Stretch         =   -1  'True
      Top             =   6720
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   284
      Left            =   1920
      Stretch         =   -1  'True
      Top             =   6720
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   283
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   6720
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   282
      Left            =   960
      Stretch         =   -1  'True
      Top             =   6720
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   281
      Left            =   480
      Stretch         =   -1  'True
      Top             =   6720
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   280
      Left            =   0
      Stretch         =   -1  'True
      Top             =   6720
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   279
      Left            =   9120
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   278
      Left            =   8640
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   277
      Left            =   8160
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   276
      Left            =   7680
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   275
      Left            =   7200
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   274
      Left            =   6720
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   273
      Left            =   6240
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   272
      Left            =   5760
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   271
      Left            =   5280
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   270
      Left            =   4800
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   269
      Left            =   4320
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   268
      Left            =   3840
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   267
      Left            =   3360
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   266
      Left            =   2880
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   265
      Left            =   2400
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   264
      Left            =   1920
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   263
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   262
      Left            =   960
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   261
      Left            =   480
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   260
      Left            =   0
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   259
      Left            =   9120
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   258
      Left            =   8640
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   257
      Left            =   8160
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   256
      Left            =   7680
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   255
      Left            =   7200
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   254
      Left            =   6720
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   253
      Left            =   6240
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   252
      Left            =   5760
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   251
      Left            =   5280
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   250
      Left            =   4800
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   249
      Left            =   4320
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   248
      Left            =   3840
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   247
      Left            =   3360
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   246
      Left            =   2880
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   245
      Left            =   2400
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   244
      Left            =   1920
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   243
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   242
      Left            =   960
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   241
      Left            =   480
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   240
      Left            =   0
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   239
      Left            =   9120
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   238
      Left            =   8640
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   237
      Left            =   8160
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   236
      Left            =   7680
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   235
      Left            =   7200
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   234
      Left            =   6720
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   233
      Left            =   6240
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   232
      Left            =   5760
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   231
      Left            =   5280
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   230
      Left            =   4800
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   229
      Left            =   4320
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   228
      Left            =   3840
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   227
      Left            =   3360
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   226
      Left            =   2880
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   225
      Left            =   2400
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   224
      Left            =   1920
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   223
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   222
      Left            =   960
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   221
      Left            =   480
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   220
      Left            =   0
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   219
      Left            =   9120
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   218
      Left            =   8640
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   217
      Left            =   8160
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   216
      Left            =   7680
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   215
      Left            =   7200
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   214
      Left            =   6720
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   213
      Left            =   6240
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   212
      Left            =   5760
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   211
      Left            =   5280
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   210
      Left            =   4800
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   209
      Left            =   4320
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   208
      Left            =   3840
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   207
      Left            =   3360
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   206
      Left            =   2880
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   205
      Left            =   2400
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   204
      Left            =   1920
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   203
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   202
      Left            =   960
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   201
      Left            =   480
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   200
      Left            =   0
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   199
      Left            =   9120
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   198
      Left            =   8640
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   197
      Left            =   8160
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   196
      Left            =   7680
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   195
      Left            =   7200
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   194
      Left            =   6720
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   193
      Left            =   6240
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   192
      Left            =   5760
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   191
      Left            =   5280
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   190
      Left            =   4800
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   189
      Left            =   4320
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   188
      Left            =   3840
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   187
      Left            =   3360
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   186
      Left            =   2880
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   185
      Left            =   2400
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   184
      Left            =   1920
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   183
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   182
      Left            =   960
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   181
      Left            =   480
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   180
      Left            =   0
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   179
      Left            =   9120
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   178
      Left            =   8640
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   177
      Left            =   8160
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   176
      Left            =   7680
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   175
      Left            =   7200
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   174
      Left            =   6720
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   173
      Left            =   6240
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   172
      Left            =   5760
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   171
      Left            =   5280
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   170
      Left            =   4800
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   169
      Left            =   4320
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   168
      Left            =   3840
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   167
      Left            =   3360
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   166
      Left            =   2880
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   165
      Left            =   2400
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   164
      Left            =   1920
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   163
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   162
      Left            =   960
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   161
      Left            =   480
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   160
      Left            =   0
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   159
      Left            =   9120
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   158
      Left            =   8640
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   157
      Left            =   8160
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   156
      Left            =   7680
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   155
      Left            =   7200
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   153
      Left            =   6240
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   152
      Left            =   5760
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   151
      Left            =   5280
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   150
      Left            =   4800
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   149
      Left            =   4320
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   148
      Left            =   3840
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   147
      Left            =   3360
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   146
      Left            =   2880
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   145
      Left            =   2400
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   144
      Left            =   1920
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   143
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   142
      Left            =   960
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   141
      Left            =   480
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   140
      Left            =   0
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   139
      Left            =   9120
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   138
      Left            =   8640
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   137
      Left            =   8160
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   136
      Left            =   7680
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   135
      Left            =   7200
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   134
      Left            =   6720
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   133
      Left            =   6240
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   132
      Left            =   5760
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   131
      Left            =   5280
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   130
      Left            =   4800
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   129
      Left            =   4320
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   128
      Left            =   3840
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   127
      Left            =   3360
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   126
      Left            =   2880
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   125
      Left            =   2400
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   124
      Left            =   1920
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   123
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   122
      Left            =   960
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   121
      Left            =   480
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   120
      Left            =   0
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   119
      Left            =   9120
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   118
      Left            =   8640
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   117
      Left            =   8160
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   116
      Left            =   7680
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   115
      Left            =   7200
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   114
      Left            =   6720
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   113
      Left            =   6240
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   112
      Left            =   5760
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   111
      Left            =   5280
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   110
      Left            =   4800
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   109
      Left            =   4320
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   108
      Left            =   3840
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   107
      Left            =   3360
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   106
      Left            =   2880
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   105
      Left            =   2400
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   104
      Left            =   1920
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   103
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   102
      Left            =   960
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   101
      Left            =   480
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   100
      Left            =   0
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   99
      Left            =   9120
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   98
      Left            =   8640
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   97
      Left            =   8160
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   96
      Left            =   7680
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   95
      Left            =   7200
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   94
      Left            =   6720
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   93
      Left            =   6240
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   92
      Left            =   5760
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   91
      Left            =   5280
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   90
      Left            =   4800
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   89
      Left            =   4320
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   88
      Left            =   3840
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   87
      Left            =   3360
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   86
      Left            =   2880
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   85
      Left            =   2400
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   84
      Left            =   1920
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   83
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   82
      Left            =   960
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   81
      Left            =   480
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   80
      Left            =   0
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   79
      Left            =   9120
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   78
      Left            =   8640
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   77
      Left            =   8160
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   76
      Left            =   7680
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   75
      Left            =   7200
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   74
      Left            =   6720
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   73
      Left            =   6240
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   72
      Left            =   5760
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   71
      Left            =   5280
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   70
      Left            =   4800
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   69
      Left            =   4320
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   68
      Left            =   3840
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   67
      Left            =   3360
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   66
      Left            =   2880
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   65
      Left            =   2400
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   64
      Left            =   1920
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   63
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   62
      Left            =   960
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   61
      Left            =   480
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   60
      Left            =   0
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   59
      Left            =   9120
      Stretch         =   -1  'True
      Top             =   960
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   58
      Left            =   8640
      Stretch         =   -1  'True
      Top             =   960
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   57
      Left            =   8160
      Stretch         =   -1  'True
      Top             =   960
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   56
      Left            =   7680
      Stretch         =   -1  'True
      Top             =   960
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   55
      Left            =   7200
      Stretch         =   -1  'True
      Top             =   960
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   54
      Left            =   6720
      Stretch         =   -1  'True
      Top             =   960
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   53
      Left            =   6240
      Stretch         =   -1  'True
      Top             =   960
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   52
      Left            =   5760
      Stretch         =   -1  'True
      Top             =   960
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   51
      Left            =   5280
      Stretch         =   -1  'True
      Top             =   960
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   50
      Left            =   4800
      Stretch         =   -1  'True
      Top             =   960
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   49
      Left            =   4320
      Stretch         =   -1  'True
      Top             =   960
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   48
      Left            =   3840
      Stretch         =   -1  'True
      Top             =   960
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   47
      Left            =   3360
      Stretch         =   -1  'True
      Top             =   960
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   46
      Left            =   2880
      Stretch         =   -1  'True
      Top             =   960
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   45
      Left            =   2400
      Stretch         =   -1  'True
      Top             =   960
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   44
      Left            =   1920
      Stretch         =   -1  'True
      Top             =   960
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   43
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   960
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   42
      Left            =   960
      Stretch         =   -1  'True
      Top             =   960
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   41
      Left            =   480
      Stretch         =   -1  'True
      Top             =   960
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   40
      Left            =   0
      Stretch         =   -1  'True
      Top             =   960
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   39
      Left            =   9120
      Stretch         =   -1  'True
      Top             =   480
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   38
      Left            =   8640
      Stretch         =   -1  'True
      Top             =   480
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   37
      Left            =   8160
      Stretch         =   -1  'True
      Top             =   480
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   36
      Left            =   7680
      Stretch         =   -1  'True
      Top             =   480
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   35
      Left            =   7200
      Stretch         =   -1  'True
      Top             =   480
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   34
      Left            =   6720
      Stretch         =   -1  'True
      Top             =   480
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   33
      Left            =   6240
      Stretch         =   -1  'True
      Top             =   480
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   32
      Left            =   5760
      Stretch         =   -1  'True
      Top             =   480
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   31
      Left            =   5280
      Stretch         =   -1  'True
      Top             =   480
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   30
      Left            =   4800
      Stretch         =   -1  'True
      Top             =   480
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   29
      Left            =   4320
      Stretch         =   -1  'True
      Top             =   480
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   28
      Left            =   3840
      Stretch         =   -1  'True
      Top             =   480
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   27
      Left            =   3360
      Stretch         =   -1  'True
      Top             =   480
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   26
      Left            =   2880
      Stretch         =   -1  'True
      Top             =   480
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   25
      Left            =   2400
      Stretch         =   -1  'True
      Top             =   480
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   24
      Left            =   1920
      Stretch         =   -1  'True
      Top             =   480
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   23
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   480
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   22
      Left            =   960
      Stretch         =   -1  'True
      Top             =   480
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   21
      Left            =   480
      Stretch         =   -1  'True
      Top             =   480
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   20
      Left            =   0
      Stretch         =   -1  'True
      Top             =   480
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   19
      Left            =   9120
      Stretch         =   -1  'True
      Top             =   0
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   18
      Left            =   8640
      Stretch         =   -1  'True
      Top             =   0
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   17
      Left            =   8160
      Stretch         =   -1  'True
      Top             =   0
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   16
      Left            =   7680
      Stretch         =   -1  'True
      Top             =   0
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   15
      Left            =   7200
      Stretch         =   -1  'True
      Top             =   0
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   14
      Left            =   6720
      Stretch         =   -1  'True
      Top             =   0
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   13
      Left            =   6240
      Stretch         =   -1  'True
      Top             =   0
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   12
      Left            =   5760
      Stretch         =   -1  'True
      Top             =   0
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   11
      Left            =   5280
      Stretch         =   -1  'True
      Top             =   0
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   10
      Left            =   4800
      Stretch         =   -1  'True
      Top             =   0
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   9
      Left            =   4320
      Stretch         =   -1  'True
      Top             =   0
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   8
      Left            =   3840
      Stretch         =   -1  'True
      Top             =   0
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   7
      Left            =   3360
      Stretch         =   -1  'True
      Top             =   0
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   6
      Left            =   2880
      Stretch         =   -1  'True
      Top             =   0
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   5
      Left            =   2400
      Stretch         =   -1  'True
      Top             =   0
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   4
      Left            =   1920
      Stretch         =   -1  'True
      Top             =   0
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   3
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   0
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   2
      Left            =   960
      Stretch         =   -1  'True
      Top             =   0
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   1
      Left            =   480
      Stretch         =   -1  'True
      Top             =   0
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   0
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "News Gothic MT"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   7695
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   9615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public STP
Public Resp
Private Sub Command1_Click()
MsgBox "Once you press Ok, please make the one change that happens after you beat the Minifortress.", , "''Change'' Data Entry"
STP = 1
End Sub

Private Sub Command2_Click()
CommonDialog1.ShowOpen
RichTextBox1.LoadFile CommonDialog1.FileName
For I = 0 To Image1.UBound
RichTextBox1.SelStart = I * 2
RichTextBox1.SelLength = 2
Image1(I).Picture = Form2.Image1(RichTextBox1.SelText).Picture
Next I
RichTextBox1.SelStart = Image1.UBound * 2 + 7
RichTextBox1.SelLength = 100
Label1.BackColor = RichTextBox1.SelText
Form2.Label1.BackColor = RichTextBox1.SelText
CommonDialog1.Color = RichTextBox1.SelText
End Sub


Private Sub Command4_Click()
Form3.Show
End Sub


Private Sub Command6_Click()
CommonDialog1.ShowColor
Label1.BackColor = CommonDialog1.Color
Form2.Label1.BackColor = CommonDialog1.Color
RichTextBox1.SelStart = Image1.UBound * 2 + 7
RichTextBox1.SelLength = 100
RichTextBox1.SelText = CommonDialog1.Color
End Sub

Private Sub Command7_Click()
Resp = MsgBox("Are you sure you want to clear this map? Everything will be the specified tile and you will lose all world data.(Ex - color and minifortress data)", vbOKCancel, "Warning")
If Resp = 1 Then
RichTextBox1.Text = ""
Resp = InputBox("Please enter the tile art number of the tile that you wish to cover the map. Note: To make the entire map land, enter 34. To make the entire map water, enter 80.", "Fill-In", 80)
If Resp > Form2.Image1.UBound Then
MsgBox "That tile art number does not exist!"
Exit Sub
End If
If Resp < 0 Then
MsgBox "That tile art number does not exist!"
Exit Sub
End If
If Resp < 10 Then
For I = 0 To Image1.UBound
RichTextBox1.Text = RichTextBox1.Text & "0" & Resp
Image1(I).Picture = Form2.Image1(Resp).Picture
Next I
End If
If Resp > 9 Then
For I = 0 To Image1.UBound
RichTextBox1.Text = RichTextBox1.Text & Resp
Image1(I).Picture = Form2.Image1(Resp).Picture
Next I
End If
End If
End Sub

Private Sub Form_Load()
STP = 0
Form2.Show
For I = 0 To Image1.UBound
RichTextBox1.Text = RichTextBox1.Text & "80"
Image1(I).Picture = Form2.Image1(80).Picture
Next I
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
End
End Sub
Private Sub Image1_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
If STP = 0 Then
Image1(Index).Picture = Form2.Image2.Picture
RichTextBox1.SelStart = Index * 2
RichTextBox1.SelLength = 2
RichTextBox1.SelText = Form2.CurIndex
End If
If STP = 1 Then
RichTextBox1.Text = RichTextBox1.Text & Form2.CurIndex
RichTextBox1.Text = RichTextBox1.Text & Index
CommonDialog1.ShowSave
RichTextBox1.SaveFile CommonDialog1.FileName
MsgBox "Good! Save complete. Change information: Change tile #" & Index & " to tile art #" & Form2.CurIndex & ".", , "Save OK"
STP = 0
End If
End Sub
Private Sub Image1_DragOver(Index As Integer, Source As Control, X As Single, Y As Single, State As Integer)
If STP = 0 Then
Image1(Index).Picture = Form2.Image2.Picture
RichTextBox1.SelStart = Index * 2
RichTextBox1.SelLength = 2
RichTextBox1.SelText = Form2.CurIndex
End If
If STP = 1 Then
RichTextBox1.SelStart = Image1.UBound * 2 + 2
RichTextBox1.SelLength = 2
RichTextBox1.SelText = Form2.CurIndex
RichTextBox1.SelStart = Image1.UBound * 2 + 4
RichTextBox1.SelLength = 3
If Index > 99 Then
If Index < 1000 Then
RichTextBox1.SelText = Index
End If
End If
If Index > 9 Then
If Index < 100 Then
RichTextBox1.SelText = "0" & Index
End If
End If
If Index > -1 Then
If Index < 10 Then
RichTextBox1.SelText = "00" & Index
End If
End If
RichTextBox1.SelStart = Image1.UBound * 2 + 7
RichTextBox1.SelLength = 100
RichTextBox1.SelText = CommonDialog1.Color


CommonDialog1.ShowSave
RichTextBox1.SaveFile CommonDialog1.FileName
MsgBox "Good! Save complete. Change information: Change tile #" & Index & " to tile art #" & Form2.CurIndex & "."
STP = 0
End If
End Sub
