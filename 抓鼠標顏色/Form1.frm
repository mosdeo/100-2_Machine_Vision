VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7785
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13245
   LinkTopic       =   "Form1"
   ScaleHeight     =   7785
   ScaleWidth      =   13245
   StartUpPosition =   3  '系統預設值
   Begin VB.PictureBox Picture1 
      Height          =   7215
      Left            =   240
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   7155
      ScaleWidth      =   9555
      TabIndex        =   0
      Top             =   240
      Width           =   9615
   End
   Begin VB.Label LabelB 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   27.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10200
      TabIndex        =   3
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label LabelG 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   27.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10200
      TabIndex        =   2
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Label LabelR 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   27.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10200
      TabIndex        =   1
      Top             =   360
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
LabelR.Caption = "R = " & Str((Picture1.Point(X, Y)) And &HFF)
LabelG.Caption = "G = " & Str((Picture1.Point(X, Y) \ &H100) And &HFF)
LabelB.Caption = "B = " & Str((Picture1.Point(X, Y) \ &H1000) And &HFF)





End Sub
