VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5475
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9375
   LinkTopic       =   "Form1"
   ScaleHeight     =   5475
   ScaleWidth      =   9375
   StartUpPosition =   3  '系統預設值
   Begin VB.HScrollBar GaryScroll 
      Height          =   615
      Left            =   1920
      TabIndex        =   6
      Top             =   4680
      Width           =   7215
   End
   Begin VB.PictureBox Picture3 
      Height          =   3975
      Left            =   6240
      ScaleHeight     =   3915
      ScaleWidth      =   2835
      TabIndex        =   5
      Top             =   480
      Width           =   2895
   End
   Begin VB.PictureBox Picture2 
      Height          =   3975
      Left            =   3240
      ScaleHeight     =   3915
      ScaleWidth      =   2835
      TabIndex        =   4
      Top             =   480
      Width           =   2895
   End
   Begin VB.PictureBox Picture1 
      Height          =   3975
      Left            =   240
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   3915
      ScaleWidth      =   2835
      TabIndex        =   0
      Top             =   480
      Width           =   2895
   End
   Begin VB.Label Label_Critical 
      Alignment       =   1  '靠右對齊
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   7
      Top             =   4800
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "黑白/負片"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6240
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "灰階"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   2
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "原始"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'設定所有變數必須宣告才能使用
Dim X, Y, nR, nG, nB, nGray As Integer

Private Sub Form_Activate()

GaryScroll.Enabled = False

For X = 1 To Picture1.ScaleWidth - 1
    DoEvents
    For Y = 1 To Picture1.ScaleHeight - 1
                   
        nR = Picture1.Point(X, Y) And &HFF
        nG = (Picture1.Point(X, Y) \ &H100) And &HFF
        nB = (Picture1.Point(X, Y) \ &H10000) And &HFF
        
        '常用灰階化參數
        nGray = 0.299 * nR + 0.587 * nG + 0.114 * nB
        
        Picture2.PSet (X, Y), RGB(nGray, nGray, nGray)
        Picture3.PSet (X, Y), RGB(255 - nGray, 255 - nGray, 255 - nGray)

    Next Y
Next X

GaryScroll.Enabled = True

End Sub

Private Sub Form_Load()
    '載入影像
    Picture1.Picture = LoadPicture(App.Path & "\girl.bmp")
    
    '計算單位為像素
    Picture1.ScaleMode = 3
    Picture2.ScaleMode = 3
    Picture3.ScaleMode = 3
    
    '設定自動重繪
    Picture1.AutoRedraw = True
    Picture2.AutoRedraw = True
    Picture3.AutoRedraw = True
    
    'Scroll Bar 初始化為0~255
    GaryScroll.Max = 255
    GaryScroll.Min = 0
    
End Sub

Private Sub GaryScroll_Change()

'GaryScroll.Value = 128

'顯示拉Bar值
Label_Critical.Caption = GaryScroll.Value

For X = 1 To Picture2.ScaleWidth - 1
    DoEvents
    For Y = 1 To Picture2.ScaleHeight - 1
    
    'Picture2.Point(X, Y) And &HFF 萃取出RGB其中一個值
    If (Picture2.Point(X, Y) And &HFF) >= GaryScroll.Value Then
        nGray = 255
    ElseIf (Picture2.Point(X, Y) And &HFF) < GaryScroll.Value Then
        nGray = 0
    End If
    
    '生出Pic3的一個像素
    Picture3.PSet (X, Y), RGB(nGray, nGray, nGray)
    
    
    Next Y
Next X

End Sub

