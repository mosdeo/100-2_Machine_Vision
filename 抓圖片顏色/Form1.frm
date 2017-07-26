VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   7530
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   17070
   LinkTopic       =   "Form1"
   ScaleHeight     =   502
   ScaleMode       =   0  '使用者自訂
   ScaleWidth      =   1138
   StartUpPosition =   1  '所屬視窗中央
   Begin VB.CommandButton Command1 
      Caption         =   "抓紅色"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7920
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
   Begin VB.PictureBox Picture2 
      Height          =   7575
      Left            =   9600
      ScaleHeight     =   7500
      ScaleMode       =   0  '使用者自訂
      ScaleWidth      =   7623.711
      TabIndex        =   1
      Top             =   0
      Width           =   7455
   End
   Begin VB.PictureBox Picture1 
      Height          =   7635
      Left            =   0
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   7575
      ScaleWidth      =   7575
      TabIndex        =   0
      Top             =   0
      Width           =   7635
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'設定所有變數必須宣告才能使用
Dim X, Y, nR, nG, nB As Integer
Dim dY As Double, dCr As Double, dCb As Double

Private Sub Command1_Click()


For X = 1 To Picture1.ScaleWidth - 1
        If X Mod 2 = 0 Then
            DoEvents
        End If
    For Y = 1 To Picture1.ScaleHeight - 1
           
        
        nR = Picture1.Point(X, Y) And &HFF
        nG = (Picture1.Point(X, Y) \ &H100) And &HFF
        nB = (Picture1.Point(X, Y) \ &H10000) And &HFF
        
        If (nR > 90) And (20 < (nR - nG) < 46) And (38 < (nR - nB) < 93) And Max(nR, nG, nB) < 250 Then
            Picture2.PSet (X, Y), RGB(nR, nG, nB)
        Else
            Picture2.PSet (X, Y), RGB(0, 0, 0)
        End If
        
'        Call RGBtoYCbCr(nR, nG, nB)
'
'        If (dY > 120) And (dCb < 95) And (dCr < 110) Then
'            Picture2.PSet (X, Y), RGB(nR, nG, nB)
'        Else
'            Picture2.PSet (X, Y), RGB(0, 0, 0)
'        End If
        
        
    Next Y
Next X

            
End Sub

Private Sub Form_Load()
    'Picture1.Picture = LoadPicture(App.Path & "\Head.bmp")
    '載入影像
    Picture1.ScaleMode = 3
    Picture2.ScaleMode = 3
    '計算單位為像素
    Picture1.AutoRedraw = True
    Picture2.AutoRedraw = True
    '設定自動重繪
End Sub
Private Function Min(ParamArray Vals())
  Dim n As Integer, MinVal
  
  MinVal = Vals(0)
  
  For n = 0 To UBound(Vals)
    If Vals(n) < MinVal Then MinVal = Vals(n)
  Next n
  Min = MinVal
End Function
Private Function Max(ParamArray Vals())
  Dim n As Integer, MaxVal
  
  For n = 0 To UBound(Vals)
    If Vals(n) > MaxVal Then MaxVal = Vals(n)
  Next n
  Max = MaxVal
End Function

Private Function RGBtoYCbCr(ByVal R, G, B)
dY = 0.299 * R + 0.587 * G + 0.114 * B
dCr = (R - dY) * 0.713 + 128
dCb = (B - dY) * 0.564 + 128
End Function
