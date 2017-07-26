VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "494702123 林高遠"
   ClientHeight    =   8040
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14520
   LinkTopic       =   "Form1"
   ScaleHeight     =   8040
   ScaleWidth      =   14520
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command5 
      Caption         =   "Laplace Filter"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   11280
      TabIndex        =   7
      Top             =   6360
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Sobel Filter"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   9840
      TabIndex        =   6
      Top             =   6360
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Prewitt Filter"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   8400
      TabIndex        =   5
      Top             =   6360
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Average Filter"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   6840
      TabIndex        =   4
      Top             =   6360
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Color to gray"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   4440
      TabIndex        =   3
      Top             =   6480
      Width           =   1095
   End
   Begin VB.PictureBox Pic3 
      Height          =   6000
      Left            =   9480
      ScaleHeight     =   5940
      ScaleWidth      =   4440
      TabIndex        =   2
      Top             =   240
      Width           =   4500
   End
   Begin VB.PictureBox Pic2 
      Height          =   6000
      Left            =   4800
      ScaleHeight     =   5940
      ScaleWidth      =   4440
      TabIndex        =   1
      Top             =   240
      Width           =   4500
   End
   Begin VB.PictureBox Pic1 
      Height          =   6000
      Left            =   120
      ScaleHeight     =   5940
      ScaleWidth      =   4440
      TabIndex        =   0
      Top             =   240
      Width           =   4500
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mask(8)   '3x3像素-濾波器變數
Dim X As Long, Y As Long '主像素座標
Dim pix(8) '3x3像素矩陣
Dim pixS As Long  '運算結果像素
Private Sub Command1_Click()
Dim nR, nG, nB, nGray As Integer

For Y = 1 To Pic1.ScaleHeight - 1
    If Y Mod 16 = 0 Then '調控轉換速度用
        DoEvents
    End If
    For X = 1 To Pic1.ScaleWidth - 1
                   
        nR = Pic1.Point(X, Y) And &HFF
        nG = (Pic1.Point(X, Y) \ &H100) And &HFF
        nB = (Pic1.Point(X, Y) \ &H10000) And &HFF
        
        '常用灰階化參數
        nGray = 0.299 * nR + 0.587 * nG + 0.114 * nB
        
        Pic2.PSet (X, Y), RGB(nGray, nGray, nGray)
        
    Next X
Next Y

End Sub

Private Sub Command2_Click()
Dim i, TempX, TempY
    
    For Y = 1 To Pic2.ScaleHeight - 1
        For X = 1 To Pic2.ScaleHeight - 1
            If X Mod 256 = 0 Then '調控轉換速度用
                DoEvents
            End If
                        
            Call RoundPixelLoad '把自己和周圍的八個pixel Load進暫存器
            
            Call mask_Average
            For i = 0 To 8
                TempX = TempX + pix(i) * mask(i)
            Next i
            pixS = Abs(TempX)
            
            
            For i = 0 To 8
                TempY = TempY + pix(i) * mask(i)
            Next i
            pixS = pixS + Abs(TempY)
            
            Pic3.PSet (X, Y), RGB(pixS, pixS, pixS)
            pixS = 0 '歸零
            TempX = 0
            TempY = 0
        Next X
    Next Y
End Sub
Private Sub Command3_Click()
Dim i, TempX, TempY
'Prewitt filter
    'call T2gray 不需要
    
    For Y = 1 To Pic2.ScaleHeight - 1
        For X = 1 To Pic2.ScaleHeight - 1
            If X Mod 128 = 0 Then '調控轉換速度用
                DoEvents
            End If

            Call RoundPixelLoad '把自己和周圍的八個pixel Load進暫存器
            
            Call mask_prewittX
            For i = 0 To 8
                TempX = TempX + pix(i) * mask(i)
            Next i
            pixS = Abs(TempX)
            
            Call mask_prewittY
            For i = 0 To 8
                TempY = TempY + pix(i) * mask(i)
            Next i
            pixS = pixS + Abs(TempY)
            
            Pic3.PSet (X, Y), RGB(pixS, pixS, pixS)
            pixS = 0 '歸零
            TempX = 0
            TempY = 0
        Next X
    Next Y
    'Set Pic2.Picture = Pic2.Image
End Sub
Private Sub Command4_Click()
Dim i, TempX, TempY
    
    For Y = 1 To Pic2.ScaleHeight - 1
        For X = 1 To Pic2.ScaleHeight - 1
            If X Mod 128 = 0 Then '調控轉換速度用
                DoEvents
            End If
                        
            Call RoundPixelLoad '把自己和周圍的八個pixel Load進暫存器
            
            Call mask_SobelX
            For i = 0 To 8
                TempX = TempX + pix(i) * mask(i)
            Next i
            pixS = Abs(TempX)
            
            Call mask_SobelY
            For i = 0 To 8
                TempY = TempY + pix(i) * mask(i)
            Next i
            pixS = pixS + Abs(TempY)
            
            Pic3.PSet (X, Y), RGB(pixS, pixS, pixS)
            pixS = 0 '歸零
            TempX = 0
            TempY = 0
        Next X
    Next Y
    'Set Pic2.Picture = Pic2.Image
End Sub
Private Sub Command5_Click()
Dim i, TempX, TempY
    
    For Y = 1 To Pic2.ScaleHeight - 1
        For X = 1 To Pic2.ScaleHeight - 1
            If X Mod 256 = 0 Then '調控轉換速度用
                DoEvents
            End If
                        
            Call RoundPixelLoad '把自己和周圍的八個pixel Load進暫存器
            
            Call mask_Laplace
            For i = 0 To 8
                TempX = TempX + pix(i) * mask(i)
            Next i
            pixS = Abs(TempX)
            
            
            For i = 0 To 8
                TempY = TempY + pix(i) * mask(i)
            Next i
            pixS = pixS + Abs(TempY)
            
            Pic3.PSet (X, Y), RGB(pixS, pixS, pixS)
            pixS = 0 '歸零
            TempX = 0
            TempY = 0
        Next X
    Next Y
End Sub
Private Sub Form_Load()
'計算單位為像素
Pic1.ScaleMode = 3
Pic2.ScaleMode = 3
Pic3.ScaleMode = 3

'設定自動重繪
Pic1.AutoRedraw = True
Pic2.AutoRedraw = True
Pic3.AutoRedraw = True

Pic1.Picture = LoadPicture(App.Path & "\yang.jpg")

End Sub
Private Sub mask_Average()
Dim i
    For i = 0 To 8
        mask(i) = 1 / 18
    Next i
End Sub

Private Sub mask_prewittX()
Dim i
    For i = 0 To 6 Step 3
        mask(0 + i) = -1
        mask(1 + i) = 0
        mask(2 + i) = 1
    Next i
End Sub
Private Sub mask_prewittY()
Dim i
    For i = 0 To 2
        mask(0 + i) = -1
        mask(3 + i) = 0
        mask(6 + i) = 1
    Next i
End Sub
Private Sub mask_SobelX()
    mask(0) = -1
    mask(1) = 0
    mask(2) = 1
    
    mask(3) = -2
    mask(4) = 0
    mask(5) = 2
    
    mask(6) = -1
    mask(7) = 0
    mask(8) = 1
End Sub
Private Sub mask_SobelY()
    mask(0) = -1
    mask(1) = -2
    mask(2) = -1
    
    mask(3) = 0
    mask(4) = 0
    mask(5) = 0
    
    mask(6) = 1
    mask(7) = 2
    mask(8) = 1
End Sub
Private Sub mask_Laplace()
    mask(0) = -1
    mask(1) = -1
    mask(2) = -1
    
    mask(3) = -1
    mask(4) = 9
    mask(5) = -1
    
    mask(6) = -1
    mask(7) = -1
    mask(8) = -1
End Sub
Private Sub RoundPixelLoad()
    pix(0) = Pic2.Point(X - 1, Y - 1) And &HFF '0 => A
    pix(1) = Pic2.Point(X, Y - 1) And &HFF     '1 => B
    pix(2) = Pic2.Point(X + 1, Y - 1) And &HFF '2 => C
    pix(3) = Pic2.Point(X - 1, Y) And &HFF     '3 => D
    pix(4) = Pic2.Point(X, Y) And &HFF         '4 => E
    pix(5) = Pic2.Point(X + 1, Y) And &HFF     '5 => F
    pix(6) = Pic2.Point(X - 1, Y + 1) And &HFF '6 => G
    pix(7) = Pic2.Point(X, Y + 1) And &HFF     '7 => H
    pix(8) = Pic2.Point(X + 1, Y + 1) And &HFF '8 => I
End Sub
