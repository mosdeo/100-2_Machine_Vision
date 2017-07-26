VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "494702123 林高遠"
   ClientHeight    =   6090
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15120
   LinkTopic       =   "Form1"
   ScaleHeight     =   406
   ScaleMode       =   3  '像素
   ScaleWidth      =   1008
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command4 
      Caption         =   "monochrome be  closingX2"
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
      Left            =   10680
      TabIndex        =   8
      Top             =   4440
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "monochrome be openingX2"
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
      Left            =   7680
      TabIndex        =   7
      Top             =   4440
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Gray to monochrome"
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
      TabIndex        =   6
      Top             =   4440
      Width           =   2175
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
      Left            =   2160
      TabIndex        =   5
      Top             =   4440
      Width           =   1095
   End
   Begin VB.PictureBox Pic5 
      Height          =   3855
      Left            =   11160
      ScaleHeight     =   253
      ScaleMode       =   3  '像素
      ScaleWidth      =   253
      TabIndex        =   4
      Top             =   720
      Width           =   3855
   End
   Begin VB.PictureBox Pic4 
      Height          =   3855
      Left            =   8640
      ScaleHeight     =   253
      ScaleMode       =   3  '像素
      ScaleWidth      =   253
      TabIndex        =   3
      Top             =   600
      Width           =   3855
   End
   Begin VB.PictureBox Pic3 
      Height          =   3855
      Left            =   5760
      ScaleHeight     =   253
      ScaleMode       =   3  '像素
      ScaleWidth      =   253
      TabIndex        =   2
      Top             =   480
      Width           =   3855
   End
   Begin VB.PictureBox Pic2 
      Height          =   3855
      Left            =   2880
      ScaleHeight     =   253
      ScaleMode       =   3  '像素
      ScaleWidth      =   253
      TabIndex        =   1
      Top             =   360
      Width           =   3855
   End
   Begin VB.PictureBox Pic1 
      Height          =   3855
      Left            =   240
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   3795
      ScaleWidth      =   3795
      TabIndex        =   0
      Top             =   240
      Width           =   3855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()
Dim X, Y, nR, nG, nB, nGray As Integer

For X = 1 To Pic1.ScaleWidth - 1
    DoEvents
    For Y = 1 To Pic1.ScaleHeight - 1
                   
        nR = Pic1.Point(X, Y) And &HFF
        nG = (Pic1.Point(X, Y) \ &H100) And &HFF
        nB = (Pic1.Point(X, Y) \ &H10000) And &HFF
        
        '常用灰階化參數
        nGray = 0.299 * nR + 0.587 * nG + 0.114 * nB
        
        Pic2.PSet (X, Y), RGB(nGray, nGray, nGray)
        
    Next Y
Next X

End Sub

Private Sub Command2_Click()
Dim X, Y, nGray As Integer
For X = 1 To Pic2.ScaleWidth - 1
    DoEvents
    For Y = 1 To Pic2.ScaleHeight - 1
    
    'Pic2.Point(X, Y) And &HFF 萃取出RGB其中一個值
    If (Pic2.Point(X, Y) And &HFF) >= 127 Then
        nGray = 255
    ElseIf (Pic2.Point(X, Y) And &HFF) < 127 Then
        nGray = 0
    End If
    
    '生出Pic3的一個像素
    Pic3.PSet (X, Y), RGB(nGray, nGray, nGray)
    
    
    Next Y
Next X
End Sub

Private Sub Command3_Click()
ErosionAndDilation_to_pic4_or_pic5 2, 4, 5, 4 '次數,方向數,蝕or長,pic4or5
ErosionAndDilation_to_pic4_or_pic5 2, 4, 4, 4
End Sub

Private Sub Command4_Click()
ErosionAndDilation_to_pic4_or_pic5 2, 4, 4, 5 '次數,方向數,蝕or長,寫入pic4or5
ErosionAndDilation_to_pic4_or_pic5 2, 4, 5, 5
End Sub

Private Sub Form_Load()
'計算單位為像素
Pic1.ScaleMode = 3
Pic2.ScaleMode = 3
Pic3.ScaleMode = 3
Pic4.ScaleMode = 3
Pic5.ScaleMode = 3

'設定自動重繪
Pic1.AutoRedraw = True
Pic2.AutoRedraw = True
Pic3.AutoRedraw = True
Pic4.AutoRedraw = True
Pic5.AutoRedraw = True

Pic1.AutoSize = True
Pic2.AutoSize = True
Pic3.AutoSize = True
Pic4.AutoSize = True
Pic5.AutoSize = True

Pic1.Picture = LoadPicture(App.Path & "\lena.jpg")

End Sub

Private Sub ErosionAndDilation_to_pic4_or_pic5(p_nTimes As Integer, p_nType As Integer, EorD As Integer, p_nPic As Integer)
    
'座標,次數
Dim X As Long, Y As Long, N As Integer

'3X3矩陣
'    Dim pixA As Byte, pixB As Byte, pixC As Byte
'    Dim pixD As Byte, pixE As Byte, pixF As Byte
'    Dim pixG As Byte, pixH As Byte, pixI As Byte
Dim pix(8) As Byte

'4方向 8方向
Dim intFour As Byte, intEight As Byte


For N = 1 To p_nTimes
    For Y = 0 To Pic3.ScaleHeight - 1 '有邊界問題所以要-1
        If Y Mod 2 = 0 Then
            DoEvents
        End If
        For X = 0 To Pic3.ScaleWidth - 1
        
            pix(0) = Pic3.Point(X - 1, Y - 1) And &HFF '0 => A
            pix(1) = Pic3.Point(X, Y - 1) And &HFF     '1 => B
            pix(2) = Pic3.Point(X + 1, Y - 1) And &HFF '2 => C
            pix(3) = Pic3.Point(X - 1, Y) And &HFF     '3 => D
            pix(4) = Pic3.Point(X, Y) And &HFF         '4 => E
            pix(5) = Pic3.Point(X + 1, Y) And &HFF     '5 => F
            pix(6) = Pic2.Point(X - 1, Y + 1) And &HFF '6 => G
            pix(7) = Pic2.Point(X, Y + 1) And &HFF     '7 => H
            pix(8) = Pic2.Point(X + 1, Y + 1) And &HFF '8 => I
            
            Select Case EorD '(5or4) '(5=E,4=D)
                Case 5: '侵蝕
                    If p_nType = 8 Then '八方向
                        
                        If (pix(0) Or pix(1) Or pix(2) Or pix(3) Or pix(5) Or pix(6) Or pix(7) Or pix(8)) = 255 Then
                            intEight = 255
                        Else
                            intEight = pix(4)
                        End If
                        
                        If p_nPic = 4 Then
                            Pic4.PSet (X, Y), RGB(intEight, intEight, intEight)
                        ElseIf p_nPic = 5 Then
                            Pic5.PSet (X, Y), RGB(intEight, intEight, intEight)
                        End If
                        
                    ElseIf p_nType = 4 Then '四方向
                        
                        If (pix(1) Or pix(3) Or pix(5) Or pix(7)) = 255 Then
                            intFour = 255
                        Else
                            intFour = pix(4)
                        End If
                        
                        If p_nPic = 4 Then
                            Pic4.PSet (X, Y), RGB(intFour, intFour, intFour)
                        ElseIf p_nPic = 5 Then
                            Pic5.PSet (X, Y), RGB(intFour, intFour, intFour)
                        End If
                        
                    End If
                    
                Case 4: '增長
                    If p_nType = 8 Then '八方向
                        
                        If (pix(0) And pix(1) And pix(2) And pix(3) And pix(5) And pix(6) And pix(7) And pix(8)) = 255 Then
                            intEight = pix(4)
                        Else
                            intEight = 0
                        End If
                        
                        If p_nPic = 4 Then
                            Pic4.PSet (X, Y), RGB(intEight, intEight, intEight)
                        ElseIf p_nPic = 5 Then
                            Pic5.PSet (X, Y), RGB(intEight, intEight, intEight)
                        End If
                        
                     
                    ElseIf p_nType = 4 Then '四方向
                        
                        If (pix(1) And pix(3) And pix(5) And pix(7)) = 255 Then
                            intFour = pix(4)
                        Else
                            intFour = 0
                        End If
                        
                        If p_nPic = 4 Then
                            Pic4.PSet (X, Y), RGB(intFour, intFour, intFour)
                        ElseIf p_nPic = 5 Then
                            Pic5.PSet (X, Y), RGB(intFour, intFour, intFour)
                        End If
                        
                    End If
            End Select
                
                
        Next X
    Next Y
Next N
End Sub
