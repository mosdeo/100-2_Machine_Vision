VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8745
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15420
   LinkTopic       =   "Form1"
   ScaleHeight     =   8745
   ScaleWidth      =   15420
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command6 
      Caption         =   "(6) 找邊緣"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   12960
      TabIndex        =   15
      Top             =   4440
      Width           =   2295
   End
   Begin VB.PictureBox PicTemplate 
      Height          =   615
      Left            =   840
      ScaleHeight     =   555
      ScaleWidth      =   435
      TabIndex        =   12
      Top             =   4200
      Width           =   495
   End
   Begin VB.PictureBox Pic 
      Height          =   3375
      Left            =   120
      ScaleHeight     =   3315
      ScaleWidth      =   2835
      TabIndex        =   11
      Top             =   240
      Width           =   2895
   End
   Begin VB.CommandButton Command5 
      Caption         =   "(5) 比樣版"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   12960
      TabIndex        =   10
      Top             =   3600
      Width           =   2295
   End
   Begin VB.CommandButton Command4 
      Caption         =   "(4) 侵蝕　"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   12960
      TabIndex        =   9
      Top             =   2760
      Width           =   2295
   End
   Begin VB.CommandButton Command3 
      Caption         =   "(3) 二值化"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   12960
      TabIndex        =   8
      Top             =   1920
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "(2) 轉灰階"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   12960
      TabIndex        =   7
      Top             =   1080
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "(1) 取藍色"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   12960
      TabIndex        =   6
      Top             =   240
      Width           =   2295
   End
   Begin VB.PictureBox PicF 
      Height          =   3375
      Left            =   9840
      ScaleHeight     =   3315
      ScaleWidth      =   2835
      TabIndex        =   5
      Top             =   3840
      Width           =   2895
   End
   Begin VB.PictureBox PicE 
      Height          =   3375
      Left            =   6600
      ScaleHeight     =   3315
      ScaleWidth      =   2835
      TabIndex        =   4
      Top             =   3840
      Width           =   2895
   End
   Begin VB.PictureBox PicD 
      Height          =   3375
      Left            =   3360
      ScaleHeight     =   3315
      ScaleWidth      =   2835
      TabIndex        =   3
      Top             =   3840
      Width           =   2895
   End
   Begin VB.PictureBox PicC 
      Height          =   3375
      Left            =   9840
      ScaleHeight     =   3315
      ScaleWidth      =   2835
      TabIndex        =   2
      Top             =   240
      Width           =   2895
   End
   Begin VB.PictureBox PicB 
      Height          =   3375
      Left            =   6600
      ScaleHeight     =   3315
      ScaleWidth      =   2835
      TabIndex        =   1
      Top             =   240
      Width           =   2895
   End
   Begin VB.PictureBox PicA 
      Height          =   3375
      Left            =   3360
      ScaleHeight     =   3315
      ScaleWidth      =   2835
      TabIndex        =   0
      Top             =   240
      Width           =   2895
   End
   Begin VB.Label LabelY 
      Caption         =   "Y"
      Height          =   495
      Left            =   6720
      TabIndex        =   14
      Top             =   8040
      Width           =   2535
   End
   Begin VB.Label LabelX 
      Caption         =   "X"
      Height          =   495
      Left            =   6720
      TabIndex        =   13
      Top             =   7440
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim X As Long, Y As Long
Dim nR, nG, nB, nGray

'*************  For Edge Filter ************
Dim mask(8)   '3x3像素-濾波器變數
Dim pix(8) '3x3像素矩陣
Dim pixS As Long  '運算結果像素
Private Sub Command1_Click()

For X = 1 To Pic.ScaleWidth - 1
        If X Mod 2 = 0 Then
            DoEvents
        End If
    For Y = 1 To Pic.ScaleHeight - 1
           
        
        nR = Pic.Point(X, Y) And &HFF
        nG = (Pic.Point(X, Y) \ &H100) And &HFF
        nB = (Pic.Point(X, Y) \ &H10000) And &HFF
        
        If (nB > (nR + nG)) Then
            PicA.PSet (X, Y), RGB(nR, nG, nB)
        Else
            PicA.PSet (X, Y), RGB(0, 0, 0)
        End If
        
    Next Y
Next X
End Sub

Private Sub Command2_Click()
For X = 1 To Pic.ScaleWidth - 1
        If X Mod 2 = 0 Then
            DoEvents
        End If
        
    For Y = 1 To Pic.ScaleHeight - 1
           
        
        nR = Pic.Point(X, Y) And &HFF
        nG = (Pic.Point(X, Y) \ &H100) And &HFF
        nB = (Pic.Point(X, Y) \ &H10000) And &HFF
        
        nGray = 0.299 * nR + 0.587 * nG + 0.114 * nB
        PicB.PSet (X, Y), RGB(nGray, nGray, nGray)

    Next Y
Next X
End Sub

Private Sub Command3_Click()
For X = 1 To PicB.ScaleWidth - 1
        If X Mod 2 = 0 Then
            DoEvents
        End If
    For Y = 1 To PicB.ScaleHeight - 1
           
        If (PicB.Point(X, Y) And &HFF) > 120 Then
            PicC.PSet (X, Y), RGB(255, 255, 255)
        Else
            PicC.PSet (X, Y), RGB(0, 0, 0)
        End If
        
    Next Y
Next X
End Sub

Private Sub Command4_Click()
ErosionAndDilation 1, 8, 5
End Sub

Private Sub Command5_Click()

PicE.Picture = PicD.Image

Dim PicX As Long, PicY As Long, PicTemplateX As Long, PicTemplateY As Long
Dim Counter As Double, XCounter As Double 'match \ mis-match 的pixel數目
Dim Sum As Double '來源影像的總像素
Dim nR_Template, nG_Template, nB_Template As Integer
    
    Sum = PicTemplate.ScaleHeight * PicTemplate.ScaleWidth
    
    For PicY = 0 To PicD.ScaleHeight - 1 Step 3
        For PicX = 0 To PicD.ScaleWidth - 1 Step 3
            DoEvents
            Counter = 0
            XCounter = 0
            
            
                            For PicTemplateY = 0 To (PicTemplate.ScaleHeight - 1)
                                For PicTemplateX = 0 To (PicTemplate.ScaleWidth - 1)
                                    If (XCounter / Sum) >= 0.11 Then
                                        GoTo Break
                                    End If
                                    
                                    '載入原圖該pixel的顏色
                                    nR = PicD.Point(PicX + PicTemplateX, PicY + PicTemplateY) And &HFF
                                    nG = (PicD.Point(PicX + PicTemplateX, PicY + PicTemplateY) \ &H100) And &HFF
                                    nB = (PicD.Point(PicX + PicTemplateX, PicY + PicTemplateY) \ &H10000) And &HFF
                                    
                                    '載入樣板該pixel的顏色
                                    nR_Template = PicTemplate.Point(PicTemplateX, PicTemplateY) And &HFF
                                    nG_Template = (PicTemplate.Point(PicTemplateX, PicTemplateY) \ &H100) And &HFF
                                    nB_Template = (PicTemplate.Point(PicTemplateX, PicTemplateY) \ &H10000) And &HFF
                                    
                                    '該pixel對比
                                    If ((nR = nR_Template) And (nG = nG_Template) And (nB = nB_Template)) Then
                                        Counter = Counter + 1
                                    Else
                                        XCounter = XCounter + 1
                                    End If
                                    
                                Next PicTemplateX
                            Next PicTemplateY
        

        
        If (Counter / Sum) > 0.89 Then
            PicE.DrawWidth = 3 '設定畫線的粗細
            PicE.Line (PicX, PicY)-(PicX + PicTemplate.ScaleWidth, PicY + PicTemplate.ScaleHeight), vbRed, B '畫方框
            PicX = PicX + PicTemplate.ScaleWidth '跳過已經比對到的區域
            GoTo Over
        End If
        
Break:

            LabelX.Caption = "X=" & PicX
            LabelY.Caption = "Y=" & PicY
        
        Next PicX
    Next PicY
Over:
End Sub

Private Sub Command6_Click()
Dim i, TempX, TempY
'Prewitt filter
    'call T2gray 不需要
    
    For X = 1 To PicB.ScaleWidth - 1
        For Y = 1 To PicB.ScaleHeight - 1
            If X Mod 4 = 0 Then '調控轉換速度用
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
            
            PicF.PSet (X, Y), RGB(pixS, pixS, pixS)
            pixS = 0 '歸零
            TempX = 0
            TempY = 0
        Next Y
    Next X
End Sub

Private Sub Form_Load()
    LabelX.FontSize = 18
    LabelY.FontSize = 18
    
    Pic.ScaleMode = 3
    Pic.AutoRedraw = True
    Pic.Picture = LoadPicture(App.Path & "\Mhorse.bmp")
    
    PicTemplate.ScaleMode = 3
    PicTemplate.AutoRedraw = True
    PicTemplate.Picture = LoadPicture(App.Path & "\Template.bmp")
    PicTemplate.AutoSize = True
    
    
    PicA.ScaleMode = 3
    PicA.AutoRedraw = True
    PicB.ScaleMode = 3
    PicB.AutoRedraw = True
    PicC.ScaleMode = 3
    PicC.AutoRedraw = True
    PicD.ScaleMode = 3
    PicD.AutoRedraw = True
    PicE.ScaleMode = 3
    PicE.AutoRedraw = True
    PicF.ScaleMode = 3
    PicF.AutoRedraw = True
End Sub
Private Sub ErosionAndDilation(p_nTimes As Integer, p_nType As Integer, EorD As Integer)
    
    '座標,次數
    Dim X As Long, Y As Long, N As Integer
    
    '3X3矩陣
    Dim pix(8) As Byte
    
    '4方向 8方向
    Dim intFour As Byte, intEight As Byte
    
    
    For N = 1 To p_nTimes
        For X = 0 To PicC.ScaleWidth - 1 '有邊界問題所以要-1
            If X Mod 2 = 0 Then
                DoEvents
            End If
            For Y = 0 To PicC.ScaleHeight - 1
            
                pix(0) = PicC.Point(X - 1, Y - 1) And &HFF '0 => A
                pix(1) = PicC.Point(X, Y - 1) And &HFF     '1 => B
                pix(2) = PicC.Point(X + 1, Y - 1) And &HFF '2 => C
                pix(3) = PicC.Point(X - 1, Y) And &HFF     '3 => D
                pix(4) = PicC.Point(X, Y) And &HFF         '4 => E
                pix(5) = PicC.Point(X + 1, Y) And &HFF     '5 => F
                pix(6) = PicC.Point(X - 1, Y + 1) And &HFF '6 => G
                pix(7) = PicC.Point(X, Y + 1) And &HFF     '7 => H
                pix(8) = PicC.Point(X + 1, Y + 1) And &HFF '8 => I
                
                Select Case EorD '(5or4) '(5=E,4=D)
                    Case 5: '侵蝕
                        If p_nType = 8 Then '八方向
                            
                            If (pix(0) Or pix(1) Or pix(2) Or pix(3) Or pix(5) Or pix(6) Or pix(7) Or pix(8)) = 255 Then
                                intEight = 255
                            Else
                                intEight = pix(4)
                            End If
                            
                            PicD.PSet (X, Y), RGB(intEight, intEight, intEight)
                         
                        ElseIf p_nType = 4 Then '四方向
                            
                            If (pix(1) Or pix(3) Or pix(5) Or pix(7)) = 255 Then
                                intFour = 255
                            Else
                                intFour = pix(4)
                            End If
                            
                            PicD.PSet (X, Y), RGB(intFour, intFour, intFour)
                        End If
                        
                    Case 4: '增長
                        If p_nType = 8 Then '八方向
                            
                            If (pix(0) And pix(1) And pix(2) And pix(3) And pix(5) And pix(6) And pix(7) And pix(8)) = 255 Then
                                intEight = pix(4)
                            Else
                                intEight = 0
                            End If
                            
                            PicD.PSet (X, Y), RGB(intEight, intEight, intEight)
                         
                        ElseIf p_nType = 4 Then '四方向
                            
                            If (pix(1) And pix(3) And pix(5) And pix(7)) = 255 Then
                                intFour = pix(4)
                            Else
                                intFour = 0
                            End If
                            
                            PicD.PSet (X, Y), RGB(intFour, intFour, intFour)
                        End If
                End Select
                    
                    
            Next Y
        Next X
    Next N
    
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
Private Sub RoundPixelLoad()
    pix(0) = PicB.Point(X - 1, Y - 1) And &HFF '0 => A
    pix(1) = PicB.Point(X, Y - 1) And &HFF     '1 => B
    pix(2) = PicB.Point(X + 1, Y - 1) And &HFF '2 => C
    pix(3) = PicB.Point(X - 1, Y) And &HFF     '3 => D
    pix(4) = PicB.Point(X, Y) And &HFF         '4 => E
    pix(5) = PicB.Point(X + 1, Y) And &HFF     '5 => F
    pix(6) = PicB.Point(X - 1, Y + 1) And &HFF '6 => G
    pix(7) = PicB.Point(X, Y + 1) And &HFF     '7 => H
    pix(8) = PicB.Point(X + 1, Y + 1) And &HFF '8 => I
End Sub
