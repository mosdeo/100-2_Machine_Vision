VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6765
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13575
   LinkTopic       =   "Form1"
   ScaleHeight     =   6765
   ScaleWidth      =   13575
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command3 
      Caption         =   "找圓"
      Height          =   1095
      Left            =   9240
      TabIndex        =   10
      Top             =   4920
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "測邊緣"
      Height          =   1095
      Left            =   5880
      TabIndex        =   9
      Top             =   4920
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "轉灰階"
      Height          =   1095
      Left            =   2520
      TabIndex        =   8
      Top             =   4920
      Width           =   1935
   End
   Begin VB.PictureBox Pic3 
      Height          =   3855
      Left            =   7080
      ScaleHeight     =   3795
      ScaleWidth      =   2715
      TabIndex        =   6
      Top             =   360
      Width           =   2775
   End
   Begin VB.Frame Frame4 
      Caption         =   "圓"
      Height          =   4215
      Left            =   10320
      TabIndex        =   5
      Top             =   120
      Width           =   3015
      Begin VB.PictureBox Pic4 
         Height          =   3855
         Left            =   120
         ScaleHeight     =   3795
         ScaleWidth      =   2715
         TabIndex        =   7
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Canny"
      Height          =   4215
      Left            =   6960
      TabIndex        =   4
      Top             =   120
      Width           =   3015
   End
   Begin VB.PictureBox Pic2 
      Height          =   3855
      Left            =   3720
      ScaleHeight     =   3795
      ScaleWidth      =   2715
      TabIndex        =   2
      Top             =   360
      Width           =   2775
   End
   Begin VB.PictureBox Pic1 
      Height          =   3855
      Left            =   360
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   3795
      ScaleWidth      =   2715
      TabIndex        =   0
      Top             =   360
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      Caption         =   "原始"
      Height          =   4215
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   3015
   End
   Begin VB.Frame Frame2 
      Caption         =   "灰階"
      Height          =   4215
      Left            =   3600
      TabIndex        =   3
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Dim mask(8)   '3x3像素-濾波器變數
Dim X As Long, Y As Long '主像素座標
Dim pix(8) '3x3像素矩陣
Dim pixS As Long  '運算結果像素

Private Type EdgePoint
    X As Long
    Y As Long
    Num As Long
End Type

Dim i As Long
Dim MaxEdgePoint(5000) As EdgePoint


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
Dim i, TempX, TempY
'Prewitt filter
    
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

Private Sub Command3_Click()
Dim ThresholdCoCirc1eDist, RndCount

'隨機抓點儲存陣列
Dim RP(4) As Integer
Dim R(4) As Double

'控制迴圈次數
'Dim RndCount As Integer
RndCount = O

'圓心座標
Dim C123X As Single, C124X As Single, C134X As Single, C234X As Single
Dim C123Y As Single, C124Y As Single, C134Y As Single, C234Y As Single

'半徑
Dim R123 As Single, R124 As Single, R134 As Single, R234 As Single

'距離
Dim Dist4to123 As Single, Dist3to124 As Single, Dist2to134 As Single, Dist1to234 As Single

'選定圓之圓心與半徑
Dim CenterX As Single, CenterY As Single, RadiusR As Single

'與圓點之距離與平方
Dim Square1 As Long, Square2 As Long, Square3 As Long, Square4 As Long

'四點座標
Dim X1 As Long, X2 As Long, X3 As Long, X4 As Long
Dim Y1 As Long, Y2 As Long, Y3 As Long, Y4 As Long

'X12 = X2 - X1
Dim Xl2 As Long, X13 As Long, X14 As Long, X23 As Long, X24 As Long, X34 As Long
Dim Y12 As Long, Y13 As Long, Y14 As Long, Y23 As Long, Y24 As Long, Y34 As Long

'SquareDist12 = X12 * X12 + Y12 * Y12
Dim SquareDist12 As Long, SquareDist13 As Long, SquareDist14 As Long, SquareDist23 As Long, SquareDist24 As Long, SquareDist34 As Long

'Denom234 =2 * ((X3 - X2) “ (Y4 -Y2) - (Y3 -Y2) * (X4 - X2))
Dim Denom123 As Long, Den0m124 As Long, Denom134 As Long, Denom234 As Long

'投票計數器
Dim CountCirclePoint As Long

'第四點與候選圓之距離
Dim DistToCircle As Single

'圓數
Dim CountCircle As Long
CountCircle = 0

'第四點和前三點組成之圓周上的距離門檻值
Dim Thresho1dCoCirc1eDist As Long
ThresholdCoCircleDist = 1

'三點成圓彼此之間的平方門檻值
Dim Threshold2PDist As Long
Threshold2PDist = 30

'找到的點數(4*根號2*R)之多少百分比
Dim Thresho1dCircleRatio As Long
Thresho1dCircleRatio = 80

'可調式參數範圍
Dim MaxX As Single, MinX As Single, MaxY As Single, MinY As Single


Dim i As Integer

For RndCount = 0 To 5000

    For i = O To 3
        R(i) = Rnd
        RP(i) = Int(R(i) * MaxEdgePoint)
    Next i
    
    X1 = EdgePoint(RP(0)).X
    Y1 = EdgePoint(RP(0)).Y
    X2 = EdgePoint(RP(1)).X
    Y2 = EdgePoint(RP(1)).Y
    X3 = EdgePoint(RP(2)).X
    Y3 = EdgePoint(RP(2)).Y
    X4 = EdgePoint(RP(3)).X
    Y4 = EdgePoint(RP(3)).Y
    X12 = X2 - X1
    X13 = X3 - X1
    X14 = X4 - X1
    X23 = X3 - X2
    X24 = X4 - X2
    X34 = X4 - X3
    Y12 = Y2 - Y1
    Y13 = Y3 - Y1
    Y14 = Y4 - Y1
    Y23 = Y3 - Y2
    Y24 = Y4 - Y2
    Y34 = Y4 - Y3
    SquareDist12 = X12 * X12 + Y12 * Y12
    SquareDist13 = X13 * X13 + Y13 * Y13
    SquareDist14 = X14 * X14 + Y14 * Y14
    SquareDist23 = X23 * X23 + Y23 * Y23
    SquareDist24 = X24 * X24 + Y24 * Y24
    SquareDist34 = X34 * X34 + Y34 * Y34

    '計算第四點到圓的距離,四組使用相同演算法
    Square1 = Xl * X1 + Y1 * Y1
    Square2 = X2 * X2 + Y2 * Y2
    Square3 = X3 * X3 + Y3 * Y3
    Square4 = X4 * X4 + Y4 * Y4
    Denom123 = 2 * (X12 * YI3 - X13 * Yl2)
    Denom124 = 2 * (X12 * YI4 - X14 * Yl2)
    Denoml34 = 2 * (X13 * Y14 - X14 * Y13)
    Denom234 = 2 * (X23 * Y24 - X24 * Y23)



    If (Denom123 = 0) Or (SquareDistl2 <= Threshold2PDist) Or (SquareDistl3 <= Threshold2PDist) Or (SquareDist23 <= Threshold2PDist) Then
        Dist4to123 = 1000
    Else
        C123X = ((Square2 - Squarel) * Y13 - (Square3 - Squarel) * Y12) / Denom123
        Cl23Y = ((Square3 - Squarel) * X12 - (Square2 - Squarel) * XI3) / Denom123
        R123 = Sqr((X1 - C123X) - (X1 - C123X) + (Y1 - Cl23Y) * (Y1 - Cl23Y))
        '(x-a)"2 + (y-b)"2 = 1'\2
        Dist4to123 = Abs(Sqr((X4 - Cl23X) * (X4 - Cl23X) + (Y4 - Cl23Y) * (Y4 - Cl23Y)) - R123)
    End If
    
    If (Denoml24 = 0) Or (SquareDistl2 <= Threshold2PDist) Or (SquareDist14 <= Threshold2PDist) Or (SquareDist24 <= Threshold2PDist) Then
        Dist3to124 = 1000
    Else
        Cl24X = ((Square2 - Square1) * YI4 - (Square4 - Square!) * Y12) / Denom124
        C124Y = ((Square4 - Square1) * X12 - (Square2 - Square1) * X14) / Denom124
        R124 = Sqr((Xl - Cl24X) * (X1 - C124X) + (Y1 - C124Y) * (Y1 - C124Y))
        Dis3to124 = Abs(Sqr((X3 - Cl24X) * (X3 - C124X) + (Y3 - C3124Y) * (Y3 - C124Y)) - R124)
    End If

    If (Denom134 = 0) Or (SquareDist13 <= Thxcsh0ld2PDist) Or (SquareDist14 <= Thresho1d2PDist) Or (SquareDist34 <= Threshold2PDist) Then
        Dist2to134 = 1000
    Else
        C134X = ((Square3 - Square1) * Y14 - (Square4 - Square1) * Y13) / Denom134
        C134Y = ((Square4 - Square1) * X13 - (Square3 - Square1) * X14) / Denom134
        R134 = Sqr((X1 - C134X) * (X1 - C134X) + (Y1 - C134Y) * (Y1 - C134Y))
        Dist2to134 = Abs(Sqr((X2 - C134X) * (X2 - Cl34X) + (Y2 - C134Y) * (Y2 - C134Y)) - R134)
    End If
    
    If (Denom234 = 0) Or (SquareDist23 <= Thresho1d2PDist) Or (SquareDist24 <= Threshold2PDist) Or (SquareDist34 <= Tb_resho1d2PDist) Then
        Dist1to234 = 1000
    Else
        C234X = ((Square3 - Square2) * Y24 - (Square4 - Square2) * Y23) / Denom234
        C234Y = ((Square4 - Square2) * X23 - (Square3 - Squa.re2) * X24) / Denom234
        R234 = Sqr((X2 - C234X) * (X2 - C234X) + (Y2 - C234Y) * (Y2 - C234Y))
        Dist1to234 = Abs(Sqr((X1 - C234X) * (X1 - C234X) + (Y1 - C234Y) * (Y1 - C234Y)) - R234)
    End If
    
    '判斷
    If (Dist4to123 = 1000) And (Dist3to124 = 1000) And (Dist2to134 = 1000) And (Dist1to234 = 1000) Then
    '全部不符合
    ElseIf (Dist4to123 <= ThresholdCoCirc1eDist) Or (Dist3to124 <= ThresholdCoCirc1eDist) Or (Dist2to134 <= ThresholdCoCircleDist) Or (DiSt2to234 <= ThreSho1dCoCircleDiSt) Then
    '至少有一符合
        Dim MinDist As Sing1e
        MinDist = ThresholdCoCircleDist
        '最小距離
        If Dist4tol23 <= ThresholdCoCircleDist Then
            CenterX -C123X
            CenterY = C123Y
            RadiusR = R123
            MinDist = Dist4tol23
        End If
        If Dist3tol24 <= MinDist Then
            CenterX = C124X
            CenterY -C124Y
            RadiusR = R124
            MinDist = Dist3to124
        End If
        If Dist2tol34 <= MinDist Then
            CenterX -C134X
            CenterY = C134Y
            RadiusR = R134
            MinDist = Dist2tol34
        End If
        If Distlto234 <= MinDist Then
            CenterX = C234X
            CenterY = C234Y
            RadiusR = R234
            MinDist = Distlto234
        End If

        '決定候選圓
        '投票
        MaxX = CenterX + RadiusR + ThresholdCoCircleDist + i
        MinX = CenterX - RadiusR - ThresholdCoCircleDist - i
        MaxY = CenterY + RadiusR + ThresholdCoCircleDist + i
        MrnY = CenterY - RadiusR - ThresholdCoCircleDist - i
        CountCirclePoint = 0
        For i = 0 To MaxEdgePoint - 1
            If ((MinX <= EdgePoint(i).X) And (EdgePoint(i).X <= MaxX)) And ((MinY <= EdgePoint(i).Y) And (EdgePoint(i).Y <= MaxY)) Then
                DistToCircle -Abs(Sqr((EdgePoint(i).X - CenterX) * (EdgePoint(i).X - CenterX) + (EdgePoint(i).Y - CenterY) * (EdgePoint(i).Y - CenterY)) - RadiusR)
                If DistToCircle <= ThresholdCoCircleDist Then
                      CountCirclePoint = CountCirclePoint + i
                End If
            End If
        Next i
        
        '超過門檻就算圓
        If (CountCirclePoint >= (4 * Sqr(2) * RadiusR * ThresholdCircleRatio) / 100) Then
        'lfCountCirclePoint>=100 Then
            'X=CenterX
            'Y=CenterY
            'RR=RadiusR
            
            'Pic4.Circle(X, Y), RR
             Pic4.Circle (CenterX, CenterY), RadiusR
             CountCircle = CountCircle + i
        End If
End Sub

Private Sub Form_Load()
'計算單位為像素
Pic1.ScaleMode = 3
Pic2.ScaleMode = 3
Pic3.ScaleMode = 3
Pic4.ScaleMode = 3

'設定自動重繪
Pic1.AutoRedraw = True
Pic2.AutoRedraw = True
Pic3.AutoRedraw = True
Pic4.AutoRedraw = True

Pic1.AutoSize = True
Pic2.AutoSize = True
Pic3.AutoSize = True
Pic4.AutoSize = True
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

