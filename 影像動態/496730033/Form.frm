VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7500
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   9840
   LinkTopic       =   "Form1"
   ScaleHeight     =   500
   ScaleMode       =   3  '像素
   ScaleWidth      =   656
   StartUpPosition =   3  '系統預設值
   Begin VB.CheckBox Check3 
      Caption         =   "No Filter"
      Height          =   375
      Left            =   4920
      TabIndex        =   10
      Top             =   5400
      Width           =   2775
   End
   Begin VB.CheckBox Check2 
      Caption         =   "No show in Pic2"
      Height          =   375
      Left            =   4920
      TabIndex        =   9
      Top             =   4920
      Width           =   2775
   End
   Begin VB.CheckBox Check1 
      Caption         =   "No Back Circle"
      Height          =   375
      Left            =   4920
      TabIndex        =   8
      Top             =   4440
      Width           =   2655
   End
   Begin VB.PictureBox Picture4 
      Height          =   3615
      Left            =   0
      ScaleHeight     =   237
      ScaleMode       =   3  '像素
      ScaleWidth      =   317
      TabIndex        =   7
      Top             =   3720
      Width           =   4815
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00FFFFFF&
      Height          =   3615
      Left            =   4920
      ScaleHeight     =   237
      ScaleMode       =   3  '像素
      ScaleWidth      =   317
      TabIndex        =   4
      Top             =   0
      Width           =   4815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "stop"
      Height          =   495
      Left            =   8400
      TabIndex        =   3
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "start"
      Height          =   495
      Left            =   6960
      TabIndex        =   2
      Top             =   3720
      Width           =   1335
   End
   Begin VB.PictureBox Picture2 
      Height          =   3615
      Left            =   4920
      ScaleHeight     =   237
      ScaleMode       =   3  '像素
      ScaleWidth      =   317
      TabIndex        =   1
      Top             =   0
      Width           =   4815
   End
   Begin VB.PictureBox Picture1 
      Height          =   3615
      Left            =   0
      ScaleHeight     =   237
      ScaleMode       =   3  '像素
      ScaleWidth      =   317
      TabIndex        =   0
      Top             =   0
      Width           =   4815
   End
   Begin VB.Label Label2 
      Caption         =   "圓數：0"
      Height          =   255
      Left            =   4920
      TabIndex        =   6
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "邊點數：0"
      Height          =   255
      Left            =   4920
      TabIndex        =   5
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Menu munWebcam 
      Caption         =   "視訊(&V)"
      Begin VB.Menu munSetViedoSource 
         Caption         =   "擷取來源(&S)"
      End
      Begin VB.Menu munSetVideoFormat 
         Caption         =   "影像格式(&F)"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim StartCap As Boolean, hwdc
Dim StartSyn As Boolean
Public PictureAcc As PictureBox
Dim Data(321, 241) As Integer
Dim DataBW(321, 241) As Boolean
Dim NewPoint As Boolean
Dim BWValue As Double
Private Type EdgePoint
X As Long
Y As Long
Num As Long
End Type
Private Type CircleCount
CenterX As Long
CenterY As Long
R As Long
End Type

Dim i As Long
Dim MaxEdgePoint As Integer

Dim NewData(321, 241) As Boolean
Dim NewCircle(300, 4) As Long
Dim NewCircleCount As Integer
Dim EdgePoint(40000) As EdgePoint
Dim CircleCount(1000) As CircleCount


Private Sub Check2_Click()
    Picture3.Cls
End Sub

Private Sub Command1_click()
    StartSyn = True

    Dim X As Long, Y As Long
    Dim intR As Long, intG As Long, intB As Long
    Dim dobH As Double, dobS As Double, dobV As Double, intEight As Double
    Dim intGray As Long, intInvers As Long
    Dim PictureColor As Long, DataPoint As Long
    
    For X = 1 To 321
        For Y = 1 To 241
            DataBW(X, Y) = 0
        Next
    Next
    NewCircleCount = 0
Synchronize:
    
    SendMessage hwdc, WM_CAP_COPY, 0, 0

    Picture2.Picture = Clipboard.GetData
    SendMessage hwdc, wm_cap_set_preview, 1, 0
    Clipboard.Clear
    DoEvents
    
    For Y = 0 To Picture2.ScaleHeight - 1
        For X = 0 To Picture2.ScaleWidth - 1
            Data(X, Y) = Picture2.Point(X, Y) And &HFF
        Next
        If Y \ 10 = 0 Then DoEvents
    Next

    If Check2.Value = 0 Then
    
    
    For Y = 1 To Picture3.ScaleHeight - 2
        For X = 1 To Picture1.ScaleWidth - 2
            intEight = Abs(-Data(X - 1, Y - 1) + Data(X + 1, Y - 1) - Data(X - 1, Y) + Data(X + 1, Y - 1) - Data(X - 1, Y + 1) + Data(X + 1, Y + 1)) + Abs(-Data(X - 1, Y - 1) - Data(X, Y - 1) - Data(X + 1, Y - 1) + Data(X - 1, Y + 1) + Data(X, Y + 1) + Data(X + 1, Y + 1))
            If intEight < 60 Then
                NewData(X, Y) = 0
            Else
                NewData(X, Y) = 1
                'EdgePoint(i).X = X
                'EdgePoint(i).Y = Y
                'EdgePoint(i).Num = i
                'i = i + 1
                'MaxEdgePoint = MaxEdgePoint + 1
            End If
            'If NewData(X, Y) = DataBW(X, Y) Then
            'Else
            '    DataBW(X, Y) = NewData
            '    If NewData = 1 Then
            '        Picture3.PSet (X, Y)
            '    Else
            '        Picture3.PSet (X, Y), &HFFFFFF
            '    End If
            'End If
        Next
        'DoEvents
    Next
    If Check3.Value = 0 Then Call dilation
    For Y = 1 To Picture3.ScaleHeight - 2
        For X = 1 To Picture1.ScaleWidth - 2
            If NewData(X, Y) = True Then
                EdgePoint(i).X = X
                EdgePoint(i).Y = Y
                EdgePoint(i).Num = i
                i = i + 1
                MaxEdgePoint = MaxEdgePoint + 1
            End If
            If NewData(X, Y) = DataBW(X, Y) Then
            Else
                DataBW(X, Y) = NewData(X, Y)
                If NewData(X, Y) = True Then
                    Picture3.PSet (X, Y)
                Else
                    Picture3.PSet (X, Y), &HFFFFFF
                End If
            End If
        Next
        DoEvents
    Next
   
    
    Else

    For Y = 1 To Picture3.ScaleHeight - 2
        For X = 1 To Picture1.ScaleWidth - 2
            intEight = Abs(-Data(X - 1, Y - 1) + Data(X + 1, Y - 1) - Data(X - 1, Y) + Data(X + 1, Y - 1) - Data(X - 1, Y + 1) + Data(X + 1, Y + 1)) + Abs(-Data(X - 1, Y - 1) - Data(X, Y - 1) - Data(X + 1, Y - 1) + Data(X - 1, Y + 1) + Data(X, Y + 1) + Data(X + 1, Y + 1))
            If intEight < 60 Then
                NewData(X, Y) = 0
            Else
                NewData(X, Y) = 1
                'EdgePoint(i).X = X
                'EdgePoint(i).Y = Y
                'EdgePoint(i).Num = i
                'i = i + 1
                'MaxEdgePoint = MaxEdgePoint + 1
            End If
        Next
        'DoEvents
    Next
    If Check3.Value = 0 Then Call dilation
    For Y = 1 To Picture3.ScaleHeight - 2
        For X = 1 To Picture1.ScaleWidth - 2
            If NewData(X, Y) = True Then
                EdgePoint(i).X = X
                EdgePoint(i).Y = Y
                EdgePoint(i).Num = i
                i = i + 1
                MaxEdgePoint = MaxEdgePoint + 1
            End If
            If NewData(X, Y) = DataBW(X, Y) Then
            Else
                DataBW(X, Y) = NewData(X, Y)
                'If NewData(X, Y) = True Then
                '    Picture3.PSet (X, Y)
                'Else
                '    Picture3.PSet (X, Y), &HFFFFFF
                'End If
            End If
        Next
        DoEvents
    Next
    
    End If
    
    Call CircleDetect
    i = 0
    Label1.Caption = "邊點數：" & MaxEdgePoint
    MaxEdgePoint = 0
    
    
If StartSyn = True Then GoTo Synchronize


End Sub

Private Sub Command2_Click()
    StartSyn = False

End Sub

Private Sub Form_Load()
Picture1.AutoRedraw = True
Picture2.AutoRedraw = True
Picture1.ScaleMode = 3
Picture2.ScaleMode = 3
Dim temp As Long

  hwdc = capCreateCaptureWindow("www.planetcodes.blogspot.com", ws_child Or ws_visible, 0, 0, 320, 240, Picture1.hWnd, 0)
  If (hwdc <> 0) Then
    temp = SendMessage(hwdc, wm_cap_driver_connect, 0, 0)
    temp = SendMessage(hwdc, wm_cap_set_preview, 1, 0)
    temp = SendMessage(hwdc, WM_CAP_SET_PREVIEWRATE, 100, 0)
    StartCap = True
  Else
    MsgBox ("No Webcam found")
  End If
End Sub


Private Sub munSetVideoFormat_Click()
    SendMessage hwdc, WM_CAP_DLG_VIDEOFORMAT, 0, 0
End Sub

Private Sub munSetViedoSource_Click()
    SendMessage hwdc, WM_CAP_DLG_VIDEOSOURCE, 0, 0
End Sub

Private Sub CircleDetect()

Dim RP(4) As Integer
Dim R(4) As Double
Dim X As Integer
'隨機抓點儲存陣列

Dim RndCount As Integer
RndCount = 0
'控制迴圈次數


Dim C123X As Single, C124X   As Single, C134X   As Single, C234X   As Single
Dim C123Y As Single, C124Y   As Single, C134Y   As Single, C234Y   As Single
'圓心坐標

Dim R123 As Single, R124 As Single, R134 As Single, R234 As Single
'半徑

Dim Dist4to123 As Single, Dist3to124 As Single, Dist2to134 As Single, Dist1to234 As Single
'距離

Dim CenterX As Single, CenterY As Single, RadiusR As Single
'選定圓之圓心及半徑

Dim Square1 As Long, Square2 As Long, Square3 As Long, Square4 As Long
'與圓點之距離平方

Dim X1 As Long, X2 As Long, X3 As Long, X4 As Long
Dim Y1 As Long, Y2 As Long, Y3 As Long, Y4 As Long
'四點座標


Dim X12 As Long, X13 As Long, X14 As Long, X23 As Long, X24 As Long, X34 As Long
Dim Y12 As Long, Y13 As Long, Y14 As Long, Y23 As Long, Y24 As Long, Y34 As Long
'X12 = X2 - X1

Dim SquareDist12 As Long, SquareDist13 As Long, SquareDist14 As Long, SquareDist23 As Long, SquareDist24 As Long, SquareDist34 As Long
'SquareDist12 = X12 * X12 + Y12 * Y12

Dim Denom123 As Long, Denom124 As Long, Denom134 As Long, Denom234 As Long
'Denom234 = 2 * ((X3 - X2) * (Y4 - Y2) - (Y3 - Y2) * (X4 - X2))

Dim CountCirclePoint As Long
'投票計數器

Dim DistToCircle As Single
'第四點與候選圓之距離

Dim CountCircle As Long
CountCircle = 0
'圓數

Dim ThresholdCoCircleDist As Long
'第四點和前三點組成之圓周上的距離門檻值
ThresholdCoCircleDist = 1

Dim Threshold2PDist As Long
'三點成圓彼此之間的距離平方的門檻值
Threshold2PDist = 30

Dim ThresholdCircleRatio As Long
'找到的圓點數(4 * 根號2 * R)之多少百分比)
ThresholdCircleRatio = 99

Dim MaxX As Single, MinX As Single, MaxY As Single, MinY As Single
'可調式範圍參數

Dim i As Integer

If NewCircleCount <> 0 Then
    For X = 0 To NewCircleCount - 1
        CenterX = NewCircle(X, 2)
        CenterY = NewCircle(X, 3)
        RadiusR = NewCircle(X, 4)
        MaxX = CenterX + RadiusR + ThresholdCoCircleDist + 1
        MinX = CenterX - RadiusR - ThresholdCoCircleDist - 1
        MaxY = CenterY + RadiusR + ThresholdCoCircleDist + 1
        MinY = CenterY - RadiusR - ThresholdCoCircleDist - 1
        
        CountCirclePoint = 0
        
        For i = 0 To MaxEdgePoint - 1
        
        If ((MinX <= EdgePoint(i).X) And (EdgePoint(i).X <= MaxX)) And ((MinY <= EdgePoint(i).Y) And (EdgePoint(i).Y <= MaxY)) Then
            DistToCircle = Abs(Sqr((EdgePoint(i).X - CenterX) * (EdgePoint(i).X - CenterX) + (EdgePoint(i).Y - CenterY) * (EdgePoint(i).Y - CenterY)) - RadiusR)
            If DistToCircle <= ThresholdCoCircleDist Then
            CountCirclePoint = CountCirclePoint + 1
            End If
        End If
        
        Next i
        
        If (CountCirclePoint >= (4 * Sqr(2) * RadiusR * ThresholdCircleRatio) / 100) Then
        'If CountCirclePoint >= 100 Then
            'X = CenterX
            'Y = CenterY
            'RR = RadiusR
            
            'Pic4.Circle (X, Y), RR
            CircleCount(CountCircle).CenterX = CenterX
            CircleCount(CountCircle).CenterY = CenterY
            CircleCount(CountCircle).R = RadiusR
            
            'Picture4.Circle (CenterX, CenterY), RadiusR, &HFF
            CountCircle = CountCircle + 1
        End If

    Next
End If
For RndCount = 0 To 1000
    
    For i = 0 To 3
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
    
    Square1 = X1 * X1 + Y1 * Y1
    Square2 = X2 * X2 + Y2 * Y2
    Square3 = X3 * X3 + Y3 * Y3
    Square4 = X4 * X4 + Y4 * Y4
    
    Denom123 = 2 * (X12 * Y13 - X13 * Y12)
    Denom124 = 2 * (X12 * Y14 - X14 * Y12)
    Denom134 = 2 * (X13 * Y14 - X14 * Y13)
    Denom234 = 2 * (X23 * Y24 - X24 * Y23)

    '計算第四點到圓的距離 四組使用相同的演算法

    If (Denom123 = 0) Or (SquareDist12 <= Threshold2PDist) Or (SquareDist13 <= Threshold2PDist) Or (SquareDist23 <= Threshold2PDist) Then
        Dist4to123 = 1000
    Else
        C123X = ((Square2 - Square1) * Y13 - (Square3 - Square1) * Y12) / Denom123
        C123Y = ((Square3 - Square1) * X12 - (Square2 - Square1) * X13) / Denom123
        R123 = Sqr((X1 - C123X) * (X1 - C123X) + (Y1 - C123Y) * (Y1 - C123Y))
        '(x-a)^2 + (y-b)^2 = r^2
        Dist4to123 = Abs(Sqr((X4 - C123X) * (X4 - C123X) + (Y4 - C123Y) * (Y4 - C123Y)) - R123)
    End If
    
    If (Denom124 = 0) Or (SquareDist12 <= Threshold2PDist) Or (SquareDist14 <= Threshold2PDist) Or (SquareDist24 <= Threshold2PDist) Then
        Dist3to124 = 1000
    Else
        C124X = ((Square2 - Square1) * Y14 - (Square4 - Square1) * Y12) / Denom124
        C124Y = ((Square4 - Square1) * X12 - (Square2 - Square1) * X14) / Denom124
        R124 = Sqr((X1 - C124X) * (X1 - C124X) + (Y1 - C124Y) * (Y1 - C124Y))
        Dist3to124 = Abs(Sqr((X3 - C124X) * (X3 - C124X) + (Y3 - C124Y) * (Y3 - C124Y)) - R124)
    End If
    
    If (Denom134 = 0) Or (SquareDist13 <= Threshold2PDist) Or (SquareDist14 <= Threshold2PDist) Or (SquareDist34 <= Threshold2PDist) Then
        Dist2to134 = 1000
    Else
        C134X = ((Square3 - Square1) * Y14 - (Square4 - Square1) * Y13) / Denom134
        C134Y = ((Square4 - Square1) * X13 - (Square3 - Square1) * X14) / Denom134
        R134 = Sqr((X1 - C134X) * (X1 - C134X) + (Y1 - C134Y) * (Y1 - C134Y))
        Dist2to134 = Abs(Sqr((X2 - C134X) * (X2 - C134X) + (Y2 - C134Y) * (Y2 - C134Y)) - R134)
    End If
        
    If (Denom234 = 0) Or (SquareDist23 <= Threshold2PDist) Or (SquareDist24 <= Threshold2PDist) Or (SquareDist34 <= Threshold2PDist) Then
        Dist1to234 = 1000
    Else
        C234X = ((Square3 - Square2) * Y24 - (Square4 - Square2) * Y23) / Denom234
        C234Y = ((Square4 - Square2) * X23 - (Square3 - Square2) * X24) / Denom234
        R234 = Sqr((X2 - C234X) * (X2 - C234X) + (Y2 - C234Y) * (Y2 - C234Y))
        Dist1to234 = Abs(Sqr((X1 - C234X) * (X1 - C234X) + (Y1 - C234Y) * (Y1 - C234Y)) - R234)
    End If
        
    '判斷
    
    If (Dist4to123 = 1000) And (Dist3to124 = 1000) And (Dist2to134 = 1000) And (Dist1to234 = 1000) Then
        '全部不符合
    ElseIf (Dist4to123 <= ThresholdCoCircleDist) Or (Dist3to124 <= ThresholdCoCircleDist) Or (Dist2to134 <= ThresholdCoCircleDist) Or (Dist1to234 <= ThresholdCoCircleDist) Then
        '至少有一符合
           
        Dim MinDist As Single
        MinDist = ThresholdCoCircleDist
        '最小距離
        
        If Dist4to123 <= ThresholdCoCircleDist Then
            CenterX = C123X
            CenterY = C123Y
            RadiusR = R123
            MinDist = Dist4to123
        End If
        
        If Dist3to124 <= MinDist Then
            CenterX = C124X
            CenterY = C124Y
            RadiusR = R124
            MinDist = Dist3to124
        End If
        
        If Dist2to134 <= MinDist Then
            CenterX = C134X
            CenterY = C134Y
            RadiusR = R134
            MinDist = Dist2to134
        End If
        
        If Dist1to234 <= MinDist Then
            CenterX = C234X
            CenterY = C234Y
            RadiusR = R234
            MinDist = Dist1to234
        End If
        


        '決定候選圓
        '投票
        
        MaxX = CenterX + RadiusR + ThresholdCoCircleDist + 1
        MinX = CenterX - RadiusR - ThresholdCoCircleDist - 1
        MaxY = CenterY + RadiusR + ThresholdCoCircleDist + 1
        MinY = CenterY - RadiusR - ThresholdCoCircleDist - 1
        '可調式範圍
        
        CountCirclePoint = 0
        
        For i = 0 To MaxEdgePoint - 1
        
        If ((MinX <= EdgePoint(i).X) And (EdgePoint(i).X <= MaxX)) And ((MinY <= EdgePoint(i).Y) And (EdgePoint(i).Y <= MaxY)) Then
            DistToCircle = Abs(Sqr((EdgePoint(i).X - CenterX) * (EdgePoint(i).X - CenterX) + (EdgePoint(i).Y - CenterY) * (EdgePoint(i).Y - CenterY)) - RadiusR)
            If DistToCircle <= ThresholdCoCircleDist Then
            CountCirclePoint = CountCirclePoint + 1
            End If
        End If
        
        Next i
        
    
        '超過門檻就算圓
        If (CountCirclePoint >= (4 * Sqr(2) * RadiusR * ThresholdCircleRatio) / 100) Then
        'If CountCirclePoint >= 100 Then
            'X = CenterX
            'Y = CenterY
            'RR = RadiusR
            
            'Pic4.Circle (X, Y), RR
            CircleCount(CountCircle).CenterX = CenterX
            CircleCount(CountCircle).CenterY = CenterY
            CircleCount(CountCircle).R = RadiusR
            
            'Picture4.Circle (CenterX, CenterY), RadiusR, &HFF
            CountCircle = CountCircle + 1
        End If
    
    End If

Next RndCount
Picture4.Cls
If Check1.Value = 0 Then
    For X = 0 To CountCircle
        Picture4.Circle (CircleCount(X).CenterX, CircleCount(X).CenterY), CircleCount(X).R, RGB(196, 196, 255)
    Next
End If
    If CountCircle <> 0 Then Call CircleCheck(CountCircle)

End Sub

Private Sub CircleCheck(ByVal CountCircle)
    Dim X As Long, Y As Long
    Dim XY(300) As Long
    Dim NewCircle2(300) As Integer
    Dim Wt As Boolean
    NewCircleCount = 0
    For X = 0 To CountCircle
        NewCircle2(X) = 1
        'Picture4.Circle (CircleCount(X).CenterX, CircleCount(X).CenterY), CircleCount(X).R, &HFF
    Next
    'NewCircle(0, 1) = XY(0)
    NewCircle(0, 2) = CircleCount(0).CenterX
    NewCircle(0, 3) = CircleCount(0).CenterY
    NewCircle(0, 4) = CircleCount(0).R
    NewCircleCount = NewCircleCount + 1
    For X = 1 To CountCircle
        For Y = 0 To NewCircleCount - 1
            If (Abs(CircleCount(X).R - NewCircle(Y, 4)) > 30) Or (Abs(CircleCount(X).CenterX - NewCircle(Y, 2)) > 30) Or (Abs(CircleCount(X).CenterY - NewCircle(Y, 3)) > 30) Then
                Wt = True
                
            Else
                'NewCircle(NewCircleCount - 1, 2) = CircleCount(X).CenterX
                'NewCircle(NewCircleCount - 1, 3) = CircleCount(X).CenterY
                'NewCircle(NewCircleCount - 1, 4) = CircleCount(X).R
                'NewCircle(NewCircleCount - 1, 2) = NewCircle2(X) / (NewCircle2(X) + 1) * NewCircle(NewCircleCount - 1, 2) + CircleCount(X).CenterX / (NewCircle2(X) + 1)
                'NewCircle(NewCircleCount - 1, 3) = NewCircle2(X) / (NewCircle2(X) + 1) * NewCircle(NewCircleCount - 1, 3) + CircleCount(X).CenterY / (NewCircle2(X) + 1)
                'NewCircle(NewCircleCount - 1, 4) = NewCircle2(X) / (NewCircle2(X) + 1) * NewCircle(NewCircleCount - 1, 4) + CircleCount(X).R / (NewCircle2(X) + 1)
                
                NewCircle2(X) = NewCircle2(X) + 1
                GoTo NextRnd
            End If
        Next
    If Wt Then
        NewCircle(NewCircleCount, 2) = CircleCount(X).CenterX
        NewCircle(NewCircleCount, 3) = CircleCount(X).CenterY
        NewCircle(NewCircleCount, 4) = CircleCount(X).R
        NewCircleCount = NewCircleCount + 1
    End If
    Wt = False
NextRnd:
    Next
    
    For X = 0 To NewCircleCount - 1
        Picture4.Circle (NewCircle(X, 2), NewCircle(X, 3)), NewCircle(X, 4) + 2, &HFF
    Next
    If NewCircleCount > 0 Then
        Label2.Caption = "圓數：" & NewCircleCount - 1
    Else
        Label2.Caption = "圓數：0"
    End If

End Sub

Sub dilation()
    Dim X As Long, Y As Long
    Dim DataBW2(321, 241) As Boolean

    For Y = 1 To Picture3.ScaleHeight - 2
        For X = 1 To Picture3.ScaleWidth - 2
            If (NewData(X - 1, Y) And NewData(X + 1, Y) And NewData(X, Y - 1) And NewData(X, Y + 1)) = 0 Then
                DataBW2(X, Y) = 0
            Else
                DataBW2(X, Y) = True
            End If
        Next
    Next
    For Y = 1 To Picture3.ScaleHeight - 2
        For X = 1 To Picture3.ScaleWidth - 2
            If (DataBW2(X - 1, Y) Or DataBW2(X + 1, Y) Or DataBW2(X, Y - 1) Or DataBW2(X, Y + 1)) = True Then
                NewData(X, Y) = True
            Else
                NewData(X, Y) = 0
            End If
        Next
    Next
    'For Y = 1 To Picture3.ScaleHeight - 2
    '    For X = 1 To Picture3.ScaleWidth - 2
    '        If NewData(X, Y) = True Then Text1.Text = Text1.Text + 1
    'Next
    'Next
End Sub
