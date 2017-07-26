VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3600
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   11025
   LinkTopic       =   "Form1"
   ScaleHeight     =   240
   ScaleMode       =   3  '像素
   ScaleWidth      =   735
   StartUpPosition =   3  '系統預設值
   Begin VB.PictureBox Picture5 
      Height          =   1815
      Left            =   2880
      ScaleHeight     =   1755
      ScaleWidth      =   2595
      TabIndex        =   9
      Top             =   3480
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.PictureBox Picture4 
      Height          =   1815
      Left            =   5160
      ScaleHeight     =   1755
      ScaleWidth      =   2475
      TabIndex        =   8
      Top             =   0
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      Caption         =   "tool"
      Height          =   3495
      Left            =   7920
      TabIndex        =   5
      Top             =   0
      Width           =   1575
      Begin VB.CheckBox circle1 
         Caption         =   "Circle"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   2760
         Width           =   975
      End
      Begin VB.CheckBox time1 
         Caption         =   "Time"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   3120
         Width           =   975
      End
      Begin VB.OptionButton Sharpen 
         Caption         =   "Sharpen"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   2040
         Width           =   975
      End
      Begin VB.OptionButton sobel 
         Caption         =   "Sobel"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1320
         Width           =   975
      End
      Begin VB.OptionButton prewitt 
         Caption         =   "Prewitt"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   1095
      End
      Begin VB.OptionButton binary 
         Caption         =   "Binary"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   975
      End
      Begin VB.OptionButton gray 
         Caption         =   "Gray"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton EdgeDetection 
         Caption         =   "EdgeDetection"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1680
         Width           =   1335
      End
      Begin VB.OptionButton Motion 
         Caption         =   "Motion"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   2400
         Width           =   975
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Stop"
      Height          =   495
      Left            =   9720
      TabIndex        =   4
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   495
      Left            =   9720
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   2760
      Top             =   2640
   End
   Begin VB.PictureBox Picture3 
      Height          =   1815
      Left            =   2640
      ScaleHeight     =   1755
      ScaleWidth      =   2355
      TabIndex        =   2
      Top             =   0
      Width           =   2415
      Begin VB.Label Label1 
         BackStyle       =   0  '透明
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   0
         Width           =   1815
      End
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      Height          =   1815
      Left            =   120
      ScaleHeight     =   117
      ScaleMode       =   3  '像素
      ScaleWidth      =   157
      TabIndex        =   1
      Top             =   3480
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.PictureBox Picture1 
      Height          =   1815
      Left            =   120
      ScaleHeight     =   117
      ScaleMode       =   3  '像素
      ScaleWidth      =   157
      TabIndex        =   0
      Top             =   0
      Width           =   2415
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
Dim StartCap As Boolean, hwdc
Dim StartSyn As Boolean
Public PictureAcc As PictureBox
Dim t As Integer
Dim maskA As Integer, maskB As Integer, maskC As Integer, maskD As Integer, maskE As Integer, maskF As Integer, maskG As Integer, maskH As Integer, maskI As Integer, int1 As Integer, int2 As Integer
Dim intS1 As Integer, intS2 As Integer

Option Explicit

Private Type EdgePoint

    X As Long
    Y As Long
    Num As Long
    
End Type

Dim i As Long
Dim MaxEdgePoint As Integer

Dim EdgePoint(10000) As EdgePoint
Private Sub Command1_click()
    
  Timer1.Enabled = True
  Timer1.Interval = 500
'If StartSyn = True Then GoTo Synchronize


End Sub

Private Sub Command2_Click()
    StartSyn = False
    Timer1.Enabled = False
    Picture3.Cls
End Sub

Private Sub Form_Load()
Picture1.AutoRedraw = True
Picture2.AutoRedraw = True
Picture3.AutoRedraw = True
Picture4.AutoRedraw = True
Picture5.AutoRedraw = True

Picture1.ScaleMode = 3
Picture2.ScaleMode = 3
Picture3.ScaleMode = 3
Picture4.ScaleMode = 3
Picture5.ScaleMode = 3


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
Private Sub mask_prewitt1()
maskA = -1
maskB = 0
maskC = 1
maskD = -1
maskE = 0
maskF = 1
maskG = -1
maskH = 0
maskI = 1
End Sub
Private Sub mask_prewitt2()

maskA = -1
maskB = -1
maskC = -1
maskD = 0
maskE = 0
maskF = 0
maskG = 1
maskH = 1
maskI = 1

End Sub
Private Sub Timer1_Timer()

  
  StartSyn = True
    
    Dim X As Long, Y As Long
    Dim intR As Integer, intG As Integer, intB As Integer
    Dim dobH As Double, dobS As Double, dobV As Double
    Dim intGray As Integer, intInvers As Integer
    Dim PictureColor As Long
    Dim intA As Integer, intBB As Integer, intC As Integer
    Dim intD As Integer, intE As Integer, intF As Integer
    Dim intGG As Integer, intH As Integer, intI As Integer
    Dim intS As Integer
    Dim intE1  As Integer
    Dim intMono As Integer
    
'Synchronize:


    SendMessage hwdc, WM_CAP_COPY, 0, 0

    Picture2.Picture = Clipboard.GetData
    SendMessage hwdc, wm_cap_set_preview, 1, 0
    Clipboard.Clear
    
    
    
    If EdgeDetection.Value = True Then
    
     For Y = 0 To Picture1.ScaleHeight - 1
        For X = 0 To Picture1.ScaleWidth - 1
        
        intE = Picture2.Point(X, Y) And &HFF
        
        If intE = intE1 Then
        intS = intE
        Else

        intA = Picture2.Point(X - 1, Y - 1) And &HFF
        intBB = Picture2.Point(X, Y - 1) And &HFF
        intC = Picture2.Point(X + 1, Y - 1) And &HFF
        intD = Picture2.Point(X - 1, Y) And &HFF
        intF = Picture2.Point(X + 1, Y) And &HFF
        intGG = Picture2.Point(X - 1, Y + 1) And &HFF
        intH = Picture2.Point(X, Y + 1) And &HFF
        intI = Picture2.Point(X + 1, Y + 1) And &HFF

        
        Call mask_prewitt1
        intS = Abs(intA * maskA + intBB * maskB + intC * maskC + intD * maskD + intE * maskE + intF * maskF + intGG * maskG + intH * maskH + intI * maskI)
        End If
        
        Picture3.PSet (X, Y), RGB(intS, intS, intS)
           
    Next X
    Next Y
    
    End If


    If Motion.Value = True Then
    
    For Y = 0 To Picture1.ScaleHeight - 1
    For X = 0 To Picture1.ScaleWidth - 1
        
    intE = Picture2.Point(X, Y) And &HFF
    intE1 = Picture4.Point(X, Y) And &HFF
    
    If (intE > intE1 - 20) And (intE < intE1 + 20) Then
    intS = 0
    Else
    intS = intE
    End If
    
    Picture2.PSet (X, Y), RGB(intS, intS, intS)

    Next X
    Next Y
    
    
    
    For Y = 0 To Picture1.ScaleHeight - 1
        For X = 0 To Picture1.ScaleWidth - 1
        
        
        intA = Picture2.Point(X - 1, Y - 1) And &HFF
        intBB = Picture2.Point(X, Y - 1) And &HFF
        intC = Picture2.Point(X + 1, Y - 1) And &HFF
        intD = Picture2.Point(X - 1, Y) And &HFF
        intE = Picture2.Point(X, Y) And &HFF
        intF = Picture2.Point(X + 1, Y) And &HFF
        intGG = Picture2.Point(X - 1, Y + 1) And &HFF
        intH = Picture2.Point(X, Y + 1) And &HFF
        intI = Picture2.Point(X + 1, Y + 1) And &HFF

        
        Call mask_prewitt1
        intS = Abs(intA * maskA + intBB * maskB + intC * maskC + intD * maskD + intE * maskE + intF * maskF + intGG * maskG + intH * maskH + intI * maskI)
        
        Picture3.PSet (X, Y), RGB(intS, intS, intS)
           
    Next X
    Next Y
    
    Set Picture4 = Picture2
    
    End If

If gray.Value = True Then
    
    
    For Y = 0 To Picture1.ScaleHeight - 1
    For X = 0 To Picture1.ScaleWidth - 1
    
    intR = Picture2.Point(X, Y) And &HFF
    intG = (Picture2.Point(X, Y) \ &H100) And &HFF
    intB = (Picture2.Point(X, Y) \ &H10000) And &HFF
    intGray = 0.299 * intR + 0.587 * intG + 0.114 * intB
    Picture3.PSet (X, Y), RGB(intGray, intGray, intGray)

  Next X
Next Y
    
    
    End If

If binary.Value = True Then

    For Y = 0 To Picture1.ScaleHeight - 1
    For X = 0 To Picture1.ScaleWidth - 1

    intR = Picture2.Point(X, Y) And &HFF
    intG = (Picture2.Point(X, Y) \ &H100) And &HFF
    intB = (Picture2.Point(X, Y) \ &H10000) And &HFF
    intGray = 0.299 * intR + 0.587 * intG + 0.114 * intB

        If intGray >= 127 Then
            intMono = 255
        Else
            intMono = 0
        End If

    Picture3.PSet (X, Y), RGB(intMono, intMono, intMono)

  Next X
Next Y

End If

If prewitt.Value = True Then

For Y = 0 To Picture1.ScaleHeight - 1
  For X = 0 To Picture1.ScaleWidth - 1
  

    intA = Picture2.Point(X - 1, Y - 1) And &HFF
    int1 = Picture2.Point(X, Y - 1) And &HFF
    intC = Picture2.Point(X + 1, Y - 1) And &HFF
    intD = Picture2.Point(X - 1, Y) And &HFF
    intE = Picture2.Point(X, Y) And &HFF
    intF = Picture2.Point(X + 1, Y) And &HFF
    int2 = Picture2.Point(X - 1, Y + 1) And &HFF
    intH = Picture2.Point(X, Y + 1) And &HFF
    intI = Picture2.Point(X + 1, Y + 1) And &HFF
   
    Call mask_prewitt1
    intS1 = Abs(intA * maskA + int1 * maskB + intC * maskC + intD * maskD + intE * maskE + intF * maskF + int2 * maskG + intH * maskH + intI * maskI)
    
    Call mask_prewitt2
    intS2 = Abs(intA * maskA + int1 * maskB + intC * maskC + intD * maskD + intE * maskE + intF * maskF + int2 * maskG + intH * maskH + intI * maskI)
    
    intS = intS1 + intS2
    Picture3.PSet (X, Y), RGB(intS, intS, intS)
    
      Next X
  Next Y
  
  End If
  
If sobel.Value = True Then
  
  For Y = 0 To Picture1.ScaleHeight - 1
  For X = 0 To Picture1.ScaleWidth - 1
  
    intA = Picture2.Point(X - 1, Y - 1) And &HFF
    int1 = Picture2.Point(X, Y - 1) And &HFF
    intC = Picture2.Point(X + 1, Y - 1) And &HFF
    intD = Picture2.Point(X - 1, Y) And &HFF
    intE = Picture2.Point(X, Y) And &HFF
    intF = Picture2.Point(X + 1, Y) And &HFF
    int2 = Picture2.Point(X - 1, Y + 1) And &HFF
    intH = Picture2.Point(X, Y + 1) And &HFF
    intI = Picture2.Point(X + 1, Y + 1) And &HFF
   
    Call mask_sobel1
    intS1 = Abs(intA * maskA + int1 * maskB + intC * maskC + intD * maskD + intE * maskE + intF * maskF + int2 * maskG + intH * maskH + intI * maskI)
    
    Call mask_sobel2
    intS2 = Abs(intA * maskA + int1 * maskB + intC * maskC + intD * maskD + intE * maskE + intF * maskF + int2 * maskG + intH * maskH + intI * maskI)
    
    intS = intS1 + intS2
    Picture3.PSet (X, Y), RGB(intS, intS, intS)

  Next X
  Next Y
  
  End If

If Sharpen.Value = True Then


  For Y = 0 To Picture1.ScaleHeight - 1
  For X = 0 To Picture1.ScaleWidth - 1
  
    
    intA = Picture2.Point(X - 1, Y - 1) And &HFF
    int1 = Picture2.Point(X, Y - 1) And &HFF
    intC = Picture2.Point(X + 1, Y - 1) And &HFF
    intD = Picture2.Point(X - 1, Y) And &HFF
    intE = Picture2.Point(X, Y) And &HFF
    intF = Picture2.Point(X + 1, Y) And &HFF
    int2 = Picture2.Point(X - 1, Y + 1) And &HFF
    intH = Picture2.Point(X, Y + 1) And &HFF
    intI = Picture2.Point(X + 1, Y + 1) And &HFF
   
    maskA = 1
    maskB = -2
    maskC = 1
    maskD = -2
    maskE = 5
    maskF = -2
    maskG = 1
    maskH = -2
    maskI = 1
    
    intS = Abs(intA * maskA + int1 * maskB + intC * maskC + intD * maskD + intE * maskE + intF * maskF + int2 * maskG + intH * maskH + intI * maskI)
    
    Picture3.PSet (X, Y), RGB(intS, intS, intS)
    
  Next X
  Next Y
    
End If


If time1.Value Then
Label1.Visible = True
Label1.Caption = Date + Time
Else
Label1.Visible = False
End If

If circle1.Value Then

    Call CircleDetect

End If


  

End Sub

Private Sub mask_sobel1()

maskA = -1
maskB = 0
maskC = 1
maskD = -2
maskE = 0
maskF = 2
maskG = -1
maskH = 0
maskI = 1

End Sub
Private Sub mask_sobel2()

maskA = -1
maskB = -2
maskC = -1
maskD = 0
maskE = 0
maskF = 0
maskG = 1
maskH = 2
maskI = 1

End Sub
Private Sub CircleDetect()

Dim RP(4) As Integer
Dim R(4) As Double
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
ThresholdCircleRatio = 80

Dim MaxX As Single, MinX As Single, MaxY As Single, MinY As Single
'可調式範圍參數

Dim i As Integer

For RndCount = 0 To 5000
    
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
            Picture4.Circle (CenterX, CenterY), RadiusR
            CountCircle = CountCircle + 1
        End If
    
    End If

Next RndCount

        
End Sub
Private Sub EdgePointCount(ByVal intS, ByVal X, ByVal Y)

    If intS >= 255 Then
    
        intS = 255
        Picture3.PSet (X, Y), RGB(0, 0, 0)
        EdgePoint(i).X = X
        EdgePoint(i).Y = Y
        EdgePoint(i).Num = i
        i = i + 1
        MaxEdgePoint = MaxEdgePoint + 1
    
    Else
    
        Picture3.PSet (X, Y), RGB(255, 255, 255)
        
    End If
    
    Label1.Caption = "邊點數:" & MaxEdgePoint
    
End Sub
