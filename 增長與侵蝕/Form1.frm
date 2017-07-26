VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7935
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   9840
   LinkTopic       =   "Form1"
   ScaleHeight     =   7935
   ScaleWidth      =   9840
   StartUpPosition =   3  '系統預設值
   Begin VB.PictureBox Pic3 
      Height          =   6000
      Left            =   5040
      ScaleHeight     =   9391.305
      ScaleMode       =   0  '使用者自訂
      ScaleWidth      =   4098.461
      TabIndex        =   9
      Top             =   120
      Width           =   4500
   End
   Begin VB.OptionButton OptionOpening 
      Caption         =   "OptionOpening"
      Height          =   615
      Left            =   3720
      TabIndex        =   6
      Top             =   7080
      Width           =   2295
   End
   Begin VB.TextBox TextDirect 
      Height          =   615
      Left            =   1680
      TabIndex        =   5
      Text            =   "4"
      Top             =   7200
      Width           =   1335
   End
   Begin VB.TextBox TextN 
      Height          =   615
      Left            =   1680
      TabIndex        =   4
      Text            =   "1"
      Top             =   6360
      Width           =   1335
   End
   Begin VB.PictureBox Pic2 
      Height          =   6000
      Left            =   5040
      ScaleHeight     =   9391.305
      ScaleMode       =   0  '使用者自訂
      ScaleWidth      =   4098.461
      TabIndex        =   3
      Top             =   120
      Width           =   4500
   End
   Begin VB.PictureBox Pic1 
      Height          =   6000
      Left            =   240
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   5984.887
      ScaleMode       =   0  '使用者自訂
      ScaleWidth      =   4425.249
      TabIndex        =   2
      Top             =   120
      Width           =   4500
   End
   Begin VB.OptionButton OptionClosing 
      Caption         =   "OptionClosing"
      Height          =   735
      Left            =   3720
      TabIndex        =   1
      Top             =   6240
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1575
      Left            =   6720
      TabIndex        =   0
      Top             =   6240
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "方向"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   8
      Top             =   7200
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "次數"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   7
      Top             =   6360
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub ErosionAndDilation(p_nTimes As Integer, p_nType As Integer, EorD As Integer)
    
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
        For Y = 0 To Pic2.ScaleHeight - 1 '有邊界問題所以要-1
            If Y Mod 16 = 0 Then
                DoEvents
            End If
            For X = 0 To Pic2.ScaleWidth - 1
            
                pix(0) = Pic2.Point(X - 1, Y - 1) And &HFF '0 => A
                pix(1) = Pic2.Point(X, Y - 1) And &HFF     '1 => B
                pix(2) = Pic2.Point(X + 1, Y - 1) And &HFF '2 => C
                pix(3) = Pic2.Point(X - 1, Y) And &HFF     '3 => D
                pix(4) = Pic2.Point(X, Y) And &HFF         '4 => E
                pix(5) = Pic2.Point(X + 1, Y) And &HFF     '5 => F
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
                            
                            Pic3.PSet (X, Y), RGB(intEight, intEight, intEight)
                         
                        ElseIf p_nType = 4 Then '四方向
                            
                            If (pix(1) Or pix(3) Or pix(5) Or pix(7)) = 255 Then
                                intFour = 255
                            Else
                                intFour = pix(4)
                            End If
                            
                            Pic3.PSet (X, Y), RGB(intFour, intFour, intFour)
                        End If
                        
                    Case 4: '增長
                        If p_nType = 8 Then '八方向
                            
                            If (pix(0) And pix(1) And pix(2) And pix(3) And pix(5) And pix(6) And pix(7) And pix(8)) = 255 Then
                                intEight = pix(4)
                            Else
                                intEight = 0
                            End If
                            
                            Pic3.PSet (X, Y), RGB(intEight, intEight, intEight)
                         
                        ElseIf p_nType = 4 Then '四方向
                            
                            If (pix(1) And pix(3) And pix(5) And pix(7)) = 255 Then
                                intFour = pix(4)
                            Else
                                intFour = 0
                            End If
                            
                            Pic3.PSet (X, Y), RGB(intFour, intFour, intFour)
                        End If
                End Select
                    
                    
            Next X
        Next Y
    
    Set Pic2.Picture = Pic3.Image
    
    Next N
    
End Sub



Private Sub Command1_Click()

Dim X As Long, Y As Long
Pic2.Picture = Pic1.Image

If OptionClosing.Value = True Then '先侵蝕,後增長
    ErosionAndDilation TextN, TextDirect, 5
    ErosionAndDilation TextN, TextDirect, 4
ElseIf OptionOpening.Value = True Then '先增長,後侵蝕
    ErosionAndDilation TextN, TextDirect, 4
    ErosionAndDilation TextN, TextDirect, 5
End If

End Sub

Private Sub Form_Load()
    Pic1.ScaleMode = 3
    Pic1.AutoRedraw = True
    Pic1.AutoSize = True
    Pic2.ScaleMode = 3
    Pic2.AutoRedraw = True
    Pic2.AutoSize = True
    Pic3.ScaleMode = 3
    Pic3.AutoRedraw = True
    Pic3.AutoSize = True
    TextN.FontSize = 20
    TextDirect.FontSize = 20
    OptionClosing.FontSize = 16
    OptionOpening.FontSize = 16
    
End Sub
