VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "494702123 �L����  �X��ϧηj�M�{��"
   ClientHeight    =   8520
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8925
   LinkTopic       =   "Form1"
   ScaleHeight     =   8520
   ScaleWidth      =   8925
   StartUpPosition =   3  '�t�ιw�]��
   Begin VB.CommandButton Command2 
      Caption         =   "Stop !"
      Height          =   495
      Left            =   5640
      TabIndex        =   10
      Top             =   6480
      Width           =   1695
   End
   Begin MSComDlg.CommonDialog CommDialog 
      Left            =   360
      Top             =   4920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�󴫼˪O"
      Height          =   495
      Left            =   5640
      TabIndex        =   9
      Top             =   7440
      Width           =   1695
   End
   Begin VB.CommandButton Command 
      Caption         =   "Mapping !"
      Height          =   495
      Left            =   5640
      TabIndex        =   3
      Top             =   5640
      Width           =   1695
   End
   Begin VB.Frame Frame2 
      Caption         =   "Template"
      Height          =   2415
      Left            =   240
      TabIndex        =   1
      Top             =   5640
      Width           =   4095
      Begin VB.TextBox Text1 
         Alignment       =   2  '�m�����
         Height          =   264
         Left            =   3240
         TabIndex        =   4
         Text            =   "0"
         Top             =   1560
         Width           =   492
      End
      Begin VB.PictureBox PicTemplate 
         AutoSize        =   -1  'True
         Height          =   1170
         Left            =   240
         Picture         =   "Form1.frx":0000
         ScaleHeight     =   74
         ScaleMode       =   3  '����
         ScaleWidth      =   76
         TabIndex        =   2
         Top             =   360
         Width           =   1200
      End
      Begin VB.Label Label1 
         Caption         =   "Number"
         Height          =   375
         Left            =   3120
         TabIndex        =   5
         Top             =   960
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Picture"
      Height          =   4335
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   8415
      Begin VB.PictureBox Pic 
         AutoSize        =   -1  'True
         Height          =   3735
         Left            =   240
         Picture         =   "Form1.frx":422A
         ScaleHeight     =   245
         ScaleMode       =   3  '����
         ScaleWidth      =   526
         TabIndex        =   6
         Top             =   360
         Width           =   7950
      End
   End
   Begin VB.Label LabelY 
      Caption         =   "Label2"
      Height          =   255
      Left            =   6840
      TabIndex        =   8
      Top             =   5040
      Width           =   1335
   End
   Begin VB.Label LabelX 
      Caption         =   "Label2"
      Height          =   255
      Left            =   4920
      TabIndex        =   7
      Top             =   5040
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim PicX, PicY, PicTemplateX, PicTemplateY As Integer

Private Sub Command1_Click()

With CommDialog
    .InitDir = App.Path '��ܪ��_�l���|
    .ShowOpen '���X�ɮ׿�ܵ���
End With

PicTemplate.Picture = LoadPicture(CommDialog.FileName) '���J�˪��Ϥ�

End Sub

Private Sub Command2_Click()
PicX = Pic.ScaleWidth
PicY = Pic.ScaleHeight
End Sub

Private Sub Form_Load()
    Pic.ScaleMode = 3
    Pic.AutoRedraw = True
    Pic.AutoSize = True
    PicTemplate.ScaleMode = 3
    PicTemplate.AutoRedraw = True
    PicTemplate.AutoSize = True
End Sub

Private Sub Command_Click()
    Dim TempNum As Integer '���쪺�˪O�ƥ�
    Dim Counter As Double, XCounter As Double 'match \ mis-match ��pixel�ƥ�
    Dim Sum As Double '�ӷ��v�����`����
    Dim nR, nG, nB, nR_Template, nG_Template, nB_Template As Integer
    
    Sum = PicTemplate.ScaleHeight * PicTemplate.ScaleWidth
    TempNum = 0
    
    
    For PicY = 0 To Pic.ScaleHeight - 1 Step 3
        For PicX = 0 To Pic.ScaleWidth - 1 Step 3
            DoEvents
            Counter = 0
            XCounter = 0
            
            
                            For PicTemplateY = 0 To (PicTemplate.ScaleHeight - 1)
                                For PicTemplateX = 0 To (PicTemplate.ScaleWidth - 1)
                                    If (XCounter / Sum) >= 0.11 Then
                                        GoTo Break
                                    End If
                                    
                                    '���J��ϸ�pixel���C��
                                    nR = Pic.Point(PicX + PicTemplateX, PicY + PicTemplateY) And &HFF
                                    nG = (Pic.Point(PicX + PicTemplateX, PicY + PicTemplateY) \ &H100) And &HFF
                                    nB = (Pic.Point(PicX + PicTemplateX, PicY + PicTemplateY) \ &H10000) And &HFF
                                    
                                    '���J�˪O��pixel���C��
                                    nR_Template = PicTemplate.Point(PicTemplateX, PicTemplateY) And &HFF
                                    nG_Template = (PicTemplate.Point(PicTemplateX, PicTemplateY) \ &H100) And &HFF
                                    nB_Template = (PicTemplate.Point(PicTemplateX, PicTemplateY) \ &H10000) And &HFF
                                    
                                    '��pixel���
                                    If ((nR = nR_Template) And (nG = nG_Template) And (nB = nB_Template)) Then
                                        Counter = Counter + 1
                                    Else
                                        XCounter = XCounter + 1
                                    End If
                                    
                                Next PicTemplateX
                            Next PicTemplateY
        

        
        If (Counter / Sum) > 0.89 Then
            Pic.DrawWidth = 3 '�]�w�e�u���ʲ�
            Pic.Line (PicX, PicY)-(PicX + PicTemplate.ScaleWidth, PicY + PicTemplate.ScaleHeight), vbRed, B '�e���
            TempNum = TempNum + 1
            Text1.Text = TempNum
            PicX = PicX + PicTemplate.ScaleWidth '���L�w�g���쪺�ϰ�
            
        End If
        
Break:

            LabelX.Caption = "X=" & PicX
            LabelY.Caption = "Y=" & PicY
        
        Next PicX
    Next PicY

End Sub

