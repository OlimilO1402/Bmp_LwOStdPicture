VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5220
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9765
   LinkTopic       =   "Form1"
   ScaleHeight     =   5220
   ScaleWidth      =   9765
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   375
      Left            =   7080
      TabIndex        =   9
      Top             =   480
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   7080
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   1080
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   4440
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   1080
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   2400
      TabIndex        =   6
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   3720
      TabIndex        =   4
      Top             =   600
      Width           =   1815
   End
   Begin VB.PictureBox Picture2 
      Height          =   3735
      Left            =   120
      ScaleHeight     =   3675
      ScaleWidth      =   4155
      TabIndex        =   3
      Top             =   1080
      Width           =   4215
      Begin VB.PictureBox Picture1 
         Height          =   3495
         Left            =   0
         ScaleHeight     =   3435
         ScaleWidth      =   3915
         TabIndex        =   5
         Top             =   0
         Width           =   3975
      End
   End
   Begin VB.CommandButton BtnLoadIPictureDisp 
      Caption         =   "Load IPictureDisp"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1935
   End
   Begin VB.CommandButton BtnChangeBackColor 
      Caption         =   "Change BackColor"
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
   Begin VB.CommandButton BtnLoadIPicture 
      Caption         =   "Load IPicture"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_PFN As String
Private m_bmp As IPicture
Private m_Pic As TIPicture
Private m_BmpPic As IPicture
Private m_apb As AlphaPB

Private Sub Command1_Click()
    Dim sp As StdPicture
    Set sp = LoadPicture(m_PFN)
    Set Picture2.Picture = sp
    'Picture1.BackColor = vbBlue 'White
    m_apb.UpdateView
End Sub

Private Sub Form_Load()
    InitIPictureVTable
    m_PFN = App.Path & "\Bird.bmp"
    Set m_apb = MNew.AlphaPB(Picture1, Picture2)
    m_apb.BlendFunc = &H1FF0000
End Sub

Private Sub BtnLoadIPicture_Click()
Try: On Error GoTo Catch
    
    Set m_bmp = LoadPicture(m_PFN)
    Text1.Text = IPicture_ToStr(m_bmp)
    Set Picture1.Picture = m_bmp
Catch:
    '
End Sub

Private Sub Command3_Click()
Try: On Error GoTo Catch
    
    Set m_BmpPic = New_IPicture(m_Pic, m_bmp)
    Dim s As String
    s = IPicture_ToStr(m_BmpPic)
    'Set Picture1.Picture = m_BmpPic ' New_IPicture(m_Pic, m_bmp)
Catch:
    Text2.Text = s
End Sub

Private Function IPicture_ToStr(bmp As IPicture) As String
Try: On Error GoTo Catch
    Dim s As String
    s = s & "Handle: " & bmp.Handle & vbCrLf
    s = s & "hPal:   " & bmp.hPal & vbCrLf
    s = s & "Type:   " & bmp.Type & vbCrLf
    s = s & "Width:  " & bmp.Width & vbCrLf
    s = s & "Height: " & bmp.Height & vbCrLf
    s = s & "CurDC:  " & bmp.CurDC & vbCrLf
Catch:
    IPicture_ToStr = s
End Function

'Private Sub BtnLoadIPictureDisp_Click()
'    Set Picture1.Picture = New_IPictureDisp(m_Pic, LoadPicture(m_PFN))
'End Sub

Private Sub BtnChangeBackColor_Click()
    Picture1.BackColor = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
End Sub


