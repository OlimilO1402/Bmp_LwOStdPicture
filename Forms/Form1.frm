VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5220
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5805
   LinkTopic       =   "Form1"
   ScaleHeight     =   5220
   ScaleWidth      =   5805
   StartUpPosition =   3  'Windows-Standard
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
      Height          =   3975
      Left            =   120
      ScaleHeight     =   3915
      ScaleWidth      =   5355
      TabIndex        =   3
      Top             =   1080
      Width           =   5415
      Begin VB.PictureBox Picture1 
         Height          =   3735
         Left            =   0
         ScaleHeight     =   3675
         ScaleWidth      =   4755
         TabIndex        =   5
         Top             =   0
         Width           =   4815
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
    Set m_bmp = LoadPicture(m_PFN)
    Set Picture1.Picture = m_bmp
    Set m_BmpPic = New_IPicture(m_Pic, m_bmp)
    Debug.Print "Handle:     " & m_BmpPic.Handle
    Debug.Print "hPal:       " & m_BmpPic.hPal
    Debug.Print "Type:       " & m_BmpPic.Type
    Debug.Print "Width:      " & m_BmpPic.Width
    Debug.Print "Height:     " & m_BmpPic.Height
    Debug.Print "CurDC:      " & m_BmpPic.CurDC
    'Debug.Print "Attributes: " & m_BmpPic.Attributes
    'Debug.Print "KeepOriFmt: " & m_BmpPic.KeepOriginalFormat
    'Debug.Print "SetHdc: " & m_BmpPic.SetHdc
    'Debug.Print "KeepOriFmt: " & m_BmpPic.KeepOriginalFormat
    
    Set Picture1.Picture = m_BmpPic ' New_IPicture(m_Pic, m_bmp)
End Sub

'Private Sub BtnLoadIPictureDisp_Click()
'    Set Picture1.Picture = New_IPictureDisp(m_Pic, LoadPicture(m_PFN))
'End Sub

Private Sub BtnChangeBackColor_Click()
    Picture1.BackColor = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
End Sub


