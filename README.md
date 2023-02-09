# Bmp_LwOStdPicture  
## Implementing IPicture for showing transparent bitmaps correctly

[![GitHub](https://img.shields.io/github/license/OlimilO1402/Bmp_LwOStdPicture?style=plastic)](https://github.com/OlimilO1402/Bmp_LwOStdPicture/blob/master/LICENSE) 
[![GitHub release (latest by date)](https://img.shields.io/github/v/release/OlimilO1402/Bmp_LwOStdPicture?style=plastic)](https://github.com/OlimilO1402/Bmp_LwOStdPicture/releases/latest)
[![Github All Releases](https://img.shields.io/github/downloads/OlimilO1402/Bmp_LwOStdPicture/total.svg)](https://github.com/OlimilO1402/Bmp_LwOStdPicture/releases/download/v1.0.0/Bmp_LwOStdPicture_v1.0.0.zip)
![GitHub followers](https://img.shields.io/github/followers/OlimilO1402?style=social)


Project started in feb. 2023.  
This example shows how to implement the IPicture-Interface for supporting alpha-channel-bitmaps through AlphaBlend-Api, by using a lightweight-object. 

this is a non working sample, it is pre beta
How to use it:
```vba
Option Explicit
Private m_PFN As String
Private m_Pic As TIPicture

Private Sub Form_Load()
    InitIPictureVTable
    m_PFN = App.Path & "\Bird.bmp"
End Sub

Private Sub BtnLoadIPicture_Click()
    Set Picture1.Picture = New_IPicture(m_Pic, LoadPicture(m_PFN))
End Sub

Private Sub BtnLoadIPictureDisp_Click()
    Set Picture1.Picture = New_IPictureDisp(m_Pic, LoadPicture(m_PFN))
End Sub

Private Sub BtnChangeBackColor_Click()
    Picture1.BackColor = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
End Sub

```

[example url link text here](https://link-url-here.org) 

![LwOStdPicture Image](Resources/LwOStdPicture.png "LwOStdPicture Image")
