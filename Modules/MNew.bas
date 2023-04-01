Attribute VB_Name = "MNew"
Option Explicit

Public Function AlphaPB(ForePB As PictureBox, BackPB As PictureBox) As AlphaPB
    Set AlphaPB = New AlphaPB: AlphaPB.New_ ForePB, BackPB
End Function

