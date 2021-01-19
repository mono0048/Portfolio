Attribute VB_Name = "Module2"
Option Explicit

Sub リンク削除()

Dim i As Long
Dim 色 As Long
Application.ScreenUpdating = False
For i = 2 To Worksheets.Count Step 1
    Worksheets(i).Cells(2, 3).ClearHyperlinks
    色 = Worksheets(i).Cells(3, 3).Interior.Gradient.ColorStop
    Worksheets(i).Cells(2, 3).Interior.Gradient.ColorStop = 色
Next
Application.ScreenUpdating = True
End Sub
