Attribute VB_Name = "Module2"
Option Explicit


Sub ƒŠƒ“ƒNíœ()

Dim i As Long
Dim F As Long
Application.ScreenUpdating = False
For i = 2 To Worksheets.Count Step 1
    Worksheets(i).Cells(2, 3).ClearHyperlinks
    F = Worksheets(i).Cells(3, 3).Interior.Gradient.ColorStop
    Worksheets(i).Cells(2, 3).Interior.Gradient.ColorStop = F
Next
Application.ScreenUpdating = True
End Sub
