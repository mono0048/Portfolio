VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet173"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

Private Sub Worksheet_selectionChange(ByVal Target As Range)
'MsgBox Target.Address(0, 0)
        If Target.Row > 8 And Target.Row < 29 And Target.Column = 3 Then
            If Target.Cells(1).Value = "" And Target.Cells(1).Offset(0, 1).Value = "" Then
                Target.Cells(1).Offset(0, 2).Font.Color = Target.Cells(1).Offset(0, 2).Interior.Color
                Target.Offset(0, 3).Font.Color = Target.Offset(0, 3).Interior.Color
                Target.Offset(0, 4).Font.Color = Target.Offset(0, 4).Interior.Color
                Target.Offset(0, 5).Font.Color = Target.Offset(0, 5).Interior.Color
             Else
                Target.Offset(0, 2).Font.Color = 1
                Target.Offset(0, 3).Font.Color = 1
                Target.Offset(0, 4).Font.Color = 1
                Target.Offset(0, 5).Font.Color = 1
             End If
         End If

        If Target.Row > 8 And Target.Row < 29 And Target.Column = 4 Then
            If Target.Cells(1).Value = "" And Target.Cells(1).Offset(0, -1).Value = "" Then
                Target.Offset(0, 1).Font.Color = Target.Offset(0, 1).Interior.Color
                Target.Offset(0, 2).Font.Color = Target.Offset(0, 2).Interior.Color
                Target.Offset(0, 3).Font.Color = Target.Offset(0, 3).Interior.Color
                Target.Offset(0, 4).Font.Color = Target.Offset(0, 4).Interior.Color
            Else
                Target.Offset(0, 1).Font.Color = 1
                Target.Offset(0, 2).Font.Color = 1
                Target.Offset(0, 3).Font.Color = 1
                Target.Offset(0, 4).Font.Color = 1
             End If
         End If


End Sub

