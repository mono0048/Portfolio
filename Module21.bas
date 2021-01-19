Attribute VB_Name = "Module21"
Option Explicit
Sub グラデーションモジュールかっこかり()

Dim mysheet As Worksheet

    Application.ScreenUpdating = False
    For Each mysheet In Worksheets
        If InStr(mysheet.Name, "全体") = 0 Then
            With mysheet.Range("C2").Interior
                '線形グラデーション
                .Pattern = xlPatternLinearGradient
                '色が変化する方向の角度を90度に設定
                .Gradient.Degree = 0
                '色情報をクリア
                .Gradient.ColorStops.Clear
                'グラデーションで使用する色を設定
                .Gradient.ColorStops.Add(0).Color = 16777215
                .Gradient.ColorStops.Add(1).Color = 14461583
            End With
        End If
    Next
    Application.ScreenUpdating = True

End Sub

Sub 色情報取得()

Dim myColor As ColorStop

    For Each myColor In Range("C2").Interior.Gradient.ColorStops
        MsgBox myColor.Color
    Next
    
End Sub

Sub 罫線設定()

Dim mysheet As Worksheet

    For Each mysheet In Worksheets
        If InStr(mysheet.Name, "全体") = 0 Then
            With mysheet.Range("C2")
                .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeTop).Weight = xlMedium
                .Borders(xlEdgeTop).ColorIndex = xlAutomatic
                .Borders(xlEdgeLeft).LineStyle = xlDash
                .Borders(xlEdgeLeft).Weight = xlThin
                .Borders(xlEdgeLeft).ColorIndex = xlAutomatic
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).Weight = xlThin
                .Borders(xlEdgeBottom).ColorIndex = xlAutomatic
            End With
        End If
    Next
    
End Sub
