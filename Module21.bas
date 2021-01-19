Attribute VB_Name = "Module21"
Option Explicit
Sub �O���f�[�V�������W���[������������()

Dim mysheet As Worksheet

    Application.ScreenUpdating = False
    For Each mysheet In Worksheets
        If InStr(mysheet.Name, "�S��") = 0 Then
            With mysheet.Range("C2").Interior
                '���`�O���f�[�V����
                .Pattern = xlPatternLinearGradient
                '�F���ω���������̊p�x��90�x�ɐݒ�
                .Gradient.Degree = 0
                '�F�����N���A
                .Gradient.ColorStops.Clear
                '�O���f�[�V�����Ŏg�p����F��ݒ�
                .Gradient.ColorStops.Add(0).Color = 16777215
                .Gradient.ColorStops.Add(1).Color = 14461583
            End With
        End If
    Next
    Application.ScreenUpdating = True

End Sub

Sub �F���擾()

Dim myColor As ColorStop

    For Each myColor In Range("C2").Interior.Gradient.ColorStops
        MsgBox myColor.Color
    Next
    
End Sub

Sub �r���ݒ�()

Dim mysheet As Worksheet

    For Each mysheet In Worksheets
        If InStr(mysheet.Name, "�S��") = 0 Then
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
