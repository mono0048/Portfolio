Attribute VB_Name = "Module3"
Option Explicit

Sub ���ԏW�v()
'*******************************************************************************************************************************
'�Z�[�t�e�B�l�b�g�쐬�l��
'1. �w���V�X�e�����g�����Q��
'2.�I�[�g�t�B���^�ɂĎw����ԓ��̔R���ނ𒊏o
'3.�������݈���̂��߂ɕ\��
'4.���āA���[�h���獷�����݈��
'***����A�g�����̎����X�V���s���B�܂��I�[�g�t�B���^���蓮�Œ��J�ɁB�\���󗓂��Ȃ����v�`�F�b�N�B��SQL�ɂđΉ��A����ǍD�B
'2020/06/10�쐬
'2020/10/12�X�V(�����g�p�J�n)
'2020/10/14�X�V(SQL�Ή��B�Q�l�V�[�g�폜)�R�[�h���ώG�����邩��蒼�����̂���
'*******************************************************************************************************************************
    
    Dim MaxC1, MaxC2, cnt, yoko, i, j, k As Long
    Dim Sh1, Sh2, Sh3, Sh4 As Worksheet
    Dim Cord, Hinmei As String
    Dim Kazu As Double
    Dim d1, d2, d3 As Date
    
    '---���[�N�V�[�g��錾---
    Set Sh1 = Worksheets("�L��")
    Set Sh2 = Worksheets("SN�����҈ꗗ")
    Set Sh3 = Worksheets("DB")
'    Set Sh4 = Worksheets("�Q�l")
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Call DB�V�[�g����
    
     '---�Q�l�V�[�g������---
    MaxC1 = Sh3.Range("A65536").End(xlUp).Row '-----�Q�l�V�[�g�̍ŏI�s
    MaxC2 = Sh2.Range("A65536").End(xlUp).Row '-----SN�����҈ꗗ�V�[�g�̍ŏI�s
    cnt = Sh2.Range("XFD2").End(xlToLeft).Column '-----SN�����҈ꗗ�V�[�g�̍ŏI��
    
    '---�W�v�J�n���ƏW�v�I�������w��---
    If Sh1.Range("C5").Value = "1" Then
        d1 = "2020/04/01"
        d2 = "2020/06/30"
'        Sh3.Range("A1").AutoFilter Field:=3, Criteria1:=">=2020/04/01", Operator:=xlAnd, Criteria2:="<=2020/06/30"
'        Sh3.Range("A1").AutoFilter Field:=6, Criteria1:=1
'        Sh3.Range("A1").AutoFilter Field:=2, Criteria1:=11
    '�`�F�b�N�}�[�N����
        Sh2.Range(Sh2.Range("Q3"), Sh2.Cells(MaxC2, 20)).Value = ""
        Sh2.Range(Sh2.Range("Q3"), Sh2.Cells(MaxC2, 17)).Value = ChrW(10003)
    ElseIf Sh1.Range("C5").Value = "2" Then
        d1 = "2020/07/01"
        d2 = "2020/09/30"
'        Sh3.Range("A1").AutoFilter Field:=3, Criteria1:=">=2020/07/01", Operator:=xlAnd, Criteria2:="<=2020/09/30"
'        Sh3.Range("A1").AutoFilter Field:=6, Criteria1:=1
'        Sh3.Range("A1").AutoFilter Field:=2, Criteria1:=11
    '�`�F�b�N�}�[�N����
        Sh2.Range(Sh2.Range("Q3"), Sh2.Cells(MaxC2, 20)).Value = ""
        Sh2.Range(Sh2.Range("R3"), Sh2.Cells(MaxC2, 18)).Value = ChrW(10003)
    ElseIf Sh1.Range("C5").Value = "3" Then
        d1 = "2020/10/01"
        d2 = "2020/12/31"
'        Sh3.Range("A1").AutoFilter Field:=3, Criteria1:=">=2020/10/01", Operator:=xlAnd, Criteria2:="<=2020/12/31"
'        Sh3.Range("A1").AutoFilter Field:=6, Criteria1:=1
'        Sh3.Range("A1").AutoFilter Field:=2, Criteria1:=11
    '�`�F�b�N�}�[�N����
        Sh2.Range(Sh2.Range("Q3"), Sh2.Cells(MaxC2, 20)).Value = ""
        Sh2.Range(Sh2.Range("S3"), Sh2.Cells(MaxC2, 19)).Value = ChrW(10003)
    ElseIf Sh1.Range("C5").Value = "4" Then
        d1 = "2021/01/01"
        d2 = "2021/03/31"
'        Sh3.Range("A1").AutoFilter Field:=3, Criteria1:=">=2021/01/01", Operator:=xlAnd, Criteria2:="<=2021/03/31"
'        Sh3.Range("A1").AutoFilter Field:=6, Criteria1:=1
'        Sh3.Range("A1").AutoFilter Field:=2, Criteria1:=11
    '�`�F�b�N�}�[�N����
        Sh2.Range(Sh2.Range("Q3"), Sh2.Cells(MaxC2, 20)).Value = ""
        Sh2.Range(Sh2.Range("T3"), Sh2.Cells(MaxC2, 20)).Value = ChrW(10003)
    Else: MsgBox "�����w�肪�Ԉ���Ă��܂��B"
    End If
'
'    '-----�I�[�g�t�B���^��f�[�^���R�s�y
'    Sh3.Range("A1").CurrentRegion.Copy Sh4.Range("A1")
'    '-----�Q�l�V�[�g�̃R�[�h�𕶎���ɕϊ�����
'    Sh4.Range("D:D").NumberFormatLocal = "@"
    Sh3.Range("F:F").NumberFormatLocal = "@" '��SQL���ߕ�����FORMAT�����ɂ���
     '---�Q�l�V�[�g������---
    MaxC1 = Sh3.Range("A65536").End(xlUp).Row '-----�Q�l�V�[�g�̍ŏI�s
    MaxC2 = Sh2.Range("A65536").End(xlUp).Row '-----SN�����҈ꗗ�V�[�g�̍ŏI�s
    cnt = Sh2.Range("XFD2").End(xlToLeft).Column '-----SN�����҈ꗗ�V�[�g�̍ŏI��
    
    For k = 3 To MaxC2
        Cord = Sh2.Range("C" & k).Text
    '---�ƍ������������Ńf�[�^�x�[�X������---
        For j = 0 To cnt - 25
            Kazu = 0
            Hinmei = Sh2.Range("W2").Offset(0, j + 1)
            For i = 2 To MaxC1
                d3 = Sh3.Range("B" & i).Value '-----�Q�l�V�[�g�̓��t��
                If d3 >= d1 And d3 <= d2 Then
                    If Sh3.Range("C" & i).Text = Cord Then
                        If Sh3.Range("F" & i).Text = Hinmei Then
                            Kazu = Kazu + Sh3.Range("H" & i).Value
                        End If
                    End If
                End If
            Next
    '---SN�����҈ꗗ�V�[�g�̕\�ɏo��---
            Sh2.Range("W" & k).Offset(0, j + 1).Value = Kazu
        Next
    Next
    '-----�I�[�g�t�B���^�[����
'    Sh3.Range("A1").AutoFilter
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

Sub DB�V�[�g����()

Dim �R�l�N�V���� As New ADODB.Connection
Dim ���R�[�h As New ADODB.Recordset
Dim �R�}���h As New ADODB.Command
Dim d1, d2 As Date
Dim Sh1, Sh4 As Worksheet
Set Sh1 = Worksheets("�L��")
Set Sh4 = Worksheets("DB")

�R�l�N�V����.Open ConnectionString:= _
"Provider=Microsoft.ACE.OLEDB.12.0;" & _
"Data Source=\\192.168.10.250\ackobai\kobai.mdb;"

'���ԓK�p�͈�
If Sh1.Range("C5").Value = "1" Then
        d1 = "2020/04/01"
        d2 = "2020/06/30"
ElseIf Sh1.Range("C5").Value = "2" Then
        d1 = "2020/07/01"
        d2 = "2020/09/30"
ElseIf Sh1.Range("C5").Value = "3" Then
        d1 = "2020/10/01"
        d2 = "2020/12/31"
ElseIf Sh1.Range("C5").Value = "4" Then
        d1 = "2021/01/01"
        d2 = "2021/03/31"
End If

With �R�}���h
    .ActiveConnection = �R�l�N�V����
    .CommandText = _
    "SELECT ��ԋ敪,���@�@�t,FORMAT(�R�[�h,'@'),�R�[�h���@����,���ރR�[�h,���i�R�[�h,���i���@����,���@�@�� FROM KT01 WHERE ��ԋ敪 = 0 and ���@�@�t BETWEEN " & "#" & d1 & "#" & " And " & "#" & d2 & "#" & " and ���i�R�[�h=1001;"
    Set ���R�[�h = .Execute
End With
Sh4.Cells.Delete
Sh4.Range("A1").Value = "��ԋ敪"
Sh4.Range("B1").Value = "���@�@�t"
Sh4.Range("C1").Value = "�R�[�h"
Sh4.Range("D1").Value = "�R�[�h���@����"
Sh4.Range("E1").Value = "���ރR�[�h"
Sh4.Range("F1").Value = "���i�R�[�h"
Sh4.Range("G1").Value = "���i���@����"
Sh4.Range("H1").Value = "���@�@��"

Sh4.Range("A2").CopyFromRecordset Data:=���R�[�h
���R�[�h.Close: Set ���R�[�h = Nothing
�R�l�N�V����.Close: Set �R�l�N�V���� = Nothing

End Sub
