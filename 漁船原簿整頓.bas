Attribute VB_Name = "���D���됮��"
Option Explicit

Sub ���D����V�[�g���擾()
'**********
'2020/10/08
'�X
'���O����
'�ǂ����̍��ڂ����L���̏ꍇ�A�S�̃y�[�W������邩��
'�e�y�[�W�ɉe���͖���
'���o�������Ȃ��V�[�g������ꍇ�A�ǂ����悤��
'**********
Dim mysheet As Worksheet
    Application.ScreenUpdating = False
    On Error Resume Next
    With ThisWorkbook
        For Each mysheet In .Worksheets
            If InStr(mysheet.Name, "�ꗗ�\") = 0 Then
                mysheet.Name = mysheet.Range("C3").Text
            End If
        Next
    End With
    Call �V�[�g�ڎ�
    Application.ScreenUpdating = True
End Sub

Sub �V�[�g�ڎ�()
  
Dim h, i, j, k As Integer
Dim iRow As Integer
Dim iColumn As Integer
Dim sName, Kumimei As String
Dim Senmei, Kumiaiinmei, Chiku, Tonsu, Nagasa, Bariki, Kennin As String
Dim Sh1 As Worksheet
Dim LastRow As Long
    
    Set Sh1 = Worksheets("�ꗗ�\")
    Sh1.Range("A1").Select
    '�ڎ��V�[�g�̐ݒ���e���N���A
    LastRow = Sh1.Cells(Rows.Count, 2).End(xlUp).Row
    Sh1.Range(Sh1.Cells(4, 2), Sh1.Cells(LastRow, 9)).ClearContents
    Sh1.Range(Sh1.Cells(4, 2), Sh1.Cells(LastRow, 9)).Hyperlinks.Delete
    '�ڎ��J�n�s���i�{���4�s�ڂ���ڎ��������j
    iRow = 4
    '�ڎ����쐬����񐔁i�{���2��ځiB��j�ɖڎ��������j
    iColumn = 2
    On Error Resume Next
    With ThisWorkbook
        '���[�N�V�[�g�̐����A���L�������J��Ԃ�
        For i = 2 To Worksheets.Count Step 1
            '��\���ƂȂ��Ă��郏�[�N�V�[�g�͖ڎ��쐬�ΏۊO�Ƃ���
            If .Worksheets(i).Visible = xlSheetVisible Then
                '�g�����ʏ��L�D�t�@�C���ւ̃n�C�p�[�����N��ݒu
                For h = 15 To 3 Step -1
                    If Not .Worksheets(i).Cells(6, h).Value = "" Then
                        Kumimei = "'" & .Worksheets(i).Cells(6, h).Text & "'"
                        .Worksheets(i).Hyperlinks.Add Anchor:=.Worksheets(i).Cells(6, h), Address:="\\192.168.10.252\share\�f�[�^�ۊ�(���R�j\�e�X�g\�g�����ʏ��L�D.xlsm", SubAddress:=Kumimei & "!A1", TextToDisplay:=.Worksheets(i).Cells(6, h).Text, ScreenTip:=.Worksheets(i).Cells(6, h).Text
'                        GoTo EndLoop0
                    End If
                Next
EndLoop0:
                '���[�N�V�[�g�����i�[
                sName = "'" & .Worksheets(i).Name & "'"
                '�ڎ��V�[�g�̑ΏۃZ���Ƀn�C�p�[�����N��ݒ�i�ڎ��쐬�Ώۃ��[�N�V�[�g��A1�Z���ւ̃����N
                Sh1.Hyperlinks.Add Anchor:=Sh1.Cells(iRow, iColumn), Address:="", SubAddress:=sName & "!A1", TextToDisplay:=.Worksheets(i).Cells(3, 3).Text, ScreenTip:=.Worksheets(i).Cells(3, 3).Text
                '�ڎ��V�[�g�̑ΏۃZ���ɃV�[�g����ݒ�
                Sh1.Cells(iRow, iColumn) = .Worksheets(i).Cells(3, 3).Text
                '�S�̃V�[�g�ɏZ���A���N�����A�n�於�A�o������]�L
                '�D������
                For j = 15 To 3 Step -1
                    If Not .Worksheets(i).Cells(5, j).Value = "" Then
                        Senmei = .Worksheets(i).Cells(5, j).Value
                        Sh1.Cells(iRow, iColumn + 1) = Senmei
                        GoTo EndLoop1
                    End If
                    'GoTo EndLoop
                Next
EndLoop1:
                '�g����������
                For j = 15 To 3 Step -1
                    If Not .Worksheets(i).Cells(6, j).Value = "" Then
                        Kumiaiinmei = .Worksheets(i).Cells(6, j).Value
                        Sh1.Cells(iRow, iColumn + 2) = Kumiaiinmei
                        GoTo EndLoop2
                    End If
                    'GoTo EndLoop
                Next
EndLoop2:
                '�n��
                Sh1.Cells(iRow, iColumn + 3) = .Worksheets(i).Cells(3, 4).Value
                '�g��������
                For j = 15 To 3 Step -1
                    If Not .Worksheets(i).Cells(14, j).Value = "" Then
                        Tonsu = .Worksheets(i).Cells(14, j).Value
                            Sh1.Cells(iRow, iColumn + 4) = Tonsu
                            GoTo EndLoop3
                    End If
                    'GoTo EndLoop
                Next
EndLoop3:
                '�D�̂̒���*��*�[������
                For j = 15 To 3 Step -1
                    If Not .Worksheets(i).Cells(15, j).Value = "" Then
                        Nagasa = .Worksheets(i).Cells(15, j).Value
                        Sh1.Cells(iRow, iColumn + 5) = Nagasa
                        GoTo EndLoop4
                    End If
                    'GoTo EndLoop
                Next
EndLoop4:
                '�n�͐�����
                For j = 15 To 3 Step -1
                    If Not .Worksheets(i).Cells(22, j).Value = "" Then
                        Bariki = .Worksheets(i).Cells(22, j).Value
                        Sh1.Cells(iRow, iColumn + 6) = Bariki
                        GoTo EndLoop5
                    End If
                    'GoTo EndLoop
                Next
EndLoop5:
                'Sh1.Cells(iRow, iColumn + 7) = Worksheets(i).Cells(14, 3).Value  '���񌟔F����
                '���񌟔F��������(********�����Ƃ������@���邾�낤***)
                For j = 7 To 1 Step -1
                    For k = 36 To 31 Step -1
                        If Not .Worksheets(i).Cells(k, 2 * j + 1).Value = "" Then
                            Kennin = .Worksheets(i).Cells(k, 2 * j + 1).Value
                            Sh1.Cells(iRow, iColumn + 7) = Kennin
                            GoTo EndLoop6
                        End If
                        'GoTo EndLoop
                    Next
                Next
EndLoop6:
                '�����N�̕����̑傫���A�t�H���g��ݒ�
                Sh1.Cells(iRow, iColumn).Font.Size = 12
                Sh1.Cells(iRow, iColumn).Font.Name = "�l�r ����"
                Sh1.Cells(iRow, iColumn).Font.Bold = True
                '���̍s��
                iRow = iRow + 1
            End If
        Next i
        Sh1.Range("B1").Value = "�ꗗ�\"
        Sh1.Range("B3").Value = "���D�o�^�ԍ�"
        Sh1.Range("C3").Value = "�D��"
        Sh1.Range("D3").Value = "�g������"
        Sh1.Range("E3").Value = "�n��"
        Sh1.Range("F3").Value = "�g����"
        Sh1.Range("G3").Value = "�D�̂̒���*��*�[��"
        Sh1.Range("H3").Value = "�n�͐�"
        Sh1.Range("I3").Value = "���񌟔F����"
    End With
    
End Sub
