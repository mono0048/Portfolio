Attribute VB_Name = "Module1"
Option Explicit

Sub �g��������䒠�V�[�g���擾()
'**********
'2020/08/26
'�X
'���O����
'�o���������L���̏ꍇ�A�S�̃y�[�W�������
'�e�y�[�W�ɉe���͖���
'���o�������Ȃ��V�[�g������ꍇ�A�ǂ����悤��
'**********
Dim mysheet As Worksheet
Application.ScreenUpdating = False
    On Error Resume Next
    For Each mysheet In Worksheets
        If InStr(mysheet.Name, "�S��") = 0 Then
            mysheet.Name = mysheet.Range("C2").Value
        End If
    Next
    
    Call �V�[�g�ڎ�
Application.ScreenUpdating = True
End Sub

Sub �V�[�g�ڎ�()
  
    Dim i As Integer
    Dim iRow As Integer
    Dim iColumn As Integer
    Dim sName, Kumimei As String
    Dim Sh1 As Worksheet
    Set Sh1 = Worksheets("�S��")
    
    Sh1.Range("A1").Select
    '�ڎ��V�[�g�̐ݒ���e���N���A
    Sh1.Range(Sh1.Cells(3, 3), Sh1.Cells(Rows.Count, 7).End(xlUp)).ClearContents
    Sh1.Range(Sh1.Cells(3, 3), Sh1.Cells(Rows.Count, 7).End(xlUp)).Hyperlinks.Delete
      
    '�ڎ��J�n�s���i�{���3�s�ڂ���ڎ��������j
    iRow = 3
    '�ڎ����쐬����񐔁i�{���3��ځiC��j�ɖڎ��������j
    iColumn = 3
  
    '���[�N�V�[�g�̐����A���L�������J��Ԃ�
    For i = 2 To Worksheets.Count Step 1
        '��\���ƂȂ��Ă��郏�[�N�V�[�g�͖ڎ��쐬�ΏۊO�Ƃ���
        If Worksheets(i).Visible = xlSheetVisible Then
            '�g�����ʏ��L�D�t�@�C���ւ̃n�C�p�[�����N��ݒu
            Kumimei = "'" & Worksheets(i).Cells(2, 3).Text & "'"
            Worksheets(i).Hyperlinks.Add Anchor:=Worksheets(i).Cells(4, 10), Address:="\\192.168.10.252\share\�f�[�^�ۊ�(���R�j\�e�X�g\�g�����ʏ��L�D.xlsm", SubAddress:=Kumimei & "!A1", TextToDisplay:="���D�o�^", ScreenTip:=Worksheets(i).Cells(4, 10).Text
            '���[�N�V�[�g�����i�[
            sName = "'" & Worksheets(i).Name & "'"
            '�ڎ��V�[�g�̑ΏۃZ���Ƀn�C�p�[�����N��ݒ�i�ڎ��쐬�Ώۃ��[�N�V�[�g��A1�Z���ւ̃����N
            Sh1.Hyperlinks.Add Anchor:=Sh1.Cells(iRow, iColumn), Address:="", SubAddress:=sName & "!A1", ScreenTip:=Worksheets(i).Name
            '�ڎ��V�[�g�̑ΏۃZ���ɃV�[�g����ݒ�
            Sh1.Cells(iRow, iColumn) = Worksheets(i).Name
            '�S�̃V�[�g�ɏZ���A���N�����A�n�於�A�o������]�L
            Sh1.Cells(iRow, iColumn + 1) = Worksheets(i).Cells(2, 5).Value
            Sh1.Cells(iRow, iColumn + 2) = Worksheets(i).Cells(2, 9).Value
            Sh1.Cells(iRow, iColumn + 3) = Worksheets(i).Cells(3, 9).Value
            Sh1.Cells(iRow, iColumn + 4) = Worksheets(i).Cells(28, 8).Value
              '���i���ǉ�
            If Not Worksheets(i).Cells(6, 7).Value = "" Then
                Sh1.Cells(iRow, iColumn + 5) = Worksheets(i).Cells(6, 7).Value
            Else
                If Not Worksheets(i).Cells(6, 6).Value = "" Then
                    Sh1.Cells(iRow, iColumn + 5) = Worksheets(i).Cells(6, 6).Value
                Else
                    If Not Worksheets(i).Cells(6, 5).Value = "" Then
                        Sh1.Cells(iRow, iColumn + 5) = Worksheets(i).Cells(6, 5).Value
                    Else
                        If Not Worksheets(i).Cells(4, 7).Value = "" Then
                            Sh1.Cells(iRow, iColumn + 5) = Worksheets(i).Cells(4, 7).Value
                        Else
                            If Not Worksheets(i).Cells(4, 6).Value = "" Then
                                Sh1.Cells(iRow, iColumn + 5) = Worksheets(i).Cells(4, 6).Value
                            Else
                                If Not Worksheets(i).Cells(4, 5).Value = "" Then
                                    Sh1.Cells(iRow, iColumn + 5) = Worksheets(i).Cells(4, 5).Value
                                Else
                                    Sh1.Cells(iRow, iColumn + 5) = Worksheets(i).Cells(4, 3).Value
                                End If
                            End If
                        End If
                    End If
                End If
            End If

  
           '�����N�̕����̑傫���A�t�H���g��ݒ�
            Sh1.Cells(iRow, iColumn).Font.Size = 12
            Sh1.Cells(iRow, iColumn).Font.Name = "�l�r ����"
            Sh1.Cells(iRow, iColumn).Font.Bold = True
  
            '���̍s��
            iRow = iRow + 1
        End If
    Next i
            Sh1.Range("B2").Value = "�ԍ�"
            Sh1.Range("C2").Value = "�g������"
            Sh1.Range("D2").Value = "�Z��"
            Sh1.Range("E2").Value = "���N����"
            Sh1.Range("F2").Value = "�n�於"
            Sh1.Range("G2").Value = "�o����"
End Sub
