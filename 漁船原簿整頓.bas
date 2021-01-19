Attribute VB_Name = "漁船原簿整頓"
Option Explicit

Sub 漁船原簿シート名取得()
'**********
'2020/10/08
'森
'懸念事項
'どっかの項目が未記入の場合、全体ページが崩れるかも
'各ページに影響は無い
'抽出したくないシートがある場合、どうしようか
'**********
Dim mysheet As Worksheet
    Application.ScreenUpdating = False
    On Error Resume Next
    With ThisWorkbook
        For Each mysheet In .Worksheets
            If InStr(mysheet.Name, "一覧表") = 0 Then
                mysheet.Name = mysheet.Range("C3").Text
            End If
        Next
    End With
    Call シート目次
    Application.ScreenUpdating = True
End Sub

Sub シート目次()
  
Dim h, i, j, k As Integer
Dim iRow As Integer
Dim iColumn As Integer
Dim sName, Kumimei As String
Dim Senmei, Kumiaiinmei, Chiku, Tonsu, Nagasa, Bariki, Kennin As String
Dim Sh1 As Worksheet
Dim LastRow As Long
    
    Set Sh1 = Worksheets("一覧表")
    Sh1.Range("A1").Select
    '目次シートの設定内容をクリア
    LastRow = Sh1.Cells(Rows.Count, 2).End(xlUp).Row
    Sh1.Range(Sh1.Cells(4, 2), Sh1.Cells(LastRow, 9)).ClearContents
    Sh1.Range(Sh1.Cells(4, 2), Sh1.Cells(LastRow, 9)).Hyperlinks.Delete
    '目次開始行数（本例は4行目から目次が作られる）
    iRow = 4
    '目次を作成する列数（本例は2列目（B列）に目次が作られる）
    iColumn = 2
    On Error Resume Next
    With ThisWorkbook
        'ワークシートの数分、下記処理を繰り返す
        For i = 2 To Worksheets.Count Step 1
            '非表示となっているワークシートは目次作成対象外とする
            If .Worksheets(i).Visible = xlSheetVisible Then
                '組合員別所有船ファイルへのハイパーリンクを設置
                For h = 15 To 3 Step -1
                    If Not .Worksheets(i).Cells(6, h).Value = "" Then
                        Kumimei = "'" & .Worksheets(i).Cells(6, h).Text & "'"
                        .Worksheets(i).Hyperlinks.Add Anchor:=.Worksheets(i).Cells(6, h), Address:="\\組合員別所有船.xlsm", SubAddress:=Kumimei & "!A1", TextToDisplay:=.Worksheets(i).Cells(6, h).Text, ScreenTip:=.Worksheets(i).Cells(6, h).Text
'                        GoTo EndLoop0
                    End If
                Next
EndLoop0:
                'ワークシート名を格納
                sName = "'" & .Worksheets(i).Name & "'"
                '目次シートの対象セルにハイパーリンクを設定（目次作成対象ワークシートのA1セルへのリンク
                Sh1.Hyperlinks.Add Anchor:=Sh1.Cells(iRow, iColumn), Address:="", SubAddress:=sName & "!A1", TextToDisplay:=.Worksheets(i).Cells(3, 3).Text, ScreenTip:=.Worksheets(i).Cells(3, 3).Text
                '目次シートの対象セルにシート名を設定
                Sh1.Cells(iRow, iColumn) = .Worksheets(i).Cells(3, 3).Text
                '全体シートに住所、生年月日、地区名、出資金を転記
                '船名判定
                For j = 15 To 3 Step -1
                    If Not .Worksheets(i).Cells(5, j).Value = "" Then
                        Senmei = .Worksheets(i).Cells(5, j).Value
                        Sh1.Cells(iRow, iColumn + 1) = Senmei
                        GoTo EndLoop1
                    End If
                    'GoTo EndLoop
                Next
EndLoop1:
                '組合員名判定
                For j = 15 To 3 Step -1
                    If Not .Worksheets(i).Cells(6, j).Value = "" Then
                        Kumiaiinmei = .Worksheets(i).Cells(6, j).Value
                        Sh1.Cells(iRow, iColumn + 2) = Kumiaiinmei
                        GoTo EndLoop2
                    End If
                    'GoTo EndLoop
                Next
EndLoop2:
                '地区
                Sh1.Cells(iRow, iColumn + 3) = .Worksheets(i).Cells(3, 4).Value
                'トン数判定
                For j = 15 To 3 Step -1
                    If Not .Worksheets(i).Cells(14, j).Value = "" Then
                        Tonsu = .Worksheets(i).Cells(14, j).Value
                            Sh1.Cells(iRow, iColumn + 4) = Tonsu
                            GoTo EndLoop3
                    End If
                    'GoTo EndLoop
                Next
EndLoop3:
                '船体の長さ*幅*深さ判定
                For j = 15 To 3 Step -1
                    If Not .Worksheets(i).Cells(15, j).Value = "" Then
                        Nagasa = .Worksheets(i).Cells(15, j).Value
                        Sh1.Cells(iRow, iColumn + 5) = Nagasa
                        GoTo EndLoop4
                    End If
                    'GoTo EndLoop
                Next
EndLoop4:
                '馬力数判定
                For j = 15 To 3 Step -1
                    If Not .Worksheets(i).Cells(22, j).Value = "" Then
                        Bariki = .Worksheets(i).Cells(22, j).Value
                        Sh1.Cells(iRow, iColumn + 6) = Bariki
                        GoTo EndLoop5
                    End If
                    'GoTo EndLoop
                Next
EndLoop5:
                'Sh1.Cells(iRow, iColumn + 7) = Worksheets(i).Cells(14, 3).Value  '次回検認期日
                '次回検認期日判定(********もっといい方法あるだろう***)
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
                'リンクの文字の大きさ、フォントを設定
                Sh1.Cells(iRow, iColumn).Font.Size = 12
                Sh1.Cells(iRow, iColumn).Font.Name = "ＭＳ 明朝"
                Sh1.Cells(iRow, iColumn).Font.Bold = True
                '次の行へ
                iRow = iRow + 1
            End If
        Next i
        Sh1.Range("B1").Value = "一覧表"
        Sh1.Range("B3").Value = "漁船登録番号"
        Sh1.Range("C3").Value = "船名"
        Sh1.Range("D3").Value = "組合員名"
        Sh1.Range("E3").Value = "地区"
        Sh1.Range("F3").Value = "トン数"
        Sh1.Range("G3").Value = "船体の長さ*幅*深さ"
        Sh1.Range("H3").Value = "馬力数"
        Sh1.Range("I3").Value = "次回検認期日"
    End With
    
End Sub
