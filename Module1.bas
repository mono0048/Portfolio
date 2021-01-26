Attribute VB_Name = "Module1"
Option Explicit

Sub 組合員名簿台帳シート名取得()
'**********
'2020/08/26
'森
'懸念事項
'出資金が未記入の場合、全体ページが崩れる
'各ページに影響は無い
'抽出したくないシートがある場合、どうしようか
'**********
Dim mysheet As Worksheet
Application.ScreenUpdating = False
    On Error Resume Next
    For Each mysheet In Worksheets
        If InStr(mysheet.Name, "全体") = 0 Then
            mysheet.Name = mysheet.Range("C2").Value
        End If
    Next
    
    Call シート目次
Application.ScreenUpdating = True
End Sub

Sub シート目次()
  
    Dim i As Integer
    Dim iRow As Integer
    Dim iColumn As Integer
    Dim sName, Kumimei As String
    Dim Sh1 As Worksheet
    Set Sh1 = Worksheets("全体")
    
    Sh1.Range("A1").Select
    '目次シートの設定内容をクリア
    Sh1.Range(Sh1.Cells(3, 3), Sh1.Cells(Rows.Count, 7).End(xlUp)).ClearContents
    Sh1.Range(Sh1.Cells(3, 3), Sh1.Cells(Rows.Count, 7).End(xlUp)).Hyperlinks.Delete
      
    '目次開始行数（本例は3行目から目次が作られる）
    iRow = 3
    '目次を作成する列数（本例は3列目（C列）に目次が作られる）
    iColumn = 3
  
    'ワークシートの数分、下記処理を繰り返す
    For i = 2 To Worksheets.Count Step 1
        '非表示となっているワークシートは目次作成対象外とする
        If Worksheets(i).Visible = xlSheetVisible Then
            '組合員別所有船ファイルへのハイパーリンクを設置
            Kumimei = "'" & Worksheets(i).Cells(2, 3).Text & "'"
            Worksheets(i).Hyperlinks.Add Anchor:=Worksheets(i).Cells(4, 10), Address:="\\組合員別所有船.xlsm", SubAddress:=Kumimei & "!A1", TextToDisplay:="漁船登録", ScreenTip:=Worksheets(i).Cells(4, 10).Text
            'ワークシート名を格納
            sName = "'" & Worksheets(i).Name & "'"
            '目次シートの対象セルにハイパーリンクを設定（目次作成対象ワークシートのA1セルへのリンク
            Sh1.Hyperlinks.Add Anchor:=Sh1.Cells(iRow, iColumn), Address:="", SubAddress:=sName & "!A1", ScreenTip:=Worksheets(i).Name
            '目次シートの対象セルにシート名を設定
            Sh1.Cells(iRow, iColumn) = Worksheets(i).Name
            '全体シートに住所、生年月日、地区名、出資金を転記
            Sh1.Cells(iRow, iColumn + 1) = Worksheets(i).Cells(2, 5).Value
            Sh1.Cells(iRow, iColumn + 2) = Worksheets(i).Cells(2, 9).Value
            Sh1.Cells(iRow, iColumn + 3) = Worksheets(i).Cells(3, 9).Value
            Sh1.Cells(iRow, iColumn + 4) = Worksheets(i).Cells(28, 8).Value
              '資格情報追加
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

  
           'リンクの文字の大きさ、フォントを設定
            Sh1.Cells(iRow, iColumn).Font.Size = 12
            Sh1.Cells(iRow, iColumn).Font.Name = "ＭＳ 明朝"
            Sh1.Cells(iRow, iColumn).Font.Bold = True
  
            '次の行へ
            iRow = iRow + 1
        End If
    Next i
            Sh1.Range("B2").Value = "番号"
            Sh1.Range("C2").Value = "組合員名"
            Sh1.Range("D2").Value = "住所"
            Sh1.Range("E2").Value = "生年月日"
            Sh1.Range("F2").Value = "地区名"
            Sh1.Range("G2").Value = "出資金"
End Sub
