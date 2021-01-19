Attribute VB_Name = "Module3"
Option Explicit

Sub 期間集計()
'*******************************************************************************************************************************
'セーフティネット作成様式
'1. 購買システムよりトラン参照
'2.オートフィルタにて指定期間内の燃油類を抽出
'3.差し込み印刷のために表化
'4.閉じて、ワードから差し込み印刷
'***現状、トランの自動更新が不安。またオートフィルタも手動で丁寧に。表も空欄がないか要チェック。←SQLにて対応、動作良好。
'2020/06/10作成
'2020/10/12更新(実務使用開始)
'2020/10/14更新(SQL対応。参考シート削除)コードが煩雑すぎるから手直しそのうち
'*******************************************************************************************************************************
    
    Dim MaxC1, MaxC2, cnt, yoko, i, j, k As Long
    Dim Sh1, Sh2, Sh3, Sh4 As Worksheet
    Dim Cord, Hinmei As String
    Dim Kazu As Double
    Dim d1, d2, d3 As Date
    
    '---ワークシートを宣言---
    Set Sh1 = Worksheets("記入")
    Set Sh2 = Worksheets("SN加入者一覧")
    Set Sh3 = Worksheets("DB")
'    Set Sh4 = Worksheets("参考")
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Call DBシート整理
    
     '---参考シートを検索---
    MaxC1 = Sh3.Range("A65536").End(xlUp).Row '-----参考シートの最終行
    MaxC2 = Sh2.Range("A65536").End(xlUp).Row '-----SN加入者一覧シートの最終行
    cnt = Sh2.Range("XFD2").End(xlToLeft).Column '-----SN加入者一覧シートの最終列
    
    '---集計開始日と集計終了日を指定---
    If Sh1.Range("C5").Value = "1" Then
        d1 = "2020/04/01"
        d2 = "2020/06/30"
'        Sh3.Range("A1").AutoFilter Field:=3, Criteria1:=">=2020/04/01", Operator:=xlAnd, Criteria2:="<=2020/06/30"
'        Sh3.Range("A1").AutoFilter Field:=6, Criteria1:=1
'        Sh3.Range("A1").AutoFilter Field:=2, Criteria1:=11
    'チェックマーク処理
        Sh2.Range(Sh2.Range("Q3"), Sh2.Cells(MaxC2, 20)).Value = ""
        Sh2.Range(Sh2.Range("Q3"), Sh2.Cells(MaxC2, 17)).Value = ChrW(10003)
    ElseIf Sh1.Range("C5").Value = "2" Then
        d1 = "2020/07/01"
        d2 = "2020/09/30"
'        Sh3.Range("A1").AutoFilter Field:=3, Criteria1:=">=2020/07/01", Operator:=xlAnd, Criteria2:="<=2020/09/30"
'        Sh3.Range("A1").AutoFilter Field:=6, Criteria1:=1
'        Sh3.Range("A1").AutoFilter Field:=2, Criteria1:=11
    'チェックマーク処理
        Sh2.Range(Sh2.Range("Q3"), Sh2.Cells(MaxC2, 20)).Value = ""
        Sh2.Range(Sh2.Range("R3"), Sh2.Cells(MaxC2, 18)).Value = ChrW(10003)
    ElseIf Sh1.Range("C5").Value = "3" Then
        d1 = "2020/10/01"
        d2 = "2020/12/31"
'        Sh3.Range("A1").AutoFilter Field:=3, Criteria1:=">=2020/10/01", Operator:=xlAnd, Criteria2:="<=2020/12/31"
'        Sh3.Range("A1").AutoFilter Field:=6, Criteria1:=1
'        Sh3.Range("A1").AutoFilter Field:=2, Criteria1:=11
    'チェックマーク処理
        Sh2.Range(Sh2.Range("Q3"), Sh2.Cells(MaxC2, 20)).Value = ""
        Sh2.Range(Sh2.Range("S3"), Sh2.Cells(MaxC2, 19)).Value = ChrW(10003)
    ElseIf Sh1.Range("C5").Value = "4" Then
        d1 = "2021/01/01"
        d2 = "2021/03/31"
'        Sh3.Range("A1").AutoFilter Field:=3, Criteria1:=">=2021/01/01", Operator:=xlAnd, Criteria2:="<=2021/03/31"
'        Sh3.Range("A1").AutoFilter Field:=6, Criteria1:=1
'        Sh3.Range("A1").AutoFilter Field:=2, Criteria1:=11
    'チェックマーク処理
        Sh2.Range(Sh2.Range("Q3"), Sh2.Cells(MaxC2, 20)).Value = ""
        Sh2.Range(Sh2.Range("T3"), Sh2.Cells(MaxC2, 20)).Value = ChrW(10003)
    Else: MsgBox "半期指定が間違っています。"
    End If
'
'    '-----オートフィルタ後データをコピペ
'    Sh3.Range("A1").CurrentRegion.Copy Sh4.Range("A1")
'    '-----参考シートのコードを文字列に変換する
'    Sh4.Range("D:D").NumberFormatLocal = "@"
    Sh3.Range("F:F").NumberFormatLocal = "@" '←SQL命令文側でFORMATしょりにする
     '---参考シートを検索---
    MaxC1 = Sh3.Range("A65536").End(xlUp).Row '-----参考シートの最終行
    MaxC2 = Sh2.Range("A65536").End(xlUp).Row '-----SN加入者一覧シートの最終行
    cnt = Sh2.Range("XFD2").End(xlToLeft).Column '-----SN加入者一覧シートの最終列
    
    For k = 3 To MaxC2
        Cord = Sh2.Range("C" & k).Text
    '---照合したい条件でデータベースを検索---
        For j = 0 To cnt - 25
            Kazu = 0
            Hinmei = Sh2.Range("W2").Offset(0, j + 1)
            For i = 2 To MaxC1
                d3 = Sh3.Range("B" & i).Value '-----参考シートの日付列
                If d3 >= d1 And d3 <= d2 Then
                    If Sh3.Range("C" & i).Text = Cord Then
                        If Sh3.Range("F" & i).Text = Hinmei Then
                            Kazu = Kazu + Sh3.Range("H" & i).Value
                        End If
                    End If
                End If
            Next
    '---SN加入者一覧シートの表に出力---
            Sh2.Range("W" & k).Offset(0, j + 1).Value = Kazu
        Next
    Next
    '-----オートフィルター解除
'    Sh3.Range("A1").AutoFilter
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

Sub DBシート整理()

Dim コネクション As New ADODB.Connection
Dim レコード As New ADODB.Recordset
Dim コマンド As New ADODB.Command
Dim d1, d2 As Date
Dim Sh1, Sh4 As Worksheet
Set Sh1 = Worksheets("記入")
Set Sh4 = Worksheets("DB")

コネクション.Open ConnectionString:= _
"Provider=Microsoft.ACE.OLEDB.12.0;" & _
"Data Source=\\192.168.10.250\ackobai\kobai.mdb;"

'期間適用範囲
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

With コマンド
    .ActiveConnection = コネクション
    .CommandText = _
    "SELECT 状態区分,日　　付,FORMAT(コード,'@'),コード名　漢字,分類コード,商品コード,商品名　漢字,数　　量 FROM KT01 WHERE 状態区分 = 0 and 日　　付 BETWEEN " & "#" & d1 & "#" & " And " & "#" & d2 & "#" & " and 商品コード=1001;"
    Set レコード = .Execute
End With
Sh4.Cells.Delete
Sh4.Range("A1").Value = "状態区分"
Sh4.Range("B1").Value = "日　　付"
Sh4.Range("C1").Value = "コード"
Sh4.Range("D1").Value = "コード名　漢字"
Sh4.Range("E1").Value = "分類コード"
Sh4.Range("F1").Value = "商品コード"
Sh4.Range("G1").Value = "商品名　漢字"
Sh4.Range("H1").Value = "数　　量"

Sh4.Range("A2").CopyFromRecordset Data:=レコード
レコード.Close: Set レコード = Nothing
コネクション.Close: Set コネクション = Nothing

End Sub
