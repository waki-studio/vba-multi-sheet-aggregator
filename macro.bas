Option Explicit

' ===================================================================
' 複数シートのデータを「集計」シートに転記するマクロ
'
' 機能:
'   各店舗シート（東京店・大阪店・福岡店）のデータを
'   「集計」シートに自動転記します。
'
' 前提条件:
'   - 各店舗シートのフォーマットが揃っていること
'     1行目: 見出し（日付、商品名、数量、売上）
'     2行目以降: データ
'   - 「集計」シートが存在すること
'
' 使い方:
'   1. 各店舗シートにデータを入力
'   2. 「集計実行」ボタンを押す
'   3. 「集計」シートにデータが自動転記される
'
' カスタマイズポイント:
'   - 対象シート名: sheetNames 配列を編集
'   - 集計先シート名: "集計" を任意の名前に変更
' ===================================================================

Sub AggregateData()

    Dim wsTarget As Worksheet      ' 集計先シート
    Dim wsSource As Worksheet      ' 各店舗シート
    Dim sheetNames As Variant      ' 対象シート名の配列
    Dim i As Integer
    Dim lastRow As Long
    Dim writeRow As Long
    Dim sourceRow As Long
    Dim totalCount As Long

    ' 対象シート名（必要に応じて編集してください）
    sheetNames = Array("東京店", "大阪店", "福岡店")

    ' 画面更新を停止（高速化）
    Application.ScreenUpdating = False

    ' 「集計」シートを取得
    On Error Resume Next
    Set wsTarget = ThisWorkbook.Worksheets("集計")
    On Error GoTo 0

    If wsTarget Is Nothing Then
        MsgBox "「集計」シートが見つかりません。先に作成してください。", vbCritical
        Application.ScreenUpdating = True
        Exit Sub
    End If

    ' 集計シートの既存データをクリア（見出しは残す）
    wsTarget.Range("A2:E10000").ClearContents

    ' 書き込み開始行
    writeRow = 2
    totalCount = 0

    ' 各店舗シートをループ
    For i = LBound(sheetNames) To UBound(sheetNames)

        ' シートの存在チェック
        On Error Resume Next
        Set wsSource = ThisWorkbook.Worksheets(sheetNames(i))
        On Error GoTo 0

        If wsSource Is Nothing Then
            MsgBox "シート「" & sheetNames(i) & "」が見つかりません。スキップします。", vbExclamation
            GoTo NextSheet
        End If

        ' シートの最終行を取得
        lastRow = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row

        ' 2行目以降をループしてコピー
        For sourceRow = 2 To lastRow
            wsTarget.Cells(writeRow, 1).Value = sheetNames(i)                       ' 店舗名
            wsTarget.Cells(writeRow, 2).Value = wsSource.Cells(sourceRow, 1).Value  ' 日付
            wsTarget.Cells(writeRow, 3).Value = wsSource.Cells(sourceRow, 2).Value  ' 商品名
            wsTarget.Cells(writeRow, 4).Value = wsSource.Cells(sourceRow, 3).Value  ' 数量
            wsTarget.Cells(writeRow, 5).Value = wsSource.Cells(sourceRow, 4).Value  ' 売上
            writeRow = writeRow + 1
            totalCount = totalCount + 1
        Next sourceRow

NextSheet:
        Set wsSource = Nothing
    Next i

    ' 画面更新を再開
    Application.ScreenUpdating = True

    ' 完了メッセージ
    MsgBox "集計が完了しました。" & vbCrLf & _
           totalCount & " 件のデータを転記しました。", vbInformation

End Sub
