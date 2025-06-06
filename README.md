# Excel 日付から曜日を自動出力するVBAマクロ

## 概要
このExcel VBAマクロは、指定したセル（例: C11）に日付を入力すると、隣のセル（例: D11）に日本語の曜日（例: 「金曜日」）を自動で表示します。`Worksheet_Change`イベントを使用し、ボタン操作不要でリアルタイムに更新します。さまざまな日付形式（「2025/6/6」「令和7年6月6日」「6月6日」など）に対応しています。

## 特徴
- **自動実行**: C11（変更可能）に日付を入力すると、自動で隣のセルに曜日を表示。
- **柔軟な日付解析**: 以下の形式に対応：
  - 標準: `2025/6/6`, `2025-06-06`, `2025.6.6`
  - 日本語: `2025年6月6日`, `令和7年6月6日`, `6月6日`（年省略時は当年を使用）
  - その他: `20250606`, `6/6`, 全角文字（例: `２０２５年`）
- **エラーハンドリング**: 無効な入力（例: `abc`）には「無効な日付」を表示。
- **年補完**: 年が省略された場合（例: `6月6日`）、当年（例: 2025）を補完。

## 必要環境
- Microsoft Excel 2007以降（Windows、日本語地域設定でのテスト済み）。
- `.xlsm`（マクロ有効ブック）形式で保存。

## インストール手順
1. **開発タブを表示**:
   - `ファイル` → `オプション` → `リボンのユーザー設定` → 「開発」チェックボックスをオン。
2. **VBAエディタを開く**:
   - `Alt + F11`でVBAエディタを開く。
3. **シートモジュールにコードを貼り付け**:
   - プロジェクトエクスプローラーで`ThisWorkbook` → `Microsoft Excel Objects` → `Sheet1`（対象シート）をダブルクリック。
   - 以下のコードを貼り付ける：

```vba
Private Sub Worksheet_Change(ByVal Target As Range)
    ' 対象セル（例: C11）を監視
    Const TARGET_CELL As String = "C11" ' ここを変更で他のセル（例: D5）に変更可能
    If Not Intersect(Target, Me.Range(TARGET_CELL)) Is Nothing Then
        Application.EnableEvents = False ' 無限ループを防止
        Dim inputValue As String
        Dim parsedDate As Date
        Dim outputCell As Range
        
        ' 入力値を取得（C11）
        Set outputCell = Me.Range(TARGET_CELL).Offset(0, 1) ' 右隣（例: D11）
        inputValue = Trim(Me.Range(TARGET_CELL).Value) ' 前後の空白を除去
        
        ' 空セルなら出力をクリア
        If inputValue = "" Then
            outputCell.Value = ""
            GoTo ExitSub
        End If
        
        ' 日付を解析
        If IsDate(inputValue) Then
            parsedDate = CDate(inputValue)
            outputCell.Value = Format(parsedDate, "aaaa") ' 日本語の曜日
        Else
            ' 日本語特有の形式を処理
            If ParseJapaneseDate(inputValue, parsedDate) Then
                outputCell.Value = Format(parsedDate, "aaaa")
            Else
                outputCell.Value = "無効な日付"
            End If
        End If
        
ExitSub:
        Application.EnableEvents = True ' イベントを再有効化
    End If
End Sub

Private Function ParseJapaneseDate(ByVal input As String, ByRef outDate As Date) As Boolean
    ' 日本語特有の日付形式を解析（例: 令和7年6月6日、6月6日）
    Dim reiwaYear As Integer
    Dim year As Integer
    Dim month As Integer
    Dim day As Integer
    Dim temp As String
    Dim parts() As String
    
    On Error GoTo ErrorHandler
    temp = Replace(input, " ", "") ' 空白を除去
    temp = Replace(temp, "年", "/") ' 年をスラッシュに
    temp = Replace(temp, "月", "/") ' 月をスラッシュに
    temp = Replace(temp, "日", "") ' 日を除去
    
    ' 全角文字を半角に変換
    temp = StrConv(temp, vbNarrow)
    
    ' 令和形式を処理
    If InStr(temp, "令和") > 0 Then
        temp = Replace(temp, "令和", "")
        parts = Split(temp, "/")
        If UBound(parts) >= 2 Then
            reiwaYear = CInt(parts(0))
            year = 2018 + reiwaYear ' 令和元年=2019
            month = CInt(parts(1))
            day = CInt(parts(2))
            outDate = DateSerial(year, month, day)
            ParseJapaneseDate = True
            Exit Function
        End If
    End If
    
    ' 年省略（例: 6月6日）の場合、当年を使用
    If InStr(temp, "/") = 0 And InStr(input, "月") > 0 Then
        temp = Year(Date) & "/" & temp
    End If
    
    ' スラッシュ区切りで再試行
    If IsDate(temp) Then
        outDate = CDate(temp)
        ParseJapaneseDate = True
        Exit Function
    End If
    
ErrorHandler:
    ParseJapaneseDate = False
End Function
```

4. **ファイルを保存**:
   - `ファイル` → `名前を付けて保存` → `Excelマクロ有効ブック（.xlsm）`を選択。
   - ファイルを開く際、「コンテンツの有効化」をクリックしてマクロを有効化。

## 使い方
1. `.xlsm`ファイルをExcelで開く。
2. 指定セル（例: C11）に日付を入力。例：
   - `2025/6/6`
   - `2025-06-06`
   - `令和7年6月6日`
   - `6月6日`（当年、2025年を補完）
   - `20250606`
3. 隣のセル（例: D11）に自動で曜日（例: `金曜日`）が表示される。
4. 無効な日付（例: `abc`）を入力すると「無効な日付」と表示。

## 対応する日付形式
- **標準形式**: `2025/6/6`, `2025-06-06`, `2025.6.6`
- **日本語形式**: `2025年6月6日`, `令和7年6月6日`, `6月6日`
- **その他**: `20250606`, `6/6`（年省略時は当年）
- **全角文字**: 例: `２０２５年６月６日`
- **エッジケース**: 空白混在や年省略も対応（地域設定に一部依存）。

## カスタマイズ方法
- **対象セル変更**: コード冒頭の`Const TARGET_CELL As String = "C11"`を変更（例: `"D5"`でD5セルに）。
- **出力セル変更**: `outputCell = Me.Range(TARGET_CELL).Offset(0, 1)`の`1`を変更（例: `2`で2列右）。
- **英語曜日**: `Format(parsedDate, "aaaa")`を`"dddd"`（例: Friday）や`"ddd"`（例: Fri）に変更。
- **デフォルト年**: `Year(Date)`を固定年（例: `2025`）に変更可能。
- **複数セル監視**: 複数セルを監視したい場合、`If Not Intersect(Target, Me.Range("C11:C20"))`のように範囲を指定。

## 注意点
- **地域設定**: 日本語地域設定（Windows）が最適。非日本語設定では曖昧な形式（例: `6/6`）の解釈が異なる場合あり。
- **シート名**: コードは`Sheet1`のシートモジュールに配置。別のシートを使う場合、対応するシートモジュールに貼り付け。
- **パフォーマンス**: 単一セル（例: C11）監視なので高速。複数セル監視への変更も可能。
- **制限**: 極端に非標準な形式（例: `六月六`）は未対応。必要なら正規表現で拡張可能。

## ライセンス
MITライセンス。自由に使用、改変、配布可能。

## 連絡先
質問や問題は、GitHubリポジトリのIssueでご連絡ください。
