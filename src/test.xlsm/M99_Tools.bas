Attribute VB_Name = "M99_Tools"
Option Explicit

' /* ---------------------------------------------------------------------------
'
'   基準セル[Cell(startCellY, startCellX)]から下方向にデータ列を取得する関数
'   Args:
'       ws          取得対象が属するワークシートオブジェクト
'       startCellY  x軸開始基準点（EX. A列 = 1, B列 = 2）
'       startCellX  y軸開始基準点 (EX. 1行 = 1, 2行 = 2)
'
'   Return:
'       Collection  オブジェクト
'
' ----------------------------------------------------------------------------*/
Function GetVerticalCollection(ws As Worksheet, startCellY As Long, startCellX As Long)

    Dim verCollection As Collection
    Set verCollection = New Collection
    Dim dy As Long: dy = 0
    Dim data As String: data = ""
    
    '/* 特定条件が満たされるまで繰り返し */
    Do
        '// 対象セルの値を取得
        data = ws.Cells(startCellY + dy, startCellX).Value
        If data = "" Then
            '// 空白のセルを検知（特定条件）
            Exit Do
        End If
        
        '// 空白のセル出ない場合、取得した値をCollectionに格納
        verCollection.Add (data)
        '// 次回参照先のセル座標を行方向に1シフト
        dy = dy + 1
    Loop
    
    Set GetVerticalCollection = verCollection

End Function

' /* ---------------------------------------------------------------------------
'
'   文字列 targetStr に対して、文字列パターン pattern が合致するかを判定する関数
'   Args:
'       targetStr   String      文字列の内容を検証したいデータ
'       pattern     String      検証したい内容（正規表現可）
'                               正規表現の記述方法はサイト等で参照（検索ワード："vba regexp" etc.）
'
'   Return:
'       Boolean                 patternに合致していた場合 True
'
' ----------------------------------------------------------------------------
' [Ex.]
'       targetStr = "10E8", pattern = "[^0-9]"  >> True    (targetStrに半角数字以外の文字が含まれる)
'       -> 狭義的な0以上の整数判定に一役    cf. If IsMatched2Pattern(str, "[^0-9]") Then (0以上の整数ではない)
'
'       【備忘録】IsNumeric関数だと "10E8" は True (100000000 と解釈される)
'
' ----------------------------------------------------------------------------*/
Function IsMatched2Pattern(targetStr As String, pattern As String) As Boolean

    Dim objReg As Object
    Set objReg = CreateObject("VBScript.RegExp")
    
    objReg.IgnoreCase = False   '// 大文字・小文字を区別
    objReg.pattern = pattern
    
    '// patternにマッチする場合はTrue
    IsMatched2Pattern = objReg.test(targetStr)

End Function
