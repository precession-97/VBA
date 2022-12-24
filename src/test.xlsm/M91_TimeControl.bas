Attribute VB_Name = "M91_TimeControl"
Option Explicit

'// 構造体
Public Type ExTime

    vHour As Long       '// 時間
    vMinute As Long     '// 分

End Type

'/* ------------------------------------------------------------------------
'
' 与えられた文字列を独自の時間型(ExTime)として変換する関数
' Args;
'   sTime       String      "hh:mm" ※ "hh:mm:ss" etc. はNG
'
' Return
'               Extime
'
' ------------------------------------------------------------------------- */
Function ConvStr2ExTime(sTime As String) As ExTime

    Dim strArray() As String
    strArray = Split(sTime, ":")
    
    Dim et As ExTime
    
    '/* 引数チェック① */
    If UBound(strArray) <> 1 Then
    
        '// sTimeに ":" が含まれていない、または 2つ以上含まれている ( -> 正常に変換できない)
        '// エラー出力
        Call Err.Raise(Number:=10001, Description:="想定外の文字列" & sTime & "を検知しました。")
        Exit Function
    
    End If
    
    '/* ▼ strArrayは(0),(1) で構成 */
    
    '/* 引数チェック② */
    If M99_Tools.IsMatched2Pattern(strArray(0), "[^0-9]") _
            Or M99_Tools.IsMatched2Pattern(strArray(1), "[^0-9]") Then
    
        '// strArray内のどちらかに半角数字でない文字が含まれている ( -> 正常に変換できない)
        '// エラー出力
        Call Err.Raise(Number:=10002, Description:="想定外の文字列" & sTime & "を検知しました。")
        Exit Function
    
    End If
    
    '/* ▼ strArray(0),(1) は半角数字のみで構成 */
    
    '/* 引数チェック③ */
    If strArray(0) = "" Or strArray(1) = "" Then
    
        '// strArray内のどちらかが文字数0の文字列である ( -> 正常に変換できない)
        '// エラー出力
        Call Err.Raise(Number:=10003, Description:="想定外の文字列" & sTime & "を検知しました。")
        Exit Function
    
    End If
    
    '// strArray(0) = "hh" -> ExTime.vHour
    et.vHour = CLng(strArray(0))
    
    '// strArray(1) = "mm" -> ExTime.vMinute
    et.vMinute = CLng(strArray(1))
    
    '// vMinuteが 0～59 となるように vMinute, vHour を調整
    Do While et.vMinute >= 60
    
        et.vHour = et.vHour + 1         '// 時間を1繰り上げ
        et.vMinute = et.vMinute - 60    '// 分を60減らす
    
    Loop
    
    ConvStr2ExTime = et

End Function

'/* ------------------------------------------------------------------------
'
' 与えられたExTime型のデータをStringとして変換する関数
'
' Args;
'   et          Extime
'
' Return
'               String      "hh:mm"
'
' ------------------------------------------------------------------------- */
Function ConvExTime2String(et As ExTime) As String

    ConvExTime2String = CStr(et.vHour) & ":" & Format(CStr(et.vMinute), "00")

End Function

'/* ------------------------------------------------------------------------
'
' 与えられたExTime型のデータ同士を比較した結果を返す関数
'
' Args;
'   et1         Extime
'   et2         Extime
'
' Return
'               String      ">" or "=" or "<"
'
' ------------------------------------------------------------------------- */
Function CompareExTime(et1 As ExTime, et2 As ExTime) As String

    '// vHour比較
    If et1.vHour > et2.vHour Then
        CompareExTime = ">"
        Exit Function
    ElseIf et1.vHour < et2.vHour Then
        CompareExTime = "<"
        Exit Function
    End If
    
    '// ▼ 時間は等しい -> 分 で比較
    '// vMinute比較
    If et1.vMinute > et2.vMinute Then
        CompareExTime = ">"
        Exit Function
    ElseIf et1.vMinute < et2.vMinute Then
        CompareExTime = "<"
        Exit Function
    End If
    
    '// ▼ 時間も分も等しい
    CompareExTime = "="
    Exit Function

End Function

'/* -----------------------------------------------------------------------------------------
'
' refTimeを一日の基準（起点）時刻とした際の、etの時刻を返す関数
'
' Args;
'   et          ExTime      変換対象の時刻      (00:00 - 23:59の範囲で指定)
'   refTime     ExTime      基準（起点）時刻    (00:00 - 23:59の範囲で指定)
'
' Return
'               ExTime      基準時刻に基づいて変換されたExTime
'-------------------------------------------------------------------------------------------
' [Ex.] et = 01:30, refTime = 04:00     >> Return 25:30  (04:00 - 27:59 の範囲になるよう変換)
'
' ----------------------------------------------------------------------------------------- */
Function ConvBasedOnRefTime(et As ExTime, refTime As ExTime) As ExTime

    '/* 引数チェック */
    If et.vHour >= 24 Or refTime.vHour >= 24 Then
        '// 00:00 - 23:59 の範囲外である値を検知（想定外）
        '// エラー出力
        Call Err.Raise(Number:=10001, Description:="想定外の引数が渡されました。")
    End If
    
    If CompareExTime(et, refTime) = ">" Or CompareExTime(et, refTime) = "=" Then
        '// 無変換で返却
        ConvBasedOnRefTime = et
    End If
    
    '// ▼ 時間変換(24時間加算)
    Dim convedEt As ExTime: convedEt = et
    convedEt.vHour = convedEt.vHour + 24
    
    ConvBasedOnRefTime = convedEt

End Function
