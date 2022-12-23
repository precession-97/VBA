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

