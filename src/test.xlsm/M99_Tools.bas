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
Function GetVerticalCollection(ws As Worksheet, startCellY As Long, startCellX)
    Dim verCollection As Collection
    Set verCollection = New Collection
    Dim dy As Long: dy = 0
    Dim data As String: data = ""
    Do
        data = ws.Cells(startCellY + dy, startCellX).Value
        If data = "" Then
            Exit Do
        End If
        verCollection.Add (data)
        dy = dy + 1
    Loop
    Set GetVerticalCollection = verCollection
End Function

