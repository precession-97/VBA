Attribute VB_Name = "M99_Tools"
Option Explicit

' /* ---------------------------------------------------------------------------
'
'   ��Z��[Cell(startCellY, startCellX)]���牺�����Ƀf�[�^����擾����֐�
'   Args:
'       ws          �擾�Ώۂ������郏�[�N�V�[�g�I�u�W�F�N�g
'       startCellY  x���J�n��_�iEX. A�� = 1, B�� = 2�j
'       startCellX  y���J�n��_ (EX. 1�s = 1, 2�s = 2)
'
'   Return:
'       Collection  �I�u�W�F�N�g
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

