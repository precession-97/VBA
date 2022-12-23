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
Function GetVerticalCollection(ws As Worksheet, startCellY As Long, startCellX As Long)

    Dim verCollection As Collection
    Set verCollection = New Collection
    Dim dy As Long: dy = 0
    Dim data As String: data = ""
    
    '/* ������������������܂ŌJ��Ԃ� */
    Do
        '// �ΏۃZ���̒l���擾
        data = ws.Cells(startCellY + dy, startCellX).Value
        If data = "" Then
            '// �󔒂̃Z�������m�i��������j
            Exit Do
        End If
        
        '// �󔒂̃Z���o�Ȃ��ꍇ�A�擾�����l��Collection�Ɋi�[
        verCollection.Add (data)
        '// ����Q�Ɛ�̃Z�����W���s������1�V�t�g
        dy = dy + 1
    Loop
    
    Set GetVerticalCollection = verCollection

End Function

' /* ---------------------------------------------------------------------------
'
'   ������ targetStr �ɑ΂��āA������p�^�[�� pattern �����v���邩�𔻒肷��֐�
'   Args:
'       targetStr   String      ������̓��e�����؂������f�[�^
'       pattern     String      ���؂��������e�i���K�\���j
'                               ���K�\���̋L�q���@�̓T�C�g���ŎQ�Ɓi�������[�h�F"vba regexp" etc.�j
'
'   Return:
'       Boolean                 pattern�ɍ��v���Ă����ꍇ True
'
' ----------------------------------------------------------------------------
' [Ex.]
'       targetStr = "10E8", pattern = "[^0-9]"  >> True    (targetStr�ɔ��p�����ȊO�̕������܂܂��)
'       -> ���`�I��0�ȏ�̐�������Ɉ��    cf. If IsMatched2Pattern(str, "[^0-9]") Then (0�ȏ�̐����ł͂Ȃ�)
'
'       �y���Y�^�zIsNumeric�֐����� "10E8" �� True (100000000 �Ɖ��߂����)
'
' ----------------------------------------------------------------------------*/
Function IsMatched2Pattern(targetStr As String, pattern As String) As Boolean

    Dim objReg As Object
    Set objReg = CreateObject("VBScript.RegExp")
    
    objReg.IgnoreCase = False   '// �啶���E�����������
    objReg.pattern = pattern
    
    '// pattern�Ƀ}�b�`����ꍇ��True
    IsMatched2Pattern = objReg.test(targetStr)

End Function
