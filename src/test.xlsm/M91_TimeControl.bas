Attribute VB_Name = "M91_TimeControl"
Option Explicit

'// �\����
Public Type ExTime

    vHour As Long       '// ����
    vMinute As Long     '// ��

End Type

'/* ------------------------------------------------------------------------
'
' �^����ꂽ�������Ǝ��̎��Ԍ^(ExTime)�Ƃ��ĕϊ�����֐�
' Args;
'   sTime       String      "hh:mm" �� "hh:mm:ss" etc. ��NG
'
' Return
'               Extime
'
' ------------------------------------------------------------------------- */
Function ConvStr2ExTime(sTime As String) As ExTime

    Dim strArray() As String
    strArray = Split(sTime, ":")
    
    Dim et As ExTime
    
    '/* �����`�F�b�N�@ */
    If UBound(strArray) <> 1 Then
    
        '// sTime�� ":" ���܂܂�Ă��Ȃ��A�܂��� 2�ȏ�܂܂�Ă��� ( -> ����ɕϊ��ł��Ȃ�)
        '// �G���[�o��
        Call Err.Raise(Number:=10001, Description:="�z��O�̕�����" & sTime & "�����m���܂����B")
        Exit Function
    
    End If
    
    '/* �� strArray��(0),(1) �ō\�� */
    
    '/* �����`�F�b�N�A */
    If M99_Tools.IsMatched2Pattern(strArray(0), "[^0-9]") _
            Or M99_Tools.IsMatched2Pattern(strArray(1), "[^0-9]") Then
    
        '// strArray���̂ǂ��炩�ɔ��p�����łȂ��������܂܂�Ă��� ( -> ����ɕϊ��ł��Ȃ�)
        '// �G���[�o��
        Call Err.Raise(Number:=10002, Description:="�z��O�̕�����" & sTime & "�����m���܂����B")
        Exit Function
    
    End If
    
    '/* �� strArray(0),(1) �͔��p�����݂̂ō\�� */
    
    '/* �����`�F�b�N�B */
    If strArray(0) = "" Or strArray(1) = "" Then
    
        '// strArray���̂ǂ��炩��������0�̕�����ł��� ( -> ����ɕϊ��ł��Ȃ�)
        '// �G���[�o��
        Call Err.Raise(Number:=10003, Description:="�z��O�̕�����" & sTime & "�����m���܂����B")
        Exit Function
    
    End If
    
    '// strArray(0) = "hh" -> ExTime.vHour
    et.vHour = CLng(strArray(0))
    
    '// strArray(1) = "mm" -> ExTime.vMinute
    et.vMinute = CLng(strArray(1))
    
    '// vMinute�� 0�`59 �ƂȂ�悤�� vMinute, vHour �𒲐�
    Do While et.vMinute >= 60
    
        et.vHour = et.vHour + 1         '// ���Ԃ�1�J��グ
        et.vMinute = et.vMinute - 60    '// ����60���炷
    
    Loop
    
    ConvStr2ExTime = et

End Function

