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

'/* ------------------------------------------------------------------------
'
' �^����ꂽExTime�^�̃f�[�^��String�Ƃ��ĕϊ�����֐�
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
' �^����ꂽExTime�^�̃f�[�^���m���r�������ʂ�Ԃ��֐�
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

    '// vHour��r
    If et1.vHour > et2.vHour Then
        CompareExTime = ">"
        Exit Function
    ElseIf et1.vHour < et2.vHour Then
        CompareExTime = "<"
        Exit Function
    End If
    
    '// �� ���Ԃ͓����� -> �� �Ŕ�r
    '// vMinute��r
    If et1.vMinute > et2.vMinute Then
        CompareExTime = ">"
        Exit Function
    ElseIf et1.vMinute < et2.vMinute Then
        CompareExTime = "<"
        Exit Function
    End If
    
    '// �� ���Ԃ�����������
    CompareExTime = "="
    Exit Function

End Function

'/* -----------------------------------------------------------------------------------------
'
' refTime������̊�i�N�_�j�����Ƃ����ۂ́Aet�̎�����Ԃ��֐�
'
' Args;
'   et          ExTime      �ϊ��Ώۂ̎���      (00:00 - 23:59�͈̔͂Ŏw��)
'   refTime     ExTime      ��i�N�_�j����    (00:00 - 23:59�͈̔͂Ŏw��)
'
' Return
'               ExTime      ������Ɋ�Â��ĕϊ����ꂽExTime
'-------------------------------------------------------------------------------------------
' [Ex.] et = 01:30, refTime = 04:00     >> Return 25:30  (04:00 - 27:59 �͈̔͂ɂȂ�悤�ϊ�)
'
' ----------------------------------------------------------------------------------------- */
Function ConvBasedOnRefTime(et As ExTime, refTime As ExTime) As ExTime

    '/* �����`�F�b�N */
    If et.vHour >= 24 Or refTime.vHour >= 24 Then
        '// 00:00 - 23:59 �͈̔͊O�ł���l�����m�i�z��O�j
        '// �G���[�o��
        Call Err.Raise(Number:=10001, Description:="�z��O�̈������n����܂����B")
    End If
    
    If CompareExTime(et, refTime) = ">" Or CompareExTime(et, refTime) = "=" Then
        '// ���ϊ��ŕԋp
        ConvBasedOnRefTime = et
    End If
    
    '// �� ���ԕϊ�(24���ԉ��Z)
    Dim convedEt As ExTime: convedEt = et
    convedEt.vHour = convedEt.vHour + 24
    
    ConvBasedOnRefTime = convedEt

End Function
