Attribute VB_Name = "Module1"

' ----------------------------------------------------------
' ���t�R�s�[
'
' ----------------------------------------------------------
Sub Macro1()
Attribute Macro1.VB_Description = "���t�R�s�["
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"

    ' �J�n�ʒu�Ɉړ�
    Range("C2").Select

    ' �����̓Z���̂P��̃Z���Ɉړ�
    Selection.End(xlDown).Select

    ' TODO: �J�����g�Z�����ŏI�Z����������A���[�v�𔲂���

    ' TODO: �J�����g�Z���̐�Έʒu�擾
    ' TODO: �J�����g�Z���Ɨׂ�D�Z����I��
    Range("C89:D89").Select

    ' �R�s�[���f�[�^��ێ�
    Selection.Copy

    ' TODO: �R�s�[���Z���̂P���̃Z����Έʒu�擾
    ' TODO: �R�s�[���Z���̂P���̃Z���Ɉړ�
    Range("C90").Select

    ' ���̓��͍σZ���܂ŃJ�[�\���ړ�
    Range(Selection, Selection.End(xlDown)).Select

    ' TODO: �R�s�[��Z���̂P��̃Z���Ɉړ�
    Range("C90:C100").Select
    ActiveSheet.Paste
End Sub
