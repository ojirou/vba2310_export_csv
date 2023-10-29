Attribute VB_Name = "Module1"
Sub MasterProcedure()
    ' �}�X�^�[�v���V�[�W���ő��̃v���V�[�W�����Ăяo��
    CopySheet ' ��Ɨp�V�[�g���쐬
    QuoteValuesInColumnsAtoV '�_�u���N�H�[�e�[�V�����ݒ�
'    QuoteValuesInColumnT�@'T�s�݂̂Ƀ_�u���N�H�[�e�[�V�����ݒ�
    Make_Csv_Quote  '�_�u���N�H�[�e�[�V�����t���Ńt�@�C���o��
End Sub
'#############################################################################
' ��Ɨp�V�[�g���쐬
'�@copy_sheet
'#############################################################################
Sub CopySheet()
    Dim ws As Worksheet
    ' �V�[�g�u�Œ莑�Y�䒠�v�����݂��邩�m�F
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("�Œ莑�Y�䒠")
    On Error GoTo 0
    ' �V�[�g�u�Œ莑�Y�䒠�v�����݂���ꍇ�A�R�s�[���āu�Œ莑�Y�䒠bak�v�Ƃ������O�ŕۑ�
    If Not ws Is Nothing Then
        Application.DisplayAlerts = False ' �x�����\���ɂ���
        ws.Copy Before:=ThisWorkbook.Sheets(1)
        Application.DisplayAlerts = True ' �x�����\���ɂ���
        ' �V�����쐬���ꂽ�V�[�g�ɖ��O��ݒ�
        ActiveSheet.Name = "�Œ莑�Y�䒠_��Ɨp"
        MsgBox "�V�[�g�u�Œ莑�Y�䒠�v���u�Œ莑�Y�䒠_��Ɨp�v�Ƃ��ăR�s�[����܂����B"
    Else
        MsgBox "�V�[�g�u�Œ莑�Y�䒠�v�����݂��܂���B"
    End If
End Sub
'#############################################################################
' �_�u���N�H�[�e�[�V�����ݒ�
'�@quote_values_A2V
'#############################################################################
Sub QuoteValuesInColumnsAtoV()
    Dim ws As Worksheet
    Dim LastRow As Long
    Dim Cell As Range
    Dim Column As Range
   ' ���[�N�V�[�g���w��
    Set ws = ThisWorkbook.Worksheets("�Œ莑�Y�䒠_��Ɨp") ' �V�[�g����K�؂Ȃ��̂ɕύX
    ' �f�[�^�̍ŏI�s���擾
    LastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    ' A�񂩂�AR��̃Z�����̒l���_�u���N�H�[�e�[�V�����ň͂�
     For Each Column In ws.Range("A4:V" & LastRow).Columns
        For Each Cell In Column.Cells
            If Not IsEmpty(Cell.Value) Then
                Cell.Value = """" & Cell.Value & """"
            End If
        Next Cell
    Next Column
    MsgBox "A�񂩂�V��̃Z�����̒l���_�u���N�H�[�e�[�V�����ň͂܂�܂����B"
End Sub
'#############################################################################
' �_�u���N�H�[�e�[�V�����ݒ�
'  quote_values_T
'#############################################################################
Sub QuoteValuesInColumnT()
    Dim ws As Worksheet
    Dim LastRow As Long
    Dim Cell As Range
    ' ���[�N�V�[�g���w��
    Set ws = ThisWorkbook.Worksheets("�Œ莑�Y�䒠_��Ɨp") ' �V�[�g����K�؂Ȃ��̂ɕύX
    ' �f�[�^�̍ŏI�s���擾
    LastRow = ws.Cells(ws.Rows.Count, "T").End(xlUp).Row
    ' Q��̃Z�����̒l���_�u���N�H�[�e�[�V�����ň͂�
    For Each Cell In ws.Range("T4:T" & LastRow)
        If Not IsEmpty(Cell.Value) Then
            Cell.Value = """" & Cell.Value & """"
        End If
    Next Cell
    MsgBox "T��̃Z�����̒l���_�u���N�H�[�e�[�V�����ň͂܂�܂����B"
End Sub
'#############################################################################
' �_�u���N�H�[�e�[�V�����t���Ńt�@�C���o��
'�@make_csv_quote
'#############################################################################
Sub Make_Csv_Quote()
    Dim i As Integer, j As Integer, FileNumber As Integer, LR As Integer, LC As Integer
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("�Œ莑�Y�䒠_��Ɨp")  ' ���[�N�V�[�g��ݒ�
    Dim OutputCsv As String, OutputTxt As String
    OutputCsv = Environ("USERPROFILE") & "\Desktop\sample.csv"
    OutputTxt = Environ("USERPROFILE") & "\Desktop\sample.txt"
    ' �ŏ���1�s���폜
    On Error Resume Next ' �G���[�𖳎����Ď��̍s�ɐi��
    ws.Rows("1:1").Delete Shift:=xlUp ' 1����3�s���폜���ď�ɃV�t�g
    On Error GoTo 0 ' �G���[�n���h�����O�����ɖ߂�
    ' W��ȍ~�̗�폜
    Dim DeleteRange As Range
    Set DeleteRange = ws.Range("W:ZZ") ' �폜��������͈̔͂��w��
    DeleteRange.Delete Shift:=xlToLeft
    ' T�񂪋�̍s�폜
    Dim lRow As Integer
    lRow = Cells(Rows.Count, "T").End(xlUp).Row  'T��̍ŏI�s���擾
    For i = lRow To 1 Step -1 '�ŏI�s����1�s�ڂ܂ŌJ��Ԃ��B
        If IsEmpty(Cells(i, "T").Value) Then  'T��̋󔒍s�𔻒�
                Rows(i).Delete '�Y���i�󔒁j����s�폜
        End If
    Next i
    Dim NewWorkbook As Workbook
    Dim CSVFilePath As String
    ' �V�������[�N�u�b�N���쐬
    Set NewWorkbook = Workbooks.Add
    ' �V�������[�N�u�b�N�Ƀf�[�^���R�s�[
    ws.Copy Before:=NewWorkbook.Sheets(1)
    ' �V�������[�N�u�b�N��CSV�t�@�C���ɕۑ�
    NewWorkbook.SaveAs OutputCsv, xlCSV
    ' �V�������[�N�u�b�N�����i�ύX����j
    NewWorkbook.Close SaveChanges:=False
    ' ���b�Z�[�W��\��
    MsgBox "�u�Œ莑�Y�䒠_��Ɨp�v�V�[�g�̓��e��CSV�t�@�C���ɃG�N�X�|�[�g����܂����B", vbInformation
'�ŏ��̍��ڍs�Q�s���폜
    Workbooks.OpenText Filename:=OutputCsv
    Rows("1:2").Delete
    ActiveWorkbook.Save
    ActiveWindow.Close True
' ��Ɨp�V�[�g���폜
    If Not ws Is Nothing Then
        Application.DisplayAlerts = False ' �x�����\���ɂ���
        ws.Delete
        Application.DisplayAlerts = True
        MsgBox sheetName & " �V�[�g���폜����܂����B", vbInformation
    Else
        MsgBox sheetName & " �V�[�g�͑��݂��܂���B", vbExclamation
    End If
End Sub
