Attribute VB_Name = "���y�[�W�폜"
Option Explicit

Sub ���y�[�W�폜()


'*****�g���܂���******


'���y�[�W�̍s�̃��X�g���擾
    Dim pageBreaks As hPageBreaks:
    Set pageBreaks = ActiveSheet.hPageBreaks
    Dim pageList As New Collection
   
'hPageBreaks�^��Collection�^�ɕϊ�
    Dim obj As Variant
    For Each obj In ActiveSheet.hPageBreaks
        pageList.Add (obj.Location.Row)
    Next
        
    Dim reList As New Collection
    Set reList = ���X�g�t��(pageList)
    
    Dim i, dwnLine, upLine As Long
    For i = 1 To reList.count Step 2
         dwnLine = reList(i) - 1
         upLine = reList(i + 1)
         ActiveSheet.Range(Rows(dwnLine), _
                           Rows(upLine)).Delete
    Next i

End Sub

Public Function ���X�g�t��(ByRef list As Collection) As Collection
    
    Dim i As Integer
    Dim reList As Collection
    Set reList = New Collection
    
    For i = list.count To 1 Step -1
        reList.Add (list(i))
    Next i
    
    Set ���X�g�t�� = reList

End Function


Public Sub ���X�g�t���e�X�g()

    Dim list As New Collection
    list.Add (1)
    list.Add (2)
    list.Add (3)
        
    Dim nothingList As New Collection

    
    Debug.Assert ���X�g�t��(list)(1) = list(3)
    Debug.Assert ���X�g�t��(list)(2) = list(2)
    Debug.Assert ���X�g�t��(list)(3) = list(1)
    
    
    MsgBox "�e�X�g����"
End Sub

Public Sub ���y�[�W�ݒ�()
    
'�P�y�[�W���擾
    Dim pageRange As Range
    On Error Resume Next
    Set pageRange = Application.InputBox( _
                        "�P�y�[�W�͈̔͂�I�����Ă��������B" _
                        , "���y�[�W�ݒ�" _
                        , Type:=8)
    If Err.Number <> 0 Then
        MsgBox "�L�����Z������܂����B"
        Exit Sub
    End If
    
'����͈͑S�̂��擾
    Dim printRange As Range
    Set printRange = Application.InputBox( _
                        "����͈͂�I�����Ă��������B" _
                        , "����͈͂̐ݒ�" _
                        , Type:=8)
    If Err.Number <> 0 Then
        MsgBox "�L�����Z������܂����B"
        Exit Sub
    End If
    
    
    With ActiveSheet
    
        '����͈͏������A�ݒ�
        .PageSetup.PrintArea = False
        .PageSetup.PrintArea = printRange.Address
        
        '���y�[�W�̏�����
        .ResetAllPageBreaks
        
        Dim i As Long
        
        '�P�y�[�W�̍ŏI�s����A����͈͂̍ŏI�s�܂�
        '�P�y�[�W�̍s�������ɉ��y�[�W��ݒ�
        For i = pageRange.Rows.count + 1 To printRange.Rows.count Step pageRange.Rows.count
             .Rows(i).PageBreak = xlPageBreakManual
        Next i
    End With
End Sub

Public Sub �C���v�b�g�`�[�폜()

'�C�ӂ͈̔͂���u�C���v�b�g�v���܂ރZ����T��

'�������Z���̏�x�s�A��y�s��͈͑I���E�폜


End Sub

Public Sub �S�V�[�g�K�p()
    Dim Ws As Worksheet
    For Each Ws In Worksheets
        Ws.Activate
        
        If Ws.Name = "" Then
            
        End If
    Next Ws
End Sub
