Sub merge_empty_same()

    Dim rngAll As Range                                     '��ü(A)ll ������ ������ ���� ����
    Dim rngC As Range                                       'A�� ��ü���� �� ��(C)ell �� ���� ����
    Dim r As Long                                           '��(r)ow ������ ����� ����
    Dim rowsCnt As Long                                     '��ü������ �������� ���� ����

    Application.ScreenUpdating = False                     'ȭ�� ������Ʈ (�Ͻ�)����
    Application.DisplayAlerts = False                       'ȭ�� ��� ����(�����ս� ��� �����ϱ� ����)


' ���ڿ��� �ִ� ���� ���� ����
'    With ActiveSheet.UsedRange                             '��ü ��뿵������
'        Set rngAll = .Offset(1).Resize(.Rows.Count - 1, 1) 'A�� ������ �������� ������ ����
'    End With
    
'    For Each rngC In rngAll                                 '�� ���� ������ ��ȯ
'        If rngC = vbNullString Then                         '���� ������ �����(A���� �ش�)
'            rngC.Offset(-1).Resize(2).Merge                 '������ ������
'        End If
'    Next rngC


'    With ActiveSheet.UsedRange                             '��ü ��뿵������
'        Set rngAll = .Offset(1).Resize(.Rows.Count - 1, 2) 'B�� ������ �������� ������ ����
'    End With

'    For Each rngC In rngAll                                 '�� ���� ������ ��ȯ
'        If rngC = vbNullString Then                         '���� ������ �����(B���� �ش�)
'            rngC.Offset(-1).Resize(2).Merge                 '������ ������
'        End If
'    Next rngC




' �Ʒ��� ��ĭ�� ����
    rowsCnt = Cells(Rows.Count, 1).End(3).Row               '��ü���� ���������� ������

' A��  �ؿ��� ���� ����
    For r = rowsCnt To 2 Step -1                            '���� �������࿡�� �� �྿ ���̸鼭 �ݺ�
        If Cells(r, 1) = Cells(r + 1, 1) Then               '���� ������ �Ʒ����� ������(A�� �ش�)
            Cells(r, 1).Resize(2).Merge                     '(A��)�� ���� ������
        End If
    Next r

'B��  �ؿ��� ���� ����
    For r = rowsCnt To 2 Step -1                            '���� �������࿡�� �� �྿ ���̸鼭 �ݺ�
        If Cells(r, 2) = Cells(r + 1, 2) Then               '���� ������ �Ʒ����� ������(B�� �ش�)
            If Cells(r + 1, 1).Value <> "" Then             '���� A���� ���� ��������
            Else
                Cells(r, 2).Resize(2).Merge                     '(B��)�� ���� ������
            End If
        End If
    Next r




'A���������� B�� C�� ����  ������ �Ʒ��� ����
    For r = 1 To rowsCnt Step 1                             '���� �������࿡�� �� �྿ ���̸鼭 �ݺ�
        If Cells(r, 1).Value <> "" Then                     '���� Cell �� ������ �ƴϸ� 
            Cells(r, 1).Resize(,3).Merge                    'A��, B��, C���� ���� 
        End If
    Next r

'B���������� C�� ����   ������ �Ʒ��� ����
    For r = 1 To rowsCnt Step 1                             '���� �������࿡�� �� �྿ ���̸鼭 �ݺ�
        If Cells(r, 2).Value <> "" Then                     '���� Cell �� ������ �ƴϸ� 
            Cells(r, 2).Resize(,2).Merge                    'A��, B��, C���� ���� 
        End If
    Next r


' ������ row , row + 1 �϶� ó�� �ֱ�
'C��   �Ʒ��� ���� ����  If Cells(r, isColumn - 1).MergeCells = False Then          '���� ��(B��)�� ���յǾ����� �Ǻ� 
    isColumn = 3                                                   'C��
    For r = rowsCnt To 2 Step -1                                   '���� �������࿡�� �� �྿ ���̸鼭 �ݺ�
        If Cells(r, isColumn) = Cells(r + 1, isColumn) Then        '���� ������ �Ʒ����� ������(A�� �ش�)
            If Cells(r + 1, isColumn - 2).Value = "" And Cells(r + 1, isColumn - 1).Value = "" Then   'A��, B�� �����϶��� 
                Cells(r, isColumn).Resize(2).Merge                 '(A��)�� ���� ������  ������ �ȵǾ� �������� ����
            End If
        End If
    Next r


    
    Application.DisplayAlerts = True                        'ȭ�� ��� ����

    Set rngAll = Nothing                                    '��ü���� �ʱ�ȭ(�޸� ����)

End Sub

