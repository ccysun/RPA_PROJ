Sub merge_empty_same()

    Dim rngAll As Range                                     '��ü(A)ll ������ ������ ���� ����
    Dim rngC As Range                                       'A�� ��ü���� �� ��(C)ell �� ���� ����
    Dim r As Long                                           '��(r)ow ������ ����� ����
    Dim rowsCnt As Long                                     '��ü������ �������� ���� ����

    Application.ScreenUpdating = False                     'ȭ�� ������Ʈ (�Ͻ�)����

    With ActiveSheet.UsedRange                             '��ü ��뿵������
        Set rngAll = .Offset(1).Resize(.Rows.Count - 1, 1) 'A�� ������ �������� ������ ����
    End With

    Application.DisplayAlerts = False                       'ȭ�� ��� ����(�����ս� ��� �����ϱ� ����)
    
    For Each rngC In rngAll                                 '�� ���� ������ ��ȯ
        If rngC = vbNullString Then                         '���� ������ �����(A���� �ش�)
            rngC.Offset(-1).Resize(2).Merge                 '������ ������
        End If
    Next rngC


    With ActiveSheet.UsedRange                             '��ü ��뿵������
        Set rngAll = .Offset(1).Resize(.Rows.Count - 1, 2) 'B�� ������ �������� ������ ����
    End With

    For Each rngC In rngAll                                 '�� ���� ������ ��ȯ
        If rngC = vbNullString Then                         '���� ������ �����(B���� �ش�)
            rngC.Offset(-1).Resize(2).Merge                 '������ ������
        End If
    Next rngC


' �Ʒ��� ��ĭ�� ����
'    rowsCnt = Cells(Rows.Count, 2).End(3).Row               '��ü���� ���������� ������
'    For r = rowsCnt To 2 Step -1                            '���� �������࿡�� �� �྿ ���̸鼭 �ݺ�
'        If Cells(r, 2) = Cells(r + 1, 2) Then               '���� ������ �Ʒ����� ������(B�� �ش�)
'            Cells(r, 2).Resize(2).Merge                     '(B��)�� ���� ������
'        End If
'    Next r
    
    Application.DisplayAlerts = True                        'ȭ�� ��� ����

    Set rngAll = Nothing                                    '��ü���� �ʱ�ȭ(�޸� ����)

End Sub

