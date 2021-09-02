Sub merge_Empty_Same_And_Sum()

    Dim rngAll As Range                                     '전체(A)ll 데이터 영역을 넣을 변수
    Dim rngC As Range                                       'A열 전체영역 각 셀(C)ell 을 넣을 변수
    Dim r As Long                                           '행(r)ow 증가에 사용할 변수
    Dim rowsCnt As Long                                     '전체데이터 마지막행 넣을 변수

    Application.ScreenUpdating = False                     '화면 업데이트 (일시)정지

    With ActiveSheet.UsedRange                             '전체 사용영역에서
        Set rngAll = .Offset(1).Resize(.Rows.Count - 1, 1) 'A열 데이터 영역만을 변수에 넣음
    End With

    Application.DisplayAlerts = False                       '화면 경고 중지(셀병합시 경고 무시하기 위해)
    For Each rngC In rngAll                                 '각 열의 각셀을 순환
        If rngC = vbNullString Then                         '만약 각셀이 비어있(A열에 해당)
            rngC.Offset(-1).Resize(2).Merge                 '윗셀과 셀병합
        End If
    Next rngC

    rowsCnt = Cells(Rows.Count, 2).End(3).Row               '전체행의 마지막행을 변수에
    For r = rowsCnt To 2 Step -1                            '제일 마지막행에서 한 행씩 줄이면서 반복
        If Cells(r, 2) = Cells(r + 1, 2) Then               '만약 각셀이 아래셀과 같으면(B열 해당)
            Cells(r, 2).Resize(2).Merge                     '(B열)두 셀을 셀병합
            Cells(r, 3) = Cells(r, 3) + Cells(r, 3).Offset(1) '(C열) 값을 더함
            Cells(r, 3).Resize(2).Merge                     '(C열) 셀병합
        End If
    Next r
    Application.DisplayAlerts = True                        '화면 경고 복원

    Set rngAll = Nothing                                    '개체변수 초기화(메모리 비우기)

End Sub
