Imports Microsoft.VisualBasic


Public Class C_DATESUM
    Dim stmt As String
    Dim do_sql As New C_SQLFUN
    Dim dyr As System.Data.DataRow
    Dim dy_n_table As New System.Data.DataTable
    'begin_date 起始日期 yyyy/MM/dd HH
    'end_date 終止日期 yyyyMM/dd HH
    Function date_sum(ByVal begin_date As String, ByVal end_date As String, ByRef td_count As Integer, ByRef th_count As Integer, ByRef th_err_msg As String) As Boolean
        date_sum = False
        th_err_msg = ""
        Dim kr_table(500) As String
        Dim pkr As Integer = 0
        Dim TFS_date As String = ""
        Dim pk As Integer = 0
        Dim pk2 As Integer = 0
        'Dim td_count As Integer = 0
        'Dim th_count As Integer = 0
        Dim all_count As Integer = 0
        If begin_date > end_date Then
            'MsgBox("終止日期時間須大於起始日期時間")
            th_err_msg = "終止日期時間須大於起始日期時間"
            Exit Function
        End If
        '        Array.Clear(kr_table, 0, kr_table.Length)
        'date_sum = ""
        'kr_table(0) = Mid(begin_date, 1, 10)
        TFS_date = Mid(begin_date, 1, 10)
        Do While TFS_date < Mid(end_date, 1, 10)
            pkr += 1
            TFS_date = CDate(Mid(begin_date, 1, 10)).AddDays(pkr).ToString("yyyy/MM/dd")
            'kr_table(pkr) = TFS_date
            If pkr > 490 Then
                'MsgBox("請假太多天", 0, "")
                th_err_msg = "請假太多天"
                Exit Function
            End If
        Loop
        stmt = "select Fod_Date from P_12 where Fod_Date >='" + Mid(begin_date, 1, 10) + "'"
        stmt += " and Fod_Date <='" + Mid(end_date, 1, 10) + "'"
        stmt += " order by Fod_Date"
        If do_sql.db_sql(stmt, do_sql.G_conn_string) = False Then
            th_err_msg = do_sql.G_errmsg
            Exit Function
        End If
        If Mid(begin_date, 1, 10) = Mid(end_date, 1, 10) Then '同一天
            If do_sql.G_table.Rows.Count > 0 Then
                td_count = 0
                th_count = 0
            Else
                pk = CInt(Mid(begin_date, 12, 2))
                If pk < 8 Then
                    pk = 8
                End If
                pk2 = CInt(Mid(end_date, 12, 2))
                If pk2 > 17 Then
                    pk2 = 17
                End If
                If pk2 >= 13 And pk < 13 Then
                    th_count = pk2 - pk - 1
                Else
                    th_count = pk2 - pk
                End If

                If th_count >= 8 Then
                    td_count = 1
                    th_count = th_count - 8
                Else
                    td_count = 0
                End If
            End If
            'date_sum = td_count.ToString + "," + th_count.ToString

            Exit Function
        End If
        If do_sql.G_table.Rows.Count > 0 Then
            all_count = do_sql.G_table.Rows.Count
            '開始日
            TFS_date = Mid(do_sql.G_table.Rows(0).Item(0).ToString(), 1, 10)
            TFS_date = TFS_date.Replace("-", "/")
            If TFS_date = Mid(begin_date, 1, 10) Then
                td_count = 0
                th_count = 0
                all_count = all_count - 1
            Else
                td_count = 0
                If CInt(Mid(begin_date, 12, 2)) >= 13 Then
                    th_count = 17 - CInt(Mid(begin_date, 12, 2))
                Else
                    th_count = 17 - CInt(Mid(begin_date, 12, 2)) - 1
                End If
            End If
            '----結束日
            TFS_date = Mid(do_sql.G_table.Rows(do_sql.G_table.Rows.Count - 1).Item(0).ToString(), 1, 10)
            TFS_date = TFS_date.Replace("-", "/")
            If TFS_date = Mid(end_date, 1, 10) Then
                all_count = all_count - 1
            Else
                If CInt(Mid(end_date, 12, 2)) >= 13 Then
                    th_count += CInt(Mid(end_date, 12, 2)) - 8 - 1
                Else
                    th_count += CInt(Mid(end_date, 12, 2)) - 8
                End If
            End If
            If th_count = 16 Then
                td_count += 2
                th_count = 0
            ElseIf th_count >= 8 Then
                td_count += 1
                th_count = th_count - 8
            End If
            '算全部
            td_count += pkr + 1 - 2 - all_count '(全部pkr + 1)-(起始終止 2)- all_count
            'date_sum = td_count.ToString + "," + th_count.ToString
        Else 'If do_sql.G_table.Rows.Count > 0 Then
            '開始日
            td_count = 0
            If CInt(Mid(begin_date, 12, 2)) >= 13 Then
                th_count = 17 - CInt(Mid(begin_date, 12, 2))
            Else
                th_count = 17 - CInt(Mid(begin_date, 12, 2)) - 1
            End If
            '----結束日
            If CInt(Mid(end_date, 12, 2)) >= 13 Then
                th_count += CInt(Mid(end_date, 12, 2)) - 8 - 1
            Else
                th_count += CInt(Mid(end_date, 12, 2)) - 8
            End If
            If th_count = 16 Then
                td_count += 2
                th_count = 0
            ElseIf th_count >= 8 Then
                td_count += 1
                th_count = th_count - 8
            End If
            '算全部
            td_count += pkr + 1 - 2  '(全部pkr + 1)-(起始終止 2)
            'date_sum = td_count.ToString + "," + th_count.ToString
        End If


    End Function
End Class
