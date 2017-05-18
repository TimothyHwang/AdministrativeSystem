
Partial Class Source_07_MOA07004
    Inherits System.Web.UI.Page
    Public do_sql As New C_SQLFUN
    Dim dr As System.Data.DataRow
    Dim n_table As New System.Data.DataTable
    Dim p As Integer = 0
    Public stmt As String
    Dim str_role As String
    Dim user_id, org_uid As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        user_id = Session("user_id")
        org_uid = Session("ORG_UID")

        'session被清空回首頁
        If user_id = "" Or org_uid = "" Then

            Response.Write(" <script language='javascript'>")
            Response.Write(" alert('畫面停留太久未使用，將重新整理回首頁');")
            Response.Write(" window.parent.location='../../index.aspx';")
            Response.Write(" </script>")

        Else

            'Call select_grid()
            str_role = Session("Role")

            If IsPostBack Then
                Exit Sub
            End If
            Txt_Sdate.Text = Now.AddDays(-10).ToString("yyyy/MM/dd")
            Txt_Edate.Text = Now.ToString("yyyy/MM/dd")
            '找出上一級單位以下全部單位
            Dim Org_Down As New C_Public
            '判斷登入者權限
            If Session("Role") = "1" Then
                stmt = "select ORG_NAME from AdminGroup order by ORG_NAME"
            ElseIf Session("Role") = "2" Then
                stmt = "select ORG_NAME from AdminGroup WHERE ORG_UID IN (" & Org_Down.getchildorg(Session("PARENT_ORG_UID")) & ") ORDER BY ORG_NAME"
            Else
                stmt = "select ORG_NAME from AdminGroup WHERE ORG_UID ='" & Session("ORG_UID") & "' ORDER BY ORG_NAME"
            End If

            If do_sql.db_sql(stmt, do_sql.G_conn_string) = False Then
                Exit Sub
            End If
            If do_sql.G_table.Rows.Count > 0 Then
                n_table = do_sql.G_table
                p = 0
                If Session("Role") = "1" Then
                    DrDown_PAUNIT.Items.Clear()
                    DrDown_PAUNIT.Items.Add("請選擇")
                    DrDown_PAUNIT.Items(p).Value = ""
                    p += 1
                End If
                For Each dr In n_table.Rows
                    DrDown_PAUNIT.Items.Add(Trim(dr("ORG_NAME").ToString))
                    DrDown_PAUNIT.Items(p).Value = Trim(dr("ORG_NAME").ToString)
                    p += 1
                Next
            End If

        End If

    End Sub

    Sub ADD_DATA()
        Dim table As New System.Data.DataTable
        Dim column As System.Data.DataColumn
        Dim row As System.Data.DataRow
        Dim table_d(100) As String
        Dim sno As Integer = 0
        stmt = "select kind_num,State_order,state_name from syskind where kind_num in(7,8,9,10) ORDER BY  kind_num,State_order,state_name"
        If do_sql.db_sql(stmt, do_sql.G_conn_string) = False Then
            Exit Sub
        End If

        column = New System.Data.DataColumn()
        column.DataType = System.Type.GetType("System.String")
        column.ColumnName = "單位"
        column.ReadOnly = True
        column.Unique = False


        ' Add the Column to the DataColumnCollection.
        table.Columns.Add(column)

        For pi As Integer = 0 To do_sql.G_table.Rows.Count - 1
            ' Create new DataColumn, set DataType, ColumnName 
            ' and add to DataTable.    
            sno += 1
            table_d(sno) = do_sql.G_table.Rows(pi).Item(2).ToString.Trim
            column = New System.Data.DataColumn()
            column.DataType = System.Type.GetType("System.Int32")
            column.ColumnName = do_sql.G_table.Rows(pi).Item(2).ToString.Trim
            column.ReadOnly = True
            column.Unique = False

            ' Add the Column to the DataColumnCollection.
            table.Columns.Add(column)
        Next
        stmt = "select P_07.PAUNIT,p_0701.nREASON,COUNT(*) PER from p_07 inner join p_0701 on p_07.eformsn=p_0701.eformsn"
        stmt += " where  p_07.nAPPTIME >= '" + Txt_Sdate.Text + "' and p_07.nAPPTIME <= '" + Txt_Edate.Text + "'"
        If str_role <> "1" Then
            stmt += " and P_07.PAUNIT='" + DrDown_PAUNIT.SelectedItem.Value + "'"
        End If
        stmt += " GROUP BY P_07.PAUNIT,p_0701.nREASON"
        If do_sql.db_sql(stmt, do_sql.G_conn_string) = False Then
            Exit Sub
        End If
        ' Create three new DataRow objects and add 
        ' them to the DataTable
        ' Dim i As Integer
        'For pi As Integer = 0 To 1
        '    row = table.NewRow()
        '    row(0) = "軍機處軍機處"
        '    row(1) = 21
        '    row(2) = 42
        '    row(3) = 71
        '    table.Rows.Add(row)

        'Next
        Dim str_unit As String = ""
        For pi As Integer = 0 To do_sql.G_table.Rows.Count - 1
            If str_unit <> do_sql.G_table.Rows(pi).Item(0).ToString.Trim Then
                If str_unit <> "" Then
                    table.Rows.Add(row)
                End If
                str_unit = do_sql.G_table.Rows(pi).Item(0).ToString.Trim
                row = table.NewRow()
                row(0) = str_unit
            End If

            For pj As Integer = 1 To sno
                If do_sql.G_table.Rows(pi).Item(1).ToString.Trim = table_d(pj) Then
                    row(pj) = do_sql.G_table.Rows(pi).Item(2).ToString.Trim
                    Exit For
                End If
            Next
            If pi = do_sql.G_table.Rows.Count - 1 Then
                table.Rows.Add(row)
            End If

        Next
        GridView1.DataSource = table
        GridView1.DataBind()


    End Sub

    Protected Sub ImgSearch_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgSearch.Click
        Call ADD_DATA()
    End Sub


    Protected Sub ImgDate1_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgDate1.Click

        Div_grid.Visible = True
        Div_grid.Style("Top") = "70px"
        Div_grid.Style("left") = "70px"

        Calendar1.SelectedDate = Txt_Sdate.Text

    End Sub

    Protected Sub ImgDate2_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgDate2.Click

        Div_grid2.Visible = True
        Div_grid2.Style("Top") = "70px"
        Div_grid2.Style("left") = "190px"

        Calendar2.SelectedDate = Txt_Edate.Text

    End Sub

    Protected Sub Calendar1_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Calendar1.SelectionChanged

        Txt_Sdate.Text = Calendar1.SelectedDate.Date
        Div_grid.Visible = False

    End Sub

    Protected Sub Calendar2_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Calendar2.SelectionChanged

        Txt_Edate.Text = Calendar2.SelectedDate.Date
        Div_grid2.Visible = False

    End Sub

    Protected Sub btnClose1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose1.Click

        Div_grid.Visible = False

    End Sub

    Protected Sub btnClose2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose2.Click

        Div_grid2.Visible = False

    End Sub

End Class
