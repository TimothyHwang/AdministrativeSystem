
Partial Class Source_00_MOA00103
    Inherits System.Web.UI.Page
    Public str_role As String
    Dim stmt As String
    Public do_sql As New C_SQLFUN
    Dim dr As System.Data.DataRow
    Dim n_table As New System.Data.DataTable
    Dim p As Integer
    Dim user_id, org_uid As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            user_id = Session("user_id")
            org_uid = Session("ORG_UID")

            'session被清空回首頁
            If user_id = "" Or org_uid = "" Then

                Response.Write(" <script language='javascript'>")
                Response.Write(" alert('畫面停留太久未使用，將重新整理回首頁');")
                Response.Write(" window.parent.location='../../index.aspx';")
                Response.Write(" </script>")

            Else

                '判斷登入者權限
                Dim LoginCheck As New C_Public

                If LoginCheck.LoginCheck(user_id, "MOA00103") <> "" Then
                    LoginCheck.LoginAction(Request.ServerVariables("REMOTE_ADDR"), user_id, "MOA00103.aspx")
                    Response.End()
                End If

                'Session("Role") = "3"
                str_role = Session("Role")
                do_sql.G_user_id = Session("user_id")
                'do_sql.G_user_id = "tempu175"
                'TextBox1.Text = str_role + "," + Session("ORG_UID") + "," + do_sql.G_user_id
                If IsPostBack Then
                    Exit Sub
                End If

                '找出上一級單位以下全部單位
                Dim Org_Down As New C_Public
                '判斷登入者權限
                If Session("Role") = "1" Then
                    stmt = "select ORG_UID,ORG_NAME from AdminGroup order by ORG_NAME"
                ElseIf Session("Role") = "2" Then
                    stmt = "select ORG_UID,ORG_NAME from AdminGroup WHERE ORG_UID IN (" & Org_Down.getchildorg(Session("PARENT_ORG_UID")) & ") ORDER BY ORG_NAME"
                Else
                    stmt = "select ORG_UID,ORG_NAME from AdminGroup WHERE ORG_UID ='" & Session("ORG_UID") & "' ORDER BY ORG_NAME"
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
                        DrDown_PAUNIT.Items(p).Value = Trim(dr("ORG_UID").ToString)
                        p += 1
                    Next
                Else
                    DrDown_PAUNIT.Items.Clear()
                    DrDown_PAUNIT.Items.Add("")
                    DrDown_PAUNIT.Items(p).Value = ""
                End If

                If DrDown_PAUNIT.SelectedItem.Value <> "" Then
                    Call select_name()
                End If

            End If

        Catch ex As Exception

        End Try

    End Sub

    Protected Sub DrDown_PAUNIT_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DrDown_PAUNIT.SelectedIndexChanged
        If str_role = "1" Or str_role = "2" Then
            Call select_name()
        End If
    End Sub
    Sub select_name()
        DrDown_emp_chinese_name.Items.Clear()
        If DrDown_PAUNIT.SelectedItem.Value <> "" Then
            stmt = "select ORG_UID,emp_chinese_name,employee_id from Employee "
            If str_role = "1" Or str_role = "2" Then
                stmt += "where ORG_UID='" + DrDown_PAUNIT.SelectedItem.Value + "'"
            Else
                stmt += "where employee_id='" + do_sql.G_user_id + "'"
            End If
            stmt += " and leave='Y'"
            stmt += " order by ORG_UID,emp_chinese_name"
            If do_sql.db_sql(stmt, do_sql.G_conn_string) = False Then
                Exit Sub
            End If
            p = 0
            n_table = do_sql.G_table
            For Each dr In n_table.Rows
                DrDown_emp_chinese_name.Items.Add(Trim(dr("emp_chinese_name").ToString))
                DrDown_emp_chinese_name.Items(p).Value = Trim(dr("employee_id").ToString)
                p += 1
            Next
        End If
    End Sub

    Protected Sub btnImgIns_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnImgIns.Click

        If DrDown_emp_chinese_name.SelectedItem.Value = "" Then
            do_sql.G_errmsg = "姓名必須點選"
            Exit Sub
        End If
        If TXT_password.Text = "" Then
            do_sql.G_errmsg = "密碼必須輸入"
            Exit Sub
        End If
        stmt = "update Employee set password='" + TXT_password.Text + "'"
        stmt += " where employee_id='" + DrDown_emp_chinese_name.SelectedItem.Value + "'"
        If do_sql.db_exec(stmt, do_sql.G_conn_string) = False Then
            Exit Sub
        End If
        do_sql.G_errmsg = "存檔成功"

    End Sub
End Class
