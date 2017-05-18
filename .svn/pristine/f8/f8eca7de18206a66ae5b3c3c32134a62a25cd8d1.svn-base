'==================================================================================================================================================================================
'   Author      : Andy Lin
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'   Action      : Finished at 2010/06/22
'   Description : First version of development
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'   Action      : Modified at 2010/06/24
'   Description : 修正執行查詢後，GridView 操作異動作業時，呈現錯誤問題。
'==================================================================================================================================================================================
Imports System.Data
Imports System.Data.SqlClient

Partial Class Source_04_MOA04017
    Inherits System.Web.UI.Page

    Dim user_id, org_uid As String
    Dim scripts As New StringBuilder

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        user_id = Session("user_id")
        org_uid = Session("ORG_UID")

        'Session被清空回首頁
        If user_id = "" Or org_uid = "" Then

            Response.Write("<script language='javascript'>")
            Response.Write("alert('畫面停留太久未使用，將重新整理回首頁');")
            Response.Write("window.parent.location='../../index.aspx';")
            Response.Write("</script>")

        Else

            '判斷登入者權限
            Dim LoginCheck As New C_Public

            If LoginCheck.LoginCheck(user_id, "MOA04017") <> "" Then
                LoginCheck.LoginAction(Request.ServerVariables("REMOTE_ADDR"), user_id, "MOA04017.aspx")
                Response.End()
            End If

        End If

        lbl_map_code.Text = ddl_map_code.SelectedValue.Trim()
        ImageButtons_ClientScriptsPreparing()

    End Sub

    Private Sub ImageButtons_ClientScriptsPreparing()

        scripts.Append("javascript:return StringValidationAfterTrim({ controls: [ ") _
               .AppendFormat("{{ client_id: '{0}', ", tb_rnum_code.ClientID) _
               .Append("blank_validation: [{ message: '請輸入房間(區域)編號' }] }, ") _
               .AppendFormat("{{ client_id: '{0}', ", tb_rnum_spec.ClientID) _
               .Append("numeric_validation: [{ message: '坪數請輸入數字' }] }, ") _
               .AppendFormat("{{ client_id: '{0}', ", ddl_map_code.ClientID) _
               .Append("blank_validation: [{ message: '請選擇圖資編號/名稱' }] }") _
               .Append("]});")

        ibtnAdd.Attributes.Add("OnClick", scripts.ToString())
        scripts.Remove(0, scripts.Length)

        scripts.Append("javascript:return StringValidationAfterTrim({ controls: [ ") _
               .AppendFormat("{{ client_id: '{0}', ", tb_rnum_spec.ClientID) _
               .Append("numeric_validation: [{ message: '坪數請輸入數字' }] } ") _
               .Append("]});")

        ibtnSearch.Attributes.Add("OnClick", scripts.ToString())
        scripts.Remove(0, scripts.Length)

    End Sub

    Protected Sub ibtnSearch_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ibtnSearch.Click

        tb_rnum_code.Text = tb_rnum_code.Text.Trim()
        tb_rnum_name.Text = tb_rnum_name.Text.Trim()
        tb_rnum_spec.Text = tb_rnum_spec.Text.Trim()
        tb_unit_code.Text = tb_unit_code.Text.Trim()

        Me.ViewState.Remove("rnum_code")
        Me.ViewState.Remove("rnum_name")
        Me.ViewState.Remove("rnum_spec")
        Me.ViewState.Remove("unit_code")
        Me.ViewState.Remove("map_code")

        If tb_rnum_code.Text <> "" Then
            Me.ViewState("rnum_code") = tb_rnum_code.Text.Trim()
        End If

        If tb_rnum_name.Text <> "" Then
            Me.ViewState("rnum_name") = tb_rnum_name.Text.Trim()
        End If

        If tb_rnum_spec.Text <> "" Then
            Me.ViewState("rnum_spec") = tb_rnum_spec.Text.Trim()
        End If

        If tb_unit_code.Text <> "" Then
            Me.ViewState("unit_code") = tb_unit_code.Text.Trim()
        End If

        If ddl_map_code.SelectedValue.Trim() <> "" Then
            Me.ViewState("map_code") = ddl_map_code.SelectedValue.Trim()
        End If

        GridView1_DataRebinding()

    End Sub

    Protected Sub ibtnAdd_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ibtnAdd.Click

        Dim sql_function As New C_SQLFUN
        Dim command As New SqlCommand(String.Empty, New SqlConnection(sql_function.G_conn_string))

        command.CommandType = Data.CommandType.Text

        command.CommandText = "Insert Into P_0411(rnum_code, rnum_name, rnum_spec, unit_code, map_code) " & _
                              "Values(@rnum_code, @rnum_name, @rnum_spec, @unit_code, @map_code)"

        command.Parameters.Add(New SqlParameter("rnum_code", SqlDbType.VarChar, 5)).Value = tb_rnum_code.Text.Trim()
        command.Parameters.Add(New SqlParameter("rnum_name", SqlDbType.NVarChar, 100)).Value = tb_rnum_name.Text.Trim()

        Try
            Decimal.Parse(tb_rnum_spec.Text)
            command.Parameters.Add(New SqlParameter("rnum_spec", SqlDbType.Decimal, 29)).Value = Decimal.Parse(tb_rnum_spec.Text)
        Catch ex As Exception
            command.Parameters.Add(New SqlParameter("rnum_spec", SqlDbType.Decimal, 29)).Value = DBNull.Value
        End Try

        command.Parameters.Add(New SqlParameter("unit_code", SqlDbType.VarChar, 4)).Value = tb_unit_code.Text.Trim()
        command.Parameters.Add(New SqlParameter("map_code", SqlDbType.VarChar, 20)).Value = ddl_map_code.SelectedValue.Trim()

        Try
            command.Connection.Open()
            command.ExecuteNonQuery()

            GridView1.DataBind()

        Catch ex As Exception

            If (TypeOf ex Is SqlException) Then
                If (CType(ex, SqlException).Number.Equals(2627)) Then
                    ClientScript.RegisterStartupScript(Me.GetType(), "SqlErrorMsg", "alert('該筆資料已存在 !');", True)
                Else
                    ClientScript.RegisterStartupScript(Me.GetType(), "SqlErrorMsg", "alert('" + ex.Message + "!');", True)
                End If
            Else
                ClientScript.RegisterStartupScript(Me.GetType(), "InsertErrorMsg", "alert('" + ex.Message + "!');", True)
            End If

        Finally

            If command.Connection.State.Equals(ConnectionState.Open) Then
                command.Connection.Close()
            End If

        End Try

    End Sub

    Protected Sub ddl_map_code_DataBound(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddl_map_code.DataBound

        ddl_map_code.Items.Insert(0, New ListItem(String.Empty, String.Empty))

        scripts.AppendFormat("var map_code = document.getElementById('{0}');", lbl_map_code.ClientID) _
               .Append("if (map_code) map_code.innerText = event.srcElement.value.Trim();")

        ddl_map_code.Attributes("OnChange") = scripts.ToString()
        scripts.Remove(0, scripts.Length)

    End Sub

    Private Sub GridView1_DataRebinding()

        SqlDataSource1.SelectParameters.Clear()

        Dim sql_statement As String = SqlDataSource1.SelectCommand

        sql_statement = sql_statement.Substring(0, sql_statement.ToUpper().IndexOf("ORDER BY"))
        sql_statement += "Where 1 = 1"

        If Not String.IsNullOrEmpty(Me.ViewState("rnum_code")) Then

            sql_statement += " And rnum_code = @rnum_code"
            SqlDataSource1.SelectParameters.Add(New Parameter("rnum_code", DbType.String, Me.ViewState("rnum_code")))

        End If

        If Not String.IsNullOrEmpty(Me.ViewState("rnum_name")) Then

            sql_statement += " And rnum_name Like '%' + @rnum_name + '%'"
            SqlDataSource1.SelectParameters.Add(New Parameter("rnum_name", DbType.String, Me.ViewState("rnum_name")))

        End If

        If Not String.IsNullOrEmpty(Me.ViewState("rnum_spec")) Then

            sql_statement += " And rnum_spec = @rnum_spec"
            SqlDataSource1.SelectParameters.Add(New Parameter("rnum_spec", DbType.Decimal, Me.ViewState("rnum_spec")))

        End If

        If Not String.IsNullOrEmpty(Me.ViewState("unit_code")) Then

            sql_statement += " And unit_code = @unit_code"
            SqlDataSource1.SelectParameters.Add(New Parameter("unit_code", DbType.String, Me.ViewState("unit_code")))

        End If

        If Not String.IsNullOrEmpty(Me.ViewState("map_code")) Then

            sql_statement += " And map_code = @map_code"
            SqlDataSource1.SelectParameters.Add(New Parameter("map_code", DbType.String, Me.ViewState("map_code")))

        End If

        sql_statement += " Order by rnum_code"

        SqlDataSource1.SelectCommand = sql_statement

    End Sub

    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound

        If e.Row.RowType.Equals(DataControlRowType.DataRow) Then

            If e.Row.RowState.Equals(DataControlRowState.Edit) Or e.Row.RowState.Equals(DataControlRowState.Alternate Or DataControlRowState.Edit) Then

                Dim control As Control = e.Row.Cells(e.Row.Cells.Count - 2).FindControl("gv_ddl_map_code")

                If TypeOf control Is DropDownList Then

                    Dim ddlControl As DropDownList = CType(control, DropDownList)

                    ddlControl.Items.Insert(0, New ListItem(String.Empty, String.Empty))

                    Dim selected_item As ListItem = ddlControl.Items.FindByValue(CType(e.Row.DataItem, DataRowView)("map_code").ToString())

                    If TypeOf selected_item Is ListItem Then
                        ddlControl.SelectedValue = selected_item.Value
                    Else
                        ddlControl.SelectedValue = String.Empty
                    End If

                    selected_item = Nothing

                    CType(e.Row.Cells(e.Row.Cells.Count - 2).FindControl("gv_lbl_map_code"), Label).Text = ddlControl.SelectedValue

                    scripts.AppendFormat("var map_code = document.getElementById('{0}');", e.Row.Cells(e.Row.Cells.Count - 2).FindControl("gv_lbl_map_code").ClientID) _
                           .Append("if (map_code) map_code.innerText = event.srcElement.value.Trim();")

                    ddlControl.Attributes("OnChange") = scripts.ToString()
                    scripts.Remove(0, scripts.Length)

                    ddlControl = Nothing
                    control = Nothing

                End If

                control = e.Row.Cells(e.Row.Cells.Count - 1).FindControl("btnImgOK")

                If TypeOf control Is ImageButton Then

                    scripts.Append("javascript:return StringValidationAfterTrim({ controls: [ ") _
                           .AppendFormat("{{ client_id: '{0}', ", CType(e.Row.Cells(2).FindControl("TextBox2"), TextBox).ClientID) _
                           .Append("numeric_validation: [{ message: '坪數請輸入數字' }] }, ") _
                           .AppendFormat("{{ client_id: '{0}', ", CType(e.Row.Cells(e.Row.Cells.Count - 2).FindControl("gv_ddl_map_code"), DropDownList).ClientID) _
                           .Append("blank_validation: [{ message: '請選擇圖資編號/名稱' }] }") _
                           .Append("]});")

                    CType(control, ImageButton).Attributes.Add("OnClick", scripts.ToString())
                    scripts.Remove(0, scripts.Length)

                    control = Nothing

                End If

            End If

        End If

    End Sub

    Protected Sub GridView1_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles GridView1.PageIndexChanging
        GridView1_DataRebinding()
    End Sub

    Protected Sub GridView1_RowCancelingEdit(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewCancelEditEventArgs) Handles GridView1.RowCancelingEdit
        GridView1_DataRebinding()
    End Sub

    Protected Sub GridView1_RowDeleting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewDeleteEventArgs) Handles GridView1.RowDeleting
        GridView1_DataRebinding()
    End Sub

    Protected Sub GridView1_RowEditing(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewEditEventArgs) Handles GridView1.RowEditing
        GridView1_DataRebinding()
    End Sub

    Protected Sub GridView1_RowUpdating(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewUpdateEventArgs) Handles GridView1.RowUpdating
        GridView1_DataRebinding()
    End Sub

    Protected Sub GridView1_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.Sorted
        GridView1_DataRebinding()
    End Sub

    Protected Sub SqlDataSource1_Updating(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.SqlDataSourceCommandEventArgs) Handles SqlDataSource1.Updating

        Dim current_row As GridViewRow = GridView1.Rows(GridView1.EditIndex)

        If TypeOf current_row Is GridViewRow Then

            Dim ddlControl As DropDownList = current_row.Cells(current_row.Cells.Count - 2).FindControl("gv_ddl_map_code")

            If TypeOf ddlControl Is DropDownList Then

                e.Command.Parameters("@map_code").Value = ddlControl.SelectedValue
                ddlControl = Nothing

            End If

            '判斷單位編號必需輸入4個字以內的英數字
            Dim tb_unit_no As TextBox = current_row.Cells(current_row.Cells.Count - 2).FindControl("TextBox3")
            If tb_unit_no.Text.Trim() <> "" Then
                Dim mlen1 As Integer = 0 : Dim mlen2 As Integer = 0 : Dim mtxt2len As Integer = 0
                For mj = 1 To Len(Trim(tb_unit_no.Text))
                    Dim mstr As String = Mid(Trim(tb_unit_no.Text.Trim()), mj, 1)
                    If Asc(mstr) >= 1 And Asc(mstr) <= 255 Then
                        mlen1 = mlen1 + 1 '這是英文字
                    Else
                        mlen2 = mlen2 + 1 '這是中文字
                    End If
                Next
                mtxt2len = mlen1 + mlen2 * 2 '這是該字串所佔的位置
                If mtxt2len > 4 Then
                    e.Cancel = True
                    ClientScript.RegisterStartupScript(Me.GetType(), "UpdateValueError", "alert('您輸入的單位編號有誤，請輸入4個字以內英數字單位編號!');", True)
                End If
            End If

            current_row = Nothing

        End If

    End Sub

End Class