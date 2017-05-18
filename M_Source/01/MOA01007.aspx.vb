'==================================================================================================================================================================================
'   Author      : Andy Lin
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'   Action      : Finished at 2010/06/24
'   Description : First Version for Development
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'   Action      : Modified at 2010/06/29
'   Description : 修正 Update 資料時，@P_KIND 參數會為 Null 的處理錯誤。
'==================================================================================================================================================================================
Imports System.Data
Imports System.Data.SqlClient

Partial Class Source_04_MOA01007
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

            If LoginCheck.LoginCheck(user_id, "MOA01007") <> "" Then
                LoginCheck.LoginAction(Request.ServerVariables("REMOTE_ADDR"), user_id, "MOA01007.aspx")
                Response.End()
            End If

        End If

        ImageButtonAdd_ClientScriptsPreparing()

    End Sub

    Protected Function ShowMoneyKindText(ByRef kind_value As String) As String

        Dim item As ListItem = ddl_p_kind.Items.FindByValue(kind_value)

        If TypeOf item Is ListItem Then
            Return item.Text.Trim()
        Else
            Return String.Empty
        End If

        item = Nothing

    End Function

    Private Sub ImageButtonAdd_ClientScriptsPreparing()

        scripts.Append("javascript:return StringValidationAfterTrim({ controls: [ ") _
               .AppendFormat("{{ client_id: '{0}', ", tb_p_money.ClientID) _
               .Append("blank_validation: [{ message: '請輸入薪資數字' }], ") _
               .Append("numeric_validation: [{ message: '薪資請輸入數字' }] },") _
               .AppendFormat("{{ client_id: '{0}', ", ddl_p_kind.ClientID) _
               .Append("blank_validation: [{ message: '請選擇薪資種類' }] } ") _
               .Append("]});")

        ibtnAdd.Attributes.Add("OnClick", scripts.ToString())
        scripts.Remove(0, scripts.Length)

    End Sub

    Protected Sub ibtnSearch_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ibtnSearch.Click

        tb_p_money.Text = tb_p_money.Text.Trim()

        Me.ViewState.Remove("P_Money")
        Me.ViewState.Remove("P_KIND")

        If tb_p_money.Text <> "" Then
            Me.ViewState("P_Money") = tb_p_money.Text.Trim()
        End If

        If ddl_p_kind.SelectedValue.Trim() <> "" Then
            Me.ViewState("P_KIND") = ddl_p_kind.SelectedValue.Trim()
        End If

        GridView1_DataRebinding()

    End Sub

    Protected Sub ibtnAdd_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ibtnAdd.Click

        Dim sql_function As New C_SQLFUN
        Dim command As New SqlCommand(String.Empty, New SqlConnection(sql_function.G_conn_string))

        command.CommandType = Data.CommandType.Text

        command.CommandText = "Insert Into P_0102(P_Money, P_KIND) " & _
                              "Values(@P_Money, @P_KIND)"

        command.Parameters.Add(New SqlParameter("P_Money", SqlDbType.Int)).Value = tb_p_money.Text.Trim()
        command.Parameters.Add(New SqlParameter("P_KIND", SqlDbType.Int)).Value = ddl_p_kind.SelectedValue.Trim()

        Try
            command.Connection.Open()
            command.ExecuteNonQuery()

            GridView1.DataBind()

        Catch ex As Exception

            If (TypeOf ex Is SqlException) Then

                If (CType(ex, SqlException).Number.Equals(2627)) Then
                    ClientScript.RegisterStartupScript(Me.GetType(), "SqlErrorMsg", "alert('該筆資料已存在 !');", True)
                End If

            End If

        Finally

            If command.Connection.State.Equals(ConnectionState.Open) Then
                command.Connection.Close()
            End If

        End Try

    End Sub

    Private Sub GridView1_DataRebinding()

        SqlDataSource1.SelectParameters.Clear()

        Dim sql_statement As String = SqlDataSource1.SelectCommand

        sql_statement = sql_statement.Substring(0, sql_statement.ToUpper().IndexOf("ORDER BY"))
        sql_statement += "Where 1 = 1"

        If Not String.IsNullOrEmpty(Me.ViewState("P_Money")) Then

            sql_statement += " And P_Money = @P_Money"
            SqlDataSource1.SelectParameters.Add(New Parameter("P_Money", DbType.Int32, Me.ViewState("P_Money")))

        End If

        If Not String.IsNullOrEmpty(Me.ViewState("P_KIND")) Then

            sql_statement += " And P_KIND = @P_KIND"
            SqlDataSource1.SelectParameters.Add(New Parameter("P_KIND", DbType.Int32, Me.ViewState("P_KIND")))

        End If

        sql_statement += " Order by P_Num"

        SqlDataSource1.SelectCommand = sql_statement

    End Sub

    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound

        If e.Row.RowType.Equals(DataControlRowType.DataRow) Then

            If e.Row.RowState.Equals(DataControlRowState.Edit) Or e.Row.RowState.Equals(DataControlRowState.Alternate Or DataControlRowState.Edit) Then

                Dim imgControl As ImageButton

                imgControl = e.Row.Cells(e.Row.Cells.Count - 1).FindControl("btnImgOK")

                Dim control As Control

                control = e.Row.Cells(1).FindControl("TextBox1")

                scripts.Append("javascript:return StringValidationAfterTrim({ controls: [ ") _
                       .AppendFormat("{{ client_id: '{0}', ", control.ClientID) _
                       .Append("blank_validation: [{ message: '請輸入薪資數字' }], ") _
                       .Append("numeric_validation: [{ message: '薪資請輸入數字' }] },")

                control = e.Row.Cells(e.Row.Cells.Count - 2).FindControl("gv_ddl_p_kind")

                Dim selected_item As ListItem = CType(control, DropDownList).Items.FindByValue( _
                                                CType(e.Row.DataItem, DataRowView)("P_KIND").ToString())

                If TypeOf selected_item Is ListItem Then
                    CType(control, DropDownList).SelectedValue = selected_item.Value
                Else
                    CType(control, DropDownList).SelectedValue = String.Empty
                End If

                selected_item = Nothing

                scripts.AppendFormat("{{ client_id: '{0}', ", control.ClientID) _
                       .Append("blank_validation: [{ message: '請選擇薪資種類' }] } ") _
                       .Append("]});")

                control = Nothing

                imgControl.Attributes.Add("OnClick", scripts.ToString())
                scripts.Remove(0, scripts.Length)

                imgControl = Nothing

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

        Dim ddlControl As DropDownList = GridView1.Rows(GridView1.EditIndex).Cells(2).FindControl("gv_ddl_p_kind")

        If TypeOf ddlControl Is DropDownList Then

            e.Command.Parameters("@P_KIND").Value = ddlControl.SelectedValue.Trim()
            ddlControl = Nothing

        End If

    End Sub

End Class