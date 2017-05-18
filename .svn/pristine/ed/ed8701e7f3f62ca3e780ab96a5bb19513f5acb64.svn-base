'==================================================================================================================================================================================
'   Author      : Andy Lin
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'   Action      : Modified at 2010/06/17
'   Description : 修正 Insert SQL Statement，改以傳遞參數的方式操作，提高安全性。
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'   Action      : Modified at 2010/06/21
'   Description : (1) 修正 Javascript 的 StringValidationAfterTrim 之使用方法。
'                 (2) 修正查詢功能，該以傳遞參數的方式執行，提升安全性。
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'   Action      : Modified at 2010/06/24
'   Description : 修正執行查詢後，GridView 操作異動作業時，呈現錯誤問題。
'==================================================================================================================================================================================
Imports System.Data
Imports System.Data.SqlClient

Partial Class Source_04_MOA04009
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

            If LoginCheck.LoginCheck(user_id, "MOA04009") <> "" Then
                LoginCheck.LoginAction(Request.ServerVariables("REMOTE_ADDR"), user_id, "MOA04009.aspx")
                Response.End()
            End If

        End If

        ImageButtonAdd_ClientScriptsPreparing()

    End Sub

    Private Sub ImageButtonAdd_ClientScriptsPreparing()

        scripts.Append("javascript:return StringValidationAfterTrim({ controls: [ ") _
               .AppendFormat("{{ client_id: '{0}', ", tb_bd_code.ClientID) _
               .Append("blank_validation: [{ message: '請輸入建物編號' }] }, ") _
               .AppendFormat("{{ client_id: '{0}', ", tb_bd_name.ClientID) _
               .Append("blank_validation: [{ message: '請輸入建物名稱' }] }, ") _
               .AppendFormat("{{ client_id: '{0}', ", tb_no.ClientID) _
               .Append("blank_validation: [{ message: '請輸入排序順序' }], ") _
               .Append("numeric_validation: [{ message: '排序順序請輸入數字' }] }") _
               .Append("]});")

        ibtnAdd.Attributes.Add("OnClick", scripts.ToString())
        scripts.Remove(0, scripts.Length)

    End Sub

    Protected Sub ibtnSearch_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ibtnSearch.Click

        tb_bd_code.Text = tb_bd_code.Text.Trim()
        tb_bd_name.Text = tb_bd_name.Text.Trim()
        tb_no.Text = tb_no.Text.Trim()

        Me.ViewState.Remove("bd_code")
        Me.ViewState.Remove("bd_name")
        Me.ViewState.Remove("no")

        If tb_bd_code.Text <> "" Then
            Me.ViewState("bd_code") = tb_bd_code.Text.Trim()
        End If

        If tb_bd_name.Text <> "" Then
            Me.ViewState("bd_name") = tb_bd_name.Text.Trim()
        End If

        If tb_no.Text <> "" Then
            Dim bg_no As Integer = 0
            Try
                bg_no = Integer.Parse(tb_no.Text.Trim())
            Catch ex As Exception
                ClientScript.RegisterStartupScript(Me.GetType(), "InputValueError", "alert('您欲查詢的建物排序順序資料不正確，請輸入數字!');", True)
            End Try
            If bg_no > 0 Then
                Me.ViewState("no") = tb_no.Text.Trim()
            End If
        End If

        GridView1_DataRebinding()

    End Sub

    Protected Sub ibtnAdd_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ibtnAdd.Click

        Dim bd_code As String = tb_bd_code.Text.Trim()
        Dim bd_name As String = tb_bd_name.Text.Trim()
        Dim bd_no As Integer = 0
        Try
            bd_no = Integer.Parse(tb_no.Text.Trim())
        Catch ex As Exception
            ClientScript.RegisterStartupScript(Me.GetType(), "InputValueError", "alert('您欲新增的建物排序順序資料不正確，請輸入數字!');", True)
        End Try

        If bd_no > 0 Then

            Dim sql_function As New C_SQLFUN
            Dim command As New SqlCommand(String.Empty, New SqlConnection(sql_function.G_conn_string))

            command.CommandType = Data.CommandType.Text

            command.CommandText = "Insert Into P_0404(bd_code, bd_name, no) " & _
                                  "Values(@bd_code, @bd_name, @no)"

            command.Parameters.Add(New SqlParameter("bd_code", SqlDbType.VarChar, 1)).Value = bd_code
            command.Parameters.Add(New SqlParameter("bd_name", SqlDbType.VarChar, 50)).Value = bd_name
            command.Parameters.Add(New SqlParameter("no", SqlDbType.Int)).Value = bd_no

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
        Else
            ClientScript.RegisterStartupScript(Me.GetType(), "InputValueError", "alert('您欲新增的建物排序順序資料不正確，請輸入大於0的數字!');", True)
        End If
    End Sub

    Private Sub GridView1_DataRebinding()

        SqlDataSource1.SelectParameters.Clear()

        Dim sql_statement As String = SqlDataSource1.SelectCommand

        sql_statement = sql_statement.Substring(0, sql_statement.ToUpper().IndexOf("ORDER BY"))
        sql_statement += "Where 1 = 1"

        If Not String.IsNullOrEmpty(Me.ViewState("bd_code")) Then

            sql_statement += " And bd_code = @bd_code"
            SqlDataSource1.SelectParameters.Add(New Parameter("bd_code", DbType.String, Me.ViewState("bd_code")))

        End If

        If Not String.IsNullOrEmpty(Me.ViewState("bd_name")) Then

            sql_statement += " And bd_name Like ('%' + @bd_name + '%')"
            SqlDataSource1.SelectParameters.Add(New Parameter("bd_name", DbType.String, Me.ViewState("bd_name")))

        End If

        If Not String.IsNullOrEmpty(Me.ViewState("no")) Then

            sql_statement += " And no = @no"
            SqlDataSource1.SelectParameters.Add(New Parameter("no", DbType.Int32, Me.ViewState("no")))

        End If

        sql_statement += " Order by no"

        SqlDataSource1.SelectCommand = sql_statement

    End Sub

    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound

        If e.Row.RowType.Equals(DataControlRowType.DataRow) Then

            If e.Row.RowState.Equals(DataControlRowState.Edit) Or e.Row.RowState.Equals(DataControlRowState.Alternate Or DataControlRowState.Edit) Then

                Dim imgControl As ImageButton

                imgControl = e.Row.Cells(3).FindControl("btnImgOK")

                Dim tbControl As TextBox

                tbControl = e.Row.Cells(1).FindControl("TextBox2")

                scripts.Append("javascript:return StringValidationAfterTrim({ controls: [ ") _
                       .AppendFormat("{{ client_id: '{0}', ", tbControl.ClientID) _
                       .Append("blank_validation: [{ message: '請輸入建物名稱' }] }, ")

                tbControl = e.Row.Cells(2).FindControl("TextBox3")

                scripts.AppendFormat("{{ client_id: '{0}', ", tbControl.ClientID) _
                       .Append("blank_validation: [{ message: '請輸入排序順序' }] , ") _
                       .Append("numeric_validation: [{ message: '排序順序請輸入數字' }] }") _
                       .Append("]});")

                imgControl.Attributes.Add("OnClick", scripts.ToString())
                scripts.Remove(0, scripts.Length)

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

End Class