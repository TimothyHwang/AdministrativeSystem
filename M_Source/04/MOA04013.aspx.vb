'==================================================================================================================================================================================
'   Author      : Andy Lin
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'   Action      : Finished at 2010/06/21
'   Description : First version of development
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'   Action      : Modified at 2010/06/25
'   Description : 修正執行查詢後，GridView 操作異動作業時，呈現錯誤問題。
'==================================================================================================================================================================================
Imports System.Data
Imports WebUtilities.Functions

Partial Class Source_04_MOA04013
    Inherits System.Web.UI.Page
    Public sPicPath As String = ConfigurationManager.AppSettings("PicPathUrl")
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

            If LoginCheck.LoginCheck(user_id, "MOA04013") <> "" Then
                LoginCheck.LoginAction(Request.ServerVariables("REMOTE_ADDR"), user_id, "MOA04013.aspx")
                Response.End()
            End If

        End If

        If Not Me.IsPostBack Then

            If Not String.IsNullOrEmpty(Request.Form("query_it_code")) Then

                Me.ViewState("it_code") = Request.Form("query_it_code").Trim()
                tb_it_code.Text = Me.ViewState("it_code")

            End If

            If Not String.IsNullOrEmpty(Request.Form("query_it_name")) Then

                Me.ViewState("it_name") = Request.Form("query_it_name").Trim()
                tb_it_name.Text = Me.ViewState("it_name")

            End If

            GridView1_DataRebinding()

        End If

    End Sub

    Protected Sub ibtnSearch_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ibtnSearch.Click

        tb_it_code.Text = tb_it_code.Text.Trim()
        tb_it_name.Text = tb_it_name.Text.Trim()

        Me.ViewState.Remove("it_code")
        Me.ViewState.Remove("it_name")

        If tb_it_code.Text <> "" Then
            Me.ViewState("it_code") = tb_it_code.Text.Trim()
        End If

        If tb_it_name.Text <> "" Then
            Me.ViewState("it_name") = tb_it_name.Text.Trim()
        End If

        GridView1_DataRebinding()

    End Sub

    Protected Sub ibtnAdd_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ibtnAdd.Click

        Dim parameters As New SortedList()

        parameters.Add("Action", "MOA04014.aspx")

        If Not String.IsNullOrEmpty(Me.ViewState("it_code")) Then
            parameters.Add("query_it_code", Me.ViewState("it_code"))
        End If

        If Not String.IsNullOrEmpty(Me.ViewState("it_name")) Then
            parameters.Add("query_it_name", Me.ViewState("it_name"))
        End If

        Response.Write(Utilities.SubmitFormGeneration(parameters))
        Response.End()

    End Sub

    Private Sub GridView1_DataRebinding()

        SqlDataSource1.SelectParameters.Clear()

        Dim sql_statement As String = SqlDataSource1.SelectCommand

        sql_statement = sql_statement.Substring(0, sql_statement.ToUpper().IndexOf("ORDER BY"))
        sql_statement += "Where 1 = 1"

        If Not String.IsNullOrEmpty(Me.ViewState("it_code")) Then
            sql_statement += " And it_code Like ('%' + @it_code + '%')"
            SqlDataSource1.SelectParameters.Add(New Parameter("it_code", DbType.String, Me.ViewState("it_code")))
        End If

        If Not String.IsNullOrEmpty(Me.ViewState("it_name")) Then
            sql_statement += " And it_name Like ('%' + @it_name + '%')"
            SqlDataSource1.SelectParameters.Add(New Parameter("it_name", DbType.String, Me.ViewState("it_name")))
        End If

        sql_statement += " Order by it_code"

        SqlDataSource1.SelectCommand = sql_statement

    End Sub

    Protected Sub GridView1_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles GridView1.PageIndexChanging
        GridView1_DataRebinding()
    End Sub

    Protected Sub GridView1_RowDeleting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewDeleteEventArgs) Handles GridView1.RowDeleting

        Dim file_filter As String = DirectCast(sender, GridView).Rows(e.RowIndex).Cells(0).Text & "*.*"

        Try

            For Each file_path As String In System.IO.Directory.GetFiles(Server.MapPath("~/M_Source/99"), file_filter)
                System.IO.File.Delete(file_path)
            Next

        Catch ex As Exception

        Finally
            file_filter = Nothing

        End Try

        GridView1_DataRebinding()

    End Sub

    Protected Sub GridView1_RowEditing(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewEditEventArgs) Handles GridView1.RowEditing

        Dim current_row As GridViewRow = GridView1.Rows(e.NewEditIndex)
        Dim parameters As New SortedList()

        parameters.Add("Action", "MOA04014.aspx")
        parameters.Add("it_code", current_row.Cells(0).Text.Trim())
        parameters.Add("it_name", current_row.Cells(1).Text.Trim())
        parameters.Add("it_spec", current_row.Cells(2).Text.Trim())
        parameters.Add("it_cost", current_row.Cells(3).Text.Trim())
        parameters.Add("manufacturer", current_row.Cells(4).Text.Trim())
        parameters.Add("memo", current_row.Cells(5).Text.Trim())
        parameters.Add("snum", current_row.Cells(8).Text.Trim())
        parameters.Add("expired_y", current_row.Cells(9).Text.Trim())
        parameters.Add("it_sort", current_row.Cells(10).Text.Trim())
        parameters.Add("it_unit", current_row.Cells(11).Text.Trim())

        current_row = Nothing

        For Each param As DictionaryEntry In parameters.Clone()
            parameters(param.Key) = Regex.Replace(CType(param.Value, String), "&nbsp;", " ")
        Next

        If Not String.IsNullOrEmpty(Me.ViewState("it_code")) Then
            parameters.Add("query_it_code", Me.ViewState("it_code"))
        End If

        If Not String.IsNullOrEmpty(Me.ViewState("it_name")) Then
            parameters.Add("query_it_name", Me.ViewState("it_name"))
        End If

        Response.Write(Utilities.SubmitFormGeneration(parameters))
        Response.End()

    End Sub

    Protected Sub GridView1_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.Sorted
        GridView1_DataRebinding()
    End Sub

End Class